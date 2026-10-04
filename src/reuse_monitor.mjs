import { defineTool } from '@deepseek-ai/dsh-tools';
import { createUserMessage } from '@deepseek-ai/dsh-llm';
import { resolve } from 'node:path';
import { readFile, writeFile, rename, lstat } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import { sha256, policyFor, workspaceFor } from './runtime.mjs';
import { safeDirectory } from './method_execution.mjs';
import { expandContract } from './contract_builder.mjs';
import { saveReuseEvidence,loadReuseEvidence } from './reuse_evidence.mjs';

const textOf=message=>typeof message==='string'?message:
  (Array.isArray(message.content)?message.content.filter(p=>p.type==='text').map(p=>p.text).join('\n'):message.content??message.text??'');
const canonical=value=>Array.isArray(value)?value.map(canonical):value&&typeof value==='object'?
  Object.fromEntries(Object.keys(value).sort().map(k=>[k,canonical(value[k])])):value;
export function installReuseMonitor(ctx,config,configuredRoot) {
  if(!config.methodDirectory || !config.methodRuntimeLock) return;
  let serial=Promise.resolve();
  const update=(exec,change)=>{
    const root=workspaceFor(ctx,configuredRoot,exec,config);
    policyFor(ctx,root,exec,config,true);
    const folder=resolve(root,'.office-plugin'),file=resolve(folder,'reuse-events.json');
    const work=serial.then(async()=>{
      await safeDirectory(root,folder);
      let state;try {
        if((await lstat(file)).isSymbolicLink())throw new Error('OFFICE_REUSE_EVENTS_LINK');
        const text=await readFile(file,'utf8');if(Buffer.byteLength(text)>1024*1024)throw new Error('OFFICE_REUSE_EVENTS_LIMIT');
        state=JSON.parse(text);
      }catch(e){if(e.code!=='ENOENT')throw e;state={events:[],successes:[]};}
      const result=await change(state);state.events=state.events.slice(-100);state.successes=state.successes.slice(-100);
      if(state.seen_user_messages)state.seen_user_messages=state.seen_user_messages.slice(-100);
      const temp=resolve(folder,randomUUID()+'.tmp');await writeFile(temp,JSON.stringify(state,null,2),{flag:'wx'});await rename(temp,file);return result;
    });serial=work.catch(()=>{});return work;
  };
  const writing=exec=>{try {policyFor(ctx,workspaceFor(ctx,configuredRoot,exec,config),exec,config,true);return true;}catch(e){if(e.message.includes('OFFICE_WRITE_DENIED'))return false;throw e;}};
  ctx.on('tools/post-execute',async(exec,result,next)=>{
    const decision=await next();
    if(!['office_check_output','office_restore_template','office_method_run'].includes(exec.name) || result.isError || !result.value?.complete || decision.kind!=='accept' || !writing(exec))return decision;
    const value=result.value;
    const signature=sha256(JSON.stringify(canonical({contract:value.contract_sha256??exec.arguments.contract,method_id:value.method_id,count:value.source_sha256.length})));
    const batch=sha256(JSON.stringify({sources:value.source_sha256,target:value.target_sha256,artifact:value.artifact_sha256??value.generated?.[0]?.sha256}));
    const event=await update(exec,async state=>{
      const previous=state.successes.find(s=>s.signature===signature&&s.batch!==batch);
      if(state.successes.some(s=>s.signature===signature&&s.batch===batch))return null;
      const item={id:randomUUID(),type:previous?'different_batch_success':'checked_success',status:'pending',
        signature,batch,created_at:new Date().toISOString(),session:exec.agent?.session?.id??null,
        script_path:exec.arguments.script_path??null,method_id:value.method_id??null,output_path:value.output_path??value.output_paths?.[0],
        source_paths:value.source_paths??exec.arguments.source_paths??null,target_path:value.target_path??exec.arguments.target_path??null};
      item.required_targets=value.required_targets;item.source_sha256=value.source_sha256;item.target_sha256=value.target_sha256;
      state.successes.push({signature,batch,session:item.session,method_id:item.method_id,output_path:item.output_path,target_sha256:value.target_sha256});
      if(exec.name==='office_method_run')return null;
      item.evidence=await saveReuseEvidence(ctx,workspaceFor(ctx,configuredRoot,exec,config),exec,item,value);
      state.events.push(item);return item;
    });
    if(!event)return decision;
    return {...decision,content:[...(decision.content??result.content),{type:'text',text:
      `复用评估事件 ${event.id}已保存：${event.type==='different_batch_success'?'同一检查要求已在新批次通过。':'本次完整检查通过。'}先交付文件，再评估可重复步骤。用office_reuse_pending(event_id)取得已核对清单与脚本版本；适合复用则保存、验证、启用，否则记录理由。`}]};
  });
  // Only continue this live turn once. inject() does not start a new/idle turn.
  ctx.on('agent/turn-stopping',async({agent,turn,signal})=>{
    if(config.reuseContinuation===false || signal.aborted || typeof agent?.inject!=='function')return;
    const exec={agent,signal};if(!writing(exec))return;
    const event=await update(exec,state=>{
      if(signal.aborted)return null;
      state.reminded_turns??=[];
      const key=`${agent.session.id}:${turn}`;
      if(state.reminded_turns.includes(key))return null;
      const item=state.events.findLast(e=>e.session===agent.session.id && e.type!=='followup_message' &&
        ['pending','method_pending','draft_saved'].includes(e.status) && !e.continuation_at);
      if(!item)return null;
      item.continuation_at=new Date().toISOString();state.reminded_turns.push(key);
      state.reminded_turns=state.reminded_turns.slice(-100);return item;
    });
    if(!event||signal.aborted)return;
    agent.inject(createUserMessage({source:{kind:'office-reuse',event_id:event.id},content:[{type:'text',text:
      `文件已经完整核对，继续一次复用评估：事件 ${event.id}。认领候选可直接取得已核对清单，不重做正确文件。若步骤能处理同类新资料，完成方法保存、验证、启用；只适合此次则说明理由。若适配失败或本轮余量不足，记录具体待处理项后结束。本轮仅此一次提醒。`}]}));
  });
  ctx.on('agent/pre-step',async({agent,messages,signal},next)=>{
    const decision=await next();if(decision.kind==='reject'||signal.aborted||!messages.length)return decision;
    const exec={agent,signal};if(!writing(exec))return decision;
    const lastUser=messages.filter(message=>message.role==='user' && message.source?.kind!=='office-reuse').at(-1);
    const userText=lastUser?textOf(lastUser):'';
    if(!userText)return decision;
    const event=await update(exec,state=>{
      const digest=sha256(lastUser.id??userText),session=agent.session.id;
      const existing=state.events.find(e=>e.type==='followup_message'&&e.digest===digest&&e.session===session);
      if(existing)return existing.status==='pending'?existing:null;
      state.seen_user_messages??=[];
      if(state.seen_user_messages.some(m=>m.digest===digest&&m.session===session))return null;
      state.seen_user_messages.push({digest,session});
      const related=state.successes.findLast(s=>s.session===session);
      if(!related)return null;
      const item={id:randomUUID(),type:'followup_message',status:'pending',digest,session,
        created_at:new Date().toISOString(),message:userText.slice(0,2000),related_success:related};state.events.push(item);return item;
    });
    if(!event)return decision;
    return {...decision,messages:[...decision.messages,createUserMessage({source:{kind:'office-reuse',event_id:event.id},content:[{type:'text',text:
      `插件记录了本次后续消息，事件 ${event.id}。先判断它是用户纠错、新需求还是仅更换数据；若影响旧方法，停用受影响版本再修正，不把全部方法停用。用office_reuse_review记录判断和证据。不要把用户消息直接当作批准的业务事实。`}]})]};
  });
  ctx.tools.register(defineTool({name:'office_reuse_review',description:
    '记录真实触发事件的判断：correction用户纠错、repeated可重复工作、new_request新需求、data_change只换数据、not_reusable不能复用。说明可重复步骤、适用结构与检查依据或不保存理由。该操作不等于保存代码；值得复用需继续save/validate/enable，已交付文件无须重复执行。',
    parameters:{event_id:{type:'string',required:true},classification:{type:'string',required:true},reason:{type:'string',required:true}},
    output:{schema:{type:'json'},render:(_a,v)=>[{type:'text',text:JSON.stringify(v)}]},
    execute:async(args,exec)=>{
      policyFor(ctx,workspaceFor(ctx,configuredRoot,exec,config),exec,config,true);
      if(!['correction','repeated','new_request','data_change','not_reusable'].includes(args.classification)||!args.reason.trim()||args.reason.length>2000)throw new Error('OFFICE_REUSE_REVIEW_INVALID');
      const item=await update(exec,state=>{
        const event=state.events.find(e=>e.id===args.event_id);
        if(!event || event.session!==(exec.agent?.session?.id??null))throw new Error('OFFICE_REUSE_EVENT_UNKNOWN: 在当前工作区用office_reuse_pending查看真实候选，再按事件ID评估');
        if(event.status!=='pending')throw new Error('OFFICE_REUSE_EVENT_REVIEWED');
        event.status=['correction','repeated'].includes(args.classification)?'method_pending':'reviewed';event.classification=args.classification;event.reason=args.reason;
        event.reviewed_at=new Date().toISOString();event.method_work_started_at=Date.now();return event;
      });
      if(!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
      return {event_id:item.id,status:item.status,classification:item.classification,code_saved:false,
        next:['correction','repeated'].includes(item.classification)?
          '用office_reuse_pending(event_id)沿用已核对清单和脚本版本，用office_help(topic=method_code)适配真实接口，再保存、验证、启用；无需重做正确文件。':
          '按本次需要继续，不强制保存方法'};
    }}));
  ctx.tools.register(defineTool({name:'office_reuse_pending',description:'查看当前工作区待整理复用候选。新会话继续整理时传event_id认领，再review/save/validate/enable；不会执行代码或重做文件。',
    parameters:{event_id:{type:'string',description:'可选，要在当前会话继续整理的真实候选ID'}},
    output:{schema:{type:'json'},render:(_a,v)=>[{type:'text',text:JSON.stringify(v)}]},
    execute:async(args,exec)=>{
      policyFor(ctx,workspaceFor(ctx,configuredRoot,exec,config),exec,config,true);
      const items=await update(exec,async state=>{
        const candidates=state.events.filter(e=>['pending','method_pending','draft_saved'].includes(e.status) && e.type!=='followup_message');
        if(args.event_id){
          const event=candidates.find(e=>e.id===args.event_id);
          if(!event)throw new Error('OFFICE_REUSE_PENDING_UNKNOWN: 该候选不在当前工作区待处理记录中');
          event.session=exec.agent?.session?.id??null;event.claimed_at=new Date().toISOString();
        }
        return Promise.all(candidates.filter(e=>!args.event_id||e.id===args.event_id).slice(-10).map(async e=>({event_id:e.id,status:e.status,type:e.type,script_path:e.script_path,
          source_paths:e.source_paths,target_path:e.target_path,output_path:e.output_path,saved_method_id:e.saved_method_id??null,
          required_targets:e.required_targets,created_at:e.created_at,
          ...(args.event_id?await loadReuseEvidence(ctx,workspaceFor(ctx,configuredRoot,exec,config),exec,e):{})})));
      });
      if(!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
      return {candidates:items,code_executed:false,next:'核对资料仍与候选对应，再继续整理、保存及验证代码'};
    }}));
  ctx.get('systemPrompt')?.section({name:'office-reusable-code',order:90,interpolate:false,text:
    '已有方法用office_method_list查用途，相符就run；工具完整通过后交付，有新要求或疑点才补查。无适用方法照常用原生脚本。用户纠错先修文件，停用确实受影响方法。完整检查保存来源关系及声明脚本版本。交付后用office_reuse_pending认领候选，一次评估是否适合重复工作；适合则review、save、validate、enable，不适合则记录理由。取得的confirmed_contract可直接用于保存，避免重新拼装；输入已变化则先修订。方法代码用宿主办公操作，详例office_help(topic=method_code)。验证是暂存检查、启用是入库、run才发布新文件。分别说明文件与方法状态。'});
  ctx.get('systemPrompt')?.section({name:'office-native-handoff',order:91,interpolate:false,text:
    '按需用原生脚本或office_read补查疑难字段，复用本次已读依据；读完相关栏目再判断缺失，不只取固定前几行。运行脚本先看office_help(topic=runtime)。候选与最终文件用不同新名称；office_restore_template按完整清单保留模板并发布。模型确定范围和来源关系，tables按字段及记录展开；不规则位置用真实bindings，Word沿用requirements。长清单用contract_path，细例office_help(topic=checklist)。按summary.next_action修错误部分；真实0正常填写，不猜缺失值。'});
  const find=(state,id,exec)=>{
    const event=state.events.find(e=>e.id===id&&e.session===(exec.agent?.session?.id??null));
    if(!event||event.status!=='method_pending')throw new Error('OFFICE_REUSE_EVENT_NOT_METHOD_PENDING');return event;
  };
  return {
    check:(id,exec,contract)=>update(exec,state=>{
      contract=expandContract(contract);
      const event=find(state,id,exec);
      if(JSON.stringify(event.required_targets)!==JSON.stringify(contract?.requirements?.map(r=>r.target_ref).sort()))
        throw new Error('OFFICE_REUSE_SCOPE_DIFFERENT: 关联事件必须对应本次完整要求，其他用途可另存不关联');
      return true;
    }),
    saved:(id,method,exec)=>update(exec,state=>{const event=find(state,id,exec);event.saved_method_id=method;event.status='draft_saved';return true;}),
    revised:(parent,method,exec,contract)=>update(exec,state=>{
      const event=state.events.find(e=>e.saved_method_id===parent&&e.session===(exec.agent?.session?.id??null));
      if(event&&JSON.stringify(event.required_targets)===JSON.stringify(contract?.requirements?.map(r=>r.target_ref).sort())){
        event.saved_method_id=method;event.status='draft_saved';return true;
      }return false;
    }),
    enabled:(method,exec,proof)=>update(exec,state=>{
      const event=state.events.find(e=>e.saved_method_id===method&&e.session===(exec.agent?.session?.id??null));
      if(event&&JSON.stringify(event.source_sha256)===JSON.stringify(proof.source_sha256)&&event.target_sha256===proof.target_sha256){event.status='method_enabled';event.enabled_at=new Date().toISOString();event.method_work_elapsed_ms=event.method_work_started_at?Date.now()-event.method_work_started_at:null;return true;}
      return false;
    }),
  };
}
