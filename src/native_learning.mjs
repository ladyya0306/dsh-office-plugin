import { writeFile,rename,open,unlink } from 'node:fs/promises';
import { resolve } from 'node:path';
import { randomUUID } from 'node:crypto';
import { performance } from 'node:perf_hooks';
import { readInput,policyFor,sha256,workspaceFor } from './runtime.mjs';
import { safeDirectory } from './method_execution.mjs';
import { hostDigest,writeHostDigest } from './native_host_state.mjs';
import { loadNativeRun } from './native_execution.mjs';
import { createUserMessage } from '@deepseek-ai/dsh-llm';
import { defineTool } from '@deepseek-ai/dsh-tools';
const ruleHash=run=>run.relationship_rule?sha256(JSON.stringify(run.relationship_rule)):
  run.group_checks?sha256(JSON.stringify({artifacts:run.group_checks.artifacts.map(a=>a.contract),consistency:run.group_checks.consistency})):null;

export class NativeLearning {
  constructor(ctx,config,configuredRoot,store) {Object.assign(this,{ctx,config,configuredRoot,store});this.pending=new Map();}
  owner(exec){return exec.agent?.session?.id??'agentless-probe';}
  file(exec){return '.office-plugin/learning/'+sha256(this.owner(exec))+'.json';}
  async read(root,exec) {
    policyFor(this.ctx,root,exec,this.config,false);
    let state;
    try{state=JSON.parse((await readInput(this.ctx,root,this.file(exec),exec.signal,1024*1024)).bytes);}
    catch(e){if(!['ENOENT','FS_NOT_FOUND'].includes(e.code))throw e;}
    const digest=await hostDigest(this.config,root,'learning',this.owner(exec));
    if(!state){
      if(digest)throw new Error('OFFICE_LEARNING_STATE_MISSING');
      return {owner:this.owner(exec),revision:0,candidates:[],corrections:[],saved:[],finish:{state:'pending',evaluations:0}};
    }
    if(state.owner!==this.owner(exec)||sha256(JSON.stringify(state))!==digest)throw new Error('OFFICE_LEARNING_STATE_CHANGED');
    return state;
  }
  async update(root,exec,fn) {
    const key=resolve(root)+':'+this.owner(exec),prior=this.pending.get(key)??Promise.resolve();
    const task=prior.then(()=>this.applyUpdate(root,exec,fn));
    // Return the actual rejection to its caller; the queue tail merely releases
    // later actions so a reported failure can be repaired.
    const tail=task.then(()=>{},()=>{});this.pending.set(key,tail);
    try{return await task;}finally{if(this.pending.get(key)===tail)this.pending.delete(key);}
  }
  async applyUpdate(root,exec,fn) {
    policyFor(this.ctx,root,exec,this.config,true);
    const folder=resolve(root,'.office-plugin/learning');await safeDirectory(root,folder);
    const lockPath=resolve(folder,sha256(this.owner(exec))+'.lock'),lock=await open(lockPath,'wx');
    try{
      const state=await this.read(root,exec),result=await fn(state);state.revision++;
      if(Buffer.byteLength(JSON.stringify(state))>1024*1024)throw new Error('OFFICE_LEARNING_STATE_LIMIT');
      if(exec.agent){
        await this.ctx.get('sessions').flush(exec.agent.session);
      }
      const temp=resolve(folder,randomUUID()+'.tmp');
      await writeFile(temp,JSON.stringify(state,null,2),{flag:'wx'});await rename(temp,resolve(root,this.file(exec)));
      await writeHostDigest(this.config,root,'learning',this.owner(exec),sha256(JSON.stringify(state)));
      return result??state;
    }finally{await lock.close();await unlink(lockPath);}
  }
  async record(root,exec,run) {
    await this.update(root,exec,state=>{
      if(state.candidates.length>=128)throw new Error('OFFICE_LEARNING_CANDIDATE_LIMIT');
      state.candidates.push({run_id:run.id,state:run.state,method_id:run.method_id,code_sha256:run.bundle[0]?.sha256??null,
        published:run.outputs.some(o=>o.state==='published'),created_at:run.created_at});
      if(run.state==='cancelled')state.finish.state='deferred';
    });
  }
  async beforeRun(root,exec) {
    const state=await this.read(root,exec);
    if(state.candidates.length>=128)throw new Error('OFFICE_LEARNING_CANDIDATE_LIMIT');
  }
  async correction(root,exec,args) {
    if(!exec.agent)throw new Error('OFFICE_LEARNING_USER_SESSION_REQUIRED');
    if(!['task_data','method_rule'].includes(args.scope))throw new Error('OFFICE_LEARNING_CORRECTION_SCOPE');
    const message=exec.agent.session.log.find(e=>e.type==='user/message'&&e.data.id===args.user_message_id&&e.data.source.kind==='user');
    if(!message)throw new Error('OFFICE_LEARNING_USER_MESSAGE_REQUIRED');
    const text=message.data.content.filter(b=>b.type==='text').map(b=>b.text).join('\n');
    if(!text.trim()||text.length>8000)throw new Error('OFFICE_LEARNING_USER_TEXT_LIMIT');
    const run=await loadNativeRun(this.ctx,this.config,root,args.run_id,exec);
    return this.update(root,exec,state=>{
      const existing=state.corrections.find(c=>c.user_message_id===args.user_message_id&&c.run_id===run.id);
      if(existing){if(existing.scope!==args.scope)throw new Error('OFFICE_LEARNING_CORRECTION_SCOPE_CHANGED');return existing;}
      const item={id:randomUUID(),run_id:run.id,method_id:run.method_id,scope:args.scope,user_message_id:message.data.id,
        user_event_seq:message.seq,user_text:text,input_sha256:run.inputs.map(f=>f.sha256),
        code_sha256:run.bundle[0]?.sha256??null,rule_sha256:ruleHash(run),state:'pending'};
      state.corrections.push(item);return item;
    });
  }
  async resolveCorrection(root,exec,args) {
    const replacement=await loadNativeRun(this.ctx,this.config,root,args.replacement_run_id,exec);
    if(replacement.state!=='checked')throw new Error('OFFICE_LEARNING_REPAIR_CHECK_REQUIRED');
    return this.update(root,exec,state=>{
      const item=state.corrections.find(c=>c.id===args.correction_id);
      if(!item)throw new Error('OFFICE_LEARNING_CORRECTION_UNKNOWN');
      if(replacement.id===item.run_id)throw new Error('OFFICE_LEARNING_NEW_RUN_REQUIRED');
      if(item.state==='resolved'&&item.replacement_run_id!==replacement.id)throw new Error('OFFICE_LEARNING_CORRECTION_ALREADY_RESOLVED');
      if(item.scope==='task_data'&&(replacement.bundle[0]?.sha256!==item.code_sha256||
          ruleHash(replacement)!==item.rule_sha256))
        throw new Error('OFFICE_LEARNING_TASK_DATA_RULE_CHANGED: 本次资料更正沿用原程序与通用规则');
      if(item.scope==='task_data'&&JSON.stringify(replacement.inputs.map(f=>f.sha256))===JSON.stringify(item.input_sha256))
        throw new Error('OFFICE_LEARNING_TASK_DATA_UNCHANGED: 更正资料后重新执行');
      item.state='resolved';item.replacement_run_id=replacement.id;item.replacement_code_sha256=replacement.bundle[0].sha256;
      return item;
    });
  }
  async beforeSave(root,exec,args,run) {
    const state=await this.read(root,exec);
    const existing=state.saved.find(s=>s.run_id===run.id);
    if(existing){if(existing.name!==args.name||existing.purpose!==args.purpose||(existing.parent_id??null)!==(args.parent_id??null))
      throw new Error('OFFICE_LEARNING_SAVE_ARGUMENT_CHANGED');return existing;}
    if(state.corrections.some(c=>c.state==='pending'))throw new Error('OFFICE_LEARNING_CORRECTION_PENDING: 先修正本次文件并关联通过检查的新执行');
    const correction=state.corrections.find(c=>c.replacement_run_id===run.id);
    if(correction?.scope==='method_rule'&&correction.method_id&&args.parent_id!==correction.method_id)
      throw new Error('OFFICE_LEARNING_PARENT_REQUIRED: 通用修订应接原方法的新版本');
    const all=await this.store.list();
    for(const row of all.filter(m=>m.execution_kind==='native_python')) {
      const {meta}=await this.store.read(row.id);
      if(meta.execution.origin_run_id!==run.id)continue;
      if(meta.name!==args.name||meta.purpose!==args.purpose||(meta.parent_id??null)!==(args.parent_id??null))
        throw new Error('OFFICE_LEARNING_SAVE_ARGUMENT_CHANGED');
      if(!meta.proofs.at(-1)?.complete)throw new Error('OFFICE_LEARNING_PARTIAL_SAVE: 方法已保留草稿，先补验证');
      return {method_id:meta.id,restored:true};
    }
    const same=all.filter(m=>m.execution_kind==='native_python'&&m.status==='active'&&
      (m.name===args.name||m.purpose===args.purpose));
    if(!args.parent_id&&same.length)throw new Error('OFFICE_LEARNING_DUPLICATE_PURPOSE: 同用途方法 '+same.map(m=>m.id).join(',')+'；复用或传parent_id修订');
    return null;
  }
  async saved(root,exec,args,meta,elapsed) {
    await this.update(root,exec,state=>{
      state.saved.push({run_id:args.run_id,method_id:meta.id,parent_id:args.parent_id??null,name:args.name,purpose:args.purpose,
        version:meta.version,code_sha256:meta.code_sha256,save_ms:elapsed});
    });
  }
  async finish(root,exec,{defer=false}={}) {
    const start=performance.now();
    return this.update(root,exec,state=>{
      if(defer){state.finish.state='deferred';return state.finish;}
      if(state.finish.evaluations)return state.finish;
      state.finish={state:'evaluated',evaluations:1,pending_corrections:state.corrections.filter(c=>c.state==='pending').map(c=>c.id),
        candidates:state.candidates.filter(c=>!state.saved.some(s=>s.run_id===c.run_id)).map(c=>({run_id:c.run_id,state:c.state})),
        assessment_ms:performance.now()-start};return state.finish;
    });
  }
  install() {
    const parameter=description=>({type:'string',required:true,description});
    const register=(name,description,parameters,fn)=>this.ctx.tools.register(defineTool({name,description,parameters,parallelSafe:false,
      output:{schema:{type:'json'},render:(_a,v)=>[{type:'text',text:JSON.stringify(v)}]},
      execute:(args,exec)=>fn(workspaceFor(this.ctx,this.configuredRoot,exec,this.config),exec,args)}));
    register('office_learning_status','查看当前会话的实际执行候选、用户纠错、保存版本和待整理状态。任务资料留在本工作区。',{},async(root,exec)=>{
      const state=await this.read(root,exec);
      const user_messages=(exec.agent?.session.log??[]).filter(e=>e.type==='user/message'&&e.data.source.kind==='user').slice(-10)
        .map(e=>({id:e.data.id,seq:e.seq,text:e.data.content.filter(b=>b.type==='text').map(b=>b.text).join('\n').slice(0,1000)}));
      return {...state,user_messages,record_path:this.file(exec)};
    });
    register('office_learning_correct','关联当前会话真实用户原话与执行记录。单次资料更正用task_data，通用计算或程序修订用method_rule；失败执行同样可记录。',{
      run_id:parameter('纠错针对的执行ID'),user_message_id:parameter('当前真实用户消息ID'),scope:parameter('task_data或method_rule')},
      (root,exec,args)=>this.correction(root,exec,args));
    register('office_learning_resolve','用独立文件检查通过的新执行关联纠错结果；单次资料修正须保持通用程序和规则一致。',{
      correction_id:parameter('本会话纠错ID'),replacement_run_id:parameter('修正后通过检查的执行ID')},
      (root,exec,args)=>this.resolveCorrection(root,exec,args));
    register('office_learning_finish','当前会话最多评估一次候选及待处理纠错。defer=true只保存待续接状态，不运行程序、不保存或启用方法。',{
      defer:{type:'boolean'}},(root,exec,args)=>this.finish(root,exec,args));
    this.ctx.on('agent/turn-stopping',async({agent,signal})=>{
      const exec={agent,signal},root=workspaceFor(this.ctx,this.configuredRoot,exec,this.config),state=await this.read(root,exec);
      const requested=agent.session.log.some(e=>e.type==='user/message'&&e.data.source.kind==='user'&&
        e.data.content.some(b=>b.type==='text'&&/(保存.{0,12}(办法|方法)|可复用|经常.{0,8}(做|用))/.test(b.text)));
      if(!requested||!state.candidates.length||state.finish.evaluations||state.finish.state==='deferred')return;
      const result=await this.finish(root,exec);
      const latest=state.candidates.findLast(c=>['executed','checked','check_failed'].includes(c.state));
      if(!state.corrections.some(c=>c.state==='pending')&&latest?.state==='checked'&&
          (latest.method_id||state.saved.some(s=>s.run_id===latest.run_id)))return;
      agent.inject(createUserMessage({source:{kind:'office-learning'},content:[{type:'text',text:
        '宿主会话整理（仅一次）：'+JSON.stringify(result)+'。本会话已发布新文件'+state.candidates.filter(c=>c.published).length+'份。查看office_learning_status；有待纠错先处理；有重复用途优先修订。适合复用的已核对程序使用office_program_save并可提供validation_summary一次保存启用；不能保存失败或未经独立核对的候选。'}]}));
    });
    this.ctx.on('agent/status',async({agent,status})=>{
      if(status!=='idle')return;
      const end=agent.session.log.findLast(e=>e.type==='turn/end');
      if(!end||!['aborted','interrupted'].includes(end.data.reason?.kind))return;
      const exec={agent,signal:new AbortController().signal},root=workspaceFor(this.ctx,this.configuredRoot,exec,this.config);
      const state=await this.read(root,exec);
      if(state.candidates.length&&state.finish.state!=='deferred')await this.finish(root,exec,{defer:true});
    });
  }
}
