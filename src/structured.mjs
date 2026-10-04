import { defineTool } from '@deepseek-ai/dsh-tools';
import { randomUUID } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { readInput, runWorker, policyFor, sha256, workspaceFor } from './runtime.mjs';
import { presentOfficeResult } from './presentation.mjs';
import { selectMethods, methodDetails, userMethodDetails } from './methods.mjs';

const str = description => ({ type: 'string', required: true, description });
const instruction = '只使用当前来源；结构关系是证据，不是已确认业务事实。不得编造缺值，不确定主体、单位或位置时核实，仍不确定就列出需问内容。已知不支持直接采用原生工具或临时脚本，不反复尝试插件。';
export function installStructured(ctx, config, configuredRoot) {
  const parserFingerprint=()=>sha256(Buffer.concat(['office_worker.py','structure_docx.py','structure_xlsx.py','source_xlsx.py',
    'structured_worker.py','structured_fill.py','docx_prepare.py','xlsx_prepare.py','xlsx_package_patch.py','package_writer.py']
    .map(file=>readFileSync(join(dirname(config.workerPath),file)))
    .concat(readFileSync(join(dirname(config.workerPath),'../vendor_candidates/value_fit.py')))));
  const reads = new Map();
  const cache = new Map();
  const presentedSources = new Map();
  const owner = exec => exec.agent?.session?.id ?? 'agentless-probe';
  const register = (name, description, parameters, execute) => ctx.tools.register(defineTool({
    name, description: instruction + description, parameters, timeoutMs: 65000,
    output: { schema: { type: 'json' }, render: (_a,v) => [{ type: 'text', text: JSON.stringify(presentOfficeResult(v)) }] },
    execute:(args,exec)=>execute(args,exec,workspaceFor(ctx,configuredRoot,exec,config)),
  }));
  const get = (id,exec) => {
    const record = reads.get(id);
    if (!record || record.owner!==owner(exec)) throw new Error('OFFICE_ID_UNKNOWN');
    if (record.root!==workspaceFor(ctx,configuredRoot,exec,config)) throw new Error('OFFICE_WORKSPACE_CHANGED: 请在当前工作区重新读取');
    if (Date.now()-record.created>600000) { reads.delete(id); throw new Error('OFFICE_ID_EXPIRED'); }
    return record;
  };
  const request = (record,operation) => ({
    operation, sources: record.sources.map(s => ({ format:s.format, bytes_b64:s.bytes.toString('base64'),scope:s.scope })),
    target: record.target ? {format:record.target.format,bytes_b64:record.target.bytes.toString('base64'),scope:record.target.scope} : null,
    requested_fields:record.fields,source_offset:record.offset,target_mode:record.targetMode,
  });
  const unchanged = async (record,exec) => {
    const root=workspaceFor(ctx,configuredRoot,exec,config);
    if (record.parser!==parserFingerprint()) throw new Error('OFFICE_PARSER_CHANGED: 解析代码版本变化，请重新读取');
    for (const s of [...record.sources, ...(record.target ? [record.target] : [])]) {
      if ((await readInput(ctx,root,s.path,exec.signal)).hash!==s.hash) throw new Error('OFFICE_INPUT_CHANGED: 重新读取，不能沿用旧版本');
    }
  };
  const open = async (path,scope,exec) => {
    const root=workspaceFor(ctx,configuredRoot,exec,config);
    const format = path.split('.').at(-1).toLowerCase();
    if (!['docx','xlsx'].includes(format)) throw new Error('OFFICE_FORMAT_UNSUPPORTED');
    return {...await readInput(ctx,root,path,exec.signal),format,scope};
  };
  register('office_read','读取1至8个Word/Excel来源及可选目标。不传目标可仅解析。来源只读不会因绘图或计算链整份拒绝，图片内容仍未读。requested_fields检索字段及附近物理格减少返回；超过1000条明确返回next_offset，source_offset续读，不把本页当全部。target_mode=review_values可核对已有值；更新必须提交原文expected_target_text。金额和主体仍须有证据，检索命中不代表匹配。默认不读数据库。同会话、工作区、文件内容、范围、选项及解析器版本的重复读取默认显示摘要；需再看完整原文时传expand_reuse=true。', {
    source_paths: {type:'array',required:true,items:{type:'string'},description:'来源相对路径；只看目标时传空数组[]并提供target_path。读取有明确字段的报告时优先requested_fields检索及上下文，需要时再补读'},
    target_path: {type:'string',description:'可选目标模板的工作区相对路径'},
    requested_fields:{type:'array',items:{type:'string'},description:'可选1至32个Excel来源检索词或明确别名；返回上下文，不自动确认键值。'},
    source_offset:{type:'number',description:'Excel来源的0起始页偏移；超过1000条时按next_offset续读，建议分别读取各来源'},
    target_mode:{type:'string',description:'fill_blanks默认仅填空；review_values允许在新副本核对更新旧值，须逐项引用旧原文'},
    expand_reuse:{type:'boolean',description:'重复读取命中摘要时传true，要求模型展示完整原文；不影响host收到的完整结构化结果。'},
    source_scopes: {type:'array',items:{type:'object',additionalProperties:false,properties:{source_index:{type:'number',required:true,description:'来源从0开始的整数索引'},sheet:str('工作表名称'),range:str('A1:B10形式范围')}}},
    target_scope: {type:'object',additionalProperties:false,properties:{sheet:str('目标工作表'),range:str('目标范围')}},
  }, async (args,exec,root) => {
    policyFor(ctx,root,exec,config,false);
    for (const [id,r] of reads) if (Date.now()-r.created>600000) reads.delete(id);
    if (reads.size>=32 || !Array.isArray(args.source_paths) || !args.source_paths.length&&!args.target_path || args.source_paths.length>8) throw new Error('OFFICE_READ_LIMIT');
    const scopes=new Map();
    for (const scope of args.source_scopes??[]) {
      if (!Number.isInteger(scope.source_index) || scope.source_index<0 || scope.source_index>=args.source_paths.length || scopes.has(scope.source_index)) throw new Error('OFFICE_SCOPE_INDEX');
      scopes.set(scope.source_index,{sheet:scope.sheet,range:scope.range});
    }
    const sources = [];
    let totalBytes=0;
    for (const [i,path] of args.source_paths.entries()) {
      const s=await open(path,scopes.get(i),exec);totalBytes+=s.bytes.length;
      if(totalBytes>20*1024*1024)throw new Error('OFFICE_TOTAL_INPUT_LIMIT: 合计超过20MiB，请分批或使用原生范围读取');
      sources.push(s);
    }
    const target = args.target_path ? await open(args.target_path,args.target_scope,exec) : null;
    if(totalBytes+(target?.bytes.length??0)>20*1024*1024)throw new Error('OFFICE_TOTAL_INPUT_LIMIT: 来源与模板合计超过20MiB');
    const record = {root,sources,target,fields:args.requested_fields,offset:args.source_offset??0,targetMode:args.target_mode??'fill_blanks',created:Date.now(),owner:owner(exec),state:'ready',parser:parserFingerprint()};
    const key=JSON.stringify({owner:record.owner,root,parser:record.parser,
      sources:sources.map(s=>[s.path,s.hash,s.scope]),target:target?[target.path,target.hash,target.scope]:null,
      fields:record.fields,offset:record.offset,targetMode:record.targetMode});
    for (const [k,v] of cache) if (Date.now()-v.created>600000) cache.delete(k);
    const cached=config.readCacheEnabled!==false ? cache.get(key) : null;
    const result=cached ? structuredClone(cached.result) : await runWorker(ctx,config,root,request(record,'inspect_v3'),exec);
    if (JSON.stringify(result.source_sha256)!==JSON.stringify(sources.map(s=>s.hash)) || result.target_sha256!==(target?.hash ?? null)) throw new Error('OFFICE_DIGEST_MISMATCH');
    record.inspection=result;
    const id=randomUUID();
    if (!cached && config.readCacheEnabled!==false) {
      if (cache.size>=32) cache.delete(cache.keys().next().value);
      cache.set(key,{created:Date.now(),result:structuredClone(result),first_read_id:id});
    }
    reads.set(id,record);
    // Different field/range requests can repeat the same physical evidence. Only
    // suppress a record when both its current file bytes and full context match.
    const evidenceKey=JSON.stringify({owner:record.owner,root,parser:record.parser,
      sources:sources.map(s=>[s.path,s.hash])});
    for(const [k,v] of presentedSources)if(Date.now()-v.created>600000)presentedSources.delete(k);
    if(!presentedSources.has(evidenceKey)){
      if(presentedSources.size>=32)presentedSources.delete(presentedSources.keys().next().value);
      presentedSources.set(evidenceKey,{created:Date.now(),records:new Map()});
    }
    const seen=presentedSources.get(evidenceKey).records,previous=[];
    for(const [candidateIndex,candidate] of result.source_candidates.entries()){
      const digest=sha256(JSON.stringify(candidate)),prior=seen.get(digest);
      if(prior)previous.push({source_ref:candidate.source_ref,read_id:prior,candidate_index:candidateIndex});
      else if(seen.size<32000)seen.set(digest,id);
    }
    return {...result,read_id:id,expires_in_seconds:600,methods:selectMethods(result,config.methodsEnabled!==false,config.disabledMethodIds??[]),
      read_reuse:{used:!!cached,expanded:Boolean(args.expand_reuse),scope:'同会话、同工作区、同内容摘要、同范围及解析器版本；每次仍核对当前文件',
        expand_parameter:'expand_reuse=true',first_read_id:cached?.first_read_id ?? id,
        previous_source_records:previous,
        source_sha256:result.source_sha256,target_sha256:result.target_sha256,
        readonly_expansion:config.recordsTools !== false ? 'office_records(read_id=本次read_id)' : null}};
  });
  register('office_fill','提交本次来源与目标的对应引用。一次完成原值重取、明确单位换算、冲突与位置检查、新文件写出及重开。review_values更新已有值须提交当前完整expected_target_text。不接受模型value，未映射字段不会自动填写；仅对applied报告完成。已有输出不覆盖。扩行、派生计算和未知宏逻辑交原生脚本，之后必须核对最终产物。', {
    read_id:str('office_read返回的标识'),output_name:str('新文件名，无路径；与目标后缀相同'),
    mappings:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{
      source_ref:str('当前source_ref'),target_ref:str('当前target_ref'),expected_target_text:{type:'string',description:'仅核对更新已有值时必须提交office_read所返回旧text；不得提交新value'}}}},
  }, async (args,exec,root) => {
    policyFor(ctx,root,exec,config,true);
    const record=get(args.read_id,exec);
    if (!record.target) throw new Error('OFFICE_TARGET_REQUIRED');
    if (typeof args.output_name!=='string' || !args.output_name.length || args.output_name.length>120 || args.output_name!==args.output_name.trim() ||
      /[\\/:*?"<>|\u0000-\u001f]/.test(args.output_name) || !args.output_name.toLowerCase().endsWith('.'+record.target.format) ||
      /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$/i.test(args.output_name.split('.')[0])) throw new Error('OFFICE_OUTPUT_NAME');
    if (!Array.isArray(args.mappings) || !args.mappings.length || args.mappings.length>100) throw new Error('OFFICE_MAPPING_LIMIT');
    const sources=new Set(record.inspection.source_candidates.map(x=>x.source_ref));
    const targets=new Set(record.inspection.target_slots.map(x=>x.target_ref));
    for (const m of args.mappings) if (Object.keys(m).some(k=>!['source_ref','target_ref','expected_target_text'].includes(k)) || !sources.has(m.source_ref) || !targets.has(m.target_ref)) throw new Error('OFFICE_REF_UNKNOWN');
    if (record.state!=='ready') throw new Error('OFFICE_PLAN_CONSUMED');
    await unchanged(record,exec);
    record.state='executing';
    if (!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);
    const result=await runWorker(ctx,config,root,{...request(record,'fill_v3'),mappings:args.mappings,output_name:args.output_name},exec);
    const file=await readInput(ctx,root,args.output_name,exec.signal,15*1024*1024);
    if (file.hash!==result.output_sha256) throw new Error('OFFICE_OUTPUT_DIGEST_MISMATCH: 已有产物待查，未报告成功');
    await unchanged(record,exec);
    if (!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);
    record.state='done';
    return {...result,output_path:args.output_name,unmapped_slots:targets.size-args.mappings.length,
      completion_scope:'仅已应用位置；空位总数不等于业务必填总数',methods_enabled:config.methodsEnabled!==false};
  });
  if (config.recordsTools !== false) {
  register('office_records','只读查看本次来源记录和方法条件，或查询本工作区历史快照；历史不表示当前值，不恢复计划和授权。', {
    read_id:{type:'string',description:'可选office_read标识；不传时仅查看方法'},
    snapshot_id:{type:'string',description:'可选保存快照标识，查询历史记录而非最新来源'},
    list_saved:{type:'boolean',description:'只读列出本工作区保存记录'},
  },async (args,exec,root) => {
    policyFor(ctx,root,exec,config,false);
    const methods=[...methodDetails(),...userMethodDetails()];
    if (args.snapshot_id || args.list_saved) return {...await runWorker(ctx,config,root,{operation:'records_view',snapshot_id:args.snapshot_id},exec),methods};
    if (!args.read_id) return {methods,global_user_directory:'插件固定library/user；仅宿主手动导入草稿，不允许模型任意写全局目录'};
    const record=get(args.read_id,exec);
    await unchanged(record,exec);
    return {...record.inspection,methods,storage:'本次10分钟版本记录；不是持久数据库'};
  });
  register('office_save_records','可选地把本次原文和位置按版本存入工作区数据库，另存待确认解释。写入仍经原生权限；不能声称已学会、已确认或已启用通用方法。首次填写不依赖保存。', {
    read_id:str('本次读取标识'),
    proposal:{type:'object',additionalProperties:false,properties:{source_ref:str('本次来源引用'),explanation:str('待确认解释/纠错建议；不自动作为事实')}},
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,true);
    const record=get(args.read_id,exec);
    await unchanged(record,exec);
    if (!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);
    const result=await runWorker(ctx,config,root,{...request(record,'records_store'),proposal:args.proposal,
      paths:{sources:record.sources.map(s=>s.path),target:record.target?.path??null}},exec);
    if (!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);
    return result;
  });
  }
}

