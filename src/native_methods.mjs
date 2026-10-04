import { defineTool } from '@deepseek-ai/dsh-tools';
import { readFile } from 'node:fs/promises';
import { resolve, dirname } from 'node:path';
import { executeNative, loadNativeRun } from './native_execution.mjs';
import { MethodStore } from './method_store.mjs';
import { workspaceFor, policyFor, readInput, sha256 } from './runtime.mjs';
import { NativeLearning } from './native_learning.mjs';
import { performance } from 'node:perf_hooks';
import { resumeGroup } from './native_group_resume.mjs';
import { installOrdinaryCapture } from './ordinary_execution_capture.mjs';
import { checkOrdinaryCandidate } from './ordinary_candidate_check.mjs';
import { searchPrograms } from './program_search.mjs';

const string=description=>({type:'string',required:true,description});
const brief=run=>JSON.parse(JSON.stringify({run_id:run.id,state:run.state,candidate:run.candidate,method_id:run.method_id,
  code_sha256:run.bundle[0]?.sha256,code_validation:run.code_validation,
  file_check:{complete:run.file_check.complete,checked:run.file_check.checked,passed:run.file_check.passed},
  outputs:run.outputs,error:run.error,runtime_fingerprint:run.runtime_fingerprint,
  relationship_binding:run.relationship_binding,
  delivery_manifest:run.delivery_manifest,
  record_path:`.office-plugin/native-runs/${run.id}/run.json`}));
export function installNativeMethods(ctx,config,configuredRoot) {
  const store=new MethodStore(config.methodDirectory);
  const learning=config.learningEnabled?new NativeLearning(ctx,config,configuredRoot,store):null;
  learning?.install();
  if(learning&&config.captureOrdinaryPython===true)installOrdinaryCapture(ctx,config,configuredRoot,learning);
  const tracked=async(root,exec,run)=>{
    if(learning)await learning.record(root,exec.signal.aborted?{...exec,signal:new AbortController().signal}:exec,run);
    return brief(run);
  };
  const register=(name,description,parameters,fn)=>ctx.tools.register(defineTool({name,description,parameters,
    timeoutMs:190000,parallelSafe:false,output:{schema:{type:'json'},render:(_a,v)=>[{type:'text',text:JSON.stringify(v)}]},
    execute:(args,exec)=>fn(args,exec,workspaceFor(ctx,configuredRoot,exec,config))}));
  const common={task:string('本次业务范围'),cwd:string('当前工作区相对运行目录，根目录写.'),
    argv:{type:'array',required:true,items:{type:'string'},description:'独立参数数组，不是shell语句；来源参数填写inputs中的相对路径'},
    inputs:{type:'array',items:{type:'string'},description:'默认从check的来源和模板产生；另有数据依赖时显式列本次全部输入'},
    outputs:{type:'json',required:true,description:'1至16份产物[{path,argv_index}]；整组放同一新子目录，全部核对后发布'},
    check:{type:'json',required:true,description:'单文件{source_paths,target_path,contract}；整组{source_paths,artifacts:[{output_index,target_path,contract,conditions}],consistency:[{kind,refs:[{output_index,target_ref}]}]}'},
    timeout_ms:{type:'number',description:'单次执行时限，1至60000毫秒'}};
  if(learning&&config.captureOrdinaryPython===true)register('office_program_check_candidate','核对DSH自动保留的普通Python程序和已生成产物，不再执行业务程序。输入、程序与产物字节须与执行时一致，完整检查通过后可保存实际程序。',{
    run_id:string('普通执行后DSH返回的执行ID'),outputs:common.outputs,check:common.check,
    conditions:{type:'json',description:'固定位置清单的来源和模板结构条件'},
    relationship_rule:{type:'json',description:'单Excel长期标签规则office_relationship_rule_v1；只存业务关系，本批位置由DSH重新确定'},
    relationship_arg_index:{type:'number',description:'原执行argv里可接收本批关系JSON的参数下标，不得占来源、模板或输出参数'}
  },async(args,exec,root)=>{
    const value=await checkOrdinaryCandidate(ctx,config,root,args,exec);
    await learning.update(root,exec,state=>{const c=state.candidates.find(c=>c.run_id===value.run_id);if(c){c.state=value.state;c.published=value.state==='checked';}});
    return value;
  });
  register('office_program_run','登记并运行明确Python入口：先保存真实副本和依赖，再沿用DSH执行器。执行立即留候选；独立文件检查通过才发布新文件。首期个人兼容配置，普通脚本可直接读写文件，Windows部分约束按实际记录。',{
    interpreter:string('配置中已允许的Python解释器绝对路径'),script_path:string('工作区单个Python入口相对路径'),
    conditions:{type:'json',description:'固定位置方法需填写；长期规则方法从本批定位结果产生结构条件'},
    relationship_rule:{type:'json',description:'长期规则：单Excel用office_relationship_rule_v1；多Excel/Word用office_group_relationship_rule_v1；标签右邻和固定文字组合用office_group_composition_rule_v1。check只给source_paths和模板路径，DSH主程序生成本批位置、清单及整组关系'},
    record_selection:{type:'json',description:'多文件规则的本批记录选择{name:[主体、合同等键值]}；只保留在本任务，复用时重新提供'},
    relationship_arg_index:{type:'number',description:'可选：argv中用于传本批处理关系JSON的下标，程序从同一check清单生成实际位置；不得占用来源或输出参数'},
    dependencies:{type:'array',items:{type:'string'},description:'明确的本地Python依赖相对路径；首期仅同目录模块，不支持本地包或动态加载'},...common,
  },async(args,exec,root)=>{await learning?.beforeRun(root,exec);return tracked(root,exec,await executeNative(ctx,config,root,args,exec));});
  register('office_program_save','从已检查执行记录保存同一份实际程序，不再把普通脚本改写成回调。输入或产物后来变化时重新核对；参数化或代码修改应先重新run，再保存新版本。',{
    run_id:string('真实执行记录ID'),name:string('可重复用途'),purpose:string('处理步骤及范围'),
    parent_id:{type:'string',description:'修改旧方法时传旧ID，新版本重新验证'},
    validation_summary:{type:'string',description:'可选：实测范围与启用理由；提供后同一保存操作完成验证和启用'},
  },async(args,exec,root)=>{
    const started=performance.now();
    policyFor(ctx,root,exec,config,true);
    if(!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
    if(args.validation_summary!==undefined&&(!args.validation_summary.trim()||args.validation_summary.length>2000))
      throw new Error('OFFICE_NATIVE_SUMMARY_REQUIRED');
    const run=await loadNativeRun(ctx,config,root,args.run_id,exec);
    if(run.state!=='checked'||!run.code_validation.passed||!run.file_check.complete)throw new Error('OFFICE_NATIVE_UNVERIFIED');
    for(const f of run.inputs)if((await readInput(ctx,root,f.path,exec.signal,15*1024*1024)).hash!==f.sha256)throw new Error('OFFICE_NATIVE_INPUT_CHANGED');
    for(const output of run.outputs)if(output.state!=='published'||(await readInput(ctx,root,output.path,exec.signal,15*1024*1024)).hash!==output.sha256)
      throw new Error('OFFICE_NATIVE_ARTIFACT_CHANGED');
    const contract=run.file_check.confirmed_contract;
    // Reuse host-owned independent proof only for the exact same current bytes.
    // The live input/output reads above verify the already checked versions;
    // a script's own success or expected-value declaration is never a proof.
    const inputHashes=new Map(run.inputs.map(f=>[f.path,f.sha256])),checked=run.file_check;
    const proofs=run.group_checks?checked.artifacts.map((c,i)=>({...c,source_paths:run.group_checks.source_paths,output:run.outputs[i]})):
      [{...checked,output:run.outputs[0]}];
    if(proofs.some(c=>c.artifact_sha256!==c.output.sha256||c.target_sha256!==inputHashes.get(c.target_path)||
      JSON.stringify(c.source_sha256)!==JSON.stringify(c.source_paths.map(p=>inputHashes.get(p)))))
      throw new Error('OFFICE_NATIVE_INDEPENDENT_PROOF_MISMATCH');
    for(const [name,hash] of run.environment.trusted_worker_versions??[]) {
      if(!/^[a-zA-Z0-9_-]+\.py$/.test(name)||sha256(await readFile(resolve(dirname(config.nativeIdentityPath),name)))!==hash)
        throw new Error('OFFICE_NATIVE_REVALIDATE: 独立检查程序版本已变化');
    }
    const files=[];
    for(const f of run.bundle) {
      const item=await readInput(ctx,root,f.snapshot,exec.signal,65536);
      if(item.hash!==f.sha256)throw new Error('OFFICE_NATIVE_CODE_CHANGED');
      files.push({...f,bytes:item.bytes});
    }
    const ruleGroup=['office_group_relationship_rule_v1','office_group_composition_rule_v1'].includes(run.relationship_rule?.schema);
    const {cwd:batchEnvironmentCwd,...portableEnvironment}=run.environment;
    const execution={kind:'native_python',entry:files[0].path,interpreter:run.interpreter,
      bundle:files.map(f=>({path:f.path,sha256:f.sha256})),environment:ruleGroup?portableEnvironment:run.environment,
      ...(ruleGroup?{}:{argument_example:run.argv,cwd_example:run.cwd}),origin_run_id:run.id,
      ...(run.relationship_rule===undefined?{}:{relationship_rule:run.relationship_rule}),
      ...(run.relationship_arg_index===undefined?{}:{relationship_arg_index:run.relationship_arg_index})};
    const code=files[0].bytes.toString('utf8');
    if(sha256(code)!==files[0].sha256)throw new Error('OFFICE_NATIVE_UTF8_REQUIRED');
    const previous=learning?await learning.beforeSave(root,exec,args,run):null;
    if(previous){
      let {meta}=await store.read(previous.method_id);
      if(meta.status==='draft'&&args.validation_summary)meta=await store.activate(meta,args.validation_summary);
      if(previous.restored)await learning.saved(root,exec,args,meta,performance.now()-started);
      return {method_id:meta.id,status:meta.status,version:meta.version,already_saved:true};
    }
    const meta=await store.create({...args,code,conditions:run.conditions,contract,execution,dependencies:files.slice(1)});
    const proof={complete:true,checked:checked.checked,passed:checked.passed,tested_at:new Date().toISOString(),
      code_sha256:meta.code_sha256,runtime_fingerprint:run.runtime_fingerprint,origin_run_id:run.id,
      source_sha256:run.inputs.map(f=>f.sha256),artifact_sha256:run.outputs[0].sha256,
      ...(run.group_checks?{artifacts:run.outputs.map((o,i)=>({...(ruleGroup?{output_index:i}:{path:o.path}),sha256:o.sha256})),consistency_checked:checked.consistency.length}:{})};
    let verified=await store.update(meta.id,meta.revision,m=>({...m,proofs:[proof]}));
    if(args.validation_summary)verified=await store.activate(verified,args.validation_summary);
    const save_ms=performance.now()-started;
    if(learning)await learning.saved(root,exec,args,verified,save_ms);
    return {method_id:meta.id,status:verified.status,version:meta.version,save_ms,code_sha256:meta.code_sha256,execution_kind:'native_python',
      proof_reused_from:run.id,independent_proof_reused:true,next:verified.status==='active'?'已启用，下一批可office_program_reuse':'填写实测范围后office_program_enable；修改代码后重新执行检查'};
  });
  register('office_program_enable','启用已通过代码验证和独立文件检查的普通脚本版本；启用不携带跨会话执行权限，运行仍按当前DSH策略。',{
    method_id:string('方法ID'),validation_summary:string('实际通过范围、限制及启用理由'),
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,true);
    if(!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
    const {meta}=await store.read(args.method_id);
    if(meta.execution?.kind!=='native_python'||meta.status==='disabled'||!meta.proofs.at(-1)?.complete||
        meta.proofs.at(-1)?.code_sha256!==meta.code_sha256)throw new Error('OFFICE_NATIVE_UNVERIFIED');
    if(!args.validation_summary.trim()||args.validation_summary.length>2000)throw new Error('OFFICE_NATIVE_SUMMARY_REQUIRED');
    const enabled=await store.activate(meta,args.validation_summary);
    return {method_id:enabled.id,status:enabled.status,code_sha256:enabled.code_sha256};
  });
  register('office_program_reuse','运行已启用普通Python副本。长期规则按当前表头及主体期间键重新定位，固定位置方法沿用保存清单；产物独立核对。执行环境变化或业务规则变化时重新验证版本。',{
    method_id:string('已启用方法ID'),...Object.fromEntries(Object.entries(common).filter(([key])=>key!=='check')),
    source_paths:{type:'array',required:true,items:{type:'string'}},target_path:{type:'string',description:'单文件模板'},
    target_paths:{type:'array',items:{type:'string'},description:'整组模板清单，顺序与outputs相同'},
    record_selection:{type:'json',description:'多文件规则本批选择{name:[匹配键值]}；方法不携带上一批客户值'},
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,true);const saved=await store.read(args.method_id);
    await learning?.beforeRun(root,exec);
    if(saved.meta.status!=='active'||saved.meta.execution?.kind!=='native_python')throw new Error('OFFICE_NATIVE_NOT_ACTIVE');
    const ruleGroup=saved.meta.contract.schema==='office_group_rule_contract_v1';
    const grouped=ruleGroup||saved.meta.contract.schema==='office_group_contract_v1';
    const count=ruleGroup?saved.meta.contract.artifact_count:saved.meta.contract.artifacts?.length;
    if(grouped&&(!Array.isArray(args.target_paths)||args.target_paths.length!==count))throw new Error('OFFICE_NATIVE_GROUP_TARGETS');
    const check=ruleGroup?{source_paths:args.source_paths,artifacts:args.target_paths.map((target_path,output_index)=>({output_index,target_path}))}:
      grouped?{source_paths:args.source_paths,artifacts:saved.meta.contract.artifacts.map((a,i)=>({...a,target_path:args.target_paths[i],
      conditions:saved.meta.conditions.artifacts[i]})),consistency:saved.meta.contract.consistency}:
      {source_paths:args.source_paths,target_path:args.target_path,contract:saved.meta.contract};
    const value=await executeNative(ctx,config,root,{...args,interpreter:saved.meta.execution.interpreter,check},exec,saved);
    if(value.state==='failed'&&value.error==='OFFICE_NATIVE_FILE_CHECK_FAILED')
      await store.update(saved.meta.id,saved.meta.revision,m=>({...m,status:'disabled',failure:value.error}));
    return tracked(root,exec,value);
  });
  register('office_program_resume','恢复已整组核对、发布途中中断的文件清单。重新核对输入、暂存产物和已发布文件，继续发布其余文件；不重跑业务、不覆盖已有文件。',{
    run_id:string('本会话真实执行ID')},async(args,exec,root)=>brief(await resumeGroup(ctx,config,root,args,exec)));
  register('office_program_list','按关键词查找共用方法库中的普通脚本版本及状态，最多返回10条。',{
    query:{type:'string'},
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,false);const rows=await store.list();
    return searchPrograms(rows,args.query);
  });
}
