import { readFile, writeFile, rename, lstat, realpath, open } from 'node:fs/promises';
import { resolve, relative, isAbsolute, dirname } from 'node:path';
import { randomUUID } from 'node:crypto';
import { readInput, policyFor, sha256, quotePwsh, outside, runWorker } from './runtime.mjs';
import { safeDirectory } from './method_execution.mjs';
import { checkArtifact } from './output_tools.mjs';
import { expandContract } from './contract_builder.mjs';
import { performance } from 'node:perf_hooks';
import { taskBinding } from './task_binding.mjs';
import { hostDigest,writeHostDigest } from './native_host_state.mjs';
import { prepareGroup,checkGroup,publishGroup } from './native_group.mjs';
import { bindGroupRule } from './native_group_relationship.mjs';

const cleanSignal=()=>new AbortController().signal;
// Host-owned receipts prevent a business script from editing a workspace JSON
// into "executed/checked" evidence. Restarted pending runs must be revalidated;
// persisted method versions remain in the shared store.
const receipts=new Map();
export const nativePath=value=>{
  if(typeof value!=='string'||!value||isAbsolute(value)||value.includes(':')||/[\x00-\x1f]/.test(value)||value.split(/[\\/]/).some(p=>p==='..'||!p))
    throw new Error('OFFICE_NATIVE_PATH_DENIED');
  return value.replaceAll('\\','/');
};
export async function nativeFolder(root,path) {
  const full=resolve(root,path);
  if(outside(root,full))throw new Error('OFFICE_NATIVE_CWD_OUTSIDE');
  if(!(await lstat(full)).isDirectory()||outside(await realpath(root),await realpath(full)))throw new Error('OFFICE_NATIVE_CWD_LINK');
  // Check every component, including a link that happens to point inside root.
  let current=root;
  for(const part of relative(root,full).split(/[\\/]/).filter(Boolean)) {
    current=resolve(current,part);
    if((await lstat(current)).isSymbolicLink())throw new Error('OFFICE_NATIVE_CWD_LINK');
  }
  return full;
}
async function persist(root,record) {
  const folder=resolve(root,'.office-plugin','native-runs',record.id);await safeDirectory(root,folder);
  const temp=resolve(folder,randomUUID()+'.tmp');
  await writeFile(temp,JSON.stringify(record,null,2),{flag:'wx'});await rename(temp,resolve(folder,'run.json'));
}
export async function sealNativeRun(config,root,record){
  await persist(root,record);const digest=sha256(JSON.stringify(record));
  if(config.nativeStateDirectory)await writeHostDigest(config,root,'run',record.id,digest);
  receipts.set(resolve(root)+':'+record.id,digest);
  if(receipts.size>256)receipts.delete(receipts.keys().next().value);
}
export async function loadNativeRun(ctx,config,root,id,exec) {
  if(!/^[a-f0-9-]{36}$/.test(id))throw new Error('OFFICE_NATIVE_RUN_ID');
  const value=JSON.parse((await readInput(ctx,root,`.office-plugin/native-runs/${id}/run.json`,exec.signal,1024*1024)).bytes);
  if(value.id!==id||value.owner!==(exec.agent?.session?.id??'agentless-probe'))throw new Error('OFFICE_NATIVE_RUN_OWNER');
  const digest=sha256(JSON.stringify(value));
  if(receipts.get(resolve(root)+':'+id)!==digest&&
      (!config.nativeStateDirectory||await hostDigest(config,root,'run',id)!==digest))
    throw new Error('OFFICE_NATIVE_RUN_RECEIPT: 执行记录被改动或宿主收据缺失，需重新执行验证');
  return value;
}
async function hostRun(ctx,config,root,exec,command,stdin,timeoutMs) {
  const policy=policyFor(ctx,root,exec,config,true);
  exec.signal.throwIfAborted();
  const spec=ctx.shell.resolve({command,workdir:exec.cwd,timeoutMs,onExpiry:'kill',signal:exec.signal,
    sandboxPolicy:policy,stdin,stdoutMaxBytes:64*1024,stderrMaxBytes:64*1024});
  const result=await(await ctx.shell.execute(spec)).result();
  return result;
}
function checkHost(result) {
  if(!['partial','enforced'].includes(result.sandbox?.enforcement)||result.sandbox?.denied)
    throw new Error('OFFICE_NATIVE_CONFINEMENT_UNAVAILABLE');
}
export async function executeNative(ctx,config,root,args,exec,saved=null) {
  args={...args,inputs:args.inputs??[...new Set([...(args.check?.source_paths??[]),args.check?.target_path,...(args.check?.artifacts??[]).map(c=>c.target_path)].filter(Boolean))]};
  policyFor(ctx,root,exec,config,true);
  if(config.nativeScriptMode!=='personal-compatible')throw new Error('OFFICE_NATIVE_MODE_DISABLED: 当前配置未允许普通脚本');
  if(!isAbsolute(args.interpreter)||resolve(args.interpreter)!==resolve(config.pythonPath))throw new Error('OFFICE_NATIVE_INTERPRETER_DENIED');
  const cwd=await nativeFolder(root,args.cwd==='.'?'.':nativePath(args.cwd));
  if(!Array.isArray(args.argv)||args.argv.length>64||args.argv.some(a=>typeof a!=='string'||a.length>4096||/[\x00\r\n]/.test(a)))throw new Error('OFFICE_NATIVE_ARGV');
  const timeout=args.timeout_ms??60000;
  if(!Number.isInteger(timeout)||timeout<1||timeout>60000)throw new Error('OFFICE_NATIVE_TIMEOUT');
  if(!args.task||typeof args.task!=='string'||args.task.length>2000)throw new Error('OFFICE_NATIVE_TASK_REQUIRED');
  if(!Array.isArray(args.inputs)||!args.inputs.length||args.inputs.length>16||!Array.isArray(args.outputs)||!args.outputs.length||args.outputs.length>16)
    throw new Error('OFFICE_NATIVE_FILES: 需1至16份产物和明确输入清单');
  const record={id:randomUUID(),owner:exec.agent?.session?.id??'agentless-probe',task:args.task,method_id:saved?.meta.id??null,
    execution_kind:'native_python',state:'preparing',candidate:true,created_at:new Date().toISOString(),
    interpreter:args.interpreter,cwd:args.cwd,argv:args.argv,timeout_ms:timeout,inputs:[],outputs:[],bundle:[],
    code_validation:{passed:false},file_check:{complete:false},enabled:false};
  const began=performance.now();record.timings_ms={};
  await persist(root,record);
  const folder=resolve(root,'.office-plugin','native-runs',record.id),bundleRoot=resolve(folder,'program');
  const freshExec={...exec,cwd};
  try {
    const inputFiles=new Map();
    const script=nativePath(saved?.meta.execution.entry??args.script_path);
    const local=saved?saved.meta.execution.bundle.map(f=>f.path):[script,...(args.dependencies??[]).map(nativePath)];
    if(local.length>32||new Set(local).size!==local.length||local.some(f=>!f.endsWith('.py')))throw new Error('OFFICE_NATIVE_BUNDLE');
    for(const [i,path] of local.entries()) {
      nativePath(path);
      const item=saved?{bytes:i===0?Buffer.from(saved.code):saved.dependencies[i-1].bytes,hash:saved.meta.execution.bundle[i].sha256}:
        await readInput(ctx,root,path,exec.signal,65536);
      if(sha256(item.bytes)!==item.hash)throw new Error('OFFICE_NATIVE_CODE_CHANGED');
      const dest=resolve(bundleRoot,path);await safeDirectory(root,dirname(dest));
      await writeFile(dest,item.bytes,{flag:'wx'});
      record.bundle.push({path,sha256:item.hash,snapshot:relative(root,dest).replaceAll('\\','/')});
    }
    for(const path of args.inputs.map(nativePath)) {
      const item=await readInput(ctx,root,path,exec.signal,15*1024*1024);record.inputs.push({path,sha256:item.hash});inputFiles.set(path,item);
    }
    let outputPath,staged,argv,check,officeRequest,group;
    if(args.check?.artifacts){
      args=await bindGroupRule(ctx,config,root,args,saved,record,inputFiles,exec);
      group=await prepareGroup(ctx,root,args,saved,record,inputFiles,folder,exec);
      argv=group.argv;officeRequest=group.officeRequest;
    }else{
      if(args.outputs.length!==1)throw new Error('OFFICE_NATIVE_GROUP_CHECKLIST');
    const output=args.outputs[0];outputPath=nativePath(output.path);
    if(!Number.isInteger(output.argv_index)||output.argv_index<0||output.argv_index>=args.argv.length||args.argv[output.argv_index]!==output.path)
      throw new Error('OFFICE_NATIVE_OUTPUT_ARGUMENT');
    if(args.inputs.includes(output.path)||local.includes(output.path))throw new Error('OFFICE_NATIVE_ORIGINAL_OUTPUT');
    await safeDirectory(root,resolve(root,dirname(outputPath)));
    await nativeFolder(root,dirname(outputPath));
    const target=await ctx.fs.resolve(resolve(root,outputPath),{cwd:root});
    if(outside(root,target.targetKey))throw new Error('OFFICE_NATIVE_OUTPUT_LINK');
    try {await lstat(resolve(root,outputPath));throw new Error('OFFICE_NATIVE_OUTPUT_EXISTS');}catch(e){if(e.code!=='ENOENT')throw e;}
    staged=resolve(folder,'artifact'+(outputPath.toLowerCase().endsWith('.xlsx')?'.xlsx':'.docx'));
    if(!/\.(xlsx|docx)$/i.test(outputPath))throw new Error('OFFICE_NATIVE_OUTPUT_FORMAT');
    argv=args.argv.map((arg,i)=>i===output.argv_index?staged:arg);
    // Convert declared inputs to absolute paths so cwd and captured script parent
    // do not change the meaning of their structured file arguments.
    for(let i=0;i<argv.length;i++) if(args.inputs.includes(argv[i]))argv[i]=resolve(root,argv[i]);
    check=args.check;
    if(!check||!args.inputs.includes(check.target_path)||!check.source_paths?.length||!check.source_paths.every(p=>args.inputs.includes(p)))throw new Error('OFFICE_NATIVE_CHECK_INPUTS');
    const pack=async path=>({format:path.split('.').at(-1).toLowerCase(),bytes_b64:inputFiles.get(nativePath(path)).bytes.toString('base64')});
    const rule=saved?.meta.execution.relationship_rule??args.relationship_rule;
    let currentConditions;
    if(rule!==undefined) {
      record.relationship_rule=structuredClone(rule);
      if(!saved&&check.contract)throw new Error('OFFICE_RELATIONSHIP_ONE_SOURCE: 长期规则与固定清单二选一');
      if((saved?.meta.execution.relationship_arg_index??args.relationship_arg_index)===undefined)
        throw new Error('OFFICE_RELATIONSHIP_ARGUMENT_REQUIRED: 程序需接收本批处理关系参数');
      const sources=[];for(const path of check.source_paths)sources.push(await pack(path));
      const start=performance.now();
      const bound=await runWorker(ctx,config,root,{operation:'bind_relationship',rule,sources,target:await pack(check.target_path)},exec);
      if(JSON.stringify(bound.source_sha256)!==JSON.stringify(check.source_paths.map(p=>inputFiles.get(nativePath(p)).hash))||
         bound.target_sha256!==inputFiles.get(nativePath(check.target_path)).hash)throw new Error('OFFICE_RELATIONSHIP_DIGEST_MISMATCH');
      check={...check,contract:bound.contract};currentConditions=bound.conditions;
      record.relationship_binding={rule_sha256:sha256(JSON.stringify(rule)),record_count:bound.record_count,field_count:bound.field_count,
        source_header_row:bound.source_header_row,target_header_row:bound.target_header_row,locations:bound.locations};
      record.timings_ms.relationship_binding=performance.now()-start;
    }
    check={...check,contract:expandContract(check.contract)};
    const bindingIndex=saved?.meta.execution.relationship_arg_index??args.relationship_arg_index;
    if(bindingIndex!==undefined) {
      if(!Number.isInteger(bindingIndex)||bindingIndex<0||bindingIndex>=argv.length||bindingIndex===output.argv_index||args.inputs.includes(args.argv[bindingIndex]))
        throw new Error('OFFICE_TASK_BINDING_ARGUMENT');
      const binding=taskBinding({contract:check.contract,sourcePaths:check.source_paths.map(p=>resolve(root,p)),
        targetPath:resolve(root,check.target_path),outputPath:staged,inputHashes:Object.fromEntries(record.inputs.map(f=>[f.path,f.sha256])),relationshipRule:record.relationship_rule});
      const bindingPath=resolve(folder,'task-binding.json'),bytes=JSON.stringify(binding,null,2);
      await writeFile(bindingPath,bytes,{flag:'wx'});argv[bindingIndex]=bindingPath;
      record.relationship_arg_index=bindingIndex;
      record.task_binding={path:relative(root,bindingPath).replaceAll('\\','/'),sha256:sha256(bytes),
        contract_sha256:binding.contract_sha256,required_count:binding.required_count};
    }
    officeRequest={sources:[],target:await pack(check.target_path),contract:check.contract,conditions:currentConditions??saved?.meta.conditions??args.conditions};
    for(const path of check.source_paths)officeRequest.sources.push(await pack(path));
    }
    record.conditions=group?{schema:'office_group_conditions_v1',artifacts:officeRequest.artifacts.map(a=>a.conditions)}:officeRequest.conditions;
    record.actual_argv=[args.interpreter,'-B','-X','utf8',resolve(bundleRoot,script),...argv];
    record.interpreter_sha256=sha256(await readFile(args.interpreter));
    const request={bundle_root:bundleRoot,files:local,original_script:resolve(root,script),workspace:root,cwd,office_request:officeRequest,
      interpreter_sha256:record.interpreter_sha256,bundle_hashes:record.bundle.map(({path,sha256})=>({path,sha256})),
      expected_fingerprints:saved?saved.meta.proofs.filter(p=>p.complete&&p.code_sha256===saved.meta.code_sha256).map(p=>p.runtime_fingerprint):null};
    let result;
    if(config.nativeLaunchPath) {
      if(!isAbsolute(config.nativeLaunchPath))throw new Error('OFFICE_NATIVE_LAUNCH_PATH');
      record.launcher_argv=[args.interpreter,'-B','-X','utf8',config.nativeLaunchPath,...record.actual_argv.slice(4)];
      record.state='launching';record.started_at=new Date().toISOString();await persist(root,record);
      const start=performance.now();
      const raw=await hostRun(ctx,config,root,freshExec,
        `& ${record.launcher_argv.map(quotePwsh).join(' ')}\nexit $LASTEXITCODE`,JSON.stringify(request),timeout);
      record.timings_ms.combined_launch=performance.now()-start;
      record.launcher_execution={exit_code:raw.exitCode,timed_out:raw.timedOut===true,aborted:raw.aborted===true,sandbox:raw.sandbox};
      checkHost(raw);exec.signal.throwIfAborted();
      if(raw.timedOut||raw.aborted)throw new Error('OFFICE_NATIVE_EXECUTION_FAILED');
      if(raw.stdout.truncated)throw new Error('OFFICE_NATIVE_LAUNCH_PROTOCOL: stdout截断，不能确认协议头');
      const cut=raw.stdout.text.indexOf('\n');
      let header;try{header=JSON.parse(cut<0?raw.stdout.text:raw.stdout.text.slice(0,cut));}
      catch{throw new Error('OFFICE_NATIVE_LAUNCH_PROTOCOL');}
      if(header.protocol!=='office_native_launch_v1'||!['ready','rejected'].includes(header.status))throw new Error('OFFICE_NATIVE_LAUNCH_PROTOCOL');
      record.timings_ms.preflight_inside_launcher=header.preflight_ms;
      if(header.preflight_detail_ms)record.timings_ms.preflight_detail=header.preflight_detail_ms;
      record.business_launch_authorized=header.status==='ready';
      if(header.status==='rejected')throw new Error(header.error?.includes('OFFICE_NATIVE_REVALIDATE')?header.error:
        'OFFICE_NATIVE_CODE_VALIDATION: '+header.error);
      record.environment=header.environment;
      const {cwd:observedCwd,...runtimeEnvironment}=header.environment;
      record.runtime_fingerprint=sha256(JSON.stringify({environment:runtimeEnvironment,interpreter_sha256:record.interpreter_sha256}));
      if(record.runtime_fingerprint!==header.runtime_fingerprint||saved&&!request.expected_fingerprints.includes(record.runtime_fingerprint))
        throw new Error('OFFICE_NATIVE_REVALIDATE: 当前环境需要新验证');
      record.code_validation={passed:true,at:new Date().toISOString()};
      result={...raw,stdout:{...raw.stdout,text:cut<0?'':raw.stdout.text.slice(cut+1)}};
    }else {
      const probeStarted=performance.now();
      const probe=await hostRun(ctx,config,root,freshExec,
        `& ${quotePwsh(args.interpreter)} -B -X utf8 ${quotePwsh(config.nativeIdentityPath)}`,JSON.stringify(request),timeout);
      checkHost(probe);record.timings_ms.preflight=performance.now()-probeStarted;
      if(probe.exitCode!==0||probe.timedOut||probe.aborted||probe.stdout.truncated)throw new Error('OFFICE_NATIVE_CODE_VALIDATION: '+probe.stderr.text.slice(0,2000));
      record.environment=JSON.parse(probe.stdout.text);
      const {cwd:observedCwd,...runtimeEnvironment}=record.environment;
      record.runtime_fingerprint=sha256(JSON.stringify({environment:runtimeEnvironment,interpreter_sha256:record.interpreter_sha256}));
      record.code_validation={passed:true,at:new Date().toISOString()};
      if(saved && !request.expected_fingerprints.includes(record.runtime_fingerprint))throw new Error('OFFICE_NATIVE_REVALIDATE: 当前环境需要新验证');
      record.state='running';record.started_at=new Date().toISOString();await persist(root,record);
      const executionStarted=performance.now();
      result=await hostRun(ctx,config,root,freshExec,
        `& ${record.actual_argv.map(quotePwsh).join(' ')}`,undefined,timeout);
      record.timings_ms.program=performance.now()-executionStarted;
    }
    await persist(root,record);
    record.execution={exit_code:result.exitCode,timed_out:result.timedOut===true,aborted:result.aborted===true,
      sandbox:result.sandbox,stdout:result.stdout.text,stderr:result.stderr.text,stdout_truncated:result.stdout.truncated};
    record.finished_at=new Date().toISOString();
    checkHost(result);exec.signal.throwIfAborted();
    if(result.timedOut||result.aborted||result.exitCode!==0)throw new Error('OFFICE_NATIVE_EXECUTION_FAILED');
    for(const f of record.bundle) if((await readInput(ctx,root,f.snapshot,exec.signal)).hash!==f.sha256)throw new Error('OFFICE_NATIVE_CODE_CHANGED');
    if(record.task_binding&&(await readInput(ctx,root,record.task_binding.path,exec.signal,1024*1024)).hash!==record.task_binding.sha256)
      throw new Error('OFFICE_TASK_BINDING_CHANGED');
    for(const f of record.inputs) if((await readInput(ctx,root,f.path,exec.signal,15*1024*1024)).hash!==f.sha256)throw new Error('OFFICE_NATIVE_INPUT_CHANGED');
    if(sha256(await readFile(args.interpreter))!==record.interpreter_sha256)throw new Error('OFFICE_NATIVE_INTERPRETER_CHANGED');
    if(group){
      const started=performance.now();record.file_check=await checkGroup(ctx,config,root,record,exec);
      record.timings_ms.independent_check=performance.now()-started;
      if(!record.file_check.complete)throw new Error('OFFICE_NATIVE_FILE_CHECK_FAILED');
      await publishGroup(ctx,config,root,record,exec,async()=>{
        await sealNativeRun(config,root,record);
      });
    }else{
    const stagedPath=relative(root,staged).replaceAll('\\','/');
    const artifact=await readInput(ctx,root,stagedPath,exec.signal,15*1024*1024);
    record.outputs=[{path:outputPath,staged_path:stagedPath,sha256:artifact.hash,state:'generated'}];
    const checkStarted=performance.now();
    record.file_check=await checkArtifact(ctx,config,root,{...check,output_path:stagedPath},exec);
    record.timings_ms.independent_check=performance.now()-checkStarted;
    if(!record.file_check.complete)throw new Error('OFFICE_NATIVE_FILE_CHECK_FAILED');
    for(const f of record.inputs)if((await readInput(ctx,root,f.path,exec.signal,15*1024*1024)).hash!==f.sha256)throw new Error('OFFICE_NATIVE_INPUT_CHANGED');
    if((await readInput(ctx,root,stagedPath,exec.signal,15*1024*1024)).hash!==artifact.hash)throw new Error('OFFICE_NATIVE_ARTIFACT_CHANGED');
    policyFor(ctx,root,exec,config,true);exec.signal.throwIfAborted();
    await nativeFolder(root,dirname(outputPath));
    if(!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
    const handle=await open(resolve(root,outputPath),'wx');
    try{await handle.writeFile(artifact.bytes);await handle.sync();}finally{await handle.close();}
    record.outputs[0].state='published';record.state='checked';    }

  }catch(e) {
    record.state=record.group_created?'recovery_required':exec.signal.aborted?'cancelled':record.execution?.timed_out||record.launcher_execution?.timed_out?'timed_out':'failed';
    record.error=e.message;record.finished_at??=new Date().toISOString();
    // Snapshot status survives cancellation. It is a candidate, never proof of
    // success, and cleanup does not run further business code.
    if(record.group_checks){
      for(const out of record.outputs)if(!out.sha256)try{out.sha256=(await readInput(ctx,root,out.staged_path,cleanSignal(),15*1024*1024)).hash;out.state='unverified';}
      catch(error){if(!['ENOENT','FS_NOT_FOUND'].includes(error.code))record.recovery_error=error.message;}
    }
    if(!record.outputs.length) {
      for(const name of ['artifact.xlsx','artifact.docx'])try {
        const path=relative(root,resolve(folder,name)).replaceAll('\\','/');
        const f=await readInput(ctx,root,path,cleanSignal(),15*1024*1024);
        record.outputs.push({staged_path:path,sha256:f.hash,state:'unverified'});
      }catch{}
    }
  }
  if(!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
  record.timings_ms.total=performance.now()-began;
  await sealNativeRun(config,root,record);
  return record;
}
