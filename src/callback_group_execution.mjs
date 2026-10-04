import { writeFile } from 'node:fs/promises';
import { resolve,relative } from 'node:path';
import { randomUUID } from 'node:crypto';
import { runtimeIdentity,safeDirectory } from './method_execution.mjs';
import { prepareGroup,checkGroup,publishGroup } from './native_group.mjs';
import { sealNativeRun } from './native_execution.mjs';
import { readInput,policyFor,sha256,quotePwsh } from './runtime.mjs';

export async function executeCallbackGroup(ctx,config,root,args,exec,saved,publish){
  const policy=policyFor(ctx,root,exec,config,true),{lock,fingerprint}=await runtimeIdentity(config);
  if(publish&&!saved.meta.proofs.some(p=>p.complete&&p.code_sha256===saved.meta.code_sha256&&p.runtime_fingerprint===fingerprint))throw new Error('OFFICE_METHOD_REVALIDATE');
  const contracts=saved.meta.contract.artifacts,targets=args.target_paths,names=args.inputs?.output_names;
  if(!Array.isArray(targets)||targets.length!==contracts.length||!Array.isArray(names)||names.length!==contracts.length||
      names.some(n=>typeof n!=='string'||/[\\/:]/.test(n)))throw new Error('OFFICE_METHOD_GROUP_TARGETS');
  const paths=names.map(name=>(args.inputs.output_directory??'delivery')+'/'+name),inputPaths=[...new Set([...args.source_paths,...targets])];
  const run={id:randomUUID(),owner:exec.agent?.session?.id??'agentless-probe',method_id:saved.meta.id,task:'已保存回调整组执行',
    execution_kind:'office_callbacks',candidate:false,state:'preparing',created_at:new Date().toISOString(),inputs:[],outputs:[],bundle:[],
    code_validation:{passed:false},file_check:{complete:false},runtime_fingerprint:fingerprint,environment:{trusted_worker_versions:lock.files.filter(f=>/worker[\\/][^\\/]+\.py$/.test(f.path)).map(f=>[f.path.split(/[\\/]/).at(-1),f.sha256])}};
  const folder=resolve(root,'.office-plugin/native-runs',run.id);await safeDirectory(root,folder);
  await writeFile(resolve(folder,'method.py'),saved.code,{flag:'wx'});run.bundle=[{path:'method.py',sha256:saved.meta.code_sha256,snapshot:relative(root,resolve(folder,'method.py')).replaceAll('\\','/')}];
  try{
    const inputFiles=new Map();let inputBytes=0;
    for(const path of inputPaths){const item=await readInput(ctx,root,path,exec.signal,15*1024*1024);inputBytes+=item.bytes.length;
      if(inputBytes>18*1024*1024)throw new Error('OFFICE_METHOD_GROUP_INPUT_LIMIT');inputFiles.set(path,item);run.inputs.push({path,sha256:item.hash});}
    const specs=contracts.map((a,i)=>({...a,target_path:targets[i],conditions:saved.meta.conditions.artifacts[i]}));
    const prepared=await prepareGroup(ctx,root,{argv:['@group',...paths],inputs:inputPaths,outputs:paths.map((path,i)=>({path,argv_index:i+1})),
      relationship_arg_index:0,check:{source_paths:args.source_paths,artifacts:specs,consistency:saved.meta.contract.consistency}},null,run,inputFiles,folder,exec);
    const request={code:saved.code,code_sha256:saved.meta.code_sha256,dependency_root:lock.dependency_root,monty_binary:lock.binary_path,monty_sha256:lock.binary_sha256,
      conditions:saved.meta.conditions,sources:prepared.officeRequest.sources,artifacts:prepared.officeRequest.artifacts.map((a,i)=>({...a,output_name:names[i],target_path:targets[i]})),
      consistency:saved.meta.contract.consistency.map(l=>({...l,refs:l.refs.map(r=>({output_path:names[r.output_index],target_ref:r.target_ref}))})),
      inputs:{...args.inputs,source_count:args.source_paths.length,target_count:targets.length},persist_generated:true};
    if(!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
    const spec=ctx.shell.resolve({command:`& ${quotePwsh(config.pythonPath)} -I -B -X utf8 ${quotePwsh(lock.runner_path)}`,workdir:resolve(folder,'artifacts'),
      sandboxPolicy:{...policy,workspaceRoot:folder},signal:exec.signal,timeoutMs:60000,onExpiry:'kill',stdin:JSON.stringify(request),stdoutMaxBytes:1024*1024});
    const result=await(await ctx.shell.execute(spec)).result();exec.signal.throwIfAborted();
    if(result.timedOut||result.aborted||result.stdout.truncated||!['partial','enforced'].includes(result.sandbox?.enforcement))throw new Error('OFFICE_METHOD_EXECUTION_STOPPED');
    const value=JSON.parse(result.stdout.text);
    if(!value.ok||result.exitCode!==0)throw new Error('OFFICE_METHOD_FAILED: '+value.error?.message);
    run.code_validation={passed:true};run.execution={exit_code:0};run.file_check=await checkGroup(ctx,config,root,run,exec);
    if(publish&&run.file_check.complete)await publishGroup(ctx,config,root,run,exec,()=>sealNativeRun(config,root,run));
    else run.state=run.file_check.complete?'validated':'failed';
  }catch(error){
    run.state=run.group_created?'recovery_required':exec.signal.aborted?'cancelled':'failed';run.error=error.message;
    for(const out of run.outputs.filter(o=>!o.sha256)){
      try{out.sha256=(await readInput(ctx,root,out.staged_path,new AbortController().signal,15*1024*1024)).hash;out.state='unverified';}
      catch(readError){if(!['ENOENT','FS_NOT_FOUND'].includes(readError.code))out.recovery_error=readError.message;}
    }
  }
  await sealNativeRun(config,root,run);
  return {ok:true,run_id:run.id,complete:run.state===(publish?'checked':'validated'),published:run.state==='checked',method_id:saved.meta.id,runtime_fingerprint:fingerprint,
    source_sha256:run.inputs.filter(f=>args.source_paths.includes(f.path)).map(f=>f.sha256),target_sha256:run.inputs.filter(f=>targets.includes(f.path)).map(f=>f.sha256),
    generated:run.outputs.map((out,i)=>({name:names[i],sha256:out.sha256??null,verification:run.file_check.artifacts?.[i]??null})),
    output_paths:run.outputs.map(o=>run.state==='checked'?o.path:o.staged_path),job_path:relative(root,folder),state:run.state,error:run.error??null,
    verification:run.file_check,delivery_manifest:run.delivery_manifest??null};
}
