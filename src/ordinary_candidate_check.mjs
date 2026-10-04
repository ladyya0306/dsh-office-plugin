import { readFile } from 'node:fs/promises';
import { resolve,dirname } from 'node:path';
import { performance } from 'node:perf_hooks';
import { readInput,policyFor,sha256,quotePwsh,runWorker } from './runtime.mjs';
import { loadNativeRun,sealNativeRun,nativePath } from './native_execution.mjs';
import { expandContract } from './contract_builder.mjs';
import { checkArtifact } from './output_tools.mjs';

// Verify the actual already executed bytes. No second business execution.
export async function checkOrdinaryCandidate(ctx,config,root,args,exec){
  policyFor(ctx,root,exec,config,true);
  const record=await loadNativeRun(ctx,config,root,args.run_id,exec);
  if(record.origin!=='ordinary_pwsh'||!['executed','check_failed','checked'].includes(record.state))throw new Error('OFFICE_CAPTURE_SUCCESS_REQUIRED');
  if(record.outputs.length!==1||!args.check?.target_path||args.check.artifacts)throw new Error('OFFICE_CAPTURE_SINGLE_CHECK_REQUIRED');
  const declared=args.outputs;if(!Array.isArray(declared)||declared.length!==1)throw new Error('OFFICE_CAPTURE_OUTPUT_ARGUMENT');
  const out=declared[0],observed=record.outputs[0],cwd=resolve(root,record.cwd);
  if(out.path!==observed.path||!Number.isInteger(out.argv_index)||out.argv_index<0||out.argv_index>=record.argv.length||
    resolve(cwd,record.argv[out.argv_index]).toLowerCase()!==resolve(root,out.path).toLowerCase())throw new Error('OFFICE_CAPTURE_OUTPUT_ARGUMENT');
  const sources=args.check.source_paths,target=args.check.target_path;
  if(!Array.isArray(sources)||!sources.length||sources.length>8||new Set(sources).size!==sources.length)throw new Error('OFFICE_CAPTURE_SOURCE_ARGUMENT');
  for(const path of [...sources,target])if(!record.inputs.some(f=>f.path===nativePath(path)))throw new Error('OFFICE_CAPTURE_INPUT_NOT_OBSERVED');
  for(const f of record.inputs)if((await readInput(ctx,root,f.path,exec.signal,15*1024*1024)).hash!==f.sha256||
    (await readInput(ctx,root,f.snapshot,exec.signal,15*1024*1024)).hash!==f.sha256)throw new Error('OFFICE_NATIVE_INPUT_CHANGED');
  for(const f of record.bundle){
    if((await readInput(ctx,root,f.path,exec.signal,65536)).hash!==f.sha256||(await readInput(ctx,root,f.snapshot,exec.signal,65536)).hash!==f.sha256)
      throw new Error('OFFICE_NATIVE_CODE_CHANGED');
  }
  if((await readInput(ctx,root,out.path,exec.signal,15*1024*1024)).hash!==observed.sha256)throw new Error('OFFICE_NATIVE_ARTIFACT_CHANGED');
  if(sha256(await readFile(config.pythonPath))!==record.interpreter_sha256)throw new Error('OFFICE_NATIVE_INTERPRETER_CHANGED');
  const pack=async path=>({format:path.split('.').at(-1).toLowerCase(),bytes_b64:(await readInput(ctx,root,path,exec.signal,15*1024*1024)).bytes.toString('base64')});
  let contract,conditions=args.conditions;
  if(args.relationship_rule){
    if(args.check.contract!==undefined)throw new Error('OFFICE_RELATIONSHIP_ONE_SOURCE');
    const index=args.relationship_arg_index;
    if(!Number.isInteger(index)||index<0||index>=record.argv.length||index===out.argv_index||
      record.inputs.some(f=>resolve(root,f.path).toLowerCase()===resolve(cwd,record.argv[index]).toLowerCase()))throw new Error('OFFICE_TASK_BINDING_ARGUMENT');
    const inputItems=[];for(const p of sources)inputItems.push(await pack(p));
    const bound=await runWorker(ctx,config,root,{operation:'bind_relationship',rule:args.relationship_rule,sources:inputItems,target:await pack(target)},exec);
    if(JSON.stringify(bound.source_sha256)!==JSON.stringify(sources.map(p=>record.inputs.find(f=>f.path===p).sha256))||bound.target_sha256!==record.inputs.find(f=>f.path===target).sha256)
      throw new Error('OFFICE_RELATIONSHIP_DIGEST_MISMATCH');
    contract=bound.contract;conditions=bound.conditions;
    record.relationship_rule=structuredClone(args.relationship_rule);record.relationship_arg_index=index;
  }else contract=expandContract(args.check.contract);
  if(!conditions?.source_guards?.length||!conditions.target_guards?.length)throw new Error('OFFICE_CAPTURE_STRUCTURE_GUARDS_REQUIRED');
  const checkStarted=performance.now();
  const checked=await checkArtifact(ctx,config,root,{source_paths:sources,target_path:target,output_path:out.path,contract},exec);
  record.timings_ms.independent_check=performance.now()-checkStarted;
  record.file_check={...checked,confirmed_contract:contract};record.conditions=conditions;
  if(!checked.complete){
    record.state='check_failed';await sealNativeRun(config,root,record);
    return {run_id:record.id,state:record.state,file_check:record.file_check,next:'修正程序后重新执行；错误程序不进入可复用库。'};
  }
  const bundleRoot=resolve(root,dirname(record.bundle[0].snapshot),...Array(record.bundle[0].path.split('/').length-1).fill('..'));
  const identityRequest={bundle_root:bundleRoot,files:record.bundle.map(f=>f.path),original_script:resolve(root,record.bundle[0].path),workspace:root,cwd};
  const spec=ctx.shell.resolve({command:`& ${quotePwsh(config.pythonPath)} -I -B -X utf8 ${quotePwsh(config.nativeIdentityPath)}`,
    workdir:cwd,signal:exec.signal,timeoutMs:60000,onExpiry:'kill',stdin:JSON.stringify(identityRequest),stdoutMaxBytes:1024*1024,sandboxPolicy:policyFor(ctx,root,exec,config,true)});
  const response=await(await ctx.shell.execute(spec)).result();
  if(response.exitCode!==0||response.timedOut||response.aborted||response.stdout.truncated||!['partial','enforced'].includes(response.sandbox?.enforcement)||response.sandbox?.denied)
    throw new Error('OFFICE_CAPTURE_VALIDATION_FAILED: '+response.stderr.text.slice(0,1500));
  record.environment=JSON.parse(response.stdout.text);
  const {cwd:observedCwd,...runtime}=record.environment;
  record.runtime_fingerprint=sha256(JSON.stringify({environment:runtime,interpreter_sha256:record.interpreter_sha256}));
  record.code_validation={passed:true};record.state='checked';record.outputs[0].state='published';
  // Preserve the complete observation record; source and target proof uses
  // the explicitly bound current task paths, other observed files stay audit-only.
  record.outputs[0].argv_index=out.argv_index;
  await sealNativeRun(config,root,record);
  return {run_id:record.id,state:'checked',file_check:{complete:true,checked:checked.checked,passed:checked.passed},
    business_program_reexecuted:false,code_sha256:record.bundle[0].sha256,next:'用office_program_save保存本次实际程序；修改后重新执行检查。'};
}
