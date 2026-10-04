import { readFile } from 'node:fs/promises';
import { resolve,dirname } from 'node:path';
import { loadNativeRun,sealNativeRun } from './native_execution.mjs';
import { checkGroup,publishGroup } from './native_group.mjs';
import { policyFor,readInput,sha256 } from './runtime.mjs';

export async function resumeGroup(ctx,config,root,args,exec){
  policyFor(ctx,root,exec,config,true);
  const run=await loadNativeRun(ctx,config,root,args.run_id,exec);
  if(!run.group_checks||!run.code_validation.passed||!run.file_check.complete)
    throw new Error('OFFICE_NATIVE_GROUP_RESUME_CHECK_REQUIRED: 整组业务与检查先完成；已生成暂存文件保留，新执行补齐全部产物后重新核对');
  for(const [name,hash] of run.environment.trusted_worker_versions??[])if(!/^[a-zA-Z0-9_-]+\.py$/.test(name)||
    sha256(await readFile(resolve(dirname(config.nativeIdentityPath),name)))!==hash)throw new Error('OFFICE_NATIVE_REVALIDATE');
  for(const file of [...run.inputs,...run.bundle.map(f=>({path:f.snapshot,sha256:f.sha256}))])
    if((await readInput(ctx,root,file.path,exec.signal,15*1024*1024)).hash!==file.sha256)throw new Error('OFFICE_NATIVE_INPUT_OR_CODE_CHANGED');
  run.file_check=await checkGroup(ctx,config,root,run,exec);
  if(!run.file_check.complete)throw new Error('OFFICE_NATIVE_FILE_CHECK_FAILED');
  if(run.state==='checked'){
    for(const out of run.outputs)if((await readInput(ctx,root,out.path,exec.signal,15*1024*1024)).hash!==out.sha256)
      throw new Error('OFFICE_NATIVE_PUBLISHED_CHANGED');
    if(!run.delivery_manifest||(await readInput(ctx,root,run.delivery_manifest,exec.signal,1024*1024)).hash!==run.delivery_manifest_sha256)
      throw new Error('OFFICE_NATIVE_MANIFEST_CHANGED');
    return run;
  }
  run.resume_attempts=(run.resume_attempts??0)+1;run.last_error=run.error;delete run.error;
  try{await publishGroup(ctx,config,root,run,exec,()=>sealNativeRun(config,root,run));}
  catch(error){run.state='recovery_required';run.error=error.message;}
  await sealNativeRun(config,root,run);return run;
}
