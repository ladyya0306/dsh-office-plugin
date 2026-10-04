import { readFile, mkdir, lstat, realpath, open } from 'node:fs/promises';
import { resolve, relative, isAbsolute } from 'node:path';
import { randomUUID } from 'node:crypto';
import { readInput, policyFor, sha256, quotePwsh, outside } from './runtime.mjs';
import { executeCallbackGroup } from './callback_group_execution.mjs';

export async function runtimeIdentity(config) {
  const lock=JSON.parse(await readFile(config.methodRuntimeLock,'utf8'));
  if (lock.version!=='1.0.0' || !isAbsolute(lock.dependency_root) || resolve(config.pythonPath)!==resolve(lock.python_path)) throw new Error('OFFICE_METHOD_RUNTIME_LOCK');
  for (const item of lock.files) if (sha256(await readFile(item.path))!==item.sha256) throw new Error('OFFICE_METHOD_RUNTIME_CHANGED: '+item.path);
  return {lock,fingerprint:sha256(JSON.stringify(lock))};
}
export async function safeDirectory(root, folder) {
  if (outside(root,folder)) throw new Error('OFFICE_METHOD_STAGE_OUTSIDE');
  const segments=relative(root,folder).split(/[\\/]/); let current=root;
  for (const segment of segments) {
    current=resolve(current,segment); await mkdir(current).catch(e=>{if(e.code!=='EEXIST') throw e;});
    const stat=await lstat(current);
    if (!stat.isDirectory() || stat.isSymbolicLink() || outside(await realpath(root),await realpath(current))) throw new Error('OFFICE_METHOD_STAGE_LINK');
  }
}
export async function executeMethod(ctx,config,root,args,exec,record,publish) {
  if(record.meta.execution?.kind==='native_python')throw new Error('OFFICE_METHOD_EXECUTOR_MISMATCH: 普通脚本使用office_program_reuse入口');
  if(record.meta.contract.schema==='office_group_contract_v1')return executeCallbackGroup(ctx,config,root,args,exec,record,publish);
  const policy=policyFor(ctx,root,exec,config,true);
  const {lock,fingerprint}=await runtimeIdentity(config);
  if (publish && !record.meta.proofs.some(p=>p.runtime_fingerprint===fingerprint && p.code_sha256===record.meta.code_sha256 && p.complete))
    throw new Error('OFFICE_METHOD_REVALIDATE: 当前解释器或代码需要重新验证');
  if (!Array.isArray(args.source_paths) || args.source_paths.length<1 || args.source_paths.length>8) throw new Error('OFFICE_SOURCE_LIMIT');
  const read=async path=>{
    const format=path.split('.').at(-1).toLowerCase();
    if (!['xlsx','docx'].includes(format)) throw new Error('OFFICE_METHOD_FORMAT_UNSUPPORTED');
    return {...await readInput(ctx,root,path,exec.signal),format};
  };
  const sources=[]; for(const path of args.source_paths) sources.push(await read(path));
  const target=await read(args.target_path);
  if ([...sources,target].reduce((v,f)=>v+f.bytes.length,0)>18*1024*1024) throw new Error('OFFICE_METHOD_INPUT_LIMIT');
  if (!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);
  const folder=resolve(root,'.office-plugin','jobs',randomUUID()); await safeDirectory(root,folder);
  const pack=f=>({format:f.format,bytes_b64:f.bytes.toString('base64'),filename:f.path.split(/[\\/]/).at(-1)});
  const request={code:record.code,code_sha256:record.meta.code_sha256,
    dependency_root:lock.dependency_root,monty_binary:lock.binary_path,monty_sha256:lock.binary_sha256,
    conditions:record.meta.conditions,contract:record.meta.contract,sources:sources.map(pack),target:pack(target),
    inputs:{...args.inputs,source_count:sources.length,source_paths:args.source_paths,target_path:args.target_path}};
  const spec=ctx.shell.resolve({command:`& ${quotePwsh(config.pythonPath)} -I -B -X utf8 ${quotePwsh(lock.runner_path)}`,
    workdir:folder,timeoutMs:60000,onExpiry:'kill',signal:exec.signal,
    sandboxPolicy:{...policy,workspaceRoot:folder},stdin:JSON.stringify(request),stdoutMaxBytes:1024*1024});
  const result=await (await ctx.shell.execute(spec)).result(); exec.signal.throwIfAborted();
  if (result.timedOut||result.aborted||result.stdout.truncated || !['partial','enforced'].includes(result.sandbox?.enforcement))
    throw new Error('OFFICE_METHOD_EXECUTION_STOPPED: 暂存区可能有产物，未报告完成');
  let value; try {value=JSON.parse(result.stdout.text);} catch {throw new Error('OFFICE_METHOD_PROTOCOL');}
  if (!value.ok||result.exitCode!==0) throw new Error('OFFICE_METHOD_FAILED: '+(value.error?.message??'执行器失败')+
    '；本次运行的是已保存代码 '+record.meta.code_sha256.slice(0,12)+'。工作区原脚本编辑不会更新该版本；修正后再次save(parent_id=旧ID)，再验证新ID。'+
    '输入已提供source_count、source_paths相对名称、target_path相对名称；nodes.ref是完整工作表!坐标，不能只按A/B列字母开头判断。');
  for(const input of [...sources,target]) if ((await readInput(ctx,root,input.path,exec.signal)).hash!==input.hash) throw new Error('OFFICE_INPUT_CHANGED');
  const files=[];
  for(const artifact of value.generated??[]) {
    if (!/^[^\\/:*?"<>|\x00-\x1f]{1,120}\.(xlsx|docx)$/i.test(artifact.name)) throw new Error('OFFICE_METHOD_OUTPUT_NAME');
    const staged=await readInput(ctx,root,relative(root,resolve(folder,artifact.name)),exec.signal,15*1024*1024);
    if (staged.hash!==artifact.sha256 || artifact.verification.artifact_sha256!==staged.hash) throw new Error('OFFICE_METHOD_OUTPUT_CHANGED');
    files.push({artifact,staged});
  }
  if(publish && value.complete && files.length===1) {
    policyFor(ctx,root,exec,config,true); exec.signal.throwIfAborted();
    // wx rejects existing files and links. Source bytes are never edited.
    const destination=resolve(root,files[0].artifact.name);
    if(outside(root,destination)) throw new Error('OFFICE_METHOD_OUTPUT_OUTSIDE');
    const handle=await open(destination,'wx');
    try {await handle.writeFile(files[0].staged.bytes); await handle.sync();} finally {await handle.close();}
    const final=await readInput(ctx,root,files[0].artifact.name,exec.signal,15*1024*1024);
    if(final.hash!==files[0].staged.hash) throw new Error('OFFICE_METHOD_PUBLISH_CHANGED');
  } else if(publish && value.complete) throw new Error('OFFICE_METHOD_MULTIPLE_OUTPUTS_UNSUPPORTED: 本版不发布多个文件，暂存文件仍在');
  if(!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);
  return {...value,method_id:record.meta.id,runtime_fingerprint:fingerprint,
    source_sha256:sources.map(s=>s.hash),target_sha256:target.hash,
    output_paths:files.map(f=>publish&&value.complete?f.artifact.name:relative(root,resolve(folder,f.artifact.name))),
    published:publish&&value.complete,job_path:relative(root,folder)};
}
