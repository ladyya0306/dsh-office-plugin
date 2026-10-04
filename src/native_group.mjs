import { mkdir,open,writeFile,lstat } from 'node:fs/promises';
import { resolve,relative,dirname,basename } from 'node:path';
import { nativePath,nativeFolder } from './native_execution.mjs';
import { safeDirectory } from './method_execution.mjs';
import { readInput,runWorker,sha256,policyFor } from './runtime.mjs';
import { expandContract } from './contract_builder.mjs';
import { taskBinding } from './task_binding.mjs';

export async function prepareGroup(ctx,root,args,saved,record,inputFiles,folder,exec) {
  const checks=args.check.artifacts;
  if(!Array.isArray(checks)||checks.length!==args.outputs.length||checks.length>16||!checks.length||
      checks.some((c,i)=>c.output_index!==i)||!Array.isArray(args.check.consistency)||!args.check.consistency.length)
    throw new Error('OFFICE_NATIVE_GROUP_CHECKLIST: 每份产物按顺序明确模板、清单及整组关系');
  const paths=args.outputs.map(o=>nativePath(o.path));
  const directory=dirname(paths[0]);
  if(directory==='.'||paths.some(p=>dirname(p)!==directory)||new Set(paths.map(p=>p.toLowerCase())).size!==paths.length)
    throw new Error('OFFICE_NATIVE_GROUP_DIRECTORY: 整组需同一新的工作区子目录和不同文件名');
  await nativeFolder(root,dirname(directory));
  try{await lstat(resolve(root,directory));throw new Error('OFFICE_NATIVE_GROUP_EXISTS');}catch(e){if(e.code!=='ENOENT')throw e;}
  const argv=[...args.argv],used=new Set();
  for(let i=0;i<args.outputs.length;i++){
    const o=args.outputs[i],path=paths[i];
    if(!/\.(xlsx|docx)$/i.test(path)||!Number.isInteger(o.argv_index)||o.argv_index<0||o.argv_index>=argv.length||
       used.has(o.argv_index)||argv[o.argv_index]!==o.path||args.inputs.includes(path)||record.bundle.some(f=>f.path===path))
      throw new Error('OFFICE_NATIVE_GROUP_OUTPUT_ARGUMENT');
    used.add(o.argv_index);
    const staged=resolve(folder,'artifacts',basename(path));await safeDirectory(root,dirname(staged));argv[o.argv_index]=staged;
    record.outputs.push({path,staged_path:relative(root,staged).replaceAll('\\','/'),state:'pending'});
  }
  for(let i=0;i<argv.length;i++)if(args.inputs.includes(argv[i]))argv[i]=resolve(root,argv[i]);
  const sources=args.check.source_paths;
  if(!sources?.length||sources.length>8||!sources.every(p=>inputFiles.has(nativePath(p))))throw new Error('OFFICE_NATIVE_CHECK_INPUTS');
  const pack=p=>{const item=inputFiles.get(nativePath(p));if(!item)throw new Error('OFFICE_NATIVE_CHECK_INPUTS');return {format:p.split('.').at(-1).toLowerCase(),bytes_b64:item.bytes.toString('base64')};};
  const normalized=checks.map((c,i)=>{
    const contract=expandContract(c.contract);
    if(contract.preserve_template===false||!inputFiles.has(nativePath(c.target_path))||c.target_path.split('.').at(-1).toLowerCase()!==paths[i].split('.').at(-1).toLowerCase())
      throw new Error('OFFICE_NATIVE_GROUP_TEMPLATE');
    return {output_index:i,target_path:c.target_path,contract,conditions:c.conditions};
  });
  const consistency=args.check.consistency.map(link=>{
    if(!link||!['text','number'].includes(link.kind)||!Array.isArray(link.refs)||link.refs.length<2||link.refs.length>16||
        new Set(link.refs.map(r=>r?.output_index)).size<2||new Set(link.refs.map(r=>`${r?.output_index}:${r?.target_ref}`)).size!==link.refs.length)
      throw new Error('OFFICE_NATIVE_GROUP_CONSISTENCY');
    return {...link,refs:link.refs.map(ref=>{
      if(!ref||!Number.isInteger(ref.output_index)||ref.output_index<0||ref.output_index>=paths.length||
          !normalized[ref.output_index].contract.requirements.some(r=>r.target_ref===ref.target_ref))throw new Error('OFFICE_NATIVE_GROUP_CONSISTENCY');
      return {output_path:paths[ref.output_index],target_ref:ref.target_ref};
    })};
  });
  const index=saved?.meta.execution.relationship_arg_index??args.relationship_arg_index;
  if(!Number.isInteger(index)||index<0||index>=argv.length||used.has(index)||args.inputs.includes(args.argv[index]))
    throw new Error('OFFICE_TASK_BINDING_ARGUMENT');
  const binding={schema:'office_task_group_binding_v1',artifacts:normalized.map((c,i)=>taskBinding({contract:c.contract,
    sourcePaths:sources.map(p=>resolve(root,p)),targetPath:resolve(root,c.target_path),outputPath:argv[args.outputs[i].argv_index],
    inputHashes:Object.fromEntries(record.inputs.map(f=>[f.path,f.sha256]))}))};
  const path=resolve(folder,'task-binding.json'),bytes=JSON.stringify(binding,null,2);await writeFile(path,bytes,{flag:'wx'});argv[index]=path;
  record.relationship_arg_index=index;record.task_binding={path:relative(root,path).replaceAll('\\','/'),sha256:sha256(bytes)};
  record.group_checks={source_paths:sources,artifacts:normalized,consistency:args.check.consistency};
  record.group_directory=directory;
  return {argv,officeRequest:{sources:sources.map(pack),artifacts:normalized.map(c=>({target:pack(c.target_path),contract:c.contract,conditions:c.conditions}))},consistency};
}

export async function checkGroup(ctx,config,root,record,exec) {
  const pack=async p=>{
    const file=await readInput(ctx,root,p,exec.signal,15*1024*1024),input=record.inputs.find(f=>f.path===p);
    if(input&&input.sha256!==file.hash)throw new Error('OFFICE_NATIVE_INPUT_CHANGED');
    return {format:p.split('.').at(-1).toLowerCase(),bytes_b64:file.bytes.toString('base64'),sha256:file.hash};
  };
  const checks=record.group_checks,sources=[];for(const p of checks.source_paths)sources.push(await pack(p));
  const artifacts=[];
  for(const item of checks.artifacts){
    const out=record.outputs[item.output_index],artifact=await pack(out.staged_path),target=await pack(item.target_path);
    if(out.sha256&&out.sha256!==artifact.sha256)throw new Error('OFFICE_NATIVE_ARTIFACT_CHANGED');
    out.sha256=artifact.sha256;if(out.state==='pending')out.state='generated';
    artifacts.push({output_path:out.path,target_path:item.target_path,artifact,target,contract:item.contract});
  }
  const consistency=checks.consistency.map(link=>({...link,refs:link.refs.map(ref=>({output_path:record.outputs[ref.output_index].path,target_ref:ref.target_ref}))}));
  const checked=await runWorker(ctx,config,root,{operation:'check_group_output',sources,artifacts,consistency},exec);
  for(let i=0;i<artifacts.length;i++){
    const item=checked.artifacts[i];
    if(item.output_path!==artifacts[i].output_path||item.artifact_sha256!==artifacts[i].artifact.sha256||item.target_sha256!==artifacts[i].target.sha256||
        JSON.stringify(item.source_sha256)!==JSON.stringify(sources.map(s=>s.sha256)))throw new Error('OFFICE_NATIVE_INDEPENDENT_PROOF_MISMATCH');
  }
  return {...checked,confirmed_contract:{schema:'office_group_contract_v1',artifacts:checks.artifacts.map(c=>({output_index:c.output_index,contract:c.contract})),consistency:checks.consistency}};
}

export async function publishGroup(ctx,config,root,record,exec,checkpoint) {
  policyFor(ctx,root,exec,config,true);exec.signal.throwIfAborted();
  if(!record.file_check.complete)throw new Error('OFFICE_NATIVE_GROUP_UNCHECKED');
  for(const input of record.inputs)if((await readInput(ctx,root,input.path,exec.signal,15*1024*1024)).hash!==input.sha256)throw new Error('OFFICE_NATIVE_INPUT_CHANGED');
  for(const out of record.outputs)if((await readInput(ctx,root,out.staged_path,exec.signal,15*1024*1024)).hash!==out.sha256)throw new Error('OFFICE_NATIVE_ARTIFACT_CHANGED');
  if(!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
  await nativeFolder(root,dirname(record.group_directory));
  if(!record.group_created){await mkdir(resolve(root,record.group_directory));record.group_created=true;await checkpoint();}
  await nativeFolder(root,record.group_directory);
  for(const out of record.outputs){
    exec.signal.throwIfAborted();
    if(out.state==='published'){
      if((await readInput(ctx,root,out.path,exec.signal,15*1024*1024)).hash!==out.sha256)throw new Error('OFFICE_NATIVE_PUBLISHED_CHANGED');
      continue;
    }
    // A crash after writing but before journaling retains a recoverable file.
    let existing;
    try{existing=await readInput(ctx,root,out.path,exec.signal,15*1024*1024);}catch(e){if(!['ENOENT','FS_NOT_FOUND'].includes(e.code))throw e;}
    if(existing){if(out.state!=='publishing'||existing.hash!==out.sha256)throw new Error('OFFICE_NATIVE_GROUP_COLLISION');}
    else{
      out.state='publishing';await checkpoint();
      const bytes=(await readInput(ctx,root,out.staged_path,exec.signal,15*1024*1024)).bytes;
      const handle=await open(resolve(root,out.path),'wx');try{await handle.writeFile(bytes);await handle.sync();}finally{await handle.close();}
    }
    if((await readInput(ctx,root,out.path,exec.signal,15*1024*1024)).hash!==out.sha256)throw new Error('OFFICE_NATIVE_PUBLISH_CHANGED');
    out.state='published';await checkpoint();
  }
  const manifest={run_id:record.id,complete:true,files:record.outputs.map(o=>({path:o.path,sha256:o.sha256})),checked:record.file_check.checked,
    consistency_checked:record.file_check.consistency.length};
  const manifestPath=record.group_directory+'/交付清单.json',bytes=JSON.stringify(manifest,null,2),digest=sha256(bytes);
  let current;
  try{current=await readInput(ctx,root,manifestPath,exec.signal,1024*1024);}catch(e){if(!['ENOENT','FS_NOT_FOUND'].includes(e.code))throw e;}
  if(current){if(record.delivery_manifest_sha256!==digest||current.hash!==digest)throw new Error('OFFICE_NATIVE_MANIFEST_COLLISION');}
  else{
    record.delivery_manifest_sha256=digest;await checkpoint();
    const handle=await open(resolve(root,manifestPath),'wx');try{await handle.writeFile(bytes);await handle.sync();}finally{await handle.close();}
  }
  record.delivery_manifest=manifestPath;record.state='checked';
}
