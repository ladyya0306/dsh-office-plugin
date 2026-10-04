import { readdir,writeFile,readFile } from 'node:fs/promises';
import { resolve,relative,dirname,isAbsolute } from 'node:path';
import { randomUUID } from 'node:crypto';
import { readInput,policyFor,workspaceFor,outside,sha256 } from './runtime.mjs';
import { nativeFolder,nativePath,sealNativeRun,loadNativeRun } from './native_execution.mjs';
import { safeDirectory } from './method_execution.mjs';
import { ordinaryPythonCommand } from './ordinary_python_command.mjs';

async function officeFiles(ctx,root,exec){
  const found=[];
  async function walk(folder,depth){
    if(depth>5)throw new Error('OFFICE_CAPTURE_DEPTH_LIMIT');
    await nativeFolder(root,folder);
    for(const entry of await readdir(resolve(root,folder),{withFileTypes:true})){
      if(entry.name.startsWith('.')||entry.name==='node_modules'||entry.name==='__pycache__')continue;
      const path=relative(root,resolve(root,folder,entry.name)).replaceAll('\\','/');
      if(entry.isSymbolicLink())throw new Error('OFFICE_CAPTURE_LINK: '+path);
      if(entry.isDirectory())await walk(path,depth+1);
      else if(entry.isFile()&&/\.(xlsx|docx)$/i.test(entry.name)){
        if(found.length>=64)throw new Error('OFFICE_CAPTURE_FILE_LIMIT');
        found.push({path,...await readInput(ctx,root,path,exec.signal,15*1024*1024)});
      }
    }
  }
  await walk('.',0);return found;
}

export function installOrdinaryCapture(ctx,config,configuredRoot,learning){
  const pending=new Map();
  const key=exec=>(exec.agent?.session?.id??'agentless-probe')+':'+exec.callId;
  ctx.on('tools/execute',async(exec,next)=>{
    if(exec.name!=='pwsh'||exec.arguments.run_in_background===true)return next();
    const plan=ordinaryPythonCommand(exec.arguments.command,config.pythonPath);
    if(!plan)return next();
    const root=workspaceFor(ctx,configuredRoot,exec,config);policyFor(ctx,root,exec,config,true);
    if(pending.size>=64)throw new Error('OFFICE_CAPTURE_PENDING_LIMIT');
    await learning.beforeRun(root,exec);
    const cwd=resolve(root,exec.arguments.workdir??'.');if(outside(root,cwd))throw new Error('OFFICE_CAPTURE_CWD_OUTSIDE');
    await nativeFolder(root,relative(root,cwd)||'.');
    const full=resolve(cwd,plan.script);if(outside(root,full))throw new Error('OFFICE_CAPTURE_SCRIPT_OUTSIDE');
    const oldOutputs=new Map(),published=new Set();
    const priorState=await learning.read(root,exec);
    for(const candidate of priorState.candidates){
      const prior=await loadNativeRun(ctx,config,root,candidate.run_id,exec);
      if(prior.origin!=='ordinary_pwsh')continue;
      for(const output of prior.outputs){
        if(output.state==='published'){published.add(output.path);continue;}
        if(!['executed','check_failed','failed'].includes(prior.state))continue;
        // Only an artifact actually produced by this session may be revised.
        // It must retain its original output argument position. User sources,
        // templates and previously published results never enter this set.
        const index=prior.argv.findIndex(arg=>resolve(root,prior.cwd,arg).toLowerCase()===resolve(root,output.path).toLowerCase());
        if(index>=0&&plan.argv[index]!==undefined&&resolve(cwd,plan.argv[index]).toLowerCase()===resolve(root,output.path).toLowerCase())oldOutputs.set(output.path,output.sha256);
      }
    }
    for(const path of published)oldOutputs.delete(path);
    const script=nativePath(relative(root,full).replaceAll('\\','/'));
    const id=randomUUID(),folder=resolve(root,'.office-plugin/native-runs',id),bundleRoot=resolve(folder,'bundle');
    const record={id,owner:exec.agent?.session?.id??'agentless-probe',task:'普通Python执行自动留存',candidate:true,
      execution_kind:'native_python',origin:'ordinary_pwsh',host_call_id:exec.callId,state:'observing',
      created_at:new Date().toISOString(),cwd:relative(root,cwd).replaceAll('\\','/')||'.',argv:plan.argv,
      interpreter:config.pythonPath,interpreter_sha256:sha256(await readFile(config.pythonPath)),bundle:[],inputs:[],outputs:[],
      file_check:{complete:false},code_validation:{passed:false},timings_ms:{},replaced_unverified_outputs:[]};
    const local=[script];
    for(const e of await readdir(dirname(full),{withFileTypes:true})){
      if(e.isFile()&&e.name.endsWith('.py')){const p=relative(root,resolve(dirname(full),e.name)).replaceAll('\\','/');if(p!==script)local.push(p);}
    }
    if(local.length>32)throw new Error('OFFICE_CAPTURE_DEPENDENCY_LIMIT');
    for(const path of local){
      const f=await readInput(ctx,root,path,exec.signal,65536),dest=resolve(bundleRoot,path);
      await safeDirectory(root,dirname(dest));await writeFile(dest,f.bytes,{flag:'wx'});
      record.bundle.push({path,sha256:f.hash,snapshot:relative(root,dest).replaceAll('\\','/')});
    }
    const files=await officeFiles(ctx,root,exec);let bytes=0;
    for(const f of files){
      bytes+=f.bytes.length;if(bytes>32*1024*1024)throw new Error('OFFICE_CAPTURE_INPUT_LIMIT');
      const dest=resolve(folder,'input-snapshot',f.path);await safeDirectory(root,dirname(dest));
      await writeFile(dest,f.bytes,{flag:'wx'});
      const item={path:f.path,sha256:f.hash,snapshot:relative(root,dest).replaceAll('\\','/')};
      if(oldOutputs.get(f.path)===f.hash)record.replaced_unverified_outputs.push(item);else record.inputs.push(item);
    }
    await sealNativeRun(config,root,record);
    pending.set(key(exec),{root,record,started:performance.now()});
    return next();
  });
  ctx.on('tools/post-execute',async(exec,result,next)=>{
    const decision=await next(),item=pending.get(key(exec));if(!item)return decision;
    pending.delete(key(exec));const {root,record,started}=item;
    // Keep failure evidence with a fresh cleanup signal; never run business code here.
    const audit={...exec,signal:new AbortController().signal};
    const value=result.value;
    record.finished_at=new Date().toISOString();record.timings_ms.program=performance.now()-started;
    record.execution={exit_code:value?.exitCode??null,timed_out:value?.timedOut===true,aborted:value?.aborted===true,
      sandbox:value?.sandbox??null,stdout_truncated:value?.stdout?.truncated===true};
    record.state=decision.kind==='accept'&&!result.isError&&value?.kind==='foreground'&&value.exitCode===0&&!value.timedOut&&!value.aborted
      &&['partial','enforced'].includes(value.sandbox?.enforcement)&&!value.sandbox?.denied?'executed':'failed';
    const after=await officeFiles(ctx,root,audit),before=new Map(record.inputs.map(f=>[f.path,f.sha256]));
    if(record.inputs.some(f=>after.find(a=>a.path===f.path)?.hash!==f.sha256)){
      record.state='failed';record.error='OFFICE_CAPTURE_INPUT_CHANGED';
    }
    for(const f of record.bundle)if((await readInput(ctx,root,f.path,audit.signal,65536)).hash!==f.sha256){
      record.state='failed';record.error='OFFICE_CAPTURE_CODE_CHANGED';
    }
    for(const f of after.filter(f=>!before.has(f.path))){
      const dest=resolve(root,'.office-plugin/native-runs',record.id,'output-snapshot',f.path);
      await safeDirectory(root,dirname(dest));await writeFile(dest,f.bytes,{flag:'wx'});
      record.outputs.push({path:f.path,sha256:f.hash,state:'unverified',snapshot:relative(root,dest).replaceAll('\\','/')});
    }
    if(!record.outputs.length&&record.state==='executed'){record.state='observed';record.candidate=false;}
    await sealNativeRun(config,root,record);if(record.candidate)await learning.record(root,audit,record);
    if(decision.kind!=='accept')return decision;
    return {...decision,content:[...(decision.content??result.content),{type:'text',text:
      `DSH已自动保留本次实际程序和文件记录，执行ID ${record.id}，状态 ${record.state}，新产物${record.outputs.length}份。`+
      (record.error?`原因：${record.error}。`:'')+
      (record.state==='executed'?'这是待核对程序。用office_program_check_candidate绑定本批完整检查范围，核对通过后office_program_save加入程序库；勿再次执行正确文件。':
       record.state==='observed'?'本次没有新Office产物，只保留读取或核对记录，无需加入程序库。':'本次失败，不加入程序库；按上述错误修正程序，再生成新文件重新执行。')}]};
  });
}
