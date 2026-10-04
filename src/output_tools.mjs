import { defineTool } from '@deepseek-ai/dsh-tools';
import { readInput, runWorker, policyFor,sha256,workspaceFor } from './runtime.mjs';
import { contractInputSchema } from './check_contract.mjs';
import { expandContract, coverageTargets } from './contract_builder.mjs';

export function outputFeedback(result, contract, restoring=false) {
  const required = coverageTargets(contract.coverage).size;
  const checks = result.checks ?? [], failed = checks.filter(c=>!c.passed);
  const missing = Math.max(0, required - (result.checked ?? 0));
  const valuesPassed = checks.length === result.checked && failed.length === 0 && !result.coverage_errors?.length;
  const action = result.complete ?
    {reason:'本次完整范围通过，交付文件。复用候选已记录，方法整理可在后续会话继续。'} :
    missing || result.coverage_errors?.length ?
      {reason:`本次核对${result.checked ?? 0}/${required}项；补齐范围和缺项来源依据，保留已正确候选。`} :
      valuesPassed && result.format_errors?.length ?
        {tool:restoring?undefined:'office_restore_template',reason:restoring?'模板交接仍有格式错误，按format_errors修复具体模板属性。':'值已全通过；用同一清单把候选值交给office_restore_template，保留格式并同步表及筛选范围。'} :
        {reason:'先按失败项核对来源引用、标签及算式。引用写错只修清单；实际值错才修改候选。',failed_targets:failed.map(c=>c.target_ref)};
  return { required, checked:result.checked??0, passed:result.passed??checks.filter(c=>c.passed).length,
    missing, failed:failed.length, format_errors:result.format_errors?.length??0,
    complete:result.complete===true, published:restoring?result.published===true:null,next_action:action };
}

export function normalizeTableRanges(args) {
  if(!Object.hasOwn(args,'table_ranges'))return args;
  const ranges=args.table_ranges;
  if(!ranges || typeof ranges!=='object' || Array.isArray(ranges))throw new Error('OFFICE_TABLE_RANGES_INVALID: 表范围需对象');
  const ordered=value=>Array.isArray(value)?value.map(ordered):value&&typeof value==='object'?
    Object.fromEntries(Object.keys(value).sort().map(key=>[key,ordered(value[key])])):value;
  if(Object.hasOwn(args.contract,'table_ranges') &&
    JSON.stringify(ordered(args.contract.table_ranges))!==JSON.stringify(ordered(ranges)))
    throw new Error('OFFICE_TABLE_RANGES_CONFLICT: 顶层和清单内的表范围不同，请明确本次范围');
  return {...args,contract:{...args.contract,table_ranges:ranges},
    parameter_normalizations:['table_ranges已归入contract.table_ranges；保存方法时采用清单内的相同范围']};
}

export async function checkArtifact(ctx, config, root, args, exec, restoring=false) {
  policyFor(ctx, root, exec, config, restoring);
  if(args.contract_path) {
    if(args.contract)throw new Error('OFFICE_CHECK_CONTRACT_ONE_SOURCE');
    const file=await readInput(ctx,root,args.contract_path,exec.signal,512*1024);
    args={...args,contract:JSON.parse(file.bytes.toString('utf8'))};
  }
  if(!args.contract||typeof args.contract!=='object'||Array.isArray(args.contract))throw new Error('OFFICE_CHECK_CONTRACT_REQUIRED');
  args=normalizeTableRanges(args);
  args={...args,contract:expandContract(args.contract)};
  const read = async path => {
    const format = path.split('.').at(-1).toLowerCase();
    if (!['xlsx','docx'].includes(format)) throw new Error('OFFICE_FORMAT_UNSUPPORTED');
    const item = await readInput(ctx, root, path, exec.signal, 15*1024*1024);
    return { ...item, format, bytes_b64: item.bytes.toString('base64') };
  };
  if (!Array.isArray(args.source_paths) || !args.source_paths.length || args.source_paths.length>8) throw new Error('OFFICE_SOURCE_LIMIT');
  const sources = [];
  for (const path of args.source_paths) sources.push(await read(path));
  const target = await read(args.target_path), artifact = await read(args.output_path);
  if ([...sources,target,artifact].reduce((n,f)=>n+f.bytes.length,0)>28*1024*1024) throw new Error('OFFICE_TOTAL_INPUT_LIMIT');
  const pack = f => ({format:f.format,bytes_b64:f.bytes_b64,filename:f.path.split(/[\\/]/).at(-1)});
  if(restoring && !config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
  const result = await runWorker(ctx,config,root,{operation:restoring?'restore_template':'check_output',sources:sources.map(pack),
    target:pack(target),artifact:pack(artifact),contract:args.contract,...(restoring?{output_name:args.output_name}:{})},exec);
  const verifiedArtifact=restoring&&result.published?await readInput(ctx,root,args.output_name,exec.signal,15*1024*1024):artifact;
  if (result.artifact_sha256 !== verifiedArtifact.hash || result.target_sha256 !== target.hash ||
      JSON.stringify(result.source_sha256)!==JSON.stringify(sources.map(s=>s.hash))) throw new Error('OFFICE_DIGEST_MISMATCH');
  for (const f of [...sources,target,artifact]) if ((await readInput(ctx,root,f.path,exec.signal,15*1024*1024)).hash!==f.hash) throw new Error('OFFICE_INPUT_CHANGED');
  if(restoring&&!config.allowAgentlessProbe)await ctx.get('sessions').flush(exec.agent.session);
  const ordered=value=>Array.isArray(value)?value.map(ordered):value&&typeof value==='object'?
    Object.fromEntries(Object.keys(value).sort().map(k=>[k,ordered(value[k])])):value;
  const feedback=outputFeedback(result,args.contract,restoring);
  return {summary:feedback,...result,next_action:feedback.next_action,...(args.parameter_normalizations?{parameter_normalizations:args.parameter_normalizations}:{}),contract_sha256:sha256(JSON.stringify(ordered(args.contract))),
    confirmed_contract:structuredClone(args.contract),
    required_targets:args.contract.requirements.map(r=>r.target_ref).sort(),output_path:restoring&&result.published?args.output_name:args.output_path,source_paths:args.source_paths,target_path:args.target_path};
}

// Keep the checked relationship for the host; show it only when a candidate is claimed.
const presentCheck=value=>{
  const {confirmed_contract,...shown}=value;
  return [{type:'text',text:JSON.stringify({...shown,checks:value.checks?.filter(c=>!c.passed),passed_checks_omitted:true})}];
};

export function installOutputTools(ctx,config,root) {
  ctx.tools.register(defineTool({name:'office_restore_template',description:
    '检查候选全部值并写入原模板新副本，保留格式和部件，同步声明表及筛选范围。候选、最终文件使用不同新名称；不覆盖原件，不猜错值。支持XLSX和普通Word安全位置。完整清单可用tables展开；用office_help看示例。',
    parameters:{source_paths:{type:'array',required:true,items:{type:'string'}},target_path:{type:'string',required:true},
      output_path:{type:'string',required:true,description:'原生脚本生成的候选文件相对路径'},output_name:{type:'string',required:true,description:'最终工作区普通新文件名，与模板同后缀'},contract:contractInputSchema,
      contract_path:{type:'string',description:'完整contract JSON相对路径，与内联二选一'},script_path:{type:'string',description:'本次实际生成候选文件的脚本相对路径，供复用评估'},
      table_ranges:{type:'json',description:'可选表范围对象；自动归入contract.table_ranges，两个位置不同则报告冲突'}},
    output:{schema:{type:'json'},render:(_a,v)=>presentCheck(v)},timeoutMs:65000,
    execute:(args,exec)=>checkArtifact(ctx,config,workspaceFor(ctx,root,exec,config),args,exec,true)}));
  ctx.tools.register(defineTool({name:'office_check_output',description:
    '按完整任务范围重读来源和产物，核对值、依据及模板。coverage须由任务确定；requirements或tables完整覆盖。长清单用contract_path。引用错只修清单，值错才改文件；按summary.next_action继续。详例用office_help(topic=checklist)。',
    parameters:{source_paths:{type:'array',required:true,items:{type:'string'}},target_path:{type:'string',required:true},
      output_path:{type:'string',required:true},contract:contractInputSchema,
      contract_path:{type:'string',description:'独立完整contract JSON工作区相对路径，与内联contract二选一；可用原生脚本一次生成，避免反复展开长参数'},
      script_path:{type:'string',description:'确实生成本次产物的原生脚本相对路径，用于复用评估；没有则省略'},
      table_ranges:{type:'json',description:'可选表范围对象；自动归入contract.table_ranges，两个位置不同则报告冲突'}},
    output:{schema:{type:'json'},render:(_a,v)=>presentCheck(v)},timeoutMs:65000,
    execute:(args,exec)=>checkArtifact(ctx,config,workspaceFor(ctx,root,exec,config),args,exec)}));
}
