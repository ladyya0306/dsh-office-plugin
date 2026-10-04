import { defineTool } from '@deepseek-ai/dsh-tools';
import { MethodStore } from './method_store.mjs';
import { executeMethod, runtimeIdentity } from './method_execution.mjs';
import { readInput, policyFor,outside,workspaceFor } from './runtime.mjs';
import { contractInputSchema,conditionsSchema } from './check_contract.mjs';
import { summarizeMethods } from './method_lookup.mjs';
import { installOfficeHelp } from './office_help.mjs';
import { expandContract } from './contract_builder.mjs';
import { resumeGroup } from './native_group_resume.mjs';
import { resolve,dirname } from 'node:path';

const str=description=>({type:'string',required:true,description});
const output={schema:{type:'json'},render:(_a,v)=>[{type:'text',text:JSON.stringify(v)}]};
export function installExecutableMethods(ctx,config,configuredRoot,reuse) {
  const register=(name,description,parameters,execute,view)=>ctx.tools.register(defineTool({name,description,parameters,
    output:view?{...output,render:(_a,value)=>[{type:'text',text:JSON.stringify(view(value))}]}:output,timeoutMs:65000,execute:(args,exec)=>execute(args,exec,workspaceFor(ctx,configuredRoot,exec,config))}));
  installOfficeHelp(register,config);
  if(!config.methodDirectory || !config.methodRuntimeLock) return;
  const store=new MethodStore(config.methodDirectory);
  const audit=async(exec)=>{if(!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);};
  register('office_method_list','按用途关键词查找方法；默认最多5条，优先启用版本。关键词相似不代表适用，draft不能直接复用。',{
    query:{type:'string',description:'可选用途关键词'},limit:{type:'number',description:'1至10整数，默认5'},
    include_inactive:{type:'boolean',description:'需要查看草稿或旧版本时传true；默认有启用候选就仅返回启用候选，没有时仍显示草稿'},
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,false);const all=await store.list();
    return summarizeMethods(all,args.query,args.limit??5,args.include_inactive??false);
  });
  register('office_method_export','复制已保存Python到工作区新文件供纠正；修改后用save(parent_id=旧ID)保存新版本并重新验证。',{
    method_id:str('方法ID'),script_name:str('普通新Python文件名，例如fix_method.py，不能含路径'),
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,true);
    if(!/^[^\\/:*?"<>|\x00-\x1f]{1,116}\.py$/.test(args.script_name))throw new Error('OFFICE_METHOD_SCRIPT_NAME');
    const record=await store.read(args.method_id);await audit(exec);
    const target=await ctx.fs.resolve(root+'/'+args.script_name,{cwd:root});
    if(outside(root,target.targetKey))throw new Error('OFFICE_METHOD_SCRIPT_OUTSIDE');
    await ctx.fs.writeText(target,record.code,{kind:'createIfAbsent'},exec.signal);await audit(exec);
    return {method_id:record.meta.id,script_path:args.script_name,code_sha256:record.meta.code_sha256,next:'修改后保存新版本，再验证新ID；旧验证不能沿用'};
  });
  register('office_method_save','保存不含当次客户事实的Python草稿，入口为def run(inputs)。代码只能调用受信任办公回调；conditions/contract声明适用结构与完整检查。版本不可变，修改后用save(parent_id=旧ID)并重新验证；不适配脚本走原生。',{
    script_path:str('工作区相对Python文件'),name:str('用户看得懂的用途'),purpose:str('处理步骤、范围及不适用情形'),
    conditions:{type:'json',description:'单模板结构条件或office_group_conditions_v1目标条件清单'},contract:contractInputSchema,
    metadata_path:{type:'string',description:'可用原生脚本一次生成工作区JSON，内容为{conditions,contract}，传相对路径避免反复在工具参数展开大清单；与内联两项二选一'},
    reuse_event_id:{type:'string',description:'真实复用评估事件ID，保存后把待处理事件关联到该代码版本'},
    parent_id:{type:'string',description:'纠正旧方法时填写旧方法ID，新版本单独保存'},
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,true);await audit(exec);
    const file=await readInput(ctx,root,args.script_path,exec.signal,65536);
    let metadata={conditions:args.conditions,contract:args.contract};
    if(args.metadata_path) {
      if(args.conditions||args.contract)throw new Error('OFFICE_METHOD_METADATA_ONE_SOURCE');
      metadata=JSON.parse((await readInput(ctx,root,args.metadata_path,exec.signal,512*1024)).bytes.toString('utf8'));
      if(!metadata||Object.keys(metadata).some(k=>!['conditions','contract'].includes(k)))throw new Error('OFFICE_METHOD_METADATA_FIELDS');
    }
    if(metadata.contract?.schema!=='office_group_contract_v1')metadata.contract=expandContract(metadata.contract);
    if(args.reuse_event_id)await reuse.check(args.reuse_event_id,exec,metadata.contract);
    const meta=await store.create({...args,...metadata,code:file.bytes.toString('utf8')});
    if(args.reuse_event_id)await reuse.saved(args.reuse_event_id,meta.id,exec);
    else if(args.parent_id)await reuse?.revised(args.parent_id,meta.id,exec,metadata.contract);
    await audit(exec);
    return {method_id:meta.id,status:meta.status,code_sha256:meta.code_sha256,
      next:'用本次样本验证，并按出错后果补必要变化或反例；未启用前仍可交付本次正确文件'};
  });
  const runParameters={method_id:str('保存的方法ID'),source_paths:{type:'array',required:true,items:{type:'string'}},
    target_path:{type:'string',description:'单模板相对路径'},target_paths:{type:'array',items:{type:'string'},description:'整组模板清单'},
    inputs:{type:'object',additionalProperties:true,properties:{},description:'单份output_name，整组output_names及新output_directory；没有直接打开路径的能力'}};
  register('office_method_validate','在独立暂存区重读本次来源与产物，按保存的完整contract检查。通过只覆盖清单范围；输入、结构或高风险算式变化需重新验证。',runParameters,async(args,exec,root)=>{
    const record=await store.read(args.method_id);
    const value=await executeMethod(ctx,config,root,args,exec,record,false);
    const verification=value.generated?.[0]?.verification;
    const complete=value.complete&&value.generated.length===(record.meta.contract.schema==='office_group_contract_v1'?record.meta.contract.artifacts.length:1);
    const proof={tested_at:new Date().toISOString(),complete,checked:value.verification?.checked??verification?.checked??0,passed:value.verification?.passed??verification?.passed??0,
      code_sha256:record.meta.code_sha256,runtime_fingerprint:value.runtime_fingerprint,
      source_sha256:value.source_sha256,target_sha256:value.target_sha256};
    await store.update(record.meta.id,record.meta.revision,m=>{m.proofs.push(proof);m.proofs=m.proofs.slice(-32);return m;});
    await audit(exec);return {...value,complete,status:'验证记录已保存，尚未自动启用',
      next:!complete?'验证未通过，不能启用；先修正具体失败，当前任务仍可走原生。':record.meta.status==='active'?
        '方法已启用。这次只生成了暂存验证文件；需要交付工作区新文件时调用office_method_run，不必重复保存或再启用。':
        '草稿验证通过，还未启用，也未发布指定工作区文件。下一步实际调用office_method_enable；需要交付时再office_method_run。文字声明不能代替操作。'};
  });
  register('office_method_resume','恢复已整组核对的回调文件发布。核对当前文件后继续，不重跑方法或覆盖已有产物。',{run_id:str('本会话真实执行ID')},
    async(args,exec,root)=>{
      const run=await resumeGroup(ctx,{...config,nativeIdentityPath:config.nativeIdentityPath??resolve(dirname(config.workerPath),'native_identity.py')},root,args,exec);
      return {run_id:run.id,state:run.state,outputs:run.outputs,delivery_manifest:run.delivery_manifest??null,error:run.error??null};
    });
  register('office_method_enable','启用代码与解释器版本一致且完整验证通过的方法；说明覆盖、限制和理由。缺少业务依据或已停用版本不得启用。',{
    method_id:str('方法ID'),validation_summary:str('已检查的变化和错误风险，尚未支持的范围'),
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,true);await audit(exec);
    const {meta}=await store.read(args.method_id),{fingerprint}=await runtimeIdentity(config);
    if(meta.status==='disabled') throw new Error('OFFICE_METHOD_DISABLED_VERSION: 保存修正后的新版本再启用');
    const latest=meta.proofs.at(-1);
    if(!latest?.complete || latest.runtime_fingerprint!==fingerprint || latest.code_sha256!==meta.code_sha256)
      throw new Error('OFFICE_METHOD_UNVERIFIED: 最新一次验证必须完整通过且版本一致');
    if(!args.validation_summary.trim() || args.validation_summary.length>2000) throw new Error('OFFICE_METHOD_VALIDATION_SUMMARY_REQUIRED');
    const enabled=await store.activate(meta,args.validation_summary);
    await reuse?.enabled(meta.id,exec,latest);
    await audit(exec);return {method_id:enabled.id,status:enabled.status,next:'下次会话先查用途及本次结构，再调用office_method_run。来源值必须读取本次文件。'};
  });
  register('office_method_run','执行启用方法：重读本次来源，按contract检查并发布工作区新文件。不得覆盖原件、猜缺失值或串主体；条件不符交还DSH，验收失败停用版本。',runParameters,async(args,exec,root)=>{
    const record=await store.read(args.method_id);
    if(record.meta.status!=='active') throw new Error('OFFICE_METHOD_NOT_ACTIVE');
    const value=await executeMethod(ctx,config,root,args,exec,record,true);
    if(!value.complete&&value.state!=='recovery_required'&&(record.meta.contract.schema!=='office_group_contract_v1'||value.verification?.artifacts)) await store.update(record.meta.id,record.meta.revision,m=>({...m,status:'disabled',failure:'本次输出检查失败',failed_at:new Date().toISOString()}));
    const successful=value.complete&&value.published;
    return {...value,reused:successful,method_executed:true,checked_coverage:record.meta.contract.coverage??record.meta.contract.artifacts.map(a=>a.contract.coverage),
      verification_summary:{origin:'插件宿主独立检查程序，依据保存的清单重读本次来源和实际产物，不是方法程序自报',
        checked: value.verification?.checked??value.generated?.[0]?.verification?.checked??0,passed:value.verification?.passed??value.generated?.[0]?.verification?.passed??0,
        coverage:record.meta.contract.coverage??record.meta.contract.artifacts.map(a=>a.contract.coverage),template_checked:record.meta.contract.preserve_template!==false,
        unchecked:'清单以外的业务含义、未声明范围和打印分页没有因此通过'},
      reuse_review:{new_method_needed:successful?false:null,event_id:null,reason:successful?
        '本次直接执行既有方法且所列检查通过；运行不产生新增入库事件，方法ID不能当事件ID。新要求或具体错误另行评估。':
        '方法执行过但本次未通过；不能判断无需新方法。查看具体失败，区分本次数据问题、方法错误和检查依据错误。方法ID不能当事件ID。'},
      next:value.complete?'已读取本次来源并检查所列范围及模板保留项。若覆盖本次任务，直接报告文件和检查范围；无需重复读取、写入、验证或新建同用途方法。任务另有要求或疑点才补查。':'结果未通过，未发布合格文件；修复或走原生。',
      scope:'只对所列检查范围完成；不能据此声称整项陌生业务完成'};
  },value=>({...value,generated:value.generated?.map(item=>({...item,verification:item.verification?{
    ...item.verification,checks:item.verification.checks?.filter(check=>!check.passed),passed_checks_omitted:true}:item.verification}))}));
  register('office_method_disable','停用有结果或解释错误的代码版本并保存纠错记录；其他方法不受影响，修正后保存新版本再验证。',{
    method_id:str('确实受影响的方法ID'),reason:str('用户指出的错误及证据，不能保存当次客户事实'),
  },async(args,exec,root)=>{
    policyFor(ctx,root,exec,config,true);await audit(exec);const {meta}=await store.read(args.method_id);
    await store.update(meta.id,meta.revision,m=>({...m,status:'disabled',failure:args.reason,failed_at:new Date().toISOString()}));
    await audit(exec);return {method_id:meta.id,status:'disabled',next:'纠错记录和新代码验证分别保存，不把旧错误版本重新自动启用'};
  });
}
