import { writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { readInput, sha256 } from './runtime.mjs';
import { safeDirectory } from './method_execution.mjs';

// The checked relation is immutable evidence, not an executable method or approval.
export async function saveReuseEvidence(ctx,root,exec,event,value) {
  if (!value.confirmed_contract) return null;
  const declared=exec.arguments.script_path
    ? await readInput(ctx,root,exec.arguments.script_path,exec.signal,128*1024) : null;
  const snapshot={contract:value.confirmed_contract,contract_sha256:value.contract_sha256,
    source_paths:value.source_paths,source_sha256:value.source_sha256,
    target_path:value.target_path,target_sha256:value.target_sha256,
    output_path:value.output_path,artifact_sha256:value.artifact_sha256,
    declared_script:declared?{path:declared.path,sha256:declared.hash,code:declared.bytes.toString('utf8')}:null,
    script_evidence:'脚本由本次工具参数指定；摘要固定其内容版本，不代表宿主已证明该脚本产生了产物'};
  const folder=resolve(root,'.office-plugin','reuse');
  await safeDirectory(root,folder);exec.signal.throwIfAborted();
  const relative=`.office-plugin/reuse/${event.id}.json`,bytes=Buffer.from(JSON.stringify(snapshot));
  if(bytes.length>768*1024)throw new Error('OFFICE_REUSE_EVIDENCE_LIMIT');
  await writeFile(resolve(root,relative),bytes,{flag:'wx'});
  return {path:relative,sha256:sha256(bytes),contract_sha256:value.contract_sha256,
    script_sha256:declared?.hash??null};
}

export async function loadReuseEvidence(ctx,root,exec,event) {
  if(!event.evidence)return null;
  const saved=await readInput(ctx,root,event.evidence.path,exec.signal,768*1024);
  if(saved.hash!==event.evidence.sha256)throw new Error('OFFICE_REUSE_EVIDENCE_CHANGED');
  const snapshot=JSON.parse(saved.bytes.toString('utf8'));
  const inputs=[...(snapshot.source_paths??[]).map((path,i)=>({path,sha256:snapshot.source_sha256[i]})),
    {path:snapshot.target_path,sha256:snapshot.target_sha256}];
  const changed=[];
  for(const file of inputs)if((await readInput(ctx,root,file.path,exec.signal)).hash!==file.sha256)changed.push(file.path);
  return {confirmed_contract:snapshot.contract,contract_sha256:snapshot.contract_sha256,
    evidence_path:event.evidence.path,script_sha256:snapshot.declared_script?.sha256??null,
    current_inputs_match:changed.length===0,changed_inputs:changed,
    next:changed.length?'输入发生变化，按本次资料修订关系后验证。':'沿用已完整核对的来源关系整理方法；代码适配后验证。'};
}
