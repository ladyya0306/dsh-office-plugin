import { expandContract, coverageTargets } from './contract_builder.mjs';
import { sha256 } from './runtime.mjs';

// A current-batch execution plan derived from the SAME confirmed relationship
// as independent checking. No values/expected answers are generated here.
export function taskBinding({contract,sourcePaths,targetPath,outputPath,inputHashes,relationshipRule}) {
  const expanded=expandContract(contract);
  if(!Array.isArray(sourcePaths)||!sourcePaths.length||new Set(sourcePaths).size!==sourcePaths.length)
    throw new Error('OFFICE_TASK_BINDING_SOURCES');
  const rules=expanded.requirements.flatMap(rule=>[rule,...(rule.operation==='compose'?rule.components??[]:[])]);
  for(const rule of rules)for(const ref of rule.source_refs??[]) {
    const index=/^s(\d+)::/.exec(ref);
    if(!index||Number(index[1])>=sourcePaths.length)throw new Error('OFFICE_TASK_BINDING_SOURCE_REF');
  }
  return {schema:'office_task_binding_v1',source_paths:sourcePaths,target_path:targetPath,output_path:outputPath,
    input_sha256:inputHashes,requirements:expanded.requirements,coverage:expanded.coverage,
    table_ranges:expanded.table_ranges??{},allow_existing_values:expanded.allow_existing_values??false,
    required_count:expanded.preserve_only?0:coverageTargets(expanded.coverage).size,
    preserve_only:expanded.preserve_only===true,contract_sha256:sha256(JSON.stringify(expanded)),
    ...(relationshipRule?.schema==='office_relationship_rule_v1'?{
      relationship_rule:structuredClone(relationshipRule),relationship_rule_sha256:sha256(JSON.stringify(relationshipRule)),
      source_sheet:relationshipRule.source_sheet,target_sheet:relationshipRule.target_sheet,
      keys:structuredClone(relationshipRule.keys),fields:structuredClone(relationshipRule.fields),
      ...(relationshipRule.table_name?{table_name:relationshipRule.table_name}:{})}:{}),
    scope:'本批位置和处理关系来自确认清单；检查仍独立读取真实来源和实际产物。该文件不含预期答案。'};
}
