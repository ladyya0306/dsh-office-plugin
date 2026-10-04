import { performance } from 'node:perf_hooks';
import { runWorker, sha256 } from './runtime.mjs';

// Only current-batch paths and business selectors enter this request. The
// reusable rule carries labels and relationships, never last batch's locations.
export async function bindGroupRule(ctx, config, root, args, saved, record, inputFiles, exec) {
  const rule = saved?.meta.execution.relationship_rule ?? args.relationship_rule;
  if (rule === undefined) return args;
  if (!['office_group_relationship_rule_v1','office_group_composition_rule_v1'].includes(rule?.schema))
    throw new Error('OFFICE_GROUP_RELATIONSHIP_SCHEMA');
  const artifacts = args.check?.artifacts, sources = args.check?.source_paths;
  if (!Array.isArray(artifacts) || artifacts.length < 2 || artifacts.length > 15 ||
      artifacts.length !== args.outputs.length || artifacts.some((item, i) => item?.output_index !== i) ||
      !Array.isArray(sources) || sources.length !== 1 || new Set(sources).size !== sources.length)
    throw new Error('OFFICE_GROUP_RELATIONSHIP_INPUTS');
  if (!saved && (args.check.consistency !== undefined || artifacts.some(item =>
    item.contract !== undefined || item.conditions !== undefined)))
    throw new Error('OFFICE_GROUP_RELATIONSHIP_ONE_RULE: 长期规则与固定位置清单二选一');
  const bindingIndex = saved?.meta.execution.relationship_arg_index ?? args.relationship_arg_index;
  if (bindingIndex === undefined)
    throw new Error('OFFICE_RELATIONSHIP_ARGUMENT_REQUIRED: 程序需接收本批处理关系参数');
  const pack = path => {
    // executeNative has already read and validated every input path through
    // the workspace filesystem. A spelling change cannot select another file.
    const item = inputFiles.get(path?.replaceAll('\\', '/'));
    if (!item) throw new Error('OFFICE_NATIVE_CHECK_INPUTS');
    return { format: path.split('.').at(-1).toLowerCase(), bytes_b64: item.bytes.toString('base64') };
  };
  const sourceItems = sources.map(pack), targets = artifacts.map(item => pack(item.target_path));
  const started = performance.now();
  const bound = await runWorker(ctx, config, root, {
    operation: rule.schema === 'office_group_composition_rule_v1' ? 'bind_group_composition' : 'bind_group_relationship',
    rule, sources: sourceItems, targets,
    ...(rule.schema === 'office_group_composition_rule_v1' ? {} : {record_selection: args.record_selection ?? {}}),
  }, exec);
  const hash = path => inputFiles.get(path.replaceAll('\\', '/')).hash;
  if (JSON.stringify(bound.source_sha256) !== JSON.stringify(sources.map(hash)) ||
      JSON.stringify(bound.target_sha256) !== JSON.stringify(artifacts.map(item => hash(item.target_path))) ||
      !Array.isArray(bound.artifacts) || bound.artifacts.length !== artifacts.length ||
      bound.artifacts.some((item, i) => item.output_index !== i) ||
      !Array.isArray(bound.consistency) || !bound.consistency.length)
    throw new Error('OFFICE_GROUP_RELATIONSHIP_DIGEST_MISMATCH');
  record.relationship_rule = structuredClone(rule);
  record.record_selection = structuredClone(args.record_selection ?? {});
  record.relationship_binding = {
    rule_sha256: sha256(JSON.stringify(rule)), record_count: bound.record_count,
    field_count: bound.field_count, artifact_count: bound.artifacts.length, locations: bound.locations,
  };
  record.timings_ms.relationship_binding = performance.now() - started;
  return { ...args, check: { source_paths: sources, artifacts: bound.artifacts.map((item, i) => ({
    ...item, target_path: artifacts[i].target_path,
  })), consistency: bound.consistency } };
}
