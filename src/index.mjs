import { defineTool } from '@deepseek-ai/dsh-tools';
import { randomUUID } from 'node:crypto';
import { isAbsolute, resolve } from 'node:path';
import { readInput, runWorker, policyFor, workspaceFor } from './runtime.mjs';
import { presentOfficeResult } from './presentation.mjs';
import { installStructured } from './structured.mjs';
import { installOutputTools } from './output_tools.mjs';
import { installExecutableMethods } from './executable_methods.mjs';
import { installReuseMonitor } from './reuse_monitor.mjs';

export const name = 'office-enhancement';
export const inject = ['tools', 'fs', 'shell', 'sandboxPolicy'];
const output = { schema: { type: 'json' }, render: (_args, value) => [{ type: 'text', text: JSON.stringify(presentOfficeResult(value)) }] };
const string = description => ({ type: 'string', required: true, description });
const lifetimeMs = 10 * 60 * 1000;
const evidenceInstruction = '只使用本次任务可核对的来源，不得无中生有，不用常识补缺值，不凭相似标签强行匹配。来源、主体、单位或目标位置有实质歧义时先用可用工具核实；仍不确定则暂停相关字段写入，向用户说明需确认内容，最终如实列出未解决项。';

export function apply(ctx, config = {}) {
  for (const key of ['workspaceRoot', 'pythonPath', 'workerPath']) {
    if (!config[key] || !isAbsolute(config[key])) throw new Error(`Office plugin requires absolute ${key}`);
  }
  if (config.workspaceMode && !['fixed','session'].includes(config.workspaceMode)) throw new Error('OFFICE_WORKSPACE_MODE_INVALID');
  const configuredRoot = resolve(config.workspaceRoot);
  const reuse=installReuseMonitor(ctx, config, configuredRoot);
  installExecutableMethods(ctx, config, configuredRoot,reuse);
  if (config.outputCheckTools !== false) installOutputTools(ctx, config, configuredRoot);
  if (config.structuredTools !== false) installStructured(ctx, config, configuredRoot);
  if (config.legacyTools === false) return;
  const inspections = new Map();
  const plans = new Map();
  const owner = exec => exec.agent?.session?.id ?? 'agentless-probe';
  const get = (map, id, exec) => {
    const record = map.get(id);
    if (!record || record.owner !== owner(exec)) throw new Error('OFFICE_ID_UNKNOWN: 标识不存在或属于其他会话');
    if (record.root!==workspaceFor(ctx,configuredRoot,exec,config)) throw new Error('OFFICE_WORKSPACE_CHANGED: 请在当前工作区重新读取');
    if (Date.now() - record.created > lifetimeMs) { map.delete(id); throw new Error('OFFICE_ID_EXPIRED: 请重新读取来源'); }
    return record;
  };
  const prune = map => {
    for (const [id, record] of map) if (Date.now() - record.created > lifetimeMs) map.delete(id);
    if (map.size >= 32) throw new Error('OFFICE_TASK_LIMIT: 未过期任务最多32项');
  };
  const requestFor = (record, operation) => ({ operation, format: record.format,
    source_b64: record.source.bytes.toString('base64'), target_b64: record.target.bytes.toString('base64') });
  const unchanged = async (record, exec) => {
    const root=workspaceFor(ctx,configuredRoot,exec,config);
    for (const original of [record.source, record.target]) {
      const current = await readInput(ctx, root, original.path, exec.signal);
      if (current.hash !== original.hash) throw new Error('OFFICE_INPUT_CHANGED: 原件变化，请新建读取任务，不沿用原计划');
    }
  };
  const register = (toolName, description, parameters, execute, parallelSafe = false) => ctx.tools.register(defineTool({
    name: toolName, description, parameters, output, timeoutMs: 65000, parallelSafe, execute:(args,exec)=>execute(args,exec,workspaceFor(ctx,configuredRoot,exec,config)) }));

  // This diagnostic tool is disabled in an ordinary profile.
  if (config.allowAgentlessProbe) register('office_probe_readonly', '仅供隔离试验的只读摘要。', { relative_path: string('工作区相对文件路径') }, async (args, exec, root) => {
    const file = await readInput(ctx, root, args.relative_path, exec.signal);
    return { status: 'read_only_probe', bytes: file.bytes.length, sha256: file.hash };
  }, true);

  register('office_inspect', evidenceInstruction + '读取一个来源和一个目标 XLSX 或 DOCX，返回候选键值和可填空位。候选不是已确认业务事实，须结合相邻主体、标签、单位和说明判断；公式不取猜测值。当前仅同格式。', {
    source_path: string('来源 XLSX/DOCX 在工作区内的相对路径'), target_path: string('目标模板 XLSX/DOCX 在工作区内的相对路径'),
  }, async (args, exec, root) => {
    policyFor(ctx, root, exec, config, false);
    prune(inspections);
    const format = args.target_path.split('.').at(-1).toLowerCase();
    if (!['xlsx', 'docx'].includes(format) || args.source_path.split('.').at(-1).toLowerCase() !== format) throw new Error('OFFICE_FORMAT_UNSUPPORTED: 当前仅同格式 XLSX/DOCX');
    const source = await readInput(ctx, root, args.source_path, exec.signal);
    const target = await readInput(ctx, root, args.target_path, exec.signal);
    const record = { root, source, target, format, owner: owner(exec), created: Date.now() };
    const value = await runWorker(ctx, config, root, requestFor(record, 'inspect'), exec);
    if (value.source_sha256 !== source.hash || value.target_sha256 !== target.hash) throw new Error('OFFICE_DIGEST_MISMATCH');
    record.inspection = value;
    const id = randomUUID(); inspections.set(id, record);
    return { inspect_id: id, expires_in_seconds: 600, ...value };
  });

  register('office_prepare_fill', evidenceInstruction + '仅提出来源到目标的字段对应关系。程序重读真实值、换算单位、检查目标并返回可复审计划，不写文件。禁止从说明、示例、其他主体或过时版本取值；不确定时先询问用户。', {
    inspect_id: string('office_inspect 返回的标识'),
    mappings: { type: 'array', required: true, items: { type: 'object', additionalProperties: false, properties: {
      source_ref: string('返回候选中的位置标识，必须原样使用'), target_ref: string('返回空位中的位置标识，必须原样使用'),
    } } },
  }, async (args, exec, root) => {
    const record = get(inspections, args.inspect_id, exec);
    if (!args.mappings.length || args.mappings.length > 100) throw new Error('OFFICE_MAPPING_LIMIT: 需要1至100项对应关系');
    const sourceRefs = new Set(record.inspection.source_candidates.map(x => x.source_ref));
    const targetRefs = new Set(record.inspection.target_slots.map(x => x.target_ref));
    for (const entry of args.mappings) if (!sourceRefs.has(entry.source_ref) || !targetRefs.has(entry.target_ref)) throw new Error('OFFICE_REF_UNKNOWN: 引用不属于本次读取结果');
    await unchanged(record, exec);
    prune(plans);
    const value = await runWorker(ctx, config, root, { ...requestFor(record, 'prepare'), mappings: args.mappings }, exec);
    if (!value.changes?.length) throw new Error('OFFICE_EMPTY_PLAN');
    const id = randomUUID();
    plans.set(id, { ...record, changes: value.changes, created: Date.now(), state: 'ready' });
    return { plan_id: id, expires_in_seconds: 600, ...value, authorization: '此计划不代表用户批准；执行仍须通过 DSH 原生权限检查。' };
  });

  register('office_apply_fill', evidenceInstruction + '按已校验计划写入新文件，不覆盖原件。成功结果的verification已经重开检查计划内值和列出的格式属性；无失败或矛盾时不必重复inspect或另写脚本重开。只对applied列出的位置报告完成，候选总数不是业务必填数；未知主体或缺资料仍须询问。仍服从DSH原生权限。', {
    plan_id: string('office_prepare_fill 返回的标识'), output_name: string('新文件名，例如 filled-report.xlsx；后缀与模板一致，不能含路径'),
  }, async (args, exec, root) => {
    const plan = get(plans, args.plan_id, exec);
    policyFor(ctx, root, exec, config, true);
    if (!/^[A-Za-z0-9][A-Za-z0-9._-]{0,119}\.(xlsx|docx)$/.test(args.output_name) || args.output_name.includes('..') || !args.output_name.endsWith(`.${plan.format}`)) throw new Error('OFFICE_OUTPUT_NAME: 首版需要不含路径的英文字母/数字文件名，后缀与模板一致');
    if (plan.state !== 'ready') throw new Error('OFFICE_PLAN_CONSUMED: 每个计划只允许一次执行尝试');
    plan.state = 'executing';
    if (!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);
    await unchanged(plan, exec);
    const value = await runWorker(ctx, config, root, { ...requestFor(plan, 'apply'), changes: plan.changes, output_name: args.output_name }, exec);
    const outputFile = await readInput(ctx, root, args.output_name, exec.signal, 15 * 1024 * 1024);
    if (outputFile.hash !== value.output_sha256) throw new Error('OFFICE_OUTPUT_DIGEST_MISMATCH: 已产生文件待检查，未报告成功');
    await unchanged(plan, exec);
    if (!config.allowAgentlessProbe) await ctx.get('sessions').flush(exec.agent.session);
    plan.state = 'done';
    return { ...value, output_path: args.output_name, provenance: plan.changes, sandbox_notice: '原生执行器限制写入；Windows 为部分约束，不代表完整操作系统沙箱。' };
  });
}
