import { createHash } from 'node:crypto';
import { isAbsolute, relative, resolve } from 'node:path';

export const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
export const quotePwsh = value => "'" + value.replaceAll("'", "''") + "'";
export const outside = (root, path) => {
  const part = relative(root, path);
  return part === '..' || part.startsWith('..\\') || part.startsWith('../') || isAbsolute(part);
};

// Only the host's session policy selects a workspace; tool arguments never do.
export function workspaceFor(ctx, configuredRoot, exec, config) {
  if (config.workspaceMode !== 'session') return configuredRoot;
  if (!exec.agent?.session) throw new Error('OFFICE_SESSION_REQUIRED: 切换工作区需要真实 DSH 会话');
  const policy = ctx.sandboxPolicy.resolve({ session: exec.agent.session });
  if (!policy.workspaceRoot || !isAbsolute(policy.workspaceRoot)) throw new Error('OFFICE_WORKSPACE_REQUIRED: 当前会话需要本地绝对工作区');
  const root = resolve(policy.workspaceRoot);
  policyFor(ctx, root, exec, config, false);
  return root;
}

export async function readInput(ctx, root, file, signal, maxBytes = 10 * 1024 * 1024) {
  if (!file || isAbsolute(file) || file.includes(':')) throw new Error('OFFICE_PATH_DENIED: 需要工作区相对路径');
  const path = resolve(root, file);
  if (outside(root, path)) throw new Error('OFFICE_PATH_DENIED: 路径越出工作区');
  signal.throwIfAborted();
  const target = await ctx.fs.resolve(path, { cwd: root });
  if (!isAbsolute(target.targetKey)) throw new Error('OFFICE_BACKEND_UNSUPPORTED: 首版仅支持本地文件后端');
  if (outside(root, target.targetKey)) throw new Error('OFFICE_PATH_DENIED: 文件链接指向工作区外');
  const bytes = await ctx.fs.readBytes(target, signal, maxBytes);
  signal.throwIfAborted();
  return { bytes, hash: sha256(bytes), path: file };
}

export function policyFor(ctx, root, exec, config, writing) {
  if (!exec.agent && !config.allowAgentlessProbe) throw new Error('OFFICE_SESSION_REQUIRED: 需要真实 DSH 会话');
  const policy = ctx.sandboxPolicy.resolve({ session: exec.agent?.session });
  if (resolve(policy.workspaceRoot).toLowerCase() !== root.toLowerCase()) throw new Error('OFFICE_WORKSPACE_MISMATCH: 会话工作区与插件配置不一致');
  if (!['workspace-write', 'read-only'].includes(policy.mode)) throw new Error('OFFICE_POLICY_DENIED: 不接受不受约束的执行模式');
  if (writing && policy.mode !== 'workspace-write') throw new Error('OFFICE_WRITE_DENIED: 当前会话只读');
  if (writing && !config.allowAgentlessProbe && (!ctx.get('sessions') || !ctx.get('sessionPersistence'))) throw new Error('OFFICE_AUDIT_UNAVAILABLE: 写入要求 DSH 会话持久化服务');
  return policy;
}

export async function runWorker(ctx, config, root, request, exec) {
  const policy = policyFor(ctx, root, exec, config, ['apply', 'apply_v3', 'fill_v3', 'records_store','restore_template'].includes(request.operation));
  exec.signal.throwIfAborted();
  const spec = ctx.shell.resolve({
    command: `& ${quotePwsh(config.pythonPath)} -I -B -X utf8 ${quotePwsh(config.workerPath)}`,
    workdir: root, timeoutMs: 60000, onExpiry: 'kill', signal: exec.signal,
    sandboxPolicy: policy, stdin: JSON.stringify(request), stdoutMaxBytes: 1024 * 1024,
  });
  const result = await (await ctx.shell.execute(spec)).result();
  exec.signal.throwIfAborted();
  if (result.timedOut || result.aborted) throw new Error('OFFICE_WORKER_STOPPED: 执行超时或取消；重试前须检查是否留下产物');
  if (!result.sandbox || !['partial', 'enforced'].includes(result.sandbox.enforcement)) throw new Error('OFFICE_CONFINEMENT_UNAVAILABLE: 原生执行器没有报告写入约束');
  if (result.stdout.truncated) throw new Error('OFFICE_WORKER_OUTPUT_LIMIT: 程序结果超限');
  let value;
  try { value = JSON.parse(result.stdout.text); }
  catch { throw new Error(`OFFICE_WORKER_PROTOCOL: 程序未返回 JSON，exitCode=${result.exitCode}`); }
  if (!value.ok || result.exitCode !== 0) throw new Error(`OFFICE_WORKER_${value.error?.code ?? 'FAILED'}: ${value.error?.message ?? '执行失败'}${value.error?.recoverable ? '；保留产物待人工检查' : ''}`);
  return value;
}
