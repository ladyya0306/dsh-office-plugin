import Schema from '@deepseek-ai/schemastery';
import { isVolatile } from '@deepseek-ai/cosmokit';
import { dirname, isAbsolute, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as Office from './refactor_index.mjs';
export const name='dsh-office-workflows';
export const inject=Office.inject;
// Settings schemas travel to the desktop page; a transform cannot carry closures.
const pathField = description => Schema.string().default('')
  .pattern(/^(?:$|[A-Za-z]:[\\/]|[\\/])/)
  .description(description + '；必须为绝对路径').volatile();
export const Config = Schema.object({
  workspaceRoot: pathField('允许处理文件的工作区'),
  pythonPath: pathField('DSH 自带 Python 解释器'),
  methodDirectory: pathField('可复用程序库；不放客户数据'),
  nativeStateDirectory: pathField('任务执行和检查记录目录'),
  workspaceMode: Schema.union(['fixed', 'session']).default('session').description('固定目录或当前会话工作区').volatile(),
  nativeScriptMode: Schema.union(['disabled', 'personal-compatible']).default('disabled').description('程序执行：关闭或个人电脑兼容验证').volatile(),
  learningEnabled: Schema.boolean().default(false).description('保留执行与修正的程序，检查通过才加入程序库').volatile(),
  captureOrdinaryPython: Schema.boolean().default(false).description('普通 Python 执行自动保留程序与文件记录').volatile(),
  outputCheckTools: Schema.boolean().default(true).description('提供独立产物检查工具').volatile(),
});
const paths = ['workspaceRoot', 'pythonPath', 'methodDirectory', 'nativeStateDirectory'];
export async function apply(ctx, config) {
  const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
  let business, closed = false, generation = 0;
  const values = () => Object.fromEntries(Object.entries(config).map(([key, value]) => [key, isVolatile(value) ? value.get() : value]));
  let pending = Promise.resolve();
  const refresh = () => {
    const requested = ++generation;
    const prior = business;
    business = undefined;
    const disposal = prior?.dispose();
    pending = pending.then(async () => {
      await disposal;
      if (closed || requested !== generation) return;
      const current = values();
      if (paths.some(key => !current[key])) return;
      for (const key of paths) {
        if (!isAbsolute(current[key])) throw new Error(key + '必须为绝对路径');
      }
      business = ctx.plugin(Office, { ...current,
        workerPath: resolve(root, 'worker/office_worker.py'),
        nativeIdentityPath: resolve(root, 'worker/native_identity.py'),
        nativeLaunchPath: resolve(root, 'worker/native_launch.py'),
      });
      await business;
    });
    return pending;
  };
  ctx.on('loader/volatile-update', () => {
    refresh().catch(error => ctx.logger.error('办公配置应用失败，办公工具停止：', error));
  });
  ctx.effect(() => async () => {
    closed = true;
    ++generation;
    await pending;
    await business?.dispose();
  });
  await refresh();
}
