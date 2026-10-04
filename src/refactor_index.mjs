// Separate development entry. Production v22 entry remains unchanged.
import { isAbsolute, resolve } from 'node:path';
import { installNativeMethods } from './native_methods.mjs';
import { installOutputTools } from './output_tools.mjs';
export const name='office-enhancement-refactor';
export const inject=['tools','fs','shell','sandboxPolicy'];
export function apply(ctx,config={}) {
  for(const key of ['workspaceRoot','pythonPath','workerPath','nativeIdentityPath','methodDirectory'])
    if(!config[key]||!isAbsolute(config[key]))throw new Error('OFFICE_REFACTOR_CONFIG: '+key);
  if(config.workspaceMode&&!['fixed','session'].includes(config.workspaceMode))throw new Error('OFFICE_WORKSPACE_MODE_INVALID');
  if(config.learningEnabled&&(!config.nativeStateDirectory||!isAbsolute(config.nativeStateDirectory)))
    throw new Error('OFFICE_REFACTOR_CONFIG: nativeStateDirectory');
  installNativeMethods(ctx,config,resolve(config.workspaceRoot));
  if(config.outputCheckTools===true)installOutputTools(ctx,config,resolve(config.workspaceRoot));
}
