import { mkdir,lstat,realpath,readFile,writeFile,rename } from 'node:fs/promises';
import { resolve,isAbsolute } from 'node:path';
import { randomUUID } from 'node:crypto';
import { outside,sha256 } from './runtime.mjs';

// Only host configuration selects this directory. It contains hashes, never
// source values, user correction text or input/output file contents.
async function location(config,root,kind,key) {
  const directory=config.nativeStateDirectory;
  if(!directory||!isAbsolute(directory)||!outside(root,resolve(directory)))
    throw new Error('OFFICE_NATIVE_HOST_STATE: 宿主收据目录需配置在任务工作区之外');
  await mkdir(directory,{recursive:true});
  if((await lstat(directory)).isSymbolicLink()||resolve(await realpath(directory)).toLowerCase()!==resolve(directory).toLowerCase())
    throw new Error('OFFICE_NATIVE_HOST_STATE_LINK');
  const path=resolve(directory,sha256(JSON.stringify([resolve(root),kind,key]))+'.json');
  try{if((await lstat(path)).isSymbolicLink())throw new Error('OFFICE_NATIVE_HOST_STATE_LINK');}
  catch(e){if(e.code!=='ENOENT')throw e;}
  return path;
}
export async function hostDigest(config,root,kind,key) {
  const path=await location(config,root,kind,key);
  try{return JSON.parse(await readFile(path,'utf8')).sha256;}
  catch(e){if(e.code==='ENOENT')return null;throw e;}
}
export async function writeHostDigest(config,root,kind,key,digest) {
  const path=await location(config,root,kind,key),temp=path+'.'+randomUUID()+'.tmp';
  await writeFile(temp,JSON.stringify({sha256:digest}),{flag:'wx'});await rename(temp,path);
}
