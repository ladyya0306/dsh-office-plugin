import { mkdir, readFile, writeFile, rename, open, unlink, readdir, lstat, realpath } from 'node:fs/promises';
import { resolve, isAbsolute } from 'node:path';
import { randomUUID } from 'node:crypto';
import { sha256 } from './runtime.mjs';
import { expandContract } from './contract_builder.mjs';

const identifier = id => {
  if (typeof id !== 'string' || !/^[a-z0-9][a-z0-9-]{0,79}$/.test(id)) throw new Error('OFFICE_METHOD_ID_INVALID');
  return id;
};
async function ordinary(path) {
  const stat = await lstat(path);
  if (stat.isSymbolicLink() || !stat.isDirectory()) throw new Error('OFFICE_METHOD_DIRECTORY_LINK');
  if ((await realpath(path)).toLowerCase() !== resolve(path).toLowerCase()) throw new Error('OFFICE_METHOD_DIRECTORY_LINK');
  return stat;
}
export class MethodStore {
  constructor(directory) {
    if (!isAbsolute(directory)) throw new Error('OFFICE_METHOD_DIRECTORY_ABSOLUTE');
    this.directory = resolve(directory);
  }
  async ready() { await mkdir(this.directory, {recursive:true}); await ordinary(this.directory); }
  folder(id) { return resolve(this.directory, identifier(id)); }
  async create({code, name, purpose, conditions, contract, parent_id = null, execution = null, dependencies = []}) {
    const grouped=contract?.schema==='office_group_contract_v1';
    const ruleGroup=execution?.kind==='native_python'&&['office_group_relationship_rule_v1','office_group_composition_rule_v1'].includes(execution.relationship_rule?.schema);
    if(ruleGroup&&!grouped)throw new Error('OFFICE_METHOD_GROUP_METADATA');
    if(grouped){
      if(!Array.isArray(contract.artifacts)||!contract.artifacts.length||contract.artifacts.length>16||
          !Array.isArray(contract.consistency)||!contract.consistency.length||conditions?.schema!=='office_group_conditions_v1'||
          conditions.artifacts?.length!==contract.artifacts.length)throw new Error('OFFICE_METHOD_GROUP_METADATA');
      contract={...contract,artifacts:contract.artifacts.map((a,i)=>{if(a.output_index!==i)throw new Error('OFFICE_METHOD_GROUP_METADATA');return {...a,contract:expandContract(a.contract)};})};
    }else contract = expandContract(contract);
    await this.ready();
    if (![name,purpose].every(v=>typeof v === 'string' && v.trim() && v.length <= 1000)) throw new Error('OFFICE_METHOD_DESCRIPTION_REQUIRED');
    if (!code.trim() || Buffer.byteLength(code)>65536) throw new Error('OFFICE_METHOD_CODE_LIMIT');
    for(const item of grouped?conditions.artifacts:[conditions]){
    if (!item || !Array.isArray(item.source_guards) || !item.source_guards.length ||
        !Array.isArray(item.target_guards) || !item.target_guards.length) throw new Error('OFFICE_METHOD_GUARDS_REQUIRED');
    const count=item.source_count;
    if(!Number.isInteger(count)||count<1||count>8||item.source_formats?.length!==count||
       item.source_guards.some(g=>!Number.isInteger(g.source_index)||g.source_index<0||g.source_index>=count)||
       new Set(item.source_guards.map(g=>g.source_index)).size!==count)
      throw new Error('OFFICE_METHOD_GUARDS_EACH_SOURCE: 每份来源需明确结构条件，同一来源可检查多个标签');
    }
    if ((grouped?contract.artifacts.map(a=>a.contract):[contract]).some(c=>
      c?.preserve_only===true
        ? !(Array.isArray(c.coverage)&&c.coverage.length===0&&Array.isArray(c.requirements)&&c.requirements.length===0&&c.preserve_template===true)
        : !c?.coverage?.length || !c?.requirements?.length || c.preserve_template === false))
      throw new Error('OFFICE_METHOD_CHECKLIST_REQUIRED: 保存独立完整范围和来源要求，不允许取消格式检查');
    if (Buffer.byteLength(JSON.stringify({conditions,contract}))>512*1024) throw new Error('OFFICE_METHOD_METADATA_LIMIT');
    const parent=parent_id?(await this.read(parent_id)).meta:null;
    if (execution) {
      if (execution.kind !== 'native_python' || !isAbsolute(execution.interpreter) ||
          !Array.isArray(execution.bundle) || execution.bundle.length !== dependencies.length + 1 ||
          execution.bundle[0].sha256 !== sha256(code)) throw new Error('OFFICE_NATIVE_MANIFEST_INVALID');
      if(execution.bundle.some(f=>typeof f.path!=='string'||isAbsolute(f.path)||f.path.includes(':')||
          !f.path.endsWith('.py')||f.path.split(/[\\/]/).some(p=>!p||p==='..'))||
          new Set(execution.bundle.map(f=>f.path)).size!==execution.bundle.length||execution.entry!==execution.bundle[0].path)
        throw new Error('OFFICE_NATIVE_MANIFEST_PATH');
      for (const [i, dep] of dependencies.entries()) {
        if (dep.sha256 !== sha256(dep.bytes) || dep.path !== execution.bundle[i+1].path || dep.sha256 !== execution.bundle[i+1].sha256)
          throw new Error('OFFICE_NATIVE_DEPENDENCY_CHANGED');
      }
    } else if (dependencies.length) throw new Error('OFFICE_NATIVE_MANIFEST_REQUIRED');
    const id = randomUUID(), folder = this.folder(id),temp=resolve(this.directory,'.draft-'+id);
    await mkdir(temp);
    // Validation above uses the actual independently checked batch. Store only
    // reusable shape here; locations, original target text and selectors stay
    // in its owner-bound run, and are regenerated for every reuse.
    if(ruleGroup){
      conditions={schema:'office_group_rule_conditions_v1',source_formats:['xlsx'],target_formats:conditions.artifacts.map(a=>a.target_format)};
      contract={schema:'office_group_rule_contract_v1',artifact_count:contract.artifacts.length,preserve_template:true};
    }
    const meta = {id,version:parent?parent.version+1:1,family_id:parent?.family_id??parent?.id??id,parent_id,name,purpose,conditions,contract,
      code_sha256:sha256(code),contract_sha256:sha256(JSON.stringify(contract)),status:'draft',revision:1,
      conditions_sha256:sha256(JSON.stringify(conditions)),
      created_at:new Date().toISOString(),proofs:[],failure:null};
    if (execution) { meta.execution=execution; meta.execution_sha256=sha256(JSON.stringify(execution)); }
    await writeFile(resolve(temp,'method.py'),code,{flag:'wx',encoding:'utf8'});
    for (const [i,dep] of dependencies.entries()) await writeFile(resolve(temp,`dependency-${i}.py`),dep.bytes,{flag:'wx'});
    await writeFile(resolve(temp,'method.json'),JSON.stringify(meta,null,2),{flag:'wx',encoding:'utf8'});
    await rename(temp,folder);
    return meta;
  }
  async read(id) {
    await this.ready(); const folder = this.folder(id); await ordinary(folder);
    for (const name of ['method.json','method.py']) if ((await lstat(resolve(folder,name))).isSymbolicLink()) throw new Error('OFFICE_METHOD_FILE_LINK');
    const text = await readFile(resolve(folder,'method.json'),'utf8');
    if (Buffer.byteLength(text)>1024*1024) throw new Error('OFFICE_METHOD_METADATA_LIMIT');
    const meta = JSON.parse(text), code = await readFile(resolve(folder,'method.py'),'utf8');
    if (meta.id!==id || !['draft','active','disabled'].includes(meta.status) || sha256(code)!==meta.code_sha256 ||
        sha256(JSON.stringify(meta.contract))!==meta.contract_sha256 || sha256(JSON.stringify(meta.conditions))!==meta.conditions_sha256) throw new Error('OFFICE_METHOD_VERSION_CHANGED');
    const family=meta.family_id??meta.id;
    let pointer;
    try{pointer=JSON.parse(await readFile(resolve(this.directory,'family-'+identifier(family)+'.json'),'utf8'));}
    catch(e){if(e.code!=='ENOENT')throw e;}
    if(pointer&&meta.status==='active'&&pointer.active_id!==id)meta.status='disabled';
    if(meta.family_id&&!pointer&&meta.status==='active')meta.status='draft';
    const dependencies=[];
    if (meta.execution) {
      if (meta.execution.kind!=='native_python' || sha256(JSON.stringify(meta.execution))!==meta.execution_sha256 ||
          meta.execution.bundle[0].sha256!==meta.code_sha256) throw new Error('OFFICE_METHOD_VERSION_CHANGED');
      for (const [i,item] of meta.execution.bundle.slice(1).entries()) {
        const file=resolve(folder,`dependency-${i}.py`);
        if((await lstat(file)).isSymbolicLink()) throw new Error('OFFICE_METHOD_FILE_LINK');
        const bytes=await readFile(file);
        if(sha256(bytes)!==item.sha256) throw new Error('OFFICE_NATIVE_DEPENDENCY_CHANGED');
        dependencies.push({...item,bytes});
      }
    }
    return {meta,code,dependencies};
  }
  async activate(meta,summary) {
    await this.ready(); await ordinary(this.folder(meta.id));
    const family=identifier(meta.family_id??meta.id),lockPath=resolve(this.directory,'family-'+family+'.lock');
    const lock=await open(lockPath,'wx');
    try {
      let pointer;
      try{pointer=JSON.parse(await readFile(resolve(this.directory,'family-'+family+'.json'),'utf8'));}
      catch(e){if(e.code!=='ENOENT')throw e;}
      if(pointer&&pointer.active_id!==meta.id&&pointer.active_id!==meta.parent_id)
        throw new Error('OFFICE_METHOD_ACTIVE_VERSION_CHANGED: 当前启用版本已改变，基于当前版本重新修订');
      const enabled=await this.update(meta.id,meta.revision,m=>({...m,status:'active',validation_summary:summary,enabled_at:new Date().toISOString()}));
      const temp=resolve(this.directory,randomUUID()+'.tmp');
      await writeFile(temp,JSON.stringify({family_id:family,active_id:meta.id}),{flag:'wx'});
      // One atomic pointer selects the current version, including sibling fixes.
      await rename(temp,resolve(this.directory,'family-'+family+'.json'));
      return enabled;
    } finally {await lock.close();await unlink(lockPath);}
  }
  async update(id, revision, change) {
    await this.ready();
    const folder = this.folder(id), lockPath=resolve(folder,'update.lock');
    const before=await ordinary(folder);
    const lock=await open(lockPath,'wx');
    try {
      const after=await ordinary(folder);
      if (before.dev!==after.dev || before.ino!==after.ino) throw new Error('OFFICE_METHOD_DIRECTORY_CHANGED');
      const {meta}=await this.read(id);
      if (meta.revision!==revision) throw new Error('OFFICE_METHOD_CONCURRENT_UPDATE');
      const next = change(structuredClone(meta)); next.revision=revision+1;
      const temp=resolve(folder,randomUUID()+'.tmp');
      await writeFile(temp,JSON.stringify(next,null,2),{flag:'wx',encoding:'utf8'});
      await rename(temp,resolve(folder,'method.json')); return next;
    } finally { await lock.close(); await unlink(lockPath); }
  }
  async list() {
    await this.ready(); const items=[];
    for (const entry of await readdir(this.directory,{withFileTypes:true})) {
      if (!entry.isDirectory() || !/^[a-z0-9][a-z0-9-]{0,79}$/.test(entry.name)) continue;
      const {meta}=await this.read(entry.name);
      items.push({id:meta.id,name:meta.name,purpose:meta.purpose,status:meta.status,
        execution_kind:meta.execution?.kind??'office_callbacks',
        parent_id:meta.parent_id,conditions:meta.conditions,proofs:meta.proofs.map(p=>({checked:p.checked,passed:p.passed,tested_at:p.tested_at})),revision:meta.revision});
      if (items.length>256) throw new Error('OFFICE_METHOD_LIST_LIMIT: 请先按目录整理超过256项的方法');
    }
    return items;
  }
}
