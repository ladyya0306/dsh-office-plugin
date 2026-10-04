// Expand confirmed table relationships. Business meaning remains with DSH.
import { contractSchema } from './check_contract.mjs';

const fail = message => { throw new Error('OFFICE_CHECKLIST_INVALID: ' + message); };
const object = value => value && typeof value === 'object' && !Array.isArray(value);
const columnNumber = column => {
  if (typeof column !== 'string' || !/^[A-Z]{1,3}$/.test(column)) fail('列字母需A至XFD');
  const number = [...column].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  if (number > 16384) fail('列超出XFD');
  return number;
};
const rowNumber = row => {
  if (!Number.isInteger(row) || row < 1 || row > 1048576) fail('行号需1至1048576整数');
  return row;
};
const columnName = number => {
  let name = '';
  while (number) { const digit = (number - 1) % 26; name = String.fromCharCode(65 + digit) + name; number = Math.floor((number - 1) / 26); }
  return name;
};
export function coverageTargets(coverage) {
  if (!Array.isArray(coverage) || !coverage.length) fail('先明确本次任务coverage范围');
  const refs = new Set();
  for (const original of coverage) {
    let spec=original;
    if (object(spec)) {
      knownKeys(spec,['sheet','range'],'coverage');
      if (typeof spec.sheet!=='string' || !spec.sheet || typeof spec.range!=='string' || !spec.range) fail('coverage需工作表和范围');
      spec=`${spec.sheet}!${spec.range}`;
    }
    if (typeof spec !== 'string' || !spec.trim()) fail('coverage每项需明确位置或矩形范围');
    const cut = spec.lastIndexOf('!'), sheet = spec.slice(0, cut), span = spec.slice(cut + 1);
    const match = cut > 0 && span.match(/^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})([1-9]\d*)$/);
    const add = ref => {
      if (refs.has(ref)) fail('coverage位置重复：' + ref);
      refs.add(ref); if (refs.size > 10000) fail('检查范围最多10000位置');
    };
    if (match) {
      const left = columnNumber(match[1]), right = columnNumber(match[3]);
      const top = rowNumber(Number(match[2])), bottom = rowNumber(Number(match[4]));
      if (left > right || top > bottom || (right - left + 1) * (bottom - top + 1) > 10000) fail('coverage矩形范围无效或超过10000格');
      for (let row = top; row <= bottom; row++) for (let column = left; column <= right; column++) add(`${sheet}!${columnName(column)}${row}`);
    } else {
      if (cut > 0 && span.includes(':')) fail('coverage需要具体单元格矩形，不能用整行或整列');
      add(spec);
    }
  }
  return refs;
}
const allowedRequirementKeys = new Set(Object.keys(contractSchema.properties.requirements.items.properties));
const explanationKeys = new Set(['note', 'description']);
function requirement(value, where, notes = []) {
  if (!object(value)) fail(where + '需对象');
  for (const key of Object.keys(value)) {
    if (!allowedRequirementKeys.has(key) && !explanationKeys.has(key)) fail(where + '未知字段：' + key);
    if (explanationKeys.has(key) && typeof value[key] !== 'string') fail(where + '.' + key + '需说明文字');
  }
  const result = structuredClone(value);
  for (const key of explanationKeys) {
    if (Object.hasOwn(result, key)) {
      notes.push({ where, text: result[key] });
      delete result[key];
    }
  }
  return result;
}
function knownKeys(value, keys, where) {
  if (!object(value)) fail(where + '需对象');
  for (const key of Object.keys(value)) if (!keys.includes(key)) fail(where + '未知字段：' + key);
}
export function expandContract(input) {
  if (!object(input)) fail('清单需对象');
  knownKeys(input,Object.keys(contractSchema.properties),'contract');
  const contract = structuredClone(input);
  if (contract.preserve_only === true) {
    if (!Array.isArray(contract.coverage) || contract.coverage.length ||
        !Array.isArray(contract.requirements) || contract.requirements.length ||
        contract.preserve_template !== true || contract.tables !== undefined ||
        (contract.table_ranges !== undefined && (!object(contract.table_ranges) || Object.keys(contract.table_ranges).length)) ||
        contract.allow_existing_values === true)
      fail('preserve_only需空coverage/requirements且preserve_template=true，只核对整份原样保留');
    return contract;
  }
  const expected = coverageTargets(contract.coverage);
  const notes = [];
  const requirements = contract.requirements === undefined ? [] : contract.requirements;
  if (!Array.isArray(requirements)) fail('requirements需数组');
  const expanded = requirements.map((item, index) => requirement(item, `requirements[${index}]`, notes));
  if (contract.tables !== undefined) {
    if (!Array.isArray(contract.tables) || !contract.tables.length) fail('tables需非空数组');
    for (const table of contract.tables) {
      knownKeys(table, ['sheet', 'fields', 'records'], 'tables项');
      if (typeof table.sheet !== 'string' || !table.sheet || table.sheet.includes('!')) fail('明确目标工作表名称');
      if (!Array.isArray(table.fields) || !table.fields.length || !Array.isArray(table.records) || !table.records.length) fail('字段和记录需非空数组');
      if (table.fields.length * table.records.length + expanded.length > 10000) fail('展开后最多10000位置');
      const columns = new Set();
      for (const field of table.fields) {
        knownKeys(field, ['column', 'source_column', 'source_label_column', 'requirement'], 'field');
        columnNumber(field.column);
        if (columns.has(field.column)) fail('字段列重复：' + field.column);
        columns.add(field.column);
        if (field.source_column !== undefined) columnNumber(field.source_column);
        if (field.source_label_column !== undefined) columnNumber(field.source_label_column);
      }
      for (const record of table.records) {
        knownKeys(record, ['target_row', 'source_index', 'source_sheet', 'source_row', 'bindings'], 'record');
        rowNumber(record.target_row);
        if (record.bindings !== undefined && !object(record.bindings)) fail('bindings需按目标列声明对象');
        for (const column of Object.keys(record.bindings ?? {})) if (!columns.has(column)) fail('未声明字段列：' + column);
        for (const field of table.fields) {
          const target = `${table.sheet}!${field.column}${record.target_row}`;
          const base = field.requirement === undefined ? {} : requirement(field.requirement, target + '字段要求', notes);
          const binding = record.bindings?.[field.column];
          const item = { ...base, ...(binding === undefined ? {} : requirement(binding, target + '来源依据', notes)) };
          if (item.target_ref !== undefined) fail('表格展开的target_ref由行列生成，不另外填写：' + target);
          const sourceRef = column => {
            if (!Number.isInteger(record.source_index) || record.source_index < 0 || record.source_index > 7 ||
                typeof record.source_sheet !== 'string' || !record.source_sheet || record.source_sheet.includes('!')) fail(target + '需明确source_index和source_sheet');
            rowNumber(record.source_row);
            return `s${record.source_index}::${record.source_sheet}!${column}${record.source_row}`;
          };
          if (field.source_column !== undefined && !item.source_refs && !item.source_matches && !item.source_absence) item.source_refs = [sourceRef(field.source_column)];
          if (field.source_label_column !== undefined && !item.source_label_refs) item.source_label_refs = [sourceRef(field.source_label_column)];
          if (!item.source_refs?.length && !item.source_matches?.length && !item.source_absence) fail(target + '缺少已确认来源；用binding声明真实来源或缺失依据');
          expanded.push({ target_ref: target, ...item });
        }
      }
    }
  }
  const actual = new Set();
  for (const item of expanded) {
    if (typeof item.target_ref !== 'string' || !item.target_ref) fail('每项需target_ref');
    if (actual.has(item.target_ref)) fail('要求位置重复：' + item.target_ref);
    actual.add(item.target_ref);
  }
  const missing = [...expected].filter(ref => !actual.has(ref));
  const extra = [...actual].filter(ref => !expected.has(ref));
  if (missing.length || extra.length) fail(`要求覆盖${actual.size}/${expected.size}；缺${missing.length}项${missing.length ? '：' + missing.slice(0, 12).join('、') : ''}；范围外${extra.length}项${extra.length ? '：' + extra.slice(0, 12).join('、') : ''}。补来源依据或明确本次任务范围，保留已正确文件。`);
  delete contract.tables;
  return notes.length ? { ...contract, requirements: expanded, notes } : { ...contract, requirements: expanded };
}
