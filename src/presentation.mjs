// Compact model-facing tables, preserving every candidate field and warning.
// The structured tool value remains unchanged for host-side consumers.
function table(items) {
  const shared = Object.keys(items[0] ?? {}).filter(key => items.every(item => Object.hasOwn(item, key)));
  const common_fields = items.length < 2 ? {} : Object.fromEntries(shared
    .filter(key => items.every(item => JSON.stringify(item[key]) === JSON.stringify(items[0][key])))
    .map(key => [key, items[0][key]]));
  const columns = shared.filter(key => !Object.hasOwn(common_fields, key));
  const result = { columns, rows: items.map(item => columns.map(key => item[key])) };
  if (Object.keys(common_fields).length) result.common_fields = common_fields;
  const extra = items.flatMap((item, row) => {
    const values = Object.fromEntries(Object.entries(item).filter(([key]) => !columns.includes(key) && !Object.hasOwn(common_fields, key)));
    return Object.keys(values).length ? [{ row, values }] : [];
  });
  if (extra.length) result.extra_fields = extra;
  return result;
}

export function presentOfficeResult(value) {
  if (value.operation !== 'inspect' || !Array.isArray(value.source_candidates) || !Array.isArray(value.target_slots)) return value;
  const { source_candidates, target_slots, ...details } = value;
  const previous=value.read_reuse?.previous_source_records??[];
  const previousIndices=new Set(previous.map(item=>item.candidate_index));
  const newSources=!value.read_reuse?.expanded?source_candidates.filter((_item,i)=>!previousIndices.has(i)):source_candidates;
  const reuseSummary = value.read_reuse?.used && !value.read_reuse?.expanded;
  if (reuseSummary) {
    return { ok: value.ok, operation: value.operation, read_id: value.read_id,
      source_candidates_summary: { count: source_candidates.length, content: '重复读取摘要；未在模型展示中展开来源原文' },
      target_slots_summary: { count: target_slots.length, content: '重复读取摘要；未在模型展示中展开目标原文' },
      ...details, candidates_are_unconfirmed: true,
      read_reuse: {...value.read_reuse,previous_source_records:undefined, display:'summary', expansion:value.read_reuse.readonly_expansion
        ? '传expand_reuse=true重新展示完整原文；或使用' + value.read_reuse.readonly_expansion
        : '传expand_reuse=true重新展示完整原文'},
      table_usage: '本次重复读取命中精确缓存；host结构化结果仍保留全部候选与位置，模型展示可按read_reuse中的方式展开。'};
  }
  // Evidence precedes inventories in DSH's inline preview of large results.
  return { ok: value.ok, operation: value.operation, read_id: value.read_id,
    source_candidates: table(newSources), target_slots: table(target_slots),
    ...details, candidates_are_unconfirmed: true,
    ...(previous.length&&!value.read_reuse?.expanded?{source_incremental:{returned:newSources.length,already_shown:previous.length,
      reused_refs:previous,note:'相同文件版本与完整上下文的物理位置已在所列read_id返回；本次read_id仍支持全部引用。需再看原文可传expand_reuse=true。'}}:{}),
    read_reuse:{...value.read_reuse,previous_source_records:undefined},
    table_usage: '每行先取common_fields共有属性，再把rows按columns逐列对应；extra_fields按从0开始的row补充片段位置等属性。全部字段、候选、上下文与警告保留；表头、单位格与标题也可能是候选，不等于应填字段。',
  };
}
