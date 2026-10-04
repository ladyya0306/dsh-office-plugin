import { readFileSync } from 'node:fs';
const builtin = JSON.parse(readFileSync(new URL('../library/builtin/methods.json', import.meta.url), 'utf8'));
export function selectMethods(inspection, enabled = true, disabledIds = []) {
  if (!enabled) return [];
  const nodes = [...inspection.source_candidates, ...inspection.target_slots];
  const labels = nodes.map(n => n.label ?? n.column_headers?.at(-1) ?? n.row_headers?.at(-1)).filter(Boolean);
  const repeat = labels.some((v,i) => labels.indexOf(v) !== i);
  const scoped = nodes.some(n => n.parent_context?.some(Boolean) || n.row_headers?.length);
  const matrix = nodes.some(n => n.row_headers?.length && n.column_headers?.length);
  const text = nodes.map(n => [n.label,n.text,...(n.row_headers??[]),...(n.column_headers??[])].filter(Boolean).join(' ')).join('\n');
  const sources = inspection.capabilities?.sources ?? [];
  const matches = {
    'subject-purpose': repeat && scoped,
    'row-column-location': matrix,
    'contact-purpose': /电话|手机|联系方式|疑似关系/.test(text),
    'sparse-source-completeness': sources.some(s => s.next_offset != null || s.read_ignored_parts?.length ||
      s.sheet_inventory?.some(i => i.declared_dimension === 'A1' && i.nonempty_or_formula_cells > 1)),
    'existing-value-review': inspection.capabilities?.target?.target_mode === 'review_values' && inspection.target_slots.some(n=>n.requires_expected_text),
    'period-and-derived-value': /20\d{2}年|20\d{2}[-/]\d{1,2}|增长率|变化率|财务比率/.test(text),
    'complete-output-check': inspection.capabilities?.target?.read === true &&
      inspection.source_candidates.length > 0 && inspection.target_slots.length > 1,
  };
  return builtin.filter(m => ['trial','active'].includes(m.status) && !disabledIds.includes(m.id))
    .filter(m => matches[m.id] === true).sort((a,b)=>(b.priority??0)-(a.priority??0)).slice(0,3)
    .map(m => ({ ...m, advice_only: true, current_values_required: true }));
}
export const methodDetails = () => structuredClone(builtin);
export function userMethodDetails() {
  // Fixed plugin-owned path, never a path supplied by the model. Imported drafts
  // are visible across workspaces but are not selected or executed automatically.
  const data=JSON.parse(readFileSync(new URL('../library/user/methods.json',import.meta.url),'utf8'));
  if (!Array.isArray(data) || data.length>32) throw new Error('OFFICE_METHOD_LIBRARY_INVALID');
  return structuredClone(data);
}

