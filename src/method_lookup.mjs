// Lookup proposes candidates only. Execution checks current structure and facts.
export function summarizeMethods(all,query,limit=5,includeInactive=false) {
  if(!Number.isInteger(limit)||limit<1||limit>10)throw new Error('OFFICE_METHOD_QUERY_LIMIT');
  const normalize=text=>text.normalize('NFKC').toLowerCase();
  const terms=[...new Set(normalize(query??'').split(/[\s,，;；、]+/u).filter(Boolean))];
  const ranked=all.map(method=>({method,terms:terms.filter(term=>normalize(method.name+' '+method.purpose).includes(term))}))
    .filter(item=>!terms.length||item.terms.length)
    .sort((a,b)=>Number(b.method.status==='active')-Number(a.method.status==='active')||b.terms.length-a.terms.length);
  const candidates=!includeInactive&&ranked.some(c=>c.method.status==='active')?ranked.filter(c=>c.method.status==='active'):ranked;
  return {total:all.length,matched:candidates.length,methods:candidates.slice(0,limit).map(({method:m,terms:matched})=>({
    id:m.id,name:m.name,purpose:m.purpose.length>240?m.purpose.slice(0,240)+'…':m.purpose,status:m.status,parent_id:m.parent_id,revision:m.revision,
    conditions:{source_count:m.conditions.source_count,source_formats:m.conditions.source_formats,target_format:m.conditions.target_format,
      source_labels:[...new Set(m.conditions.source_guards.map(g=>g.text))],target_labels:m.conditions.target_guards.map(g=>g.text)},
    last_saved_validation:m.proofs.at(-1)??null,verified_for_current_inputs:false,...(terms.length?{matched_keywords:matched}:{})
  })),hidden_inactive:ranked.length-candidates.length,
    lookup_scope:'关键词只给候选；摘要可能截短，验证仅是历史，不表示本次适用或已检查。执行仍核对本次结构、来源和完整清单。需查看草稿或旧版可用include_inactive=true；没有启用候选时仍展示草稿。',
    ...(all.length&&!candidates.length?{next:'库非空，关键词未匹配。可不带query查看启用方法；不因此重新编写已有方法。'}:{})};
}
