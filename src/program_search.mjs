const normalize=value=>value.normalize('NFC').toLowerCase().trim();
const order=rows=>[...rows].sort((a,b)=>(a.status==='active'?0:1)-(b.status==='active'?0:1)||String(a.id).localeCompare(String(b.id)));

export function searchPrograms(rows,query) {
  if(!Array.isArray(rows))throw new Error('OFFICE_PROGRAM_SEARCH_ROWS');
  const native=rows.filter(row=>row?.execution_kind==='native_python');
  if(query===undefined||query===null||query==='')return {methods:order(native).slice(0,10),search_mode:'inventory'};
  if(typeof query!=='string')throw new Error('OFFICE_PROGRAM_SEARCH_QUERY');
  const text=normalize(query);
  if(text.length>200)throw new Error('OFFICE_PROGRAM_SEARCH_QUERY_LIMIT');
  if(!text)return {methods:order(native).slice(0,10),search_mode:'inventory'};
  const terms=[...new Set(text.split(/[\s\p{P}\p{S}]+/u).filter(Boolean))];
  if(!terms.length||terms.length>32)throw new Error('OFFICE_PROGRAM_SEARCH_TERMS');
  const haystack=row=>normalize(`${row.name??''} ${row.purpose??''}`);
  const whole=native.filter(row=>haystack(row).includes(text));
  if(whole.length)return {methods:order(whole).slice(0,10),search_mode:'substring'};
  const matched=native.map(row=>({...row,_matched:terms.filter(term=>haystack(row).includes(term)).length})).filter(row=>row._matched);
  if(matched.length)return {methods:matched.sort((a,b)=>b._matched-a._matched||(a.status==='active'?0:1)-(b.status==='active'?0:1)||String(a.id).localeCompare(String(b.id))).slice(0,10).map(({_matched,...row})=>row),search_mode:'keyword'};
  return {methods:order(native.filter(row=>row.status==='active')).slice(0,10),search_mode:'active_inventory_fallback',message:'没有关键词匹配，列出已启用程序及适用范围，运行前按当前文件检查'};
}
