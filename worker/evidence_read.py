"""Read current evidence needed by a checklist, including cells after page one."""
import copy
import json
from worker.structured_worker import _read
from worker.office_worker import WorkerError, _decode_b64, _sha256

MAX_PREFLIGHT_EVIDENCE_CACHE_RESULTS = 8

def read_evidence(item,index,contract,guards=(),cache=None):
    refs={guard['ref'] for guard in guards if guard.get('source_index')==index}
    needs_all=False
    prefix=f's{index}::'
    for rule in contract.get('requirements',[]):
        needs_all=needs_all or bool(rule.get('source_matches')) or rule.get('source_absence',{}).get('source_index')==index
        for value in [*rule.get('source_refs',[]),rule.get('source_unit_ref'),
                      *rule.get('source_label_refs',[]),
                      rule.get('target_unit_source_ref'),rule.get('subject_source_ref')]:
            if isinstance(value,str) and value.startswith(prefix) and not value[len(prefix):].startswith('$'):
                refs.add(value[len(prefix):])
    selected=sorted(refs) if refs and not needs_all and item.get('format')=='xlsx' else None
    relevant_guards=sorted((guard for guard in guards if guard.get('source_index')==index),
                           key=lambda guard: json.dumps(guard, ensure_ascii=False, sort_keys=True,
                                                        separators=(',', ':')))
    # This cache exists only for one trusted native preflight.  Decode before
    # deriving the key, so altered input bytes cannot reuse earlier evidence.
    # The key keeps read ranges and guards separate even where their selected
    # refs happen to overlap.
    cache_key=None
    if cache is not None:
        data=_decode_b64(item.get('bytes_b64'),'source')
        cache_key=(
            index, _sha256(data), item.get('format'),
            json.dumps(item.get('scope'), ensure_ascii=False, sort_keys=True, separators=(',', ':')),
            item.get('filename'), tuple(selected) if selected is not None else None, needs_all,
            tuple(json.dumps(guard, ensure_ascii=False, sort_keys=True, separators=(',', ':'))
                  for guard in relevant_guards),
        )
        existing=cache.get(cache_key)
        if existing is not None:
            return copy.deepcopy(existing)
    result=_read(item,'source',requested_refs=selected)
    next_offset=result['capabilities'].get('next_offset')
    while next_offset is not None:
        if len(result['nodes'])>=10000:
            raise WorkerError('evidence_limit','本次必要来源超过10000项，请缩小明确范围或交DSH处理')
        page=_read(item,'source',offset=next_offset,requested_refs=selected)
        result['nodes'].extend(page['nodes'])
        offset=page['capabilities'].get('next_offset')
        if offset is not None and offset<=next_offset:
            raise WorkerError('evidence_pagination','来源分页没有前进')
        next_offset=offset
    result['capabilities']={**result['capabilities'],'next_offset':None,
        'pages_combined':True,'selected_evidence_only':selected is not None,
        'all_text_nodes_read': selected is None and not item.get('scope')}
    for node in result['nodes']:
        for key in ('row_header_refs','column_header_refs'):
            if key in node: node[key]=[prefix+ref for ref in node[key]]
    result['read_scope_meta']={'ref':'$read_scope','kind':'read_scope',
        'all_text_nodes_read':result['capabilities']['all_text_nodes_read'],
        'sheets':[sheet['sheet'] for sheet in result['capabilities'].get('sheet_inventory',[])]}
    if item.get('filename'):
        result['nodes'].append({'ref':'$filename','kind':'filename_identity','value':item['filename'],'formula':False})
    if cache is not None and len(cache) < MAX_PREFLIGHT_EVIDENCE_CACHE_RESULTS:
        cache[cache_key]=copy.deepcopy(result)
    return result
