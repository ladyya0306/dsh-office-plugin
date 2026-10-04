"""Check every requested file and declared cross-file facts from actual bytes."""
from decimal import Decimal
from worker.office_worker import WorkerError
from worker.output_check import check_output, _number
from worker.structured_worker import _read


def check_group(request):
    items = request.get('artifacts')
    if not isinstance(items, list) or not 1 <= len(items) <= 16:
        raise WorkerError('group_file_limit', '整组需1至16份产物及各自模板与检查清单')
    if any(not isinstance(item, dict) for item in items):
        raise WorkerError('group_file_identity', '每份产物需完整对象')
    names = [item.get('output_path') for item in items]
    if any(not isinstance(name, str) or not name for name in names) or len(set(names)) != len(names):
        raise WorkerError('group_file_identity', '整组产物名称需明确且不重复')
    verified, values = [], {}
    for item in items:
        result = check_output({'sources': request['sources'], 'target': item['target'],
                               'artifact': item['artifact'], 'contract': item['contract']})
        result.update(output_path=item['output_path'], target_path=item['target_path'])
        verified.append(result)
        document = _read(item['artifact'], 'source')
        nodes = {n['ref']: n for n in document['nodes']}
        if document['format'] == 'docx':
            from worker.docx_output_check import docx_output_nodes
            original = _read(item['target'], 'target', target_mode=
                             'review_values' if item['contract'].get('allow_existing_values') else 'fill_blanks')
            nodes.update(docx_output_nodes(original, document, set(item['contract']['coverage'])))
        values[item['output_path']] = nodes
    links = request.get('consistency')
    if not isinstance(links, list) or not 1 <= len(links) <= 10000:
        raise WorkerError('group_consistency_required', '整组需明确主体、编号、金额和日期等跨文件核对关系')
    checks = []
    covered = {item['output_path']: {r['target_ref'] for r in item['contract']['requirements']} for item in items}
    for link in links:
        refs = link.get('refs') if isinstance(link, dict) else None
        if not isinstance(refs, list) or not 2 <= len(refs) <= 16 or link.get('kind') not in {'text', 'number'}:
            raise WorkerError('group_consistency_invalid', '每项一致性需2至16个文件位置及text或number类型')
        collected = []
        seen = set()
        for ref in refs:
            if not isinstance(ref, dict):
                raise WorkerError('group_consistency_invalid', '跨文件引用需文件名和字段位置')
            key = (ref.get('output_path'), ref.get('target_ref'))
            if key in seen or key[1] not in covered.get(key[0], set()):
                raise WorkerError('group_consistency_outside', '跨文件引用重复或不在已独立核对的完整清单内：'+str(key))
            seen.add(key)
            node = values[key[0]].get(key[1])
            if not node or node.get('formula') or node.get('value_error'):
                raise WorkerError('group_consistency_unreadable', '跨文件字段缺少可读实际值：'+str(key))
            value = node.get('value')
            if link['kind'] == 'number':
                value = _number(value)
            elif not isinstance(value, str):
                raise WorkerError('group_consistency_type', '文字一致性字段需实际文字：'+str(key))
            collected.append(value)
        if len({key[0] for key in seen}) < 2:
            raise WorkerError('group_consistency_invalid', '跨文件一致性需至少两个不同文件')
        checks.append({'label': link.get('label', ''), 'refs': refs, 'passed': all(v == collected[0] for v in collected),
                       'actual': [str(v) if isinstance(v, Decimal) else v for v in collected]})
    return {'ok': True, 'operation': 'check_group_output', 'complete': all(r['complete'] for r in verified) and all(c['passed'] for c in checks),
            'checked': sum(r['checked'] for r in verified), 'passed': sum(r['passed'] for r in verified),
            'artifacts': verified, 'consistency': checks, 'file_count': len(items)}
