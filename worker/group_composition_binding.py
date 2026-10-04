"""Current-label binder for reusable multi-template composition rules.

The saved rule stores labels and fixed text patterns only.  This module reads
the current source and templates on every call and returns current refs/guards.
"""
from copy import deepcopy
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter

from worker.evidence_read import read_evidence
from worker.structured_worker import _read
from worker.office_worker import WorkerError
from worker.group_relationship_binding import fail, text, object_only, docx_target
from worker.structure_docx import TAG


def _next(ref):
    sheet, cell = ref.rsplit('!', 1); col, row = coordinate_from_string(cell)
    return sheet + '!' + get_column_letter(column_index_from_string(col) + 1) + str(row)


def _unique(nodes, label, what):
    found = [n for n in nodes if n.get('value') == label]
    if len(found) != 1: fail('target_label_ambiguous' if found else 'target_label_missing', what, [label])
    return found[0]


def _source_field(field, source_nodes):
    node = _unique(source_nodes, field['label'], '来源标签需唯一精确匹配')
    ref = _next(node['ref']); value = next((n for n in source_nodes if n['ref'] == ref), None)
    if not value or value.get('value') is None or value.get('formula') or value.get('value_error'):
        fail('value_invalid', '来源标签右邻值缺失、含公式或读取错误', [ref])
    requirement = {'operation': field.get('operation', 'copy'), 'source_refs': ['s0::' + ref]}
    if requirement['operation'] == 'extract': requirement['text_extract'] = field['text_extract']
    for key in ('source_unit', 'target_unit', 'decimals', 'blank_markers'):
        if key in field: requirement[key] = deepcopy(field[key])
    return requirement, {'source_index': 0, 'ref': node['ref'], 'text': field['label']}


def _xlsx_target(target, label):
    node = _unique(target['nodes'], label, '目标标签需唯一精确匹配')
    ref = _next(node['ref']); value = next((n for n in target['nodes'] if n['ref'] == ref), None)
    if not value or ref not in {s['target_ref'] for s in target['slots']}: fail('target_unwritable', '目标标签右邻格不可写', [ref])
    return [(ref, value.get('value'))], {'ref': node['ref'], 'text': label}


def _docx_targets(target, label, occurrences):
    # Existing helper protects table labels.  A paragraph exactly ending in a
    # colon has a zero-width host span; retain its bookmark-safe ref.
    cells=[n for n in target['nodes'] if n.get('kind')=='cell' and n.get('text','').strip()==label]
    if cells:
        pairs=[]; guards=[]
        for node in cells:
            element=target['elements'][node['ref']]; siblings=element.getparent().findall(TAG('tc')); index=siblings.index(element)+1
            right=siblings[index] if index<len(siblings) else None
            slots=[s for s in target['slots'] if s.get('kind')=='cell' and target['elements'].get(s['target_ref']) is right]
            if len(slots)!=1: fail('target_unwritable','Word标签右邻位置不可唯一写入',[node['ref']])
            pairs.append((slots[0]['target_ref'], slots[0].get('text'))); guards.append({'ref':node['ref'],'text':node['text']})
        if len(pairs)!=occurrences: fail('target_label_ambiguous','Word目标标签物理次数不匹配',[label])
        return pairs, guards
    candidates = [s for s in target['slots'] if s.get('write_kind') == 'text_span'
                  and s.get('text') == label and s['span_start'] == s['span_end']]
    if len(candidates) != occurrences: fail('target_label_ambiguous','Word段落标签次数不匹配',[label])
    return [(s['target_ref'], '') for s in candidates], [{'ref':s['ref'],'text':label} for s in candidates]


def validate(rule, targets):
    object_only(rule, {'schema','source_sheet','source_fields','artifacts','consistency_fields'}, '组合长期规则')
    if rule.get('schema') != 'office_group_composition_rule_v1': fail('rule_invalid', '组合长期规则版本错误')
    text(rule.get('source_sheet'), '来源工作表')
    fields = rule.get('source_fields')
    if not isinstance(fields, list) or not 1 <= len(fields) <= 32: fail('rule_invalid', 'source_fields需1至32项')
    ids = set()
    for field in fields:
        object_only(field, {'id','label','operation','text_extract','source_unit','target_unit','decimals','blank_markers'}, '来源字段')
        text(field.get('id'), '来源字段id'); text(field.get('label'), '来源标签')
        if field['id'] in ids: fail('rule_invalid', '来源字段id重复')
        ids.add(field['id']); op = field.get('operation','copy')
        if op not in {'copy','extract','blank'}: fail('rule_invalid', '来源字段操作只支持copy/extract/blank')
        if op == 'extract' and (not isinstance(field.get('text_extract'), dict) or set(field['text_extract'])-{'after','before','through','cast','normalize_parentheses'}): fail('rule_invalid', 'extract需明确且只含允许text_extract字段')
        if op == 'blank' and (not isinstance(field.get('blank_markers'), list) or not 1<=len(field['blank_markers'])<=16 or any(not isinstance(x,str) or len(x)>128 for x in field['blank_markers'])): fail('rule_invalid', 'blank需1至16项明确blank_markers')
        if ('source_unit' in field) != ('target_unit' in field) or ('source_unit' in field and (field['source_unit'] not in {'元','千元','万元','亿元'} or field['target_unit'] not in {'元','千元','万元','亿元'})): fail('rule_invalid','金额单位需成对且在允许范围')
        if 'decimals' in field and (isinstance(field['decimals'],bool) or not isinstance(field['decimals'],int) or not 0<=field['decimals']<=10): fail('rule_invalid','decimals需0至10整数')
    artifacts = rule.get('artifacts')
    if not isinstance(artifacts,list) or len(artifacts) != len(targets) or not 2 <= len(artifacts) <= 15: fail('rule_invalid','产物需按当前目标顺序声明2至15项')
    if any(not isinstance(a,dict) for a in artifacts): fail('rule_invalid','产物需对象')
    if [a.get('output_index') for a in artifacts] != list(range(len(artifacts))) or any(isinstance(a.get('output_index'),bool) or not isinstance(a.get('output_index'),int) for a in artifacts): fail('rule_invalid','产物编号需连续整数')
    for a, target in zip(artifacts, targets):
        object_only(a, {'output_index','format','fields','preserve_only','allow_existing_values'}, '组合产物')
        if any(key in a and not isinstance(a[key],bool) for key in ('preserve_only','allow_existing_values')): fail('rule_invalid','保留及覆盖条件需布尔值')
        if a.get('format') != target.get('format') or a.get('format') not in {'xlsx','docx'}: fail('rule_invalid','产物格式与当前目标不符')
        if a.get('preserve_only') is True:
            if a.get('fields',[]) not in ([], None): fail('rule_invalid','整份保留不能声明字段')
            continue
        if not isinstance(a.get('fields'),list) or not a['fields']: fail('rule_invalid','非保留产物需声明字段')
        for f in a['fields']:
            object_only(f, {'target_label','field_id','text_template','components','expected_occurrences'}, '组合目标字段')
            text(f.get('target_label'),'目标标签')
            direct = isinstance(f.get('field_id'),str); composed = isinstance(f.get('text_template'),str) and isinstance(f.get('components'),list)
            if direct == composed or (direct and f['field_id'] not in ids) or (composed and (not 1<=len(f['components'])<=16 or any(not isinstance(x,str) or x not in ids for x in f['components']))): fail('rule_invalid','目标字段需恰有field_id或1至16项text_template/components')
            if 'expected_occurrences' in f and (isinstance(f['expected_occurrences'],bool) or not isinstance(f['expected_occurrences'],int) or not 1<=f['expected_occurrences']<=16): fail('rule_invalid','expected_occurrences需1至16整数')
    consistency=rule.get('consistency_fields')
    if not isinstance(consistency,list) or not consistency: fail('rule_invalid','需声明非空consistency_fields')
    for item in consistency:
        object_only(item,{'field_id','kind','roles'},'一致字段')
        if item.get('field_id') not in ids or item.get('kind') not in {'text','number'}: fail('rule_invalid','一致字段无效')
        if 'roles' in item and (not isinstance(item['roles'],list) or not 1<=len(item['roles'])<=16 or any(not isinstance(role,str) or not role for role in item['roles'])): fail('rule_invalid','一致字段roles需1至16项明确角色')


def bind_group_composition(request):
    object_only(request, {'operation','rule','sources','targets'}, '组合绑定请求')
    if request.get('operation') != 'bind_group_composition': fail('input_invalid','组合绑定操作错误')
    sources, targets = request.get('sources'), request.get('targets')
    if not isinstance(sources,list) or len(sources)!=1 or not isinstance(sources[0],dict) or sources[0].get('format')!='xlsx' or not isinstance(targets,list) or not 2<=len(targets)<=15 or any(not isinstance(item,dict) for item in targets): fail('input_invalid','需1份XLSX来源和2至15份当前目标列表')
    rule=request.get('rule'); validate(rule, targets)
    source=read_evidence(sources[0],0,{'requirements':[]})
    source_nodes=[n for n in source['nodes'] if n.get('ref','').split('!',1)[0] == rule['source_sheet']]
    if not source_nodes: fail('source_sheet_missing','当前来源缺少规则声明工作表',[rule['source_sheet']])
    fields={f['id']:f for f in rule['source_fields']}; resolved={}; guards=[]
    for fid, field in fields.items():
        resolved[fid], guard = _source_field(field, source_nodes); guards.append(guard)
    artifacts=[]; locations=[]; grouped={}; hashes=[]
    for spec, item in zip(rule['artifacts'], targets):
        target=_read(item,'target',target_mode='review_values'); hashes.append(target['sha256'])
        if not target['capabilities'].get('read'): fail('target_unwritable','目标不可读')
        if spec.get('preserve_only'):
            titles=[node for node in target['nodes'] if isinstance(node.get('value'),str) and node['value'].strip()]
            if not titles: fail('target_label_missing','整份保留目标缺少稳定标题guard')
            title=titles[0]
            artifacts.append({'output_index':spec['output_index'],'contract':{'preserve_only':True,'coverage':[],'requirements':[],'preserve_template':True},'conditions':{'source_count':1,'source_formats':['xlsx'],'target_format':spec['format'],'source_guards':guards,'target_guards':[{'ref':title['ref'],'text':title['value']}]}}); continue
        if not target['capabilities'].get('write'): fail('target_unwritable','目标不可写或受保护')
        reqs=[]; tguards=[]
        for mapping in spec['fields']:
            expected_count=mapping.get('expected_occurrences',1)
            pairs, guard = (_xlsx_target(target,mapping['target_label']) if spec['format']=='xlsx' else _docx_targets(target,mapping['target_label'],expected_count))
            if len(pairs) != expected_count: fail('target_label_ambiguous','目标标签物理次数与expected_occurrences不一致',[mapping['target_label']])
            tguards.extend(guard if isinstance(guard,list) else [guard])
            for ref, old in pairs:
                if 'field_id' in mapping:
                    req={**deepcopy(resolved[mapping['field_id']]),'target_ref':ref}
                    role=mapping['field_id']
                else:
                    components=[deepcopy(resolved[fid]) for fid in mapping['components']]
                    req={'operation':'compose','target_ref':ref,'expected_text':'' if old is None else str(old),'text_template':mapping['text_template'],'components':components}
                    role='+'.join(mapping['components'])
                reqs.append(req); grouped.setdefault(role,[]).append({'output_index':spec['output_index'],'target_ref':ref}); locations.append({'output_index':spec['output_index'],'target_ref':ref,'target_label':mapping['target_label']})
        contract={'coverage':[x['target_ref'] for x in reqs],'requirements':reqs,'preserve_template':True}
        if spec.get('allow_existing_values') is True: contract['allow_existing_values']=True
        artifacts.append({'output_index':spec['output_index'],'contract':contract,'conditions':{'source_count':1,'source_formats':['xlsx'],'target_format':spec['format'],'source_guards':guards,'target_guards':tguards}})
    consistency=[]
    for item in rule['consistency_fields']:
        roles=item.get('roles') or [item['field_id']]
        missing=[role for role in roles if role not in grouped]
        if missing: fail('consistency_missing','声明一致字段未生成位置',missing)
        refs=[ref for role in roles for ref in grouped[role]]
        if len({ref['output_index'] for ref in refs}) < 2: fail('consistency_missing','声明一致字段未跨至少2份产物',roles)
        consistency.append({'label':item['field_id'],'kind':item['kind'],'refs':refs})
    return {'ok':True,'operation':'bind_group_composition','artifacts':artifacts,'consistency':consistency,'source_sha256':[source['sha256']],'target_sha256':hashes,'field_count':sum(len(a['contract'].get('requirements',[])) for a in artifacts),'record_count':1,'locations':locations}
