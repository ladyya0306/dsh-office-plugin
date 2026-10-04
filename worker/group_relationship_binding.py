"""Bind one current XLSX source to current XLSX/DOCX group templates."""
from worker.office_worker import WorkerError
from worker.evidence_read import read_evidence
from worker.structured_worker import _read
from worker.relationship_binding import column, header, record_keys


def fail(code, message, refs=()):
    raise WorkerError('group_relationship_'+code, message+('：'+'、'.join(refs) if refs else ''),
                      {'references': list(refs)})


def text(value, where):
    if not isinstance(value, str) or not value.strip() or len(value) > 80:
        fail('rule_invalid', where+'需为1至80字的明确文字')
    return value


def object_only(value, keys, where):
    if not isinstance(value, dict) or set(value)-set(keys): fail('rule_invalid', where+'含未知字段')


def field_rule(value):
    object_only(value, {'id','source_labels','operation','source_unit_label','target_unit','decimals','blank_markers'}, '字段规则')
    text(value.get('id'), '字段id')
    labels=value.get('source_labels')
    if not isinstance(labels, list) or not 1 <= len(labels) <= 8 or any(not isinstance(x,str) for x in labels) or len(set(labels)) != len(labels): fail('rule_invalid','来源字段需为1至8个不重复标签')
    for item in labels: text(item,'来源字段')
    op=value.get('operation','copy')
    if not isinstance(op,str) or op not in {'copy','difference','ratio','blank'} or len(labels)!=(2 if op in {'difference','ratio'} else 1):
        fail('rule_invalid','字段操作或来源字段数量不支持')
    if bool(value.get('source_unit_label')) != bool(value.get('target_unit')): fail('rule_invalid','单位字段与目标单位需成对填写')
    if value.get('source_unit_label'):
        text(value['source_unit_label'],'来源单位字段')
        if not isinstance(value['target_unit'],str) or value['target_unit'] not in {'元','千元','万元','亿元'} or op=='blank': fail('rule_invalid','金额目标单位无效')
    if 'decimals' in value and (isinstance(value['decimals'],bool) or not isinstance(value['decimals'],int) or not 0 <= value['decimals'] <= 10): fail('rule_invalid','小数位需0至10')
    if op=='blank' and (not isinstance(value.get('blank_markers'),list) or not 1<=len(value['blank_markers'])<=16 or any(not isinstance(x,str) or len(x)>128 for x in value['blank_markers'])): fail('rule_invalid','空值规则需1至16项明确文字标记，每项最多128字')


def validate(rule, selections, target_count):
    object_only(rule, {'schema','source_sheet','keys','fields','artifacts','consistency_fields'}, '长期规则')
    if rule.get('schema')!='office_group_relationship_rule_v1': fail('rule_invalid','长期规则版本错误')
    text(rule.get('source_sheet'),'来源工作表')
    if '!' in rule['source_sheet']: fail('rule_invalid','来源工作表名不能含!')
    keys=rule.get('keys'); fields=rule.get('fields'); artifacts=rule.get('artifacts')
    if not isinstance(keys,list) or not 1 <= len(keys) <= 8 or any(not isinstance(k,str) for k in keys) or len(set(keys)) != len(keys): fail('rule_invalid','来源匹配键需为1至8个不重复标签')
    for key in keys: text(key,'来源匹配键')
    if not isinstance(fields,list) or not 1 <= len(fields) <= 32 or any(not isinstance(f,dict) for f in fields): fail('rule_invalid','字段需为1至32项')
    for item in fields: field_rule(item)
    ids=[item['id'] for item in fields]
    by_id={item['id']:item for item in fields}
    if len(set(ids))!=len(ids): fail('rule_invalid','字段id重复')
    if not isinstance(artifacts,list) or not 2 <= len(artifacts) <= 15 or any(not isinstance(a,dict) for a in artifacts) or len(artifacts)!=target_count: fail('rule_invalid','产物需按输入顺序连续声明2至15项')
    if [a.get('output_index') for a in artifacts] != list(range(len(artifacts))) or any(isinstance(a.get('output_index'),bool) for a in artifacts): fail('rule_invalid','产物编号需连续整数')
    for artifact in artifacts:
        object_only(artifact, {'output_index','format','target_sheet','keys','fields','table_name','record_selector'}, '产物')
        if artifact.get('format')=='xlsx':
            text(artifact.get('target_sheet'),'目标工作表')
            if '!' in artifact['target_sheet']: fail('rule_invalid','目标工作表名不能含!')
            if 'table_name' in artifact: fail('rule_invalid','本期多文件长期规则尚不支持table_name')
            if not isinstance(artifact.get('keys'),list) or not artifact['keys']: fail('rule_invalid','XLSX产物需声明匹配键')
            for key in artifact['keys']:
                object_only(key, {'source_label','target_label'}, 'XLSX匹配键'); text(key.get('source_label'),'来源匹配键'); text(key.get('target_label'),'目标匹配键')
                if key['source_label'] not in keys: fail('rule_invalid','XLSX匹配键不属于来源匹配键')
            if [k['source_label'] for k in artifact['keys']] != keys or len({k['target_label'] for k in artifact['keys']}) != len(keys): fail('rule_invalid','XLSX匹配键必须按来源键顺序完整对应且目标标签不重复')
        elif artifact.get('format')=='docx': text(artifact.get('record_selector'),'Word记录选择')
        else: fail('rule_invalid','产物格式只能xlsx或docx')
        if not isinstance(artifact.get('fields'),list) or not 1 <= len(artifact['fields']) <= 32 or any(not isinstance(i,dict) for i in artifact['fields']): fail('rule_invalid','产物需声明1至32个字段')
        seen_ids=[]; seen_labels=[]
        for item in artifact['fields']:
            object_only(item, {'field_id','target_label'}, '产物字段'); text(item.get('field_id'),'字段id'); text(item.get('target_label'),'目标标签')
            if item['field_id'] not in ids: fail('rule_invalid','产物引用未知字段id')
            seen_ids.append(item['field_id']); seen_labels.append(item['target_label'])
        if len(set(seen_ids)) != len(seen_ids) or len(set(seen_labels)) != len(seen_labels): fail('rule_invalid','产物字段id或目标标签重复')
        if artifact.get('format')=='xlsx' and not {k['target_label'] for k in artifact['keys']} <= set(seen_labels): fail('rule_invalid','XLSX匹配键也必须作为输出字段核对')
        for key in keys:
            identity=[by_id[item['field_id']] for item in artifact['fields']
                if by_id[item['field_id']]['source_labels']==[key] and by_id[item['field_id']].get('operation','copy')=='copy'
                and not by_id[item['field_id']].get('source_unit_label')]
            if not identity: fail('rule_invalid','每份产物都需直接核对主体和合同匹配键：'+key)
        if artifact['format']=='xlsx':
            for key in artifact['keys']:
                output=next(f for f in artifact['fields'] if f['target_label']==key['target_label'])
                field=by_id[output['field_id']]
                if field['source_labels']!=[key['source_label']] or field.get('operation','copy')!='copy' or field.get('source_unit_label'):
                    fail('rule_invalid','目标匹配键需复制其对应来源匹配键：'+key['target_label'])
        if artifact.get('format')=='docx' and artifact['record_selector'] not in selections: fail('rule_invalid','Word记录选择不存在')
    if not isinstance(rule.get('consistency_fields'),list) or not rule['consistency_fields']: fail('rule_invalid','需声明跨文件一致字段')
    seen_consistency=[]
    for item in rule['consistency_fields']:
        object_only(item, {'field_id','kind'}, '一致字段');
        if item.get('field_id') not in ids or not isinstance(item.get('kind'),str) or item.get('kind') not in {'text','number'}: fail('rule_invalid','一致字段无效')
        seen_consistency.append(item['field_id'])
    if len(set(seen_consistency)) != len(seen_consistency): fail('rule_invalid','一致字段重复')


def selected_records(source, rule, selections):
    labels=set(rule['keys'])
    for field in rule['fields']:
        labels.update(field['source_labels'])
        if field.get('source_unit_label'): labels.add(field['source_unit_label'])
    row, headers=header(source['nodes'], rule['source_sheet'], labels)
    records,cells=record_keys(source['nodes'],rule['source_sheet'],row,[column(headers[k]['ref']) for k in rule['keys']])
    if not 1 <= len(selections) <= 128: fail('selection_invalid','本批记录选择需1至128项')
    found={}; used=set()
    for name, values in selections.items():
        if not isinstance(name,str) or not 1 <= len(name) <= 80 or not isinstance(values,list) or len(values)!=len(rule['keys']) or any(isinstance(v,bool) or not isinstance(v,(str,int,float)) or (isinstance(v,str) and not v) or (isinstance(v,float) and (v != v or v in (float('inf'),float('-inf')))) for v in values): fail('selection_invalid','本批记录选择需与匹配键等长的明确值数组')
        matches=[(key,record) for key,record in records.items() if len(key)==len(values) and all(type(k[1]) is type(v) and k[1]==v for k,v in zip(key,values))]
        if len(matches)!=1: fail('selection_missing' if not matches else 'selection_ambiguous','本批记录选择未唯一匹配来源', [ref for _, record in matches for ref in record['refs']])
        key, record=matches[0]
        if key in used: fail('selection_duplicate','同一来源记录不能用不同选择名重复选择',record['refs'])
        used.add(key); found[name]={**record,'key':key}
    return row,headers,cells,found


def source_requirement(field, source_row, headers, cells, sheet, target_ref):
    refs=['s0::'+sheet+'!'+column(headers[name]['ref'])+str(source_row) for name in field['source_labels']]
    for ref in refs:
        node=cells.get(ref[4:])
        if not node or node.get('value') is None or node.get('formula') or node.get('value_error'): fail('value_invalid','来源字段缺少可用值、含公式或读取错误',[ref])
    result={'target_ref':target_ref,'source_refs':refs,'operation':field.get('operation','copy')}
    if result['operation']=='blank': result['blank_markers']=field['blank_markers']
    if 'decimals' in field: result['decimals']=field['decimals']
    if field.get('source_unit_label'):
        unit_ref=sheet+'!'+column(headers[field['source_unit_label']]['ref'])+str(source_row); unit=cells.get(unit_ref,{})
        if unit.get('formula') or unit.get('value_error') or unit.get('value') not in {'元','千元','万元','亿元'}: fail('unit_invalid','来源单位不明确',[unit_ref])
        result.update(source_unit=unit['value'],target_unit=field['target_unit'],source_unit_ref='s0::'+unit_ref)
    return result


def docx_target(target, label):
    # Parsed label/value fragments share their physical paragraph's text. Count
    # physical objects only, including labels whose neighbours are not writable.
    labels=[node for node in target['nodes'] if (node.get('kind')=='cell' and node.get('text','').strip()==label) or (node.get('kind')=='paragraph' and '|v:' not in node['ref'] and node.get('text','').startswith(label) and len(node.get('text',''))>len(label) and node['text'][len(label)] in '：:')]
    if len(labels)!=1: fail('target_label_missing' if not labels else 'target_label_ambiguous','Word目标标签需唯一完整物理标签',[label])
    node=labels[0]
    if node['kind']=='cell':
        element=target['elements'][node['ref']]
        siblings=element.getparent().findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
        index=siblings.index(element)+1
        right=siblings[index] if index<len(siblings) else None
        slots=[slot for slot in target['slots'] if slot.get('kind')=='cell' and target['elements'].get(slot['target_ref']) is right]
    else: slots=[slot for slot in target['slots'] if slot.get('target_ref','').startswith(node['ref']+'|s:') and slot.get('text','').startswith(label)]
    if len(slots)!=1: fail('target_unwritable' if not slots else 'target_label_ambiguous','Word标签右邻位置或单段落占位不可唯一写入',[node['ref']])
    slot=slots[0]
    return slot['target_ref'], slot['text'] if slot.get('requires_expected_text') else None, {'ref':node['ref'],'text':node['value']}


def bind_group_relationship(request):
    object_only(request, {'operation','rule','sources','targets','record_selection'}, '绑定请求')
    if 'operation' in request and request['operation'] != 'bind_group_relationship': fail('input_invalid','绑定请求操作错误')
    sources=request.get('sources'); targets=request.get('targets'); selections=request.get('record_selection')
    if not isinstance(sources,list) or len(sources)!=1 or not isinstance(sources[0],dict) or sources[0].get('format')!='xlsx' or not isinstance(targets,list) or not 2 <= len(targets) <= 15 or not isinstance(selections,dict): fail('input_invalid','需1份XLSX来源、2至15份目标和本批记录选择')
    if any(not isinstance(t,dict) or t.get('format') not in {'xlsx','docx'} for t in targets): fail('input_invalid','目标格式只能xlsx或docx')
    rule=request.get('rule'); validate(rule,selections,len(targets))
    source=read_evidence(sources[0],0,{'requirements':[]})
    if not source['capabilities'].get('read'): fail('input_unreadable','来源不可读')
    header_row,headers,cells,records=selected_records(source,rule,selections)
    artifacts=[]; locations=[]; grouped={}; target_hashes=[]
    by_id={field['id']:field for field in rule['fields']}
    for spec,target_item in zip(rule['artifacts'],targets):
        if target_item['format']!=spec['format']: fail('input_format','目标顺序或格式与规则不符')
        target=_read(target_item,'target',target_mode='review_values')
        target_hashes.append(target['sha256'])
        if not target['capabilities'].get('read') or not target['capabilities'].get('write'): fail('target_unwritable','目标不可写或受保护')
        requirements=[]; has_existing_values=False
        target_guards=[]
        if spec['format']=='xlsx':
            labels={k['target_label'] for k in spec['keys']}|{f['target_label'] for f in spec['fields']}; tr,th=header(target['nodes'],spec['target_sheet'],labels)
            target_guards=[{'ref':th[k['target_label']]['ref'],'text':k['target_label']} for k in spec['keys']]
            target_keys, _=record_keys(target['nodes'],spec['target_sheet'],tr,[column(th[k['target_label']]['ref']) for k in spec['keys']])
            writable={slot['target_ref'] for slot in target['slots']}
            for name,record in records.items():
                target_record=target_keys.get(record['key'])
                if not target_record: fail('record_missing','目标中缺少本批选择记录',record['refs'])
                for item in spec['fields']:
                    ref=spec['target_sheet']+'!'+column(th[item['target_label']]['ref'])+str(target_record['row'])
                    if ref not in writable: fail('target_unwritable','目标字段不可写',[ref])
                    requirements.append(source_requirement(by_id[item['field_id']],record['row'],headers,cells,rule['source_sheet'],ref))
                    grouped.setdefault((item['field_id'],name),[]).append({'output_index':spec['output_index'],'target_ref':ref})
                locations.append({'output_index':spec['output_index'],'record_selector':name,'source_row':record['row'],'target_row':target_record['row']})
        else:
            record=records[spec['record_selector']]
            for item in spec['fields']:
                ref,expected,guard=docx_target(target,item['target_label'])
                target_guards.append(guard)
                if expected is not None:
                    has_existing_values=True
                    target_guards.append({'ref':ref,'text':expected})
                requirement=source_requirement(by_id[item['field_id']],record['row'],headers,cells,rule['source_sheet'],ref)
                if 'decimals' in requirement: requirement['render']='fixed_decimal'
                requirements.append(requirement)
                grouped.setdefault((item['field_id'],spec['record_selector']),[]).append({'output_index':spec['output_index'],'target_ref':ref})
            locations.append({'output_index':spec['output_index'],'record_selector':spec['record_selector'],'source_row':record['row']})
        contract={'coverage':[r['target_ref'] for r in requirements],'requirements':requirements,'preserve_template':True}
        if has_existing_values: contract['allow_existing_values']=True
        artifacts.append({'output_index':spec['output_index'],'contract':contract,'conditions':{'source_count':1,'source_formats':['xlsx'],'target_format':spec['format'],'source_guards':[{'source_index':0,'ref':headers[k]['ref'],'text':k} for k in rule['keys']],'target_guards':target_guards}})
    consistency=[]
    wanted={item['field_id']:item['kind'] for item in rule['consistency_fields']}
    for (field_id,name),refs in grouped.items():
        if field_id in wanted:
            if len(refs)<2: fail('consistency_missing','一致字段在本批记录中必须出现在至少2份产物',[field_id,name])
            consistency.append({'label':field_id+'@'+name,'kind':wanted[field_id],'refs':refs})
    if not consistency: fail('consistency_missing','未生成跨产物一致字段')
    if sum(len(a['contract']['requirements']) for a in artifacts)>10000: fail('coverage_limit','本批展开字段超过10000')
    for field_id in wanted:
        entries=[refs for (current,_),refs in grouped.items() if current==field_id]
        if not entries or any(len({ref['output_index'] for ref in refs})<2 for refs in entries): fail('consistency_missing','每个一致字段的每条本批记录需跨至少2份产物',[field_id])
    return {'ok':True,'operation':'bind_group_relationship','artifacts':artifacts,'consistency':consistency,'source_sha256':[source['sha256']],'target_sha256':target_hashes,'field_count':sum(len(a['contract']['requirements']) for a in artifacts),'record_count':len(records),'locations':locations}
