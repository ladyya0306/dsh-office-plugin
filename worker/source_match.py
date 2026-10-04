"""Resolve current XLSX rows by explicit keys from the untouched template."""
import re


def source_match(spec,sources,template):
    if not isinstance(spec,dict) or not isinstance(spec.get('sheet'),str):
        raise ValueError('按记录查数需要明确工作表')
    column=spec.get('value_column');keys=spec.get('key_columns')
    if not isinstance(column,str) or not re.fullmatch('[A-Z]{1,3}',column) or not isinstance(keys,list) or not 1<=len(keys)<=8:
        raise ValueError('按记录查数需值所在列及1至8个明确匹配键')
    expected=[]
    for key in keys:
        if not isinstance(key,dict) or not re.fullmatch('[A-Z]{1,3}',str(key.get('column',''))):
            raise ValueError('匹配键需明确来源列')
        target=template.get(key.get('target_ref'))
        if not target or target.get('value') is None or target.get('formula') or target.get('value_error'):
            raise ValueError('原模板的匹配键缺失或未经计算，不能猜')
        expected.append((key['column'],target['value'],target.get('value_type')))
    start=spec.get('start_row',2)
    if isinstance(start,bool) or not isinstance(start,int) or start<1:raise ValueError('起始数据行不明确')
    candidates=[];first=expected[0][0]
    pattern=re.compile(r'^(s[0-7])::'+re.escape(spec['sheet'])+'!'+re.escape(first)+r'([1-9][0-9]*)$')
    for ref,node in sources.items():
        match=pattern.fullmatch(ref)
        if not match or int(match[2])<start:continue
        prefix=match[1]+'::'+spec['sheet']+'!';row=match[2]
        if all(sources.get(prefix+col+row,{}).get('value')==value and sources.get(prefix+col+row,{}).get('value_type')==value_type and not sources.get(prefix+col+row,{}).get('formula') and not sources.get(prefix+col+row,{}).get('value_error') for col,value,value_type in expected):
            candidates.append(prefix+column+row)
    if len(candidates)!=1:
        raise ValueError('按本次原模板的主体/订单/期间应唯一匹配，实际找到'+str(len(candidates))+'条；不能任选或把缺失当0')
    if candidates[0] not in sources:raise ValueError('本次记录存在但所需字段未读到或缺失，不能当0')
    value=sources[candidates[0]]
    if value.get('value') is None or value.get('formula') or value.get('value_error'):
        raise ValueError('本次字段缺失或公式/类型未确认，不能当0')
    return candidates[0]
