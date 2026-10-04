"""Bounded data rules for explicit text extraction and Decimal expressions.

No file access, eval, imports from saved scripts, or arbitrary regular expressions.
Business meaning and the chosen delimiters/formula remain the caller's declaration.
"""
from decimal import Decimal


def extract_text(value, spec):
    if not isinstance(value, str) or len(value) > 4096:
        raise ValueError('文本拆分需要不超过4096字符的明确来源文字')
    allowed = {'after', 'before', 'through', 'cast', 'normalize_parentheses'}
    if not isinstance(spec, dict) or not spec or set(spec) - allowed:
        raise ValueError('text_extract只接受after、before、through、cast和normalize_parentheses')
    if spec.get('before') and spec.get('through'):
        raise ValueError('截取末端before与through二选一')
    result = value
    if spec.get('normalize_parentheses'):
        result = result.replace('（', '(').replace('）', ')')
    for key in ('after', 'before', 'through'):
        delimiter = spec.get(key)
        if delimiter is None:
            continue
        if not isinstance(delimiter, str) or not 1 <= len(delimiter) <= 128:
            raise ValueError('文本边界需1至128字符的明确文字')
        if spec.get('normalize_parentheses'):
            delimiter = delimiter.replace('（', '(').replace('）', ')')
        # Ambiguous repeated anchors must be resolved by the caller, not guessed.
        if result.count(delimiter) != 1:
            raise ValueError('文本边界未找到或不唯一：' + delimiter)
        index = result.index(delimiter)
        if key == 'after':
            result = result[index + len(delimiter):]
        else:
            result = result[:index + (len(delimiter) if key == 'through' else 0)]
    result = result.strip()
    cast = spec.get('cast', 'text')
    if cast == 'integer':
        if not result or len(result) > 100 or not result.isascii() or not result.isdecimal():
            raise ValueError('拆分结果不是明确的非负整数')
        return int(result)
    if cast != 'text' or not result:
        raise ValueError('文本拆分类型需text或integer；空值需另列缺失依据')
    return result


def evaluate_expression(expression, numbers):
    """A small arithmetic tree with indexed current source values, never Python."""
    visited = 0

    def evaluate(node, depth):
        nonlocal visited
        visited += 1
        if visited > 64 or depth > 12 or not isinstance(node, dict):
            raise ValueError('明确算式最多64个节点、12层')
        if set(node) == {'source'}:
            index = node['source']
            if isinstance(index, bool) or not isinstance(index, int) or not 0 <= index < len(numbers):
                raise ValueError('算式来源编号不属于本次明确值')
            return numbers[index]
        if set(node) == {'constant'}:
            raw = node['constant']
            if isinstance(raw, bool) or raw is None or len(str(raw)) > 128:
                raise ValueError('算式常数需明确有限数值')
            number = Decimal(str(raw))
            if not number.is_finite() or abs(number.adjusted()) > 100:
                raise ValueError('算式常数超出数值范围')
            return number
        if set(node) != {'operator', 'args'} or not isinstance(node['args'], list):
            raise ValueError('算式节点需source、constant或operator与args')
        op = node['operator']
        values = [evaluate(item, depth + 1) for item in node['args']]
        if not 1 <= len(values) <= 32:
            raise ValueError('算式每项需1至32个明确参与值')
        if op in {'sum', 'average'}:
            value = sum(values, Decimal(0))
            if op == 'average':
                value /= len(values)
        elif op == 'product':
            value = Decimal(1)
            for number in values:
                value *= number
        elif op == 'difference' and len(values) == 2:
            value = values[0] - values[1]
        elif op in {'ratio', 'percent'} and len(values) == 2:
            if values[1] == 0:
                raise ValueError('分母为0，不能默认结果0')
            value = values[0] / values[1]
            if op == 'percent':
                value *= 100
        else:
            raise ValueError('算式仅支持和、平均、积、差、比、百分比')
        if not value.is_finite() or abs(value.adjusted()) > 100:
            raise ValueError('算式中间结果超出数值范围')
        return value

    return evaluate(expression, 0)


def explicit_label(node, label, label_ref, nodes, value_ref):
    """Verify caller-selected label against returned physical row/column evidence."""
    evidence = nodes.get(label_ref)
    if not evidence or evidence.get('formula') or evidence.get('value') != label:
        raise ValueError('来源标签引用不存在或原文不符')
    if label_ref not in node.get('row_header_refs', []) + node.get('column_header_refs', []):
        raise ValueError('来源标签不属于该值本次返回的同行或同列上下文')
    if label_ref.split('::', 1)[0] != value_ref.split('::', 1)[0]:
        raise ValueError('来源标签和数值不能跨来源文件')


def blank_expected(requirement, nodes):
    refs = requirement.get('source_refs', [])
    absence = requirement.get('source_absence')
    if refs and absence:
        raise ValueError('明确空值与栏目缺失两种依据不能混用')
    if refs:
        markers = requirement.get('blank_markers')
        if not isinstance(markers, list) or not markers or len(markers) > 16:
            raise ValueError('来源明确空值需要逐项声明blank_markers')
        if any(isinstance(v, (bool, int, float, dict, list)) for v in markers):
            raise ValueError('空值标记只接受文字或null；真实0不能视为缺失')
        for ref in refs:
            node = nodes.get(ref)
            if not node or node.get('formula') or node.get('value_error') or node.get('value') not in markers:
                raise ValueError('当前来源不是声明的空值标记')
        labels=requirement.get('source_labels')
        label_refs=requirement.get('source_label_refs')
        if labels is not None and (not isinstance(labels,list) or len(labels)!=len(refs)):
            raise ValueError('空值来源标签需逐项对应')
        if label_refs is not None and (not isinstance(label_refs,list) or len(label_refs)!=len(refs)):
            raise ValueError('空值来源标签引用需逐项对应')
        for index,ref in enumerate(refs):
            node=nodes[ref]
            label=labels[index] if labels else requirement.get('source_label')
            if label and label_refs:
                explicit_label(node,label,label_refs[index],nodes,ref)
            elif label and label not in node.get('row_headers',[])[:1]+node.get('column_headers',[])[:1]:
                raise ValueError('空值来源主标签不符')
        return None
    if not isinstance(absence, dict) or set(absence) != {'source_index', 'sheet', 'labels'}:
        raise ValueError('栏目缺失需source_absence声明来源编号、工作表及全部待查标签')
    index, sheet, labels = absence['source_index'], absence['sheet'], absence['labels']
    if isinstance(index, bool) or not isinstance(index, int) or not 0 <= index <= 7:
        raise ValueError('缺失检查来源编号无效')
    if not isinstance(sheet, str) or not sheet or not isinstance(labels, list) or not 1 <= len(labels) <= 32:
        raise ValueError('缺失检查需工作表及1至32个明确标签')
    if any(not isinstance(label, str) or not label.strip() or len(label) > 128 for label in labels):
        raise ValueError('缺失检查标签无效')
    meta = nodes.get(f's{index}::$read_scope', {})
    if not meta.get('all_text_nodes_read') or sheet not in meta.get('sheets', []):
        raise ValueError('未读完声明工作表的文字格，不能据此确认栏目缺失')
    normalized = {label.strip().rstrip(':：') for label in labels}
    for ref, node in nodes.items():
        if not ref.startswith(f's{index}::{sheet}!'):
            continue
        if node.get('formula') or node.get('value_error'):
            raise ValueError('声明工作表含未计算公式或读取错误，不能证明栏目不存在')
        value = node.get('value')
        if isinstance(value, str) and value.strip().rstrip(':：') in normalized:
            raise ValueError('本次声明栏目仍然存在，不能按栏目缺失留空')
    return None
