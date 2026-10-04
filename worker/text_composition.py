"""Strict composition of independently verified values into fixed target text."""
from decimal import Decimal
import re


class TextCompositionError(ValueError):
    pass


_ROOT_KEYS = {'operation', 'target_ref', 'expected_text', 'text_template', 'components'}
_COMPONENT_KEYS = {
    'source_refs', 'source_matches', 'operation', 'text_extract', 'source_unit', 'target_unit',
    'source_unit_ref', 'target_unit_source_ref', 'decimals', 'source_labels', 'source_label_refs',
    'source_label', 'filename_identity', 'expression', 'blank_markers', 'source_absence',
}


def _parse_template(template):
    if not isinstance(template, str) or not template or len(template) > 4096:
        raise TextCompositionError('text_template需为1至4096字符的明确固定文本模板')
    tokens, literals, index = [], [], 0
    literal = []
    while index < len(template):
        char = template[index]
        if char == '{':
            if literal:
                piece = ''.join(literal); tokens.append(('literal', piece)); literals.append(piece); literal = []
            close = template.find('}', index + 1)
            if close < 0:
                raise TextCompositionError('text_template占位符缺少右花括号')
            raw = template[index + 1:close]
            if not raw.isdecimal() or str(int(raw)) != raw:
                raise TextCompositionError('text_template只支持无格式参数的{0}至{15}')
            tokens.append(('component', int(raw))); index = close + 1; continue
        if char == '}':
            raise TextCompositionError('text_template不支持转义花括号或未知格式语法')
        literal.append(char); index += 1
    if literal:
        piece = ''.join(literal); tokens.append(('literal', piece)); literals.append(piece)
    return tokens, literals


def _guard_fixed_text(expected_text, literals):
    if not expected_text.strip():
        return
    if not literals:
        raise TextCompositionError('非空模板需声明并保留固定文字边界，不能整段清除')
    position = 0
    for literal in (item for item in literals if item):
        if expected_text.count(literal) != 1:
            raise TextCompositionError('原非空文字中的固定片段未唯一保留：' + literal)
        found = expected_text.find(literal, position)
        if found < position:
            raise TextCompositionError('原非空文字中的固定片段顺序不一致')
        position = found + len(literal)
    fixed = ''.join(literals)
    for token in re.findall(r'人民币|美元|港元|亿元|万元|千元|元', expected_text):
        if fixed.count(token) < expected_text.count(token):
            raise TextCompositionError('模板固定币种或金额单位必须保留：' + token)


def _component(item, index, target_ref):
    if not isinstance(item, dict) or set(item) - _COMPONENT_KEYS:
        raise TextCompositionError('components含未知参数')
    if item.get('operation') == 'compose':
        raise TextCompositionError('components禁止嵌套compose')
    if 'expected_value' in item or 'target_ref' in item:
        raise TextCompositionError('components不能携带答案或目标位置')
    if not isinstance(item.get('operation'), str) or not item['operation']:
        raise TextCompositionError('第%d个组件需声明operation' % index)
    result = dict(item)
    # _expected does not consume target_ref for source arithmetic, but passing a
    # fixed current target keeps evaluator contracts uniform and auditable.
    result['target_ref'] = target_ref
    return result


def _render_value(value, component):
    if value is None:
        if component.get('operation') != 'blank':
            raise TextCompositionError('只有带blank依据的组件可组合为空文字')
        return ''
    if isinstance(value, Decimal):
        decimals = component.get('decimals')
        if decimals is not None:
            if isinstance(decimals, bool) or not isinstance(decimals, int) or not 0 <= decimals <= 10:
                raise TextCompositionError('decimals需为0至10整数')
            return format(value, '.%df' % decimals)
        return str(value)
    return str(value)


def compose_expected(requirement, nodes, template_nodes, evaluator):
    """Evaluate components and render them into one fixed current target text.

    ``evaluator`` is the host's normal per-field expected-value function (for
    example ``output_check._expected``).  It is called once for each component;
    errors deliberately propagate rather than being converted to empty text.
    """
    if not isinstance(requirement, dict) or set(requirement) != _ROOT_KEYS:
        raise TextCompositionError('compose要求只包含operation、target_ref、expected_text、text_template、components')
    if requirement.get('operation') != 'compose':
        raise TextCompositionError('compose要求operation=compose')
    target_ref = requirement.get('target_ref')
    expected_text = requirement.get('expected_text')
    if not isinstance(target_ref, str) or not target_ref or not isinstance(expected_text, str) or len(expected_text) > 4096:
        raise TextCompositionError('compose目标引用或expected_text无效')
    if not isinstance(template_nodes, dict) or target_ref not in template_nodes:
        raise TextCompositionError('compose目标引用不属于本次模板读取')
    node = template_nodes[target_ref]
    current = None if not isinstance(node, dict) else node.get('value')
    if ('' if current is None else current) != expected_text:
        raise TextCompositionError('模板当前位置原文与expected_text guard不一致')
    components = requirement.get('components')
    if not isinstance(components, list) or not 1 <= len(components) <= 16:
        raise TextCompositionError('components需为1至16项')
    tokens, literals = _parse_template(requirement.get('text_template'))
    indexes = [value for kind, value in tokens if kind == 'component']
    if sorted(indexes) != list(range(len(components))) or len(set(indexes)) != len(indexes):
        raise TextCompositionError('text_template必须恰好一次引用每个组件编号')
    placeholder_span = (node.get('write_kind') == 'text_span'
                        and re.fullmatch(r'[_＿□☐]{2,}', expected_text) is not None)
    if not placeholder_span:
        _guard_fixed_text(expected_text, literals)
    if not callable(evaluator):
        raise TextCompositionError('compose需要明确的组件求值器')
    values = []
    for index, item in enumerate(components):
        component = _component(item, index, target_ref)
        values.append(_render_value(evaluator(component, nodes, template_nodes), component))
    return ''.join(values[value] if kind == 'component' else value for kind, value in tokens)
