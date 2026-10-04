"""Deterministic locations for writing only declared values inside fixed text.

The module has no document I/O.  A caller first supplies the current physical
paragraph/cell text, then consumes the returned character spans (for DOCX these
map directly to the existing ``|s:start:end`` convention).
"""
from datetime import datetime
from decimal import Decimal, InvalidOperation


class PreciseTextError(ValueError):
    """A fixed-text replacement cannot be established without guessing."""


_PLACEHOLDER = set('_＿□☐—-')
_ALLOWED_PART = {'name', 'kind', 'before', 'after', 'span', 'missing'}


def _text(value, label, minimum=1, maximum=128):
    if not isinstance(value, str) or not minimum <= len(value) <= maximum:
        raise PreciseTextError(label + '需为%d至%d字符的明确文字' % (minimum, maximum))
    return value


def _number(value):
    if not isinstance(value, str) or not value or value != value.strip() or len(value) > 100:
        raise PreciseTextError('数值位置不是明确文本')
    normalized = value.replace(',', '')
    if not normalized or any(char not in '0123456789+-.eE' for char in normalized):
        raise PreciseTextError('数值位置含未声明文字')
    try:
        result = Decimal(normalized)
    except InvalidOperation as exc:
        raise PreciseTextError('数值位置不是有限十进制数') from exc
    if not result.is_finite():
        raise PreciseTextError('数值位置不是有限十进制数')
    return result


def _date(value):
    if not isinstance(value, str) or not value or value != value.strip() or len(value) > 20:
        raise PreciseTextError('日期位置不是明确日期文本')
    for pattern in ('%Y-%m-%d', '%Y/%m/%d', '%Y年%m月%d日'):
        try:
            return datetime.strptime(value, pattern).date().isoformat()
        except ValueError:
            pass
    raise PreciseTextError('日期必须完整为YYYY-MM-DD、YYYY/MM/DD或YYYY年M月D日；尾部杂字不接受')


def _missing(value, policy):
    if policy not in {'keep', 'empty'}:
        raise PreciseTextError('缺项规则只能为keep或empty')
    if policy == 'empty':
        if value != '':
            raise PreciseTextError('缺项规则empty要求当前位置为空')
    elif not value or any(char not in _PLACEHOLDER for char in value):
        raise PreciseTextError('缺项规则keep要求当前位置是明确占位符')
    return None


def _anchor_span(text, before, after):
    before_positions = []
    offset = 0
    while True:
        found = text.find(before, offset)
        if found < 0:
            break
        before_positions.append(found)
        offset = found + 1
    candidates = []
    for start in before_positions:
        content_start = start + len(before)
        offset = content_start
        while True:
            end = text.find(after, offset)
            if end < 0:
                break
            candidates.append((content_start, end))
            offset = end + 1
    if len(candidates) != 1:
        raise PreciseTextError('before/after边界未找到或不唯一')
    return candidates[0]


def _part(text, part):
    if not isinstance(part, dict) or set(part) - _ALLOWED_PART:
        raise PreciseTextError('文本分项含未知参数')
    name = _text(part.get('name'), '分项名称')
    kind = part.get('kind')
    if kind not in {'number', 'date', 'text'}:
        raise PreciseTextError('分项类型只能为number、date或text')
    missing = part.get('missing')
    anchored = 'before' in part or 'after' in part
    explicit = 'span' in part
    if anchored == explicit:
        raise PreciseTextError('每个分项需唯一选择before/after或span边界')
    if anchored:
        before, after = _text(part.get('before'), 'before边界'), _text(part.get('after'), 'after边界')
        start, end = _anchor_span(text, before, after)
    else:
        span = part.get('span')
        if (not isinstance(span, list) or len(span) != 2 or any(isinstance(x, bool) or not isinstance(x, int) for x in span)
                or not 0 <= span[0] <= span[1] <= len(text)):
            raise PreciseTextError('span需为当前位置内的[start,end]整数边界')
        start, end = span
        before, after = text[:start], text[end:]
    current = text[start:end]
    if missing is not None:
        value = _missing(current, missing)
    elif kind == 'number':
        value = _number(current)
    elif kind == 'date':
        value = _date(current)
    else:
        if not current or current != current.strip() or len(current) > 512:
            raise PreciseTextError('文本位置需为非空且无首尾空白的明确文字')
        value = current
    return {'name': name, 'kind': kind, 'start': start, 'end': end,
            'prefix': text[:start], 'suffix': text[end:], 'current_text': current,
            'value': value, 'missing': missing}


def locate_precise_text(text, specification):
    """Return verified, non-overlapping text spans from one current text value.

    ``specification`` is deliberately small::

        {"expected_text": current_text, "parts": [
          {"name": "amount", "kind": "number", "before": "金额：", "after": "万元"},
          {"name": "start", "kind": "date", "before": "期限：", "after": "至"},
        ]}

    ``expected_text`` is a guard against applying stale coordinates.  A part may
    instead give one explicit ``span``.  Missing values must expressly choose
    ``keep`` (placeholder) or ``empty``.  No caller-supplied regular expression
    is accepted.
    """
    if not isinstance(text, str) or len(text) > 4096:
        raise PreciseTextError('当前固定文字需为不超过4096字符的字符串')
    if not isinstance(specification, dict) or set(specification) != {'expected_text', 'parts'}:
        raise PreciseTextError('精确文本规则只接受expected_text和parts')
    if specification['expected_text'] != text:
        raise PreciseTextError('当前文字与expected_text guard不一致')
    parts = specification['parts']
    if not isinstance(parts, list) or not 1 <= len(parts) <= 16:
        raise PreciseTextError('精确文本规则需1至16个分项')
    located = [_part(text, item) for item in parts]
    if len({item['name'] for item in located}) != len(located):
        raise PreciseTextError('分项名称不能重复')
    ordered = sorted(located, key=lambda item: (item['start'], item['end']))
    if any(left['end'] > right['start'] for left, right in zip(ordered, ordered[1:])):
        raise PreciseTextError('分项位置重叠')
    return {'expected_text': text, 'parts': located}
