"""Task-scoped value adaptation for fixed text already printed in a template.

Origin: D:\\DSH-worktrees\\fill-efficiency\\toolV2\\office_kit\\value_fit.py
Git blob: 55e0874579099db252a754662614c71a23c7531b
Source SHA-256: F08DA5D9C29B6ABA2DBAD6ADCBD859A814635867CF167DCB8D65D306ACA0034B

The source repository has no declared LICENSE, NOTICE, or COPYING file.  This
is a task-scoped internal reuse under the requesting user's authorization, not
a public licence grant; do not label or redistribute it as MIT or any other
licence until the source owner supplies one.

This module deliberately has no filesystem, Office-package, model, shell, or
workflow dependency.  It only adapts a supplied string to literal text which a
template has already printed.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation, localcontext
import re


class ValueFitError(ValueError):
    """A proposed value cannot safely fit the already printed template text."""


_PRINTED_ENDINGS = ("有限公司", "万元", "千元", "人民币", "公司", "元", "年", "号")
_YUAN_AMOUNT = re.compile(r"(?:人民币\s*)?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?\s*(?:万|亿)?")
_AMOUNT_VALUE = re.compile(
    r"^\s*(?:人民币\s*)?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?\s*(?P<unit>亿元|万元|千元|元|亿|万)?\s*$"
)
_AMOUNT_UNITS = ("万元", "亿元", "千元", "元", "万", "亿")
_AMOUNT_FACTORS = {"元": Decimal(1), "千元": Decimal(1000), "万": Decimal(10_000), "万元": Decimal(10_000),
                   "亿": Decimal(100_000_000), "亿元": Decimal(100_000_000)}
_VALUE_UNITS = ("万元", "亿元", "千元", "元", "年", "个月", "月", "日", "人", "户", "份", "笔", "个", "次", "家", "%")
_UNIT_BLANK_RE = re.compile(r"(?:[ \u3000]+|_{2,}|＿{2,}|[□☐]+)(?=(?:万元|亿元|千元|元|年|个月|月|日|人|户|份|笔|个|次|家|%))")
_PLACEHOLDER_MARK_RE = re.compile(r"[ \u3000]+|_{2,}|＿{2,}|[□☐]+")
_DATE_PART = r"\d{4}(?:\s*年(?:\s*\d{1,2}\s*月(?:\s*\d{1,2}\s*日?)?)?|[-./]\d{1,2}(?:[-./]\d{1,2})?)?"
_DATE_RANGE = re.compile(_DATE_PART + r"\s*起?\s*(?:至|到|[-—–~～]+)\s*" + _DATE_PART)


def _checked_span(text: str, start: int, end: int) -> None:
    if not isinstance(text, str) or not isinstance(start, int) or not isinstance(end, int):
        raise ValueFitError("模板文本和位置必须分别为字符串与整数")
    if start < 0 or end < start or end > len(text):
        raise ValueFitError("填写位置超出模板文本范围")


def _printed_amount_unit(text: str, start: int, end: int) -> str | None:
    right = text[end:].lstrip()
    suffix = next((unit for unit in _AMOUNT_UNITS if right.startswith(unit)), None)
    if suffix:
        return suffix
    label_unit = re.search(r"[（(]\s*(亿元|万元|千元|元)\s*[）)]\s*[：:]?\s*$", text[:start])
    return label_unit.group(1) if label_unit else None


def value_fit_issue(text: str, start: int, end: int, source_value: str) -> str | None:
    """Return an explainable incompatibility without changing the template."""
    _checked_span(text, start, end)
    right = text[end:].lstrip()
    value = str(source_value).strip()
    if right.startswith(("年", "个月", "月")) and _DATE_RANGE.search(value):
        unit = "年" if right.startswith("年") else ("个月" if right.startswith("个月") else "月")
        return f"模板此处固定印有“{unit}”，需要期限数量；来源值是日期区间，不能直接填写"
    amount_unit = _printed_amount_unit(text, start, end)
    if amount_unit and not _AMOUNT_VALUE.fullmatch(value):
        return f"模板此处固定印有人民币金额单位“{amount_unit}”，来源值不是可识别的人民币数字金额；不能进行跨币种换算"
    return None


def unit_wrapper_span(text: str) -> tuple[int, int] | None:
    """Return the sole blank immediately before a known unit, else ``None``."""
    if not isinstance(text, str):
        return None
    candidates = []
    for match in _UNIT_BLANK_RE.finditer(text):
        right = text[match.end():].lstrip()
        if any(right.startswith(unit) for unit in _VALUE_UNITS):
            candidates.append(match.span())
    if len(candidates) != 1 or len(list(_PLACEHOLDER_MARK_RE.finditer(text))) != 1:
        return None
    return candidates[0]


def _converted_amount(value: str, target_unit: str) -> str:
    match = _AMOUNT_VALUE.fullmatch(value.strip())
    if not match:
        raise ValueFitError(f"模板此处固定印有人民币金额单位“{target_unit}”，来源值不是可识别的人民币数字金额；不能进行跨币种换算")
    source_unit = match.group("unit")
    number = re.sub(r"[ ,，\s]", "", value.strip())
    number = re.sub(r"^人民币", "", number).strip()
    number = re.sub(r"(?:亿元|万元|千元|元|亿|万)$", "", number).strip()
    try:
        with localcontext() as context:
            context.prec = max(50, len(number) + 20)
            amount = Decimal(number)
            if source_unit:
                amount = amount * _AMOUNT_FACTORS[source_unit] / _AMOUNT_FACTORS[target_unit]
    except InvalidOperation as exc:
        raise ValueFitError("来源金额无法精确解析，不能按模板单位换算") from exc
    rendered = format(amount.normalize(), "f")
    return "0" if rendered in ("-0", "") else rendered


def fit_value(text: str, start: int, end: int, source_value: str) -> str:
    """Adapt a value without inventing units or overwriting fixed literal text."""
    issue = value_fit_issue(text, start, end, source_value)
    if issue:
        raise ValueFitError(issue)
    value = str(source_value)
    left = text[:start].rstrip()
    right = text[end:].lstrip()
    printed_amount_unit = _printed_amount_unit(text, start, end)
    if printed_amount_unit:
        value = _converted_amount(value, printed_amount_unit)
    if left.endswith("人民币") and value.strip().startswith("人民币"):
        value = value.strip()[len("人民币"):].lstrip()
    number_prefix = re.search(r"(?:合同编号(?:为)?|编号为)\s*[：:]?\s*([^\n，。；;（）()【】]+)$", left)
    if number_prefix and right.startswith("号"):
        prefix = number_prefix.group(1).strip()
        if prefix and value.strip().startswith(prefix):
            value = value.strip()[len(prefix):].lstrip()
    if right.startswith("元"):
        amount_before_printed_unit = value[:-1] if value.endswith("元") else value
        if not _YUAN_AMOUNT.fullmatch(amount_before_printed_unit.strip()):
            raise ValueFitError("模板已印“元”，源值不是可直接拼接的人民币数字金额")
    if left.endswith("【") and right.startswith("】"):
        if value.startswith("【"):
            value = value[1:]
        if value.endswith("】号"):
            value = value[:-2]
        elif value.endswith("】"):
            value = value[:-1]
    for ending in _PRINTED_ENDINGS:
        if right.startswith(ending) and value.endswith(ending):
            value = value[:-len(ending)]
            break
    for opening, closing in (("（", "）"), ("(", ")"), ("《", "》")):
        if left.endswith(opening) and right.startswith(closing) and value.startswith(opening) and value.endswith(closing):
            value = value[1:-1]
            break
    if not value and str(source_value).strip():
        raise ValueFitError("值与模板固定文字完全重叠，无法判断是否应留空")
    return value
