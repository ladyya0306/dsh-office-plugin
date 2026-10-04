"""Deterministic XLSX mapping preparation with no paths, I/O, or model access."""
from __future__ import annotations

from decimal import Decimal, InvalidOperation
from datetime import date, datetime, time
import math
from pathlib import Path
import re
import sys
from typing import Any, Callable


# ``python -I -B worker/office_worker.py`` omits the project root from
# sys.path.  This trusted, file-derived insertion is the only import-path
# adjustment; it never reads a user-provided path.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from vendor_candidates.value_fit import ValueFitError, fit_value


MONEY_UNITS = {"元", "千元", "万元", "亿元"}
PERCENT_UNIT = "%"
MAX_SAFE_INTEGER = 2**53 - 1


class PrepareError(ValueError):
    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


def _unit(ws: Any, cell: Any, numeric_context: bool) -> str:
    value = ws.cell(cell.row, cell.column + 1).value
    rendered = "" if value is None else str(value).strip()
    if not rendered or rendered == "单位":
        return ""
    if rendered in MONEY_UNITS | {PERCENT_UNIT}:
        return rendered
    # A text field may sit beside the next business value in a horizontal
    # layout.  Only a numeric/money mapping treats an unknown neighbour as a
    # declared (and therefore rejectable) unit; text mappings do not guess.
    return rendered if numeric_context else ""


def _numeric_or_money(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float, Decimal)):
        return True
    return isinstance(value, str) and bool(re.fullmatch(r"\s*(?:人民币\s*)?\d+(?:[,.]\d+)?(?:亿元|万元|千元|元)?\s*", value))


def _scalar(value: Any, reference: str) -> Any:
    if value is None or isinstance(value, (dict, list, tuple, set)):
        raise PrepareError("invalid_source", f"source_ref 必须提供标量值：{reference}")
    if isinstance(value, (datetime, date, time)):
        raise PrepareError("unsupported_value_type", f"日期/时间来源暂不支持，拒绝隐式序列化：{reference}")
    if isinstance(value, int) and not isinstance(value, bool) and abs(value) > MAX_SAFE_INTEGER:
        raise PrepareError("numeric_precision", f"整数超过 JSON 安全范围 ±{MAX_SAFE_INTEGER}：{reference}")
    if isinstance(value, float) and (not math.isfinite(value) or Decimal(str(value)) != Decimal(repr(value))):
        raise PrepareError("numeric_precision", f"浮点数不是可安全 JSON 表达的有限值：{reference}")
    return value


def _number(rendered: str, source_ref: str, target_ref: str) -> int | float:
    try:
        decimal = Decimal(rendered)
    except InvalidOperation as exc:
        raise PrepareError("unit_conversion_failed", f"金额换算未得到数字：{source_ref} → {target_ref}") from exc
    if not decimal.is_finite():
        raise PrepareError("unit_conversion_failed", f"金额换算结果不是有限数字：{source_ref} → {target_ref}")
    if decimal == decimal.to_integral_value():
        integer = int(decimal)
        if abs(integer) > MAX_SAFE_INTEGER:
            raise PrepareError("numeric_precision", f"换算整数超过 JSON 安全范围 ±{MAX_SAFE_INTEGER}：{source_ref} → {target_ref}")
        return integer
    converted = float(decimal)
    if not math.isfinite(converted) or Decimal(str(converted)) != decimal:
        raise PrepareError("numeric_precision", f"换算小数不能无损表示为 JSON 数字：{source_ref} → {target_ref}")
    return converted


def _money_text(value: Any, unit: str, source_ref: str) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal, str)):
        raise PrepareError("invalid_source", f"金额来源必须是数字或数字文本：{source_ref}")
    text = str(value).strip()
    if not text:
        raise PrepareError("invalid_source", f"金额来源不能为空：{source_ref}")
    for known in MONEY_UNITS | {PERCENT_UNIT}:
        if text.endswith(known):
            raise PrepareError("unit_conflict", f"来源单元格已含单位且相邻单位也已声明：{source_ref}")
    return text + unit


def _derive_value(value: Any, source_unit: str, target_unit: str, source_ref: str, target_ref: str) -> tuple[Any, dict[str, str]]:
    if source_unit == target_unit == "":
        return _scalar(value, source_ref), {"kind": "identity", "source_unit": "", "target_unit": ""}
    if source_unit == PERCENT_UNIT and target_unit == PERCENT_UNIT:
        if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
            raise PrepareError("unit_conflict", f"百分比来源必须是数值单元格，按 Excel 原始数值保留：{source_ref}")
        return value, {"kind": "percent_raw", "source_unit": "%", "target_unit": "%",
                       "rule": "来源和目标均为 %；不乘除 100，保留 Excel 原始数值"}
    if source_unit in MONEY_UNITS and target_unit in MONEY_UNITS:
        if source_unit == target_unit:
            return _scalar(value, source_ref), {"kind": "same_money_unit", "source_unit": source_unit,
                                                 "target_unit": target_unit, "rule": "同单位保留来源原值"}
        try:
            rendered = fit_value("____" + target_unit, 0, 4, _money_text(value, source_unit, source_ref))
        except ValueFitError as exc:
            raise PrepareError("unit_conversion_failed", f"金额单位不能确定性换算：{source_ref} → {target_ref}；{exc}") from exc
        return _number(rendered, source_ref, target_ref), {
            "kind": "money_unit_conversion", "source_unit": source_unit, "target_unit": target_unit,
            "rule": "仅元、千元、万元、亿元间精确十进制换算",
        }
    if source_unit and source_unit not in MONEY_UNITS | {PERCENT_UNIT}:
        raise PrepareError("unknown_unit", f"来源单位不在允许集合（元、千元、万元、亿元、%）：{source_unit}")
    if target_unit and target_unit not in MONEY_UNITS | {PERCENT_UNIT}:
        raise PrepareError("unknown_unit", f"目标单位不在允许集合（元、千元、万元、亿元、%）：{target_unit}")
    raise PrepareError("unit_conflict", f"单位不一致，拒绝猜测：{source_unit or '无单位'} → {target_unit or '无单位'}")


def derive_changes(
    mappings: Any,
    resolve_source: Callable[[Any], tuple[Any, Any, str]],
    resolve_target: Callable[[Any], tuple[Any, Any, str]],
    assert_safe_target: Callable[[Any, Any, str], None],
) -> list[dict[str, Any]]:
    """Return deterministic values for source/target mappings, without writes."""
    if not isinstance(mappings, list) or not mappings:
        raise PrepareError("invalid_mappings", "mappings 必须是非空数组")
    changes: list[dict[str, Any]] = []
    seen_targets: set[str] = set()
    for mapping in mappings:
        if not isinstance(mapping, dict) or set(mapping) != {"source_ref", "target_ref"}:
            raise PrepareError("invalid_mapping", "每项 mappings 只能包含 source_ref 和 target_ref；不得提交 value")
        source_ws, source_cell, source_ref = resolve_source(mapping["source_ref"])
        target_ws, target_cell, target_ref = resolve_target(mapping["target_ref"])
        if target_ref in seen_targets:
            raise PrepareError("duplicate_target", f"target_ref 重复：{target_ref}")
        seen_targets.add(target_ref)
        assert_safe_target(target_ws, target_cell, target_ref)
        value = _scalar(source_cell.value, source_ref)
        numeric_context = _numeric_or_money(value)
        derived, conversion = _derive_value(value, _unit(source_ws, source_cell, numeric_context), _unit(target_ws, target_cell, numeric_context),
                                             source_ref, target_ref)
        changes.append({"source_ref": source_ref, "target_ref": target_ref, "value": derived,
                        "conversion": conversion})
    return changes
