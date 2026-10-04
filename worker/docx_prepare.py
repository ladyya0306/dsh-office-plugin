"""Deterministic DOCX reference discovery and mapping preparation."""
from __future__ import annotations

import re
from typing import Any, Callable


class DocxPrepareError(ValueError):
    def __init__(self, code: str, message: str):
        super().__init__(message); self.code = code; self.message = message


LABEL_VALUE = re.compile(r"^\s*(?P<label>[^：:]{1,50})[：:]\s*(?P<value>.+?)\s*$")
NARRATIVE = re.compile(r"(?P<label>[\u4e00-\u9fffA-Za-z0-9]{2,24})(?:为|：|:)\s*(?P<value>[^，。；;]+)")
AMOUNT_NARRATIVE = re.compile(r"(?P<label>[\u4e00-\u9fffA-Za-z0-9]{2,24}?)(?P<value>人民币(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?(?:亿元|万元|千元|元))")
EMPTY_LABEL = re.compile(r"^\s*(?P<label>[^：:]{1,50})[：:]\s*$")
SKIP_LABEL = re.compile(r"说明|备注|提示|签字|审批|日期")
EXPLICIT_UNIT = re.compile(r"[（(]\s*(亿元|万元|千元|元|%)\s*[）)]")


def unit_of(value: Any) -> str:
    text = str(value).strip()
    return next((unit for unit in ("亿元", "万元", "千元", "元", "%") if text.endswith(unit)), "")


def explicit_unit(label: Any) -> str:
    match = EXPLICIT_UNIT.search(str(label))
    return match.group(1) if match else ""


def paragraph_values(text: str) -> list[dict[str, Any]]:
    """Return only explicit label/value facts and their source character spans."""
    exact = LABEL_VALUE.match(text)
    if exact:
        return [{"label": exact.group("label").strip(), "value": exact.group("value").strip(),
                 "start": exact.start("value"), "end": exact.end("value")}]
    matches = list(AMOUNT_NARRATIVE.finditer(text)) + list(NARRATIVE.finditer(text))
    unique: dict[tuple[int, int], dict[str, Any]] = {}
    for match in matches:
        value = match.group("value").strip()
        unique[(match.start("value"), match.end("value"))] = {
            "label": match.group("label").strip(), "value": value,
            "start": match.start("value"), "end": match.end("value"),
        }
    return [unique[key] for key in sorted(unique)]


def source_candidates(document: Any) -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        values = paragraph_values(text)
        for ordinal, item in enumerate(values):
            value = item["value"]
            reference = f"p:{index}" if len(values) == 1 else f"p:{index}:v:{ordinal}"
            found.append({"source_ref": reference, "label": item["label"], "value": value, "unit": unit_of(value),
                          "nearby": {"paragraph": text[:300]}, "span_start": item["start"], "span_end": item["end"]})
    for ti, table in enumerate(document.tables):
        for ri, row in enumerate(table.rows):
            for ci in range(1, len(row.cells)):
                label, value = row.cells[ci - 1].text.strip(), row.cells[ci].text.strip()
                if label and value and not SKIP_LABEL.search(label):
                    found.append({"source_ref": f"t:{ti}:r:{ri}:c:{ci}", "label": label, "value": value,
                                  "unit": unit_of(value), "nearby": {"left": label}})
    return found


def target_slots(document: Any, is_safe: Callable[[Any, str], bool]) -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    for index, paragraph in enumerate(document.paragraphs):
        match = EMPTY_LABEL.match(paragraph.text)
        if match and not SKIP_LABEL.search(match.group("label")) and is_safe(paragraph, f"p:{index}"):
            label = match.group("label").strip()
            found.append({"target_ref": f"p:{index}", "label": label, "unit": explicit_unit(label), "nearby": {"paragraph": paragraph.text}})
    for ti, table in enumerate(document.tables):
        for ri, row in enumerate(table.rows):
            for ci in range(1, len(row.cells)):
                label, cell = row.cells[ci - 1].text.strip(), row.cells[ci]
                ref = f"t:{ti}:r:{ri}:c:{ci}"
                if label and not SKIP_LABEL.search(label) and not cell.text.strip() and is_safe(cell, ref):
                    right = row.cells[ci + 1].text.strip() if ci + 1 < len(row.cells) else ""
                    unit = right if right in {"元", "千元", "万元", "亿元", "%"} else explicit_unit(label)
                    found.append({"target_ref": ref, "label": label, "unit": unit, "nearby": {"left": label, "right": right}})
    return found


def derive_changes(mappings: Any, resolve_source: Callable[[Any], tuple[str, str]], resolve_target: Callable[[Any], tuple[Any, str, str]]) -> list[dict[str, Any]]:
    if not isinstance(mappings, list) or not mappings:
        raise DocxPrepareError("invalid_mappings", "mappings 必须是非空数组")
    changes, seen = [], set()
    for mapping in mappings:
        if not isinstance(mapping, dict) or set(mapping) != {"source_ref", "target_ref"}:
            raise DocxPrepareError("invalid_mapping", "每项 mappings 只能包含 source_ref 和 target_ref；不得提交 value")
        source, source_ref = resolve_source(mapping["source_ref"])
        _target, target_ref, target_unit = resolve_target(mapping["target_ref"])
        if target_ref in seen:
            raise DocxPrepareError("duplicate_target", f"target_ref 重复：{target_ref}")
        seen.add(target_ref)
        value = str(source).strip()
        if not value:
            raise DocxPrepareError("invalid_source", f"source_ref 为空：{source_ref}")
        source_unit = unit_of(value)
        if source_unit and target_unit and source_unit != target_unit:
            raise DocxPrepareError("unsupported_docx_conversion", f"DOCX 显式单位不一致，首版不换算：{source_unit} → {target_unit}")
        changes.append({"source_ref": source_ref, "target_ref": target_ref, "value": value,
                        "conversion": {"kind": "identity", "source_unit": source_unit, "target_unit": target_unit}})
    return changes
