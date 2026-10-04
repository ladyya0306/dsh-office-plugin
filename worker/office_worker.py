"""Offline, stdin/stdout Office worker.  The first supported format is XLSX.

The worker accepts base64 package bytes, never paths.  It has no network,
subprocess, model, shell, or configuration loading code.  ``apply`` writes a
new, exclusive ``.xlsx`` file beneath the current working directory only.
"""
from __future__ import annotations

import base64
import binascii
from datetime import date, datetime, time
import hashlib
import io
import json
import os
from pathlib import Path
import re
import sys
import zipfile
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
if __name__ == "__main__":
    # Structured modules must share protocol exception types with this entry.
    sys.modules.setdefault("worker.office_worker", sys.modules[__name__])

from worker.xlsx_prepare import PrepareError, derive_changes
from worker.xlsx_package_patch import XlsxPackageError, patch_xlsx
from worker.docx_prepare import DocxPrepareError, derive_changes as derive_docx_changes, explicit_unit, paragraph_values, source_candidates as docx_source_candidates, target_slots as docx_target_slots
from worker.package_writer import replace_parts


_OPENPYXL_DEPENDENCIES: tuple[Any, type] | None = None
_DOCX_DOCUMENT: Any = None


def _openpyxl_dependencies() -> tuple[Any, type]:
    """Load the external XLSX runtime only after an XLSX request is accepted."""
    global _OPENPYXL_DEPENDENCIES
    if _OPENPYXL_DEPENDENCIES is None:
        try:
            from openpyxl import load_workbook
            from openpyxl.cell.cell import MergedCell
        except ModuleNotFoundError as exc:
            raise WorkerError(
                "dependency_unavailable",
                "XLSX worker 需要 openpyxl；隔离运行环境必须将该依赖安装在受信任的解释器路径中",
            ) from exc
        _OPENPYXL_DEPENDENCIES = (load_workbook, MergedCell)
    return _OPENPYXL_DEPENDENCIES


def _is_merged_cell(cell: Any) -> bool:
    return isinstance(cell, _openpyxl_dependencies()[1])


def _docx_document_type():
    global _DOCX_DOCUMENT
    if _DOCX_DOCUMENT is None:
        try:
            from docx import Document
        except ModuleNotFoundError as exc:
            raise WorkerError("dependency_unavailable", "DOCX worker 需要 python-docx") from exc
        _DOCX_DOCUMENT = Document
    return _DOCX_DOCUMENT


MAX_INPUT_BYTES = 10 * 1024 * 1024
MAX_REQUEST_BYTES = 28 * 1024 * 1024
MAX_ZIP_MEMBERS = 2048
MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024
MAX_COMPRESSION_RATIO = 100
MAX_CANDIDATES = 1000
MAX_CHANGES = 100
MAX_OUTPUT_BYTES = 15 * 1024 * 1024
MAX_WORKSHEETS = 32
MAX_WORKSHEET_ROWS = 10_000
MAX_WORKSHEET_COLUMNS = 256
MAX_TOTAL_SCANNED_CELLS = 100_000
OUTPUT_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,119}$")
PLACEHOLDER_RE = re.compile(r"^[\s_＿—\-□☐]*$")
CELL_RE = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]{0,6}$")


class WorkerError(ValueError):
    def __init__(self, code: str, message: str, details: dict[str, Any] | None = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details or {}


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _decode_b64(value: Any, field: str) -> bytes:
    if not isinstance(value, str):
        raise WorkerError("invalid_request", f"{field} 必须是 base64 字符串")
    try:
        data = base64.b64decode(value, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise WorkerError("invalid_base64", f"{field} 不是有效 base64") from exc
    if not data:
        raise WorkerError("empty_input", f"{field} 不能为空")
    if len(data) > MAX_INPUT_BYTES:
        raise WorkerError("input_too_large", f"{field} 超过 {MAX_INPUT_BYTES} 字节限制")
    return data


def _check_zip(data: bytes, label: str) -> frozenset[str]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            infos = archive.infolist()
    except zipfile.BadZipFile as exc:
        raise WorkerError("invalid_package", f"{label} 不是有效的 Office ZIP 包") from exc
    if len(infos) > MAX_ZIP_MEMBERS:
        raise WorkerError("zip_too_many_members", f"{label} ZIP 成员数超过 {MAX_ZIP_MEMBERS}")
    names: set[str] = set()
    file_members: set[str] = set()
    for info in infos:
        if info.filename in names:
            raise WorkerError("zip_duplicate_member", f"{label} ZIP 包含重复 member：{info.filename}")
        names.add(info.filename)
        if info.is_dir():
            # Deflate may encode an empty directory entry with a few bytes of
            # compressed framing.  ZipInfo.file_size is the payload boundary.
            if info.file_size != 0:
                raise WorkerError("zip_suspicious", f"{label} ZIP 包含带载荷的伪目录：{info.filename}")
            continue
        file_members.add(info.filename)
    total = sum(info.file_size for info in infos)
    if total > MAX_UNCOMPRESSED_BYTES:
        raise WorkerError("zip_too_large", f"{label} 解压后超过 {MAX_UNCOMPRESSED_BYTES} 字节限制")
    for info in infos:
        if info.compress_size == 0 and info.file_size > 0:
            raise WorkerError("zip_suspicious", f"{label} 包含零压缩大小的非空成员")
        if info.compress_size and info.file_size / info.compress_size > MAX_COMPRESSION_RATIO:
            raise WorkerError("zip_suspicious", f"{label} 包含超过 {MAX_COMPRESSION_RATIO}:1 的压缩成员")
    return frozenset(file_members)


def _zip_file_contents(data: bytes, label: str) -> dict[str, bytes]:
    members = _check_zip(data, label)
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in members}


def _relationship_parts(data: bytes, label: str) -> dict[str, bytes]:
    return {name: value for name, value in _zip_file_contents(data, label).items() if name.endswith(".rels")}


def _check_xlsx_features(data: bytes, label: str) -> frozenset[str]:
    members = _check_zip(data, label)
    unsupported_prefixes = (
        "xl/externalLinks/", "xl/embeddings/", "xl/activeX/", "xl/ctrlProps/",
        "xl/drawings/", "xl/charts/", "xl/slicers/", "xl/threadedComments/", "xl/persons/",
    )
    unsupported = sorted(name for name in members if name == "xl/vbaProject.bin" or name == "xl/calcChain.xml"
                         or name.startswith(unsupported_prefixes))
    if unsupported:
        listed = "、".join(unsupported[:5])
        raise WorkerError("unsupported_xlsx_feature",
                          f"{label} 包含首版不处理的外部链接、宏、嵌入对象或绘图相关部件：{listed}")
    return members


def _check_docx_features(data: bytes, label: str) -> frozenset[str]:
    members = _check_zip(data, label)
    unsupported = [name for name in members if name == "word/vbaProject.bin" or name.startswith(("word/embeddings/", "word/activeX/", "word/afchunk"))]
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            has_textbox = any(b"<w:txbxContent" in archive.read(name) for name in members if name.startswith("word/") and name.endswith(".xml"))
    except KeyError as exc:
        raise WorkerError("invalid_package", f"{label} DOCX ZIP 成员无法读取") from exc
    if unsupported or has_textbox:
        detail = "、".join(sorted(unsupported)[:4]) or "文本框"
        raise WorkerError("unsupported_docx_feature", f"{label} 包含首版不支持的嵌入物、宏或文本框：{detail}")
    if "word/document.xml" not in members:
        raise WorkerError("invalid_package", f"{label} 缺少 word/document.xml")
    return members


def _load_docx(data: bytes, label: str):
    _check_docx_features(data, label)
    try:
        return _docx_document_type()(io.BytesIO(data))
    except WorkerError:
        raise
    except Exception as exc:
        raise WorkerError("invalid_docx", f"{label} 无法按 DOCX 打开：{exc}") from exc


def _load_xlsx(data: bytes, label: str):
    _check_xlsx_features(data, label)
    try:
        workbook = _openpyxl_dependencies()[0](io.BytesIO(data), read_only=False, data_only=False)
    except WorkerError:
        raise
    except Exception as exc:  # openpyxl exception types vary by malformed package.
        raise WorkerError("invalid_xlsx", f"{label} 无法按 XLSX 打开：{exc}") from exc
    try:
        _validate_workbook_shape(workbook, label)
        return workbook
    except Exception:
        workbook.close()
        raise


def _validate_workbook_shape(wb: Any, label: str) -> None:
    if len(wb.worksheets) > MAX_WORKSHEETS:
        raise WorkerError("workbook_too_large", f"{label} 工作表数超过 {MAX_WORKSHEETS}")
    scanned = 0
    for ws in wb.worksheets:
        rows, columns = ws.max_row, ws.max_column
        if rows > MAX_WORKSHEET_ROWS or columns > MAX_WORKSHEET_COLUMNS:
            raise WorkerError("workbook_too_large",
                              f"{label} 的 {ws.title} 尺寸为 {rows}×{columns}，超过 {MAX_WORKSHEET_ROWS}×{MAX_WORKSHEET_COLUMNS}")
        scanned += rows * columns
        if scanned > MAX_TOTAL_SCANNED_CELLS:
            raise WorkerError("workbook_too_large", f"{label} 扫描格数超过 {MAX_TOTAL_SCANNED_CELLS}")


def _text(value: Any) -> str | None:
    if value is None:
        return None
    rendered = str(value).strip()
    return rendered if rendered else None


def _is_formula(cell: Any) -> bool:
    return cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("="))


def _merged_range(ws: Any, coordinate: str) -> str | None:
    for merged in ws.merged_cells.ranges:
        if coordinate in merged:
            return str(merged)
    return None


def _neighbors(ws: Any, cell: Any) -> dict[str, str]:
    values: dict[str, str] = {}
    for name, row, column in (("left", cell.row, cell.column - 1), ("right", cell.row, cell.column + 1),
                              ("up", cell.row - 1, cell.column), ("down", cell.row + 1, cell.column)):
        if row < 1 or column < 1:
            continue
        value = _text(ws.cell(row, column).value)
        if value:
            values[name] = value[:200]
    return values


def _label_and_unit(ws: Any, cell: Any) -> tuple[str | None, str | None, dict[str, str]]:
    nearby = _neighbors(ws, cell)
    label = nearby.get("left") or nearby.get("up")
    unit = nearby.get("right")
    return label, unit, nearby


def _iter_source_candidates(wb: Any) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    candidates: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    for ws in wb.worksheets:
        if "!" in ws.title:
            warnings.append({"sheet": ws.title, "code": "unaddressable_sheet_name",
                             "message": "工作表名称含 !，首版 sheet!A1 引用不能表示；该表不输出候选"})
            continue
        for row in ws.iter_rows():
            for cell in row:
                if _is_merged_cell(cell) or cell.value is None or _is_formula(cell):
                    continue
                if isinstance(cell.value, (datetime, date, time)):
                    raise WorkerError("unsupported_value_type", f"日期/时间来源暂不支持，拒绝隐式 JSON 序列化：{ws.title}!{cell.coordinate}")
                label, unit, nearby = _label_and_unit(ws, cell)
                if not label:
                    continue
                candidates.append({"source_ref": f"{ws.title}!{cell.coordinate}", "label": label,
                                   "value": cell.value, "unit": unit or "", "nearby": nearby})
                if len(candidates) >= MAX_CANDIDATES:
                    warnings.append({"code": "candidate_limit", "message": f"来源候选已截断为 {MAX_CANDIDATES} 项"})
                    return candidates, warnings
    return candidates, warnings


def _slot_status(ws: Any, cell: Any) -> tuple[bool, dict[str, Any]]:
    if _is_merged_cell(cell):
        return False, {"reason": "merged_non_top_left"}
    merged = _merged_range(ws, cell.coordinate)
    if _is_formula(cell):
        return False, {"reason": "formula", "merged_range": merged}
    if ws.protection.sheet and cell.protection.locked:
        return False, {"reason": "protected", "merged_range": merged}
    if cell.value is not None and not (isinstance(cell.value, str) and PLACEHOLDER_RE.fullmatch(cell.value)):
        return False, {"reason": "non_blank", "merged_range": merged}
    label, unit, nearby = _label_and_unit(ws, cell)
    if not label:
        return False, {"reason": "missing_adjacent_label", "merged_range": merged}
    return True, {"label": label, "unit": unit or "", "nearby": nearby, "merged_range": merged}


def _iter_target_slots(wb: Any) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    slots: list[dict[str, Any]] = []
    blocked: list[dict[str, Any]] = []
    for ws in wb.worksheets:
        if "!" in ws.title:
            blocked.append({"sheet": ws.title, "code": "unaddressable_sheet_name",
                            "message": "工作表名称含 !，首版 sheet!A1 引用不能表示"})
            continue
        for row in ws.iter_rows():
            for cell in row:
                if _is_merged_cell(cell):
                    continue
                ok, detail = _slot_status(ws, cell)
                if ok:
                    slots.append({"target_ref": f"{ws.title}!{cell.coordinate}", **detail})
                    if len(slots) >= MAX_CANDIDATES:
                        blocked.append({"code": "candidate_limit", "message": f"目标空位已截断为 {MAX_CANDIDATES} 项"})
                        return slots, blocked
                elif detail["reason"] in {"formula", "protected"}:
                    blocked.append({"target_ref": f"{ws.title}!{cell.coordinate}", **detail})
    return slots, blocked


def _parse_ref(reference: Any, role: str) -> tuple[str, str]:
    if not isinstance(reference, str) or reference.count("!") != 1:
        raise WorkerError("invalid_reference", f"{role} 必须是 sheet!A1，且工作表名称不能含 !")
    sheet, coordinate = reference.split("!", 1)
    if not sheet or not CELL_RE.fullmatch(coordinate):
        raise WorkerError("invalid_reference", f"{role} 必须是 sheet!A1 的单一单元格")
    return sheet, coordinate.upper()


def _cell_for_ref(wb: Any, reference: Any, role: str):
    sheet, coordinate = _parse_ref(reference, role)
    if sheet not in wb.sheetnames:
        raise WorkerError("unknown_sheet", f"{role} 的工作表不存在：{sheet}")
    cell = wb[sheet][coordinate]
    if _is_merged_cell(cell):
        raise WorkerError("merged_non_top_left", f"{role} 指向合并区域的非左上单元格：{reference}")
    return wb[sheet], cell


def _validate_output_name(value: Any, suffix: str = ".xlsx") -> str:
    if not isinstance(value, str) or not OUTPUT_NAME_RE.fullmatch(value) or not value.endswith(suffix) or ".." in value:
        raise WorkerError("invalid_output_name", f"output_name 必须是当前工作目录下的普通 {suffix} 文件名，不能含路径、冒号或 ..")
    return value


def _validate_request(request: Any) -> tuple[str, str, str, bytes, bytes]:
    if not isinstance(request, dict):
        raise WorkerError("invalid_request", "请求必须是 JSON 对象")
    operation = request.get("operation")
    if operation not in {"inspect", "prepare", "apply"}:
        raise WorkerError("invalid_operation", "operation 只能是 inspect、prepare 或 apply")
    format_name = request.get("format")
    source_format = request.get("source_format", format_name)
    if format_name not in {"xlsx", "docx"} or source_format not in {"xlsx", "docx"}:
        raise WorkerError("unsupported_format", "format/source_format 只能是 xlsx 或 docx")
    if source_format != format_name:
        raise WorkerError("unsupported_cross_format", "首版只支持 source_format 与 format 相同")
    return operation, format_name, source_format, _decode_b64(request.get("source_b64"), "source_b64"), _decode_b64(request.get("target_b64"), "target_b64")


def _workbook_summary(source: bytes, target: bytes, source_wb: Any, target_wb: Any) -> dict[str, Any]:
    candidates, source_warnings = _iter_source_candidates(source_wb)
    slots, target_hints = _iter_target_slots(target_wb)
    return {"source_sha256": _sha256(source), "target_sha256": _sha256(target),
            "source_candidates": candidates, "target_slots": slots,
            "hints": {"source": source_warnings, "target": target_hints}}


def _cell_fidelity_state(cell: Any, include_value: bool) -> tuple[Any, ...]:
    state: tuple[Any, ...] = (
        tuple(cell._style) if cell.has_style else None,
        cell.number_format,
        cell.hyperlink.target if cell.hyperlink else None,
        cell.comment.text if cell.comment else None,
    )
    return (cell.value, cell.data_type, *state) if include_value else state


def _dimensions_state(ws: Any) -> tuple[tuple[Any, ...], tuple[Any, ...]]:
    rows = tuple(sorted((index, dim.height, dim.hidden, dim.outlineLevel, dim.collapsed,
                         tuple(dim._style) if dim.has_style else None)
                        for index, dim in ws.row_dimensions.items()
                        if dim.height is not None or dim.hidden or dim.outlineLevel or dim.collapsed or dim.has_style))
    columns = tuple(sorted((index, dim.width, dim.hidden, dim.outlineLevel, dim.collapsed, dim.min, dim.max)
                           for index, dim in ws.column_dimensions.items()))
    return rows, columns


def _fidelity_snapshot(wb: Any, editable: set[tuple[str, str]]) -> dict[str, Any]:
    sheets: dict[str, Any] = {}
    for ws in wb.worksheets:
        states: dict[str, tuple[Any, ...]] = {}
        for cell in ws._cells.values():
            if _is_merged_cell(cell):
                continue
            editable_cell = (ws.title, cell.coordinate) in editable
            if not editable_cell and cell.value is None and not cell.has_style and not cell.hyperlink and not cell.comment:
                # openpyxl may materialize default blank cells during scanning
                # and remove them again on save.  They are not visible content
                # or formatting and cannot be a fidelity regression.
                continue
            states[cell.coordinate] = _cell_fidelity_state(cell, not editable_cell)
        sheets[ws.title] = {
            "merged": str(ws.merged_cells), "dimensions": _dimensions_state(ws),
            "freeze_panes": str(ws.freeze_panes), "sheet_state": ws.sheet_state,
            "print_area": str(ws.print_area), "page_setup": str(ws.page_setup),
            "page_margins": str(ws.page_margins), "print_options": str(ws.print_options), "cells": states,
        }
    return {"sheetnames": tuple(wb.sheetnames), "sheets": sheets}


def _fidelity_issues(before: dict[str, Any], after_wb: Any, editable: set[tuple[str, str]]) -> list[str]:
    after = _fidelity_snapshot(after_wb, editable)
    issues: list[str] = []
    if before["sheetnames"] != after["sheetnames"]:
        return ["工作表名称或顺序改变"]
    for name in before["sheetnames"]:
        left, right = before["sheets"][name], after["sheets"][name]
        for key, label in (("merged", "合并区域"), ("dimensions", "行高或列宽"), ("freeze_panes", "冻结窗格"),
                           ("sheet_state", "隐藏状态"), ("print_area", "打印区域"), ("page_setup", "页面设置"),
                           ("page_margins", "页边距"), ("print_options", "打印选项")):
            if left[key] != right[key]:
                issues.append(f"{name} {label}改变")
        coordinates = set(left["cells"]) | set(right["cells"])
        for coordinate in sorted(coordinates):
            if left["cells"].get(coordinate) != right["cells"].get(coordinate):
                issues.append(f"{name}!{coordinate} 的内容或样式改变")
                if len(issues) >= 20:
                    return issues
    return issues


def _write_exclusive(name: str, data: bytes) -> str:
    if len(data) > MAX_OUTPUT_BYTES:
        raise WorkerError("output_too_large", f"输出超过 {MAX_OUTPUT_BYTES} 字节限制")
    cwd = Path.cwd().resolve()
    destination = (cwd / name).resolve()
    if destination.parent != cwd:
        raise WorkerError("invalid_output_name", "output_name 不得离开当前工作目录")
    try:
        fd = os.open(destination, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    except FileExistsError as exc:
        raise WorkerError("output_exists", f"输出文件已存在：{name}") from exc
    try:
        with os.fdopen(fd, "wb") as output:
            output.write(data)
            output.flush()
            os.fsync(output.fileno())
    except OSError as exc:
        # The name was exclusively reserved.  Preserve the bytes that reached
        # disk so the caller can inspect or remove this recoverable partial
        # output explicitly; never delete a possibly useful failure artifact.
        raise WorkerError("output_partial", f"输出写入或同步失败，已保留可能不完整的文件：{name}", {
            "output_name": name, "output_path": str(destination), "recoverable": True,
        }) from exc
    return str(destination)


def _mapping_source(wb: Any, reference: Any) -> tuple[Any, Any, str]:
    ws, cell = _cell_for_ref(wb, reference, "source_ref")
    if cell.value is None or _is_formula(cell):
        raise WorkerError("invalid_source", f"source_ref 必须指向非公式、非空的来源值：{reference}")
    return ws, cell, f"{ws.title}!{cell.coordinate}"


def _mapping_target(wb: Any, reference: Any) -> tuple[Any, Any, str]:
    ws, cell = _cell_for_ref(wb, reference, "target_ref")
    return ws, cell, f"{ws.title}!{cell.coordinate}"


def _assert_safe_mapping_target(ws: Any, cell: Any, reference: str) -> None:
    allowed, detail = _slot_status(ws, cell)
    if not allowed:
        raise WorkerError("unsafe_target", f"target_ref 不可写：{reference}（{detail['reason']}）")


def _derive_mapping_changes(source_wb: Any, target_wb: Any, mappings: Any) -> list[dict[str, Any]]:
    try:
        return derive_changes(
            mappings,
            lambda reference: _mapping_source(source_wb, reference),
            lambda reference: _mapping_target(target_wb, reference),
            _assert_safe_mapping_target,
        )
    except PrepareError as exc:
        raise WorkerError(exc.code, exc.message) from exc


def _prepare(request: dict[str, Any], source: bytes, target: bytes, source_wb: Any, target_wb: Any) -> dict[str, Any]:
    changes = _derive_mapping_changes(source_wb, target_wb, request.get("mappings"))
    return {"source_sha256": _sha256(source), "target_sha256": _sha256(target), "changes": changes}


_DOCX_P = re.compile(r"^p:(\d+)$")
_DOCX_T = re.compile(r"^t:(\d+):r:(\d+):c:(\d+)$")
_DOCX_PV = re.compile(r"^p:(\d+):v:(\d+)$")


def _docx_ref(document: Any, reference: Any, role: str):
    if not isinstance(reference, str):
        raise WorkerError("invalid_reference", f"{role} 必须是 p:索引 或 t:表:r:行:c:列")
    match = _DOCX_P.fullmatch(reference)
    if match:
        index = int(match.group(1))
        if index >= len(document.paragraphs): raise WorkerError("invalid_reference", f"{role} 段落不存在：{reference}")
        return document.paragraphs[index], f"p:{index}", "p"
    match = _DOCX_T.fullmatch(reference)
    if match:
        ti, ri, ci = map(int, match.groups())
        try: return document.tables[ti].rows[ri].cells[ci], f"t:{ti}:r:{ri}:c:{ci}", "t"
        except IndexError as exc: raise WorkerError("invalid_reference", f"{role} 表格单元格不存在：{reference}") from exc
    raise WorkerError("invalid_reference", f"{role} 必须是 p:索引 或 t:表:r:行:c:列")


def _docx_safe_target(item: Any, reference: str) -> bool:
    text = item.text.strip()
    if reference.startswith("p:"):
        return bool(re.fullmatch(r"[^：:]{1,50}[：:]", text)) and not re.search(r"说明|备注|提示|签字|审批|日期", text)
    return not text


def _docx_source_value(item: Any, reference: str) -> str:
    text = item.text.strip()
    if reference.startswith("t:"):
        return text
    values = paragraph_values(text)
    match = _DOCX_PV.fullmatch(reference)
    if match:
        ordinal = int(match.group(2))
        if ordinal < len(values): return values[ordinal]["value"]
        raise WorkerError("invalid_reference", f"source_ref 片段不存在：{reference}")
    if len(values) == 1: return values[0]["value"]
    raise WorkerError("ambiguous_source", f"段落来源必须只有一个明确标签值；当前不能用 {reference} 猜测多个叙述值")


def _docx_changes(source_doc: Any, target_doc: Any, mappings: Any) -> list[dict[str, Any]]:
    try:
        physical_targets = set()
        def source(reference: Any):
            match = _DOCX_PV.fullmatch(reference) if isinstance(reference, str) else None
            item, ref, _kind = _docx_ref(source_doc, f"p:{match.group(1)}" if match else reference, "source_ref")
            if match: ref = reference
            return _docx_source_value(item, ref), ref
        def target(reference: Any):
            item, ref, kind = _docx_ref(target_doc, reference, "target_ref")
            if not _docx_safe_target(item, ref): raise WorkerError("unsafe_target", f"target_ref 不可写：{ref}")
            physical = item._p if kind == "p" else item._tc
            if physical in physical_targets:
                raise WorkerError("duplicate_target", "多个目标引用指向同一物理段落或合并单元格")
            physical_targets.add(physical)
            if kind == "p":
                unit = explicit_unit(item.text.split("：", 1)[0].split(":", 1)[0])
            else:
                match = _DOCX_T.fullmatch(ref)
                ti, ri, ci = map(int, match.groups())
                row = target_doc.tables[ti].rows[ri]
                right = row.cells[ci + 1].text.strip() if ci + 1 < len(row.cells) else ""
                unit = right if right in {"元", "万元", "亿元", "%"} else explicit_unit(row.cells[ci - 1].text)
            return item, ref, unit
        return derive_docx_changes(mappings, source, target)
    except DocxPrepareError as exc:
        raise WorkerError(exc.code, exc.message) from exc


def _docx_summary(source: bytes, target: bytes, source_doc: Any, target_doc: Any) -> dict[str, Any]:
    return {"source_sha256": _sha256(source), "target_sha256": _sha256(target),
            "source_candidates": docx_source_candidates(source_doc),
            "target_slots": docx_target_slots(target_doc, _docx_safe_target), "hints": {"source": [], "target": []}}


def _docx_run_styles(item: Any) -> tuple[str | None, ...]:
    return tuple(run._r.rPr.xml if run._r.rPr is not None else None for paragraph in item.paragraphs for run in paragraph.runs) if hasattr(item, "paragraphs") else tuple(run._r.rPr.xml if run._r.rPr is not None else None for run in item.runs)


def _docx_fidelity_snapshot(document: Any, editable: set[str]) -> dict[str, tuple[str, tuple[str | None, ...]]]:
    entries = {}
    for index, paragraph in enumerate(document.paragraphs): entries[f"p:{index}"] = (paragraph.text, _docx_run_styles(paragraph))
    for ti, table in enumerate(document.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells): entries[f"t:{ti}:r:{ri}:c:{ci}"] = (cell.text, _docx_run_styles(cell))
    return entries


def _verify_docx_fidelity(before: dict[str, tuple[str, tuple[str | None, ...]]], after: Any, editable: set[str]) -> None:
    now = _docx_fidelity_snapshot(after, editable)
    if set(before) != set(now): raise WorkerError("fidelity_failed", "输出 DOCX 段落或表格结构改变")
    for ref, (text, styles) in before.items():
        new_text, new_styles = now[ref]
        if ref not in editable and (text != new_text or styles != new_styles): raise WorkerError("fidelity_failed", f"非目标 DOCX 内容或 runs 样式改变：{ref}")
        if ref in editable and (not new_text.startswith(text) or new_styles[:len(styles)] != styles): raise WorkerError("fidelity_failed", f"目标 DOCX 原有文字或 runs 样式改变：{ref}")


def _docx_editable_aliases(document: Any, references: set[str]) -> set[str]:
    """Aliases of an authorized merged cell are the same physical write."""
    editable = set(references)
    cells = set()
    for reference in references:
        item, _ref, kind = _docx_ref(document, reference, "target_ref")
        if kind == "t": cells.add(item._tc)
    if cells:
        for ti, table in enumerate(document.tables):
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    if cell._tc in cells: editable.add(f"t:{ti}:r:{ri}:c:{ci}")
    return editable


def _apply_docx(request: dict[str, Any], source: bytes, target: bytes, source_doc: Any, target_doc: Any) -> dict[str, Any]:
    submitted = request.get("changes")
    if not isinstance(submitted, list) or not submitted or len(submitted) > MAX_CHANGES: raise WorkerError("invalid_changes", "changes 必须是非空且受限数组")
    name = _validate_output_name(request.get("output_name"), ".docx")
    mappings = [{"source_ref": x.get("source_ref"), "target_ref": x.get("target_ref")} for x in submitted if isinstance(x, dict)]
    if len(mappings) != len(submitted) or any(set(x) - {"source_ref","target_ref","value","conversion"} or "value" not in x for x in submitted): raise WorkerError("invalid_change", "changes 必须包含 source_ref、target_ref、value")
    changes = _docx_changes(source_doc, target_doc, mappings)
    for expected, given in zip(changes, submitted, strict=True):
        if not isinstance(given["value"], str) or given["value"] != expected["value"]: raise WorkerError("value_mismatch", f"changes.value 不匹配：{expected['target_ref']}")
    editable = _docx_editable_aliases(target_doc, {change["target_ref"] for change in changes})
    fidelity_before = _docx_fidelity_snapshot(target_doc, editable)
    for change in changes:
        item, ref, kind = _docx_ref(target_doc, change["target_ref"], "target_ref")
        paragraph = item if kind == "p" else item.paragraphs[0]
        paragraph.add_run(change["value"])
    if "word/document.xml" not in _zip_file_contents(target, "target"):
        raise WorkerError("invalid_package", "目标 DOCX 缺少 word/document.xml")
    output = replace_parts(target, {"word/document.xml": target_doc.part.blob})
    target_parts = _zip_file_contents(target, "target")
    output_parts = _zip_file_contents(output, "输出")
    if set(target_parts) != set(output_parts): raise WorkerError("package_parts_changed", "输出 DOCX ZIP 文件部件集合改变，拒绝写出")
    if _relationship_parts(target, "target") != _relationship_parts(output, "输出"):
        raise WorkerError("fidelity_failed", "DOCX 关系部件字节改变")
    changed_parts = [name for name in target_parts if name != "word/document.xml" and target_parts[name] != output_parts[name]]
    if changed_parts:
        raise WorkerError("fidelity_failed", f"DOCX 未编辑文件部件字节改变：{changed_parts[0]}")
    reopened = _load_docx(output, "输出")
    for change in changes:
        item, _ref, _kind = _docx_ref(reopened, change["target_ref"], "target_ref")
        if change["value"] not in item.text: raise WorkerError("verification_failed", f"输出重开后未找到写入：{change['target_ref']}")
    _verify_docx_fidelity(fidelity_before, reopened, editable)
    destination = _write_exclusive(name, output)
    return {"applied": changes, "output_name": name, "output_path": destination, "output_sha256": _sha256(output), "source_sha256": _sha256(source), "target_sha256": _sha256(target), "verification": {"reopened": True, "package_members_preserved": True, "fidelity": {"ok": True, "scope": "验证非目标文本/runs 样式、目标原有 runs 样式前缀和 ZIP 成员集合；未声称所有 DOCX 结构完全保真"}}}


def _apply(request: dict[str, Any], source: bytes, target: bytes, source_wb: Any, target_wb: Any) -> dict[str, Any]:
    submitted_changes = request.get("changes")
    if not isinstance(submitted_changes, list) or not submitted_changes or len(submitted_changes) > MAX_CHANGES:
        raise WorkerError("invalid_changes", f"changes 必须是 1 到 {MAX_CHANGES} 项的数组")
    name = _validate_output_name(request.get("output_name"))
    mappings: list[dict[str, Any]] = []
    for submitted in submitted_changes:
        if not isinstance(submitted, dict) or not {"source_ref", "target_ref", "value"} <= submitted.keys():
            raise WorkerError("invalid_change", "每项 changes 必须包含 source_ref、target_ref、value")
        if set(submitted) - {"source_ref", "target_ref", "value", "conversion"}:
            raise WorkerError("invalid_change", "changes 不支持未声明字段")
        mappings.append({"source_ref": submitted["source_ref"], "target_ref": submitted["target_ref"]})
    changes = _derive_mapping_changes(source_wb, target_wb, mappings)
    for expected, submitted in zip(changes, submitted_changes, strict=True):
        if type(submitted["value"]) is not type(expected["value"]) or submitted["value"] != expected["value"]:
            raise WorkerError("value_mismatch", f"changes.value 与来源和单位确定性推导结果不一致：{expected['target_ref']}")
        if "conversion" in submitted and submitted["conversion"] != expected["conversion"]:
            raise WorkerError("conversion_mismatch", f"changes.conversion 与确定性推导结果不一致：{expected['target_ref']}")
    planned_targets: set[tuple[str, str]] = set()
    for change in changes:
        if isinstance(change, dict) and "target_ref" in change:
            sheet, coordinate = _parse_ref(change["target_ref"], "target_ref")
            planned_targets.add((sheet, coordinate))
    fidelity_before = _fidelity_snapshot(target_wb, planned_targets)
    seen_targets: set[str] = set()
    applied: list[dict[str, Any]] = []
    for change in changes:
        source_ws, source_cell = _cell_for_ref(source_wb, change["source_ref"], "source_ref")
        target_ws, target_cell = _cell_for_ref(target_wb, change["target_ref"], "target_ref")
        target_key = f"{target_ws.title}!{target_cell.coordinate}"
        if target_key in seen_targets:
            raise WorkerError("duplicate_target", f"target_ref 重复：{target_key}")
        seen_targets.add(target_key)
        target_cell.value = change["value"]
        applied.append({"source_ref": f"{source_ws.title}!{source_cell.coordinate}", "target_ref": target_key,
                        "value": change["value"]})
    try:
        output = patch_xlsx(target, applied)
    except XlsxPackageError as exc:
        raise WorkerError(exc.code, exc.message) from exc
    if len(output) > MAX_OUTPUT_BYTES:
        raise WorkerError("output_too_large", f"输出超过 {MAX_OUTPUT_BYTES} 字节限制")
    target_members = _check_xlsx_features(target, "target")
    output_members = _check_xlsx_features(output, "输出")
    if target_members != output_members:
        removed = sorted(target_members - output_members)
        added = sorted(output_members - target_members)
        detail = "；".join(filter(None, [
            f"缺少：{'、'.join(removed[:4])}" if removed else "",
            f"新增：{'、'.join(added[:4])}" if added else "",
        ]))
        raise WorkerError("package_parts_changed", f"输出 ZIP 成员集合改变，拒绝写出：{detail}")
    reopened = _load_xlsx(output, "输出")
    try:
        for change in applied:
            sheet, coord = _parse_ref(change["target_ref"], "target_ref")
            if reopened[sheet][coord].value != change["value"]:
                raise WorkerError("verification_failed", f"输出重开后未找到预期写入：{change['target_ref']}")
        fidelity_issues = _fidelity_issues(fidelity_before, reopened, planned_targets)
        if fidelity_issues:
            raise WorkerError("fidelity_failed", "输出重开后出现未授权的工作簿改变：" + "；".join(fidelity_issues[:8]))
    finally:
        reopened.close()
    destination = _write_exclusive(name, output)
    return {"applied": applied, "output_name": name, "output_path": destination,
            "output_sha256": _sha256(output), "source_sha256": _sha256(source),
            "target_sha256": _sha256(target), "verification": {
                "reopened": True,
                "written_cells": len(applied),
                "package_members_preserved": True,
                "fidelity": {
                    "ok": True,
                    "scope": "除明确 target_ref 的值外，工作表顺序、合并区域、页面属性、行列尺寸和有意义单元格的内容/样式未变",
                },
            }}


def _run_request(request: Any) -> dict[str, Any]:
    if isinstance(request,dict) and request.get('operation')=='bind_group_composition':
        from worker.group_composition_binding import bind_group_composition
        return bind_group_composition(request)
    if isinstance(request,dict) and request.get('operation')=='bind_group_relationship':
        from worker.group_relationship_binding import bind_group_relationship
        return bind_group_relationship(request)
    if isinstance(request,dict) and request.get('operation')=='bind_relationship':
        from worker.relationship_binding import bind_relationship
        return bind_relationship(request)
    if isinstance(request,dict) and request.get('operation')=='restore_template':
        from worker.template_restore import restore_template
        return restore_template(request)
    if isinstance(request, dict) and request.get("operation") == "check_output":
        from worker.output_check import check_output
        return check_output(request)
    if isinstance(request, dict) and request.get("operation") == "check_group_output":
        from worker.group_output_check import check_group
        return check_group(request)
    if isinstance(request, dict) and request.get("operation") in {"records_store", "records_view"}:
        from worker.provenance_store import run_records
        return run_records(request)
    if isinstance(request, dict) and request.get("operation") in {"inspect_v3", "prepare_v3", "apply_v3", "fill_v3"}:
        from worker.structured_worker import run_v3
        return run_v3(request)
    operation, format_name, _source_format, source, target = _validate_request(request)
    if format_name == "docx":
        source_doc, target_doc = _load_docx(source, "source"), _load_docx(target, "target")
        if operation == "inspect": return {"ok": True, "operation": "inspect", **_docx_summary(source, target, source_doc, target_doc)}
        if operation == "prepare": return {"ok": True, "operation": "prepare", "source_sha256": _sha256(source), "target_sha256": _sha256(target), "changes": _docx_changes(source_doc, target_doc, request.get("mappings"))}
        return {"ok": True, "operation": "apply", **_apply_docx(request, source, target, source_doc, target_doc)}
    source_wb = _load_xlsx(source, "source")
    target_wb = _load_xlsx(target, "target")
    try:
        summary = _workbook_summary(source, target, source_wb, target_wb)
        if operation == "inspect":
            return {"ok": True, "operation": "inspect", **summary}
        if operation == "prepare":
            return {"ok": True, "operation": "prepare", **_prepare(request, source, target, source_wb, target_wb)}
        return {"ok": True, "operation": "apply", **_apply(request, source, target, source_wb, target_wb)}
    finally:
        source_wb.close()
        target_wb.close()


def run_request(request: Any) -> dict[str, Any]:
    """Run one request and normalize all worker-domain failures to protocol JSON."""
    try:
        return _run_request(request)
    except WorkerError as exc:
        return {"ok": False, "error": {"code": exc.code, "message": exc.message, **exc.details}}
    except Exception as exc:  # The protocol never emits Python tracebacks.
        return {"ok": False, "error": {"code": "internal_error", "message": str(exc)}}


def _read_stdin_request() -> Any:
    stream = getattr(sys.stdin, "buffer", sys.stdin)
    raw = stream.read(MAX_REQUEST_BYTES + 1)
    raw_bytes = raw.encode("utf-8") if isinstance(raw, str) else raw
    if len(raw_bytes) > MAX_REQUEST_BYTES:
        raise WorkerError("request_too_large", f"stdin 请求超过 {MAX_REQUEST_BYTES} 字节限制")
    return json.loads(raw_bytes)


def main() -> int:
    try:
        request = _read_stdin_request()
        response = run_request(request)
    except WorkerError as exc:
        response = {"ok": False, "error": {"code": exc.code, "message": exc.message, **exc.details}}
    except json.JSONDecodeError:
        response = {"ok": False, "error": {"code": "invalid_json", "message": "stdin 不是有效 JSON"}}
    except Exception as exc:  # Do not expose tracebacks or filesystem internals through the protocol.
        response = {"ok": False, "error": {"code": "internal_error", "message": str(exc)}}
    sys.stdout.write(json.dumps(response, ensure_ascii=False, separators=(",", ":")) + "\n")
    return 0 if response.get("ok") else 2


if __name__ == "__main__":
    raise SystemExit(main())
