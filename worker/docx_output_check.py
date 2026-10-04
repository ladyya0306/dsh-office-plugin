"""Independent DOCX output and template-fidelity checks.

This module deliberately works from the physical OOXML objects in the target
read.  It does not use source values or business labels to decide whether a
change is correct.  Empty/review cells and direct paragraph spans are the only
write shapes understood here.
"""
import io
import re
import zipfile

from lxml import etree

from worker.structure_docx import NS, TAG, own_text, read_docx


_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_DOC = "word/document.xml"
_CELL = re.compile(r"^word/document\.xml\|c:(\d+)$")
_SPAN = re.compile(r"^word/document\.xml\|p:(\d+)\|s:(\d+):(\d+)$")


def _parse_xml(data):
    parser = etree.XMLParser(resolve_entities=False, no_network=True,
                              remove_comments=False, recover=False)
    return etree.fromstring(data, parser)


def _zip_parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError("DOCX包含重复部件名")
        return {name: archive.read(name) for name in names}, names


def _tree(value):
    if isinstance(value, dict):
        trees = value.get("trees")
        if isinstance(trees, dict) and _DOC in trees:
            return trees[_DOC]
        data = value.get("bytes")
        if isinstance(data, (bytes, bytearray)):
            parts, _ = _zip_parts(bytes(data))
            return _parse_xml(parts[_DOC])
    if isinstance(value, (bytes, bytearray)):
        parts, _ = _zip_parts(bytes(value))
        return _parse_xml(parts[_DOC])
    raise ValueError("缺少word/document.xml结构")


def _direct_text(p):
    return "".join(p.xpath("./w:r/w:t/text()", namespaces=NS))


def _physical(tree, kind, index):
    tag = "w:tc" if kind == "cell" else "w:p"
    nodes = tree.xpath("//" + tag, namespaces=NS)
    return nodes[index] if 0 <= index < len(nodes) else None


def _slot_kind(ref):
    match = _CELL.fullmatch(ref)
    if match:
        return "cell", int(match.group(1))
    match = _SPAN.fullmatch(ref)
    if match:
        return "span", int(match.group(1)), int(match.group(2)), int(match.group(3))
    return None


def _span_values(original, output, spans):
    """Return values inserted at spans, rejecting ambiguous separators."""
    spans = sorted((int(a), int(b), ref) for a, b, ref in spans)
    for i, (start, end, _) in enumerate(spans):
        if start < 0 or end < start or end > len(original):
            raise ValueError("段落范围超出原模板")
        if i and start < spans[i - 1][1]:
            raise ValueError("段落范围重叠")
    if not spans:
        return {}
    if len(spans)>32:
        raise ValueError('同段超过32个占位范围，请由DSH拆分或处理')
    prefix = original[:spans[0][0]]
    suffix = original[spans[-1][1]:]
    if not output.startswith(prefix) or not output.endswith(suffix):
        raise ValueError("段落固定前后文被改变")
    end_of_values = len(output) - len(suffix)
    if end_of_values < len(prefix):
        raise ValueError("段落输出长度不合法")

    # One span has fixed boundaries and therefore one possible value.
    if len(spans) == 1:
        return {spans[0][2]: output[len(prefix):end_of_values]}

    separators = [original[spans[i][1]:spans[i + 1][0]]
                  for i in range(len(spans) - 1)]
    if any(separator == "" for separator in separators):
        raise ValueError("同段多个占位范围之间没有固定分隔，无法唯一解析")

    results = []

    def walk(position, separator_index, values):
        if len(results)>1:return
        if separator_index == len(separators):
            if position <= end_of_values:
                final = dict(values)
                final[spans[-1][2]] = output[position:end_of_values]
                results.append(final)
            return
        separator = separators[separator_index]
        search = position
        while True:
            found = output.find(separator, search, end_of_values + len(separator))
            if found < 0 or found > end_of_values:
                break
            # The separator must be wholly before the final fixed suffix.
            if found + len(separator) <= end_of_values:
                next_values = dict(values)
                next_values[spans[separator_index][2]] = output[position:found]
                walk(found + len(separator), separator_index + 1, next_values)
            search = found + 1
            if len(results)>1:return

    walk(len(prefix), 0, {})
    # Repeated identical values at different separator positions are still
    # ambiguous; retain every parse rather than deduplicating by value.
    if len(results) != 1:
        raise ValueError("同段多个占位范围的固定分隔解析不唯一")
    return results[0]


def _slot_records(original_target):
    slots = original_target.get("slots") if isinstance(original_target, dict) else None
    if not isinstance(slots, list):
        raise ValueError("目标读取结果缺少slots")
    return {slot.get("target_ref"): slot for slot in slots
            if isinstance(slot, dict) and isinstance(slot.get("target_ref"), str)}


def _extract(original_target, artifact, coverage=None):
    """Read values using original physical slots; return values and diagnostics."""
    original_tree = _tree(original_target)
    artifact_tree = _tree(artifact)
    slots = _slot_records(original_target)
    requested = set(slots) if coverage is None else set(coverage)
    errors = []
    unknown = sorted(requested - set(slots))
    if unknown:
        errors.append("覆盖位置不是本次目标slot：" + ",".join(unknown[:8]))

    by_paragraph = {}
    physical = {}
    for ref, slot in slots.items():
        parsed = _slot_kind(ref)
        if not parsed:
            errors.append("不支持的目标引用：" + ref)
            continue
        kind, index, *span = parsed
        before = _physical(original_tree, kind, index)
        after = _physical(artifact_tree, kind, index)
        if before is None or after is None:
            errors.append("输出缺少目标物理对象：" + ref)
            continue
        if kind == "cell":
            text = own_text(after)
            physical[ref] = {"kind": "cell", "text": text,
                             "original_text": own_text(before)}
        else:
            start, end = span
            original_text = _direct_text(before)
            by_paragraph.setdefault(index, (before, after, original_text, []))[3].append(
                (start, end, ref))

    for index, (before, after, original_text, spans) in by_paragraph.items():
        try:
            values = _span_values(original_text, _direct_text(after), spans)
            for start, end, ref in spans:
                physical[ref] = {"kind": "span", "text": values[ref],
                                 "original_text": original_text[start:end],
                                 "paragraph_text": _direct_text(after)}
        except ValueError as exc:
            errors.append("段落p:%d：%s" % (index, str(exc)))

    return original_tree, artifact_tree, slots, physical, errors


def docx_output_nodes(original_target: dict, artifact: dict,
                      coverage: set[str]) -> dict:
    """Return ``target_ref -> node`` values from the reopened artifact.

    The original target supplies the physical indices and span boundaries.
    ``coverage`` is accepted for the common caller contract; all original
    slots are read so a later contract check can report an omitted requirement.
    This function does not infer values from labels or sources.
    """
    try:
        _, _, slots, physical, _ = _extract(original_target, artifact, coverage)
    except (ValueError, KeyError, TypeError, zipfile.BadZipFile, etree.XMLSyntaxError):
        return {}
    result = {}
    for ref, slot in slots.items():
        value = physical.get(ref, {}).get("text")
        node = dict(slot)
        node["ref"] = ref
        node["text"] = value
        node["value"] = value
        node["output_text"] = value
        result[ref] = node
    return result


def _allowed_paths(original_tree, slots, coverage):
    cells, paragraphs = set(), set()
    errors = []
    for ref in coverage:
        slot = slots.get(ref)
        if slot is None:
            errors.append("覆盖位置不是本次目标slot：" + str(ref))
            continue
        parsed = _slot_kind(ref)
        if not parsed:
            errors.append("不支持的目标引用：" + str(ref))
            continue
        kind, index, *_ = parsed
        element = _physical(original_tree, kind, index)
        if element is None:
            errors.append("模板缺少目标物理对象：" + ref)
            continue
        path = original_tree.getroottree().getpath(element)
        (cells if kind == "cell" else paragraphs).add(path)
    return cells, paragraphs, errors


def _tag_name(node):
    return node.tag if isinstance(node.tag, str) else None


def _attrs(node, allow_text):
    result = dict(node.attrib)
    if allow_text and _tag_name(node) == TAG("t"):
        result.pop(_XML_SPACE, None)
    return result


def _ordinary_new_run(node):
    return (_tag_name(node) == TAG("r") and not node.attrib and
            len(node) == 1 and _tag_name(node[0]) == TAG("t") and
            set(node[0].attrib) <= {_XML_SPACE})


def _children_match(left, right, allow, path, errors, cell_paths=None,
                    paragraph_paths=None, tree=None):
    """Compare children, allowing only documented empty-run insertions."""
    if _tag_name(left) == TAG("r") and allow and not any(_tag_name(x) == TAG("t") for x in left):
        # A writer may add w:t to an existing empty run, but cannot alter any
        # other child or add a second text node.
        if len(right) == len(left) + 1:
            extra = [x for x in right if _tag_name(x) == TAG("t")]
            if len(extra) == 1:
                reduced = list(right)
                reduced.remove(extra[0])
                if all(_tag_name(x) != TAG("t") for x in left):
                    return _compare_child_lists(left, reduced, allow, path, errors,
                                                cell_paths, paragraph_paths, tree)
        return _compare_child_lists(left, right, allow, path, errors,
                                    cell_paths, paragraph_paths, tree)
    if _tag_name(left) == TAG("p") and allow and not any(_tag_name(x) == TAG("r") for x in left):
        if len(right) == len(left) + 1:
            extra = [x for x in right if _ordinary_new_run(x)]
            if len(extra) == 1:
                reduced = list(right)
                reduced.remove(extra[0])
                return _compare_child_lists(left, reduced, allow, path, errors,
                                            cell_paths, paragraph_paths, tree)
    return _compare_child_lists(left, right, allow, path, errors,
                                cell_paths, paragraph_paths, tree)


def _compare_child_lists(left, right, allow, path, errors, cell_paths=None,
                         paragraph_paths=None, tree=None):
    if len(left) != len(right):
        errors.append("XML子节点结构改变：" + path)
        return
    for index, (a, b) in enumerate(zip(left, right)):
        _compare_node(a, b, allow, path + "/" + str(index), errors,
                      cell_paths, paragraph_paths, tree)


def _compare_node(left, right, allow, path, errors, cell_paths=None, paragraph_paths=None, tree=None):
    if _tag_name(left) != _tag_name(right):
        errors.append("XML节点改变：" + path)
        return
    tag = _tag_name(left)
    current_path = tree.getroottree().getpath(left) if tree is not None and isinstance(tag, str) else None
    if paragraph_paths and current_path in paragraph_paths:
        allow = True
    if cell_paths and current_path in cell_paths:
        allow = True
    text_allowed = allow and tag == TAG("t")
    if _attrs(left, text_allowed) != _attrs(right, text_allowed):
        errors.append("XML属性或样式改变：" + path)
    if not text_allowed and (left.text or "") != (right.text or ""):
        errors.append("未覆盖文字改变：" + path)
    if (left.tail or "") != (right.tail or ""):
        errors.append("XML尾文字改变：" + path)
    _children_match(left, right, allow, path, errors,
                    cell_paths, paragraph_paths, tree)


def _uncovered_span_errors(original_tree, artifact_tree, slots, coverage):
    errors = []
    groups = {}
    for ref, slot in slots.items():
        parsed = _slot_kind(ref)
        if parsed and parsed[0] == "span":
            groups.setdefault(parsed[1], []).append((parsed[2], parsed[3], ref))
    for index, spans in groups.items():
        before = _physical(original_tree, "span", index)
        after = _physical(artifact_tree, "span", index)
        if before is None or after is None:
            continue
        try:
            values = _span_values(_direct_text(before), _direct_text(after), spans)
        except ValueError:
            continue
        for start, end, ref in spans:
            if ref not in coverage and values.get(ref) != _direct_text(before)[start:end]:
                errors.append("未覆盖段落范围被改变：" + ref)
    return errors


def check_docx_format(template: bytes, artifact: bytes,
                      coverage: set[str]) -> list[str]:
    """Compare a DOCX artifact to its template within physical slot coverage."""
    errors = []
    try:
        before_parts, before_names = _zip_parts(template)
        after_parts, after_names = _zip_parts(artifact)
    except (ValueError, zipfile.BadZipFile, KeyError, OSError) as exc:
        return ["DOCX包无法读取：" + str(exc)]
    if set(before_names) != set(after_names):
        errors.append("DOCX包部件集合改变")
    for name in sorted(set(before_names) & set(after_names)):
        if name != _DOC and before_parts[name] != after_parts[name]:
            errors.append("未授权部件改变：" + name)
    if _DOC not in before_parts or _DOC not in after_parts:
        return errors + ["缺少word/document.xml"]
    try:
        from worker.office_worker import _check_zip
        original = read_docx(template, "target", target_mode="review_values")
        artifact_read = read_docx(artifact, "target", target_mode="review_values")
        before_tree, after_tree, slots, _, extract_errors = _extract(original, artifact_read, set(coverage))
        errors.extend(extract_errors)
        cells, paragraphs, path_errors = _allowed_paths(before_tree, slots, set(coverage))
        errors.extend(path_errors)
        errors.extend(_uncovered_span_errors(before_tree, after_tree, slots, set(coverage)))
        _compare_node(before_tree, after_tree, False, "/document", errors,
                      cells, paragraphs, before_tree)
    except (ValueError, KeyError, TypeError, zipfile.BadZipFile,
            etree.XMLSyntaxError, OSError) as exc:
        errors.append("DOCX结构无法独立复核：" + str(exc))
    # Preserve ordering while avoiding repeated diagnostics from recursive XML
    # comparison and slot validation.
    return list(dict.fromkeys(errors))

