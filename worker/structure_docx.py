"""OOXML structure records, including nested tables and source textboxes."""
import re
from lxml import etree
from worker.docx_prepare import paragraph_values, explicit_unit, unit_of
from worker.xlsx_package_patch import _xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
TAG = lambda name: "{%s}%s" % (W, name)
BLANK = re.compile(r"[_＿□☐]{2,}")
UNITS = {"元", "千元", "万元", "亿元", "%"}


def own_text(element):
    # Exclude nested table/box paragraphs from their enclosing physical cell.
    paragraphs = [element] if element.tag == TAG("p") else list(element.findall(TAG("p")))
    return "\n".join("".join(p.xpath(".//w:t/text()", namespaces=NS))
                     for p in paragraphs)


def _safe_paragraph(p):
    if p.xpath("ancestor::w:sdt",namespaces=NS):
        return False
    # Bookmarks carry no text and remain unchanged during span replacement.
    # Template fidelity independently checks their positions and attributes.
    allowed = {TAG(x) for x in ("p", "pPr", "r", "rPr", "t", "bookmarkStart", "bookmarkEnd")}
    # Formatting descendants under pPr/rPr are untouched and allowed.
    for node in p.iter():
        if node.tag not in allowed and not any(a.tag in {TAG("pPr"), TAG("rPr")} for a in node.iterancestors()):
            return False
    return True


def _paragraph_spans(p, text):
    spans = [(m.start(), m.end()) for m in BLANK.finditer(text)]
    if not spans and re.fullmatch(r"\s*[^：:]{1,100}[：:]\s*", text):
        spans = [(len(text), len(text))]
    return spans


def _context(element):
    chain = []
    for ancestor in reversed(list(element.iterancestors())):
        if ancestor.tag == TAG("tc"):
            value = own_text(ancestor).strip()
            if value:
                chain.append(value[:300])
    return chain


def read_docx(data, role, target_mode='fill_blanks'):
    import io, zipfile
    from worker.office_worker import _check_zip, WorkerError
    members = _check_zip(data, role)
    blocked = sorted(n for n in members if n == "word/vbaProject.bin" or
                     n.startswith(("word/embeddings/", "word/activeX/", "word/afchunk")))
    if "word/document.xml" not in members:
        raise WorkerError("invalid_package", "缺少word/document.xml")
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {"word/document.xml": _xml(archive.read("word/document.xml"))}
        if role == "source":
            for name in sorted(members):
                if re.fullmatch(r"word/(?:header|footer)\d+\.xml", name):
                    parts[name] = _xml(archive.read(name))
    records, slots, elements = [], [], {}
    text_cache={}
    def text_of(element):
        if element not in text_cache:text_cache[element]=own_text(element)
        return text_cache[element]
    writable_cells = set()
    for part, tree in parts.items():
        paragraphs = tree.xpath("//w:p", namespaces=NS)
        cells = tree.xpath("//w:tc", namespaces=NS)
        if len(paragraphs) + len(cells) > 2000:
            raise WorkerError("structure_limit", "Word结构超过2000个物理对象，未截断；请使用原生范围读取")
        sections, active = {}, []
        for element in tree.iter():
            if element.tag == TAG("p") and element.getparent().tag == TAG("body"):
                style = element.xpath("string(w:pPr/w:pStyle/@w:val)",namespaces=NS)
                heading = re.fullmatch(r"(?:Heading|标题)([1-9])",style)
                if heading:
                    level=int(heading[1]); active=[h for h in active if h[0]<level]
                    active.append((level,text_of(element)[:300]))
            if element.tag in {TAG("tc"),TAG("p")}:
                sections[element]=[text for _,text in active]
        indices = {p: i for i, p in enumerate(paragraphs)}
        for i, cell in enumerate(cells):
            ref = f"{part}|c:{i}"
            properties = cell.find(TAG("tcPr"))
            merge = properties.find(TAG("vMerge")) if properties is not None else None
            continuation = merge is not None and merge.get(TAG("val"), "continue") != "restart"
            row = cell.getparent()
            siblings = row.findall(TAG("tc"))
            column = siblings.index(cell)
            table = row.getparent()
            rows = table.findall(TAG("tr"))
            row_index = rows.index(row)
            grid_col = sum(int(c.xpath("string(w:tcPr/w:gridSpan/@w:val)", namespaces=NS) or 1) for c in siblings[:column])
            column_headers = []
            for above in rows[:row_index]:
                offset = 0
                for c in above.findall(TAG("tc")):
                    span = int(c.xpath("string(w:tcPr/w:gridSpan/@w:val)", namespaces=NS) or 1)
                    if offset <= grid_col < offset + span and text_of(c).strip():
                        column_headers.append(text_of(c).strip()[:300])
                    offset += span
            left = [text_of(c).strip()[:300] for c in siblings[:column] if text_of(c).strip()]
            right = [text_of(c).strip() for c in siblings[column+1:] if text_of(c).strip()]
            text = text_of(cell)
            preceding = table.xpath("preceding-sibling::w:p[position() <= 3]",namespaces=NS)
            context = list(dict.fromkeys(sections.get(cell,[]) + _context(cell) + [text_of(p)[:300] for p in preceding if text_of(p).strip()]))
            unit = next((v for v in right[:1] if v in UNITS), "") or _unit_from_headers(left + column_headers)
            record = {"ref": ref, "kind": "cell", "text": text, "value": text,
                      "parent_context": context, "row_headers": left, "column_headers": column_headers,
                      "unit": unit, "unit_evidence": right[:1] + left + column_headers,
                      "physical_path": tree.getroottree().getpath(cell),
                      "merged_continuation": continuation, "relation_status": "unconfirmed"}
            elements[ref] = cell
            if not continuation:
                records.append(record)
            direct_p = cell.findall(TAG("p"))
            blank = bool(re.fullmatch(r"[\s_＿□☐—-]*",text))
            writable = (role=='target' and part == "word/document.xml" and not continuation and not blocked and
                        (blank or target_mode=='review_values') and not cell.findall(TAG("tbl")) and
                        len(direct_p) == 1 and _safe_paragraph(direct_p[0]) and
                        not cell.xpath("ancestor::w:txbxContent", namespaces=NS) and
                        not tree.xpath("//w:documentProtection", namespaces=NS))
            if writable:
                writable_cells.add(cell)
                slots.append({**record, "target_ref": ref, "write_kind": "empty_cell" if blank else 'replace_cell',
                              'requires_expected_text':not blank,
                              "span_start":0,"span_end":len(text)})
        for i, p in enumerate(paragraphs):
            # Descendant drawing/textbox text is a separate physical paragraph.
            text = "".join(p.xpath("./w:r/w:t/text()", namespaces=NS))
            ref = f"{part}|p:{i}"
            elements[ref] = p
            textbox = bool(p.xpath("ancestor::w:txbxContent", namespaces=NS))
            preceding = p.xpath("preceding::w:p[position() <= 3]", namespaces=NS)
            context = list(dict.fromkeys(sections.get(p,[]) + _context(p) + [text_of(x)[:200] for x in preceding if text_of(x).strip()]))
            base = {"ref": ref, "kind": "textbox" if textbox else "paragraph",
                    "text": text, "value": text, "parent_context": context,
                    "row_headers": [], "column_headers": [], "unit": "",
                    "physical_path": tree.getroottree().getpath(p), "relation_status": "unconfirmed"}
            # Raw paragraphs remain available even without recognized label/value forms.
            if not p.xpath("ancestor::w:tc", namespaces=NS) or textbox:
                records.append(base)
            for j, item in enumerate(paragraph_values(text)):
                fragment = f"{ref}|v:{j}"
                elements[fragment] = p
                records.append({**base, "ref": fragment, **item,
                                "unit": unit_of(item["value"]) or explicit_unit(item["label"]),
                                "unit_evidence": [item["label"], item["value"]]})
            hidden_merge = any(tc.xpath("w:tcPr/w:vMerge[not(@w:val) or @w:val='continue']",namespaces=NS)
                               for tc in p.xpath("ancestor::w:tc",namespaces=NS))
            full_cell_slot = any(tc in writable_cells for tc in p.xpath("ancestor::w:tc",namespaces=NS))
            if part == "word/document.xml" and not textbox and not blocked and not hidden_merge and not full_cell_slot and _safe_paragraph(p):
                for start, end in _paragraph_spans(p, text):
                    slotref = f"{ref}|s:{start}:{end}"
                    elements[slotref] = p
                    fixed = text[max(0,start-100):start] + text[end:end+100]
                    slots.append({**base, "target_ref": slotref, "write_kind": "text_span",
                                  "span_start": start, "span_end": end,
                                  "unit": explicit_unit(fixed) or
                                  next((u for u in ("亿元","万元","千元","元","%") if text[end:].lstrip().startswith(u)), "")})
    # Word protection is stored in settings.xml, not document.xml.
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        protection = (_xml(archive.read("word/settings.xml")).xpath("//w:documentProtection", namespaces=NS)
                      if "word/settings.xml" in members else [])
        unknown = [p for p in protection if p.get(TAG('enforcement'),'').strip().lower()
                   not in {'0','false','off','1','true','on'}]
        if unknown and role=='target':
            raise WorkerError('invalid_document_protection','Word保护开关缺失或无效，需明确其含义后再写入')
        # An explicit off flag is retained in the package and imposes no restriction.
        # Missing/unknown values remain blocked until their meaning can be established.
        protected = any(p.get(TAG('enforcement'), '').strip().lower() not in {'0','false','off'}
                        for p in protection)
    if protected:
        slots = []
    capabilities = {"read": True, "write": not blocked and not protected and role=='target',
                    'target_mode':target_mode,
                    "source_textbox_read": True, "target_textbox_write": False,
                    "unsupported_parts": blocked, "protected": protected,
                    "source_textbox_types":["w:txbxContent/w:p"],"target_content_control_write":False}
    return {"nodes": records, "slots": slots, "capabilities": capabilities,
            "trees": parts, "elements": elements}


def _unit_from_headers(headers):
    units = {explicit_unit(v) for v in headers if explicit_unit(v)}
    return next(iter(units)) if len(units) == 1 else ""


def replace_span(p, start, end, value):
    texts = p.xpath("./w:r/w:t", namespaces=NS)
    full = "".join(t.text or "" for t in texts)
    if not 0 <= start <= end <= len(full):
        raise ValueError("invalid span")
    if not texts:
        runs = p.findall(TAG("r"))
        r = runs[0] if runs else etree.SubElement(p, TAG("r"))
        t = etree.SubElement(r, TAG("t"))
        t.text = value
        return
    offset, inserted = 0, False
    for t in texts:
        original = t.text or ""
        a, b = offset, offset + len(original)
        offset = b
        if start == end:
            if not inserted and a <= start <= b:
                at = start-a
                t.text = original[:at] + value + original[at:]
                inserted = True
        elif b > start and a < end:
            left = original[:max(0,start-a)]
            right = original[max(0,end-a):] if end < b else ""
            t.text = left + (value if not inserted else "") + right
            inserted = True
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    if not inserted:
        raise ValueError("span insertion failed")

