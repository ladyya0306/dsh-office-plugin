"""Multi-source changes and minimal package writing from current physical refs."""
import copy
import io
import re
import zipfile
from lxml import etree
from worker.office_worker import (WorkerError, _validate_output_name, _write_exclusive,
                                  _sha256, _load_xlsx, _fidelity_snapshot, _fidelity_issues)
from worker.structure_docx import NS, TAG, replace_span
from worker.xlsx_prepare import _derive_value, PrepareError
from worker.xlsx_package_patch import patch_xlsx
from worker.package_writer import replace_parts


def derive(inputs, target, mappings):
    if not isinstance(mappings,list) or not 1<=len(mappings)<=100:
        raise WorkerError("mapping_limit","必须提交1至100项对应关系")
    sources = {f"s{i}::{n['ref']}": n for i,s in enumerate(inputs) for n in s["nodes"]}
    slots = {n["target_ref"]:n for n in target["slots"]}
    changes, used = [], set()
    spans = {}
    for mapping in mappings:
        if not isinstance(mapping,dict) or not {'source_ref','target_ref'}<=set(mapping) or set(mapping)-{'source_ref','target_ref','expected_target_text'}:
            raise WorkerError("invalid_mapping","只允许本次source_ref和target_ref，不接受模型value")
        sr, tr = mapping["source_ref"],mapping["target_ref"]
        if sr not in sources or tr not in slots:
            raise WorkerError("unknown_reference","引用不在本次读取范围内")
        if tr in used:
            raise WorkerError("duplicate_target","一个物理目标只能安排一次写入")
        used.add(tr)
        node, slot = sources[sr],slots[tr]
        if slot.get('requires_expected_text') and mapping.get('expected_target_text')!=slot['text']:
            raise WorkerError('target_previous_value_mismatch','已有目标内容必须提交本次读取的完整expected_target_text；不允许直接覆盖')
        if 'expected_target_text' in mapping and mapping['expected_target_text']!=slot['text']:
            raise WorkerError('target_previous_value_mismatch','目标旧值与本次记录不符')
        if node.get("formula") or node.get("value_error"):
            raise WorkerError("invalid_source","不能使用公式或不支持类型")
        if node.get('value_type') in {'datetime', 'time'}:
            raise WorkerError('unsupported_temporal_write', '已读取日期/时间及原始依据；旧映射写入尚未支持日期类型，请使用原生程序并独立核对')
        value = node["value"]
        if value is None or str(value).strip()=="":
            raise WorkerError("empty_source","来源没有值")
        if node.get("unit_ambiguous") or slot.get("unit_ambiguous"):
            raise WorkerError("ambiguous_unit","多项单位依据冲突，不能猜单位")
        su,tu = node.get("unit",""),slot.get("unit","")
        source_format = inputs[int(sr.split("::",1)[0][1:])]["format"]
        if source_format != target["format"] and "%" in {su,tu}:
            raise WorkerError("unsupported_percent_presentation", "跨格式百分比的原始数值与显示值尚未验收，交原生核实，不能直接照抄")
        raw = value
        if isinstance(value,str) and su in {"元","万元","亿元"}:
            match = re.fullmatch(r"\s*(?:人民币\s*)?([-+]?\d+(?:\.\d+)?)\s*"+su+r"\s*",value)
            if match:
                raw = match[1]
        try:
            value, conversion = _derive_value(raw,su,tu,sr,tr)
            # Excel monetary values must be numbers, not numeric-looking text.
            if target["format"]=="xlsx" and tu in {"元","万元","亿元"} and isinstance(value,str):
                from worker.xlsx_prepare import _number
                value = _number(value,sr,tr)
        except PrepareError as exc:
            raise WorkerError(exc.code,exc.message) from exc
        if target["format"]=="docx":
            # A standalone Word cell preserves a same-unit source spelling;
            # an inline blank followed by a fixed unit must not duplicate it.
            if (slot.get("write_kind") in {"empty_cell","replace_cell"} and su==tu and su in {"元","万元","亿元"}
                and isinstance(node["value"],str) and raw!=node["value"]):
                value = node["value"]
            else:
                value = str(value)
        if slot.get("write_kind")=="text_span":
            key = tr.split("|s:")[0]
            a,b = slot["span_start"],slot["span_end"]
            for x,y in spans.setdefault(key,[]):
                if a<y and b>x or a==b==x==y:
                    raise WorkerError("overlapping_targets","文本插值范围重叠")
            spans[key].append((a,b))
        change={"source_ref":sr,"target_ref":tr,"value":value,"conversion":conversion}
        if slot.get('requires_expected_text'):
            change['previous_text']=slot['text'];change['operation']='review_replace'
        changes.append(change)
    return changes


def patch_and_verify(target, changes):
    data = target["bytes"]
    if target["format"]=="xlsx":
        applied = [{"target_ref":x["target_ref"],"value":x["value"]} for x in changes]
        wb = _load_xlsx(data,"target")
        editable = {tuple(c["target_ref"].split("!",1)) for c in changes}
        try:
            before = _fidelity_snapshot(wb,editable)
        finally:
            wb.close()
        output = patch_xlsx(data,applied)
        reopened = _load_xlsx(output,"output")
        try:
            for change in changes:
                sheet,coord = change["target_ref"].split("!",1)
                cell = reopened[sheet][coord]
                if type(cell.value)!=type(change["value"]) or cell.value!=change["value"]:
                    raise WorkerError("verification_failed","目标值或类型重开后不符")
            errors = _fidelity_issues(before,reopened,editable)
            if errors:
                raise WorkerError("fidelity_failed","；".join(errors[:6]))
        finally:
            reopened.close()
    else:
        tree = target["trees"]["word/document.xml"]
        before = copy.deepcopy(tree)
        slots = {x["target_ref"]:x for x in target["slots"]}
        changes = sorted(changes,key=lambda x:slots[x["target_ref"]].get("span_start",-1),reverse=True)
        expected = {}
        for c in changes:
            slot = slots[c["target_ref"]]
            element = target["elements"][c["target_ref"]]
            p = element.find(TAG("p")) if slot["write_kind"] in {"empty_cell","replace_cell"} else element
            texts = p.xpath("./w:r/w:t/text()",namespaces=NS)
            text = "".join(texts)
            start,end = slot["span_start"],slot["span_end"]
            replace_span(p,start,end,c["value"])
            expected[tree.getroottree().getpath(p)] = "".join(p.xpath("./w:r/w:t/text()",namespaces=NS))
        output = replace_parts(data,{"word/document.xml":etree.tostring(tree,xml_declaration=True,encoding="UTF-8",standalone=True)})
        from worker.xlsx_package_patch import _xml
        with zipfile.ZipFile(io.BytesIO(output)) as z:
            reopened = _xml(z.read("word/document.xml"))
        for path,value in expected.items():
            p = reopened.xpath(path,namespaces=reopened.nsmap)[0]
            if "".join(p.xpath("./w:r/w:t/text()",namespaces=NS))!=value:
                raise WorkerError("verification_failed","Word最终文本不符")
        # Remove the allowed text/run changes and compare all other XML.
        for path in expected:
            left = before.xpath(path,namespaces=before.nsmap)[0]
            right = reopened.xpath(path,namespaces=reopened.nsmap)[0]
            lp = left.find(TAG("pPr"))
            rp = right.find(TAG("pPr"))
            if (etree.tostring(lp) if lp is not None else None)!=(etree.tostring(rp) if rp is not None else None):
                raise WorkerError("fidelity_failed","段落样式改变")
            oldruns,newruns = left.findall(TAG("r")),right.findall(TAG("r"))
            if oldruns and len(oldruns)!=len(newruns):
                raise WorkerError("fidelity_failed","现有runs结构改变")
            for a,b in zip(oldruns,newruns):
                for r in (a,b):
                    for t in r.findall(TAG("t")):
                        # Text may be added to an existing empty run. Compare
                        # every other child/style, rather than requiring an
                        # empty run to already contain a w:t element.
                        r.remove(t)
                if etree.tostring(a)!=etree.tostring(b):
                    raise WorkerError("fidelity_failed","原有文字样式改变")
            right.getparent().replace(right,copy.deepcopy(left))
        if etree.tostring(before,method="c14n")!=etree.tostring(reopened,method="c14n"):
            raise WorkerError("fidelity_failed","非目标XML结构或格式改变")
    with zipfile.ZipFile(io.BytesIO(data)) as a,zipfile.ZipFile(io.BytesIO(output)) as b:
        if set(a.namelist())!=set(b.namelist()):
            raise WorkerError("fidelity_failed","包部件集合改变")
        for name in a.namelist():
            permitted = name=="word/document.xml" if target["format"]=="docx" else name.startswith("xl/worksheets/") and name.endswith(".xml")
            if not permitted and a.read(name)!=b.read(name):
                raise WorkerError("fidelity_failed","未授权部件改变："+name)
    return output


def apply_changes(target, changes, name):
    if (not isinstance(name,str) or not 1<=len(name)<=120 or name!=name.strip() or
        re.search(r'[\\/:*?"<>|\x00-\x1f]',name) or not name.lower().endswith("."+target["format"]) or
        re.fullmatch(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])",name.split(".",1)[0],re.I)):
        raise WorkerError("invalid_output_name","输出必须是普通新文件名，无路径或Windows保留字符，后缀与模板一致")
    output = patch_and_verify(target,changes)
    _write_exclusive(name,output)
    return {"applied":changes,"output_name":name,"output_sha256":_sha256(output),
            "verification":{"reopened":True,"values_and_types":True,"non_target_parts_preserved":True,
                            "scope":"计划内值、XML样式与非目标包部件；尚不代表全部页面视觉验收"}}

