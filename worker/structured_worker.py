"""Structured operations, while old v1 worker operations remain compatible."""
from worker.office_worker import WorkerError, _decode_b64, _sha256
from worker.structure_docx import read_docx
from worker.structure_xlsx import read_xlsx
from worker.structured_fill import derive, apply_changes


def _read(item,role,fields=None,offset=0,target_mode='fill_blanks',requested_refs=None):
    if not isinstance(item,dict) or item.get("format") not in {"docx","xlsx"}:
        raise WorkerError("invalid_input","输入必须声明docx/xlsx格式")
    data = _decode_b64(item.get("bytes_b64"),role)
    scope = item.get("scope")
    if item["format"]=="docx" and scope:
        raise WorkerError("invalid_scope","本版Word按完整结构读取；范围仅接受Excel")
    try:
        if item['format']=='xlsx' and role=='source':
            from worker.source_xlsx import read_source
            records=read_source(data,scope,fields,offset,requested_refs)
        else:
            records = read_docx(data,role,target_mode) if item["format"]=="docx" else read_xlsx(data,role,scope,target_mode)
    except WorkerError as exc:
        if exc.code in {"unsupported_xlsx_feature","unsupported_docx_feature"}:
            records = {"nodes":[],"slots":[],"capabilities":{"read":False,"write":False,"reason":exc.message}}
        else:
            raise
    if len(records["nodes"])>1000 or len(records["slots"])>1000:
        raise WorkerError("structure_limit","结构超过1000项，未截断；请使用原生路径")
    return {**records,"format":item["format"],"bytes":data,"sha256":_sha256(data)}


def run_v3(request):
    source_items = request.get("sources")
    if not isinstance(source_items,list) or len(source_items)>8 or not source_items and not request.get('target'):
        raise WorkerError("source_limit","来源需1至8份；只看目标时sources可为空数组")
    mode=request.get('target_mode','fill_blanks')
    if mode not in {'fill_blanks','review_values'}:
        raise WorkerError('invalid_target_mode','目标模式只能fill_blanks/review_values')
    inputs = [_read(item,"source",request.get('requested_fields'),request.get('source_offset',0)) for item in source_items]
    target = _read(request["target"],"target",target_mode=mode) if request.get("target") else None
    common = {"source_sha256":[s["sha256"] for s in inputs],"target_sha256":target["sha256"] if target else None}
    operation = request["operation"]
    if operation=="inspect_v3":
        candidates = [{**n,"source_ref":f"s{i}::{n['ref']}","source_index":i}
                      for i,s in enumerate(inputs) for n in s["nodes"]]
        return {"ok":True,"operation":"inspect",**common,
                "source_candidates":candidates,"target_slots":target["slots"] if target else [],
                "capabilities":{"sources":[s["capabilities"] for s in inputs],
                                "target":target["capabilities"] if target else None},
                "interpretation":"原文及结构引用，不表示主体/键值关系已确认；仅引用本次证据。"}
    if target is None:
        raise WorkerError("missing_target","写入需要本次目标文件")
    changes = derive(inputs,target,request.get("mappings"))
    if operation=="prepare_v3":
        return {"ok":True,"operation":"prepare",**common,"changes":changes}
    submitted = request.get("changes")
    if operation!="fill_v3" and submitted!=changes:
        raise WorkerError("value_mismatch","执行内容必须与当前输入确定性推导完全相同")
    return {"ok":True,"operation":"apply",**common,
            **apply_changes(target,changes,request.get("output_name"))}

