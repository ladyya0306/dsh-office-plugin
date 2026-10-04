"""Copy independently checked candidate values into the original template bytes."""
import base64
from worker.office_worker import WorkerError, _write_exclusive
from worker.output_check import check_output, _coverage
from worker.structured_worker import _read
from worker.method_io import OfficeCallbacks


def restore_template(request):
    contract=request.get('contract')
    if not isinstance(contract,dict) or contract.get('preserve_template') is False:
        raise WorkerError('invalid_restore_contract','模板保留需完整清单且preserve_template不能为false')
    candidate_request={**request,'contract':{**contract,'preserve_template':False}}
    checked=check_output(candidate_request)
    if not checked['complete']:
        return {**checked,'operation':'restore_template','published':False,
                'stage':'candidate_values','next':'先修正候选值或检查依据，程序不以格式修复掩盖内容错误'}
    fmt=request['target']['format']
    office=OfficeCallbacks(request)
    candidate=_read(request['artifact'],'source')
    nodes={node['ref']:node for node in candidate['nodes']}
    coverage=_coverage(contract['coverage'],fmt)
    if fmt=='docx':
        from worker.docx_output_check import docx_output_nodes
        nodes.update(docx_output_nodes(office.target,candidate,coverage))
    changes=[]
    for ref in sorted(coverage):
        value=nodes.get(ref,{}).get('value')
        changes.append({'target_ref':ref,'value': '' if fmt=='docx' and value is None else value})
    name=request.get('output_name')
    office.write_cells(changes,name,contract.get('table_ranges',{}))
    data=office.outputs[name]
    verified=check_output({**request,'artifact':{'format':fmt,'bytes_b64':base64.b64encode(data).decode()}})
    if not verified['complete']:
        return {**verified,'artifact_sha256':checked['artifact_sha256'],
                'staged_artifact_sha256':verified['artifact_sha256'],
                'operation':'restore_template','published':False,'stage':'restored_verification'}
    _write_exclusive(name,data)
    return {**verified,'operation':'restore_template','published':True,'output_name':name,
            'changed':len(changes),'next':'当前来源值和模板保留已按完整清单独立检查；据实际输出交付，并评估可重复步骤'}
