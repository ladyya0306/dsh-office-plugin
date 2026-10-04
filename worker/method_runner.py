"""Fixed bridge to Monty: no CPython exec/import of saved method source."""
import base64
import hashlib
import json
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from worker.office_worker import WorkerError, _sha256
from worker.method_io import OfficeCallbacks, plain, output_name
from worker.output_check import check_output


def run(request):
    code=request.get('code')
    if not isinstance(code,str) or not code.strip() or len(code.encode('utf-8'))>65536:
        raise WorkerError('method_code_limit','方法程序需非空且不超过64KiB')
    if _sha256(code.encode('utf-8'))!=request.get('code_sha256'):
        raise WorkerError('method_code_changed','代码版本不一致')
    dependency=Path(request['dependency_root']).resolve()
    binary=Path(request['monty_binary']).resolve()
    if dependency not in binary.parents or _sha256(binary.read_bytes())!=request['monty_sha256']:
        raise WorkerError('method_runtime_changed','解释器版本不一致')
    sys.path.insert(0,str(dependency))
    from pydantic_monty import Monty
    grouped='artifacts' in request
    if grouped:
        from worker.method_group_io import GroupCallbacks
        office=GroupCallbacks(request)
    else: office=OfficeCallbacks(request)
    office.check_conditions(request.get('conditions'))
    prints=[]
    def on_print(stream,text):
        if sum(len(v) for v in prints)+len(text)>32768:
            raise WorkerError('method_print_limit','方法打印超过32KiB')
        prints.append(text)
    # No mounts, OS handler, host object wrappers or environment exposure.
    with Monty(binary_path=str(binary),min_processes=1,max_processes=1) as pool:
        with pool.checkout(script_name='method.py',limits={
            'max_memory':64*1024*1024,'max_feed_duration_secs':5.0,
            'max_recursion_depth':100,'max_suspensions':256}) as session:
            value=session.feed_run(code+'\nrun(inputs)\n',
                inputs={'inputs':plain(request.get('inputs',{}))},
                external_lookup=office.functions(),print_callback=on_print)
    result=plain(value)
    if not office.outputs:
        return {'ok':True,'operation':'method_execute','complete':False,'data':result,
                'reason':'程序未生成文件；本版不把纯数据结果启用为文件处理方法'}
    generated=[]
    group_check=None
    if grouped:
        if set(office.outputs)!={item['output_name'] for item in request['artifacts']}:
            raise WorkerError('method_group_incomplete','程序生成数量与本批交付清单不同；停止整组发布')
        from worker.group_output_check import check_group
        group_check=check_group({'sources':request['sources'],'artifacts':[
            {**item,'output_path':item['output_name'],'artifact':{'format':item['target']['format'],
                'bytes_b64':base64.b64encode(office.outputs[item['output_name']]).decode()}}
            for item in request['artifacts']], 'consistency':request['consistency']})
    for name,data in office.outputs.items():
        item=next((item for item in request['artifacts'] if item['output_name']==name),None) if grouped else None
        fmt=item['target']['format'] if grouped else office.target['format']
        verified=next(r for r in group_check['artifacts'] if r['output_path']==name) if grouped else check_output({'sources':request['sources'],'target':request['target'],
            'artifact':{'format':fmt,'bytes_b64':base64.b64encode(data).decode()},
            'contract':request['contract']})
        output_name(name,fmt)
        # Staged artifact only; host alone publishes after independent checks.
        if not grouped or not request.get('persist_generated'):
            with open(name,'xb') as stream: stream.write(data)
        elif _sha256(Path(name).read_bytes())!=_sha256(data):
            raise WorkerError('method_output_changed','回调暂存产物已改变')
        generated.append({'name':name,'sha256':_sha256(data),'verification':verified})
    return {'ok':True,'operation':'method_execute','complete':group_check['complete'] if grouped else all(g['verification']['complete'] for g in generated),
            **({'group_check':group_check} if grouped else {}),
            'generated':generated,'data':result,'code_sha256':request['code_sha256'],
            'runtime':{'name':'pydantic-monty','version':'1.0.0','binary_sha256':request['monty_sha256'],
                       'filesystem_mounts':False,'os_handler':False,'limits':{'memory_bytes':64*1024*1024,'seconds':5}},
            'scope':'方法所列明确要求；非全部陌生业务或完整操作系统沙箱'}


if __name__=='__main__':
    try:
        raw=sys.stdin.buffer.read(28*1024*1024+1)
        if len(raw)>28*1024*1024: raise WorkerError('method_request_limit','请求超过28MiB')
        result=run(json.loads(raw.decode('utf-8')))
    except Exception as exc:
        result={'ok':False,'error':{'code':getattr(exc,'code','method_execution_failed'),
                                  'message':str(exc)},'scope':'未报告成功；暂存区可能有产物待查'}
    print(json.dumps(result,ensure_ascii=False,allow_nan=False))
    raise SystemExit(0 if result.get('ok') else 1)
