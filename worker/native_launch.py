"""Validate then launch a child interpreter with the exact frozen entry.

The host shell still owns the process tree, policy, deadline and cancellation.
Only one leading protocol line belongs to this trusted launcher. Candidate output
cannot replace it. No candidate module is imported into the launcher or plugin.
"""
import hashlib
import json
import os
from pathlib import Path
import sys
import time
import subprocess
from native_identity import validate

PROTOCOL = 'office_native_launch_v1'


def emit(value):
    print(json.dumps({'protocol': PROTOCOL, **value}, ensure_ascii=False, separators=(',', ':')), flush=True)


def launch(request, entry, arguments):
    started = time.perf_counter()
    detail = {}
    try:
        expected_entry = (Path(request['bundle_root']) / request['files'][0]).resolve()
        if Path(entry).resolve() != expected_entry:
            raise ValueError('OFFICE_NATIVE_ENTRY_MISMATCH')
        environment = validate(request, timing_sink=detail)
        runtime = {k: v for k, v in environment.items() if k != 'cwd'}
        interpreter_hash = hashlib.sha256(Path(sys.executable).read_bytes()).hexdigest()
        if interpreter_hash != request['interpreter_sha256']:
            raise ValueError('OFFICE_NATIVE_INTERPRETER_CHANGED')
        fingerprint = hashlib.sha256(json.dumps({'environment': runtime,
            'interpreter_sha256': interpreter_hash}, ensure_ascii=False, separators=(',', ':')).encode()).hexdigest()
        proofs = request.get('expected_fingerprints')
        if proofs is not None and fingerprint not in proofs:
            raise ValueError('OFFICE_NATIVE_REVALIDATE: 当前环境需要新验证')
        for file in request['bundle_hashes']:
            path = Path(request['bundle_root']) / file['path']
            if hashlib.sha256(path.read_bytes()).hexdigest() != file['sha256']:
                raise ValueError('OFFICE_NATIVE_CODE_CHANGED')
        emit({'status': 'ready', 'environment': environment, 'runtime_fingerprint': fingerprint,
              'preflight_detail_ms': detail,
              'preflight_ms': (time.perf_counter() - started) * 1000})
    except Exception as exc:
        emit({'status': 'rejected', 'error': str(exc), 'preflight_detail_ms': detail,
              'preflight_ms': (time.perf_counter() - started) * 1000})
        return 78
    # Windows execv does not reliably preserve argv/exit status (covered by tests).
    # A shell-free child keeps normal Python entry semantics and inherits the
    # confined token. The host owns and kills the entire process tree.
    return subprocess.call([sys.executable, '-B', '-X', 'utf8', str(expected_entry), *arguments],
                           stdin=sys.stdin, stdout=sys.stdout, stderr=sys.stderr, shell=False)


if __name__ == '__main__':
    raise SystemExit(launch(json.load(sys.stdin), sys.argv[1], sys.argv[2:]))
