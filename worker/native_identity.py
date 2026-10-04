"""Fixed metadata/compile probe; never import candidate code or run its business logic."""
import ast
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import sys
import time

def _record_timing(timing_sink, name, started):
    if timing_sink is not None:
        timing_sink[name] = (time.perf_counter() - started) * 1000


def _package_inventory(distributions):
    def identity(distribution):
        metadata = distribution.metadata
        return metadata['Name'], metadata['Version']
    return sorted(identity(distribution) for distribution in distributions)


def validate(request, timing_sink=None):
    compile_started = time.perf_counter()
    root = Path(request['bundle_root']).resolve()
    declared = set(request['files'])
    issues = []
    for name in request['files']:
        path = (root / name).resolve()
        if root not in path.parents or path.suffix != '.py':
            raise ValueError('invalid bundle path')
        try:
            tree = ast.parse(path.read_bytes(), filename=name)
            compile(tree, name, 'exec')
        except SyntaxError as error:
            issues.append(f'{name}:{error.lineno}:{error.offset}: syntax error: {error.msg}')
            continue
        # Static local imports must travel with the entry. Dynamic loading is outside
        # this initial supported contract and must be separately declared/reviewed.
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                modules = [a.name for a in node.names]
            elif isinstance(node, ast.ImportFrom):
                if node.level:
                    issues.append(f'{name}:{node.lineno}: relative package imports require a later supported profile')
                    continue
                modules = [node.module] if node.module else []
            else:
                continue
            for module in modules:
                top = module.split('.')[0]
                for base in [Path(request['original_script']).parent, Path(request['cwd']), root / Path(request['files'][0]).parent]:
                    single, package = base / (top + '.py'), base / top
                    if package.is_dir():
                        issues.append(f'{name}:{node.lineno}: local packages require a later supported profile: ' + top)
                    if single.is_file():
                        rel = single.relative_to(root if root in single.parents else Path(request['workspace'])).as_posix()
                        if rel not in declared:
                            issues.append(f'{name}:{node.lineno}: undeclared local code dependency: ' + rel)
                        if single.parent != Path(request['original_script']).parent and single.parent == Path(request['cwd']):
                            issues.append(f'{name}:{node.lineno}: cwd-only local imports require a later supported profile')
            if any(module == 'importlib' for module in modules):
                issues.append(f'{name}:{node.lineno}: dynamic code loading is outside initial automatic reuse')
        if any(isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id in
               {'__import__', 'eval', 'exec'} for n in ast.walk(tree)):
            issues.append(name + ': dynamic code loading is outside initial automatic reuse')
    if issues:
        _record_timing(timing_sink, 'compile_import_checks_ms', compile_started)
        raise ValueError('OFFICE_NATIVE_PREFLIGHT_ISSUES: ' + json.dumps(list(dict.fromkeys(issues)), ensure_ascii=False))
    _record_timing(timing_sink, 'compile_import_checks_ms', compile_started)
    packages_started = time.perf_counter()
    packages = _package_inventory(importlib.metadata.distributions())
    _record_timing(timing_sink, 'package_metadata_ms', packages_started)
    office_started = time.perf_counter()
    if request.get('office_request'):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
        from worker.method_io import OfficeCallbacks
        office_request = request['office_request']
        # Preflight-only cache.  It is discarded after this validation call;
        # candidate execution and output verification always parse independently.
        evidence_cache = {}
        for item in office_request.get('artifacts', [office_request]):
            office = OfficeCallbacks({**item, 'sources': office_request['sources']}, evidence_cache=evidence_cache)
            office.check_conditions(item['conditions'])
    _record_timing(timing_sink, 'office_structure_ms', office_started)
    hashes_started = time.perf_counter()
    value = {'executable': os.path.realpath(sys.executable), 'python': sys.version,
             'packages': packages, 'cwd': os.getcwd(),
             'trusted_worker_versions': [(p.name, hashlib.sha256(p.read_bytes()).hexdigest())
                                         for p in sorted(Path(__file__).resolve().parent.glob('*.py'))]}
    _record_timing(timing_sink, 'worker_hashes_ms', hashes_started)
    return value


if __name__ == '__main__':
    print(json.dumps(validate(json.load(sys.stdin)), ensure_ascii=False))
