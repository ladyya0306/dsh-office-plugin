"""Optional workspace provenance storage; never stores executable plans."""
import json
import os
from pathlib import Path
import sqlite3
import uuid
from worker.office_worker import WorkerError, _sha256

APPLICATION_ID = 0x44534833
DB_NAME = "office-plugin-records.sqlite3"


def _connect(writing):
    root = Path.cwd().resolve()
    path = root / DB_NAME
    if path.is_symlink() or path.resolve().parent != root:
        raise WorkerError("records_path_denied","数据库链接越出工作区")
    if not path.exists():
        if not writing:
            return None
        try:
            fd = os.open(path,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
            os.close(fd)
        except FileExistsError as exc:
            raise WorkerError("records_race","数据库同时创建，请检查后重试") from exc
        connection = sqlite3.connect(path)
        try:
            with connection:
                connection.execute(f"PRAGMA application_id={APPLICATION_ID}")
                connection.execute("CREATE TABLE snapshots (id TEXT PRIMARY KEY, created TEXT DEFAULT CURRENT_TIMESTAMP, manifest TEXT NOT NULL, facts TEXT NOT NULL)")
                connection.execute("CREATE TABLE cases (id TEXT PRIMARY KEY, snapshot_id TEXT NOT NULL, source_ref TEXT NOT NULL, proposal TEXT NOT NULL, status TEXT NOT NULL, created TEXT DEFAULT CURRENT_TIMESTAMP)")
        except Exception:
            connection.close()
            raise
        return connection
    probe = sqlite3.connect(path.as_uri()+"?mode=ro",uri=True)
    try:
        if probe.execute("PRAGMA application_id").fetchone()[0]!=APPLICATION_ID:
            raise WorkerError("records_not_owned","已有同名文件不属于插件，拒绝使用或覆盖")
    except sqlite3.DatabaseError as exc:
        raise WorkerError("records_not_owned","已有同名文件不是插件数据库，拒绝使用或覆盖") from exc
    finally:
        probe.close()
    return sqlite3.connect(path) if writing else sqlite3.connect(path.as_uri()+"?mode=ro",uri=True)


def run_records(request):
    writing = request["operation"]=="records_store"
    connection = _connect(writing)
    if connection is None:
        return {"ok":True,"stored":False,"snapshots":[],"cases":[]}
    try:
        if writing:
            from worker.structured_worker import run_v3
            inspection = run_v3({**request,"operation":"inspect_v3"})
            manifest = {"sources":inspection["source_sha256"],"target":inspection["target_sha256"],
                        "paths":request.get("paths"),
                        "source_scopes":[s.get("scope") for s in request["sources"]],
                        "target_scope":(request.get("target") or {}).get("scope"),
                        'requested_fields':request.get('requested_fields'),'source_offset':request.get('source_offset',0),
                        'target_mode':request.get('target_mode','fill_blanks'),
                        "parser_version":_sha256(b"".join((Path(__file__).parent/name).read_bytes() for name in
                                                   ("structure_docx.py","structure_xlsx.py","source_xlsx.py","structured_worker.py")))}
            facts = {"source_candidates":inspection["source_candidates"],
                     "target_slots":inspection["target_slots"],"capabilities":inspection["capabilities"]}
            text = json.dumps(manifest,ensure_ascii=False,sort_keys=True)
            identity = _sha256(text.encode())
            proposal = request.get("proposal")
            with connection:
                connection.execute("INSERT OR IGNORE INTO snapshots(id,manifest,facts) VALUES (?,?,?)",
                                   (identity,text,json.dumps(facts,ensure_ascii=False)))
                if proposal is not None:
                    if not isinstance(proposal,dict) or set(proposal)!={"source_ref","explanation"} or not isinstance(proposal["explanation"],str) or not 1<=len(proposal["explanation"])<=2000:
                        raise WorkerError("invalid_proposal","提议必须是source_ref和1至2000字explanation")
                    if proposal["source_ref"] not in {n["source_ref"] for n in facts["source_candidates"]}:
                        raise WorkerError("unknown_reference","提议引用不在本次来源")
                    connection.execute("INSERT INTO cases(id,snapshot_id,source_ref,proposal,status) VALUES (?,?,?,?,?)",
                                       (str(uuid.uuid4()),identity,proposal["source_ref"],proposal["explanation"],"unconfirmed_proposal"))
            return {"ok":True,"stored":True,"snapshot_id":identity,"database":DB_NAME,
                    "confirmation":"原文按版本保存；提议未被认定为用户确认，不自动启用全局方法"}
        identity = request.get("snapshot_id")
        rows = connection.execute("SELECT id,created,manifest,facts FROM snapshots WHERE id=?",(identity,)).fetchall() if identity else connection.execute("SELECT id,created,manifest FROM snapshots ORDER BY created DESC LIMIT 32").fetchall()
        snapshots = [{"id":r[0],"created":r[1],"manifest":json.loads(r[2]),**({"facts":json.loads(r[3])} if len(r)>3 else {})} for r in rows]
        cases = [{"id":r[0],"snapshot_id":r[1],"source_ref":r[2],"proposal":r[3],"status":r[4]}
                 for r in connection.execute("SELECT id,snapshot_id,source_ref,proposal,status FROM cases WHERE snapshot_id=? LIMIT 100",(identity,))]
        return {"ok":True,"stored":True,"snapshots":snapshots,"cases":cases,
                "reuse":"历史查看；没有核对当前摘要，不可当当前值，也不恢复任何授权"}
    except sqlite3.Error as exc:
        raise WorkerError("records_storage_failed",str(exc)) from exc
    finally:
        connection.close()

