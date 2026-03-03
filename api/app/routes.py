from __future__ import annotations

from fastapi import APIRouter, UploadFile, File, HTTPException
import requests

from .schema import (
    ScanRequest,
    SendSelectedRequest,
    DeleteSelectedRequest,
    AutoWatchRequest,
    ConfigUpdateRequest,
    RpaStartRequest,
)
from .contacts import read_contacts, build_name_index
from .scanner import build_tasks
from .db import (
    insert_task,
    list_tasks,
    status_counts,
    set_tasks_status,
    set_all_pending_to_queued,
    delete_task,
    delete_tasks,
    clear_tasks,
)
from .worker import worker
from .watch import watcher
from .config import load_config, update_config
from .rpa_runner import start_run as start_rpa_run, stop_run as stop_rpa_run, status as rpa_status, tail_log as rpa_tail_log

router = APIRouter()


@router.get("/health")
def health():
    return {"status": "ok"}


@router.get("/status")
def get_status():
    return status_counts()


@router.get("/tasks")
def get_tasks():
    return list_tasks()


@router.delete("/tasks/{task_id}")
def remove_task(task_id: int):
    deleted = delete_task(task_id)
    return {"deleted": deleted}


@router.post("/tasks/delete-selected")
def remove_selected_tasks(req: DeleteSelectedRequest):
    deleted = delete_tasks(req.taskIds)
    return {"deleted": deleted}


@router.post("/tasks/clear")
def remove_all_tasks():
    deleted = clear_tasks()
    return {"deleted": deleted}


@router.post("/scan")
def scan_files(req: ScanRequest):
    update_config({"root_path": req.rootPath})
    contacts = read_contacts("/data/contacts.xlsx")
    name_index = build_name_index(contacts)
    tasks = build_tasks(req.rootPath, name_index)
    for t in tasks:
        insert_task(t)
    return {"created": len(tasks)}


@router.post("/send/batch")
def send_batch():
    count = set_all_pending_to_queued()
    worker.start()
    return {"queued": count}


@router.post("/send/selected")
def send_selected(req: SendSelectedRequest):
    set_tasks_status(req.taskIds, "queued")
    worker.start()
    return {"queued": len(req.taskIds)}


@router.post("/auto-watch")
def auto_watch(req: AutoWatchRequest):
    cfg = load_config()
    if req.enabled:
        watcher.start(cfg.root_path)
    else:
        watcher.stop()
    return {"enabled": req.enabled}


@router.post("/contacts/upload")
def upload_contacts(file: UploadFile = File(...)):
    content = file.file.read()
    with open("/data/contacts.xlsx", "wb") as f:
        f.write(content)
    return {"ok": True}


@router.get("/config")
def get_config():
    cfg = load_config()
    return {
        "corp_id": cfg.corp_id,
        "agent_id": cfg.agent_id,
        "secret": "****" if cfg.secret else "",
        "root_path": cfg.root_path,
        "rate_limit_per_sec": cfg.rate_limit_per_sec,
        "max_concurrency": cfg.max_concurrency,
    }


@router.post("/config")
def set_config(req: ConfigUpdateRequest):
    data = req.model_dump(exclude_unset=True)
    cfg = update_config(data)
    return {"ok": True, "root_path": cfg.root_path}


@router.get("/ip")
def get_public_ip():
    r = requests.get("https://api.ipify.org?format=json", timeout=10)
    r.raise_for_status()
    return r.json()


@router.get("/rpa/status")
def get_rpa_status():
    return rpa_status()


@router.get("/rpa/log-tail")
def get_rpa_log_tail(lines: int = 160):
    return rpa_tail_log(lines=lines)


@router.post("/rpa/start")
def start_rpa(req: RpaStartRequest):
    try:
        return start_rpa_run(req.model_dump())
    except FileNotFoundError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=409, detail=str(e))


@router.post("/rpa/stop")
def stop_rpa():
    return stop_rpa_run()
