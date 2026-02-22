from __future__ import annotations

from fastapi import APIRouter, UploadFile, File
import requests

from .schema import ScanRequest, SendSelectedRequest, AutoWatchRequest, ConfigUpdateRequest
from .contacts import read_contacts, build_name_index
from .scanner import build_tasks
from .db import (
    insert_task,
    list_tasks,
    status_counts,
    set_tasks_status,
    set_all_pending_to_queued,
)
from .worker import worker
from .watch import watcher
from .config import load_config, update_config

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
