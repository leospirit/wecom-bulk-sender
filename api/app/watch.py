from __future__ import annotations

import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from .config import load_config
from .contacts import read_contacts, build_name_index
from .scanner import build_tasks
from .db import insert_task, set_tasks_status
from .worker import worker


class _Handler(FileSystemEventHandler):
    def __init__(self, root_path: str):
        super().__init__()
        self.root_path = root_path
        self._lock = threading.Lock()

    def on_created(self, event):
        if event.is_directory:
            return
        self._scan()

    def _scan(self):
        with self._lock:
            cfg = load_config()
            contacts = read_contacts("/data/contacts.xlsx")
            idx = build_name_index(contacts)
            for t in build_tasks(cfg.root_path, idx):
                new_id = insert_task(t)
                if new_id and t.get("status") == "pending":
                    set_tasks_status([new_id], "queued")
                    worker.start()


class Watcher:
    def __init__(self):
        self._observer: Observer | None = None

    def start(self, root_path: str):
        if self._observer:
            return
        handler = _Handler(root_path)
        observer = Observer()
        observer.schedule(handler, root_path, recursive=True)
        observer.start()
        self._observer = observer

    def stop(self):
        if not self._observer:
            return
        self._observer.stop()
        self._observer.join(timeout=2)
        self._observer = None


watcher = Watcher()
