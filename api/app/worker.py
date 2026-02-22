from __future__ import annotations

import threading
import time
from concurrent.futures import ThreadPoolExecutor

from .config import load_config
from .db import fetch_tasks_by_status, update_task_status
from .wecom import WeComClient


class SendWorker:
    def __init__(self):
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()

    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()

    def _run(self):
        while not self._stop.is_set():
            cfg = load_config()
            if not cfg.corp_id or not cfg.secret or not cfg.agent_id:
                time.sleep(2)
                continue
            client = WeComClient(cfg.corp_id, cfg.secret, cfg.agent_id)
            tasks = fetch_tasks_by_status("queued", limit=50)
            if not tasks:
                time.sleep(1)
                continue

            interval = 1.0 / max(cfg.rate_limit_per_sec, 0.1)
            with ThreadPoolExecutor(max_workers=max(cfg.max_concurrency, 1)) as ex:
                futures = []
                for t in tasks:
                    if self._stop.is_set():
                        break
                    update_task_status(t["id"], "sending")
                    futures.append(ex.submit(self._send_one, client, t))
                    time.sleep(interval)
                for f in futures:
                    try:
                        f.result()
                    except Exception:
                        pass

    def _send_one(self, client: WeComClient, task: dict):
        try:
            media_id = client.upload_image(task["file_path"])
            client.send_image(task["user_id"], media_id)
            update_task_status(task["id"], "sent")
        except Exception as e:
            update_task_status(task["id"], "failed", str(e))


worker = SendWorker()
