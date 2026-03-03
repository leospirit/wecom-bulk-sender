from __future__ import annotations

import csv
import os
import subprocess
import sys
import threading
import time
from collections import Counter, deque
from pathlib import Path
from typing import Any


_LOCK = threading.Lock()
_PROC: subprocess.Popen | None = None
_STATE: dict[str, Any] = {
    "started_at": None,
    "finished_at": None,
    "return_code": None,
    "command": [],
    "log_file": "",
    "results_csv": "",
}


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _resolve_path(path_str: str | None, default_rel: str) -> Path:
    raw = (path_str or "").strip()
    p = Path(raw) if raw else Path(default_rel)
    if not p.is_absolute():
        p = _repo_root() / p
    return p


def _build_command(cfg: dict[str, Any], log_file: Path, results_csv: Path) -> list[str]:
    script = _repo_root() / "tools" / "wecom_rpa_sender.py"
    tasks_csv = _resolve_path(cfg.get("tasks_csv"), "tools/rpa_tasks.real.csv")
    if not script.exists():
        raise FileNotFoundError(f"RPA sender not found: {script}")
    if not tasks_csv.exists():
        raise FileNotFoundError(f"Tasks CSV not found: {tasks_csv}")

    cmd = [
        sys.executable,
        str(script),
        "--tasks-csv",
        str(tasks_csv),
        "--send-mode",
        str(cfg.get("send_mode", "clipboard")),
        "--main-title-re",
        str(cfg.get("main_title_re", r".*(WeCom|WXWork|企业微信).*")),
        "--interval-sec",
        str(cfg.get("interval_sec", 3)),
        "--timeout-sec",
        str(cfg.get("timeout_sec", 20)),
        "--max-retries",
        str(cfg.get("max_retries", 2)),
        "--retry-delay-sec",
        str(cfg.get("retry_delay_sec", 1.0)),
        "--stabilize-open-rounds",
        str(cfg.get("stabilize_open_rounds", 2)),
        "--stabilize-focus-rounds",
        str(cfg.get("stabilize_focus_rounds", 2)),
        "--stabilize-send-rounds",
        str(cfg.get("stabilize_send_rounds", 1)),
        "--open-chat-strategy",
        str(cfg.get("open_chat_strategy", "keyboard_first")),
        "--resume-from",
        str(cfg.get("resume_from", 1)),
        "--results-csv",
        str(results_csv),
        "--log-file",
        str(log_file),
    ]
    wecom_exe = (cfg.get("wecom_exe") or "").strip()
    if wecom_exe:
        cmd.extend(["--wecom-exe", wecom_exe])
    if cfg.get("dry_run"):
        cmd.append("--dry-run")
    if cfg.get("paste_only", True):
        cmd.append("--paste-only")
    if cfg.get("no_chat_verify", True):
        cmd.append("--no-chat-verify")
    if cfg.get("resume_failed"):
        cmd.append("--resume-failed")
        cmd.extend(["--resume-results-csv", str(results_csv)])
    if cfg.get("stop_on_fail", True):
        cmd.append("--stop-on-fail")
    if cfg.get("skip_missing_image", True):
        cmd.append("--skip-missing-image")
    if cfg.get("debug_chat_text"):
        cmd.append("--debug-chat-text")
    return cmd


def _current_running() -> bool:
    return _PROC is not None and _PROC.poll() is None


def _refresh_exit_state():
    if _PROC is None:
        return
    rc = _PROC.poll()
    if rc is None:
        return
    if _STATE.get("finished_at") is None:
        _STATE["finished_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        _STATE["return_code"] = rc


def start_run(cfg: dict[str, Any]) -> dict[str, Any]:
    global _PROC
    with _LOCK:
        _refresh_exit_state()
        if _current_running():
            raise RuntimeError("RPA run is already in progress.")

        log_file = _resolve_path(cfg.get("log_file"), "run-logs/rpa-sender.log")
        results_csv = _resolve_path(cfg.get("results_csv"), "run-logs/rpa-results.csv")
        log_file.parent.mkdir(parents=True, exist_ok=True)
        results_csv.parent.mkdir(parents=True, exist_ok=True)
        log_file.write_text("", encoding="utf-8")

        cmd = _build_command(cfg, log_file, results_csv)
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        _PROC = subprocess.Popen(
            cmd,
            cwd=str(_repo_root()),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            env=env,
            creationflags=creationflags,
        )
        _STATE.update(
            {
                "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "finished_at": None,
                "return_code": None,
                "command": cmd,
                "log_file": str(log_file),
                "results_csv": str(results_csv),
            }
        )
        return status()


def stop_run() -> dict[str, Any]:
    global _PROC
    with _LOCK:
        if _PROC is None or _PROC.poll() is not None:
            _refresh_exit_state()
            return status()
        _PROC.terminate()
        try:
            _PROC.wait(timeout=8)
        except subprocess.TimeoutExpired:
            _PROC.kill()
            _PROC.wait(timeout=5)
        _refresh_exit_state()
        return status()


def _read_result_counts(path: Path) -> dict[str, int]:
    if not path.exists():
        return {}
    counter: Counter[str] = Counter()
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                status = (row.get("status") or "").strip()
                if status:
                    counter[status] += 1
    except Exception:
        return {}
    return dict(counter)


def status() -> dict[str, Any]:
    with _LOCK:
        _refresh_exit_state()
        log_raw = str(_STATE.get("log_file") or "").strip()
        res_raw = str(_STATE.get("results_csv") or "").strip()
        log_file = Path(log_raw) if log_raw else None
        results_csv = Path(res_raw) if res_raw else None
        return {
            "running": _current_running(),
            "pid": _PROC.pid if _PROC else None,
            "started_at": _STATE.get("started_at"),
            "finished_at": _STATE.get("finished_at"),
            "return_code": _STATE.get("return_code"),
            "command": _STATE.get("command") or [],
            "log_file": str(log_file) if log_file is not None else "",
            "results_csv": str(results_csv) if results_csv is not None else "",
            "result_counts": _read_result_counts(results_csv) if results_csv is not None else {},
        }


def tail_log(lines: int = 160) -> dict[str, Any]:
    with _LOCK:
        _refresh_exit_state()
        log_raw = str(_STATE.get("log_file") or "").strip()
        if not log_raw:
            return {"path": "", "lines": []}
        log_file = Path(log_raw)
        n = max(20, min(int(lines), 800))
        if not log_file.exists():
            return {"path": str(log_file), "lines": []}
        buf: deque[str] = deque(maxlen=n)
        try:
            with log_file.open("r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    buf.append(line.rstrip("\n"))
        except Exception:
            return {"path": str(log_file), "lines": []}
        return {"path": str(log_file), "lines": list(buf)}
