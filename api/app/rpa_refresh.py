from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path
from typing import Any


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _build_refresh_command(score_api_base: str = "http://host.docker.internal:8010") -> list[str]:
    script = _repo_root() / "tools" / "build_rpa_tasks_from_score.py"
    contacts = "/data/contacts.xlsx"
    output = _repo_root() / "tools" / "rpa_tasks.with-message.csv"
    delta_output = _repo_root() / "tools" / "rpa_tasks.delta.csv"
    pending_output = _repo_root() / "tools" / "rpa_tasks.pending.csv"
    state_csv = _repo_root() / "run-logs" / "rpa_send_state.csv"
    return [
        sys.executable,
        str(script),
        "--report-driven",
        "--contacts-xlsx",
        contacts,
        "--output-csv",
        str(output),
        "--delta-output-csv",
        str(delta_output),
        "--pending-output-csv",
        str(pending_output),
        "--state-csv",
        str(state_csv),
        "--score-api-base",
        str(score_api_base),
    ]


def _build_backfill_command() -> list[str]:
    script = _repo_root() / "tools" / "backfill_rpa_send_state.py"
    results_csv = _repo_root() / "run-logs" / "rpa-results.csv"
    tasks_csv = _repo_root() / "tools" / "rpa_tasks.with-message.csv"
    state_csv = _repo_root() / "run-logs" / "rpa_send_state.csv"
    return [
        sys.executable,
        str(script),
        "--results-csv",
        str(results_csv),
        "--tasks-csv",
        str(tasks_csv),
        "--state-csv",
        str(state_csv),
    ]


def _parse_refresh_summary(stdout: str) -> dict[str, Any]:
    m = re.search(r"total=(\d+)\s+enriched=(\d+)\s+no_match_or_empty=(\d+)\s+errors=(\d+)(?:\s+delta=(\d+))?(?:\s+pending=(\d+))?", str(stdout or ""))
    if not m:
        return {"total": 0, "enriched": 0, "no_match_or_empty": 0, "errors": 0, "delta": 0, "pending": 0}
    return {
        "total": int(m.group(1)),
        "enriched": int(m.group(2)),
        "no_match_or_empty": int(m.group(3)),
        "errors": int(m.group(4)),
        "delta": int(m.group(5) or 0),
        "pending": int(m.group(6) or 0),
    }


def _parse_backfill_summary(stdout: str) -> dict[str, Any]:
    m = re.search(
        r"handled_results=(\d+)\s+matched=(\d+)\s+unmatched=(\d+)\s+ambiguous=(\d+)\s+ignored_status=(\d+)\s+merged_total=(\d+)",
        str(stdout or ""),
    )
    if not m:
        return {
            "handled_results": 0,
            "matched": 0,
            "unmatched": 0,
            "ambiguous": 0,
            "ignored_status": 0,
            "merged_total": 0,
        }
    return {
        "handled_results": int(m.group(1)),
        "matched": int(m.group(2)),
        "unmatched": int(m.group(3)),
        "ambiguous": int(m.group(4)),
        "ignored_status": int(m.group(5)),
        "merged_total": int(m.group(6)),
    }


def run_refresh(score_api_base: str = "http://host.docker.internal:8010") -> dict[str, Any]:
    cmd = _build_refresh_command(score_api_base=score_api_base)
    proc = subprocess.run(
        cmd,
        cwd=str(_repo_root()),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    payload = {
        "ok": proc.returncode == 0,
        "return_code": proc.returncode,
        "command": cmd,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
    }
    payload.update(_parse_refresh_summary(proc.stdout))
    return payload


def run_backfill() -> dict[str, Any]:
    cmd = _build_backfill_command()
    proc = subprocess.run(
        cmd,
        cwd=str(_repo_root()),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    payload = {
        "ok": proc.returncode == 0,
        "return_code": proc.returncode,
        "command": cmd,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
    }
    payload.update(_parse_backfill_summary(proc.stdout))
    return payload
