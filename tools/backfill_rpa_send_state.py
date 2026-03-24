#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
from pathlib import Path


HANDLED_SEND_STATUSES = {"ok", "sent", "pasted_only", "pasted_unverified"}
STATE_FIELDNAMES = [
    "parent_name",
    "student_name",
    "class_name",
    "report_submission_id",
    "image_path",
    "status",
    "text_status",
    "updated_at",
]


def _norm(value: str) -> str:
    return str(value or "").strip()


def _norm_message(value: str) -> str:
    return _norm(value).replace("\r\n", "\n").replace("\r", "\n")


def _image_name(value: str) -> str:
    raw = _norm(value)
    return Path(raw).name if raw else ""


def state_row_key(row: dict) -> tuple[str, str, str]:
    return (
        _norm(row.get("class_name", "")).casefold(),
        _norm(row.get("student_name", "")).casefold(),
        _norm(row.get("report_submission_id", "")).casefold(),
    )


def merge_state_rows(previous_rows: list[dict], new_rows: list[dict]) -> list[dict]:
    new_map = {state_row_key(r): r for r in new_rows}
    merged: list[dict] = []
    used = set()
    for row in previous_rows:
        key = state_row_key(row)
        if key in new_map:
            merged.append(new_map[key])
            used.add(key)
        else:
            merged.append(row)
    for row in new_rows:
        key = state_row_key(row)
        if key not in used:
            merged.append(row)
    return merged


def read_csv_rows(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_state_rows(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=STATE_FIELDNAMES)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def build_backfill_state_rows(result_rows: list[dict], task_rows: list[dict]) -> tuple[list[dict], dict[str, int]]:
    summary = {
        "handled_results": 0,
        "matched": 0,
        "unmatched": 0,
        "ambiguous": 0,
        "ignored_status": 0,
    }

    task_index: dict[tuple[str, str, str, str], list[dict]] = {}
    fallback_index: dict[tuple[str, str, str], list[dict]] = {}
    for task in task_rows:
        fallback_key = (
            _norm(task.get("parent_name", "")).casefold(),
            _norm(task.get("student_name", "")).casefold(),
            _image_name(task.get("image_path", "")).casefold(),
        )
        key = (
            fallback_key[0],
            fallback_key[1],
            fallback_key[2],
            _norm_message(task.get("message_text", "")).casefold(),
        )
        task_index.setdefault(key, []).append(task)
        fallback_index.setdefault(fallback_key, []).append(task)

    state_rows: list[dict] = []
    for result in result_rows:
        status = _norm(result.get("status", ""))
        if status not in HANDLED_SEND_STATUSES:
            summary["ignored_status"] += 1
            continue

        summary["handled_results"] += 1
        key = (
            _norm(result.get("parent_name", "")).casefold(),
            _norm(result.get("student_name", "")).casefold(),
            _image_name(result.get("image_path", "")).casefold(),
            _norm_message(result.get("message_text", "")).casefold(),
        )
        matches = task_index.get(key, [])
        if not matches:
            fallback_key = (key[0], key[1], key[2])
            fallback_matches = fallback_index.get(fallback_key, [])
            if len(fallback_matches) == 1:
                matches = fallback_matches
        if not matches:
            summary["unmatched"] += 1
            continue
        if len(matches) > 1:
            summary["ambiguous"] += 1
            continue

        task = matches[0]
        state_rows.append(
            {
                "parent_name": _norm(task.get("parent_name", "")),
                "student_name": _norm(task.get("student_name", "")),
                "class_name": _norm(task.get("class_name", "")),
                "report_submission_id": _norm(task.get("report_submission_id", "")),
                "image_path": _norm(task.get("image_path", "")),
                "status": status,
                "text_status": _norm(result.get("text_status", "")),
                "updated_at": _norm(result.get("timestamp", "")),
            }
        )
        summary["matched"] += 1

    return state_rows, summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Backfill rpa_send_state.csv from historical rpa-results.csv")
    repo_root = Path(__file__).resolve().parents[1]
    parser.add_argument("--results-csv", default=str(repo_root / "run-logs" / "rpa-results.csv"))
    parser.add_argument("--tasks-csv", default=str(repo_root / "tools" / "rpa_tasks.with-message.csv"))
    parser.add_argument("--state-csv", default=str(repo_root / "run-logs" / "rpa_send_state.csv"))
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    results_csv = Path(args.results_csv)
    tasks_csv = Path(args.tasks_csv)
    state_csv = Path(args.state_csv)

    result_rows = read_csv_rows(results_csv)
    task_rows = read_csv_rows(tasks_csv)
    previous_state_rows = read_csv_rows(state_csv)

    new_state_rows, summary = build_backfill_state_rows(result_rows, task_rows)
    merged_state_rows = merge_state_rows(previous_state_rows, new_state_rows)
    write_state_rows(state_csv, merged_state_rows)

    print(
        "handled_results={handled_results} matched={matched} unmatched={unmatched} ambiguous={ambiguous} "
        "ignored_status={ignored_status} merged_total={merged_total}".format(
            **summary,
            merged_total=len(merged_state_rows),
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
