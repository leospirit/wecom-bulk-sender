#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlencode
from urllib.request import urlopen


def _norm_key(text: str) -> str:
    s = str(text or "").strip().lower()
    out = []
    for ch in s:
        if ("a" <= ch <= "z") or ("0" <= ch <= "9") or ("\u4e00" <= ch <= "\u9fff"):
            out.append(ch)
    return "".join(out)


def _stem_candidates(image_path: str) -> list[str]:
    stem = Path(str(image_path or "")).stem
    if not stem:
        return []
    outs = [stem]
    for suffix in ("_report", "-report", " report"):
        if stem.lower().endswith(suffix):
            outs.append(stem[: -len(suffix)])
    return [x for x in outs if x]


def _http_json(url: str, timeout: float) -> Any:
    with urlopen(url, timeout=timeout) as resp:
        data = resp.read().decode("utf-8", errors="replace")
    return json.loads(data)


def _load_reports(base_url: str, timeout: float) -> list[dict[str, Any]]:
    payload = _http_json(f"{base_url.rstrip('/')}/api/reports", timeout)
    if isinstance(payload, list):
        rows = payload
    elif isinstance(payload, dict) and isinstance(payload.get("items"), list):
        rows = payload["items"]
    else:
        rows = []
    out: list[dict[str, Any]] = []
    for row in rows:
        if isinstance(row, dict):
            out.append(row)
    out.sort(key=lambda x: float(x.get("timestamp") or 0.0), reverse=True)
    return out


def _extract_report_names(report: dict[str, Any]) -> list[str]:
    names = []
    for k in ("student_name", "display_name", "original_filename", "id"):
        raw = str(report.get(k) or "").strip()
        if raw:
            names.append(raw)
            if k == "original_filename":
                names.append(Path(raw).stem)
    dedup: list[str] = []
    seen: set[str] = set()
    for name in names:
        key = _norm_key(name)
        if not key or key in seen:
            continue
        seen.add(key)
        dedup.append(name)
    return dedup


def _pick_report_id(task_row: dict[str, str], reports: list[dict[str, Any]]) -> str:
    candidates: list[str] = []
    for key in ("student_name", "parent_name"):
        raw = str(task_row.get(key) or "").strip()
        if raw:
            candidates.append(raw)
    for stem in _stem_candidates(task_row.get("image_path", "")):
        candidates.append(stem)

    norm_candidates = [_norm_key(x) for x in candidates if _norm_key(x)]
    if not norm_candidates:
        return ""

    best_report_id = ""
    best_score = -1
    best_ts = -1.0

    for report in reports:
        report_id = str(report.get("id") or "").strip()
        if not report_id:
            continue
        report_ts = float(report.get("timestamp") or 0.0)
        report_names = _extract_report_names(report)
        norm_report_names = [_norm_key(x) for x in report_names if _norm_key(x)]
        if not norm_report_names:
            continue

        score = 0
        for c in norm_candidates:
            for r in norm_report_names:
                if c == r:
                    score = max(score, 100)
                elif r.startswith(c) or c.startswith(r):
                    score = max(score, 85)
                elif c in r or r in c:
                    score = max(score, 72)
        if score > best_score or (score == best_score and report_ts > best_ts):
            best_score = score
            best_ts = report_ts
            best_report_id = report_id

    return best_report_id if best_score >= 72 else ""


def _fetch_message(
    base_url: str,
    submission_id: str,
    timeout: float,
    top_n: int,
    per_phoneme: int,
    max_links: int,
) -> dict[str, Any]:
    params = urlencode(
        {
            "top_n": max(1, int(top_n)),
            "per_phoneme": max(1, int(per_phoneme)),
            "max_links": max(1, int(max_links)),
        }
    )
    url = f"{base_url.rstrip('/')}/api/reports/{quote(submission_id, safe='')}/phoneme-video-message?{params}"
    data = _http_json(url, timeout)
    return data if isinstance(data, dict) else {}


def _read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        fieldnames = [str(x or "").strip() for x in (reader.fieldnames or []) if str(x or "").strip()]
        rows = [dict(r) for r in reader]
    return fieldnames, rows


def _write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def main() -> int:
    parser = argparse.ArgumentParser(description="Enrich WeCom RPA tasks CSV with score-reading phoneme video messages.")
    parser.add_argument("--input-csv", default="tools/rpa_tasks.real.csv", help="Base tasks CSV")
    parser.add_argument("--output-csv", default="tools/rpa_tasks.with-message.csv", help="Output CSV with message_text")
    parser.add_argument("--score-api-base", default="http://127.0.0.1:8000", help="Score-reading API base URL")
    parser.add_argument("--timeout-sec", type=float, default=12.0, help="HTTP timeout")
    parser.add_argument("--top-n", type=int, default=3, help="Weak phoneme count")
    parser.add_argument("--per-phoneme", type=int, default=2, help="Videos per phoneme in source payload")
    parser.add_argument("--max-links", type=int, default=3, help="Max links in integrated message")
    args = parser.parse_args()

    input_path = Path(args.input_csv)
    if not input_path.exists():
        print(f"[ERR] input csv not found: {input_path}")
        return 2

    try:
        reports = _load_reports(args.score_api_base, timeout=float(args.timeout_sec))
    except Exception as exc:
        print(f"[ERR] cannot fetch reports from score api: {exc}")
        return 3
    if not reports:
        print("[ERR] report list is empty, cannot enrich tasks.")
        return 4

    fieldnames, rows = _read_csv(input_path)
    if not rows:
        print("[ERR] input csv has no rows.")
        return 5
    if "parent_name" not in fieldnames or "image_path" not in fieldnames:
        print("[ERR] csv must include parent_name,image_path")
        return 6

    extra_fields = ["report_submission_id", "weak_phonemes", "video_links", "message_text", "message_source"]
    for col in extra_fields:
        if col not in fieldnames:
            fieldnames.append(col)

    ok_count = 0
    miss_count = 0
    err_count = 0

    for row in rows:
        for col in extra_fields:
            row[col] = str(row.get(col) or "").strip()

        report_id = _pick_report_id(row, reports)
        if not report_id:
            row["message_source"] = "no_report_match"
            miss_count += 1
            continue

        row["report_submission_id"] = report_id
        try:
            payload = _fetch_message(
                base_url=args.score_api_base,
                submission_id=report_id,
                timeout=float(args.timeout_sec),
                top_n=int(args.top_n),
                per_phoneme=int(args.per_phoneme),
                max_links=int(args.max_links),
            )
        except Exception as exc:
            row["message_source"] = f"fetch_error:{str(exc)[:120]}"
            err_count += 1
            continue

        weak_phonemes = payload.get("weak_phonemes") or []
        links = payload.get("links") or []
        if isinstance(weak_phonemes, list):
            row["weak_phonemes"] = ",".join(str(x or "").strip() for x in weak_phonemes if str(x or "").strip())
        if isinstance(links, list):
            row["video_links"] = " | ".join(str((x or {}).get("url") or "").strip() for x in links if isinstance(x, dict))

        message_text = str(payload.get("message_text") or "").strip()
        if message_text:
            row["message_text"] = message_text
            row["message_source"] = "score_api"
            ok_count += 1
        else:
            row["message_source"] = "empty_message"
            miss_count += 1

    output_path = Path(args.output_csv)
    _write_csv(output_path, fieldnames, rows)

    print(f"[OK] output: {output_path}")
    print(f"[OK] total={len(rows)} enriched={ok_count} no_match_or_empty={miss_count} errors={err_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
