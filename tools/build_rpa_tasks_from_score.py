#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlencode
from urllib.request import urlopen
import zipfile
import xml.etree.ElementTree as ET

try:
    import pandas as pd
except ModuleNotFoundError:
    pd = None

DEFAULT_EXPORT_IMAGES_DIR = Path(r"D:\score_reading_fresh\data\out")


def _norm_key(text: str) -> str:
    s = str(text or "").strip().lower()
    out = []
    for ch in s:
        if ("a" <= ch <= "z") or ("0" <= ch <= "9") or ("\u4e00" <= ch <= "\u9fff"):
            out.append(ch)
    return "".join(out)


def _default_delta_output_csv() -> str:
    return "tools/rpa_tasks.delta.csv"


def _default_pending_output_csv() -> str:
    return "tools/rpa_tasks.pending.csv"


def _default_state_csv() -> str:
    return "run-logs/rpa_send_state.csv"


HANDLED_SEND_STATUSES = {"ok", "sent", "pasted_only", "pasted_unverified"}


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




def _norm_cjk(text: str) -> str:
    return "".join(ch for ch in str(text or "").strip() if "\u4e00" <= ch <= "\u9fff" or ch.isdigit())


def _extract_report_student_name(report: dict[str, Any]) -> str:
    for key in ("student_name", "display_name", "original_filename"):
        raw = str(report.get(key) or "").strip()
        if not raw:
            continue
        base = Path(raw).stem
        out = []
        for ch in base:
            if "一" <= ch <= "鿿":
                out.append(ch)
            else:
                break
        if out:
            return "".join(out)
    return ""


def _extract_parent_student_name(parent_name: str) -> str:
    text = str(parent_name or "").strip()
    for suffix in ("妈妈", "爸爸", "家长"):
        if text.endswith(suffix):
            return text[: -len(suffix)]
    return text


def _extract_parent_student_names(parent_name: str) -> list[str]:
    base = _extract_parent_student_name(parent_name)
    if not base:
        return []
    names: list[str] = []
    for part in base.replace("／", "/").split("/"):
        cleaned = str(part or "").strip()
        if cleaned and cleaned not in names:
            names.append(cleaned)
    return names


def _extract_class_name(dept_text: str) -> str:
    text = str(dept_text or "")
    groups = [group.strip() for group in text.split(";") if group.strip()]
    for group in groups:
        parts = [p.strip() for p in group.split("/") if p.strip()]
        for part in reversed(parts):
            if "班" in part and "家长" not in part:
                return part
    return ""


def _parent_priority(name: str) -> int:
    text = str(name or "")
    if text.endswith("妈妈"):
        return 0
    if text.endswith("爸爸"):
        return 1
    return 2


def _choose_better_parent(current: dict | None, candidate: dict) -> dict:
    if current is None:
        return candidate
    cur_p = _parent_priority(current.get("姓名", ""))
    cand_p = _parent_priority(candidate.get("姓名", ""))
    if cand_p < cur_p:
        return candidate
    return current


def _build_contact_indexes(rows: list[dict[str, str]]) -> dict[str, dict]:
    by_student: dict[str, dict] = {}
    by_class_student: dict[tuple[str, str], dict] = {}
    for raw in rows:
        row = dict(raw)
        parent_name = str(row.get("姓名") or "").strip()
        user_id = str(row.get("账号") or "").strip()
        class_name = _extract_class_name(str(row.get("部门") or ""))
        student_names = _extract_parent_student_names(parent_name)
        if not parent_name or not user_id or not student_names:
            continue
        row["_student_name"] = student_names[0]
        row["_student_names"] = student_names
        row["_class_name"] = class_name
        for student_name in student_names:
            s_key = _norm_cjk(student_name)
            if s_key:
                by_student[s_key] = _choose_better_parent(by_student.get(s_key), row)
            cs_key = (_norm_cjk(class_name), _norm_cjk(student_name))
            if cs_key[1]:
                by_class_student[cs_key] = _choose_better_parent(by_class_student.get(cs_key), row)
    return {"by_student": by_student, "by_class_student": by_class_student}


def _match_contact_for_report(report: dict[str, Any], indexes: dict[str, dict]) -> dict[str, Any] | None:
    student_key = _norm_cjk(_extract_report_student_name(report))
    class_key = _norm_cjk(str(report.get("class_name") or ""))
    if not student_key:
        return None
    if class_key:
        matched = indexes.get("by_class_student", {}).get((class_key, student_key))
        if matched is not None:
            return matched
    return indexes.get("by_student", {}).get(student_key)


def _latest_reports_by_student(reports: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str], dict[str, Any]] = {}
    for report in reports:
        student_key = _norm_cjk(_extract_report_student_name(report))
        if not student_key:
            continue
        class_key = _norm_cjk(str(report.get("class_name") or ""))
        key = (class_key, student_key)
        current = grouped.get(key)
        if current is None or float(report.get("timestamp") or 0.0) > float(current.get("timestamp") or 0.0):
            grouped[key] = report
    return sorted(grouped.values(), key=lambda x: float(x.get("timestamp") or 0.0), reverse=True)


def _html_to_image_path(html_path: str, report_id: str) -> str:
    p = Path(str(html_path or ""))
    if p.suffix.lower() == ".html":
        return str(p.with_suffix(".png"))
    if report_id:
        return str(p / f"{report_id}.png") if p.exists() and p.is_dir() else str(p.with_name(f"{report_id}.png"))
    return str(p)


def _parse_exported_report_filename(filename: str) -> tuple[str, str]:
    stem = Path(str(filename or "")).stem
    lowered = stem.lower()
    for suffix in ("_report", "+report", "-report", " report"):
        if lowered.endswith(suffix):
            stem = stem[: -len(suffix)]
            break
    parts = [part.strip() for part in stem.replace("+", "_").split("_") if part.strip()]
    if not parts:
        return "", ""
    student_name = parts[0]
    class_name = ""
    for part in parts[1:]:
        if "班" in part:
            class_name = part
            break
    return student_name, class_name


def _build_exported_image_index(export_images_dir: Path | None) -> dict[str, dict[Any, Path]]:
    by_student: dict[str, tuple[float, Path]] = {}
    by_class_student: dict[tuple[str, str], tuple[float, Path]] = {}
    root = Path(export_images_dir) if export_images_dir else DEFAULT_EXPORT_IMAGES_DIR
    if not root.exists():
        return {"by_student": {}, "by_class_student": {}}
    for path in root.rglob("*_report.png"):
        student_name, class_name = _parse_exported_report_filename(path.name)
        student_key = _norm_cjk(student_name)
        class_key = _norm_cjk(class_name)
        if not student_key:
            continue
        mtime = path.stat().st_mtime
        current = by_student.get(student_key)
        if current is None or mtime > current[0]:
            by_student[student_key] = (mtime, path)
        if class_key:
            current = by_class_student.get((class_key, student_key))
            if current is None or mtime > current[0]:
                by_class_student[(class_key, student_key)] = (mtime, path)
    return {
        "by_student": {k: v[1] for k, v in by_student.items()},
        "by_class_student": {k: v[1] for k, v in by_class_student.items()},
    }


def _resolve_exported_image_path(report: dict[str, Any], export_index: dict[str, dict[Any, Path]]) -> str:
    student_key = _norm_cjk(_extract_report_student_name(report))
    class_key = _norm_cjk(str(report.get("class_name") or ""))
    if not student_key:
        return ""
    if class_key:
        matched = export_index.get("by_class_student", {}).get((class_key, student_key))
        if matched:
            return str(matched)
    matched = export_index.get("by_student", {}).get(student_key)
    return str(matched) if matched else ""


def _resolve_report_class_name(report: dict[str, Any], contact: dict[str, Any], image_path: str) -> str:
    report_class = str(report.get("class_name") or "").strip()
    if report_class:
        return report_class
    if image_path:
        _student_name, exported_class = _parse_exported_report_filename(Path(image_path).name)
        if exported_class:
            return exported_class
    return str(contact.get("_class_name") or "").strip()


def _build_rows_from_reports(
    reports: list[dict[str, Any]],
    contacts_rows: list[dict[str, str]],
    export_images_dir: Path | None = None,
) -> tuple[list[dict[str, str]], dict[str, int]]:
    indexes = _build_contact_indexes(contacts_rows)
    export_index = _build_exported_image_index(export_images_dir)
    latest_reports = _latest_reports_by_student(reports)
    rows: list[dict[str, str]] = []
    matched = 0
    unmatched = 0
    for report in latest_reports:
        contact = _match_contact_for_report(report, indexes)
        if contact is None:
            unmatched += 1
            continue
        matched += 1
        report_id = str(report.get("id") or "").strip()
        image_path = _resolve_exported_image_path(report, export_index) or _html_to_image_path(
            str(report.get("html_path") or report.get("path") or ""),
            report_id,
        )
        row = {
            "parent_name": str(contact.get("姓名") or "").strip(),
            "student_name": str(_extract_report_student_name(report) or contact.get("_student_name") or "").strip(),
            "class_name": _resolve_report_class_name(report, contact, image_path),
            "wecom_user_id": str(contact.get("账号") or "").strip(),
            "report_submission_id": report_id,
            "image_path": image_path,
        }
        rows.append(row)
    stats = {"matched": matched, "unmatched": unmatched, "deduped_reports": len(latest_reports)}
    return rows, stats


def _prepare_report_driven_rows(
    reports: list[dict[str, Any]],
    contacts_rows: list[dict[str, str]],
    export_images_dir: Path | None = None,
) -> list[dict[str, str]]:
    rows, _stats = _build_rows_from_reports(reports, contacts_rows, export_images_dir=export_images_dir)
    return rows


def _xlsx_col_to_index(col_ref: str) -> int:
    value = 0
    for ch in col_ref:
        if 'A' <= ch <= 'Z':
            value = value * 26 + (ord(ch) - ord('A') + 1)
    return max(value - 1, 0)


def _read_contacts_xlsx_stdlib(path: Path) -> list[dict[str, str]]:
    ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    name_key = "姓名"
    user_key = "账号"
    dept_key = "部门"
    with zipfile.ZipFile(path) as zf:
        shared = []
        if 'xl/sharedStrings.xml' in zf.namelist():
            root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
            for si in root.findall('a:si', ns):
                text = ''.join(t.text or '' for t in si.findall('.//a:t', ns))
                shared.append(text)

        wb = ET.fromstring(zf.read('xl/workbook.xml'))
        rels = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
        rel_map = {rel.attrib.get('Id'): rel.attrib.get('Target') for rel in rels}
        first_sheet = wb.find('a:sheets/a:sheet', ns)
        if first_sheet is None:
            return []
        rid = first_sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        target = rel_map.get(rid, 'worksheets/sheet1.xml')
        sheet_path = 'xl/' + target.lstrip('/')
        sheet = ET.fromstring(zf.read(sheet_path))

        rows = []
        for row in sheet.findall('.//a:sheetData/a:row', ns):
            vals = {}
            for c in row.findall('a:c', ns):
                ref = c.attrib.get('r', '')
                col = ''.join(ch for ch in ref if ch.isalpha())
                idx = _xlsx_col_to_index(col)
                cell_type = c.attrib.get('t')
                v = c.find('a:v', ns)
                value = ''
                if v is not None and v.text is not None:
                    value = v.text
                    if cell_type == 's':
                        try:
                            value = shared[int(value)]
                        except Exception:
                            pass
                is_node = c.find('a:is', ns)
                if is_node is not None:
                    value = ''.join(t.text or '' for t in is_node.findall('.//a:t', ns))
                vals[idx] = str(value).strip()
            if vals:
                max_idx = max(vals.keys())
                rows.append([vals.get(i, '') for i in range(max_idx + 1)])

    header_row = 0
    for i, row in enumerate(rows):
        if name_key in row and user_key in row:
            header_row = i
            break
    headers = rows[header_row] if rows else []
    header_map = {name: idx for idx, name in enumerate(headers) if name}
    out = []
    for row in rows[header_row + 1:]:
        name = row[header_map[name_key]].strip() if name_key in header_map and header_map[name_key] < len(row) else ''
        user_id = row[header_map[user_key]].strip() if user_key in header_map and header_map[user_key] < len(row) else ''
        dept = row[header_map[dept_key]].strip() if dept_key in header_map and header_map[dept_key] < len(row) else ''
        if name and user_id:
            out.append({name_key: name, user_key: user_id, dept_key: dept})
    return out


def _read_contacts_xlsx(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    if pd is not None:
        try:
            df = pd.read_excel(path, sheet_name=0, header=None)
            header_row = 0
            for i in range(len(df)):
                row = [str(x or '').strip() for x in df.iloc[i].tolist()]
                if '??' in row and '??' in row:
                    header_row = i
                    break
            real = pd.read_excel(path, sheet_name=0, header=header_row)
            cols = [c for c in ['??', '??', '??'] if c in real.columns]
            if len(cols) >= 2:
                return real[cols].fillna('').to_dict(orient='records')
        except Exception:
            pass

    try:
        return _read_contacts_xlsx_stdlib(path)
    except Exception:
        return []


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


def _delta_row_key(row: dict[str, str]) -> tuple[str, str]:
    return (str(row.get("class_name") or "").strip(), str(row.get("student_name") or "").strip())


def _state_row_key(row: dict[str, str]) -> tuple[str, str, str]:
    return (
        str(row.get("class_name") or "").strip(),
        str(row.get("student_name") or "").strip(),
        str(row.get("report_submission_id") or "").strip(),
    )


def _build_delta_rows(new_rows: list[dict[str, str]], previous_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    if not previous_rows:
        return list(new_rows)
    previous_by_key = {_delta_row_key(row): str(row.get("report_submission_id") or "").strip() for row in previous_rows}
    delta_rows: list[dict[str, str]] = []
    for row in new_rows:
        key = _delta_row_key(row)
        report_id = str(row.get("report_submission_id") or "").strip()
        if previous_by_key.get(key) != report_id:
            delta_rows.append(row)
    return delta_rows


def _build_pending_rows(new_rows: list[dict[str, str]], state_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    handled_keys = {
        _state_row_key(row)
        for row in state_rows
        if str(row.get("status") or "").strip() in HANDLED_SEND_STATUSES
    }
    if not handled_keys:
        return list(new_rows)
    pending_rows: list[dict[str, str]] = []
    for row in new_rows:
        if _state_row_key(row) not in handled_keys:
            pending_rows.append(row)
    return pending_rows


def main() -> int:
    parser = argparse.ArgumentParser(description="Enrich WeCom RPA tasks CSV with score-reading phoneme video messages.")
    parser.add_argument("--input-csv", default="tools/rpa_tasks.real.csv", help="Base tasks CSV")
    parser.add_argument("--output-csv", default="tools/rpa_tasks.with-message.csv", help="Output CSV with message_text")
    parser.add_argument("--delta-output-csv", default=_default_delta_output_csv(), help="Output CSV containing only new/updated tasks")
    parser.add_argument("--pending-output-csv", default=_default_pending_output_csv(), help="Output CSV containing current tasks not yet handled")
    parser.add_argument("--state-csv", default=_default_state_csv(), help="CSV ledger of handled send-state rows")
    parser.add_argument("--score-api-base", default="http://127.0.0.1:8000", help="Score-reading API base URL")
    parser.add_argument("--timeout-sec", type=float, default=12.0, help="HTTP timeout")
    parser.add_argument("--top-n", type=int, default=3, help="Weak phoneme count")
    parser.add_argument("--per-phoneme", type=int, default=2, help="Videos per phoneme in source payload")
    parser.add_argument("--max-links", type=int, default=3, help="Max links in integrated message")
    parser.add_argument("--report-driven", action="store_true", help="Build tasks from reports + contacts instead of input CSV")
    parser.add_argument("--contacts-xlsx", default="/data/contacts.xlsx", help="Contacts Excel for report-driven mode")
    parser.add_argument("--export-images-dir", default=str(DEFAULT_EXPORT_IMAGES_DIR), help="Directory containing exported report PNG files")
    args = parser.parse_args()

    try:
        reports = _load_reports(args.score_api_base, timeout=float(args.timeout_sec))
    except Exception as exc:
        print(f"[ERR] cannot fetch reports from score api: {exc}")
        return 3
    if not reports:
        print("[ERR] report list is empty, cannot enrich tasks.")
        return 4

    output_path = Path(args.output_csv)
    previous_rows: list[dict[str, str]] = []
    previous_fieldnames: list[str] = []
    if output_path.exists():
        previous_fieldnames, previous_rows = _read_csv(output_path)

    if args.report_driven:
        contacts_path = Path(args.contacts_xlsx)
        contacts_rows = _read_contacts_xlsx(contacts_path)
        if not contacts_rows:
            print(f"[ERR] contacts xlsx not found or unreadable: {contacts_path}")
            return 2
        rows, report_stats = _build_rows_from_reports(
            reports,
            contacts_rows,
            export_images_dir=Path(args.export_images_dir),
        )
        fieldnames = ["parent_name", "student_name", "class_name", "wecom_user_id", "report_submission_id", "image_path"]
        if not rows:
            print(f"[ERR] no matched report rows generated. stats={report_stats}")
            return 5
    else:
        input_path = Path(args.input_csv)
        if not input_path.exists():
            print(f"[ERR] input csv not found: {input_path}")
            return 2
        fieldnames, rows = _read_csv(input_path)
        if not rows:
            print("[ERR] input csv has no rows.")
            return 5
        if "parent_name" not in fieldnames or "image_path" not in fieldnames:
            print("[ERR] csv must include parent_name,image_path")
            return 6

    extra_fields = ["report_submission_id", "focus_word_phonemes", "weak_phonemes", "video_links", "message_text", "message_source"]
    for col in extra_fields:
        if col not in fieldnames:
            fieldnames.append(col)

    ok_count = 0
    miss_count = 0
    err_count = 0

    for row in rows:
        for col in extra_fields:
            row[col] = str(row.get(col) or "").strip()

        report_id = str(row.get("report_submission_id") or "").strip() or _pick_report_id(row, reports)
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

        focus_word_phonemes = payload.get("focus_word_phonemes") or []
        weak_phonemes = payload.get("weak_phonemes") or []
        links = payload.get("links") or []
        if isinstance(focus_word_phonemes, list):
            row["focus_word_phonemes"] = ",".join(
                str(x or "").strip() for x in focus_word_phonemes if str(x or "").strip()
            )
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

    _write_csv(output_path, fieldnames, rows)
    delta_rows = _build_delta_rows(rows, previous_rows)
    delta_output_path = Path(args.delta_output_csv)
    delta_fieldnames = previous_fieldnames or fieldnames
    for col in fieldnames:
        if col not in delta_fieldnames:
            delta_fieldnames.append(col)
    _write_csv(delta_output_path, delta_fieldnames, delta_rows)
    state_rows: list[dict[str, str]] = []
    state_path = Path(args.state_csv)
    if state_path.exists():
        _state_fieldnames, state_rows = _read_csv(state_path)
    pending_rows = _build_pending_rows(rows, state_rows)
    pending_output_path = Path(args.pending_output_csv)
    pending_fieldnames = previous_fieldnames or fieldnames
    for col in fieldnames:
        if col not in pending_fieldnames:
            pending_fieldnames.append(col)
    _write_csv(pending_output_path, pending_fieldnames, pending_rows)

    print(f"[OK] output: {output_path}")
    print(f"[OK] delta_output: {delta_output_path}")
    print(f"[OK] pending_output: {pending_output_path}")
    print(f"[OK] total={len(rows)} enriched={ok_count} no_match_or_empty={miss_count} errors={err_count} delta={len(delta_rows)} pending={len(pending_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


