# Exported Image Path Matching Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Make report-driven task refresh prefer real exported report PNGs from the score-reading output directory instead of assuming `<submission_id>.png` beside report HTML.

**Architecture:** Extend the report-driven task builder to scan a configurable export-image directory, parse exported report filenames into student/class keys, and resolve `image_path` from that index before falling back to the current derived-path logic. Keep all existing report matching, contact matching, and message enrichment behavior intact.

**Tech Stack:** Python 3.12, pytest, existing WeCom sender CLI tooling.

---

### Task 1: Add failing tests for exported image matching

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py`

**Step 1: Write the failing tests**

Add tests covering:
- exported filename parsing for `学生_班级_单元_report.png`
- report-driven refresh choosing a real exported PNG over derived `<submission_id>.png`
- fallback to derived path when no exported PNG exists

**Step 2: Run test to verify it fails**

Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py -q`

Expected: FAIL due to missing exported-image matching behavior.

**Step 3: Commit**

Do not commit yet. Move to implementation after confirming RED.

### Task 2: Implement exported image indexing and matching

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py`

**Step 1: Add exported image helpers**

Implement minimal helpers to:
- normalize names/classes for matching
- parse filename stems ending in `_report` or `+report`
- build an exported image index from `--export-images-dir`
- resolve best image path for a report using class+student first, student-only second, newest file as tie-breaker

**Step 2: Wire exported image matching into report-driven row generation**

Update report-driven generation so `image_path` uses:
1. exported image match
2. fallback `_html_to_image_path(...)`

**Step 3: Add CLI option**

Add `--export-images-dir` with a practical default pointing at the score-reading output directory used in this environment.

**Step 4: Run targeted tests**

Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py -q`

Expected: PASS.

### Task 3: Verify refresh entrypoints still work

**Files:**
- Inspect/modify if needed: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py`
- Inspect/modify if needed: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_refresh.py`

**Step 1: Confirm existing entrypoints can rely on the new default**

Only change these files if they need to explicitly pass `--export-images-dir`. Prefer no change if the default is sufficient.

**Step 2: Run refresh-related tests**

Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\tests\test_rpa_refresh.py -q`

Expected: PASS.

### Task 4: Validate with real report-driven generation

**Files:**
- No source change required unless verification fails

**Step 1: Run report-driven rebuild against real data**

Run: `python C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py --report-driven --contacts-xlsx C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\data\contacts.xlsx --output-csv C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\rpa_tasks.with-message.csv --score-api-base http://127.0.0.1:8000`

Expected: summary with non-zero total and no `Skip missing image` root cause when corresponding exported PNGs exist.

**Step 2: Inspect generated CSV sample**

Confirm several `image_path` values point at real exported PNG files under `D:\score_reading_fresh\data\out`.

### Task 5: Final verification

**Files:**
- No source change required unless verification fails

**Step 1: Python syntax check**

Run:
- `python -m py_compile C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py`
- include any additional modified Python files if changed

**Step 2: Summarize outcomes**

Report:
- which tests passed
- whether real CSV now points at exported PNGs
- any remaining gaps
