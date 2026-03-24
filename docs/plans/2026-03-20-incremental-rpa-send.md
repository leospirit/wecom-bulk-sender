# Incremental RPA Task Refresh Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Generate a delta task CSV during refresh and allow local/web RPA runs to target either the full task list or only newly added/updated reports.

**Architecture:** Keep `rpa_tasks.with-message.csv` as the canonical full snapshot. During each refresh, compare the newly built snapshot against the previous snapshot by `class_name + student_name`; write rows with new keys or changed `report_submission_id` to `rpa_tasks.delta.csv`. Surface the send scope in both GUI and web by switching the active `tasks_csv` path between the full and delta CSVs instead of changing sender semantics.

**Tech Stack:** Python CLI + FastAPI backend + Tkinter GUI + React web UI + pytest

---

### Task 1: Add delta snapshot generation to refresh script

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py`
- Test: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py`

**Step 1: Write the failing tests**
- Add tests for:
  - previous full CSV missing => delta contains all rows
  - matching `class_name + student_name` with same `report_submission_id` => excluded from delta
  - matching `class_name + student_name` with changed `report_submission_id` => included in delta
  - delta file path defaults to `tools/rpa_tasks.delta.csv`

**Step 2: Run test to verify it fails**
Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py -q`
Expected: FAIL on missing delta helpers/CLI behavior.

**Step 3: Write minimal implementation**
- Add helpers for row keying and delta comparison.
- Add optional `--delta-output-csv` parameter with default `tools/rpa_tasks.delta.csv`.
- Before overwriting full CSV, load the existing full CSV if present.
- Write full CSV as today, then write delta CSV containing only new/updated rows.
- Extend stdout summary with `delta=` count.

**Step 4: Run test to verify it passes**
Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py -q`
Expected: PASS.

### Task 2: Wire delta paths into local GUI refresh + start flow

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py`
- Test: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py`

**Step 1: Write the failing tests**
- Add tests for:
  - default delta CSV path helper returns `tools/rpa_tasks.delta.csv`
  - refresh command includes `--delta-output-csv`
  - refresh finalize leaves active CSV on delta CSV when send scope is `delta`
  - send scope switching updates chosen CSV path between full and delta

**Step 2: Run test to verify it fails**
Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py -q`
Expected: FAIL on missing delta support.

**Step 3: Write minimal implementation**
- Add `default_delta_csv_path()` helper.
- Add send-scope state with values `all` / `delta`, default `delta`.
- Pass `--delta-output-csv` in refresh command.
- After refresh, set active `csv_path` according to selected scope.
- Add a small UI selector for send scope and keep `csv_path` synchronized.

**Step 4: Run test to verify it passes**
Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py -q`
Expected: PASS.

### Task 3: Wire delta paths into API refresh and web UI

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_refresh.py`
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\schema.py`
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_runner.py`
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\web\src\App.tsx`
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\static\rpa-lite\app.js`
- Test: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\tests\test_rpa_refresh.py`

**Step 1: Write the failing tests**
- Add tests for:
  - API refresh command includes `--delta-output-csv`
  - parsed refresh summary exposes `delta`
  - start request defaults to delta CSV path
  - runner resolves default tasks CSV to delta CSV when no path is given

**Step 2: Run test to verify it fails**
Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\tests\test_rpa_refresh.py -q`
Expected: FAIL.

**Step 3: Write minimal implementation**
- Add `delta_output` path handling in API refresh command builder and summary parser.
- Change API/web/rpa-lite defaults to `tools/rpa_tasks.delta.csv` for active send runs.
- Preserve manual override of `tasks_csv` so users can still choose full CSV.
- Add web UI send-scope selector that switches `tasks_csv` between full and delta defaults.

**Step 4: Run test to verify it passes**
Run: `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\tests\test_rpa_refresh.py -q`
Expected: PASS.

### Task 4: Verify end-to-end refresh output and build health

**Files:**
- Verify only; no planned code files

**Step 1: Run targeted test suites**
Run:
- `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py -q`
- `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py -q`
- `python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\tests\test_rpa_refresh.py -q`
Expected: all PASS.

**Step 2: Run syntax/build verification**
Run:
- `python -m py_compile C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_refresh.py C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_runner.py C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\schema.py`
- `npm --prefix C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\web run build`
Expected: success.

**Step 3: Run real refresh command**
Run:
`python C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py --report-driven --contacts-xlsx C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\data\contacts.xlsx --output-csv C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\rpa_tasks.with-message.csv --delta-output-csv C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\rpa_tasks.delta.csv --score-api-base http://localhost`
Expected: success with `delta=` count in stdout and both CSVs created.

**Step 4: Sanity-check delta behavior**
- Confirm `rpa_tasks.delta.csv` exists.
- Confirm when no report changed, delta CSV is empty/header-only.
- Confirm when a report changes, delta CSV only contains affected student rows.
