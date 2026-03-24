# RPA Send State Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add a CSV-backed send-state ledger so refresh builds a pending task CSV and RPA runs default to rows not yet handled.

**Architecture:** Keep the existing full and delta snapshots intact. Add a new pending snapshot computed from the current full snapshot minus handled task keys stored in `run-logs/rpa_send_state.csv`. Update sender to append/update that ledger from real task results and switch defaults from delta to pending.

**Tech Stack:** Python, CSV, FastAPI helpers, existing RPA sender script, pytest.

---

### Task 1: Add failing tests for pending/state helpers

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_build_rpa_tasks_from_score_refresh.py`
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\tests\test_rpa_refresh.py`

**Steps:**
1. Add tests for default pending/state paths.
2. Add tests for filtering pending rows from handled state keys.
3. Add tests for parsing refresh summary with `pending=`.
4. Run targeted tests and verify failure.

### Task 2: Add pending/state support to refresh builder

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py`

**Steps:**
1. Add helpers for default pending/state paths.
2. Add CSV read/write helpers for state rows.
3. Add handled-status filtering logic keyed by `class_name + student_name + report_submission_id`.
4. Add CLI args `--pending-output-csv` and `--state-csv`.
5. Write pending CSV and extend stdout summary with `pending=`.
6. Run targeted tests and make them pass.

### Task 3: Persist send state from sender results

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_sender.py`

**Steps:**
1. Add state row key and merge helpers.
2. Add `--state-csv` CLI arg.
3. Convert current run result rows into state rows.
4. Merge with existing state CSV, replacing by stable key.
5. Write state CSV at end of run.
6. Run targeted tests or py_compile.

### Task 4: Switch refresh/start defaults from delta to pending

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_refresh.py`
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_runner.py`
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\schema.py`
- Optional follow-up if needed: local/web defaults in GUI/UI files

**Steps:**
1. Update refresh command builder to pass pending/state output paths.
2. Update summary parser to include `pending`.
3. Change default `tasks_csv` from `tools/rpa_tasks.delta.csv` to `tools/rpa_tasks.pending.csv`.
4. Update tests accordingly.

### Task 5: Verify end-to-end behavior

**Files:**
- Verify outputs only

**Steps:**
1. Run targeted pytest suites.
2. Run `py_compile` on modified Python files.
3. Optionally run refresh script once and verify full/delta/pending/state files exist.
4. Summarize how to use pending mode for multi-batch pushes.
