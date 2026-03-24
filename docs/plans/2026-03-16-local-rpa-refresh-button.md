# Local RPA Refresh Button Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add a local desktop `刷新发送任务` button that rebuilds `tools/rpa_tasks.with-message.csv` from reports and contacts, then makes that CSV the active send file.

**Architecture:** Keep the desktop GUI thin. Reuse the existing report-driven rebuild script and only add GUI-side command construction, background execution, summary parsing, and UI state updates. Avoid duplicating report matching logic in the local client.

**Tech Stack:** Python, Tkinter, subprocess, threading, pytest

---

### Task 1: Add testable refresh helpers

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py`
- Test: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py`

**Step 1: Write the failing test**

Add tests for:
- default CSV path is `tools/rpa_tasks.with-message.csv`
- refresh command contains report-driven arguments
- refresh summary parsing returns the expected counters

**Step 2: Run test to verify it fails**

Run:
```bash
python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py -q
```

Expected:
- FAIL because helpers do not exist yet

**Step 3: Write minimal implementation**

In `tools/wecom_rpa_gui.py`, add pure helpers:
- `build_refresh_tasks_command(repo_root: Path) -> list[str]`
- `parse_refresh_tasks_summary(output: str) -> dict[str, int]`

Also switch the GUI default CSV path to `tools/rpa_tasks.with-message.csv`.

**Step 4: Run test to verify it passes**

Run:
```bash
python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py -q
```

Expected:
- PASS

**Step 5: Commit**

```bash
git add tools/wecom_rpa_gui.py tools/tests/test_wecom_rpa_gui_refresh.py
git commit -m "test: cover local refresh command helpers"
```

### Task 2: Add refresh button and background execution

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py`
- Test: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py`

**Step 1: Write the failing test**

Add tests for GUI-side refresh behavior:
- success path updates `csv_path` to `tools/rpa_tasks.with-message.csv`
- success path updates status text with parsed counters

Keep the tests focused on a small method that processes refresh results instead of full Tk rendering.

**Step 2: Run test to verify it fails**

Run:
```bash
python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py -q
```

Expected:
- FAIL because refresh result handling does not exist yet

**Step 3: Write minimal implementation**

In `tools/wecom_rpa_gui.py`:
- add a refresh button to Step 1
- add methods:
  - `refresh_tasks()`
  - worker-thread wrapper that runs the refresh command
  - a small result handler that updates:
    - `csv_path`
    - `status_text`
    - log pane
    - task inspection
- disable refresh while it is running

**Step 4: Run test to verify it passes**

Run:
```bash
python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py -q
```

Expected:
- PASS

**Step 5: Commit**

```bash
git add tools/wecom_rpa_gui.py tools/tests/test_wecom_rpa_gui_refresh.py
git commit -m "feat: add local refresh button for report-driven tasks"
```

### Task 3: Verify local integration

**Files:**
- Modify: `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py` (only if needed)

**Step 1: Run targeted tests**

Run:
```bash
python -m pytest C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\test_wecom_rpa_gui_refresh.py -q
```

Expected:
- PASS

**Step 2: Run syntax verification**

Run:
```bash
python -m py_compile C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py
```

Expected:
- no output

**Step 3: Run the refresh command once for real**

Run:
```bash
python C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py --report-driven --contacts-xlsx C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\data\contacts.xlsx --output-csv C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\rpa_tasks.with-message.csv --score-api-base http://127.0.0.1:8000
```

Expected:
- refresh summary with non-error counters

**Step 4: Optional manual GUI smoke test**

Run:
```bash
python C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py
```

Verify:
- button is visible
- click `刷新发送任务`
- status/log updates
- `任务 CSV` field points to `tools/rpa_tasks.with-message.csv`

**Step 5: Commit**

```bash
git add tools/wecom_rpa_gui.py tools/tests/test_wecom_rpa_gui_refresh.py
git commit -m "chore: verify local refresh workflow"
```
