# WeCom Bulk Image Sender Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a Dockerized web app that scans folders of student image files, maps to “学生名+妈妈” users in WeCom, and sends images with queue monitoring.

**Architecture:** React (Vite) frontend talks to a FastAPI backend. Backend reads Excel contacts, scans folders, queues jobs in SQLite, sends images via WeCom APIs, and exposes status APIs. Docker Compose runs web + api with a shared data volume for inputs/logs.

**Tech Stack:** Python 3.12, FastAPI, SQLite, Pandas, Pydantic, React + Vite, Docker Compose.

---

### Task 1: Scaffold repository and Docker Compose

**Files:**
- Create: `docker-compose.yml`
- Create: `api/Dockerfile`
- Create: `web/Dockerfile`
- Create: `README.md`

**Step 1: Write the failing test**
- N/A (scaffold)

**Step 2: Run test to verify it fails**
- N/A

**Step 3: Write minimal implementation**
- Compose with `api` (FastAPI) and `web` (Vite) services.
- Create shared volume mount `./data:/data`.
- Expose `api` on 8000 and `web` on 5173 (or 80 via nginx).

**Step 4: Run test to verify it passes**
- N/A

**Step 5: Commit**
```bash
git add docker-compose.yml api/Dockerfile web/Dockerfile README.md
git commit -m "chore: scaffold docker compose"
```

### Task 2: Backend config and health endpoint

**Files:**
- Create: `api/app/main.py`
- Create: `api/app/config.py`
- Create: `api/app/models.py`
- Create: `api/tests/test_health.py`

**Step 1: Write the failing test**
```python
from fastapi.testclient import TestClient
from app.main import app

def test_health():
    client = TestClient(app)
    r = client.get("/api/health")
    assert r.status_code == 200
    assert r.json()["status"] == "ok"
```

**Step 2: Run test to verify it fails**
Run: `pytest api/tests/test_health.py -v`
Expected: FAIL (app not defined)

**Step 3: Write minimal implementation**
- `config.yaml` load (corpId/agentId/secret, rate limits, worker settings, root dir default `/data/inbox`).
- `GET /api/health` returns `{status: "ok"}`.

**Step 4: Run test to verify it passes**
Run: `pytest api/tests/test_health.py -v`
Expected: PASS

**Step 5: Commit**
```bash
git add api/app/main.py api/app/config.py api/app/models.py api/tests/test_health.py
git commit -m "feat: add backend config and health endpoint"
```

### Task 3: Excel parsing and name matching

**Files:**
- Create: `api/app/contacts.py`
- Create: `api/tests/test_contacts.py`

**Step 1: Write the failing test**
```python
from app.contacts import build_name_index, extract_student_name

def test_extract_student_name():
    assert extract_student_name("刘家骏6单元背诵_report") == "刘家骏"


def test_build_name_index():
    rows = [
        {"姓名": "刘家骏妈妈", "账号": "u1"},
        {"姓名": "刘家骏爸爸", "账号": "u2"},
    ]
    idx = build_name_index(rows)
    assert idx["刘家骏妈妈"] == "u1"
```

**Step 2: Run test to verify it fails**
Run: `pytest api/tests/test_contacts.py -v`
Expected: FAIL (functions missing)

**Step 3: Write minimal implementation**
- Parse Excel `成员列表` starting at row with headers, read `姓名` and `账号`.
- `extract_student_name` returns leading CJK characters.
- `build_name_index` returns dict of name -> userId.

**Step 4: Run test to verify it passes**
Run: `pytest api/tests/test_contacts.py -v`
Expected: PASS

**Step 5: Commit**
```bash
git add api/app/contacts.py api/tests/test_contacts.py
git commit -m "feat: parse contacts and extract student name"
```

### Task 4: File scanning and task creation

**Files:**
- Create: `api/app/scanner.py`
- Create: `api/app/db.py`
- Create: `api/app/schema.py`
- Create: `api/tests/test_scanner.py`

**Step 1: Write the failing test**
```python
from app.scanner import build_tasks

def test_build_tasks(tmp_path):
    f = tmp_path / "刘家骏6单元背诵_report.png"
    f.write_bytes(b"x")
    name_index = {"刘家骏妈妈": "u1"}
    tasks = build_tasks(str(tmp_path), name_index)
    assert tasks[0].user_id == "u1"
```

**Step 2: Run test to verify it fails**
Run: `pytest api/tests/test_scanner.py -v`
Expected: FAIL

**Step 3: Write minimal implementation**
- Recursive scan folder.
- Filter image extensions.
- Build tasks with status `pending` and fields (path, student, parent, userId).
- SQLite schema for tasks.

**Step 4: Run test to verify it passes**
Run: `pytest api/tests/test_scanner.py -v`
Expected: PASS

**Step 5: Commit**
```bash
git add api/app/scanner.py api/app/db.py api/app/schema.py api/tests/test_scanner.py
git commit -m "feat: scan files and create tasks"
```

### Task 5: WeCom client and send worker

**Files:**
- Create: `api/app/wecom.py`
- Create: `api/app/worker.py`
- Create: `api/tests/test_wecom.py`

**Step 1: Write the failing test**
```python
from app.wecom import WeComClient

def test_build_token_url():
    c = WeComClient(corp_id="c", secret="s")
    assert "gettoken" in c.token_url
```

**Step 2: Run test to verify it fails**
Run: `pytest api/tests/test_wecom.py -v`
Expected: FAIL

**Step 3: Write minimal implementation**
- Token fetch with caching.
- Media upload for image.
- Message send to single `userId`.
- Worker reads pending tasks and updates status with retry/limit.

**Step 4: Run test to verify it passes**
Run: `pytest api/tests/test_wecom.py -v`
Expected: PASS

**Step 5: Commit**
```bash
git add api/app/wecom.py api/app/worker.py api/tests/test_wecom.py
git commit -m "feat: wecom client and send worker"
```

### Task 6: Backend API for UI

**Files:**
- Create: `api/app/routes.py`
- Modify: `api/app/main.py`
- Create: `api/tests/test_api.py`

**Step 1: Write the failing test**
```python
from fastapi.testclient import TestClient
from app.main import app

def test_list_tasks():
    client = TestClient(app)
    r = client.get("/api/tasks")
    assert r.status_code == 200
```

**Step 2: Run test to verify it fails**
Run: `pytest api/tests/test_api.py -v`
Expected: FAIL

**Step 3: Write minimal implementation**
- Endpoints: load contacts, scan, list tasks, start/stop worker, toggle auto-watch, send selected.
- Return status counts.

**Step 4: Run test to verify it passes**
Run: `pytest api/tests/test_api.py -v`
Expected: PASS

**Step 5: Commit**
```bash
git add api/app/routes.py api/app/main.py api/tests/test_api.py
git commit -m "feat: api endpoints for ui"
```

### Task 7: Frontend UI

**Files:**
- Create: `web/src/App.tsx`
- Create: `web/src/api.ts`
- Create: `web/src/components/TaskTable.tsx`
- Create: `web/src/components/Controls.tsx`
- Create: `web/src/styles.css`

**Step 1: Write the failing test**
- N/A (UI)

**Step 2: Run test to verify it fails**
- N/A

**Step 3: Write minimal implementation**
- Controls: choose root path, upload contacts file, scan, start batch send, auto-watch toggle.
- Table: task list, checkbox selection, “send selected” button, status chips.

**Step 4: Run test to verify it passes**
- N/A

**Step 5: Commit**
```bash
git add web/src/App.tsx web/src/api.ts web/src/components web/src/styles.css
git commit -m "feat: build frontend ui"
```

### Task 8: Documentation and smoke test

**Files:**
- Modify: `README.md`

**Step 1: Write the failing test**
- N/A

**Step 2: Run test to verify it fails**
- N/A

**Step 3: Write minimal implementation**
- Add Windows/Mac quickstart.
- Explain volume mount and folder path.
- Describe “batch vs selected send”.

**Step 4: Run test to verify it passes**
- N/A

**Step 5: Commit**
```bash
git add README.md
git commit -m "docs: add usage and setup"
```
