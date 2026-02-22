from __future__ import annotations

import sqlite3
from pathlib import Path
from datetime import datetime

DB_PATH = Path("/data/app.db")


def _conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    return sqlite3.connect(DB_PATH)


def init_db():
    with _conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT UNIQUE,
                student_name TEXT,
                parent_name TEXT,
                user_id TEXT,
                status TEXT,
                error TEXT,
                created_at TEXT,
                updated_at TEXT
            )
            """
        )
        conn.commit()


def insert_task(task: dict) -> int | None:
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        try:
            cur = conn.execute(
                """
                INSERT INTO tasks (file_path, student_name, parent_name, user_id, status, error, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    task["file_path"],
                    task.get("student_name", ""),
                    task.get("parent_name", ""),
                    task.get("user_id"),
                    task.get("status", "pending"),
                    task.get("error"),
                    now,
                    now,
                ),
            )
            conn.commit()
            return cur.lastrowid
        except sqlite3.IntegrityError:
            return None


def list_tasks() -> list[dict]:
    with _conn() as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute("SELECT * FROM tasks ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]


def update_task_status(task_id: int, status: str, error: str | None = None) -> None:
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        conn.execute(
            "UPDATE tasks SET status=?, error=?, updated_at=? WHERE id=?",
            (status, error, now, task_id),
        )
        conn.commit()


def set_tasks_status(ids: list[int], status: str) -> None:
    if not ids:
        return
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        conn.executemany(
            "UPDATE tasks SET status=?, updated_at=? WHERE id=?",
            [(status, now, i) for i in ids],
        )
        conn.commit()


def fetch_tasks_by_status(status: str, limit: int = 50) -> list[dict]:
    with _conn() as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT * FROM tasks WHERE status=? ORDER BY id ASC LIMIT ?",
            (status, limit),
        ).fetchall()
        return [dict(r) for r in rows]


def set_all_pending_to_queued() -> int:
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        cur = conn.execute(
            "UPDATE tasks SET status=?, updated_at=? WHERE status='pending'",
            ("queued", now),
        )
        conn.commit()
        return cur.rowcount


def status_counts() -> dict:
    with _conn() as conn:
        rows = conn.execute(
            "SELECT status, COUNT(*) as c FROM tasks GROUP BY status"
        ).fetchall()
        counts = {"total": 0, "pending": 0, "queued": 0, "sending": 0, "sent": 0, "failed": 0, "skipped": 0}
        for status, c in rows:
            counts[status] = c
        counts["total"] = sum(v for k, v in counts.items() if k != "total")
        return counts
