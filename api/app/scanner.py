from __future__ import annotations

import os
from pathlib import Path
from .contacts import extract_student_name

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}


def scan_files(root_path: str) -> list[str]:
    root = Path(root_path)
    if not root.exists():
        return []
    files = []
    for dirpath, _, filenames in os.walk(root):
        for name in filenames:
            p = Path(dirpath) / name
            if p.suffix.lower() in IMAGE_EXTS:
                files.append(str(p))
    return files


def build_tasks(root_path: str, name_index: dict) -> list[dict]:
    tasks = []
    for file_path in scan_files(root_path):
        filename = Path(file_path).stem
        student = extract_student_name(filename)
        if not student:
            tasks.append(
                {
                    "file_path": file_path,
                    "student_name": "",
                    "parent_name": "",
                    "user_id": None,
                    "status": "skipped",
                    "error": "无法提取学生姓名",
                }
            )
            continue
        parent_name = f"{student}妈妈"
        user_id = name_index.get(parent_name)
        if not user_id:
            tasks.append(
                {
                    "file_path": file_path,
                    "student_name": student,
                    "parent_name": parent_name,
                    "user_id": None,
                    "status": "skipped",
                    "error": "通讯录未匹配到妈妈",
                }
            )
            continue
        tasks.append(
            {
                "file_path": file_path,
                "student_name": student,
                "parent_name": parent_name,
                "user_id": user_id,
                "status": "pending",
                "error": None,
            }
        )
    return tasks
