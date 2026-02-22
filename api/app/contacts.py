from __future__ import annotations

import re
from pathlib import Path
import pandas as pd

CJK_RE = re.compile(r"^([\u4e00-\u9fff]+)")


def extract_student_name(filename: str) -> str | None:
    m = CJK_RE.match(filename)
    return m.group(1) if m else None


def _find_header_row(df: pd.DataFrame) -> int:
    for i in range(len(df)):
        row = df.iloc[i].astype(str).tolist()
        if "姓名" in row and "账号" in row:
            return i
    return 0


def read_contacts(path: str | Path) -> list[dict]:
    p = Path(path)
    if not p.exists():
        return []
    raw = pd.read_excel(p, sheet_name="成员列表", header=None)
    header_row = _find_header_row(raw)
    df = pd.read_excel(p, sheet_name="成员列表", header=header_row)
    df = df[["姓名", "账号"]].dropna()
    rows = df.to_dict(orient="records")
    return rows


def build_name_index(rows: list[dict]) -> dict:
    idx: dict[str, str] = {}
    for r in rows:
        name = str(r.get("姓名") or "").strip()
        user_id = str(r.get("账号") or "").strip()
        if name and user_id:
            idx[name] = user_id
    return idx
