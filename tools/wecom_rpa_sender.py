#!/usr/bin/env python
"""
Windows WeCom RPA sender (pywinauto).

This is a practical skeleton for batch sending report images to parents by UI automation.
Selectors may need small adjustments on different WeCom versions.
"""
from __future__ import annotations

import argparse
import csv
import logging
import re
import subprocess
import sys
import time
import struct
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys
import win32clipboard
import win32con
import win32gui
import win32process


@dataclass
class Task:
    parent_name: str
    image_path: Path
    student_name: str = ""
    search_keyword: str = ""
    confirm_keyword: str = ""
    message_text: str = ""


TUNE = {
    "stabilize_open_rounds": 2,
    "stabilize_focus_rounds": 2,
    "stabilize_send_rounds": 1,
    "open_chat_strategy": "keyboard_first",
}


def setup_logger(log_file: Path) -> logging.Logger:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("wecom_rpa")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(sh)
    return logger


def read_tasks(csv_path: Path) -> list[Task]:
    tasks: list[Task] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"parent_name", "image_path"}
        cols = {c.strip() for c in (reader.fieldnames or [])}
        if not required.issubset(cols):
            raise ValueError("CSV must contain columns: parent_name,image_path")

        for row in reader:
            parent = (row.get("parent_name") or "").strip()
            image = (row.get("image_path") or "").strip()
            student = (row.get("student_name") or "").strip()
            search_keyword = (row.get("search_keyword") or "").strip()
            confirm_keyword = (row.get("confirm_keyword") or "").strip()
            raw_message_text = str(row.get("message_text") or "")
            message_text = (
                raw_message_text.replace("\\r\\n", "\n").replace("\\n", "\n").strip()
            )
            if not parent or not image:
                continue
            tasks.append(
                Task(
                    parent_name=parent,
                    image_path=Path(image),
                    student_name=student,
                    search_keyword=search_keyword or parent,
                    confirm_keyword=confirm_keyword or parent,
                    message_text=message_text,
                )
            )
    return tasks


def wait_until(timeout_sec: float, cond, interval_sec: float = 0.2) -> bool:
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if cond():
            return True
        time.sleep(interval_sec)
    return False


def _main_pid(main_win) -> int | None:
    try:
        return main_win.process_id()
    except Exception:
        return None


def _ensure_wecom_foreground(main_win, timeout_sec: float = 1.2) -> bool:
    target_pid = _main_pid(main_win)
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        try:
            main_win.set_focus()
        except Exception:
            pass
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, fg_pid = win32process.GetWindowThreadProcessId(hwnd)
            if target_pid is None or fg_pid == target_pid:
                return True
        except Exception:
            pass
        time.sleep(0.08)
    return False


def _safe_send_keys(main_win, keys: str, **kwargs):
    _ensure_wecom_foreground(main_win)
    send_keys(keys, **kwargs)


def ensure_wecom_running(wecom_exe: str | None, title_re: str, timeout_sec: float) -> Application:
    app = Application(backend="uia")
    try:
        app.connect(title_re=title_re)
        return app
    except Exception:
        # Fallback: connect by process executable name when title matching fails.
        try:
            app.connect(path="WXWork.exe")
            return app
        except Exception:
            pass
        if not wecom_exe:
            raise RuntimeError("WeCom not running/matched. Keep WeCom window visible or provide --wecom-exe.")

    app = Application(backend="uia").start(wecom_exe)
    ok = wait_until(
        timeout_sec,
        lambda: Desktop(backend="uia").window(title_re=title_re).exists()
        or Desktop(backend="uia").window(title_re=".*(企业微信|WeCom|WXWork).*").exists(),
    )
    if not ok:
        raise RuntimeError("WeCom window not found after start.")
    try:
        app.connect(title_re=title_re)
    except Exception:
        app.connect(path="WXWork.exe")
    return app


def _select_main_window(title_re: str, process_id: int | None = None):
    desktop = Desktop(backend="uia")
    if process_id:
        try:
            windows = desktop.windows(process=process_id)
        except Exception:
            windows = desktop.windows()
    else:
        windows = desktop.windows()
    candidates = []
    fallback = []
    process_fallback = []
    for w in windows:
        try:
            if not w.is_visible():
                continue
            title = (w.window_text() or "").strip()
            if not title:
                # Some WeCom builds expose empty-title top-level windows.
                rect = w.rectangle()
                area = max(0, rect.width()) * max(0, rect.height())
                if process_id:
                    process_fallback.append((area, w))
                continue
            rect = w.rectangle()
            area = max(0, rect.width()) * max(0, rect.height())
            if re.search(title_re, title, flags=re.IGNORECASE):
                candidates.append((area, w))
            elif re.search(r"(企业微信|WeCom|WXWork)", title, flags=re.IGNORECASE):
                fallback.append((area, w))
            elif process_id:
                process_fallback.append((area, w))
        except Exception:
            continue
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]
    if fallback:
        fallback.sort(key=lambda x: x[0], reverse=True)
        return fallback[0][1]
    if process_fallback:
        process_fallback.sort(key=lambda x: x[0], reverse=True)
        return process_fallback[0][1]
    return None


def wait_for_main_window(title_re: str, timeout_sec: float, app: Application | None = None):
    title_re = (title_re or "").replace("\r", "").replace("\n", "").strip() or r".*(企业微信|WeCom|WXWork).*"

    process_id = None
    if app is not None:
        try:
            process_id = app.process
        except Exception:
            process_id = None

    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        w = _select_main_window(title_re, process_id=process_id)
        if w is None:
            # If app is attached to a helper process without top-level window,
            # fall back to scanning all desktop windows.
            w = _select_main_window(title_re, process_id=None)
        if w is not None:
            try:
                if w.is_minimized():
                    w.restore()
            except Exception:
                pass
            return w
        time.sleep(0.3)
    raise RuntimeError(f"WeCom main window not found by title regex: {title_re}")


def pick_search_edit(main_win):
    edits = main_win.descendants(control_type="Edit")
    edits = [e for e in edits if e.is_visible() and e.is_enabled()]
    if not edits:
        raise RuntimeError("Search box not found")

    # Heuristic: WeCom session search box is in left/top sidebar.
    rect = main_win.rectangle()
    left_max = rect.left + int(rect.width() * 0.45)
    top_max = rect.top + int(rect.height() * 0.28)
    sidebar_top = [
        e
        for e in edits
        if e.rectangle().left <= left_max and e.rectangle().top <= top_max
    ]
    if sidebar_top:
        sidebar_top.sort(key=lambda e: (e.rectangle().top, e.rectangle().left))
        return sidebar_top[0]

    edits.sort(key=lambda e: (e.rectangle().top, e.rectangle().left))
    return edits[0]


def pick_chat_input_edit(main_win):
    edits = main_win.descendants(control_type="Edit")
    edits = [e for e in edits if e.is_visible() and e.is_enabled()]
    if not edits:
        raise RuntimeError("Chat input box not found")

    rect = main_win.rectangle()
    min_top = rect.top + int(rect.height() * 0.60)
    min_left = rect.left + int(rect.width() * 0.22)
    min_width = int(rect.width() * 0.25)

    candidates = []
    for e in edits:
        r = e.rectangle()
        if r.top < min_top:
            continue
        if r.left < min_left:
            continue
        if r.width() < min_width:
            continue
        candidates.append(e)

    if not candidates:
        candidates = [e for e in edits if e.rectangle().top >= min_top] or edits

    # Prefer the largest lower editor-like control.
    candidates.sort(key=lambda e: (e.rectangle().width() * e.rectangle().height(), e.rectangle().top), reverse=True)
    return candidates[0]


def open_chat(main_win, parent_name: str, timeout_sec: float):
    _stabilize_wecom_ui(main_win, rounds=int(TUNE["stabilize_open_rounds"]))
    queries = _keyword_variants(parent_name)
    per_try_timeout = max(0.8, min(3.0, timeout_sec / max(1, len(queries))))

    # Keyboard-first flow: Ctrl+F to global search, type keyword, click first result row.
    # This avoids UIA search-box dependency on some WeCom builds.
    hit = False
    for q in queries:
        _stabilize_wecom_ui(main_win, rounds=int(TUNE["stabilize_open_rounds"]))
        _safe_send_keys(main_win, "^f")
        time.sleep(0.2)
        _safe_send_keys(main_win, "^a{BACKSPACE}")
        _safe_send_keys(main_win, q, with_spaces=True, pause=0.03)
        time.sleep(0.28)
        if _open_search_result(main_win, q, per_try_timeout):
            # On some WeCom builds UIA cannot detect the editor reliably.
            # If result row is clicked, proceed and force focus the editor area.
            time.sleep(0.15)
            _prepare_chat_input_focus(main_win)
            hit = True
            break

    if not hit:
        raise RuntimeError(f"Search result not found/clickable: {parent_name}")

    # Best-effort focus chat input. Do not fail hard here; sending step still has its own focusing logic.
    wait_until(min(timeout_sec, 1.5), lambda: _focus_chat_input(main_win))


def _click_left_result_row(main_win) -> bool:
    rect = main_win.rectangle()
    x = int(rect.width() * 0.15)
    y = int(rect.height() * 0.19)
    try:
        main_win.click_input(coords=(x, y))
        return True
    except Exception:
        return False


def _open_search_result(main_win, keyword: str, timeout_sec: float) -> bool:
    strategy = str(TUNE.get("open_chat_strategy", "keyboard_first")).strip().lower()
    if strategy == "click_first":
        if _click_search_result(main_win, keyword, timeout_sec=min(1.2, timeout_sec)):
            return True
        if _click_left_result_row(main_win):
            return True
        for keys in ("{ENTER}", "{DOWN}{ENTER}"):
            try:
                _safe_send_keys(main_win, keys)
                time.sleep(0.16)
                return True
            except Exception:
                continue
        return False

    # keyboard_first / hybrid
    for keys in ("{ENTER}", "{DOWN}{ENTER}"):
        try:
            _safe_send_keys(main_win, keys)
            time.sleep(0.16)
            return True
        except Exception:
            continue
    if _click_search_result(main_win, keyword, timeout_sec=min(1.2, timeout_sec)):
        return True
    return _click_left_result_row(main_win)


def _click_search_result(main_win, keyword: str, timeout_sec: float) -> bool:
    # Click a visible text item in the left result pane that best matches keyword.
    def _norm(s: str) -> str:
        return "".join(ch for ch in (s or "") if ch.isalnum() or ("\u4e00" <= ch <= "\u9fff")).lower()

    targets = [_norm(x) for x in _keyword_variants(keyword)]
    targets = [t for t in targets if t]
    if not targets:
        return False

    deadline = time.time() + max(1.2, timeout_sec)
    rect = main_win.rectangle()
    left_max = rect.left + int(rect.width() * 0.45)
    top_min = rect.top + int(rect.height() * 0.08)
    top_max = rect.top + int(rect.height() * 0.72)

    while time.time() < deadline:
        best = None
        best_score = -1
        texts = main_win.descendants(control_type="Text")
        for t in texts:
            try:
                if not t.is_visible():
                    continue
                r = t.rectangle()
                if r.left > left_max or r.top < top_min or r.top > top_max:
                    continue
                txt = (t.window_text() or "").strip()
                nt = _norm(txt)
                if len(nt) < 2:
                    continue
                score = 0
                for target in targets:
                    if target == nt:
                        score = max(score, 100)
                    elif target in nt:
                        score = max(score, 85)
                    elif nt in target:
                        score = max(score, 65)
                if score > best_score:
                    best_score = score
                    best = t
            except Exception:
                continue

        if best is not None and best_score >= 60:
            try:
                best.click_input()
                return True
            except Exception:
                # Try clicking near the text center as fallback.
                try:
                    br = best.rectangle()
                    x = br.left + max(2, br.width() // 2)
                    y = br.top + max(2, br.height() // 2)
                    main_win.click_input(coords=(x - rect.left, y - rect.top))
                    return True
                except Exception:
                    pass
        time.sleep(0.12)
    return False


def _keyword_variants(keyword: str) -> list[str]:
    # Build query variants for composite names like "卢天若/卢天亦妈妈".
    raw = (keyword or "").strip()
    if not raw:
        return []

    cands: list[str] = [raw]
    parts = [p.strip() for p in re.split(r"[\\/|｜]", raw) if p.strip()]
    cands.extend(parts)

    if "妈妈" in raw:
        cands.append(raw.replace("妈妈", ""))
    for p in parts:
        if "妈妈" in p:
            cands.append(p.replace("妈妈", ""))
        else:
            cands.append(f"{p}妈妈")

    out: list[str] = []
    seen = set()
    for c in cands:
        n = _normalize_text(c)
        if len(n) < 2 or n in seen:
            continue
        seen.add(n)
        out.append(c)
    return out


def _normalize_text(s: str) -> str:
    return "".join(ch for ch in s if ch.isalnum() or ("\u4e00" <= ch <= "\u9fff")).lower()


def _collect_header_texts(main_win) -> list[str]:
    rect = main_win.rectangle()
    max_top = rect.top + int(rect.height() * 0.36)
    min_left = rect.left + int(rect.width() * 0.22)
    texts: list[str] = []
    seen = set()

    for c in main_win.descendants(control_type="Text"):
        try:
            if not c.is_visible():
                continue
            r = c.rectangle()
            if r.top > max_top or r.left < min_left:
                continue
            txt = (c.window_text() or "").strip()
            if not txt or len(txt) > 80:
                continue
            if txt in seen:
                continue
            seen.add(txt)
            texts.append(txt)
        except Exception:
            continue
    return texts


def _candidate_keywords(task: Task) -> list[str]:
    kws: list[str] = []
    base = (task.confirm_keyword or task.parent_name or "").strip()
    student = (task.student_name or "").strip()

    if base:
        kws.append(base)
        parts = [p.strip() for p in re.split(r"[\\/|｜]", base) if p.strip()]
        kws.extend(parts)
        if "妈妈" in base:
            for p in parts:
                if "妈妈" not in p:
                    kws.append(f"{p}妈妈")
    if student:
        kws.append(student)

    out: list[str] = []
    seen = set()
    for k in kws:
        nk = _normalize_text(k)
        if len(nk) < 2:
            continue
        if nk in seen:
            continue
        seen.add(nk)
        out.append(k)
    return out


def verify_chat_selected(main_win, task: Task) -> bool:
    header_texts = _collect_header_texts(main_win)
    norm_header = [_normalize_text(t) for t in header_texts if _normalize_text(t)]
    if not norm_header:
        return False

    for kw in _candidate_keywords(task):
        nkw = _normalize_text(kw)
        if not nkw:
            continue
        for ht in norm_header:
            if nkw in ht:
                return True
    return False


def _focus_chat_input(main_win) -> bool:
    try:
        inp = pick_chat_input_edit(main_win)
        inp.click_input()
        return True
    except Exception:
        return False


def send_image_via_dialog(main_win, image_path: Path, timeout_sec: float, paste_only: bool = False):
    if not image_path.exists():
        raise FileNotFoundError(f"Image not found: {image_path}")

    # Try common shortcut to open send-file dialog.
    _ensure_wecom_foreground(main_win)
    _focus_chat_input(main_win)
    _safe_send_keys(main_win, "^o")

    desktop = Desktop(backend="uia")

    dialog_title_re = r"^(打开|Open|选择文件|Select).*"

    def _dlg_exists() -> bool:
        return desktop.window(title_re=dialog_title_re).exists(timeout=0.1)

    if not wait_until(timeout_sec, _dlg_exists):
        raise RuntimeError("Open file dialog not found. You may need to adjust shortcut/selector.")

    dlg = desktop.window(title_re=dialog_title_re)
    dlg.set_focus()

    image_str = str(image_path)
    image_escaped = image_str.replace("+", "{+}")

    # First try control-based path input.
    edits = [e for e in dlg.descendants(control_type="Edit") if e.is_visible() and e.is_enabled()]
    if edits:
        target = edits[-1]
        target.click_input()
        _safe_send_keys(main_win, "^a{BACKSPACE}")
        _safe_send_keys(main_win, image_escaped, with_spaces=True, pause=0.01)
        _safe_send_keys(main_win, "{ENTER}")
    else:
        # Fallback: use "File name" mnemonic and paste path.
        _safe_send_keys(main_win, "%n")
        time.sleep(0.1)
        _safe_send_keys(main_win, "^a{BACKSPACE}")
        _copy_text_to_clipboard(image_str)
        _safe_send_keys(main_win, "^v")
        _safe_send_keys(main_win, "{ENTER}")

    # If dialog still open, click Open button as a fallback.
    if desktop.window(title_re=dialog_title_re).exists(timeout=0.3):
        try:
            dlg = desktop.window(title_re=dialog_title_re)
            buttons = [b for b in dlg.descendants(control_type="Button") if b.is_visible() and b.is_enabled()]
            open_btn = None
            for b in buttons:
                txt = (b.window_text() or "").strip()
                if re.search(r"(打开|Open)", txt, flags=re.IGNORECASE):
                    open_btn = b
                    break
            if open_btn is not None:
                open_btn.click_input()
            else:
                dlg.set_focus()
                _safe_send_keys(main_win, "{ENTER}")
        except Exception:
            dlg.set_focus()
            _safe_send_keys(main_win, "{ENTER}")

    # In some clients, there is a preview step that needs one more Enter to send.
    if not paste_only:
        time.sleep(0.8)
        _safe_send_keys(main_win, "{ENTER}")


def _send_image_via_clipboard(main_win, image_path: Path, paste_only: bool = False) -> bool:
    # Best effort focus: if chat input selector fails, still try paste on main window.
    _prepare_chat_input_focus(main_win)
    baseline_images = _count_input_area_images(main_win)
    _copy_file_to_clipboard(image_path)
    time.sleep(0.2)

    pasted = False
    confirmed = False
    paste_attempts = 1 if paste_only else 3
    for _ in range(paste_attempts):
        _prepare_chat_input_focus(main_win)
        if _paste_into_chat_input(main_win):
            pasted = True
            time.sleep(0.25)
            if _attachment_paste_confirmed(main_win, image_path, baseline_images, timeout_sec=1.4):
                confirmed = True
                break
        time.sleep(0.15)

    # Clipboard mode stays clipboard-only: no dialog fallback.
    if not pasted:
        raise RuntimeError("Clipboard paste hotkey failed (input focus not ready).")
    if not confirmed:
        if paste_only:
            # In paste-only rehearsal mode, treat unverified paste as soft success.
            # Some WeCom builds don't expose input attachment controls to UIA reliably.
            return False
        raise RuntimeError("Clipboard paste not confirmed in chat input.")

    if paste_only:
        return True

    if not paste_only:
        time.sleep(0.9)
        _safe_send_keys(main_win, "{ENTER}")
    return True


def _copy_file_to_clipboard(path: Path):
    # Put a real file object (CF_HDROP) on clipboard, not plain text path.
    p = str(path.resolve())
    if not Path(p).exists():
        raise FileNotFoundError(f"Image not found: {p}")

    # First try WinForms Clipboard.SetFileDropList (very stable on Windows apps).
    p_escaped = p.replace("'", "''")
    ps_cmd = (
        "Add-Type -AssemblyName System.Windows.Forms; "
        "$list = New-Object System.Collections.Specialized.StringCollection; "
        f"$null = $list.Add('{p_escaped}'); "
        "[System.Windows.Forms.Clipboard]::SetFileDropList($list)"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Sta", "-Command", ps_cmd],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if _clipboard_has_target_file(Path(p)):
            return
    except Exception:
        pass

    # Fallback to direct CF_HDROP payload.
    payload = p + "\0\0"
    encoded = payload.encode("utf-16le")
    dropfiles = struct.pack("<IiiII", 20, 0, 0, 0, 1)
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_HDROP, dropfiles + encoded)
    finally:
        win32clipboard.CloseClipboard()

    if not _clipboard_has_target_file(Path(p)):
        raise RuntimeError("Failed to put target file on clipboard (CF_HDROP target mismatch)")


def _clipboard_has_files() -> bool:
    try:
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP):
                return False
            files = win32clipboard.GetClipboardData(win32con.CF_HDROP)
            return bool(files)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return False


def _normalize_win_path(path_str: str) -> str:
    s = (path_str or "").strip().strip('"').replace("/", "\\")
    try:
        s = str(Path(s).resolve())
    except Exception:
        pass
    return s.rstrip("\\").lower()


def _clipboard_has_target_file(path: Path) -> bool:
    target = _normalize_win_path(str(path))
    if not target:
        return False
    try:
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP):
                return False
            files = win32clipboard.GetClipboardData(win32con.CF_HDROP)
            if not files:
                return False
            for f in files:
                if _normalize_win_path(str(f)) == target:
                    return True
            return False
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return False


def _paste_into_chat_input(main_win) -> bool:
    _ensure_wecom_foreground(main_win)
    try:
        inp = pick_chat_input_edit(main_win)
        inp.click_input()
        try:
            inp.type_keys("^v", with_spaces=True, set_foreground=False, pause=0.01)
        except Exception:
            _safe_send_keys(main_win, "^v")
        return True
    except Exception:
        # Fallback: selector may fail on some WeCom builds, but hotkey can still work.
        try:
            rect = main_win.rectangle()
            x = int(rect.width() * 0.72)
            y = int(rect.height() * 0.92)
            main_win.click_input(coords=(x, y))
            time.sleep(0.08)
            _safe_send_keys(main_win, "^v")
            return True
        except Exception:
            return False


def _attachment_hint_present(main_win, image_path: Path, timeout_sec: float = 1.2) -> bool:
    # Detect likely attachment chip text in input area as a best-effort signal.
    input_rect = _chat_input_rect(main_win)
    targets = {image_path.name, image_path.stem}
    norm_targets = {_normalize_text(t) for t in targets if _normalize_text(t)}
    if not norm_targets:
        return False

    deadline = time.time() + max(0.4, timeout_sec)
    while time.time() < deadline:
        texts = main_win.descendants(control_type="Text")
        for t in texts:
            try:
                if not t.is_visible():
                    continue
                r = t.rectangle()
                if not _rect_inside(r, input_rect):
                    continue
                txt = (t.window_text() or "").strip()
                nt = _normalize_text(txt)
                if not nt:
                    continue
                for q in norm_targets:
                    if len(q) >= 3 and (q in nt or nt in q):
                        return True
            except Exception:
                continue
        time.sleep(0.12)
    return False


def _count_input_area_images(main_win) -> int:
    input_rect = _chat_input_rect(main_win)
    count = 0
    for c in main_win.descendants(control_type="Image"):
        try:
            if not c.is_visible():
                continue
            r = c.rectangle()
            if not _rect_inside(r, input_rect):
                continue
            count += 1
        except Exception:
            continue
    return count


def _attachment_paste_confirmed(main_win, image_path: Path, baseline_images: int, timeout_sec: float = 1.4) -> bool:
    deadline = time.time() + max(0.4, timeout_sec)
    while time.time() < deadline:
        if _count_input_area_images(main_win) > baseline_images:
            return True
        # Keep as secondary fallback only (inside input area), as some clients do show file-name chips.
        if _attachment_hint_present(main_win, image_path, timeout_sec=0.15):
            return True
        time.sleep(0.1)
    return False


def _rect_inside(child, parent) -> bool:
    cx = child.left + max(1, child.width() // 2)
    cy = child.top + max(1, child.height() // 2)
    return parent.left <= cx <= parent.right and parent.top <= cy <= parent.bottom


def _chat_input_rect(main_win):
    try:
        return pick_chat_input_edit(main_win).rectangle()
    except Exception:
        rect = main_win.rectangle()
        left = rect.left + int(rect.width() * 0.22)
        top = rect.top + int(rect.height() * 0.60)
        return type(rect)(left, top, rect.right - 6, rect.bottom - 6)


def _prepare_chat_input_focus(main_win, clear_input: bool = False):
    # Close possible global search box/popups and return focus to chat area.
    _stabilize_wecom_ui(main_win, rounds=int(TUNE["stabilize_focus_rounds"]))

    if _focus_chat_input(main_win):
        if clear_input:
            try:
                _safe_send_keys(main_win, "^a{BACKSPACE}")
            except Exception:
                pass
        return

    # Coordinate fallback: click in lower-right editor region.
    rect = main_win.rectangle()
    rel_x = int(rect.width() * 0.74)
    rel_y = int(rect.height() * 0.92)
    try:
        main_win.click_input(coords=(rel_x, rel_y))
        time.sleep(0.08)
        if clear_input:
            try:
                _safe_send_keys(main_win, "^a{BACKSPACE}")
            except Exception:
                pass
    except Exception:
        _ensure_wecom_foreground(main_win)


def _stabilize_wecom_ui(main_win, rounds: int = 2):
    _ensure_wecom_foreground(main_win)
    for _ in range(max(0, int(rounds))):
        try:
            _safe_send_keys(main_win, "{ESC}")
        except Exception:
            pass
        time.sleep(0.08)


def _copy_text_to_clipboard(text: str):
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, str(text or ""))
    finally:
        win32clipboard.CloseClipboard()


def send_text_message(main_win, text: str, timeout_sec: float, paste_only: bool = False):
    body = str(text or "").strip()
    if not body:
        return

    _prepare_chat_input_focus(main_win)
    _copy_text_to_clipboard(body)
    time.sleep(0.1)
    if not _paste_into_chat_input(main_win):
        _prepare_chat_input_focus(main_win)
        _safe_send_keys(main_win, "^v")
    time.sleep(0.2)
    if paste_only:
        return
    _safe_send_keys(main_win, "{ENTER}")

    err = _detect_wecom_error_dialog(main_win, timeout_sec=min(max(timeout_sec, 1.0), 4.0))
    if err:
        raise RuntimeError(f"WeCom text send dialog error: {err}")


def _detect_wecom_error_dialog(main_win, timeout_sec: float = 1.2) -> str | None:
    desktop = Desktop(backend="uia")
    deadline = time.time() + timeout_sec
    target_pid = None
    try:
        target_pid = main_win.process_id()
    except Exception:
        target_pid = None

    while time.time() < deadline:
        for w in desktop.windows():
            try:
                if not w.is_visible():
                    continue
                pid = getattr(w.element_info, "process_id", None)
                if target_pid and pid != target_pid:
                    continue
                title = (w.window_text() or "").strip()
                cls = getattr(w.element_info, "class_name", "") or ""
                text_parts = [title]
                for t in w.descendants(control_type="Text"):
                    try:
                        s = (t.window_text() or "").strip()
                        if s:
                            text_parts.append(s)
                    except Exception:
                        continue
                merged = " | ".join(text_parts)
                if re.search(r"(失败|错误|error|fail|频繁|限制|无权限|not allow)", merged, flags=re.IGNORECASE):
                    # Close pop dialog to avoid blocking next step.
                    try:
                        w.set_focus()
                        _safe_send_keys(main_win, "{ESC}")
                    except Exception:
                        pass
                    return merged[:240]
                # Some builds use generic title; still check common dialog class.
                if cls == "#32770" and re.search(r"(提示|warning|提醒|notice)", merged, flags=re.IGNORECASE):
                    try:
                        w.set_focus()
                        _safe_send_keys(main_win, "{ESC}")
                    except Exception:
                        pass
                    return merged[:240]
            except Exception:
                continue
        time.sleep(0.15)
    return None


def send_image(main_win, image_path: Path, timeout_sec: float, mode: str, paste_only: bool = False) -> bool:
    _stabilize_wecom_ui(main_win, rounds=int(TUNE["stabilize_send_rounds"]))
    if mode == "clipboard":
        return _send_image_via_clipboard(main_win, image_path, paste_only=paste_only)
    elif mode == "dialog":
        send_image_via_dialog(main_win, image_path, timeout_sec, paste_only=paste_only)
        return True
    else:
        # auto mode: dialog first, then fallback clipboard
        try:
            send_image_via_dialog(main_win, image_path, timeout_sec, paste_only=paste_only)
            confirmed = True
        except Exception:
            confirmed = _send_image_via_clipboard(main_win, image_path, paste_only=paste_only)

    if not paste_only:
        err = _detect_wecom_error_dialog(main_win, timeout_sec=min(max(timeout_sec, 1.0), 4.0))
        if err:
            raise RuntimeError(f"WeCom send dialog error: {err}")
    return confirmed if mode == "auto" else True


def _norm_key_part(s: str) -> str:
    return (s or "").strip().lower()


def task_key(task: Task) -> tuple[str, str]:
    return (_norm_key_part(task.parent_name), _norm_key_part(str(task.image_path)))


def row_key(row: dict) -> tuple[str, str]:
    return (_norm_key_part(row.get("parent_name", "")), _norm_key_part(row.get("image_path", "")))


def read_result_rows(path: Path) -> list[dict]:
    if not path.exists():
        return []
    rows: list[dict] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(dict(row))
    return rows


def index_result_status(rows: list[dict]) -> dict[tuple[str, str], str]:
    out: dict[tuple[str, str], str] = {}
    for r in rows:
        out[row_key(r)] = (r.get("status") or "").strip()
    return out


def merge_result_rows(previous_rows: list[dict], new_rows: list[dict]) -> list[dict]:
    new_map = {row_key(r): r for r in new_rows}
    merged: list[dict] = []
    used = set()
    for r in previous_rows:
        k = row_key(r)
        if k in new_map:
            merged.append(new_map[k])
            used.add(k)
        else:
            merged.append(r)
    for r in new_rows:
        k = row_key(r)
        if k not in used:
            merged.append(r)
    return merged


def recover_ui_after_failure(main_win):
    # Best-effort UI recovery between retries/tasks.
    try:
        _ensure_wecom_foreground(main_win, timeout_sec=0.8)
        _safe_send_keys(main_win, "{ESC}")
        time.sleep(0.08)
        _safe_send_keys(main_win, "{ESC}")
        time.sleep(0.08)
    except Exception:
        pass


def write_results(path: Path, rows: Iterable[dict]):
    rows = list(rows)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "parent_name",
                "student_name",
                "image_path",
                "message_text",
                "text_status",
                "status",
                "error",
                "attempts",
                "timestamp",
            ],
        )
        writer.writeheader()
        for r in rows:
            writer.writerow(r)


def dump_visible_windows(limit: int = 80) -> list[str]:
    lines: list[str] = []
    desktop = Desktop(backend="uia")
    for w in desktop.windows():
        try:
            if not w.is_visible():
                continue
            title = (w.window_text() or "").strip()
            pid = getattr(w.element_info, "process_id", None)
            cls = getattr(w.element_info, "class_name", "")
            rect = w.rectangle()
            area = max(0, rect.width()) * max(0, rect.height())
            lines.append(f"pid={pid} area={area} class={cls} title={title}")
        except Exception:
            continue
    lines.sort(key=lambda s: int(re.search(r"area=(\d+)", s).group(1)) if re.search(r"area=(\d+)", s) else 0, reverse=True)
    return lines[:limit]


def main() -> int:
    parser = argparse.ArgumentParser(description="WeCom report sender via UI automation")
    parser.add_argument(
        "--tasks-csv",
        required=True,
        help="CSV with columns: parent_name,image_path[,student_name,search_keyword,confirm_keyword,message_text]",
    )
    parser.add_argument("--wecom-exe", default=None, help="Optional WeCom executable path")
    parser.add_argument("--main-title-re", default=".*企业微信.*", help="Main window title regex")
    parser.add_argument("--interval-sec", type=float, default=1.2, help="Delay between tasks")
    parser.add_argument("--timeout-sec", type=float, default=10.0, help="Timeout for UI waits")
    parser.add_argument(
        "--send-mode",
        choices=["clipboard", "dialog", "auto"],
        default="clipboard",
        help="Image send mode. Default uses clipboard paste.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Only navigate chats, do not send images")
    parser.add_argument("--paste-only", action="store_true", help="Paste image into chat input without sending")
    parser.add_argument("--no-chat-verify", action="store_true", help="Disable chat title verify before sending")
    parser.add_argument("--max-retries", type=int, default=2, help="Retry count per task on failure")
    parser.add_argument("--retry-delay-sec", type=float, default=1.0, help="Delay before next retry")
    parser.add_argument(
        "--stabilize-open-rounds",
        type=int,
        default=2,
        help="Anti-disturb ESC rounds before opening/searching chat (0-5)",
    )
    parser.add_argument(
        "--stabilize-focus-rounds",
        type=int,
        default=2,
        help="Anti-disturb ESC rounds before focusing chat input (0-5)",
    )
    parser.add_argument(
        "--stabilize-send-rounds",
        type=int,
        default=1,
        help="Anti-disturb ESC rounds before sending/pasting image (0-5)",
    )
    parser.add_argument(
        "--open-chat-strategy",
        choices=["keyboard_first", "click_first", "hybrid"],
        default="keyboard_first",
        help="Strategy to open search result chat.",
    )
    parser.add_argument("--stop-on-fail", action="store_true", help="Stop whole run when one task fails")
    parser.add_argument("--skip-missing-image", action="store_true", help="Mark missing image as skipped instead of failed")
    parser.add_argument("--resume-from", type=int, default=1, help="Resume from task index (1-based)")
    parser.add_argument("--resume-failed", action="store_true", help="Skip tasks already marked ok in previous result file")
    parser.add_argument(
        "--resume-results-csv",
        default="run-logs/rpa-results.csv",
        help="Previous result file used by --resume-failed",
    )
    parser.add_argument("--debug-windows", action="store_true", help="Print visible desktop windows before run")
    parser.add_argument("--debug-chat-text", action="store_true", help="Print detected header texts for each task")
    parser.add_argument("--skip-text-message", action="store_true", help="Do not send message_text even if CSV provides it")
    parser.add_argument("--results-csv", default="run-logs/rpa-results.csv", help="Result output file")
    parser.add_argument("--log-file", default="run-logs/rpa-sender.log", help="Log file")
    args = parser.parse_args()

    logger = setup_logger(Path(args.log_file))
    TUNE["stabilize_open_rounds"] = max(0, min(5, int(args.stabilize_open_rounds)))
    TUNE["stabilize_focus_rounds"] = max(0, min(5, int(args.stabilize_focus_rounds)))
    TUNE["stabilize_send_rounds"] = max(0, min(5, int(args.stabilize_send_rounds)))
    TUNE["open_chat_strategy"] = str(args.open_chat_strategy).strip().lower() or "keyboard_first"
    logger.info(
        "Anti-disturb tuning: open_rounds=%s focus_rounds=%s send_rounds=%s open_strategy=%s",
        TUNE["stabilize_open_rounds"],
        TUNE["stabilize_focus_rounds"],
        TUNE["stabilize_send_rounds"],
        TUNE["open_chat_strategy"],
    )
    tasks = read_tasks(Path(args.tasks_csv))
    if not tasks:
        logger.error("No valid tasks loaded from CSV")
        return 2

    total_tasks = len(tasks)
    indexed_tasks = [(i + 1, t) for i, t in enumerate(tasks)]

    resume_from = max(1, int(args.resume_from))
    if resume_from > total_tasks:
        logger.error("--resume-from (%s) exceeds total tasks (%s)", resume_from, total_tasks)
        return 2
    if resume_from > 1:
        indexed_tasks = [x for x in indexed_tasks if x[0] >= resume_from]
        logger.info("Resume from index %s: pending %s/%s task(s)", resume_from, len(indexed_tasks), total_tasks)

    previous_rows: list[dict] = []
    previous_status_map: dict[tuple[str, str], str] = {}
    if args.resume_failed:
        prev_path = Path(args.resume_results_csv)
        previous_rows = read_result_rows(prev_path)
        previous_status_map = index_result_status(previous_rows)
        resume_skip_status = {"ok", "sent"}
        before = len(indexed_tasks)
        indexed_tasks = [
            (idx, t)
            for idx, t in indexed_tasks
            if previous_status_map.get(task_key(t), "") not in resume_skip_status
        ]
        logger.info(
            "Resume failed mode: pending %s/%s task(s) after skipping previous ok",
            len(indexed_tasks),
            before,
        )
        if not indexed_tasks:
            logger.info("No pending tasks to run.")
            return 0

    logger.info("Loaded %s task(s)", len(tasks))
    logger.info("Switch to WeCom desktop and do not touch keyboard/mouse during run.")
    if args.debug_windows:
        for line in dump_visible_windows():
            logger.info("[win] %s", line)

    try:
        app = ensure_wecom_running(args.wecom_exe, args.main_title_re, args.timeout_sec)
    except Exception as e:
        logger.exception("Cannot attach WeCom: %s", e)
        return 3

    try:
        main_win = wait_for_main_window(args.main_title_re, args.timeout_sec, app=app)
        main_win.set_focus()
    except Exception as e:
        logger.exception("Cannot resolve WeCom main window: %s", e)
        return 4

    results: list[dict] = []
    max_attempts = max(1, int(args.max_retries) + 1)

    run_total = len(indexed_tasks)
    for run_idx, (orig_idx, task) in enumerate(indexed_tasks, start=1):
        row = {
            "parent_name": task.parent_name,
            "student_name": task.student_name,
            "image_path": str(task.image_path),
            "message_text": task.message_text,
            "text_status": "pending",
            "status": "pending",
            "error": "",
            "attempts": 0,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        }

        if not task.image_path.exists():
            row["status"] = "skipped_missing_image" if args.skip_missing_image else "failed"
            row["error"] = f"Image not found: {task.image_path}"
            row["attempts"] = 0
            row["text_status"] = "not_sent_missing_image"
            if args.skip_missing_image:
                logger.warning("[%s/%s|#%s] Skip missing image: %s", run_idx, run_total, orig_idx, task.image_path)
            else:
                logger.error("[%s/%s|#%s] Missing image: %s", run_idx, run_total, orig_idx, task.image_path)
        else:
            last_err = ""
            for attempt in range(1, max_attempts + 1):
                row["attempts"] = attempt
                try:
                    logger.info(
                        "[%s/%s|#%s][try %s/%s] Open chat: %s",
                        run_idx,
                        run_total,
                        orig_idx,
                        attempt,
                        max_attempts,
                        task.parent_name,
                    )
                    open_chat(main_win, task.search_keyword or task.parent_name, args.timeout_sec)
                    if args.debug_chat_text:
                        logger.info(
                            "[%s/%s|#%s] Header texts: %s",
                            run_idx,
                            run_total,
                            orig_idx,
                            " | ".join(_collect_header_texts(main_win)),
                        )
                    if not args.no_chat_verify:
                        ok = wait_until(args.timeout_sec, lambda: verify_chat_selected(main_win, task), interval_sec=0.25)
                        if not ok:
                            raise RuntimeError(
                                f"Chat verify failed for '{task.parent_name}'. "
                                "Set search_keyword/confirm_keyword in CSV to disambiguate."
                            )
                    if args.dry_run:
                        logger.info("[%s/%s|#%s] Dry run, skip sending", run_idx, run_total, orig_idx)
                        row["text_status"] = "dry_run"
                        row["status"] = "dry_run"
                    else:
                        action = "Paste image only" if args.paste_only else "Send image"
                        logger.info(
                            "[%s/%s|#%s] %s (%s): %s",
                            run_idx,
                            run_total,
                            orig_idx,
                            action,
                            args.send_mode,
                            task.image_path,
                        )
                        paste_confirmed = send_image(
                            main_win,
                            task.image_path,
                            args.timeout_sec,
                            args.send_mode,
                            paste_only=args.paste_only,
                        )
                        if args.paste_only:
                            if paste_confirmed:
                                row["status"] = "pasted_only"
                            else:
                                row["status"] = "pasted_unverified"
                                row["error"] = "Paste not UIA-confirmed (likely pasted, please spot-check chat input)."
                                logger.warning(
                                    "[%s/%s|#%s] Paste unverified by UIA: %s",
                                    run_idx,
                                    run_total,
                                    orig_idx,
                                    task.parent_name,
                                )
                        else:
                            row["status"] = "sent"

                        if args.skip_text_message:
                            row["text_status"] = "skipped_by_flag"
                        elif str(task.message_text or "").strip():
                            logger.info(
                                "[%s/%s|#%s] %s message_text after image (%s chars)",
                                run_idx,
                                run_total,
                                orig_idx,
                                "Paste-only" if args.paste_only else "Send",
                                len(task.message_text.strip()),
                            )
                            send_text_message(
                                main_win,
                                task.message_text,
                                args.timeout_sec,
                                paste_only=args.paste_only,
                            )
                            row["text_status"] = "pasted_only" if args.paste_only else "sent"
                        else:
                            row["text_status"] = "skipped_empty"
                    if row.get("status") != "pasted_unverified":
                        row["error"] = ""
                    break
                except Exception as e:
                    last_err = str(e)
                    logger.exception(
                        "[%s/%s|#%s][try %s/%s] Failed: %s",
                        run_idx,
                        run_total,
                        orig_idx,
                        attempt,
                        max_attempts,
                        task.parent_name,
                    )
                    recover_ui_after_failure(main_win)
                    if attempt < max_attempts:
                        time.sleep(max(0.2, args.retry_delay_sec))
                    else:
                        row["status"] = "failed"
                        row["error"] = last_err

        results.append(row)
        time.sleep(max(args.interval_sec, 0.0))

        if args.stop_on_fail and row["status"] == "failed":
            logger.error("Stop on fail is enabled. Abort at task #%s (%s/%s).", orig_idx, run_idx, run_total)
            break

    output_rows = results
    output_path = Path(args.results_csv)
    if args.resume_failed and output_path.resolve() == Path(args.resume_results_csv).resolve():
        output_rows = merge_result_rows(previous_rows, results)
    write_results(output_path, output_rows)
    sent = sum(1 for r in results if r["status"] in {"ok", "sent"})
    pasted = sum(1 for r in results if r["status"] in {"pasted_only", "pasted_unverified"})
    pasted_unverified = sum(1 for r in results if r["status"] == "pasted_unverified")
    dry_run = sum(1 for r in results if r["status"] == "dry_run")
    failed = sum(1 for r in results if r["status"] == "failed")
    skipped = sum(1 for r in results if str(r["status"]).startswith("skipped"))
    logger.info(
        "Done. sent=%s pasted=%s pasted_unverified=%s dry_run=%s failed=%s skipped=%s",
        sent,
        pasted,
        pasted_unverified,
        dry_run,
        failed,
        skipped,
    )
    logger.info("Result file: %s", args.results_csv)
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
