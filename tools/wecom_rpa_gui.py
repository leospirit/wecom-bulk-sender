#!/usr/bin/env python
from __future__ import annotations

import csv
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText


class RpaGuiApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("企业微信批量助手")
        self.root.geometry("1220x780")
        self.root.minsize(1060, 700)

        self.repo_root = Path(__file__).resolve().parents[1]
        self.sender_script = self.repo_root / "tools" / "wecom_rpa_sender.py"
        self.default_csv = self.repo_root / "tools" / "rpa_tasks.real.csv"
        self.config_path = self.repo_root / "run-logs" / "rpa-ui-config.json"

        self.proc: subprocess.Popen | None = None
        self.log_q: queue.Queue[tuple[str, str]] = queue.Queue()
        self.running = False
        self._auto_minimized = False
        self.progress_total = 0
        self.progress_current = 0

        self.csv_path = tk.StringVar(value=str(self.default_csv))
        self.send_mode = tk.StringVar(value="clipboard")
        self.main_title_re = tk.StringVar(value=r".*(WeCom|WXWork|企业微信).*")
        self.interval_sec = tk.StringVar(value="3")
        self.timeout_sec = tk.StringVar(value="12")
        self.max_retries = tk.StringVar(value="2")
        self.retry_delay_sec = tk.StringVar(value="1.0")
        self.stabilize_open_rounds = tk.StringVar(value="2")
        self.stabilize_focus_rounds = tk.StringVar(value="2")
        self.stabilize_send_rounds = tk.StringVar(value="1")
        self.open_chat_strategy = tk.StringVar(value="keyboard_first")
        self.resume_from = tk.StringVar(value="1")

        self.paste_only = tk.BooleanVar(value=True)
        self.no_chat_verify = tk.BooleanVar(value=True)
        self.resume_failed = tk.BooleanVar(value=False)
        self.stop_on_fail = tk.BooleanVar(value=True)
        self.skip_missing_image = tk.BooleanVar(value=True)
        self.debug_chat_text = tk.BooleanVar(value=False)

        self.results_csv = tk.StringVar(value=str(self.repo_root / "run-logs" / "rpa-results.csv"))
        self.log_file = tk.StringVar(value=str(self.repo_root / "run-logs" / "rpa-sender.log"))

        self.status_text = tk.StringVar(value="就绪")
        self.progress_text = tk.StringVar(value="0/0")
        self.cmd_preview = tk.StringVar(value="")
        self.tasks_info = tk.StringVar(value="未检查任务文件")
        self.mode_name = tk.StringVar(value="测试模式（仅粘贴）")
        self.mode_desc = tk.StringVar(value="不会发送消息，仅粘贴图片，适合联调")
        self.mode_buttons: dict[str, tk.Button] = {}
        self.custom_buttons: list[tk.Button] = []
        self.theme_mode = tk.StringVar(value="light")
        self.run_started_at: float | None = None
        self.summary_total = tk.StringVar(value="总数 0")
        self.summary_ok = tk.StringVar(value="成功 0")
        self.summary_failed = tk.StringVar(value="失败 0")
        self.summary_skipped = tk.StringVar(value="跳过 0")
        self.summary_elapsed = tk.StringVar(value="耗时 0s")
        self._done_ok = 0
        self._done_failed = 0
        self._done_skipped = 0

        self._apply_style()
        self._load_config()
        self._build_ui()
        self._bind_var_traces()
        self._refresh_cmd_preview()
        self.root.after(100, self._drain_logs)

    def _apply_style(self):
        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.style.configure("App.TFrame")
        self.style.configure("Card.TLabelframe", borderwidth=0, relief="flat")
        self.style.configure("Card.TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"))
        self.style.configure("App.TLabel", font=("Microsoft YaHei UI", 10))
        self.style.configure("Sub.TLabel", font=("Microsoft YaHei UI", 9))
        self.style.configure("State.TLabel", font=("Microsoft YaHei UI", 10, "bold"))
        self.style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"))
        self.style.configure("TNotebook.Tab", padding=(14, 8), font=("Microsoft YaHei UI", 9, "bold"))

    def _theme_tokens(self, mode: str) -> dict[str, str]:
        if mode == "dark":
            return {
                "bg": "#0b1220",
                "panel": "#111b2e",
                "card": "#16233b",
                "text": "#e5edf9",
                "sub": "#9bb0cf",
                "accent": "#4f8cff",
                "header": "#081329",
                "header_sub": "#8fb4ff",
                "log_bg": "#050b16",
                "log_fg": "#d6e4ff",
                "log_insert": "#8fb4ff",
                "entry_bg": "#0f1a2e",
                "entry_fg": "#dbe8ff",
                "prog_trough": "#1d2e4d",
            }
        return {
            "bg": "#f4f7fb",
            "panel": "#f4f7fb",
            "card": "#ffffff",
            "text": "#1f2a37",
            "sub": "#5b6675",
            "accent": "#0f4cdb",
            "header": "#0f4cdb",
            "header_sub": "#dbe7ff",
            "log_bg": "#0f172a",
            "log_fg": "#e5e7eb",
            "log_insert": "#93c5fd",
            "entry_bg": "#ffffff",
            "entry_fg": "#1f2a37",
            "prog_trough": "#d8e2f2",
        }

    def _apply_theme(self, mode: str):
        self.theme_mode.set(mode)
        t = self._theme_tokens(mode)

        self.root.configure(bg=t["bg"])
        self.style.configure("App.TFrame", background=t["bg"])
        self.style.configure("Card.TLabelframe", background=t["card"])
        self.style.configure("Card.TLabelframe.Label", background=t["card"], foreground=t["text"])
        self.style.configure("App.TLabel", background=t["bg"], foreground=t["text"])
        self.style.configure("Sub.TLabel", background=t["bg"], foreground=t["sub"])
        self.style.configure("State.TLabel", background=t["bg"], foreground=t["accent"])
        self.style.configure("TCheckbutton", background=t["bg"], foreground=t["text"])
        self.style.configure("TNotebook", background=t["bg"], borderwidth=0)
        self.style.configure("TNotebook.Tab", background=t["panel"], foreground=t["sub"])
        self.style.map("TNotebook.Tab", background=[("selected", t["card"])], foreground=[("selected", t["text"])])
        self.style.configure("Horizontal.TProgressbar", troughcolor=t["prog_trough"], background=t["accent"], bordercolor=t["prog_trough"])
        self.style.configure("TEntry", fieldbackground=t["entry_bg"], foreground=t["entry_fg"])
        self.style.configure("Cmd.TEntry", fieldbackground=t["entry_bg"], foreground=t["entry_fg"])

        if hasattr(self, "header"):
            self.header.configure(bg=t["header"])
            self.header_title.configure(bg=t["header"], fg="#ffffff")
            self.header_sub.configure(bg=t["header"], fg=t["header_sub"])
            self._apply_btn_palette(self.btn_theme, "#ffffff", "#edf2ff", "#dbe6ff", "#0f4cdb")

        if hasattr(self, "stats_frame"):
            self.stats_frame.configure(bg=t["bg"])
            chip_bg = "#1d2e4d" if mode == "dark" else "#ffffff"
            chip_fg = "#dce9ff" if mode == "dark" else "#1f2a37"
            for lbl in getattr(self, "stat_labels", []):
                lbl.configure(bg=chip_bg, fg=chip_fg)

        if hasattr(self, "log_box"):
            self.log_box.configure(bg=t["log_bg"], fg=t["log_fg"], insertbackground=t["log_insert"])

        for btn in self.custom_buttons:
            variant = getattr(btn, "_variant", "secondary")
            self._apply_variant_palette(btn, variant)

        self._sync_mode_button_from_state()
        self._refresh_theme_button_text()

    def _refresh_theme_button_text(self):
        if hasattr(self, "btn_theme"):
            self.btn_theme.configure(text="浅色" if self.theme_mode.get() == "dark" else "深色")

    def _toggle_theme(self):
        new_mode = "dark" if self.theme_mode.get() == "light" else "light"
        self._apply_theme(new_mode)
        self.status_text.set("已切换为深色主题" if new_mode == "dark" else "已切换为浅色主题")

    def _apply_btn_palette(self, btn: tk.Button, normal: str, hover: str, active: str, fg: str):
        btn._normal_bg = normal  # type: ignore[attr-defined]
        btn._hover_bg = hover  # type: ignore[attr-defined]
        btn._active_bg = active  # type: ignore[attr-defined]
        btn._text_fg = fg  # type: ignore[attr-defined]
        btn.configure(
            bg=normal,
            fg=fg,
            activebackground=active,
            activeforeground=fg,
            disabledforeground="#93a1b5",
        )

    def _variant_palette(self) -> dict[str, tuple[str, str, str, str]]:
        if self.theme_mode.get() == "dark":
            return {
                "primary": ("#4f8cff", "#3d7af0", "#2f67dc", "#ffffff"),
                "secondary": ("#213554", "#264061", "#2b4970", "#dce9ff"),
                "danger": ("#d14b58", "#bf3e4b", "#ab303d", "#ffffff"),
                "chip": ("#1e3558", "#26426a", "#2d4d78", "#dce9ff"),
            }
        return {
            "primary": ("#0f4cdb", "#0d45c8", "#0a3aab", "#ffffff"),
            "secondary": ("#e9eef8", "#dfe8f7", "#d3e0f4", "#1f2a37"),
            "danger": ("#ef4444", "#dc2626", "#b91c1c", "#ffffff"),
            "chip": ("#eef2ff", "#e1e9ff", "#d3ddff", "#1f2a37"),
        }

    def _apply_variant_palette(self, btn: tk.Button, variant: str):
        palette = self._variant_palette()
        normal, hover, active, fg = palette.get(variant, palette["secondary"])
        self._apply_btn_palette(btn, normal, hover, active, fg)

    def _make_button(self, parent, text: str, command, variant: str = "secondary", width: int | None = None) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=12,
            pady=8,
            font=("Microsoft YaHei UI", 10, "bold"),
            highlightthickness=0,
            width=width or 0,
        )
        btn._variant = variant  # type: ignore[attr-defined]
        self._apply_variant_palette(btn, variant)

        def _on_enter(_):
            if str(btn.cget("state")) != "disabled":
                btn.configure(bg=getattr(btn, "_hover_bg"))

        def _on_leave(_):
            if str(btn.cget("state")) != "disabled":
                btn.configure(bg=getattr(btn, "_normal_bg"))

        btn.bind("<Enter>", _on_enter)
        btn.bind("<Leave>", _on_leave)
        self.custom_buttons.append(btn)
        return btn

    def _set_mode_button_active(self, name: str):
        if not self.mode_buttons:
            return
        for key, btn in self.mode_buttons.items():
            if key == name:
                if self.theme_mode.get() == "dark":
                    self._apply_btn_palette(btn, "#4f8cff", "#3d7af0", "#2f67dc", "#ffffff")
                else:
                    self._apply_btn_palette(btn, "#0f4cdb", "#0d45c8", "#0a3aab", "#ffffff")
            else:
                if self.theme_mode.get() == "dark":
                    self._apply_btn_palette(btn, "#1e3558", "#26426a", "#2d4d78", "#dce9ff")
                else:
                    self._apply_btn_palette(btn, "#eef2ff", "#e1e9ff", "#d3ddff", "#1f2a37")

    def _sync_mode_button_from_state(self):
        if not self.mode_buttons:
            return
        if self.resume_failed.get():
            key = "resume_failed"
            self.mode_name.set("续跑失败项")
            self.mode_desc.set("仅处理上次失败/未执行任务，避免重复发送")
        elif self.paste_only.get():
            key = "test"
            self.mode_name.set("测试模式（仅粘贴）")
            self.mode_desc.set("不会发送消息，仅粘贴图片，适合联调")
        else:
            key = "send"
            self.mode_name.set("正式发送（稳态）")
            self.mode_desc.set("执行发送，失败会自动重试，适合正式批量")
        self._set_mode_button_active(key)

    def _build_ui(self):
        root = self.root
        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(2, weight=1)

        self._build_header(root)
        self._build_status_bar(root)

        main = ttk.Frame(root, style="App.TFrame")
        main.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 10))
        main.grid_columnconfigure(0, weight=0)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(0, weight=1)

        left = ttk.Frame(main, style="App.TFrame")
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        left.grid_columnconfigure(0, weight=1)

        right = ttk.Frame(main, style="App.TFrame")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(0, weight=1)

        self._build_step_cards(left)
        self._build_tabs(right)
        self._apply_theme(self.theme_mode.get())

    def _build_header(self, root: tk.Tk):
        self.header = tk.Frame(root, bg="#0f4cdb", height=82)
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.grid_columnconfigure(0, weight=1)

        self.header_title = tk.Label(
            self.header,
            text="企业微信批量助手",
            bg="#0f4cdb",
            fg="#ffffff",
            font=("Microsoft YaHei UI", 15, "bold"),
        )
        self.header_title.grid(row=0, column=0, sticky="w", padx=14, pady=(12, 0))
        self.header_sub = tk.Label(
            self.header,
            text="流程：1 选择任务文件  ->  2 选择运行模式  ->  3 执行",
            bg="#0f4cdb",
            fg="#dbe7ff",
            font=("Microsoft YaHei UI", 9),
        )
        self.header_sub.grid(row=1, column=0, sticky="w", padx=14, pady=(4, 12))

        self.btn_theme = self._make_button(self.header, "深色", self._toggle_theme, "secondary")
        self.btn_theme.grid(row=0, column=1, rowspan=2, sticky="e", padx=14, pady=16)

    def _build_status_bar(self, root: tk.Tk):
        row = ttk.Frame(root, style="App.TFrame")
        row.grid(row=1, column=0, sticky="ew", padx=12, pady=(10, 6))
        row.grid_columnconfigure(2, weight=1)

        ttk.Label(row, text="状态", style="Sub.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(row, textvariable=self.status_text, style="State.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 14))

        self.progress = ttk.Progressbar(row, orient="horizontal", mode="determinate")
        self.progress.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        ttk.Label(row, textvariable=self.progress_text, style="Sub.TLabel").grid(row=0, column=3, sticky="e")

        self.stats_frame = tk.Frame(row, bg="#f4f7fb")
        self.stats_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        for i in range(5):
            self.stats_frame.grid_columnconfigure(i, weight=1)

        self.stat_labels: list[tk.Label] = []
        self.stat_labels.append(self._make_stat_chip(self.stats_frame, self.summary_total, 0))
        self.stat_labels.append(self._make_stat_chip(self.stats_frame, self.summary_ok, 1))
        self.stat_labels.append(self._make_stat_chip(self.stats_frame, self.summary_failed, 2))
        self.stat_labels.append(self._make_stat_chip(self.stats_frame, self.summary_skipped, 3))
        self.stat_labels.append(self._make_stat_chip(self.stats_frame, self.summary_elapsed, 4))

    def _make_stat_chip(self, parent: tk.Frame, var: tk.StringVar, col: int) -> tk.Label:
        lbl = tk.Label(
            parent,
            textvariable=var,
            bg="#ffffff",
            fg="#1f2a37",
            font=("Microsoft YaHei UI", 9, "bold"),
            padx=10,
            pady=6,
            relief="flat",
            bd=0,
        )
        lbl.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 8, 0))
        return lbl

    def _build_step_cards(self, parent: ttk.Frame):
        self._build_step1(parent)
        self._build_step2(parent)
        self._build_step3(parent)

    def _build_step1(self, parent: ttk.Frame):
        card = ttk.LabelFrame(parent, text="第 1 步：选择任务文件", style="Card.TLabelframe")
        card.grid(row=0, column=0, sticky="ew")
        card.grid_columnconfigure(1, weight=1)

        ttk.Label(card, text="任务 CSV", style="App.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(card, textvariable=self.csv_path, width=42).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=8)
        self._make_button(card, "▦ 选择", self._pick_csv, "secondary").grid(row=0, column=2, padx=(0, 10), pady=8)

        self._make_button(card, "✓ 检查任务", self.inspect_tasks, "secondary").grid(row=1, column=0, padx=10, pady=(0, 8), sticky="w")
        ttk.Label(card, textvariable=self.tasks_info, style="Sub.TLabel").grid(row=1, column=1, columnspan=2, sticky="w", pady=(0, 8))

    def _build_step2(self, parent: ttk.Frame):
        card = ttk.LabelFrame(parent, text="第 2 步：选择运行模式", style="Card.TLabelframe")
        card.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        btns = ttk.Frame(card, style="App.TFrame")
        btns.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        for i in range(3):
            btns.grid_columnconfigure(i, weight=1)

        b_test = self._make_button(btns, "◇ 测试（仅粘贴）", lambda: self.apply_preset("test"), "chip")
        b_send = self._make_button(btns, "▶ 正式发送", lambda: self.apply_preset("send"), "chip")
        b_resume = self._make_button(btns, "↻ 续跑失败项", lambda: self.apply_preset("resume_failed"), "chip")
        b_test.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        b_send.grid(row=0, column=1, sticky="ew", padx=6)
        b_resume.grid(row=0, column=2, sticky="ew", padx=(6, 0))
        self.mode_buttons = {"test": b_test, "send": b_send, "resume_failed": b_resume}
        self._sync_mode_button_from_state()

        ttk.Label(card, textvariable=self.mode_name, style="State.TLabel").grid(row=1, column=0, sticky="w", padx=10)
        ttk.Label(card, textvariable=self.mode_desc, style="Sub.TLabel").grid(row=2, column=0, sticky="w", padx=10, pady=(2, 10))

    def _build_step3(self, parent: ttk.Frame):
        card = ttk.LabelFrame(parent, text="第 3 步：执行", style="Card.TLabelframe")
        card.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        card.grid_columnconfigure(0, weight=1)

        toggles = ttk.Frame(card, style="App.TFrame")
        toggles.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        ttk.Checkbutton(toggles, text="仅粘贴不发送", variable=self.paste_only).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(toggles, text="只续跑失败", variable=self.resume_failed).grid(row=0, column=1, sticky="w", padx=(18, 0))

        actions = ttk.Frame(card, style="App.TFrame")
        actions.grid(row=1, column=0, sticky="ew", padx=10, pady=(6, 10))
        for i in range(3):
            actions.grid_columnconfigure(i, weight=1)

        self.btn_run = self._make_button(actions, "▶ 开始运行", self.run_task, "primary")
        self.btn_run.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        self.btn_quick = self._make_button(actions, "↻ 快速测试", self.quick_test, "secondary")
        self.btn_quick.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        self.btn_stop = self._make_button(actions, "■ 停止", self.stop_task, "danger")
        self.btn_stop.configure(state="disabled")
        self.btn_stop.grid(row=1, column=1, sticky="ew", padx=6)
        self._make_button(actions, "⌫ 清空日志", self.clear_log, "secondary").grid(row=1, column=2, sticky="ew", padx=(6, 0))

    def _build_tabs(self, parent: ttk.Frame):
        tabs = ttk.Notebook(parent)
        tabs.grid(row=0, column=0, sticky="nsew")

        tab_log = ttk.Frame(tabs, style="App.TFrame")
        tab_log.grid_columnconfigure(0, weight=1)
        tab_log.grid_rowconfigure(1, weight=1)

        cmd_card = ttk.LabelFrame(tab_log, text="命令预览", style="Card.TLabelframe")
        cmd_card.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
        cmd_card.grid_columnconfigure(0, weight=1)
        self.cmd_entry = ttk.Entry(cmd_card, textvariable=self.cmd_preview, style="Cmd.TEntry")
        self.cmd_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self._make_button(cmd_card, "⎘ 复制", self._copy_cmd, "secondary").grid(row=0, column=1, padx=(0, 10), pady=10)

        log_card = ttk.LabelFrame(tab_log, text="运行日志", style="Card.TLabelframe")
        log_card.grid(row=1, column=0, sticky="nsew", padx=6, pady=(4, 6))
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(0, weight=1)

        self.log_box = ScrolledText(
            log_card,
            wrap="word",
            bg="#0f172a",
            fg="#e5e7eb",
            insertbackground="#93c5fd",
            font=("Consolas", 10),
            relief="flat",
        )
        self.log_box.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        tab_adv = ttk.Frame(tabs, style="App.TFrame")
        tab_adv.grid_columnconfigure(1, weight=1)

        row = 0
        def add_label_entry(label: str, var: tk.Variable, width: int = 30):
            nonlocal row
            ttk.Label(tab_adv, text=label, style="App.TLabel").grid(row=row, column=0, sticky="w", padx=10, pady=7)
            ttk.Entry(tab_adv, textvariable=var, width=width).grid(row=row, column=1, sticky="ew", padx=(0, 10), pady=7)
            row += 1

        ttk.Label(tab_adv, text="发送方式", style="App.TLabel").grid(row=row, column=0, sticky="w", padx=10, pady=7)
        ttk.OptionMenu(tab_adv, self.send_mode, self.send_mode.get(), "clipboard", "dialog", "auto").grid(row=row, column=1, sticky="w", padx=(0, 10), pady=7)
        row += 1

        add_label_entry("窗口标题正则", self.main_title_re)
        add_label_entry("任务间隔秒", self.interval_sec)
        add_label_entry("超时秒", self.timeout_sec)
        add_label_entry("最大重试次数", self.max_retries)
        add_label_entry("重试间隔秒", self.retry_delay_sec)
        add_label_entry("稳态轮次-开会话", self.stabilize_open_rounds)
        add_label_entry("稳态轮次-聚焦输入", self.stabilize_focus_rounds)
        add_label_entry("稳态轮次-发送前", self.stabilize_send_rounds)
        ttk.Label(tab_adv, text="打开会话策略", style="App.TLabel").grid(row=row, column=0, sticky="w", padx=10, pady=7)
        ttk.OptionMenu(
            tab_adv,
            self.open_chat_strategy,
            self.open_chat_strategy.get(),
            "keyboard_first",
            "click_first",
            "hybrid",
        ).grid(row=row, column=1, sticky="w", padx=(0, 10), pady=7)
        row += 1
        add_label_entry("从第几条开始", self.resume_from)

        ttk.Checkbutton(tab_adv, text="关闭会话校验", variable=self.no_chat_verify).grid(row=row, column=0, sticky="w", padx=10, pady=6)
        ttk.Checkbutton(tab_adv, text="失败即停", variable=self.stop_on_fail).grid(row=row, column=1, sticky="w", pady=6)
        row += 1
        ttk.Checkbutton(tab_adv, text="缺图跳过", variable=self.skip_missing_image).grid(row=row, column=0, sticky="w", padx=10, pady=6)
        ttk.Checkbutton(tab_adv, text="调试会话文本", variable=self.debug_chat_text).grid(row=row, column=1, sticky="w", pady=6)
        row += 1

        add_label_entry("结果文件", self.results_csv)
        self._make_button(tab_adv, "↗ 打开结果文件", lambda: self._open_file(self.results_csv.get()), "secondary").grid(
            row=row - 1, column=2, padx=(0, 10), pady=7
        )

        add_label_entry("日志文件", self.log_file)
        self._make_button(tab_adv, "↗ 打开日志文件", lambda: self._open_file(self.log_file.get()), "secondary").grid(
            row=row - 1, column=2, padx=(0, 10), pady=7
        )

        tabs.add(tab_log, text="运行与日志")
        tabs.add(tab_adv, text="高级参数")

        self._log("界面已就绪")
        self._log(f"发送脚本: {self.sender_script}")

    def _var_map(self) -> dict[str, tk.Variable]:
        return {
            "csv_path": self.csv_path,
            "send_mode": self.send_mode,
            "main_title_re": self.main_title_re,
            "interval_sec": self.interval_sec,
            "timeout_sec": self.timeout_sec,
            "max_retries": self.max_retries,
            "retry_delay_sec": self.retry_delay_sec,
            "stabilize_open_rounds": self.stabilize_open_rounds,
            "stabilize_focus_rounds": self.stabilize_focus_rounds,
            "stabilize_send_rounds": self.stabilize_send_rounds,
            "open_chat_strategy": self.open_chat_strategy,
            "resume_from": self.resume_from,
            "paste_only": self.paste_only,
            "no_chat_verify": self.no_chat_verify,
            "resume_failed": self.resume_failed,
            "stop_on_fail": self.stop_on_fail,
            "skip_missing_image": self.skip_missing_image,
            "debug_chat_text": self.debug_chat_text,
            "results_csv": self.results_csv,
            "log_file": self.log_file,
            "theme_mode": self.theme_mode,
        }

    def _bind_var_traces(self):
        for var in self._var_map().values():
            try:
                var.trace_add("write", lambda *_: self._refresh_cmd_preview())
            except Exception:
                pass

    def _load_config(self):
        if not self.config_path.exists():
            return
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception:
            return
        for name, var in self._var_map().items():
            if name in data:
                try:
                    var.set(data[name])
                except Exception:
                    pass

    def _save_config(self):
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        data = {name: var.get() for name, var in self._var_map().items()}
        try:
            self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _pick_csv(self):
        path = filedialog.askopenfilename(
            title="选择任务 CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=str(self.repo_root / "tools"),
        )
        if path:
            self.csv_path.set(path)
            self.inspect_tasks()

    def inspect_tasks(self):
        csv_path = Path(self.csv_path.get().strip())
        if not csv_path.exists():
            self.tasks_info.set("文件不存在")
            return
        try:
            with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                total = 0
                valid = 0
                for row in reader:
                    total += 1
                    if (row.get("parent_name") or "").strip() and (row.get("image_path") or "").strip():
                        valid += 1
            self.tasks_info.set(f"共 {total} 行，可执行 {valid} 行")
            self.summary_total.set(f"总数 {valid}")
            self.status_text.set("任务文件检查完成")
        except Exception as e:
            self.tasks_info.set(f"读取失败: {e}")

    def _set_running(self, running: bool):
        self.running = running
        self.btn_run.configure(state="disabled" if running else "normal")
        self.btn_stop.configure(state="normal" if running else "disabled")
        if hasattr(self, "btn_quick"):
            self.btn_quick.configure(state="disabled" if running else "normal")
        if running:
            self.status_text.set("运行中")
        elif self.status_text.get() == "运行中":
            self.status_text.set("就绪")
        if not running:
            self._save_config()

    def _log(self, msg: str):
        self.log_box.insert("end", msg.rstrip() + "\n")
        self.log_box.see("end")

    def clear_log(self):
        self.log_box.delete("1.0", "end")

    def _copy_cmd(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.cmd_preview.get())
        self.status_text.set("命令已复制")

    def _open_file(self, p: str):
        path = Path((p or "").strip())
        if not path.exists():
            messagebox.showwarning("文件不存在", str(path))
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def _drain_logs(self):
        try:
            while True:
                kind, msg = self.log_q.get_nowait()
                if kind == "line":
                    self._log(msg)
                    self._update_progress_from_line(msg)
                    self._parse_runtime_line(msg)
                elif kind == "done":
                    self._log(msg)
                    self._on_run_finished(msg)
                    self._set_running(False)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_logs)

    def _update_progress_from_line(self, line: str):
        m = re.search(r"\[(\d+)/(\d+)", line)
        if not m:
            return
        cur = int(m.group(1))
        total = int(m.group(2))
        self.progress_total = max(self.progress_total, total)
        self.progress_current = max(self.progress_current, cur)
        self.progress.configure(maximum=max(1, self.progress_total), value=self.progress_current)
        self.progress_text.set(f"{self.progress_current}/{self.progress_total}")

    def _reset_summary(self):
        self._done_ok = 0
        self._done_failed = 0
        self._done_skipped = 0
        self._done_pasted = 0
        self._done_pasted_unverified = 0
        self._done_dry = 0
        self.summary_total.set("总数 0")
        self.summary_ok.set("成功 0")
        self.summary_failed.set("失败 0")
        self.summary_skipped.set("跳过 0")
        self.summary_elapsed.set("耗时 0s")

    def _parse_runtime_line(self, line: str):
        m_loaded = re.search(r"Loaded\s+(\d+)\s+task", line)
        if m_loaded:
            total = int(m_loaded.group(1))
            self.summary_total.set(f"总数 {total}")
            self.progress_total = max(self.progress_total, total)
            if self.progress_total > 0:
                self.progress.configure(maximum=self.progress_total)

        m_done_new = re.search(
            r"Done\.\s+sent=(\d+)\s+pasted=(\d+)(?:\s+pasted_unverified=(\d+))?\s+dry_run=(\d+)\s+failed=(\d+)\s+skipped=(\d+)",
            line,
        )
        if m_done_new:
            sent = int(m_done_new.group(1))
            pasted = int(m_done_new.group(2))
            pasted_unverified = int(m_done_new.group(3) or 0)
            dry_run = int(m_done_new.group(4))
            failed = int(m_done_new.group(5))
            skipped = int(m_done_new.group(6))
            self._done_ok = sent
            self._done_pasted = pasted
            self._done_pasted_unverified = pasted_unverified
            self._done_dry = dry_run
            self._done_failed = failed
            self._done_skipped = skipped
            if pasted > 0 and sent == 0:
                if pasted_unverified > 0:
                    self.summary_ok.set(f"仅粘贴 {pasted}（未验证 {pasted_unverified}）")
                else:
                    self.summary_ok.set(f"仅粘贴 {pasted}")
            elif pasted > 0 and sent > 0:
                if pasted_unverified > 0:
                    self.summary_ok.set(f"发送 {sent} / 粘贴 {pasted}（未验证 {pasted_unverified}）")
                else:
                    self.summary_ok.set(f"发送 {sent} / 粘贴 {pasted}")
            else:
                self.summary_ok.set(f"发送 {sent}")
            self.summary_failed.set(f"失败 {failed}")
            self.summary_skipped.set(f"跳过 {skipped}")
            return

        m_done_old = re.search(r"Done\.\s+ok=(\d+)\s+failed=(\d+)(?:\s+skipped=(\d+))?", line)
        if m_done_old:
            self._done_ok = int(m_done_old.group(1))
            self._done_failed = int(m_done_old.group(2))
            self._done_skipped = int(m_done_old.group(3) or 0)
            self.summary_ok.set(f"发送 {self._done_ok}")
            self.summary_failed.set(f"失败 {self._done_failed}")
            self.summary_skipped.set(f"跳过 {self._done_skipped}")

    def _on_run_finished(self, done_line: str):
        elapsed = 0.0
        if self.run_started_at is not None:
            elapsed = max(0.0, time.time() - self.run_started_at)
        self.summary_elapsed.set(f"耗时 {int(elapsed)}s")
        if self._auto_minimized:
            try:
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
            except Exception:
                pass
            self._auto_minimized = False

        rc_match = re.search(r"退出码:\s*(\d+)", done_line)
        rc = int(rc_match.group(1)) if rc_match else 1

        title = "运行完成" if rc == 0 else "运行结束（有异常）"
        msg = (
            f"总数: {self.summary_total.get().replace('总数 ', '')}\n"
            f"发送成功: {self._done_ok}\n"
            f"仅粘贴: {self._done_pasted}\n"
            f"粘贴未验证: {getattr(self, '_done_pasted_unverified', 0)}\n"
            f"dry-run: {self._done_dry}\n"
            f"失败: {self._done_failed}\n"
            f"跳过: {self._done_skipped}\n"
            f"耗时: {int(elapsed)} 秒"
        )
        self.status_text.set(
            f"完成：成功{self._done_ok} 失败{self._done_failed} 跳过{self._done_skipped}"
        )
        if self._done_failed > 0 or rc != 0:
            messagebox.showwarning(title, msg)
        else:
            messagebox.showinfo(title, msg)

    def _num_value(self, s: str, kind: str) -> str:
        raw = (s or "").strip()
        if kind == "int":
            int(raw)
        else:
            float(raw)
        return raw

    def _build_cmd(self) -> list[str]:
        csv_path = self.csv_path.get().strip()
        if not csv_path:
            raise ValueError("任务 CSV 不能为空")
        if not Path(csv_path).exists():
            raise ValueError(f"任务 CSV 不存在: {csv_path}")
        if not self.sender_script.exists():
            raise ValueError(f"发送脚本不存在: {self.sender_script}")

        interval = self._num_value(self.interval_sec.get(), "float")
        timeout = self._num_value(self.timeout_sec.get(), "float")
        max_retries = self._num_value(self.max_retries.get(), "int")
        retry_delay = self._num_value(self.retry_delay_sec.get(), "float")
        stabilize_open_rounds = self._num_value(self.stabilize_open_rounds.get(), "int")
        stabilize_focus_rounds = self._num_value(self.stabilize_focus_rounds.get(), "int")
        stabilize_send_rounds = self._num_value(self.stabilize_send_rounds.get(), "int")
        resume_from = self._num_value(self.resume_from.get(), "int")

        cmd = [
            sys.executable,
            str(self.sender_script),
            "--tasks-csv",
            csv_path,
            "--send-mode",
            self.send_mode.get().strip() or "clipboard",
            "--main-title-re",
            self.main_title_re.get().strip() or r".*(WeCom|WXWork|企业微信).*",
            "--interval-sec",
            interval,
            "--timeout-sec",
            timeout,
            "--max-retries",
            max_retries,
            "--retry-delay-sec",
            retry_delay,
            "--stabilize-open-rounds",
            stabilize_open_rounds,
            "--stabilize-focus-rounds",
            stabilize_focus_rounds,
            "--stabilize-send-rounds",
            stabilize_send_rounds,
            "--open-chat-strategy",
            self.open_chat_strategy.get().strip() or "keyboard_first",
            "--resume-from",
            resume_from,
            "--results-csv",
            self.results_csv.get().strip(),
            "--log-file",
            self.log_file.get().strip(),
        ]

        if self.paste_only.get():
            cmd.append("--paste-only")
        if self.no_chat_verify.get():
            cmd.append("--no-chat-verify")
        if self.resume_failed.get():
            cmd.append("--resume-failed")
            cmd.extend(["--resume-results-csv", self.results_csv.get().strip()])
        if self.stop_on_fail.get():
            cmd.append("--stop-on-fail")
        if self.skip_missing_image.get():
            cmd.append("--skip-missing-image")
        if self.debug_chat_text.get():
            cmd.append("--debug-chat-text")
        return cmd

    def _refresh_cmd_preview(self):
        try:
            cmd = self._build_cmd()
            self.cmd_preview.set(" ".join(f'"{c}"' if " " in c else c for c in cmd))
        except Exception:
            self.cmd_preview.set("")
        self._sync_mode_button_from_state()

    def apply_preset(self, name: str):
        if name == "test":
            self.send_mode.set("clipboard")
            self.paste_only.set(True)
            self.no_chat_verify.set(True)
            self.resume_failed.set(False)
            self.stop_on_fail.set(True)
            self.skip_missing_image.set(True)
            self.debug_chat_text.set(False)
            self.interval_sec.set("3")
            self.timeout_sec.set("12")
            self.max_retries.set("2")
            self.retry_delay_sec.set("1.0")
            self.stabilize_open_rounds.set("1")
            self.stabilize_focus_rounds.set("1")
            self.stabilize_send_rounds.set("1")
            self.open_chat_strategy.set("keyboard_first")
            self.resume_from.set("1")
            self.mode_name.set("测试模式（仅粘贴）")
            self.mode_desc.set("不会发送消息，仅粘贴图片，适合联调")
        elif name == "send":
            self.send_mode.set("clipboard")
            self.paste_only.set(False)
            self.no_chat_verify.set(True)
            self.resume_failed.set(False)
            self.stop_on_fail.set(True)
            self.skip_missing_image.set(True)
            self.debug_chat_text.set(False)
            self.interval_sec.set("3")
            self.timeout_sec.set("15")
            self.max_retries.set("3")
            self.retry_delay_sec.set("1.5")
            self.stabilize_open_rounds.set("2")
            self.stabilize_focus_rounds.set("2")
            self.stabilize_send_rounds.set("1")
            self.open_chat_strategy.set("keyboard_first")
            self.resume_from.set("1")
            self.mode_name.set("正式发送（稳态）")
            self.mode_desc.set("执行发送，失败会自动重试，适合正式批量")
        elif name == "resume_failed":
            self.send_mode.set("clipboard")
            self.paste_only.set(False)
            self.no_chat_verify.set(True)
            self.resume_failed.set(True)
            self.stop_on_fail.set(False)
            self.skip_missing_image.set(True)
            self.debug_chat_text.set(False)
            self.interval_sec.set("3")
            self.timeout_sec.set("15")
            self.max_retries.set("3")
            self.retry_delay_sec.set("1.5")
            self.stabilize_open_rounds.set("2")
            self.stabilize_focus_rounds.set("2")
            self.stabilize_send_rounds.set("1")
            self.open_chat_strategy.set("keyboard_first")
            self.resume_from.set("1")
            self.mode_name.set("续跑失败项")
            self.mode_desc.set("仅处理上次失败/未执行任务，避免重复发送")
        self._refresh_cmd_preview()
        self.status_text.set("已应用模式预设")
        self._sync_mode_button_from_state()

    def quick_test(self):
        self.apply_preset("test")
        self.run_task()

    def run_task(self):
        if self.running:
            return

        if self.paste_only.get():
            ok = messagebox.askyesno(
                "仅粘贴模式",
                "当前是“仅粘贴不发送”模式。\n运行后不会发送到家长，只会尝试粘贴图片。\n是否继续？",
            )
            if not ok:
                return
        else:
            ok = messagebox.askyesno(
                "正式发送确认",
                "当前是“正式发送”模式。\n运行后会实际发送给家长。\n是否继续？",
            )
            if not ok:
                return

        try:
            cmd = self._build_cmd()
        except Exception as e:
            messagebox.showerror("参数错误", str(e))
            return

        self.progress_total = 0
        self.progress_current = 0
        self.run_started_at = time.time()
        self._reset_summary()
        self.progress.configure(value=0, maximum=1)
        self.progress_text.set("0/0")
        self._refresh_cmd_preview()
        self._log("=" * 90)
        self._log("执行命令:")
        self._log(self.cmd_preview.get())
        try:
            if self.root.state() != "iconic":
                self.root.iconify()
                self._auto_minimized = True
        except Exception:
            self._auto_minimized = False
        self._set_running(True)

        def _worker():
            try:
                self.proc = subprocess.Popen(
                    cmd,
                    cwd=str(self.repo_root),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )
                assert self.proc.stdout is not None
                for line in self.proc.stdout:
                    self.log_q.put(("line", line.rstrip("\n")))
                rc = self.proc.wait()
                self.log_q.put(("done", f"[UI] 进程结束，退出码: {rc}"))
            except Exception as e:
                self.log_q.put(("done", f"[UI] 启动失败: {e}"))
            finally:
                self.proc = None

        threading.Thread(target=_worker, daemon=True).start()

    def stop_task(self):
        if not self.proc:
            return
        try:
            self.proc.terminate()
            self._log("[UI] 已发送停止信号")
        except Exception as e:
            self._log(f"[UI] 停止失败: {e}")

    def on_close(self):
        if self.running and self.proc:
            if not messagebox.askyesno("退出", "任务仍在运行，确认退出？"):
                return
            try:
                self.proc.terminate()
            except Exception:
                pass
        self._save_config()
        self.root.destroy()


def main():
    root = tk.Tk()
    app = RpaGuiApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
