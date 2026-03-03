from __future__ import annotations

from pydantic import BaseModel
from typing import List, Optional, Literal


class ScanRequest(BaseModel):
    rootPath: str


class SendSelectedRequest(BaseModel):
    taskIds: List[int]


class DeleteSelectedRequest(BaseModel):
    taskIds: List[int]


class AutoWatchRequest(BaseModel):
    enabled: bool


class ConfigUpdateRequest(BaseModel):
    corp_id: Optional[str] = None
    agent_id: Optional[str] = None
    secret: Optional[str] = None
    root_path: Optional[str] = None
    rate_limit_per_sec: Optional[float] = None
    max_concurrency: Optional[int] = None


class RpaStartRequest(BaseModel):
    tasks_csv: str = "tools/rpa_tasks.real.csv"
    wecom_exe: Optional[str] = None
    main_title_re: str = r".*(WeCom|WXWork|企业微信).*"
    send_mode: Literal["clipboard", "dialog", "auto"] = "clipboard"
    dry_run: bool = False
    paste_only: bool = True
    no_chat_verify: bool = True
    interval_sec: float = 3.0
    timeout_sec: float = 20.0
    max_retries: int = 2
    retry_delay_sec: float = 1.0
    stabilize_open_rounds: int = 2
    stabilize_focus_rounds: int = 2
    stabilize_send_rounds: int = 1
    open_chat_strategy: Literal["keyboard_first", "click_first", "hybrid"] = "keyboard_first"
    resume_from: int = 1
    resume_failed: bool = False
    stop_on_fail: bool = True
    skip_missing_image: bool = True
    debug_chat_text: bool = False
    results_csv: str = "run-logs/rpa-results.csv"
    log_file: str = "run-logs/rpa-sender.log"
