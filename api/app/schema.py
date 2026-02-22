from __future__ import annotations

from pydantic import BaseModel
from typing import List, Optional


class ScanRequest(BaseModel):
    rootPath: str


class SendSelectedRequest(BaseModel):
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
