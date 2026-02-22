from __future__ import annotations

import os
from dataclasses import dataclass, asdict
from pathlib import Path
import yaml

CONFIG_PATH = Path(os.getenv("APP_CONFIG", "/data/config.yaml"))


@dataclass
class AppConfig:
    corp_id: str = ""
    agent_id: str = ""
    secret: str = ""
    root_path: str = "/data/inbox"
    rate_limit_per_sec: float = 1.0
    max_concurrency: int = 2


def load_config() -> AppConfig:
    if not CONFIG_PATH.exists():
        cfg = AppConfig()
        save_config(cfg)
        return cfg
    data = yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8")) or {}
    return AppConfig(**data)


def save_config(cfg: AppConfig) -> None:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        yaml.safe_dump(asdict(cfg), allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )


def update_config(partial: dict) -> AppConfig:
    cfg = load_config()
    for k, v in partial.items():
        if hasattr(cfg, k) and v is not None:
            setattr(cfg, k, v)
    save_config(cfg)
    return cfg
