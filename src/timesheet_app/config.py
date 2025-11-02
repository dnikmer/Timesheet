"""Configuration helpers for the Timesheet application."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from typing import Optional


APP_DIR = Path.home() / ".timesheet_app"
CONFIG_FILE = APP_DIR / "config.json"


@dataclass
class AppConfig:
    """Persisted configuration."""

    excel_path: Optional[str] = None

    @classmethod
    def load(cls) -> "AppConfig":
        """Load configuration from disk, returning defaults when missing."""

        if CONFIG_FILE.exists():
            try:
                data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    allowed = {item.name for item in fields(cls)}
                    filtered = {key: value for key, value in data.items() if key in allowed}
                    return cls(**filtered)
                return cls()
            except (json.JSONDecodeError, TypeError, ValueError):
                # Fall back to defaults if the file is corrupted.
                pass
        return cls()

    def save(self) -> None:
        """Persist configuration to disk."""

        APP_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")
