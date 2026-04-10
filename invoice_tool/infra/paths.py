from __future__ import annotations

import os
import platform
from pathlib import Path


def get_config_dir() -> Path:
    if platform.system() == "Windows":
        base = Path(os.environ.get("APPDATA", Path.home()))
        config_dir = base / "InvoiceTool"
    else:
        config_dir = Path.home() / ".invoice_tool"
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir


def is_relative_to(path: Path, base: Path) -> bool:
    try:
        path.relative_to(base)
        return True
    except ValueError:
        return False


CONFIG_DIR = get_config_dir()
CONFIG_FILE = CONFIG_DIR / "config.json"
HISTORY_FILE = CONFIG_DIR / "history.json"
LOG_FILE = CONFIG_DIR / "app.log"
