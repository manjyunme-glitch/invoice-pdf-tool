from __future__ import annotations

import os
import platform
import tempfile
from pathlib import Path
from typing import Tuple


CONFIG_DIR_FALLBACK_REASON = ""


def _prepare_config_dir(preferred: Path, fallback: Path) -> Tuple[Path, str]:
    try:
        preferred.mkdir(parents=True, exist_ok=True)
        return preferred, ""
    except OSError as preferred_error:
        try:
            fallback.mkdir(parents=True, exist_ok=True)
            return fallback, f"首选配置目录不可用，已切换到临时目录：{preferred_error}"
        except OSError as fallback_error:
            return preferred, f"配置目录和临时回退目录均不可用：{preferred_error}；{fallback_error}"


def get_config_dir() -> Path:
    global CONFIG_DIR_FALLBACK_REASON
    if platform.system() == "Windows":
        base = Path(os.environ.get("APPDATA", Path.home()))
        preferred = base / "InvoiceTool"
    else:
        preferred = Path.home() / ".invoice_tool"
    fallback = Path(tempfile.gettempdir()) / "InvoiceTool-fallback"
    config_dir, warning = _prepare_config_dir(preferred, fallback)
    CONFIG_DIR_FALLBACK_REASON = warning
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
ACTIVE_TASK_FILE = CONFIG_DIR / "active_task.json"
LOG_FILE = CONFIG_DIR / "app.log"
