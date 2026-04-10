from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .logging_setup import logger


def load_json(path: Path, default: Any) -> Any:
    try:
        if path.exists():
            return json.loads(path.read_text("utf-8"))
    except (json.JSONDecodeError, PermissionError, OSError) as exc:
        logger.error(f"加载 {path.name} 失败: {exc}")
    return default


def save_json(path: Path, data: Any) -> None:
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
    except (PermissionError, OSError) as exc:
        logger.error(f"保存 {path.name} 失败: {exc}")
