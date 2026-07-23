from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Optional, Tuple, Type, Union
from uuid import uuid4

from .logging_setup import logger


ExpectedJsonType = Optional[Union[Type[Any], Tuple[Type[Any], ...]]]


def quarantine_invalid_json(path: Path) -> Optional[Path]:
    if not path.exists():
        return None
    backup = path.with_name(
        f"{path.stem}.corrupt-{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex[:8]}"
        f"{path.suffix or '.json'}"
    )
    try:
        path.replace(backup)
    except OSError as exc:
        logger.error("隔离损坏的 %s 失败: %s", path.name, exc)
        return None
    logger.warning("损坏的 %s 已隔离为 %s", path.name, backup.name)
    return backup


def load_json(
    path: Path,
    default: Any,
    *,
    expected_type: ExpectedJsonType = None,
    quarantine_invalid: bool = False,
) -> Any:
    try:
        if path.exists():
            value = json.loads(path.read_text("utf-8"))
            if expected_type is not None and not isinstance(value, expected_type):
                expected_name = (
                    ", ".join(item.__name__ for item in expected_type)
                    if isinstance(expected_type, tuple)
                    else expected_type.__name__
                )
                raise ValueError(f"JSON 根节点类型无效，应为 {expected_name}")
            return value
    except (json.JSONDecodeError, UnicodeError, PermissionError, OSError, ValueError) as exc:
        logger.error(f"加载 {path.name} 失败: {exc}")
        if quarantine_invalid:
            quarantine_invalid_json(path)
    return default


def save_json(path: Path, data: Any) -> bool:
    temporary_path = path.with_name(f".{path.name}.{uuid4().hex}.tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with temporary_path.open("w", encoding="utf-8", newline="\n") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary_path, path)
        return True
    except (PermissionError, OSError, TypeError, ValueError) as exc:
        logger.error(f"保存 {path.name} 失败: {exc}")
        try:
            temporary_path.unlink(missing_ok=True)
        except OSError:
            pass
        return False
