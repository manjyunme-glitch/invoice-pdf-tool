from __future__ import annotations

import json
import os
import platform
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, Optional

from ..infra.privacy import redact_sensitive_text


def redact_diagnostic_text(value: object) -> str:
    return redact_sensitive_text(value)


def build_diagnostic_snapshot(
    *,
    app_version: str,
    capabilities: Mapping[str, bool],
    config_path: Path,
    history_path: Path,
    log_path: Path,
    recent_errors: Iterable[Mapping[str, Any]] = (),
    config_schema_version: int = 1,
    config_directory_fallback_reason: str = "",
) -> Dict[str, Any]:
    errors = []
    for error in list(recent_errors)[-20:]:
        message = error.get("message") or error.get("detail") or error.get("summary", "")
        errors.append(
            {
                "time": redact_diagnostic_text(error.get("time", "")),
                "level": redact_diagnostic_text(error.get("level", "ERROR")),
                "message": redact_diagnostic_text(message),
            }
        )
    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "application": {
            "version": app_version,
            "frozen": bool(getattr(sys, "frozen", False)),
            "config_schema_version": config_schema_version,
        },
        "runtime": {
            "python": platform.python_version(),
            "implementation": platform.python_implementation(),
            "platform": platform.platform(),
            "architecture": platform.machine(),
        },
        "capabilities": {str(key): bool(value) for key, value in capabilities.items()},
        "storage": {
            "config_exists": config_path.exists(),
            "history_exists": history_path.exists(),
            "log_exists": log_path.exists(),
            "log_size_bytes": log_path.stat().st_size if log_path.exists() else 0,
            "config_directory_fallback": bool(config_directory_fallback_reason),
            "config_directory_warning": redact_diagnostic_text(config_directory_fallback_reason),
        },
        "recent_errors": errors,
    }


def create_diagnostic_bundle(
    target: Path,
    snapshot: Mapping[str, Any],
    *,
    log_path: Optional[Path] = None,
    max_log_lines: int = 500,
) -> Path:
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary_fd, temporary_name = tempfile.mkstemp(
        prefix=f".{target.name}.",
        suffix=".tmp",
        dir=str(target.parent),
    )
    os.close(temporary_fd)
    temporary = Path(temporary_name)
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(
                "diagnostics.json",
                json.dumps(dict(snapshot), ensure_ascii=False, indent=2),
            )
            archive.writestr(
                "README.txt",
                "此诊断包不包含原始配置、历史记录、Excel 或 PDF。\n"
                "日志内容已尝试移除本地路径、网络路径和长数字标识。\n"
                "发送前仍建议人工检查文件内容。\n",
            )
            if log_path and log_path.exists():
                try:
                    lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
                    sanitized = "\n".join(redact_diagnostic_text(line) for line in lines[-max_log_lines:])
                    archive.writestr("sanitized-log.txt", sanitized)
                except OSError as exc:
                    archive.writestr("log-read-error.txt", redact_diagnostic_text(exc))
        os.replace(temporary, target)
    except Exception:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    return target
