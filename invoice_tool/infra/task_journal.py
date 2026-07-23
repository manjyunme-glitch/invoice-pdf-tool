from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional
from uuid import uuid4

from .logging_setup import logger
from .storage import load_json, save_json


class TaskJournal:
    """A single-task, crash-recovery journal stored outside user work folders."""

    SCHEMA_VERSION = 1

    def __init__(self, path: Path) -> None:
        self.path = path
        self._lock = threading.Lock()
        self._data: Optional[Dict[str, Any]] = None

    @staticmethod
    def _now() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _quarantine_invalid(self) -> Optional[Path]:
        if not self.path.exists():
            return None
        quarantine_path = self.path.with_name(
            f"{self.path.stem}.invalid-{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex[:8]}"
            f"{self.path.suffix or '.json'}"
        )
        try:
            self.path.replace(quarantine_path)
        except OSError as exc:
            logger.error("隔离损坏的任务恢复日志失败，已阻止新文件任务: %s", exc)
            return None
        logger.warning("损坏的任务恢复日志已隔离为 %s", quarantine_path.name)
        return quarantine_path

    def load(self) -> Optional[Dict[str, Any]]:
        with self._lock:
            value = load_json(self.path, None)
            if value is None:
                if self.path.exists():
                    self._quarantine_invalid()
                self._data = None
                return None
            if not isinstance(value, dict):
                logger.error("任务恢复日志根节点结构无效")
                self._quarantine_invalid()
                self._data = None
                return None
            schema_version = value.get("schema_version")
            if isinstance(schema_version, int) and schema_version > self.SCHEMA_VERSION:
                logger.error("任务恢复日志来自较新版本，已保留并阻止新文件任务")
                self._data = None
                return None
            if schema_version != self.SCHEMA_VERSION or not str(value.get("task_id", "")).strip():
                logger.error("任务恢复日志结构无效")
                self._quarantine_invalid()
                self._data = None
                return None
            self._data = value
            return dict(value)

    def begin(self, task_type: str, folder: Path, metadata: Optional[Dict[str, Any]] = None) -> str:
        task_id = uuid4().hex
        now = self._now()
        data: Dict[str, Any] = {
            "schema_version": self.SCHEMA_VERSION,
            "task_id": task_id,
            "type": task_type,
            "folder": str(folder),
            "started_at": now,
            "updated_at": now,
            "moves": [],
            "report_files": [],
            "report_entries": [],
        }
        if metadata:
            data["metadata"] = metadata
        with self._lock:
            if self.path.exists():
                raise OSError("存在尚未归档的任务恢复日志，请先解决历史记录保存问题")
            if not save_json(self.path, data):
                raise OSError("无法创建任务恢复日志，任务未启动")
            self._data = data
        return task_id

    def _update(self, task_id: str, key: str, value: Any) -> bool:
        with self._lock:
            if self._data is None:
                loaded = load_json(self.path, None)
                if isinstance(loaded, dict):
                    self._data = loaded
            if not self._data or self._data.get("task_id") != task_id:
                logger.error("任务恢复日志与当前任务不匹配")
                return False
            items = self._data.setdefault(key, [])
            if not isinstance(items, list):
                logger.error("任务恢复日志字段 %s 结构无效", key)
                return False
            items.append(value)
            previous_updated_at = self._data.get("updated_at")
            self._data["updated_at"] = self._now()
            if save_json(self.path, self._data):
                return True
            items.pop()
            self._data["updated_at"] = previous_updated_at
            return False

    def record_move(self, task_id: str, move: Dict[str, Any]) -> bool:
        return self._update(task_id, "moves", dict(move))

    def record_report(self, task_id: str, entry: Dict[str, Any]) -> bool:
        path = str(entry.get("path", ""))
        with self._lock:
            if self._data is None:
                loaded = load_json(self.path, None)
                if isinstance(loaded, dict):
                    self._data = loaded
            if not self._data or self._data.get("task_id") != task_id:
                logger.error("任务恢复日志与当前任务不匹配")
                return False
            reports = self._data.setdefault("report_files", [])
            entries = self._data.setdefault("report_entries", [])
            if not isinstance(reports, list) or not isinstance(entries, list):
                logger.error("任务恢复日志的报告字段结构无效")
                return False
            reports.append(path)
            entries.append(dict(entry))
            previous_updated_at = self._data.get("updated_at")
            self._data["updated_at"] = self._now()
            if save_json(self.path, self._data):
                return True
            reports.pop()
            entries.pop()
            self._data["updated_at"] = previous_updated_at
            return False

    def clear(self, task_id: Optional[str] = None) -> bool:
        with self._lock:
            if task_id:
                current = self._data
                if current is None and self.path.exists():
                    loaded = load_json(self.path, None)
                    current = loaded if isinstance(loaded, dict) else None
                if not current or current.get("task_id") != task_id:
                    return False
            try:
                if self.path.exists():
                    self.path.unlink()
                self._data = None
                return True
            except OSError as exc:
                logger.error("清理任务恢复日志失败: %s", exc)
                return False
