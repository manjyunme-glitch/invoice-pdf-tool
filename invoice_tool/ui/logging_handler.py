from __future__ import annotations

from datetime import datetime
import logging
import tkinter as tk
from typing import Callable, Dict


UiDispatcher = Callable[[Callable[[], None]], None]


class TkTextHandler(logging.Handler):
    """将 logging 输出渲染到 tk.Text。"""

    LEVEL_TAG = {
        logging.DEBUG: "info",
        logging.INFO: "info",
        logging.WARNING: "warning",
        logging.ERROR: "error",
        logging.CRITICAL: "error",
    }

    PREFIX_TAG = {
        "✅": "success",
        "📊": "header",
        "🚀": "header",
        "===": "header",
    }

    def __init__(self, text_widget: tk.Text, dispatch_ui: UiDispatcher) -> None:
        super().__init__()
        self.text_widget = text_widget
        self.dispatch_ui = dispatch_ui
        self._active = True

    def emit(self, record: logging.LogRecord) -> None:
        if not self._active:
            return
        msg = self.format(record)
        tag = self.LEVEL_TAG.get(record.levelno, "info")
        for prefix, mapped_tag in self.PREFIX_TAG.items():
            if record.getMessage().startswith(prefix):
                tag = mapped_tag
                break

        def write() -> None:
            if not self._active:
                return
            self.text_widget.insert("end", msg + "\n", tag)
            self.text_widget.see("end")

        self.dispatch_ui(write)

    def close(self) -> None:
        self._active = False
        super().close()


class RecentErrorHandler(logging.Handler):
    """提取 error/critical 日志并回传给 UI。"""

    def __init__(self, callback: Callable[[Dict[str, str]], None], dispatch_ui: UiDispatcher) -> None:
        super().__init__(level=logging.ERROR)
        self.callback = callback
        self.dispatch_ui = dispatch_ui
        self._active = True

    def emit(self, record: logging.LogRecord) -> None:
        if not self._active or record.levelno < logging.ERROR:
            return

        formatted = self.format(record)
        summary = record.getMessage().splitlines()[0].strip()
        entry = {
            "time": datetime.fromtimestamp(record.created).strftime("%H:%M:%S"),
            "level": record.levelname,
            "summary": summary,
            "detail": formatted,
        }

        def push() -> None:
            if self._active:
                self.callback(entry)

        self.dispatch_ui(push)

    def close(self) -> None:
        self._active = False
        super().close()
