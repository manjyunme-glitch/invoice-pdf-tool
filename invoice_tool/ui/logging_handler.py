from __future__ import annotations

from datetime import datetime
import logging
import threading
import tkinter as tk
from typing import Callable, Dict


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

    def __init__(self, text_widget: tk.Text, root: tk.Tk) -> None:
        super().__init__()
        self.text_widget = text_widget
        self.root = root

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        tag = self.LEVEL_TAG.get(record.levelno, "info")
        for prefix, mapped_tag in self.PREFIX_TAG.items():
            if record.getMessage().startswith(prefix):
                tag = mapped_tag
                break

        def write() -> None:
            self.text_widget.insert("end", msg + "\n", tag)
            self.text_widget.see("end")

        if threading.current_thread() is threading.main_thread():
            write()
        else:
            self.root.after(0, write)


class RecentErrorHandler(logging.Handler):
    """提取 error/critical 日志并回传给 UI。"""

    def __init__(self, root: tk.Tk, callback: Callable[[Dict[str, str]], None]) -> None:
        super().__init__(level=logging.ERROR)
        self.root = root
        self.callback = callback

    def emit(self, record: logging.LogRecord) -> None:
        if record.levelno < logging.ERROR:
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
            self.callback(entry)

        if threading.current_thread() is threading.main_thread():
            push()
        else:
            self.root.after(0, push)
