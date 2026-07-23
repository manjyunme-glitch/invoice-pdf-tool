import logging
import sys
import threading
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.ui.logging_handler import RecentErrorHandler, TkTextHandler


class FakeText:
    def __init__(self):
        self.writes = []

    def insert(self, position, text, tag):
        self.writes.append((position, text, tag))

    def see(self, _position):
        return None


class RecentErrorHandlerTests(unittest.TestCase):
    def test_recent_error_handler_only_collects_error_records(self):
        entries = []
        handler = RecentErrorHandler(entries.append, lambda callback: callback())
        handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s %(message)s", datefmt="%H:%M:%S"))

        info_record = logging.makeLogRecord({"levelno": logging.INFO, "levelname": "INFO", "msg": "just info"})
        error_record = logging.makeLogRecord({"levelno": logging.ERROR, "levelname": "ERROR", "msg": "something failed"})

        handler.emit(info_record)
        handler.emit(error_record)

        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["level"], "ERROR")
        self.assertEqual(entries[0]["summary"], "something failed")
        self.assertIn("something failed", entries[0]["detail"])

    def test_background_log_handlers_only_use_supplied_ui_dispatcher(self):
        callbacks = []
        entries = []
        text_widget = FakeText()
        text_handler = TkTextHandler(text_widget, callbacks.append)
        error_handler = RecentErrorHandler(entries.append, callbacks.append)
        info_record = logging.makeLogRecord({"levelno": logging.INFO, "levelname": "INFO", "msg": "worker info"})
        error_record = logging.makeLogRecord({"levelno": logging.ERROR, "levelname": "ERROR", "msg": "worker error"})

        worker = threading.Thread(
            target=lambda: (text_handler.emit(info_record), error_handler.emit(error_record))
        )
        worker.start()
        worker.join()

        self.assertEqual(len(callbacks), 2)
        self.assertEqual(text_widget.writes, [])
        self.assertEqual(entries, [])
        for callback in callbacks:
            callback()
        self.assertEqual(text_widget.writes[0][1], "worker info\n")
        self.assertEqual(entries[0]["summary"], "worker error")

    def test_closed_log_handler_ignores_queued_widget_update(self):
        callbacks = []
        text_widget = FakeText()
        handler = TkTextHandler(text_widget, callbacks.append)
        record = logging.makeLogRecord({"levelno": logging.INFO, "levelname": "INFO", "msg": "stale"})

        handler.emit(record)
        handler.close()
        callbacks[0]()

        self.assertEqual(text_widget.writes, [])


if __name__ == "__main__":
    unittest.main()
