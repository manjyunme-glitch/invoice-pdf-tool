import logging
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.ui.logging_handler import RecentErrorHandler


class FakeRoot:
    def after(self, _delay, callback):
        callback()


class RecentErrorHandlerTests(unittest.TestCase):
    def test_recent_error_handler_only_collects_error_records(self):
        entries = []
        handler = RecentErrorHandler(FakeRoot(), entries.append)
        handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s %(message)s", datefmt="%H:%M:%S"))

        info_record = logging.makeLogRecord({"levelno": logging.INFO, "levelname": "INFO", "msg": "just info"})
        error_record = logging.makeLogRecord({"levelno": logging.ERROR, "levelname": "ERROR", "msg": "something failed"})

        handler.emit(info_record)
        handler.emit(error_record)

        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["level"], "ERROR")
        self.assertEqual(entries[0]["summary"], "something failed")
        self.assertIn("something failed", entries[0]["detail"])


if __name__ == "__main__":
    unittest.main()
