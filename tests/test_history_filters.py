import sys
import unittest
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.ui.app import filter_history_records


class HistoryFilterTests(unittest.TestCase):
    def test_filter_history_records_supports_type_date_and_keyword(self):
        records = [
            {
                "time": "2026-04-08 09:00:00",
                "folder": r"D:\pdf\alpha",
                "count": 2,
                "type": "整理",
                "moves": [{"filename": "a.pdf"}],
            },
            {
                "time": "2026-04-05 10:00:00",
                "folder": r"D:\pdf\beta",
                "count": 1,
                "type": "筛选",
                "moves": [{"filename": "b.pdf"}],
                "report_files": [r"D:\out\report.xlsx"],
            },
            {
                "time": "2026-03-01 10:00:00",
                "folder": r"D:\pdf\gamma",
                "count": 1,
                "type": "筛选",
                "moves": [{"filename": "c.pdf"}],
            },
        ]

        now = datetime(2026, 4, 8, 12, 0, 0)

        self.assertEqual(
            filter_history_records(records, type_filter="筛选", date_filter="全部", keyword="", now=now),
            [1, 2],
        )
        self.assertEqual(
            filter_history_records(records, type_filter="全部", date_filter="最近7天", keyword="", now=now),
            [0, 1],
        )
        self.assertEqual(
            filter_history_records(records, type_filter="全部", date_filter="全部", keyword="report", now=now),
            [1],
        )
        self.assertEqual(
            filter_history_records(records, type_filter="整理", date_filter="最近30天", keyword="alpha", now=now),
            [0],
        )


if __name__ == "__main__":
    unittest.main()
