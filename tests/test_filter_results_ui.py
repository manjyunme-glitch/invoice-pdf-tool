import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.core import FilterResultRow, FilterService
from invoice_tool.runtime import PANDAS_SUPPORT, pd
from invoice_tool.ui.app import filter_filter_result_rows, sort_filter_result_rows


class FilterResultUiTests(unittest.TestCase):
    def test_filter_helpers_support_status_and_keyword(self):
        rows = [
            FilterResultRow(status="已导出", invoice_number="1001", pdf_name="a.pdf", detail="ok"),
            FilterResultRow(status="未匹配", invoice_number="1002", pdf_name="", detail="未找到对应PDF"),
            FilterResultRow(status="重复冲突", invoice_number="1001", pdf_name="", detail="a.pdf, b.pdf"),
        ]

        filtered = filter_filter_result_rows(rows, status_filter="未匹配", keyword="")
        self.assertEqual([row.invoice_number for row in filtered], ["1002"])

        searched = filter_filter_result_rows(rows, status_filter="全部", keyword="a.pdf")
        self.assertEqual([row.status for row in searched], ["已导出", "重复冲突"])

        sorted_rows = sort_filter_result_rows(rows, sort_key="status", descending=False)
        self.assertEqual([row.status for row in sorted_rows], ["已导出", "未匹配", "重复冲突"])

    @unittest.skipUnless(PANDAS_SUPPORT, "pandas is required for preview row tests")
    def test_filter_preview_exposes_rows_for_table_view(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "sample.xlsx"
            pdf_folder = root / "pdfs"
            pdf_folder.mkdir()

            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": ["1001", "1002"]}).to_excel(writer, sheet_name="Sheet1", index=False)

            (pdf_folder / "dzfp_1001_测试公司_20240101.pdf").write_text("pdf", encoding="utf-8")
            (pdf_folder / "dup_1001_另一家公司_20240102.pdf").write_text("pdf", encoding="utf-8")

            preview = FilterService.preview(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                invoice_index=1,
            )

            statuses = [row.status for row in preview.result_rows]
            self.assertIn("可匹配", statuses)
            self.assertIn("未匹配", statuses)
            self.assertIn("重复冲突", statuses)


if __name__ == "__main__":
    unittest.main()
