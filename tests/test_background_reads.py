import sys
import tempfile
import unittest
from concurrent.futures import CancelledError
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.core.services import FilterService, OrganizeService
from invoice_tool.core.strategies import SegmentFilenameParser
from invoice_tool.core.workbook import WorkbookAnalyzerService
from invoice_tool.runtime import PANDAS_SUPPORT


class BackgroundReadServiceTests(unittest.TestCase):
    def test_organize_preview_returns_plain_renderable_result(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            folder = Path(temporary_directory)
            (folder / "dzfp_1001_示例公司.pdf").write_bytes(b"pdf")
            (folder / "格式不完整.pdf").write_bytes(b"pdf")

            result = OrganizeService.preview(
                folder,
                company_index=2,
                filename_parser=SegmentFilenameParser(),
            )

            self.assertEqual(result.total_count, 2)
            self.assertEqual(result.selectable_count, 1)
            self.assertEqual(result.invalid_count, 1)
            self.assertEqual(result.organized_count, 0)
            self.assertEqual(result.rows[0].relative_path, "dzfp_1001_示例公司.pdf")

    def test_organize_preview_honors_cancel_before_scanning(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            with self.assertRaises(CancelledError):
                OrganizeService.preview(
                    Path(temporary_directory),
                    company_index=2,
                    cancel_requested=lambda: True,
                )

    @unittest.skipUnless(PANDAS_SUPPORT, "pandas is required")
    def test_workbook_analysis_honors_cancel_before_opening_file(self):
        with self.assertRaises(CancelledError):
            WorkbookAnalyzerService.analyze(
                Path("missing.xlsx"),
                cancel_requested=lambda: True,
            )

    @unittest.skipUnless(PANDAS_SUPPORT, "pandas is required")
    def test_filter_preview_honors_cancel_before_opening_file(self):
        with self.assertRaises(CancelledError):
            FilterService.preview(
                Path("missing.xlsx"),
                Path("missing-pdfs"),
                1,
                cancel_requested=lambda: True,
            )


if __name__ == "__main__":
    unittest.main()
