import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.core import (
    DEFAULT_RULE_PRESET_ID,
    DEFAULT_EXCLUDE_KEYWORDS,
    FilterService,
    InvoiceFilter,
    SegmentFilenameParser,
    SmartInvoiceColumnResolver,
    get_rule_preset,
    list_rule_presets,
)
from invoice_tool.runtime import PANDAS_SUPPORT, pd


class FakeReportExporter:
    def __init__(self):
        self.called = False

    def export_filter_report(self, output_dir, matched, not_found, excel_col_name):
        self.called = True
        report_path = output_dir / "fake_report.txt"
        report_path.write_text(
            f"matched={len(matched)};not_found={len(not_found)};col={excel_col_name}",
            encoding="utf-8",
        )
        return report_path


class StrategyTests(unittest.TestCase):
    def test_segment_filename_parser_supports_custom_separator(self):
        parser = SegmentFilenameParser(separator="-")
        self.assertEqual(
            parser.split_parts("dzfp-1001-测试公司-20240101.pdf"),
            ["dzfp", "1001", "测试公司", "20240101"],
        )
        self.assertEqual(parser.parse_segment("dzfp-1001-测试公司-20240101.pdf", 2), "测试公司")

    def test_column_resolver_can_use_custom_exact_names(self):
        resolver = SmartInvoiceColumnResolver(
            exact_column_names=("票据号",),
            exclude_keywords=DEFAULT_EXCLUDE_KEYWORDS,
        )
        result = InvoiceFilter.find_invoice_column(
            ["备注发票号", "票据号", "开票日期"],
            column_resolver=resolver,
        )
        self.assertEqual(result, "票据号")

    def test_rule_preset_registry_exposes_default_and_supplier_preset(self):
        preset_ids = [preset.preset_id for preset in list_rule_presets()]
        self.assertIn(DEFAULT_RULE_PRESET_ID, preset_ids)
        supplier = get_rule_preset("supplier_archive")
        self.assertEqual(supplier.company_name_index, 1)
        self.assertEqual(supplier.invoice_number_index, 2)

    @unittest.skipUnless(PANDAS_SUPPORT, "pandas is required for strategy service tests")
    def test_filter_service_accepts_injected_strategies_and_report_exporter(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "sample.xlsx"
            pdf_folder = root / "pdfs"
            out_folder = root / "out"
            pdf_folder.mkdir()

            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"票据号": ["1001"]}).to_excel(writer, sheet_name="Sheet1", index=False)

            (pdf_folder / "dzfp-1001-测试公司-20240101.pdf").write_text("pdf", encoding="utf-8")

            parser = SegmentFilenameParser(separator="-")
            resolver = SmartInvoiceColumnResolver(exact_column_names=("票据号",))
            report_exporter = FakeReportExporter()

            result = FilterService.run(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                output_dir=out_folder,
                invoice_index=1,
                recursive=False,
                column_resolver=resolver,
                filename_parser=parser,
                report_exporter=report_exporter,
            )

            self.assertEqual(result.found_count, 1)
            self.assertEqual(result.copy_fail_count, 0)
            self.assertTrue(report_exporter.called)
            self.assertIsNotNone(result.report_path)
            self.assertTrue(result.report_path.exists())
            self.assertTrue((out_folder / "dzfp-1001-测试公司-20240101.pdf").exists())
            self.assertEqual([row.status for row in result.result_rows], ["已导出"])
            self.assertEqual(result.result_rows[0].invoice_number, "1001")


if __name__ == "__main__":
    unittest.main()
