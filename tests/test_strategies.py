import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock


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
from invoice_tool.core.strategies import OpenpyxlFilterReportExporter
from invoice_tool.runtime import OPENPYXL_SUPPORT, PANDAS_SUPPORT, openpyxl, pd


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

    def test_segment_filename_parser_handles_bounds_and_negative_indices(self):
        parser = SegmentFilenameParser(separator="_")
        self.assertIsNone(parser.parse_segment("a_b_c.pdf", 5))
        self.assertIsNone(parser.parse_segment("a_b_c.pdf", -5))
        self.assertEqual(parser.parse_segment("a_b_c.pdf", -1), "c")
        self.assertEqual(parser.parse_segment("a_b_c.pdf", -2), "b")

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

    @unittest.skipUnless(OPENPYXL_SUPPORT, "openpyxl is required for report tests")
    def test_default_report_exporter_never_overwrites_same_second_report(self):
        fixed_now = datetime(2026, 7, 22, 10, 30, 45)
        exporter = OpenpyxlFilterReportExporter()
        with tempfile.TemporaryDirectory() as td, mock.patch(
            "invoice_tool.core.strategies.datetime"
        ) as datetime_mock:
            datetime_mock.now.return_value = fixed_now
            output_dir = Path(td)
            first = exporter.export_filter_report(output_dir, [], [], "发票号码")
            second = exporter.export_filter_report(output_dir, [], [], "发票号码")

            self.assertIsNotNone(first)
            self.assertIsNotNone(second)
            self.assertNotEqual(first, second)
            self.assertEqual(first.name, "筛选结果报告_20260722_103045.xlsx")
            self.assertEqual(second.name, "筛选结果报告_20260722_103045_2.xlsx")
            self.assertTrue(first.exists())
            self.assertTrue(second.exists())
            self.assertFalse(list(output_dir.glob(".*.tmp.xlsx")))

    @unittest.skipUnless(OPENPYXL_SUPPORT, "openpyxl is required for report tests")
    def test_report_contains_all_status_details_and_escapes_formula_like_text(self):
        exporter = OpenpyxlFilterReportExporter()
        with tempfile.TemporaryDirectory() as td:
            output_dir = Path(td)
            report = exporter.export_filter_report(
                output_dir,
                [{"invoice_number": "=1+1", "filename": "+danger.pdf", "time": "now"}],
                ["@missing"],
                "=发票列",
                result_rows=[
                    {
                        "status": "复制失败",
                        "invoice_number": "=1+1",
                        "pdf_name": "+danger.pdf",
                        "detail": "@原因",
                        "path": "-危险路径",
                    },
                    {
                        "status": "同名冲突",
                        "invoice_number": "1002",
                        "pdf_name": "normal.pdf",
                        "detail": "已保留原文件",
                        "path": "",
                    },
                ],
            )

            workbook = openpyxl.load_workbook(report, data_only=False)
            try:
                self.assertIn("处理明细", workbook.sheetnames)
                self.assertEqual(workbook["已成功导出"]["B2"].value, "'=1+1")
                self.assertEqual(workbook["已成功导出"]["C2"].value, "'+danger.pdf")
                self.assertEqual(workbook["缺失清单"]["B2"].value, "'@missing")
                details = workbook["处理明细"]
                self.assertEqual(details["C2"].value, "'=1+1")
                self.assertEqual(details["D2"].value, "'+danger.pdf")
                self.assertEqual(details["E2"].value, "'@原因")
                self.assertEqual(details["F2"].value, "'-危险路径")
                summary_values = {
                    workbook["汇总"].cell(row=row, column=1).value:
                    workbook["汇总"].cell(row=row, column=2).value
                    for row in range(1, workbook["汇总"].max_row + 1)
                }
                self.assertEqual(summary_values["状态：复制失败"], 1)
                self.assertEqual(summary_values["状态：同名冲突"], 1)
            finally:
                workbook.close()

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
