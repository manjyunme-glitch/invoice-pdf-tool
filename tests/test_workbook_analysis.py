import sys
import tempfile
import unittest
from concurrent.futures import CancelledError
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.core import FilterService, InvoiceFilter, WorkbookAnalyzerService
from invoice_tool.runtime import PANDAS_SUPPORT, pd

try:
    import xlwt

    XLWT_SUPPORT = True
except ImportError:
    xlwt = None
    XLWT_SUPPORT = False


@unittest.skipUnless(PANDAS_SUPPORT, "pandas is required for workbook analysis tests")
class WorkbookAnalysisTests(unittest.TestCase):
    @unittest.skipUnless(XLWT_SUPPORT, "xlwt is required to generate a legacy .xls fixture")
    def test_legacy_xls_is_read_with_xlrd(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "legacy.xls"
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet("旧版表")
            sheet.write(0, 0, "发票号码")
            sheet.write(0, 1, "公司名称")
            sheet.write(1, 0, "00123456")
            sheet.write(1, 1, "测试公司")
            workbook.save(str(excel_path))

            sheets = InvoiceFilter.list_excel_sheets(excel_path)
            result = InvoiceFilter.read_invoice_records(
                excel_path,
                sheet_name="旧版表",
                company_column_name="公司名称",
            )

            self.assertEqual(sheets, ["旧版表"])
            self.assertEqual(result["invoice_numbers"], ["00123456"])
            self.assertEqual(result["records"][0]["company_name"], "测试公司")

    def test_empty_invoice_sheet_returns_empty_result_without_crashing(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "empty.xlsx"
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": pd.Series(dtype=object)}).to_excel(
                    writer,
                    sheet_name="空表",
                    index=False,
                )

            result = InvoiceFilter.read_invoice_records(excel_path, sheet_name="空表")

            self.assertEqual(result["source_row_count"], 0)
            self.assertEqual(result["invoice_numbers"], [])
            self.assertEqual(result["records"], [])

    def test_missing_invoice_column_has_actionable_error(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "missing-column.xlsx"
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"金额": [100], "备注": ["测试"]}).to_excel(
                    writer,
                    index=False,
                )

            with self.assertRaisesRegex(ValueError, "未找到发票号码列"):
                InvoiceFilter.read_invoice_records(excel_path)

    def test_invoice_values_preserve_text_zeroes_and_normalize_scientific_notation(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "number-formats.xlsx"
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame(
                    {"发票号码": ["00123", "1E+3", 456.0, None, "00123"]}
                ).to_excel(writer, index=False)

            result = InvoiceFilter.read_invoice_records(excel_path)

            self.assertEqual(result["invoice_numbers"], ["00123", "1000", "456"])

    def test_large_sheet_reads_with_cancellable_tuple_iteration(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "large.xlsx"
            row_count = 5_000
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame(
                    {
                        "发票号码": [f"{index:08d}" for index in range(row_count)],
                        "公司名称": ["测试公司"] * row_count,
                    }
                ).to_excel(writer, index=False)

            result = InvoiceFilter.read_invoice_records(excel_path)

            self.assertEqual(result["source_row_count"], row_count)
            self.assertEqual(len(result["invoice_numbers"]), row_count)
            self.assertEqual(result["invoice_numbers"][0], "00000000")
            self.assertEqual(result["invoice_numbers"][-1], "00004999")

    def test_large_sheet_can_cancel_during_row_iteration(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "cancel.xlsx"
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": range(1_000)}).to_excel(writer, index=False)
            calls = 0

            def cancel_requested():
                nonlocal calls
                calls += 1
                return calls >= 3

            with self.assertRaises(CancelledError):
                InvoiceFilter.read_invoice_records(
                    excel_path,
                    cancel_requested=cancel_requested,
                )

    def test_workbook_analyzer_profiles_multi_sheet_candidates(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "complex.xlsx"

            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame(
                    {
                        "发票号码": ["1001", "1002"],
                        "公司名称": ["甲公司", "乙公司"],
                        "金额": [100, 200],
                    }
                ).to_excel(writer, sheet_name="标准模板", index=False)
                pd.DataFrame(
                    {
                        "票号": ["2001", "2002"],
                        "客户名称": ["丙公司", "丁公司"],
                        "税额": [6, 12],
                    }
                ).to_excel(writer, sheet_name="自制模板", index=False)
                pd.DataFrame(
                    {
                        "摘要": ["服务费", "材料费"],
                        "金额": [50, 80],
                    }
                ).to_excel(writer, sheet_name="说明页", index=False)

            result = WorkbookAnalyzerService.analyze(excel_path)
            profiles = {profile.sheet_name: profile for profile in result.sheet_profiles}

            self.assertEqual(result.total_sheet_count, 3)
            self.assertEqual(result.usable_sheet_count, 2)
            self.assertEqual(result.recommended_sheet_name, "标准模板")

            standard = profiles["标准模板"]
            self.assertTrue(standard.recommended)
            self.assertEqual(standard.selected_invoice_column, "发票号码")
            self.assertEqual(standard.selected_company_column, "公司名称")
            self.assertTrue(standard.sample_rows)

            custom = profiles["自制模板"]
            self.assertTrue(custom.usable)
            self.assertIn("票号", [item.column_name for item in custom.invoice_candidates])
            self.assertIn("客户名称", [item.column_name for item in custom.company_candidates])

            notes = profiles["说明页"]
            self.assertFalse(notes.usable)
            self.assertEqual(notes.issue, "未识别到发票列")

    def test_filter_preview_accepts_manual_invoice_column_override(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "manual.xlsx"
            pdf_folder = root / "pdfs"
            pdf_folder.mkdir()

            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame(
                    {
                        "票据编号": ["9001", "9002"],
                        "往来单位": ["甲公司", "乙公司"],
                    }
                ).to_excel(writer, sheet_name="Sheet1", index=False)

            (pdf_folder / "dzfp_9001_甲公司_20240101.pdf").write_text("pdf", encoding="utf-8")

            preview = FilterService.preview(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                invoice_index=1,
                sheet_name="Sheet1",
                invoice_column_name="票据编号",
            )

            self.assertEqual(preview.excel_column_name, "票据编号")
            self.assertEqual(len(preview.matched), 1)
            self.assertEqual(preview.matched[0]["invoice"], "9001")
            self.assertEqual(preview.not_found, ["9002"])

    def test_workbook_analysis_handles_numeric_like_headers_without_crashing(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "numeric_headers.xlsx"
            pdf_folder = root / "pdfs"
            pdf_folder.mkdir()

            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame(
                    {
                        79423535.93: ["1001", "1002"],
                        "公司名称": ["甲公司", "乙公司"],
                    }
                ).to_excel(writer, sheet_name="8月进页", index=False)

            result = WorkbookAnalyzerService.analyze(excel_path)
            profile = result.sheet_profiles[0]
            self.assertEqual(profile.sheet_name, "8月进页")
            self.assertIn("79423535.93", profile.columns)

            (pdf_folder / "dzfp_1001_甲公司_20240101.pdf").write_text("pdf", encoding="utf-8")
            preview = FilterService.preview(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                invoice_index=1,
                sheet_name="8月进页",
                invoice_column_name="79423535.93",
            )

            self.assertEqual(preview.excel_column_name, "79423535.93")
            self.assertEqual(len(preview.matched), 1)
            self.assertEqual(preview.not_found, ["1002"])

    def test_filter_preview_supports_row_condition_and_company_exclusion(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "deduct.xlsx"
            pdf_folder = root / "pdfs"
            pdf_folder.mkdir()

            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame(
                    {
                        "发票号码": ["1001", "1002", "1003"],
                        "购方名称": ["重庆测试公司", "乱标记供应商", "重庆测试公司"],
                        "是否抵扣": ["是", "是", "否"],
                    }
                ).to_excel(writer, sheet_name="进项", index=False)

            (pdf_folder / "dzfp_1001_重庆测试公司_20240101.pdf").write_text("pdf", encoding="utf-8")
            (pdf_folder / "dzfp_1002_乱标记供应商_20240101.pdf").write_text("pdf", encoding="utf-8")
            (pdf_folder / "dzfp_1003_重庆测试公司_20240101.pdf").write_text("pdf", encoding="utf-8")

            preview = FilterService.preview(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                invoice_index=1,
                sheet_name="进项",
                invoice_column_name="发票号码",
                company_column_name="购方名称",
                filter_column_name="是否抵扣",
                filter_mode="等于任一",
                filter_values="是",
                company_exclude_keywords="乱标记",
            )

            self.assertEqual(preview.source_row_count, 3)
            self.assertEqual(preview.filtered_out_count, 2)
            self.assertEqual(preview.invoice_numbers, ["1001"])
            self.assertEqual(len(preview.matched), 1)

    def test_company_exclusion_requires_company_column(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "missing_company.xlsx"
            pdf_folder = root / "pdfs"
            pdf_folder.mkdir()
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": ["1001"]}).to_excel(writer, sheet_name="Sheet1", index=False)

            with self.assertRaisesRegex(ValueError, "尚未选择公司列"):
                FilterService.preview(
                    excel_path=excel_path,
                    pdf_folder=pdf_folder,
                    invoice_index=1,
                    company_exclude_keywords="测试",
                )

    def test_row_filter_requires_filter_column(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "missing_filter_column.xlsx"
            pdf_folder = root / "pdfs"
            pdf_folder.mkdir()
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": ["1001"]}).to_excel(writer, sheet_name="Sheet1", index=False)

            with self.assertRaisesRegex(ValueError, "尚未选择条件列"):
                FilterService.preview(
                    excel_path=excel_path,
                    pdf_folder=pdf_folder,
                    invoice_index=1,
                    filter_mode="等于任一",
                    filter_values="是",
                )

    def test_sheet_names_with_surrounding_spaces_remain_readable(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "spaced_sheet.xlsx"
            pdf_folder = root / "pdfs"
            pdf_folder.mkdir()
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": ["1001"]}).to_excel(writer, sheet_name=" Sheet1 ", index=False)

            sheets = InvoiceFilter.list_excel_sheets(excel_path)
            preview = FilterService.preview(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                invoice_index=1,
                sheet_name=" Sheet1 ",
            )

            self.assertEqual(sheets, [" Sheet1 "])
            self.assertEqual(preview.sheet_name, " Sheet1 ")


if __name__ == "__main__":
    unittest.main()
