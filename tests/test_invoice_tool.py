import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

MODULE_PATH = ROOT / "发票处理工具v5.py"
SPEC = importlib.util.spec_from_file_location("invoice_tool_entry", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


class Var:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class ButtonStub:
    def __init__(self):
        self.calls = []

    def config(self, **kwargs):
        self.calls.append(kwargs)


class InvoiceToolTests(unittest.TestCase):
    def test_parse_filename_uses_stem_when_last_segment_selected(self):
        company, valid = MODULE.InvoiceOrganizer.parse_filename(
            "dzfp_123456_测试公司_20240101.pdf",
            3,
        )
        self.assertTrue(valid)
        self.assertEqual(company, "20240101")

    def test_build_pdf_mapping_excludes_output_dir_and_counts_stats(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src = root / "src"
            out = src / "out"
            src.mkdir()
            out.mkdir()

            (src / "a_1001_A.pdf").write_text("ok", encoding="utf-8")
            (src / "dup_1001_B.pdf").write_text("dup", encoding="utf-8")
            (src / "bad.pdf").write_text("bad", encoding="utf-8")
            (out / "skip_2002_C.pdf").write_text("skip", encoding="utf-8")

            mapping, conflicts, stats = MODULE.InvoiceFilter.build_pdf_mapping(
                src,
                invoice_index=1,
                recursive=True,
                exclude_dirs=[out],
            )

            self.assertEqual(mapping, {"1001": "a_1001_A.pdf"})
            self.assertEqual(len(conflicts), 1)
            self.assertEqual(stats["scanned"], 3)
            self.assertEqual(stats["valid_named"], 2)
            self.assertEqual(stats["invalid_named"], 1)
            self.assertEqual(stats["duplicates"], 1)

    @unittest.skipUnless(MODULE.PANDAS_SUPPORT, "pandas is required for Excel tests")
    def test_read_invoice_numbers_supports_sheet_and_aliases(self):
        with tempfile.TemporaryDirectory() as td:
            excel_path = Path(td) / "sample.xlsx"

            with MODULE.pd.ExcelWriter(excel_path) as writer:
                MODULE.pd.DataFrame({"其他列": [1, 2]}).to_excel(writer, sheet_name="首页", index=False)
                MODULE.pd.DataFrame({"销项票号": ["1001", "1001", "1002.0", None]}).to_excel(
                    writer,
                    sheet_name="目标",
                    index=False,
                )

            invoice_numbers, col, sheet_name, columns = MODULE.InvoiceFilter.read_invoice_numbers(
                excel_path,
                sheet_name="目标",
                extra_aliases=["销项票号"],
            )

            self.assertEqual(invoice_numbers, ["1001", "1002"])
            self.assertEqual(col, "销项票号")
            self.assertEqual(sheet_name, "目标")
            self.assertEqual(columns, ["销项票号"])

    def test_validate_filter_paths_blocks_recursive_output_inside_source(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel = root / "sample.xlsx"
            pdf = root / "pdfs"
            out = pdf / "导出"
            excel.write_text("placeholder", encoding="utf-8")
            pdf.mkdir()

            app = SimpleNamespace(
                excel_path=Var(str(excel)),
                pdf_folder=Var(str(pdf)),
                output_folder=Var(str(out)),
                manual_output_folder=Var(str(out)),
                auto_output_by_sheet=Var(False),
                excel_sheet_name=Var("Sheet1"),
                filter_recursive=Var(True),
            )
            app._get_effective_output_folder_path = lambda: MODULE.InvoiceToolApp._get_effective_output_folder_path(app)

            with mock.patch.object(MODULE.messagebox, "showerror") as showerror:
                result = MODULE.InvoiceToolApp._validate_filter_paths(app)

            self.assertIsNone(result)
            self.assertTrue(showerror.called)
            self.assertIn("导出文件夹不能位于PDF源文件夹内部", showerror.call_args[0][1])

    def test_validate_filter_paths_returns_resolved_paths_for_valid_inputs(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel = root / "sample.xlsx"
            pdf = root / "pdfs"
            out = root / "exports"
            excel.write_text("placeholder", encoding="utf-8")
            pdf.mkdir()
            out.mkdir()

            app = SimpleNamespace(
                excel_path=Var(str(excel)),
                pdf_folder=Var(str(pdf)),
                output_folder=Var(str(out)),
                manual_output_folder=Var(str(out)),
                auto_output_by_sheet=Var(False),
                excel_sheet_name=Var("Sheet1"),
                filter_recursive=Var(False),
            )
            app._get_effective_output_folder_path = lambda: MODULE.InvoiceToolApp._get_effective_output_folder_path(app)

            with mock.patch.object(MODULE.messagebox, "showerror") as showerror:
                result = MODULE.InvoiceToolApp._validate_filter_paths(app)

            self.assertEqual(result, (excel, pdf, out.resolve()))
            showerror.assert_not_called()

    def test_undo_all_moves_keeps_failed_entries(self):
        m1 = {"source": "src1", "target": "target1", "filename": "a.pdf"}
        m2 = {"source": "src2", "target": "target2", "filename": "b.pdf"}
        app = SimpleNamespace(
            current_session_history=[m1, m2],
            undo_btn=ButtonStub(),
            undo_all_btn=ButtonStub(),
            all_history=[{"type": "整理", "moves": [m1, m2], "count": 2}],
            _save_history=lambda: None,
            _refresh_history_tree=lambda: None,
            _scan_files=lambda: None,
        )

        with mock.patch.object(MODULE.messagebox, "askyesno", return_value=True), \
             mock.patch.object(MODULE.messagebox, "showinfo"), \
             mock.patch.object(
                 MODULE.InvoiceOrganizer,
                 "rollback_single_move",
                 side_effect=[(True, ""), (False, "失败")],
             ):
            MODULE.InvoiceToolApp._undo_all_moves(app)

        self.assertEqual(app.current_session_history, [m1])
        self.assertEqual(app.all_history[0]["moves"], [m1])
        self.assertEqual(app.all_history[0]["count"], 1)
        self.assertEqual(app.undo_btn.calls[-1]["state"], "normal")
        self.assertEqual(app.undo_all_btn.calls[-1]["state"], "normal")

    def test_effective_output_folder_uses_excel_parent_and_sheet_when_auto_mode_enabled(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel = root / "2025年8月商品库存表.xlsx"
            excel.write_text("placeholder", encoding="utf-8")
            app = SimpleNamespace(
                excel_path=Var(str(excel)),
                excel_sheet_name=Var("8月进项"),
                auto_output_by_sheet=Var(True),
                manual_output_folder=Var(""),
                output_folder=Var(""),
            )

            path = MODULE.InvoiceToolApp._get_effective_output_folder_path(app)
            self.assertEqual(path, excel.parent / "8月进项")


    def test_sync_filter_context_resets_row_filters_when_sheet_changes(self):
        excel = Path("E:/vs-code/fapiao_v5/sample.xlsx")
        app = SimpleNamespace(
            excel_path=Var(str(excel)),
            excel_sheet_name=Var("8月进项"),
            row_filter_column_name=Var("是否抵扣"),
            row_filter_mode=Var("等于任一"),
            row_filter_values=Var("是"),
            company_exclude_keywords=Var("乱标记"),
            _active_filter_context=(str(excel), "7月进项"),
        )
        app._current_filter_context = lambda sheet_name=None: MODULE.InvoiceToolApp._current_filter_context(app, sheet_name)
        app._reset_sheet_row_filters = lambda: MODULE.InvoiceToolApp._reset_sheet_row_filters(app)

        changed = MODULE.InvoiceToolApp._sync_filter_context(app, "8月进项")

        self.assertTrue(changed)
        self.assertEqual(app.row_filter_column_name.get(), "")
        self.assertEqual(app.row_filter_mode.get(), "不过滤")
        self.assertEqual(app.row_filter_values.get(), "")
        self.assertEqual(app.company_exclude_keywords.get(), "")
        self.assertEqual(app._active_filter_context, (str(excel), "8月进项"))

    def test_sync_filter_context_keeps_row_filters_when_sheet_is_same(self):
        excel = Path("E:/vs-code/fapiao_v5/sample.xlsx")
        app = SimpleNamespace(
            excel_path=Var(str(excel)),
            excel_sheet_name=Var("8月进项"),
            row_filter_column_name=Var("是否抵扣"),
            row_filter_mode=Var("等于任一"),
            row_filter_values=Var("是"),
            company_exclude_keywords=Var("乱标记"),
            _active_filter_context=(str(excel), "8月进项"),
        )
        app._current_filter_context = lambda sheet_name=None: MODULE.InvoiceToolApp._current_filter_context(app, sheet_name)
        app._reset_sheet_row_filters = lambda: MODULE.InvoiceToolApp._reset_sheet_row_filters(app)

        changed = MODULE.InvoiceToolApp._sync_filter_context(app, "8月进项")

        self.assertFalse(changed)
        self.assertEqual(app.row_filter_column_name.get(), "是否抵扣")
        self.assertEqual(app.row_filter_mode.get(), "等于任一")
        self.assertEqual(app.row_filter_values.get(), "是")
        self.assertEqual(app.company_exclude_keywords.get(), "乱标记")


if __name__ == "__main__":
    unittest.main()
