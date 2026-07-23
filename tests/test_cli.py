import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool import app as app_module
from invoice_tool import cli as cli_module
from invoice_tool.cli import main as cli_main
from invoice_tool.runtime import PANDAS_SUPPORT, pd


class CliTests(unittest.TestCase):
    def test_gui_falls_back_to_native_tk_when_drag_drop_root_fails(self):
        native_root = mock.Mock()
        dnd = mock.Mock()
        dnd.Tk.side_effect = RuntimeError("missing tkdnd binary")

        with (
            mock.patch("invoice_tool.app.DND_SUPPORT", True),
            mock.patch("invoice_tool.app.TkinterDnD", dnd),
            mock.patch("invoice_tool.app.tk.Tk", return_value=native_root) as native_tk,
            mock.patch("invoice_tool.app.InvoiceToolApp") as app_class,
            mock.patch("invoice_tool.app.MODERN_UI", False),
        ):
            app_module.run_gui()

        dnd.Tk.assert_called_once_with()
        native_tk.assert_called_once_with()
        app_class.assert_called_once_with(native_root)
        native_root.mainloop.assert_called_once_with()

    def test_gui_continues_when_modern_theme_initialization_fails(self):
        root = mock.Mock()
        style_module = mock.Mock()
        style_module.Style.side_effect = RuntimeError("theme unavailable")

        with (
            mock.patch("invoice_tool.app.DND_SUPPORT", False),
            mock.patch("invoice_tool.app.tk.Tk", return_value=root),
            mock.patch("invoice_tool.app.InvoiceToolApp") as app_class,
            mock.patch("invoice_tool.app.MODERN_UI", True),
            mock.patch("invoice_tool.app.ttkb", style_module),
        ):
            app_module.run_gui()

        style_module.Style.assert_called_once_with(theme="cosmo")
        app_class.assert_called_once_with(root)
        root.mainloop.assert_called_once_with()

    def test_should_hold_console_only_for_frozen_explorer_launch(self):
        original_frozen = getattr(sys, "frozen", None)
        try:
            sys.frozen = True
            self.assertTrue(cli_module._should_hold_console(None, 0, parent_name="explorer.exe"))
            self.assertTrue(cli_module._should_hold_console("filter", 1, parent_name="explorer.exe"))
            self.assertFalse(cli_module._should_hold_console("filter", 0, parent_name="explorer.exe"))
            self.assertFalse(cli_module._should_hold_console(None, 0, parent_name="powershell.exe"))
        finally:
            if original_frozen is None:
                delattr(sys, "frozen")
            else:
                sys.frozen = original_frozen

    def test_presets_command_outputs_json(self):
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            exit_code = cli_main(["presets", "--json"])

        payload = json.loads(stdout.getvalue())
        self.assertEqual(exit_code, 0)
        self.assertTrue(payload["presets"])
        self.assertIn("standard_digital", [item["id"] for item in payload["presets"]])

    def test_app_main_routes_cli_arguments_to_cli_entry(self):
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            exit_code = app_module.main(["presets", "--json"])

        payload = json.loads(stdout.getvalue())
        self.assertEqual(exit_code, 0)
        self.assertTrue(payload["presets"])

    @unittest.skipUnless(PANDAS_SUPPORT, "pandas is required for CLI filter tests")
    def test_filter_dry_run_outputs_preview_summary(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            excel_path = root / "sample.xlsx"
            pdf_folder = root / "pdfs"
            output_folder = root / "out"
            pdf_folder.mkdir()
            output_folder.mkdir()

            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": ["1001", "1002"]}).to_excel(writer, sheet_name="Sheet1", index=False)

            (pdf_folder / "dzfp_1001_测试公司_20240101.pdf").write_text("pdf", encoding="utf-8")

            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                exit_code = cli_main(
                    [
                        "filter",
                        "--excel",
                        str(excel_path),
                        "--pdf-folder",
                        str(pdf_folder),
                        "--output-folder",
                        str(output_folder),
                        "--dry-run",
                        "--json",
                    ]
                )

            payload = json.loads(stdout.getvalue())
            self.assertEqual(exit_code, 0)
            self.assertEqual(payload["mode"], "filter-dry-run")
            self.assertEqual(payload["matched_count"], 1)
            self.assertEqual(payload["not_found_count"], 1)

    def test_organize_dry_run_outputs_summary(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            folder = root / "pdfs"
            folder.mkdir()
            (folder / "dzfp_1001_测试公司_20240101.pdf").write_text("pdf", encoding="utf-8")
            (folder / "bad.pdf").write_text("pdf", encoding="utf-8")

            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                exit_code = cli_main(
                    [
                        "organize",
                        "--folder",
                        str(folder),
                        "--dry-run",
                        "--json",
                    ]
                )

            payload = json.loads(stdout.getvalue())
            self.assertEqual(exit_code, 0)
            self.assertEqual(payload["mode"], "organize-dry-run")
            self.assertEqual(payload["scanned"], 2)
            self.assertEqual(payload["valid"], 1)
            self.assertEqual(payload["selected"], 1)


if __name__ == "__main__":
    unittest.main()
