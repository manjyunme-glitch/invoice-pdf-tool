import sys
import tempfile
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
import unittest
import zipfile
from concurrent.futures import CancelledError
from pathlib import Path
from tkinter import ttk
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.ui import InvoiceToolApp
from invoice_tool.application.configuration import normalize_config
from invoice_tool.core.models import (
    FilterPreviewResult,
    FilterResultRow,
    OrganizePreviewResult,
    OrganizePreviewRow,
    OrganizeResultRow,
    OrganizeTaskResult,
    PdfScanStats,
    WorkbookAnalysisResult,
    WorkbookSheetProfile,
)
from invoice_tool.infra.storage import load_json, save_json
from invoice_tool.ui.app import FILTER_WORKFLOW_STEPS, ORGANIZE_WORKFLOW_STEPS, UI_THEME_PRESETS
from invoice_tool.ui.workspace_app import WORKSPACE_PAGES, WorkspacePageStack
from invoice_tool.runtime import MODERN_UI, ttkb


class UiLayoutTests(unittest.TestCase):
    def setUp(self):
        if MODERN_UI and ttkb is not None:
            ttkb.Style.instance = None
        self._data_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self._data_dir.cleanup()
        if MODERN_UI and ttkb is not None:
            ttkb.Style.instance = None

    def _create_app(self, root: tk.Tk) -> InvoiceToolApp:
        data_dir = Path(self._data_dir.name)
        return InvoiceToolApp(
            root,
            config_file=data_dir / "config.json",
            history_file=data_dir / "history.json",
            active_task_file=data_dir / "active_task.json",
        )

    def _wait_for_idle(self, root: tk.Tk, app: InvoiceToolApp, timeout: float = 8.0) -> None:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            root.update()
            with app._lock:
                if not app.is_running:
                    root.update()
                    return
            time.sleep(0.01)
        self.fail("后台任务未在预期时间内结束")

    @staticmethod
    def _relative_y(widget: tk.Widget, root: tk.Widget) -> int:
        y = 0
        current = widget
        while current is not root:
            y += current.winfo_y()
            current = current.master
        return y

    @staticmethod
    def _contrast_ratio(first: str, second: str) -> float:
        def luminance(color: str) -> float:
            channels = [int(color[index:index + 2], 16) / 255 for index in (1, 3, 5)]
            linear = [
                value / 12.92
                if value <= 0.04045
                else ((value + 0.055) / 1.055) ** 2.4
                for value in channels
            ]
            return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]

        lighter, darker = sorted(
            (luminance(first), luminance(second)),
            reverse=True,
        )
        return (lighter + 0.05) / (darker + 0.05)

    def test_filter_result_tree_uses_grid_scrollbars(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            tree_frame = app.filter_result_tree.master
            managers = [child.winfo_manager() for child in tree_frame.winfo_children()]
            self.assertTrue(managers)
            self.assertTrue(all(manager == "grid" for manager in managers))
        finally:
            app._on_closing()

    def test_filter_page_uses_continuous_workflow_with_fixed_action_bar(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            root.update_idletasks()
            expected_stages = tuple(stage[0] for stage in FILTER_WORKFLOW_STEPS)
            self.assertEqual(tuple(app.filter_workflow_cards), expected_stages)
            self.assertEqual(tuple(app.filter_workflow_sections), ("input", "rules", "results"))
            self.assertEqual(app.filter_action_bar.pack_info()["side"], "bottom")
            self.assertIsNot(app.filter_action_bar, app.filter_scroll_panel)
            self.assertIsNot(app.filter_preview_btn.master, app.filter_scroll_panel)
            self.assertIs(app.filter_preview_btn.master, app.filter_run_btn.master)
            self.assertIs(app.filter_run_btn.master, app.filter_retry_btn.master)
            self.assertIs(app.filter_retry_btn.master, app.pause_filter_btn.master)
            self.assertEqual(app.pause_filter_btn.cget("state"), "disabled")
            self.assertIs(app.filter_run_btn.master, app.cancel_flt_btn.master)
            self.assertIs(app.filter_progress.master, app.filter_run_btn.master)
        finally:
            app._on_closing()

    def test_filter_workflow_stage_updates_progress_cards_and_status(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app._set_filter_workflow_stage("execute", "后台任务正在执行")
            self.assertEqual(app.filter_workflow_stage.get(), "execute")
            self.assertEqual(app.filter_workflow_status_text.get(), "后台任务正在执行")
            self.assertTrue(app.filter_workflow_cards["input"].cget("text").startswith("✓"))
            self.assertEqual(app.filter_workflow_cards["execute"].cget("bg"), app.palette["primary"])
            self.assertEqual(app._scroll_filter_workflow_to("preview"), "break")
        finally:
            app._on_closing()

    def test_filter_workflow_navigation_aligns_each_scroll_destination(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            root.update_idletasks()
            canvas = app.filter_scroll_canvas
            panel = app.filter_scroll_panel
            scroll_region = canvas.bbox("all")
            self.assertIsNotNone(scroll_region)
            assert scroll_region is not None
            content_top = scroll_region[1]
            content_height = scroll_region[3] - content_top
            max_offset = max(content_height - canvas.winfo_height(), 0)

            expected_sections = {
                "input": "input",
                "rules": "rules",
                "preview": "results",
                "execute": "results",
                "results": "results",
            }
            for stage_key, section_key in expected_sections.items():
                canvas.yview_moveto(0)
                app._scroll_filter_workflow_to(stage_key)
                root.update_idletasks()
                section = app.filter_workflow_sections[section_key]
                expected_offset = min(
                    max(section.winfo_y() - content_top - 8, 0),
                    max_offset,
                )
                actual_offset = canvas.yview()[0] * content_height
                self.assertAlmostEqual(
                    actual_offset,
                    expected_offset,
                    delta=2,
                    msg=f"{stage_key} 没有滚动到 {section_key} 的正确位置",
                )
        finally:
            app._on_closing()

    def test_disabled_action_buttons_remain_readable_and_primary_type_is_preserved(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            for theme_id in ("day", "night"):
                if app.ui_theme.get() != theme_id:
                    app._set_ui_theme(theme_id)
                for role in (
                    "primary",
                    "success",
                    "warning",
                    "danger",
                    "secondary",
                    "neutral",
                ):
                    normal_bg, hover_bg, foreground = app._button_colors(role)
                    self.assertGreaterEqual(
                        self._contrast_ratio(normal_bg, foreground),
                        4.5,
                    )
                    self.assertGreaterEqual(
                        self._contrast_ratio(hover_bg, foreground),
                        4.5,
                    )
                for button in (
                    app.filter_retry_btn,
                    app.pause_filter_btn,
                    app.cancel_flt_btn,
                ):
                    disabled_bg = str(button.cget("bg"))
                    disabled_fg = str(button.cget("disabledforeground"))
                    self.assertIn(
                        disabled_fg,
                        {
                            app.palette["button_disabled_fg"],
                            app.palette["button_disabled_accent_fg"],
                        },
                    )
                    self.assertGreaterEqual(
                        self._contrast_ratio(disabled_bg, disabled_fg),
                        4.5,
                    )

            primary_font = tkfont.Font(root=root, font=app.filter_run_btn.cget("font"))
            self.assertGreaterEqual(abs(int(primary_font.cget("size"))), 12)
            self.assertEqual(primary_font.cget("weight"), "bold")
        finally:
            app._on_closing()

    def test_workspace_palette_is_not_overwritten_by_modern_ttk_theme(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            for theme_id in ("day", "night"):
                if app.ui_theme.get() != theme_id:
                    app._set_ui_theme(theme_id)
                self.assertEqual(
                    str(app.workspace_sidebar.cget("bg")),
                    app.palette["hero_card_bg"],
                )
                self.assertEqual(
                    str(app.workspace_main.cget("bg")),
                    app.palette["root_bg"],
                )
                brand = app.workspace_sidebar.winfo_children()[0]
                self.assertEqual(str(brand.cget("bg")), app.palette["hero_card_bg"])
                brand_labels = [
                    child
                    for child in brand.winfo_children()
                    if isinstance(child, tk.Label)
                ]
                self.assertTrue(brand_labels)
                self.assertTrue(
                    all(
                        str(label.cget("bg")) == app.palette["hero_card_bg"]
                        for label in brand_labels
                    )
                )
        finally:
            app._on_closing()

    def test_filter_workflow_stage_survives_theme_rebuild(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app._set_filter_workflow_stage("results", "预览完成")
            app._set_ui_theme("night")
            self.assertEqual(app.filter_workflow_stage.get(), "results")
            self.assertEqual(app.filter_workflow_status_text.get(), "预览完成")
            self.assertEqual(app.filter_workflow_cards["results"].cget("bg"), app.palette["primary"])
        finally:
            app._on_closing()

    def test_organize_page_uses_continuous_workflow_with_fixed_action_bar(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            root.update_idletasks()
            expected_stages = tuple(stage[0] for stage in ORGANIZE_WORKFLOW_STEPS)
            self.assertEqual(tuple(app.organize_workflow_cards), expected_stages)
            self.assertEqual(tuple(app.organize_workflow_sections), ("input", "preview", "results"))
            self.assertEqual(app.organize_action_bar.pack_info()["side"], "bottom")
            self.assertIs(app.start_btn.master, app.undo_btn.master)
            self.assertIs(app.start_btn.master, app.retry_org_btn.master)
            self.assertIs(app.retry_org_btn.master, app.pause_org_btn.master)
            self.assertEqual(app.retry_org_btn.cget("state"), "disabled")
            self.assertEqual(app.pause_org_btn.cget("state"), "disabled")
            self.assertIs(app.undo_btn.master, app.undo_all_btn.master)
            self.assertIs(app.undo_all_btn.master, app.cancel_org_btn.master)
            self.assertIs(app.organize_progress.master, app.start_btn.master)
            self.assertIsNot(app.start_btn.master, app.organize_action_bar)
        finally:
            app._on_closing()

    def test_organize_workflow_stage_updates_cards_result_and_focus(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app._set_organize_workflow_stage("execute", "正在整理")
            app._update_organize_result("整理进行中", "计划处理 3 个文件")
            self.assertEqual(app.organize_workflow_stage.get(), "execute")
            self.assertEqual(app.organize_workflow_status_text.get(), "正在整理")
            self.assertEqual(app.organize_result_title.get(), "整理进行中")
            self.assertEqual(app.organize_result_detail.get(), "计划处理 3 个文件")
            self.assertTrue(app.organize_workflow_cards["input"].cget("text").startswith("✓"))
            self.assertEqual(app.organize_workflow_cards["execute"].cget("bg"), app.palette["primary"])
            self.assertEqual(app._focus_organize_workflow_step("confirm"), "break")
        finally:
            app._on_closing()

    def test_organize_workflow_state_survives_theme_rebuild(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app._set_organize_workflow_stage("results", "整理完成")
            app._update_organize_result("整理完成", "成功 3 · 失败 0")
            app._set_ui_theme("night")
            self.assertEqual(app.organize_workflow_stage.get(), "results")
            self.assertEqual(app.organize_workflow_status_text.get(), "整理完成")
            self.assertEqual(app.organize_result_title.get(), "整理完成")
            self.assertEqual(app.organize_result_detail.get(), "成功 3 · 失败 0")
            self.assertEqual(app.organize_workflow_cards["results"].cget("bg"), app.palette["primary"])
        finally:
            app._on_closing()

    def test_organize_scan_advances_to_confirmation_with_summary(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        invoice_dir = Path(self._data_dir.name) / "待整理发票"
        invoice_dir.mkdir()
        (invoice_dir / "销售_1001_示例公司.pdf").write_bytes(b"pdf")
        (invoice_dir / "格式不完整.pdf").write_bytes(b"pdf")
        try:
            app.organize_folder_path.set(str(invoice_dir))
            app._scan_files()
            self._wait_for_idle(root, app)
            self.assertEqual(app.organize_workflow_stage.get(), "confirm")
            self.assertEqual(len(app.preview_data), 2)
            self.assertEqual(sum(1 for value in app.file_check_vars.values() if value.get()), 1)
            self.assertEqual(app.organize_result_title.get(), "扫描到 2 个 PDF")
            self.assertIn("可处理 1", app.organize_result_detail.get())
            self.assertIn("文件名无效 1", app.organize_result_detail.get())
        finally:
            app._on_closing()

    def test_organize_history_rerun_summary_uses_loaded_selection_not_all_files(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app._pending_organize_rerun_files = {"retry.pdf", "missing.pdf"}
            app._apply_organize_preview(
                OrganizePreviewResult(
                    rows=[
                        OrganizePreviewRow(
                            relative_path="retry.pdf",
                            company="甲公司",
                            target="甲公司",
                            selectable=True,
                        ),
                        OrganizePreviewRow(
                            relative_path="other.pdf",
                            company="乙公司",
                            target="乙公司",
                            selectable=True,
                        ),
                    ],
                    total_count=2,
                    selectable_count=2,
                )
            )

            self.assertTrue(app.file_check_vars["retry.pdf"].get())
            self.assertFalse(app.file_check_vars["other.pdf"].get())
            self.assertIn("选择 1 个", app.organize_workflow_status_text.get())
            self.assertIn("1 个历史文件", app.organize_workflow_status_text.get())
        finally:
            app._on_closing()

    def test_organize_scan_runs_off_tk_main_thread_and_restores_controls(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        invoice_dir = Path(self._data_dir.name) / "后台扫描"
        invoice_dir.mkdir()
        started = threading.Event()
        release = threading.Event()
        worker_threads = []

        def slow_preview(*args, **kwargs):
            worker_threads.append(threading.current_thread())
            started.set()
            release.wait(3)
            return OrganizePreviewResult()

        try:
            app.organize_folder_path.set(str(invoice_dir))
            with patch("invoice_tool.ui.app.OrganizeService.preview", side_effect=slow_preview):
                self.assertTrue(app._scan_files())
                self.assertTrue(started.wait(1))
                self.assertIsNot(worker_threads[0], threading.main_thread())
                self.assertTrue(app.is_running)
                self.assertEqual(app.org_scan_btn.cget("state"), "disabled")
                self.assertEqual(app.cancel_org_btn.cget("state"), "normal")
                release.set()
                self._wait_for_idle(root, app)
            self.assertEqual(app.org_scan_btn.cget("state"), "normal")
            self.assertEqual(app.cancel_org_btn.cget("state"), "disabled")
        finally:
            release.set()
            if app.is_running:
                self._wait_for_idle(root, app)
            app._on_closing()

    def test_organize_failure_detail_enables_targeted_retry(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        invoice_dir = Path(self._data_dir.name) / "retry"
        invoice_dir.mkdir()
        filename = "dzfp_1001_示例公司.pdf"
        (invoice_dir / filename).write_bytes(b"pdf")
        try:
            app.organize_folder_path.set(str(invoice_dir))
            app.preview_data = {
                filename: {
                    "filename": filename,
                    "company": "示例公司",
                    "target": "示例公司",
                    "valid": True,
                    "already_organized": False,
                }
            }
            app.file_check_vars = {filename: tk.BooleanVar(value=False)}
            result = OrganizeTaskResult(
                fail_count=1,
                result_rows=[
                    OrganizeResultRow(
                        status="失败",
                        filename=filename,
                        company="示例公司",
                        detail="文件被占用",
                        retryable=True,
                    )
                ],
            )
            app._apply_organize_execution_result(result)

            self.assertEqual(app.retry_org_btn.cget("state"), "normal")
            self.assertEqual(app.organize_failed_files, [filename])
            values = app.file_tree.item(app.file_tree.get_children()[0], "values")
            self.assertIn("文件被占用", values[3])
            with patch.object(app, "_execute_organize") as execute:
                app._retry_failed_organize()
            self.assertTrue(app.file_check_vars[filename].get())
            execute.assert_called_once_with()
        finally:
            app._on_closing()

    def test_write_task_pause_button_toggles_pause_and_resume(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            with app._lock:
                app.is_running = True
                app._active_task_kind = "write"
            app.pause_org_btn.config(state="normal")

            app._toggle_pause_task()
            self.assertTrue(app._pause_flag.is_set())
            self.assertIn("继续", app.pause_org_btn.cget("text"))
            self.assertIn("当前文件", app.organize_workflow_status_text.get())

            app._toggle_pause_task()
            self.assertFalse(app._pause_flag.is_set())
            self.assertIn("暂停", app.pause_org_btn.cget("text"))
        finally:
            with app._lock:
                app.is_running = False
                app._active_task_kind = ""
            app._pause_flag.clear()
            app._on_closing()

    def test_filter_copy_failure_enables_safe_retry_action(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app._set_filter_results(
                [FilterResultRow(status="复制失败", invoice_number="1001", detail="权限不足")]
            )
            self.assertEqual(app.filter_retry_btn.cget("state"), "normal")
            with (
                patch("invoice_tool.ui.app.messagebox.askyesno", return_value=True),
                patch.object(app, "_run_filter") as run_filter,
            ):
                app._retry_failed_filter()
            run_filter.assert_called_once_with(skip_confirmation=True)
            self.assertIn("重试 1 个", app.filter_workflow_status_text.get())
        finally:
            app._on_closing()

    def test_filter_execution_requires_preview_for_current_inputs(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        data_dir = Path(self._data_dir.name)
        excel_path = data_dir / "book.xlsx"
        pdf_folder = data_dir / "pdfs"
        output_folder = data_dir / "output"
        excel_path.write_bytes(b"placeholder")
        pdf_folder.mkdir()
        try:
            app.excel_path.set(str(excel_path))
            app.pdf_folder.set(str(pdf_folder))
            app.auto_output_by_sheet.set(False)
            app.manual_output_folder.set(str(output_folder))
            app._last_filter_preview_signature = ("stale",)
            app._last_filter_preview_result = FilterPreviewResult(
                invoice_numbers=["1001"],
                excel_column_name="发票号码",
                sheet_name="Sheet1",
                columns=["发票号码"],
                mapping={},
                conflicts=[],
                matched=[],
                not_found=["1001"],
                pdf_stats=PdfScanStats(),
            )

            with (
                patch("invoice_tool.ui.app.messagebox.showwarning") as warning,
                patch.object(app, "_try_begin_task") as begin_task,
            ):
                app._run_filter()

            warning.assert_called_once()
            begin_task.assert_not_called()
            self.assertFalse((data_dir / "active_task.json").exists())
            self.assertIn("请先预览", app.filter_workflow_status_text.get())
        finally:
            app._on_closing()

    def test_filter_execution_final_confirmation_is_non_mutating_when_cancelled(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        data_dir = Path(self._data_dir.name)
        excel_path = data_dir / "book.xlsx"
        pdf_folder = data_dir / "pdfs"
        output_folder = data_dir / "output"
        excel_path.write_bytes(b"placeholder")
        pdf_folder.mkdir()
        try:
            app.excel_path.set(str(excel_path))
            app.pdf_folder.set(str(pdf_folder))
            app.auto_output_by_sheet.set(False)
            app.manual_output_folder.set(str(output_folder))
            preview = FilterPreviewResult(
                invoice_numbers=["1001"],
                excel_column_name="发票号码",
                sheet_name="Sheet1",
                columns=["发票号码"],
                mapping={"1001": "invoice_1001.pdf"},
                conflicts=[],
                matched=[{"invoice": "1001", "pdf": "invoice_1001.pdf"}],
                not_found=[],
                pdf_stats=PdfScanStats(scanned=1, valid_named=1),
            )
            app._last_filter_preview_result = preview
            app._last_filter_preview_signature = app._filter_preview_context_signature(output_folder)

            with (
                patch("invoice_tool.ui.app.messagebox.askyesno", return_value=False) as confirm,
                patch.object(app, "_try_begin_task") as begin_task,
            ):
                app._run_filter()

            confirm.assert_called_once()
            begin_task.assert_not_called()
            self.assertFalse(output_folder.exists())
            self.assertFalse((data_dir / "active_task.json").exists())
            self.assertIn("尚未复制", app.filter_workflow_status_text.get())
        finally:
            app._on_closing()

    def test_workbook_analysis_runs_off_tk_main_thread(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        excel_path = Path(self._data_dir.name) / "book.xlsx"
        excel_path.write_bytes(b"placeholder")
        started = threading.Event()
        release = threading.Event()
        worker_threads = []

        def slow_analysis(*args, **kwargs):
            worker_threads.append(threading.current_thread())
            started.set()
            release.wait(3)
            return WorkbookAnalysisResult(
                workbook_name=excel_path.name,
                sheet_profiles=[
                    WorkbookSheetProfile(
                        sheet_name="Sheet1",
                        row_count=1,
                        column_count=1,
                        columns=["发票号码"],
                        selected_invoice_column="发票号码",
                        usable=True,
                        recommended=True,
                    )
                ],
                recommended_sheet_name="Sheet1",
                total_sheet_count=1,
                usable_sheet_count=1,
            )

        try:
            app.excel_path.set(str(excel_path))
            with patch("invoice_tool.ui.app.WorkbookAnalyzerService.analyze", side_effect=slow_analysis):
                self.assertTrue(app._refresh_workbook_analysis())
                self.assertTrue(started.wait(1))
                self.assertIsNot(worker_threads[0], threading.main_thread())
                self.assertEqual(app.workbook_analysis_btn.cget("state"), "disabled")
                release.set()
                self._wait_for_idle(root, app)
            self.assertEqual(app.excel_sheet_name.get(), "Sheet1")
            self.assertEqual(app.workbook_analysis_result.total_sheet_count, 1)
            self.assertEqual(app.workbook_analysis_btn.cget("state"), "normal")
        finally:
            release.set()
            if app.is_running:
                self._wait_for_idle(root, app)
            app._on_closing()

    def test_filter_preview_can_be_cancelled_without_file_changes(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        root_dir = Path(self._data_dir.name)
        excel_path = root_dir / "book.xlsx"
        excel_path.write_bytes(b"placeholder")
        pdf_dir = root_dir / "pdf"
        output_dir = root_dir / "output"
        pdf_dir.mkdir()
        output_dir.mkdir()
        source_pdf = pdf_dir / "dzfp_1001_测试公司.pdf"
        source_pdf.write_bytes(b"pdf")
        started = threading.Event()
        worker_threads = []

        def cancellable_preview(*args, cancel_requested=None, **kwargs):
            worker_threads.append(threading.current_thread())
            started.set()
            deadline = time.monotonic() + 3
            while time.monotonic() < deadline and not cancel_requested():
                time.sleep(0.01)
            raise CancelledError()

        try:
            app.excel_path.set(str(excel_path))
            app.excel_sheet_name.set("Sheet1")
            app.pdf_folder.set(str(pdf_dir))
            app.auto_output_by_sheet.set(False)
            app.manual_output_folder.set(str(output_dir))
            app._sync_output_folder_mode_ui()
            with patch("invoice_tool.ui.app.FilterService.preview", side_effect=cancellable_preview):
                self.assertTrue(app._preview_filter())
                self.assertTrue(started.wait(1))
                self.assertIsNot(worker_threads[0], threading.main_thread())
                self.assertEqual(app.cancel_flt_btn.cget("state"), "normal")
                app._cancel_task()
                self._wait_for_idle(root, app)
            self.assertTrue(source_pdf.exists())
            self.assertEqual(list(output_dir.iterdir()), [])
            self.assertIn("预览已取消", app.filter_summary_title.get())
            self.assertEqual(app.cancel_flt_btn.cget("state"), "disabled")
        finally:
            if app.is_running:
                app._cancel_flag.set()
                self._wait_for_idle(root, app)
            app._on_closing()

    def test_history_page_uses_master_detail_with_fixed_action_bar(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            root.update_idletasks()
            self.assertEqual(app.history_action_bar.pack_info()["side"], "bottom")
            self.assertIs(app.history_tree.master.master, app.history_split)
            self.assertIs(app.history_detail_panel.master, app.history_split)
            self.assertIs(app.history_rollback_btn.master, app.history_rerun_btn.master)
            self.assertIs(app.history_rollback_btn.master, app.history_view_btn.master)
            self.assertIs(app.history_view_btn.master, app.history_open_btn.master)
            self.assertIs(app.history_open_btn.master, app.history_clear_btn.master)
            self.assertIs(app.history_clear_btn.master, app.history_refresh_btn.master)
            self.assertIsNot(app.history_rollback_btn.master, app.history_action_bar)
        finally:
            app._on_closing()

    def test_history_filter_task_can_be_loaded_without_executing_files(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        root_dir = Path(self._data_dir.name)
        excel_path = root_dir / "history.xlsx"
        excel_path.write_bytes(b"excel")
        pdf_dir = root_dir / "pdf-history"
        output_dir = root_dir / "out-history"
        pdf_dir.mkdir()
        output_dir.mkdir()
        try:
            app.all_history = [
                {
                    "time": "2026-07-22 11:00:00",
                    "folder": str(pdf_dir),
                    "count": 1,
                    "type": "筛选",
                    "moves": [],
                    "result_rows": [{"status": "复制失败", "invoice_number": "1001"}],
                    "rerun": {
                        "type": "筛选",
                        "excel_path": str(excel_path),
                        "pdf_folder": str(pdf_dir),
                        "output_dir": str(output_dir),
                        "invoice_index": 1,
                        "recursive": False,
                        "sheet_name": "Sheet1",
                        "invoice_column_name": "发票号码",
                        "company_column_name": "公司名称",
                        "filter_column_name": "状态",
                        "filter_mode": "等于任一",
                        "filter_values": "正常",
                        "company_exclude_keywords": "测试",
                        "rule_preset_id": "standard_digital",
                        "custom_invoice_aliases": "票号",
                    },
                }
            ]
            app._refresh_history_tree()
            self.assertEqual(app.history_rerun_btn.cget("state"), "normal")

            with patch("invoice_tool.ui.app.messagebox.askyesno", return_value=True):
                app._load_history_for_rerun()

            self.assertEqual(app._workspace_page_key, "filter")
            self.assertEqual(app.excel_path.get(), str(excel_path))
            self.assertEqual(app.pdf_folder.get(), str(pdf_dir))
            self.assertEqual(app.manual_output_folder.get(), str(output_dir))
            self.assertEqual(app.row_filter_values.get(), "正常")
            self.assertEqual(list(output_dir.iterdir()), [])
            self.assertIn("尚未复制", app.status_var.get())
        finally:
            app._on_closing()

    def test_history_selection_shows_safe_rollback_summary(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app.all_history = [
                {
                    "time": "2026-07-22 09:30:00",
                    "folder": str(Path(self._data_dir.name)),
                    "count": 1,
                    "type": "整理",
                    "moves": [
                        {
                            "filename": "示例发票.pdf",
                            "source": str(Path(self._data_dir.name) / "示例发票.pdf"),
                            "target": str(Path(self._data_dir.name) / "公司" / "示例发票.pdf"),
                            "operation_root": str(Path(self._data_dir.name)),
                            "fingerprint": {"algorithm": "sha256", "sha256": "a" * 64, "size": 3},
                        }
                    ],
                }
            ]
            app._refresh_history_tree()
            self.assertEqual(app.history_detail_title.get(), "2026-07-22 09:30:00")
            self.assertIn("可校验回滚", app.history_detail_safety.get())
            self.assertEqual(app.history_rollback_btn.cget("state"), "normal")
            self.assertEqual(app.history_preview_listbox.get(0), "示例发票.pdf")
        finally:
            app._on_closing()

    def test_corrupt_history_remains_viewable_but_disables_rollback(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app.all_history = [
                {
                    "time": "损坏时间",
                    "folder": "",
                    "type": "整理",
                    "moves": None,
                    "report_files": None,
                }
            ]
            app._refresh_history_tree()
            self.assertEqual(len(app.filtered_history_indices), 1)
            self.assertIn("历史结构损坏", app.history_detail_safety.get())
            self.assertEqual(app.history_rollback_btn.cget("state"), "disabled")
            self.assertEqual(app.history_view_btn.cget("state"), "normal")
        finally:
            app._on_closing()

    def test_history_selection_survives_theme_rebuild(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app.all_history = [
                {"time": "2026-07-22 10:00:00", "folder": "first", "count": 0, "type": "整理", "moves": []},
                {"time": "2026-07-22 09:00:00", "folder": "second", "count": 0, "type": "筛选", "moves": []},
            ]
            app._refresh_history_tree()
            second_item = app.history_tree.get_children()[1]
            app.history_tree.selection_set(second_item)
            app._on_history_select()
            app._set_ui_theme("night")
            self.assertEqual(app._get_selected_history_index(), 1)
            self.assertEqual(app.history_detail_folder.get(), "second")
        finally:
            app._on_closing()

    def test_settings_tab_uses_scrollable_canvas(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            canvases = [child for child in app.settings_scroll_outer.winfo_children() if isinstance(child, tk.Canvas)]
            scrollbars = [child for child in app.settings_scroll_outer.winfo_children() if isinstance(child, ttk.Scrollbar)]
            self.assertTrue(canvases)
            self.assertTrue(scrollbars)
            self.assertEqual(app.settings_action_bar.pack_info()["side"], "bottom")
            self.assertIs(app.settings_save_btn.master, app.settings_import_btn.master)
            self.assertIs(app.settings_import_btn.master, app.settings_export_btn.master)
            self.assertIs(app.settings_export_btn.master, app.settings_reset_btn.master)
            self.assertIsNot(app.settings_save_btn.master, app.settings_action_bar)
        finally:
            app._on_closing()

    def test_settings_exposes_version_and_embedded_release_notes(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            self.assertEqual(app.release_notes_btn.cget("text"), "查看更新说明")
            with patch("invoice_tool.ui.app.messagebox.showinfo") as showinfo:
                app._show_release_notes()
            title, detail = showinfo.call_args.args
            self.assertIn("v6.0.0", title)
            self.assertIn("文件移动", detail)
            self.assertIn("诊断包", detail)
        finally:
            app._on_closing()

    def test_log_drawer_can_copy_sanitized_visible_content(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app.log_text.insert("end", "[12:00:00] 已脱敏日志\n")

            app._copy_log()
            root.update()

            self.assertIn("[12:00:00] 已脱敏日志", root.clipboard_get())
            self.assertIn("已复制", app.status_var.get())
            self.assertEqual(app.log_copy_btn.cget("text"), "复制")
        finally:
            app._on_closing()

    def test_config_commit_creates_backup_and_applies_runtime_values(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            current = app._collect_runtime_config()
            plan = normalize_config(
                {"ui_theme": "night", "company_name_index": 5},
                base=current,
                preset_ids=app._preset_by_id.keys(),
            )
            backup_path = app._commit_config_plan(plan, "测试配置")

            self.assertTrue(backup_path.exists())
            self.assertEqual(app.ui_theme.get(), "night")
            self.assertEqual(app.company_name_index.get(), 5)
            saved = load_json(Path(self._data_dir.name) / "config.json", {})
            self.assertEqual(saved["ui_theme"], "night")
            self.assertEqual(saved["company_name_index"], 5)
            self.assertIn("测试配置完成", app.settings_status_text.get())
        finally:
            app._on_closing()

    def test_failed_config_commit_restores_future_config_write_protection(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            snapshot = {"config_schema_version": 99, "future": "keep"}
            app._config_write_blocked_reason = "future schema"
            app._blocked_config_snapshot = snapshot
            current = app._collect_runtime_config()
            plan = normalize_config({"ui_theme": "night"}, base=current)

            with patch.object(app, "_save_config", return_value=False):
                with self.assertRaises(OSError):
                    app._commit_config_plan(plan, "测试失败")

            self.assertEqual(app._config_write_blocked_reason, "future schema")
            self.assertIs(app._blocked_config_snapshot, snapshot)
            self.assertEqual(app.ui_theme.get(), current["ui_theme"])
        finally:
            app._config_write_blocked_reason = ""
            app._on_closing()

    def test_diagnostic_export_handler_writes_sanitized_bundle(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        target = Path(self._data_dir.name) / "diagnostics.zip"
        log_path = Path(self._data_dir.name) / "app.log"
        log_path.write_text(r"ERROR C:\Finance\客户A\invoice.pdf 123456789012", encoding="utf-8")
        app.recent_errors = [
            {
                "time": "12:00:00",
                "level": "ERROR",
                "summary": "读取失败",
                "detail": r"读取 C:\Finance\客户A\invoice.pdf 失败",
            }
        ]
        try:
            with (
                patch("invoice_tool.ui.app.LOG_FILE", log_path),
                patch("invoice_tool.ui.app.filedialog.asksaveasfilename", return_value=str(target)),
                patch("invoice_tool.ui.app.messagebox.showinfo") as showinfo,
            ):
                app._export_diagnostic_bundle()

            self.assertTrue(target.exists())
            self.assertTrue(showinfo.called)
            self.assertIn("诊断包已导出", app.settings_status_text.get())
            with zipfile.ZipFile(target) as archive:
                combined = "\n".join(
                    archive.read(name).decode("utf-8") for name in archive.namelist()
                )
            self.assertNotIn("Finance", combined)
            self.assertNotIn("123456789012", combined)
        finally:
            app._on_closing()

    def test_future_config_is_not_overwritten_on_close(self):
        config_path = Path(self._data_dir.name) / "config.json"
        self.assertTrue(save_json(config_path, {"config_schema_version": 99, "future": "keep"}))
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            self.assertTrue(app._config_write_blocked_reason)
            self.assertFalse(app._save_config())
        finally:
            app._on_closing()
        self.assertEqual(load_json(config_path, {})["future"], "keep")

    def test_compact_header_keeps_notebook_near_top(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            root.update_idletasks()
            notebook_top = self._relative_y(app.notebook, app.root)
            self.assertLess(notebook_top, 210)
        finally:
            app._on_closing()

    def test_workspace_uses_task_navigation_and_tabless_page_stack(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            root.update_idletasks()
            self.assertEqual(tuple(app.workspace_nav_buttons), tuple(page.key for page in WORKSPACE_PAGES))
            self.assertIsInstance(app.notebook, WorkspacePageStack)
            self.assertEqual(app._workspace_page_key, "filter")
            self.assertEqual(app.notebook.select(), str(app.filter_frame))
            self.assertEqual(app.filter_frame.winfo_manager(), "grid")
            self.assertEqual(app.organize_frame.winfo_manager(), "")
            self.assertEqual(app.history_frame.winfo_manager(), "")
            self.assertEqual(app.settings_frame.winfo_manager(), "")
            self.assertEqual(app.workspace_page_title.get(), "Excel 筛选与 PDF 匹配")
        finally:
            app._on_closing()

    def test_workspace_navigation_updates_page_header_and_selection(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app._select_workspace_page("history", persist=False, focus_navigation=False)
            root.update()
            self.assertEqual(app.notebook.select(), str(app.history_frame))
            self.assertEqual(app.filter_frame.winfo_manager(), "")
            self.assertEqual(app.history_frame.winfo_manager(), "grid")
            self.assertEqual(app.workspace_page_title.get(), "任务历史")
            self.assertIn("安全回滚", app.workspace_workflow.get())
            self.assertEqual(app.workspace_page_counter.cget("text"), "03 / 04")
        finally:
            app._on_closing()

    def test_workspace_shortcut_persists_last_page_without_dropping_config(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            self.assertEqual(app._navigate_workspace_shortcut("settings"), "break")
            self.assertEqual(app._workspace_page_key, "settings")
            self.assertEqual(app.config["workspace_page"], "settings")
            self.assertEqual(app.config["rule_preset_id"], app.rule_preset_id.get())
        finally:
            app._on_closing()

    def test_workspace_theme_rebuild_preserves_current_task_page(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            app._select_workspace_page("history", persist=False, focus_navigation=False)
            app._set_ui_theme("night")
            self.assertEqual(app._workspace_page_key, "history")
            self.assertEqual(app.notebook.select(), str(app.history_frame))
            self.assertEqual(app.workspace_page_title.get(), "任务历史")
            self.assertEqual(app.ui_theme.get(), "night")
        finally:
            app._on_closing()

    def test_workspace_geometry_scales_with_tk_dpi(self):
        root = tk.Tk()
        root.withdraw()
        root.tk.call("tk", "scaling", InvoiceToolApp.BASE_TK_SCALING * 1.5)
        app = self._create_app(root)
        try:
            self.assertAlmostEqual(app._workspace_scale, 1.5, delta=0.05)
            self.assertEqual(
                app.workspace_sidebar.cget("width"),
                app._scaled(app.NAV_BASE_WIDTH),
            )
        finally:
            app._on_closing()

    def test_workspace_scale_supports_125_150_and_200_percent(self):
        root = tk.Tk()
        root.withdraw()
        try:
            for factor in (1.25, 1.5, 2.0):
                with self.subTest(factor=factor):
                    root.tk.call("tk", "scaling", InvoiceToolApp.BASE_TK_SCALING * factor)
                    self.assertAlmostEqual(
                        InvoiceToolApp._dpi_scale(root),
                        factor,
                        delta=0.06,
                    )
        finally:
            root.destroy()

    def test_workbook_analysis_details_are_collapsed_by_default(self):
        root = tk.Tk()
        root.withdraw()
        app = self._create_app(root)
        try:
            self.assertFalse(app.workbook_analysis_expanded.get())
            self.assertEqual(app.workbook_analysis_content.winfo_manager(), "")
        finally:
            app._on_closing()

    def test_status_colors_are_theme_specific(self):
        status_keys = [
            "status_success",
            "status_missing",
            "status_skip",
            "status_error",
            "status_conflict",
            "status_preview",
        ]
        for key in status_keys:
            self.assertIn(key, UI_THEME_PRESETS["day"])
            self.assertIn(key, UI_THEME_PRESETS["night"])
            self.assertNotEqual(UI_THEME_PRESETS["day"][key], UI_THEME_PRESETS["night"][key])


if __name__ == "__main__":
    unittest.main()
