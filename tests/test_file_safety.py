from __future__ import annotations

import queue
import tempfile
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from invoice_tool.core.file_safety import fingerprint_file
from invoice_tool.core.filtering import InvoiceFilter
from invoice_tool.core.organizer import InvoiceOrganizer
from invoice_tool.core.report import ReportExporter
from invoice_tool.core.services import FilterService, OrganizeService
from invoice_tool.infra.task_journal import TaskJournal
from invoice_tool.infra.storage import load_json, save_json
from invoice_tool.runtime import PANDAS_SUPPORT, pd
from invoice_tool.ui.app import InvoiceToolApp


class FileSafetyTests(unittest.TestCase):
    def test_organize_restores_file_when_recovery_journal_callback_fails(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "invoice.pdf"
            source.write_bytes(b"invoice")

            result = OrganizeService.run(
                folder=root,
                files=[source.name],
                preview_data={source.name: {"valid": True, "company": "测试公司"}},
                operation_callback=lambda _move: (_ for _ in ()).throw(OSError("journal full")),
            )

            self.assertEqual(result.success_count, 0)
            self.assertEqual(result.fail_count, 1)
            self.assertEqual(result.moves, [])
            self.assertTrue(source.exists())
            self.assertEqual(source.read_bytes(), b"invoice")
            self.assertFalse((root / "测试公司" / source.name).exists())
            self.assertIn("恢复日志写入失败", result.result_rows[0].detail)

    def test_organize_rejects_company_path_escape(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory) / "invoices"
            root.mkdir()
            source = root / "invoice.pdf"
            source.write_bytes(b"invoice")

            result = OrganizeService.run(
                folder=root,
                files=[source.name],
                preview_data={source.name: {"valid": True, "company": ".."}},
            )

            self.assertEqual(result.success_count, 0)
            self.assertEqual(result.fail_count, 1)
            self.assertEqual(result.moves, [])
            self.assertEqual(result.result_rows[0].status, "失败")
            self.assertTrue(result.result_rows[0].retryable)
            self.assertIn("公司目录名称不安全", result.result_rows[0].detail)
            self.assertTrue(source.exists())
            self.assertFalse((root.parent / source.name).exists())

    def test_recursive_organize_is_idempotent_for_direct_target_folder(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            company_folder = root / "Acme"
            company_folder.mkdir()
            source = company_folder / "invoice.pdf"
            source.write_bytes(b"already organized")

            result = OrganizeService.run(
                folder=root,
                files=[str(Path("Acme") / source.name)],
                preview_data={
                    str(Path("Acme") / source.name): {"valid": True, "company": "Acme"},
                },
            )

            self.assertEqual(result.success_count, 0)
            self.assertEqual(result.skip_count, 1)
            self.assertEqual(result.moves, [])
            self.assertEqual(source.read_bytes(), b"already organized")
            self.assertEqual([path.name for path in company_folder.iterdir()], ["invoice.pdf"])

    @unittest.skipUnless(PANDAS_SUPPORT, "pandas required")
    def test_duplicate_pdf_invoice_is_blocked_from_export(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            excel_path = root / "sample.xlsx"
            pdf_folder = root / "pdfs"
            output_root = root / "output"
            pdf_folder.mkdir()
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": ["1001"]}).to_excel(writer, sheet_name="Sheet1", index=False)
            (pdf_folder / "a_1001_A.pdf").write_bytes(b"first")
            (pdf_folder / "b_1001_B.pdf").write_bytes(b"second")

            result = FilterService.run(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                output_dir=output_root,
                invoice_index=1,
            )

            self.assertEqual(result.found_count, 0)
            self.assertEqual(result.not_found, [])
            self.assertEqual(len(result.conflicts), 1)
            self.assertEqual([row.status for row in result.result_rows], ["重复冲突"])
            self.assertFalse((output_root / "a_1001_A.pdf").exists())
            self.assertFalse((output_root / "b_1001_B.pdf").exists())

    def test_company_target_rejects_current_directory_and_reserved_name(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            for company in (".", "CON", "Acme/West"):
                with self.subTest(company=company):
                    with self.assertRaises(ValueError):
                        InvoiceOrganizer.resolve_company_target(root, company)

    def test_company_target_accepts_chinese_spaces_and_parentheses(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            target = InvoiceOrganizer.resolve_company_target(root, "测试 公司（北京）")
            self.assertEqual(target, (root / "测试 公司（北京）").resolve())

    def test_organize_records_fingerprint_and_valid_rollback_succeeds(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "invoice.pdf"
            source.write_bytes(b"original invoice")

            result = OrganizeService.run(
                folder=root,
                files=[source.name],
                preview_data={source.name: {"valid": True, "company": "测试公司"}},
            )

            self.assertEqual(result.success_count, 1)
            move = result.moves[0]
            self.assertEqual(move["operation_root"], str(root.resolve()))
            self.assertEqual(move["fingerprint"]["sha256"], fingerprint_file(Path(move["target"]))["sha256"])

            ok, error = InvoiceOrganizer.rollback_single_move(move)
            self.assertTrue(ok, error)
            self.assertEqual(source.read_bytes(), b"original invoice")

    def test_rollback_refuses_replacement_file(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "invoice.pdf"
            source.write_bytes(b"original invoice")
            result = OrganizeService.run(
                folder=root,
                files=[source.name],
                preview_data={source.name: {"valid": True, "company": "测试公司"}},
            )
            move = result.moves[0]
            target = Path(move["target"])
            target.write_bytes(b"replacement")

            ok, error = InvoiceOrganizer.rollback_single_move(move)

            self.assertFalse(ok)
            self.assertIn("内容已变化", error)
            self.assertEqual(target.read_bytes(), b"replacement")
            self.assertFalse(source.exists())

    def test_rollback_refuses_to_overwrite_new_source_file(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "invoice.pdf"
            source.write_bytes(b"original invoice")
            result = OrganizeService.run(
                folder=root,
                files=[source.name],
                preview_data={source.name: {"valid": True, "company": "测试公司"}},
            )
            move = result.moves[0]
            source.write_bytes(b"new source file")

            ok, error = InvoiceOrganizer.rollback_single_move(move)

            self.assertFalse(ok)
            self.assertIn("原位置已有同名文件", error)
            self.assertEqual(source.read_bytes(), b"new source file")
            self.assertTrue(Path(move["target"]).exists())

    def test_legacy_history_without_fingerprint_is_not_destructive(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            target = root / "company" / "invoice.pdf"
            target.parent.mkdir()
            target.write_bytes(b"keep me")
            move = {
                "source": str(root / "invoice.pdf"),
                "target": str(target),
                "filename": target.name,
            }

            ok, error = InvoiceOrganizer.rollback_single_move(move)

            self.assertFalse(ok)
            self.assertIn("旧历史", error)
            self.assertEqual(target.read_bytes(), b"keep me")

    def test_filter_rollback_refuses_changed_export(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_root = Path(temporary_directory)
            target = output_root / "invoice.pdf"
            target.write_bytes(b"exported")
            record = {
                "target": str(target),
                "filename": target.name,
                "output_root": str(output_root),
                "fingerprint": fingerprint_file(target),
            }
            target.write_bytes(b"user replacement")

            ok, error = InvoiceOrganizer.delete_recorded_file(record)

            self.assertFalse(ok)
            self.assertIn("内容已变化", error)
            self.assertEqual(target.read_bytes(), b"user replacement")

    def test_filter_service_records_verified_export_for_safe_rollback(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_folder = root / "pdfs"
            output_root = root / "output"
            pdf_folder.mkdir()
            source = pdf_folder / "invoice_1001.pdf"
            source.write_bytes(b"pdf payload")
            excel_result = {
                "invoice_numbers": ["1001"],
                "invoice_column_name": "发票号码",
                "sheet_name": "Sheet1",
                "columns": ["发票号码"],
                "company_column_name": "",
                "filter_column_name": "",
                "filter_mode": "不过滤",
                "filter_values": [],
                "source_row_count": 1,
                "filtered_out_count": 0,
            }
            recorded = []
            with mock.patch.object(InvoiceFilter, "read_invoice_records", return_value=excel_result), \
                 mock.patch.object(
                     InvoiceFilter,
                     "build_pdf_mapping",
                     return_value=({"1001": source.name}, [], {"scanned": 1, "valid_named": 1, "invalid_named": 0, "duplicates": 0}),
                 ), \
                 mock.patch.object(ReportExporter, "export_filter_report", return_value=None):
                result = FilterService.run(
                    excel_path=root / "sample.xlsx",
                    pdf_folder=pdf_folder,
                    output_dir=output_root,
                    invoice_index=1,
                    operation_callback=recorded.append,
                )

            self.assertEqual(result.found_count, 1)
            self.assertEqual(recorded, result.moves)
            move = result.moves[0]
            self.assertEqual(move["output_root"], str(output_root.resolve()))
            ok, error = InvoiceOrganizer.delete_recorded_file(move)
            self.assertTrue(ok, error)
            self.assertFalse(Path(move["target"]).exists())
            self.assertFalse(any(path.name.endswith(".tmp") for path in output_root.iterdir()))

    def test_filter_removes_copy_when_recovery_journal_callback_fails(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_folder = root / "pdfs"
            output_root = root / "output"
            pdf_folder.mkdir()
            source = pdf_folder / "invoice_1001.pdf"
            source.write_bytes(b"pdf payload")
            excel_result = {
                "invoice_numbers": ["1001"],
                "invoice_column_name": "发票号码",
                "sheet_name": "Sheet1",
                "columns": ["发票号码"],
                "company_column_name": "",
                "filter_column_name": "",
                "filter_mode": "不过滤",
                "filter_values": [],
                "source_row_count": 1,
                "filtered_out_count": 0,
            }
            with mock.patch.object(InvoiceFilter, "read_invoice_records", return_value=excel_result), \
                 mock.patch.object(
                     InvoiceFilter,
                     "build_pdf_mapping",
                     return_value=({"1001": source.name}, [], {"scanned": 1, "valid_named": 1, "invalid_named": 0, "duplicates": 0}),
                 ), \
                 mock.patch.object(ReportExporter, "export_filter_report", return_value=None):
                result = FilterService.run(
                    excel_path=root / "sample.xlsx",
                    pdf_folder=pdf_folder,
                    output_dir=output_root,
                    invoice_index=1,
                    operation_callback=lambda _move: (_ for _ in ()).throw(OSError("journal full")),
                )

            self.assertEqual(result.found_count, 0)
            self.assertEqual(result.copy_fail_count, 1)
            self.assertEqual(result.moves, [])
            self.assertFalse((output_root / source.name).exists())
            self.assertEqual(source.read_bytes(), b"pdf payload")

    def test_filter_removes_report_when_report_journal_callback_fails(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_folder = root / "pdfs"
            output_root = root / "output"
            pdf_folder.mkdir()
            output_root.mkdir()
            excel_result = {
                "invoice_numbers": [],
                "invoice_column_name": "发票号码",
                "sheet_name": "Sheet1",
                "columns": ["发票号码"],
                "company_column_name": "",
                "filter_column_name": "",
                "filter_mode": "不过滤",
                "filter_values": [],
                "source_row_count": 0,
                "filtered_out_count": 0,
            }
            report_path = output_root / "report.xlsx"

            def create_report(*_args, **_kwargs):
                report_path.write_bytes(b"report")
                return report_path

            with mock.patch.object(InvoiceFilter, "read_invoice_records", return_value=excel_result), \
                 mock.patch.object(
                     InvoiceFilter,
                     "build_pdf_mapping",
                     return_value=({}, [], {"scanned": 0, "valid_named": 0, "invalid_named": 0, "duplicates": 0}),
                 ), \
                 mock.patch.object(ReportExporter, "export_filter_report", side_effect=create_report):
                with self.assertRaisesRegex(OSError, "报告恢复日志写入失败"):
                    FilterService.run(
                        excel_path=root / "sample.xlsx",
                        pdf_folder=pdf_folder,
                        output_dir=output_root,
                        invoice_index=1,
                        report_callback=lambda _path: (_ for _ in ()).throw(OSError("journal full")),
                    )

            self.assertFalse(report_path.exists())

    @unittest.skipUnless(PANDAS_SUPPORT, "pandas required")
    def test_existing_different_export_is_reported_as_conflict_and_preserved(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            excel_path = root / "sample.xlsx"
            pdf_folder = root / "pdfs"
            output_root = root / "output"
            pdf_folder.mkdir()
            output_root.mkdir()
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": ["1001"]}).to_excel(writer, sheet_name="Sheet1", index=False)
            source = pdf_folder / "invoice_1001.pdf"
            source.write_bytes(b"source")
            target = output_root / source.name
            target.write_bytes(b"user file")

            result = FilterService.run(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                output_dir=output_root,
                invoice_index=1,
            )

            self.assertEqual(result.found_count, 0)
            self.assertEqual(result.skip_count, 0)
            self.assertEqual(result.target_conflict_count, 1)
            self.assertEqual(result.result_rows[0].status, "同名冲突")
            self.assertEqual(target.read_bytes(), b"user file")

    @unittest.skipUnless(PANDAS_SUPPORT, "pandas required")
    def test_existing_identical_export_is_skipped(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            excel_path = root / "sample.xlsx"
            pdf_folder = root / "pdfs"
            output_root = root / "output"
            pdf_folder.mkdir()
            output_root.mkdir()
            with pd.ExcelWriter(excel_path) as writer:
                pd.DataFrame({"发票号码": ["1001"]}).to_excel(writer, sheet_name="Sheet1", index=False)
            source = pdf_folder / "invoice_1001.pdf"
            source.write_bytes(b"same")
            (output_root / source.name).write_bytes(b"same")

            result = FilterService.run(
                excel_path=excel_path,
                pdf_folder=pdf_folder,
                output_dir=output_root,
                invoice_index=1,
            )

            self.assertEqual(result.found_count, 0)
            self.assertEqual(result.skip_count, 1)
            self.assertEqual(result.target_conflict_count, 0)
            self.assertEqual(result.result_rows[0].status, "已跳过")

    def test_report_failure_keeps_copied_file_in_recovery_journal(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_folder = root / "pdfs"
            output_root = root / "output"
            pdf_folder.mkdir()
            source = pdf_folder / "invoice_1001.pdf"
            source.write_bytes(b"pdf payload")
            journal_path = root / "active_task.json"
            journal = TaskJournal(journal_path)
            task_id = journal.begin("筛选", pdf_folder)
            excel_result = {
                "invoice_numbers": ["1001"],
                "invoice_column_name": "发票号码",
                "sheet_name": "Sheet1",
                "columns": ["发票号码"],
                "company_column_name": "",
                "filter_column_name": "",
                "filter_mode": "不过滤",
                "filter_values": [],
                "source_row_count": 1,
                "filtered_out_count": 0,
            }
            with mock.patch.object(InvoiceFilter, "read_invoice_records", return_value=excel_result), \
                 mock.patch.object(
                     InvoiceFilter,
                     "build_pdf_mapping",
                     return_value=({"1001": source.name}, [], {"scanned": 1, "valid_named": 1, "invalid_named": 0, "duplicates": 0}),
                 ), \
                 mock.patch.object(ReportExporter, "export_filter_report", side_effect=OSError("report failed")):
                with self.assertRaises(OSError):
                    FilterService.run(
                        excel_path=root / "sample.xlsx",
                        pdf_folder=pdf_folder,
                        output_dir=output_root,
                        invoice_index=1,
                        operation_callback=lambda move: journal.record_move(task_id, move),
                    )

            recovered = TaskJournal(journal_path).load()
            self.assertIsNotNone(recovered)
            self.assertEqual(len(recovered["moves"]), 1)
            copied = Path(recovered["moves"][0]["target"])
            self.assertTrue(copied.exists())
            self.assertEqual(copied.read_bytes(), b"pdf payload")


class TaskJournalTests(unittest.TestCase):
    def test_load_json_tolerates_invalid_utf8(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            target = Path(temporary_directory) / "history.json"
            target.write_bytes(b"\xff\xfe\x00")

            loaded = load_json(target, [])

            self.assertEqual(loaded, [])
            self.assertTrue(target.exists())

    def test_invalid_config_or_history_is_quarantined_before_default_can_be_saved(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            history_path = root / "history.json"
            original = b'{"unexpected": "object"}'
            history_path.write_bytes(original)

            loaded = load_json(
                history_path,
                [],
                expected_type=list,
                quarantine_invalid=True,
            )

            self.assertEqual(loaded, [])
            self.assertFalse(history_path.exists())
            backups = list(root.glob("history.corrupt-*.json"))
            self.assertEqual(len(backups), 1)
            self.assertEqual(backups[0].read_bytes(), original)
            self.assertTrue(save_json(history_path, loaded))
            self.assertEqual(backups[0].read_bytes(), original)

    def test_atomic_json_save_cleans_temporary_file_on_serialization_error(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            target = Path(temporary_directory) / "config.json"

            saved = save_json(target, {"not_json": object()})

            self.assertFalse(saved)
            self.assertFalse(target.exists())
            self.assertEqual(list(Path(temporary_directory).iterdir()), [])

    def test_corrupt_journal_is_quarantined_and_does_not_block_new_task(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            journal_path = root / "active_task.json"
            journal_path.write_bytes(b"\xffbroken")
            journal = TaskJournal(journal_path)

            self.assertIsNone(journal.load())
            self.assertFalse(journal_path.exists())
            quarantined = list(root.glob("active_task.invalid-*.json"))
            self.assertEqual(len(quarantined), 1)
            self.assertEqual(quarantined[0].read_bytes(), b"\xffbroken")

            task_id = journal.begin("整理", root)
            self.assertTrue(task_id)
            self.assertTrue(journal_path.exists())

    def test_future_journal_is_preserved_and_blocks_old_writer(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            journal_path = root / "active_task.json"
            self.assertTrue(
                save_json(
                    journal_path,
                    {"schema_version": 99, "task_id": "future-task", "moves": []},
                )
            )
            journal = TaskJournal(journal_path)

            self.assertIsNone(journal.load())
            self.assertTrue(journal_path.exists())
            with self.assertRaises(OSError):
                journal.begin("整理", root)

    def test_clear_with_task_id_never_removes_a_different_unloaded_journal(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            journal_path = root / "active_task.json"
            writer = TaskJournal(journal_path)
            writer.begin("整理", root)

            fresh_instance = TaskJournal(journal_path)
            self.assertFalse(fresh_instance.clear("different-task"))
            self.assertTrue(journal_path.exists())

    def test_failed_journal_update_rolls_back_in_memory_entry(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            journal_path = root / "active_task.json"
            journal = TaskJournal(journal_path)
            task_id = journal.begin("整理", root)
            first = {"filename": "first.pdf"}
            second = {"filename": "second.pdf"}

            with mock.patch("invoice_tool.infra.task_journal.save_json", return_value=False):
                self.assertFalse(journal.record_move(task_id, first))
            self.assertTrue(journal.record_move(task_id, second))

            loaded = TaskJournal(journal_path).load()
            self.assertIsNotNone(loaded)
            self.assertEqual(loaded["moves"], [second])

    def test_journal_round_trip_and_recovery_to_history(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            journal_path = Path(temporary_directory) / "active_task.json"
            journal = TaskJournal(journal_path)
            task_id = journal.begin("整理", Path(temporary_directory))
            move = {
                "source": str(Path(temporary_directory) / "invoice.pdf"),
                "target": str(Path(temporary_directory) / "company" / "invoice.pdf"),
                "filename": "invoice.pdf",
            }
            self.assertTrue(journal.record_move(task_id, move))

            app = SimpleNamespace(
                _task_journal=TaskJournal(journal_path),
                all_history=[],
                _save_history=lambda: True,
            )
            recovered = InvoiceToolApp._recover_interrupted_task(app)

            self.assertEqual(recovered, 1)
            self.assertEqual(app.all_history[0]["task_id"], task_id)
            self.assertTrue(app.all_history[0]["recovered"])
            self.assertFalse(journal_path.exists())

    def test_new_task_cannot_overwrite_unarchived_journal(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            journal_path = Path(temporary_directory) / "active_task.json"
            journal = TaskJournal(journal_path)
            journal.begin("整理", Path(temporary_directory))

            with self.assertRaises(OSError):
                journal.begin("筛选", Path(temporary_directory))

    def test_worker_threads_are_non_daemon(self):
        release = threading.Event()
        app = SimpleNamespace(_lock=threading.Lock(), _worker_thread=None)

        InvoiceToolApp._start_worker(app, release.wait, (), name="test-safe-worker")
        worker = app._worker_thread
        try:
            self.assertIsNotNone(worker)
            self.assertFalse(worker.daemon)
        finally:
            release.set()
            worker.join(timeout=2)

    def test_pause_waiter_blocks_until_resume_and_cancel_also_unblocks(self):
        app = SimpleNamespace(
            _pause_flag=threading.Event(),
            _cancel_flag=threading.Event(),
        )
        app._pause_flag.set()
        worker = threading.Thread(target=InvoiceToolApp._wait_if_paused, args=(app,))
        worker.start()
        worker.join(timeout=0.15)
        self.assertTrue(worker.is_alive())

        app._pause_flag.clear()
        worker.join(timeout=1)
        self.assertFalse(worker.is_alive())

        app._pause_flag.set()
        app._cancel_flag.set()
        InvoiceToolApp._wait_if_paused(app)

    def test_background_ui_events_are_queued_until_main_thread_drain(self):
        calls = []
        app = SimpleNamespace(
            _ui_events=queue.Queue(),
            _ui_event_pump_id="scheduled",
            _close_finalized=True,
        )

        worker = threading.Thread(
            target=InvoiceToolApp._post_ui,
            args=(app, lambda: calls.append("done")),
        )
        worker.start()
        worker.join(timeout=2)

        self.assertEqual(calls, [])
        InvoiceToolApp._drain_ui_events(app)
        self.assertEqual(calls, ["done"])

    def test_close_requests_cancel_and_waits_for_running_task(self):
        callbacks = []
        finalized = []

        class RootStub:
            @staticmethod
            def after(delay, callback):
                callbacks.append(callback)

        class StatusStub:
            value = ""

            def set(self, value):
                self.value = value

        app = SimpleNamespace(
            _close_finalized=False,
            _closing_requested=False,
            _lock=threading.Lock(),
            is_running=True,
            _cancel_flag=threading.Event(),
            status_var=StatusStub(),
            root=RootStub(),
        )
        app._wait_for_task_before_close = lambda: InvoiceToolApp._wait_for_task_before_close(app)
        app._finalize_close = lambda: finalized.append(True)

        with mock.patch("invoice_tool.ui.app.messagebox.askyesno", return_value=True):
            InvoiceToolApp._on_closing(app)

        self.assertTrue(app._cancel_flag.is_set())
        self.assertTrue(app._closing_requested)
        self.assertEqual(finalized, [])
        self.assertEqual(len(callbacks), 1)

        app.is_running = False
        callbacks.pop()()
        self.assertEqual(finalized, [True])


if __name__ == "__main__":
    unittest.main()
