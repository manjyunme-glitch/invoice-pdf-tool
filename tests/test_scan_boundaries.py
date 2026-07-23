from __future__ import annotations

import os
import tempfile
import unittest
from concurrent.futures import CancelledError
from pathlib import Path
from unittest import mock

from invoice_tool.core.organizer import InvoiceOrganizer


class ScanBoundaryTests(unittest.TestCase):
    def test_empty_folder_and_case_insensitive_pdf_suffix(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            self.assertEqual(InvoiceOrganizer.scan_pdf_files(root), [])

            (root / "中文 名称（测试）.PDF").write_bytes(b"pdf")
            (root / "ignore.txt").write_text("text", encoding="utf-8")

            self.assertEqual(
                InvoiceOrganizer.scan_pdf_files(root),
                [Path("中文 名称（测试）.PDF")],
            )

    def test_invalid_scan_root_has_actionable_error(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            missing = root / "missing"
            regular_file = root / "not-a-folder"
            regular_file.write_bytes(b"file")

            with self.assertRaisesRegex(FileNotFoundError, "PDF目录不存在"):
                InvoiceOrganizer.scan_pdf_files(missing)
            with self.assertRaisesRegex(NotADirectoryError, "PDF路径不是文件夹"):
                InvoiceOrganizer.scan_pdf_files(regular_file)

    def test_recursive_scan_prunes_excluded_tree_before_enumerating_its_files(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "source"
            excluded = source / "output"
            included = source / "nested"
            excluded.mkdir(parents=True)
            included.mkdir()
            (excluded / "private_1001.pdf").write_bytes(b"excluded")
            (included / "invoice_1002.pdf").write_bytes(b"included")

            real_scandir = os.scandir
            enumerated = []

            def tracking_scandir(path):
                enumerated.append(Path(path).resolve())
                return real_scandir(path)

            with mock.patch("os.scandir", side_effect=tracking_scandir):
                matches = InvoiceOrganizer.scan_pdf_files(
                    source,
                    recursive=True,
                    exclude_dirs=[excluded],
                )

            self.assertEqual(matches, [Path("nested") / "invoice_1002.pdf"])
            self.assertNotIn(excluded.resolve(), enumerated)

    def test_recursive_scan_cancels_during_enumeration(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            for index in range(100):
                (root / f"invoice_{index}.pdf").write_bytes(b"x")
            calls = 0

            def cancel_requested():
                nonlocal calls
                calls += 1
                return calls > 20

            with self.assertRaises(CancelledError):
                InvoiceOrganizer.scan_pdf_files(
                    root,
                    recursive=True,
                    cancel_requested=cancel_requested,
                )

    def test_bulk_scan_returns_all_files_without_loading_contents(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            for index in range(1_000):
                (root / f"invoice_{index:04d}.pdf").touch()

            matches = InvoiceOrganizer.scan_pdf_files(root)

            self.assertEqual(len(matches), 1_000)
            self.assertEqual(matches[0], Path("invoice_0000.pdf"))
            self.assertEqual(matches[-1], Path("invoice_0999.pdf"))


if __name__ == "__main__":
    unittest.main()
