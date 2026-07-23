from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path

from invoice_tool.cli import build_parser
from invoice_tool.version import APP_VERSION, WINDOWS_EXE_BASENAME, __version__


ROOT = Path(__file__).resolve().parents[1]


class ReleaseMetadataTests(unittest.TestCase):
    def test_version_has_one_python_source_of_truth(self):
        self.assertEqual(__version__, "6.0.0")
        self.assertEqual(APP_VERSION, "v6.0.0")
        self.assertEqual(WINDOWS_EXE_BASENAME, "invoice-pdf-tool-v6.0.0-windows-x64")
        parser = build_parser()
        self.assertIn(APP_VERSION, parser.prog)
        self.assertIn(APP_VERSION, parser.description or "")

    def test_release_files_agree_on_executable_name_and_version(self):
        expected_name = WINDOWS_EXE_BASENAME
        for relative_path in (
            "README.md",
            "打包为EXE.bat",
            "invoice-pdf-tool-v6.spec",
            "version_info.txt",
        ):
            text = (ROOT / relative_path).read_text(encoding="utf-8")
            self.assertIn("6.0.0", text, relative_path)
            self.assertIn(expected_name, text, relative_path)

    def test_v6_entry_is_importable_without_starting_gui(self):
        entry_path = ROOT / "发票处理工具v6.py"
        spec = importlib.util.spec_from_file_location("invoice_tool_v6_entry", entry_path)
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)

        self.assertTrue(callable(module.main))

    def test_build_dependency_is_not_required_for_runtime_install(self):
        runtime_requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8").lower()
        build_requirements = (ROOT / "requirements-build.txt").read_text(encoding="utf-8").lower()

        self.assertNotIn("pyinstaller", runtime_requirements)
        self.assertIn("pyinstaller", build_requirements)


if __name__ == "__main__":
    unittest.main()
