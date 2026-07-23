from __future__ import annotations

import builtins
import importlib.util
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
RUNTIME_PATH = ROOT / "invoice_tool" / "runtime.py"


class RuntimeFallbackTests(unittest.TestCase):
    def test_all_optional_dependencies_can_be_absent_without_runtime_import_failure(self):
        blocked_roots = {"pandas", "ttkbootstrap", "tkinterdnd2", "openpyxl", "xlrd"}
        original_import = builtins.__import__

        def blocked_import(name, globals=None, locals=None, fromlist=(), level=0):
            if name.split(".", 1)[0] in blocked_roots:
                raise ImportError(f"blocked optional dependency: {name}")
            return original_import(name, globals, locals, fromlist, level)

        spec = importlib.util.spec_from_file_location("invoice_tool_runtime_fallback_test", RUNTIME_PATH)
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        with mock.patch("builtins.__import__", side_effect=blocked_import):
            spec.loader.exec_module(module)

        self.assertFalse(module.PANDAS_SUPPORT)
        self.assertFalse(module.MODERN_UI)
        self.assertFalse(module.DND_SUPPORT)
        self.assertFalse(module.OPENPYXL_SUPPORT)
        self.assertFalse(module.XLRD_SUPPORT)
        self.assertIsNone(module.pd)
        self.assertIsNone(module.ttkb)
        self.assertIsNone(module.TkinterDnD)
        self.assertIsNone(module.openpyxl)
        self.assertIsNone(module.xlrd)


if __name__ == "__main__":
    unittest.main()
