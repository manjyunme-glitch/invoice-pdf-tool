import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.ui import app as ui_app_module
from invoice_tool.ui import InvoiceToolApp as PublicInvoiceToolApp


class UiImportTests(unittest.TestCase):
    def test_ui_module_exposes_logging_for_log_drawer(self):
        self.assertTrue(hasattr(ui_app_module, "logging"))
        self.assertIsNotNone(ui_app_module.logging.Formatter)

    def test_public_ui_app_points_to_visual_refresh_class(self):
        self.assertTrue(PublicInvoiceToolApp.__module__.endswith("v521_app"))


if __name__ == "__main__":
    unittest.main()
