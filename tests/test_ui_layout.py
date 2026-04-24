import sys
import tkinter as tk
import unittest
from pathlib import Path
from tkinter import ttk


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.ui import InvoiceToolApp
from invoice_tool.ui.app import UI_THEME_PRESETS


class UiLayoutTests(unittest.TestCase):
    @staticmethod
    def _relative_y(widget: tk.Widget, root: tk.Widget) -> int:
        y = 0
        current = widget
        while current is not root:
            y += current.winfo_y()
            current = current.master
        return y

    def test_filter_result_tree_uses_grid_scrollbars(self):
        root = tk.Tk()
        root.withdraw()
        app = InvoiceToolApp(root)
        try:
            tree_frame = app.filter_result_tree.master
            managers = [child.winfo_manager() for child in tree_frame.winfo_children()]
            self.assertTrue(managers)
            self.assertTrue(all(manager == "grid" for manager in managers))
        finally:
            app._on_closing()

    def test_settings_tab_uses_scrollable_canvas(self):
        root = tk.Tk()
        root.withdraw()
        app = InvoiceToolApp(root)
        try:
            outer_frames = app.settings_frame.winfo_children()
            self.assertTrue(outer_frames)
            outer = outer_frames[0]
            canvases = [child for child in outer.winfo_children() if isinstance(child, tk.Canvas)]
            scrollbars = [child for child in outer.winfo_children() if isinstance(child, ttk.Scrollbar)]
            self.assertTrue(canvases)
            self.assertTrue(scrollbars)
        finally:
            app._on_closing()

    def test_compact_header_keeps_notebook_near_top(self):
        root = tk.Tk()
        root.withdraw()
        app = InvoiceToolApp(root)
        try:
            root.update_idletasks()
            notebook_top = self._relative_y(app.notebook, app.root)
            self.assertLess(notebook_top, 210)
        finally:
            app._on_closing()

    def test_workbook_analysis_details_are_collapsed_by_default(self):
        root = tk.Tk()
        root.withdraw()
        app = InvoiceToolApp(root)
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
