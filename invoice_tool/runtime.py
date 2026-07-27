from __future__ import annotations

import platform


if platform.system() == "Windows":
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except AttributeError:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


try:
    import pandas as pd

    PANDAS_SUPPORT = True
except ImportError:
    pd = None
    PANDAS_SUPPORT = False


try:
    import ttkbootstrap as ttkb

    # The application owns the complete palette for legacy Tk widgets.
    # ttkbootstrap patches their constructors and otherwise replaces explicit
    # colors (for example, the dark workspace sidebar) with its theme defaults.
    # Keep ttkbootstrap for ttk controls while opting legacy widgets out of that
    # automatic recoloring by default.
    from ttkbootstrap.widgets import TK_WIDGETS

    for _tk_widget in TK_WIDGETS:
        _original_init = _tk_widget.__init__
        if getattr(_original_init, "_invoice_tool_preserves_tk_colors", False):
            continue

        def _preserve_explicit_tk_colors(
            self,
            *args,
            __original_init=_original_init,
            **kwargs,
        ):
            kwargs.setdefault("autostyle", False)
            __original_init(self, *args, **kwargs)

        _preserve_explicit_tk_colors._invoice_tool_preserves_tk_colors = True
        _tk_widget.__init__ = _preserve_explicit_tk_colors

    MODERN_UI = True
except ImportError:
    ttkb = None
    MODERN_UI = False


try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_SUPPORT = True
except ImportError:
    DND_FILES = None
    TkinterDnD = None
    DND_SUPPORT = False


try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    OPENPYXL_SUPPORT = True
except ImportError:
    openpyxl = None
    Font = PatternFill = Alignment = Border = Side = None
    OPENPYXL_SUPPORT = False


try:
    import xlrd

    XLRD_SUPPORT = True
except ImportError:
    xlrd = None
    XLRD_SUPPORT = False
