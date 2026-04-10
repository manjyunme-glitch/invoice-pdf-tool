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

