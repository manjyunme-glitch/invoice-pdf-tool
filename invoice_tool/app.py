from __future__ import annotations

import sys
import tkinter as tk
from typing import List, Optional

from .runtime import DND_SUPPORT, MODERN_UI, TkinterDnD, ttkb
from .ui import InvoiceToolApp


def run_gui() -> None:
    if DND_SUPPORT:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    if MODERN_UI:
        try:
            ttkb.Style(theme="cosmo")
        except Exception:
            pass

    InvoiceToolApp(root)
    root.mainloop()


def main(argv: Optional[List[str]] = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        run_gui()
        return 0

    if args[0] == "gui":
        run_gui()
        return 0

    from .cli import main as cli_main

    return cli_main(args)
