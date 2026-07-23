from __future__ import annotations

from typing import List, Optional

from .version import APP_VERSION, __version__


def main(argv: Optional[List[str]] = None) -> int:
    """Lazy public entry point so importing package metadata does not initialize Tk."""

    from .app import main as app_main

    return app_main(argv)

__all__ = ["APP_VERSION", "__version__", "main"]
