from __future__ import annotations

import logging

from .paths import LOG_FILE


logger = logging.getLogger("InvoiceTool")
logger.setLevel(logging.DEBUG)
logger.propagate = False

if not any(isinstance(handler, logging.FileHandler) and getattr(handler, "baseFilename", "") == str(LOG_FILE) for handler in logger.handlers):
    file_handler = logging.FileHandler(str(LOG_FILE), encoding="utf-8", mode="a")
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    )
    logger.addHandler(file_handler)
