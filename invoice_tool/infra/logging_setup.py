from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler

from .paths import CONFIG_DIR_FALLBACK_REASON, LOG_FILE
from .privacy import RedactingFormatter


logger = logging.getLogger("InvoiceTool")
logger.setLevel(logging.DEBUG)
logger.propagate = False

if not any(
    isinstance(handler, logging.FileHandler)
    and getattr(handler, "baseFilename", "") == str(LOG_FILE)
    for handler in logger.handlers
):
    try:
        file_handler = RotatingFileHandler(
            str(LOG_FILE),
            encoding="utf-8",
            maxBytes=2 * 1024 * 1024,
            backupCount=3,
        )
        file_handler.setFormatter(
            RedactingFormatter(
                "%(asctime)s [%(levelname)s] %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        )
        logger.addHandler(file_handler)
    except OSError:
        # The GUI can still attach its in-memory handlers after startup.
        logger.addHandler(logging.NullHandler())

if CONFIG_DIR_FALLBACK_REASON:
    logger.warning(CONFIG_DIR_FALLBACK_REASON)
