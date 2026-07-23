from __future__ import annotations

import logging
import re


WINDOWS_PATH_PATTERN = re.compile(r"(?i)(?<![\w])(?:[a-z]:\\)[^\r\n\t\"<>|，,；;。]*")
UNC_PATH_PATTERN = re.compile(r"\\\\[^\\\s]+\\[^\r\n\t\"<>|，,；;。]*")
DOCUMENT_NAME_PATTERN = re.compile(
    r"(?i)(?<![\w.])[^\\/:*?\"<>|\r\n\t，,；;。]*?\.(?:pdf|xlsx?|xls|csv)(?![\w.])"
)
LONG_NUMBER_PATTERN = re.compile(r"(?<!\d)\d{8,}(?!\d)")
EMAIL_PATTERN = re.compile(r"(?i)(?<![\w.+-])[\w.+-]+@[\w.-]+\.[a-z]{2,}(?![\w.-])")


def redact_sensitive_text(value: object) -> str:
    """Remove common business identifiers from text intended for persistent logs."""

    text = str(value)
    text = WINDOWS_PATH_PATTERN.sub("<PATH>", text)
    text = UNC_PATH_PATTERN.sub("<NETWORK_PATH>", text)
    text = DOCUMENT_NAME_PATTERN.sub("<DOCUMENT>", text)
    text = EMAIL_PATTERN.sub("<EMAIL>", text)
    text = LONG_NUMBER_PATTERN.sub("<NUMBER>", text)
    return text


class RedactingFormatter(logging.Formatter):
    """Logging formatter that sanitizes the final message, including tracebacks."""

    def format(self, record: logging.LogRecord) -> str:
        return redact_sensitive_text(super().format(record))
