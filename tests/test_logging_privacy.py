import logging
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.infra.privacy import RedactingFormatter, redact_sensitive_text


class LoggingPrivacyTests(unittest.TestCase):
    def test_sensitive_text_redacts_paths_documents_numbers_and_email(self):
        raw = (
            r"读取 C:\Finance\客户甲\dzfp_123456789012_客户甲.pdf 失败；"
            r"来源 \\server\share\客户乙.xlsx，联系 finance@example.com，"
            "备用文件 客户丙_876543210987.pdf"
        )

        sanitized = redact_sensitive_text(raw)

        for sensitive in (
            "Finance",
            "server",
            "客户甲",
            "客户乙",
            "客户丙",
            "123456789012",
            "876543210987",
            "finance@example.com",
        ):
            self.assertNotIn(sensitive, sanitized)
        self.assertIn("<PATH>", sanitized)
        self.assertIn("<NETWORK_PATH>", sanitized)
        self.assertIn("<DOCUMENT>", sanitized)
        self.assertIn("<EMAIL>", sanitized)

    def test_redacting_formatter_sanitizes_arguments_and_traceback(self):
        formatter = RedactingFormatter("%(levelname)s %(message)s")
        try:
            raise OSError(r"无法打开 C:\Private\客户甲\123456789012.pdf")
        except OSError:
            exc_info = sys.exc_info()
        record = logging.LogRecord(
            name="privacy-test",
            level=logging.ERROR,
            pathname=__file__,
            lineno=1,
            msg="导出失败：%s",
            args=(r"C:\Private\客户甲\123456789012.pdf",),
            exc_info=exc_info,
        )

        rendered = formatter.format(record)

        self.assertNotIn("Private", rendered)
        self.assertNotIn("客户甲", rendered)
        self.assertNotIn("123456789012", rendered)
        self.assertIn("<PATH>", rendered)


if __name__ == "__main__":
    unittest.main()
