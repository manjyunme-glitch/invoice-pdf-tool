import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.application.diagnostics import (
    build_diagnostic_snapshot,
    create_diagnostic_bundle,
    redact_diagnostic_text,
)


class DiagnosticsTests(unittest.TestCase):
    def test_redaction_removes_local_paths_network_paths_and_long_numbers(self):
        text = r"读取 C:\Finance\客户A\invoice.pdf 失败，发票号 123456789012，来源 \\server\share\a.pdf"
        redacted = redact_diagnostic_text(text)

        self.assertNotIn("Finance", redacted)
        self.assertNotIn("server", redacted)
        self.assertNotIn("123456789012", redacted)
        self.assertIn("<PATH>", redacted)
        self.assertIn("<NUMBER>", redacted)

    def test_snapshot_contains_environment_but_not_storage_paths(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            log_path = root / "app.log"
            log_path.write_text("error", encoding="utf-8")
            snapshot = build_diagnostic_snapshot(
                app_version="v6.0.0",
                capabilities={"pandas": True, "dnd": False},
                config_path=root / "config.json",
                history_path=root / "history.json",
                log_path=log_path,
                recent_errors=[{"time": "now", "level": "ERROR", "detail": rf"失败 {root}\1234567890"}],
            )

            serialized = json.dumps(snapshot, ensure_ascii=False)
            self.assertNotIn(str(root), serialized)
            self.assertNotIn("1234567890", serialized)
            self.assertTrue(snapshot["storage"]["log_exists"])
            self.assertEqual(snapshot["storage"]["log_size_bytes"], 5)
            self.assertFalse(snapshot["storage"]["config_directory_fallback"])
            self.assertIn("<PATH>", snapshot["recent_errors"][0]["message"])

    def test_bundle_contains_only_sanitized_diagnostics_and_log_tail(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            log_path = root / "app.log"
            log_path.write_text(
                "ignored\n" + r"ERROR C:\Finance\客户A\invoice.pdf 123456789012" + "\n",
                encoding="utf-8",
            )
            target = root / "diagnostics.zip"
            snapshot = {"application": {"version": "v6.0.0"}}

            create_diagnostic_bundle(target, snapshot, log_path=log_path, max_log_lines=1)

            with zipfile.ZipFile(target) as archive:
                self.assertEqual(
                    set(archive.namelist()),
                    {"diagnostics.json", "README.txt", "sanitized-log.txt"},
                )
                sanitized_log = archive.read("sanitized-log.txt").decode("utf-8")
                self.assertNotIn("Finance", sanitized_log)
                self.assertNotIn("123456789012", sanitized_log)
                self.assertIn("<PATH>", sanitized_log)


if __name__ == "__main__":
    unittest.main()
