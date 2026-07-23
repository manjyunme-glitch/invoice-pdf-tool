from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from invoice_tool.infra.paths import _prepare_config_dir


class ConfigPathTests(unittest.TestCase):
    def test_unusable_preferred_config_path_uses_writable_fallback(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            preferred = root / "preferred"
            fallback = root / "fallback"
            preferred.write_bytes(b"not a directory")

            selected, warning = _prepare_config_dir(preferred, fallback)

            self.assertEqual(selected, fallback)
            self.assertTrue(fallback.is_dir())
            self.assertIn("已切换到临时目录", warning)

    def test_both_unusable_paths_return_preferred_with_explicit_warning(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            preferred = root / "preferred"
            fallback = root / "fallback"
            preferred.write_bytes(b"not a directory")
            fallback.write_bytes(b"not a directory")

            selected, warning = _prepare_config_dir(preferred, fallback)

            self.assertEqual(selected, preferred)
            self.assertIn("均不可用", warning)


if __name__ == "__main__":
    unittest.main()
