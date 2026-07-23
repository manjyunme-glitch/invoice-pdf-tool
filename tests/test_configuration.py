import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from invoice_tool.application.configuration import (
    CONFIG_DEFAULTS,
    ConfigurationError,
    backup_config,
    default_config_plan,
    load_config_plan,
    normalize_config,
    save_config_export,
)
from invoice_tool.infra.storage import load_json, save_json


class ConfigurationTests(unittest.TestCase):
    def test_normalize_config_applies_defaults_and_preserves_unknown_keys(self):
        plan = normalize_config(
            {
                "ui_theme": "NIGHT",
                "company_name_index": 4,
                "future_setting": {"enabled": True},
            }
        )

        self.assertEqual(plan.config["ui_theme"], "night")
        self.assertEqual(plan.config["company_name_index"], 4)
        self.assertEqual(plan.config["workspace_page"], CONFIG_DEFAULTS["workspace_page"])
        self.assertEqual(plan.config["future_setting"], {"enabled": True})
        self.assertTrue(any("future_setting" in warning for warning in plan.warnings))

    def test_import_keeps_current_values_when_incoming_values_are_invalid(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "import.json"
            self.assertTrue(
                save_json(
                    path,
                    {
                        "company_name_index": -1,
                        "auto_output_by_sheet": "yes",
                        "rule_preset_id": "missing",
                        "excel_sheet_name": " 工作表 ",
                    },
                )
            )
            current = {
                **CONFIG_DEFAULTS,
                "company_name_index": 3,
                "auto_output_by_sheet": False,
                "rule_preset_id": "custom",
            }

            plan = load_config_plan(path, current, preset_ids={"standard_digital", "custom"})

            self.assertEqual(plan.config["company_name_index"], 3)
            self.assertFalse(plan.config["auto_output_by_sheet"])
            self.assertEqual(plan.config["rule_preset_id"], "custom")
            self.assertEqual(plan.config["excel_sheet_name"], " 工作表 ")
            self.assertEqual(len(plan.warnings), 3)

    def test_future_config_schema_is_rejected(self):
        with self.assertRaises(ConfigurationError):
            normalize_config({"config_schema_version": 999})

    def test_default_plan_resets_known_values_and_preserves_future_values(self):
        current = {
            **CONFIG_DEFAULTS,
            "ui_theme": "night",
            "company_name_index": 7,
            "future_setting": "keep",
        }

        plan = default_config_plan(current)

        self.assertEqual(plan.config["ui_theme"], "day")
        self.assertEqual(plan.config["company_name_index"], 2)
        self.assertEqual(plan.config["future_setting"], "keep")
        self.assertEqual({change.key for change in plan.changes}, {"ui_theme", "company_name_index"})

    def test_export_and_backup_are_valid_and_never_reuse_backup_name(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            export_path = directory / "settings.json"
            config = {**CONFIG_DEFAULTS, "ui_theme": "night"}
            self.assertTrue(save_config_export(export_path, config))
            self.assertEqual(load_json(export_path, {})["ui_theme"], "night")

            fixed_time = datetime(2026, 7, 22, 12, 30, 0)
            first = backup_config(config, directory, now=fixed_time)
            second = backup_config(config, directory, now=fixed_time)

            self.assertNotEqual(first, second)
            self.assertTrue(first.exists())
            self.assertTrue(second.exists())
            self.assertEqual(load_json(first, {})["config_schema_version"], 1)


if __name__ == "__main__":
    unittest.main()
