from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, Optional, Tuple

from ..infra.storage import load_json, save_json


CONFIG_SCHEMA_VERSION = 1
CONFIG_DEFAULTS: Dict[str, Any] = {
    "config_schema_version": CONFIG_SCHEMA_VERSION,
    "help_seen": False,
    "workspace_page": "filter",
    "ui_theme": "day",
    "rule_preset_id": "standard_digital",
    "company_name_index": 2,
    "invoice_number_index": 1,
    "organize_folder": "",
    "excel_path": "",
    "excel_sheet_name": "",
    "pdf_folder": "",
    "selected_invoice_column_name": "",
    "selected_company_column_name": "",
    "row_filter_column_name": "",
    "row_filter_mode": "不过滤",
    "row_filter_values": "",
    "company_exclude_keywords": "",
    "auto_output_by_sheet": True,
    "output_folder": "",
    "invoice_column_aliases": "",
    "company_column_aliases": "",
}

STRING_KEYS = frozenset(
    {
        "rule_preset_id",
        "organize_folder",
        "excel_path",
        "excel_sheet_name",
        "pdf_folder",
        "selected_invoice_column_name",
        "selected_company_column_name",
        "row_filter_column_name",
        "row_filter_values",
        "company_exclude_keywords",
        "output_folder",
        "invoice_column_aliases",
        "company_column_aliases",
    }
)
PATH_KEYS = frozenset({"organize_folder", "excel_path", "pdf_folder", "output_folder"})
ROW_FILTER_MODES = frozenset({"不过滤", "等于任一", "包含任一", "不等于任一", "不包含任一"})
WORKSPACE_PAGES = frozenset({"filter", "organize", "history", "settings"})
UI_THEMES = frozenset({"day", "night"})


class ConfigurationError(ValueError):
    """Raised when an imported configuration cannot be safely interpreted."""


@dataclass(frozen=True)
class ConfigChange:
    key: str
    old_value: Any
    new_value: Any


@dataclass(frozen=True)
class ConfigPlan:
    config: Dict[str, Any]
    changes: Tuple[ConfigChange, ...]
    warnings: Tuple[str, ...]


def _validate_known_value(
    key: str,
    value: Any,
    *,
    preset_ids: Optional[frozenset[str]] = None,
) -> Tuple[bool, Any, str]:
    if key == "config_schema_version":
        if isinstance(value, int) and not isinstance(value, bool) and value == CONFIG_SCHEMA_VERSION:
            return True, value, ""
        return False, CONFIG_SCHEMA_VERSION, f"{key} 必须为 {CONFIG_SCHEMA_VERSION}"
    if key in {"help_seen", "auto_output_by_sheet"}:
        if isinstance(value, bool):
            return True, value, ""
        return False, CONFIG_DEFAULTS[key], f"{key} 必须为布尔值"
    if key in {"company_name_index", "invoice_number_index"}:
        if isinstance(value, int) and not isinstance(value, bool) and 0 <= value <= 10:
            return True, value, ""
        return False, CONFIG_DEFAULTS[key], f"{key} 必须是 0 到 10 的整数"
    if key == "ui_theme":
        normalized = str(value).strip().lower() if isinstance(value, str) else ""
        if normalized in UI_THEMES:
            return True, normalized, ""
        return False, CONFIG_DEFAULTS[key], "ui_theme 只能是 day 或 night"
    if key == "workspace_page":
        normalized = str(value).strip().lower() if isinstance(value, str) else ""
        if normalized in WORKSPACE_PAGES:
            return True, normalized, ""
        return False, CONFIG_DEFAULTS[key], "workspace_page 不是受支持的任务页"
    if key == "row_filter_mode":
        normalized = str(value).strip() if isinstance(value, str) else ""
        if normalized in ROW_FILTER_MODES:
            return True, normalized, ""
        return False, CONFIG_DEFAULTS[key], "row_filter_mode 不是受支持的筛选模式"
    if key in STRING_KEYS:
        if not isinstance(value, str):
            return False, CONFIG_DEFAULTS[key], f"{key} 必须为字符串"
        limit = 32767 if key in PATH_KEYS else 4096
        if len(value) > limit:
            return False, CONFIG_DEFAULTS[key], f"{key} 超过允许长度 {limit}"
        normalized = value.strip() if key not in {"excel_sheet_name"} else value
        if key == "rule_preset_id" and preset_ids and normalized not in preset_ids:
            return False, CONFIG_DEFAULTS[key], f"未知规则预设：{normalized}"
        return True, normalized, ""
    return True, value, ""


def normalize_config(
    raw: Mapping[str, Any],
    *,
    base: Optional[Mapping[str, Any]] = None,
    preset_ids: Optional[Iterable[str]] = None,
    include_defaults: bool = True,
) -> ConfigPlan:
    if not isinstance(raw, Mapping):
        raise ConfigurationError("配置根节点必须是 JSON 对象")
    schema_version = raw.get("config_schema_version", CONFIG_SCHEMA_VERSION)
    if isinstance(schema_version, int) and not isinstance(schema_version, bool) and schema_version > CONFIG_SCHEMA_VERSION:
        raise ConfigurationError(
            f"配置版本 {schema_version} 高于当前支持版本 {CONFIG_SCHEMA_VERSION}，已停止导入"
        )

    normalized: Dict[str, Any] = dict(base or {})
    if include_defaults:
        for key, value in CONFIG_DEFAULTS.items():
            normalized.setdefault(key, value)
    known_preset_ids = frozenset(str(item) for item in preset_ids) if preset_ids is not None else None
    warnings = []
    for key, value in raw.items():
        if key not in CONFIG_DEFAULTS:
            normalized[key] = value
            warnings.append(f"保留未知配置项：{key}")
            continue
        valid, normalized_value, reason = _validate_known_value(
            key,
            value,
            preset_ids=known_preset_ids,
        )
        if valid:
            normalized[key] = normalized_value
        else:
            warnings.append(f"忽略无效配置项 {key}：{reason}")

    normalized["config_schema_version"] = CONFIG_SCHEMA_VERSION
    reference = dict(base or {})
    changes = tuple(
        ConfigChange(key=key, old_value=reference.get(key), new_value=value)
        for key, value in normalized.items()
        if reference.get(key) != value
    )
    return ConfigPlan(config=normalized, changes=changes, warnings=tuple(warnings))


def load_config_plan(
    path: Path,
    current: Mapping[str, Any],
    *,
    preset_ids: Optional[Iterable[str]] = None,
) -> ConfigPlan:
    raw = load_json(path, None)
    if not isinstance(raw, dict):
        raise ConfigurationError("配置文件不是有效的 JSON 对象，未执行导入")
    return normalize_config(raw, base=current, preset_ids=preset_ids, include_defaults=True)


def default_config_plan(current: Mapping[str, Any]) -> ConfigPlan:
    preserved_unknown = {key: value for key, value in current.items() if key not in CONFIG_DEFAULTS}
    reset_config = {**preserved_unknown, **CONFIG_DEFAULTS}
    changes = tuple(
        ConfigChange(key=key, old_value=current.get(key), new_value=value)
        for key, value in reset_config.items()
        if current.get(key) != value
    )
    return ConfigPlan(config=reset_config, changes=changes, warnings=())


def save_config_export(path: Path, config: Mapping[str, Any]) -> bool:
    export_data = dict(config)
    export_data["config_schema_version"] = CONFIG_SCHEMA_VERSION
    return save_json(path, export_data)


def backup_config(
    config: Mapping[str, Any],
    directory: Path,
    *,
    now: Optional[datetime] = None,
) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    timestamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    counter = 0
    while True:
        suffix = "" if counter == 0 else f"_{counter}"
        candidate = directory / f"config.backup.{timestamp}{suffix}.json"
        if not candidate.exists():
            break
        counter += 1
    if not save_json(candidate, dict(config)):
        raise OSError(f"配置备份写入失败：{candidate}")
    return candidate
