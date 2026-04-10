from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple


@dataclass(frozen=True)
class RulePreset:
    preset_id: str
    name: str
    description: str
    company_name_index: int
    invoice_number_index: int
    invoice_column_aliases: Tuple[str, ...] = ()
    filename_separator: str = "_"
    exact_column_names: Tuple[str, ...] = ()
    exclude_keywords: Tuple[str, ...] = ()
    report_style: str = "standard"


DEFAULT_RULE_PRESET_ID = "standard_digital"


RULE_PRESETS: Dict[str, RulePreset] = {
    "standard_digital": RulePreset(
        preset_id="standard_digital",
        name="标准数电发票",
        description="适合 `dzfp_发票号码_公司名称_时间戳.pdf` 这类标准下划线命名。",
        company_name_index=2,
        invoice_number_index=1,
    ),
    "finance_export": RulePreset(
        preset_id="finance_export",
        name="财务导出增强",
        description="保持标准文件名分段规则，同时补充财务系统常见列名，如票号、销项票号、开票号码。",
        company_name_index=2,
        invoice_number_index=1,
        invoice_column_aliases=("票号", "销项票号", "开票号码", "票据号"),
    ),
    "supplier_archive": RulePreset(
        preset_id="supplier_archive",
        name="供应商归档命名",
        description="适合 `类型_公司名称_发票号码_时间` 这类归档命名，发票号位于第3段、公司位于第2段。",
        company_name_index=1,
        invoice_number_index=2,
        invoice_column_aliases=("票号", "开票号码"),
    ),
    "custom": RulePreset(
        preset_id="custom",
        name="手动配置",
        description="保留你手动填写的段位和列别名，不自动覆盖当前设置。",
        company_name_index=2,
        invoice_number_index=1,
    ),
}


def list_rule_presets() -> List[RulePreset]:
    return [RULE_PRESETS[preset_id] for preset_id in ("standard_digital", "finance_export", "supplier_archive", "custom")]


def get_rule_preset(preset_id: str) -> RulePreset:
    return RULE_PRESETS.get(preset_id, RULE_PRESETS[DEFAULT_RULE_PRESET_ID])
