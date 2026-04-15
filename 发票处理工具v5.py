#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v5.2.1 GUI 入口：独立于旧版本项目的当前启动文件。"""

from tkinter import messagebox

from invoice_tool.app import main
from invoice_tool.core import (
    DEFAULT_EXACT_COLUMN_NAMES,
    DEFAULT_EXCLUDE_KEYWORDS,
    DEFAULT_RULE_PRESET_ID,
    FilterPreviewResult,
    FilterResultRow,
    FilterReportExporterStrategy,
    FilterService,
    FilterTaskResult,
    FilenameParserStrategy,
    InvoiceFilter,
    InvoiceColumnResolverStrategy,
    InvoiceOrganizer,
    OpenpyxlFilterReportExporter,
    OrganizeService,
    OrganizeTaskResult,
    PdfScanStats,
    ReportExporter,
    RulePreset,
    SegmentFilenameParser,
    SmartInvoiceColumnResolver,
    get_rule_preset,
    list_rule_presets,
)
from invoice_tool.infra import (
    CONFIG_DIR,
    CONFIG_FILE,
    HISTORY_FILE,
    LOG_FILE,
    get_config_dir,
    is_relative_to,
    load_json,
    logger,
    save_json,
)
from invoice_tool.runtime import (
    DND_FILES,
    DND_SUPPORT,
    MODERN_UI,
    OPENPYXL_SUPPORT,
    PANDAS_SUPPORT,
    TkinterDnD,
    openpyxl,
    pd,
    ttkb,
)
from invoice_tool.ui import InvoiceToolApp, TkTextHandler

__all__ = [
    "CONFIG_DIR",
    "CONFIG_FILE",
    "DEFAULT_EXACT_COLUMN_NAMES",
    "DEFAULT_EXCLUDE_KEYWORDS",
    "DEFAULT_RULE_PRESET_ID",
    "DND_FILES",
    "DND_SUPPORT",
    "FilterPreviewResult",
    "FilterResultRow",
    "FilterReportExporterStrategy",
    "FilterService",
    "FilterTaskResult",
    "FilenameParserStrategy",
    "HISTORY_FILE",
    "InvoiceFilter",
    "InvoiceColumnResolverStrategy",
    "InvoiceOrganizer",
    "InvoiceToolApp",
    "LOG_FILE",
    "MODERN_UI",
    "OpenpyxlFilterReportExporter",
    "OPENPYXL_SUPPORT",
    "OrganizeService",
    "OrganizeTaskResult",
    "PANDAS_SUPPORT",
    "PdfScanStats",
    "RulePreset",
    "ReportExporter",
    "SegmentFilenameParser",
    "SmartInvoiceColumnResolver",
    "TkTextHandler",
    "TkinterDnD",
    "get_rule_preset",
    "get_config_dir",
    "is_relative_to",
    "list_rule_presets",
    "load_json",
    "logger",
    "main",
    "messagebox",
    "openpyxl",
    "pd",
    "save_json",
    "ttkb",
]


if __name__ == "__main__":
    raise SystemExit(main())
