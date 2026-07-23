from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from ..core.strategies import (
    FilenameParserStrategy,
    FilterReportExporterStrategy,
    InvoiceColumnResolverStrategy,
)


@dataclass(frozen=True)
class OrganizeExecutionRequest:
    folder: Path
    files: Tuple[str, ...]
    preview_data: Dict[str, Dict[str, Any]]
    recursive: bool = False


@dataclass(frozen=True)
class OrganizePreviewRequest:
    folder: Path
    company_index: int
    recursive: bool
    filename_parser: FilenameParserStrategy


@dataclass(frozen=True)
class WorkbookAnalysisRequest:
    excel_path: Path
    extra_invoice_aliases: Tuple[str, ...]
    extra_company_aliases: Tuple[str, ...]
    selected_sheet_name: str
    selected_invoice_column_name: str
    selected_company_column_name: str


@dataclass(frozen=True)
class FilterPreviewRequest:
    excel_path: Path
    pdf_folder: Path
    output_dir: Path
    invoice_index: int
    recursive: bool
    sheet_name: str
    invoice_column_name: Optional[str]
    company_column_name: Optional[str]
    filter_column_name: Optional[str]
    filter_mode: str
    filter_values: Optional[str]
    company_exclude_keywords: Optional[str]
    extra_aliases: Tuple[str, ...]
    exclude_dirs: Tuple[Path, ...]
    filename_parser: FilenameParserStrategy
    column_resolver: InvoiceColumnResolverStrategy
    active_filter_desc: str
    context_signature: Tuple[str, ...]


@dataclass(frozen=True)
class FilterExecutionRequest:
    excel_path: Path
    pdf_folder: Path
    output_dir: Path
    invoice_index: int
    recursive: bool
    sheet_name: str
    invoice_column_name: Optional[str]
    company_column_name: Optional[str]
    filter_column_name: Optional[str]
    filter_mode: str
    filter_values: Optional[str]
    company_exclude_keywords: Optional[str]
    extra_aliases: Tuple[str, ...]
    exclude_dirs: Tuple[Path, ...]
    filename_parser: FilenameParserStrategy
    column_resolver: InvoiceColumnResolverStrategy
    report_exporter: FilterReportExporterStrategy
    active_filter_desc: str
    rule_preset_id: str = ""
    custom_invoice_aliases: str = ""
