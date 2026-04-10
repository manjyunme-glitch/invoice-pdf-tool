from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional


@dataclass
class SheetColumnCandidate:
    column_name: str
    score: int = 0
    reason: str = ""
    sample_values: List[str] = field(default_factory=list)


@dataclass
class WorkbookSheetProfile:
    sheet_name: str
    row_count: int = 0
    column_count: int = 0
    columns: List[str] = field(default_factory=list)
    invoice_candidates: List[SheetColumnCandidate] = field(default_factory=list)
    company_candidates: List[SheetColumnCandidate] = field(default_factory=list)
    selected_invoice_column: str = ""
    selected_company_column: str = ""
    sample_rows: List[Dict[str, str]] = field(default_factory=list)
    recommended: bool = False
    usable: bool = False
    issue: str = ""


@dataclass
class WorkbookAnalysisResult:
    workbook_name: str = ""
    sheet_profiles: List[WorkbookSheetProfile] = field(default_factory=list)
    recommended_sheet_name: str = ""
    total_sheet_count: int = 0
    usable_sheet_count: int = 0


@dataclass
class FilterResultRow:
    status: str
    invoice_number: str = ""
    pdf_name: str = ""
    detail: str = ""
    path: str = ""


@dataclass
class PdfScanStats:
    scanned: int = 0
    valid_named: int = 0
    invalid_named: int = 0
    duplicates: int = 0


@dataclass
class OrganizeTaskResult:
    moves: List[Dict[str, str]] = field(default_factory=list)
    success_count: int = 0
    fail_count: int = 0
    elapsed: float = 0.0
    cancelled: bool = False


@dataclass
class FilterPreviewResult:
    invoice_numbers: List[str]
    excel_column_name: str
    sheet_name: str
    columns: List[str]
    mapping: Dict[str, str]
    conflicts: List[str]
    matched: List[Dict[str, str]]
    not_found: List[str]
    pdf_stats: PdfScanStats
    company_column_name: str = ""
    filter_column_name: str = ""
    filter_mode: str = ""
    filter_values: List[str] = field(default_factory=list)
    source_row_count: int = 0
    filtered_out_count: int = 0
    result_rows: List[FilterResultRow] = field(default_factory=list)


@dataclass
class FilterTaskResult:
    found_count: int = 0
    skip_count: int = 0
    copy_fail_count: int = 0
    not_found: List[str] = field(default_factory=list)
    moves: List[Dict[str, str]] = field(default_factory=list)
    elapsed: float = 0.0
    cancelled: bool = False
    report_path: Optional[Path] = None
    pdf_stats: PdfScanStats = field(default_factory=PdfScanStats)
    sheet_name: str = ""
    excel_column_name: str = ""
    company_column_name: str = ""
    filter_column_name: str = ""
    filter_mode: str = ""
    filter_values: List[str] = field(default_factory=list)
    source_row_count: int = 0
    filtered_out_count: int = 0
    columns: List[str] = field(default_factory=list)
    conflicts: List[str] = field(default_factory=list)
    result_rows: List[FilterResultRow] = field(default_factory=list)
