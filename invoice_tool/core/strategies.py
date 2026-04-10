from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Protocol, Sequence, Tuple

from ..infra.logging_setup import logger
from ..runtime import (
    Alignment,
    Border,
    Font,
    OPENPYXL_SUPPORT,
    PatternFill,
    Side,
    openpyxl,
)


DEFAULT_EXACT_COLUMN_NAMES: Tuple[str, ...] = (
    "发票号码",
    "发票号",
    "数电发票号码",
    "数电发票号",
    "电子发票号码",
    "电子发票号",
)

DEFAULT_EXCLUDE_KEYWORDS: Tuple[str, ...] = ("备注", "说明", "原发票")


class FilenameParserStrategy(Protocol):
    def split_parts(self, filename: str) -> List[str]:
        ...

    def parse_segment(self, filename: str, index: int) -> Optional[str]:
        ...


@dataclass(frozen=True)
class SegmentFilenameParser:
    separator: str = "_"
    use_stem: bool = True

    def split_parts(self, filename: str) -> List[str]:
        raw_name = Path(filename).stem if self.use_stem else Path(filename).name
        return [part.strip() for part in raw_name.split(self.separator)]

    def parse_segment(self, filename: str, index: int) -> Optional[str]:
        parts = self.split_parts(filename)
        if len(parts) > index:
            value = parts[index].strip()
            if value:
                return value
        return None


class InvoiceColumnResolverStrategy(Protocol):
    def find_invoice_column(
        self,
        columns: Sequence[str],
        extra_aliases: Optional[List[str]] = None,
    ) -> Optional[str]:
        ...


@dataclass(frozen=True)
class SmartInvoiceColumnResolver:
    exact_column_names: Tuple[str, ...] = DEFAULT_EXACT_COLUMN_NAMES
    exclude_keywords: Tuple[str, ...] = DEFAULT_EXCLUDE_KEYWORDS
    fallback_keyword: str = "发票号"
    normalize_spaces: bool = True

    def _normalize(self, value: str) -> str:
        text = str(value).strip()
        return text.replace(" ", "") if self.normalize_spaces else text

    def find_invoice_column(
        self,
        columns: Sequence[str],
        extra_aliases: Optional[List[str]] = None,
    ) -> Optional[str]:
        column_names = [str(column).strip() for column in columns]
        normalized_map = {self._normalize(column): column for column in column_names}

        exact_targets: List[str] = []
        seen_targets = set()
        for target in list(self.exact_column_names) + list(extra_aliases or []):
            normalized = self._normalize(str(target))
            if normalized and normalized not in seen_targets:
                seen_targets.add(normalized)
                exact_targets.append(normalized)

        for target in exact_targets:
            if target in normalized_map:
                return normalized_map[target]

        fuzzy_targets = [target for target in exact_targets if len(target) >= 2]
        for column in column_names:
            if any(keyword in column for keyword in self.exclude_keywords):
                continue
            if self.fallback_keyword and self.fallback_keyword in column:
                return column
            normalized_column = self._normalize(column)
            if any(alias in normalized_column for alias in fuzzy_targets):
                return column
        return None


class FilterReportExporterStrategy(Protocol):
    def export_filter_report(
        self,
        output_dir: Path,
        matched: List[Dict[str, str]],
        not_found: List[str],
        excel_col_name: str,
    ) -> Optional[Path]:
        ...


@dataclass(frozen=True)
class OpenpyxlFilterReportExporter:
    success_sheet_name: str = "已成功导出"
    missing_sheet_name: str = "缺失清单"
    summary_sheet_name: str = "汇总"
    report_prefix: str = "筛选结果报告"

    def export_filter_report(
        self,
        output_dir: Path,
        matched: List[Dict[str, str]],
        not_found: List[str],
        excel_col_name: str,
    ) -> Optional[Path]:
        if not OPENPYXL_SUPPORT:
            logger.warning("未安装 openpyxl，无法生成筛选报告")
            return None

        workbook = openpyxl.Workbook()

        ws_success = workbook.active
        ws_success.title = self.success_sheet_name

        header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        headers_success = ["序号", "发票号码", "PDF文件名", "导出时间"]
        for column_index, header in enumerate(headers_success, 1):
            cell = ws_success.cell(row=1, column=column_index, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

        for row_index, item in enumerate(matched, 2):
            ws_success.cell(row=row_index, column=1, value=row_index - 1).border = thin_border
            ws_success.cell(
                row=row_index,
                column=2,
                value=item.get("invoice_number", item.get("invoice", "")),
            ).border = thin_border
            ws_success.cell(
                row=row_index,
                column=3,
                value=item.get("filename", item.get("pdf", "")),
            ).border = thin_border
            ws_success.cell(row=row_index, column=4, value=item.get("time", "")).border = thin_border

        ws_success.column_dimensions["A"].width = 8
        ws_success.column_dimensions["B"].width = 28
        ws_success.column_dimensions["C"].width = 55
        ws_success.column_dimensions["D"].width = 22

        ws_missing = workbook.create_sheet(self.missing_sheet_name)
        headers_missing = ["序号", "发票号码", "状态"]
        red_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")

        for column_index, header in enumerate(headers_missing, 1):
            cell = ws_missing.cell(row=1, column=column_index, value=header)
            cell.font = header_font
            cell.fill = red_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

        for row_index, invoice_number in enumerate(not_found, 2):
            ws_missing.cell(row=row_index, column=1, value=row_index - 1).border = thin_border
            ws_missing.cell(row=row_index, column=2, value=invoice_number).border = thin_border
            ws_missing.cell(row=row_index, column=3, value="未找到对应PDF").border = thin_border

        ws_missing.column_dimensions["A"].width = 8
        ws_missing.column_dimensions["B"].width = 28
        ws_missing.column_dimensions["C"].width = 22

        ws_summary = workbook.create_sheet(self.summary_sheet_name)
        summary_rows = [
            ("报告生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("Excel发票号码列", excel_col_name),
            ("总发票数", len(matched) + len(not_found)),
            ("成功导出", len(matched)),
            ("未找到", len(not_found)),
            ("匹配率", f"{len(matched) / max(len(matched) + len(not_found), 1) * 100:.1f}%"),
        ]
        for row_index, (key, value) in enumerate(summary_rows, 1):
            ws_summary.cell(row=row_index, column=1, value=key).font = Font(name="微软雅黑", bold=True)
            ws_summary.cell(row=row_index, column=2, value=value)
        ws_summary.column_dimensions["A"].width = 20
        ws_summary.column_dimensions["B"].width = 30

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = output_dir / f"{self.report_prefix}_{timestamp}.xlsx"
        workbook.save(str(report_path))
        logger.info(f"📊 筛选报告已导出：{report_path.name}")
        return report_path
