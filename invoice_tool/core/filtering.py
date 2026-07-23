from __future__ import annotations

import math
import re
from concurrent.futures import CancelledError
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..runtime import OPENPYXL_SUPPORT, PANDAS_SUPPORT, XLRD_SUPPORT, pd
from .organizer import InvoiceOrganizer
from .strategies import (
    DEFAULT_EXACT_COLUMN_NAMES,
    DEFAULT_EXCLUDE_KEYWORDS,
    FilenameParserStrategy,
    InvoiceColumnResolverStrategy,
    SegmentFilenameParser,
    SmartInvoiceColumnResolver,
)


class InvoiceFilter:
    """发票筛选相关纯逻辑。"""

    EXACT_COL_NAMES = list(DEFAULT_EXACT_COLUMN_NAMES)
    EXCLUDE_KEYWORDS = list(DEFAULT_EXCLUDE_KEYWORDS)
    DEFAULT_COLUMN_RESOLVER = SmartInvoiceColumnResolver()
    DEFAULT_FILENAME_PARSER = SegmentFilenameParser()

    @staticmethod
    def _decimal_to_plain_text(value: Decimal) -> str:
        text = format(value, "f")
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        return text or "0"

    @staticmethod
    def build_column_lookup(columns: List[Any]) -> Dict[str, Any]:
        lookup: Dict[str, Any] = {}
        for column in columns:
            display_name = str(column).strip()
            if not display_name:
                display_name = str(column)
            lookup.setdefault(display_name, column)
        return lookup

    @staticmethod
    def build_column_position_lookup(columns: List[Any]) -> Dict[str, int]:
        lookup: Dict[str, int] = {}
        for position, column in enumerate(columns):
            display_name = str(column).strip()
            if not display_name:
                display_name = str(column)
            lookup.setdefault(display_name, position)
        return lookup

    @staticmethod
    def normalize_excel_text(value: Any) -> str:
        if value is None:
            return ""
        if hasattr(pd, "isna") and pd.isna(value):
            return ""

        if isinstance(value, bool):
            return "是" if value else "否"

        if isinstance(value, int):
            return str(value)

        if isinstance(value, Decimal):
            if not value.is_finite():
                return ""
            return InvoiceFilter._decimal_to_plain_text(value)

        if isinstance(value, float):
            if not math.isfinite(value):
                return ""
            if value.is_integer():
                return format(value, ".0f")
            text = format(value, ".15f").rstrip("0").rstrip(".")
            return text or "0"

        text = str(value).strip()
        if not text:
            return ""

        scientific_pattern = re.compile(r"^[+-]?\d+(?:\.\d+)?[eE][+-]?\d+$")
        if scientific_pattern.match(text):
            try:
                decimal_value = Decimal(text)
                text = InvoiceFilter._decimal_to_plain_text(decimal_value)
            except (InvalidOperation, ValueError):
                pass

        if text.endswith(".0") and text[:-2].isdigit():
            text = text[:-2]
        return text

    @staticmethod
    def normalize_invoice_number(value: Any) -> str:
        """normalize_excel_text 已处理 float/.0/科学计数法，直接复用即可。"""
        return InvoiceFilter.normalize_excel_text(value)

    @staticmethod
    def parse_aliases(raw_aliases: Optional[str]) -> List[str]:
        if not raw_aliases:
            return []
        return [part.strip() for part in re.split(r"[,，;\n；]+", raw_aliases) if part.strip()]

    @staticmethod
    def filter_tokens(raw_values: Optional[str]) -> List[str]:
        return InvoiceFilter.parse_aliases(raw_values)

    @staticmethod
    def match_filter_condition(cell_value: str, mode: str, tokens: List[str]) -> bool:
        normalized_value = cell_value.strip().lower()
        normalized_tokens = [token.strip().lower() for token in tokens if token.strip()]
        if not normalized_tokens or not mode or mode == "不过滤":
            return True
        if mode == "等于任一":
            return normalized_value in normalized_tokens
        if mode == "包含任一":
            return any(token in normalized_value for token in normalized_tokens)
        if mode == "不等于任一":
            return normalized_value not in normalized_tokens
        if mode == "不包含任一":
            return all(token not in normalized_value for token in normalized_tokens)
        return True

    @staticmethod
    def find_invoice_column(
        columns: List[str],
        extra_aliases: Optional[List[str]] = None,
        column_resolver: Optional[InvoiceColumnResolverStrategy] = None,
    ) -> Optional[str]:
        resolver = column_resolver or InvoiceFilter.DEFAULT_COLUMN_RESOLVER
        return resolver.find_invoice_column(columns, extra_aliases=extra_aliases)

    @staticmethod
    def ensure_excel_support(excel_path: Path) -> None:
        suffix = excel_path.suffix.lower()
        if suffix not in {".xlsx", ".xls"}:
            raise ValueError(f"不支持的 Excel 格式：{suffix or '无扩展名'}（仅支持 .xlsx / .xls）")
        if suffix == ".xlsx" and not OPENPYXL_SUPPORT:
            raise ValueError("当前环境未安装 openpyxl，无法读取 .xlsx 文件")
        if suffix == ".xls" and not XLRD_SUPPORT:
            raise ValueError("当前环境未安装 xlrd，无法读取 .xls 文件；请重新安装完整依赖")

    @staticmethod
    def resolve_sheet_name(sheet_names: List[str], requested: Optional[str]) -> str:
        if not sheet_names:
            raise ValueError("Excel中未找到任何工作表")
        requested_name = "" if requested is None else str(requested)
        if not requested_name:
            return sheet_names[0]
        if requested_name in sheet_names:
            return requested_name
        trimmed = requested_name.strip()
        matches = [name for name in sheet_names if name.strip() == trimmed]
        if len(matches) == 1:
            return matches[0]
        raise ValueError(f"Excel中未找到工作表：{requested_name}")

    @staticmethod
    def list_excel_sheets(excel_path: Path) -> List[str]:
        if not PANDAS_SUPPORT:
            raise ValueError("当前环境未安装 pandas，无法读取 Excel")
        InvoiceFilter.ensure_excel_support(excel_path)
        try:
            excel_file = pd.ExcelFile(str(excel_path))
        except FileNotFoundError:
            raise FileNotFoundError(f"Excel文件不存在：{excel_path}")
        except PermissionError:
            raise PermissionError("无法读取Excel（可能被其他程序占用）")
        except Exception as exc:
            raise ValueError(f"Excel读取失败：{exc}")

        try:
            sheets = [str(name) for name in excel_file.sheet_names]
        finally:
            excel_file.close()

        if not sheets:
            raise ValueError("Excel中未找到任何工作表")
        return sheets

    @staticmethod
    def read_invoice_numbers(
        excel_path: Path,
        sheet_name: Optional[str] = None,
        invoice_column_name: Optional[str] = None,
        company_column_name: Optional[str] = None,
        filter_column_name: Optional[str] = None,
        filter_mode: str = "不过滤",
        filter_values: Optional[str] = None,
        company_exclude_keywords: Optional[str] = None,
        extra_aliases: Optional[List[str]] = None,
        column_resolver: Optional[InvoiceColumnResolverStrategy] = None,
    ) -> Tuple[List[str], str, str, List[str]]:
        result = InvoiceFilter.read_invoice_records(
            excel_path,
            sheet_name=sheet_name,
            invoice_column_name=invoice_column_name,
            company_column_name=company_column_name,
            filter_column_name=filter_column_name,
            filter_mode=filter_mode,
            filter_values=filter_values,
            company_exclude_keywords=company_exclude_keywords,
            extra_aliases=extra_aliases,
            column_resolver=column_resolver,
        )
        return result["invoice_numbers"], result["invoice_column_name"], result["sheet_name"], result["columns"]

    @staticmethod
    def read_invoice_records(
        excel_path: Path,
        sheet_name: Optional[str] = None,
        invoice_column_name: Optional[str] = None,
        company_column_name: Optional[str] = None,
        filter_column_name: Optional[str] = None,
        filter_mode: str = "不过滤",
        filter_values: Optional[str] = None,
        company_exclude_keywords: Optional[str] = None,
        extra_aliases: Optional[List[str]] = None,
        column_resolver: Optional[InvoiceColumnResolverStrategy] = None,
        cancel_requested: Optional[Callable[[], bool]] = None,
    ) -> Dict[str, Any]:
        if not PANDAS_SUPPORT:
            raise ValueError("当前环境未安装 pandas，无法读取 Excel")
        if cancel_requested and cancel_requested():
            raise CancelledError("Excel 读取已取消")
        InvoiceFilter.ensure_excel_support(excel_path)
        try:
            excel_file = pd.ExcelFile(str(excel_path))
        except FileNotFoundError:
            raise FileNotFoundError(f"Excel文件不存在：{excel_path}")
        except PermissionError:
            raise PermissionError("无法读取Excel（可能被其他程序占用）")
        except Exception as exc:
            raise ValueError(f"Excel读取失败：{exc}")

        try:
            sheet_names = [str(name) for name in excel_file.sheet_names]
            if not sheet_names:
                raise ValueError("Excel中未找到任何工作表")
            target_sheet = InvoiceFilter.resolve_sheet_name(sheet_names, sheet_name)
            dataframe = pd.read_excel(excel_file, sheet_name=target_sheet, dtype=object)
        finally:
            excel_file.close()
        if cancel_requested and cancel_requested():
            raise CancelledError("Excel 读取已取消")

        raw_columns = dataframe.columns.tolist()
        columns = [str(column).strip() if str(column).strip() else str(column) for column in raw_columns]
        column_positions = InvoiceFilter.build_column_position_lookup(raw_columns)
        manual_column = str(invoice_column_name or "").strip()
        if manual_column:
            if manual_column not in column_positions:
                raise ValueError(f"工作表“{target_sheet}”中未找到指定的发票列：{manual_column}")
            column_name = manual_column
            invoice_column_position = column_positions[manual_column]
        else:
            column_name = InvoiceFilter.find_invoice_column(
                columns,
                extra_aliases=extra_aliases,
                column_resolver=column_resolver,
            )
            invoice_column_position = column_positions.get(column_name, -1) if column_name else -1
        if column_name is None or invoice_column_position < 0:
            raise ValueError(
                "Excel中未找到发票号码列！\n"
                "支持的列名：发票号码、发票号、数电发票号码"
            )

        company_column_display = str(company_column_name or "").strip()
        company_column_position: Optional[int] = None
        if company_column_display:
            if company_column_display not in column_positions:
                raise ValueError(f"工作表“{target_sheet}”中未找到指定的公司列：{company_column_display}")
            company_column_position = column_positions[company_column_display]

        filter_column_display = str(filter_column_name or "").strip()
        filter_column_position: Optional[int] = None
        if filter_column_display:
            if filter_column_display not in column_positions:
                raise ValueError(f"工作表“{target_sheet}”中未找到指定的条件列：{filter_column_display}")
            filter_column_position = column_positions[filter_column_display]

        filter_token_list = InvoiceFilter.filter_tokens(filter_values)
        company_exclude_list = InvoiceFilter.filter_tokens(company_exclude_keywords)
        if company_exclude_list and company_column_position is None:
            raise ValueError("已设置公司排除关键字，但尚未选择公司列")
        if filter_token_list and filter_mode != "不过滤" and filter_column_position is None:
            raise ValueError("已设置行筛选条件，但尚未选择条件列")
        records: List[Dict[str, str]] = []
        filtered_out_count = 0
        source_row_count = int(len(dataframe.index))

        for row_number, row_values in enumerate(
            dataframe.itertuples(index=False, name=None),
            start=2,
        ):
            if row_number % 250 == 0 and cancel_requested and cancel_requested():
                raise CancelledError("Excel 读取已取消")
            invoice_number = InvoiceFilter.normalize_invoice_number(
                row_values[invoice_column_position]
            )
            if not invoice_number:
                continue

            company_name = (
                InvoiceFilter.normalize_excel_text(row_values[company_column_position])
                if company_column_position is not None
                else ""
            )
            filter_text = (
                InvoiceFilter.normalize_excel_text(row_values[filter_column_position])
                if filter_column_position is not None
                else ""
            )

            if filter_column_position is not None and not InvoiceFilter.match_filter_condition(filter_text, filter_mode, filter_token_list):
                filtered_out_count += 1
                continue

            if company_exclude_list and company_name:
                normalized_company = company_name.lower()
                if any(token.lower() in normalized_company for token in company_exclude_list):
                    filtered_out_count += 1
                    continue

            records.append(
                {
                    "invoice_number": invoice_number,
                    "company_name": company_name,
                    "row_number": str(row_number),
                    "filter_value": filter_text,
                }
            )

        seen = set()
        unique: List[str] = []
        for item in records:
            invoice_number = item["invoice_number"]
            if invoice_number not in seen:
                seen.add(invoice_number)
                unique.append(invoice_number)
        return {
            "invoice_numbers": unique,
            "invoice_column_name": column_name,
            "company_column_name": company_column_display,
            "filter_column_name": filter_column_display,
            "filter_mode": filter_mode if filter_token_list else "不过滤",
            "filter_values": filter_token_list,
            "sheet_name": target_sheet,
            "columns": columns,
            "records": records,
            "source_row_count": source_row_count,
            "filtered_out_count": filtered_out_count,
        }

    @staticmethod
    def build_pdf_mapping(
        pdf_folder: Path,
        invoice_index: int,
        recursive: bool = False,
        exclude_dirs: Optional[List[Path]] = None,
        filename_parser: Optional[FilenameParserStrategy] = None,
        cancel_requested: Optional[Callable[[], bool]] = None,
    ) -> Tuple[Dict[str, str], List[str], Dict[str, int]]:
        mapping: Dict[str, str] = {}
        conflicts: List[str] = []
        conflict_paths: Dict[str, List[str]] = {}
        parser = filename_parser or InvoiceFilter.DEFAULT_FILENAME_PARSER
        stats = {
            "scanned": 0,
            "valid_named": 0,
            "invalid_named": 0,
            "duplicates": 0,
        }
        pdf_files = InvoiceOrganizer.scan_pdf_files(
            pdf_folder,
            recursive,
            exclude_dirs=exclude_dirs,
            cancel_requested=cancel_requested,
        )
        stats["scanned"] = len(pdf_files)

        for pdf_file in pdf_files:
            if cancel_requested and cancel_requested():
                raise CancelledError("PDF 映射已取消")
            invoice_number = InvoiceFilter.normalize_invoice_number(
                parser.parse_segment(str(pdf_file), invoice_index)
            )
            if invoice_number:
                stats["valid_named"] += 1
                if invoice_number in conflict_paths:
                    conflict_paths[invoice_number].append(str(pdf_file))
                    stats["duplicates"] += 1
                    continue
                if invoice_number in mapping:
                    first_path = mapping.pop(invoice_number)
                    conflict_paths[invoice_number] = [first_path, str(pdf_file)]
                    stats["duplicates"] += 1
                    continue
                mapping[invoice_number] = str(pdf_file)
                continue
            stats["invalid_named"] += 1
        conflicts.extend(
            f"发票号 {invoice_number} 重复: {', '.join(paths)}"
            for invoice_number, paths in conflict_paths.items()
        )
        return mapping, conflicts, stats

    @staticmethod
    def preview_match(invoice_numbers: List[str], pdf_mapping: Dict[str, str]) -> Dict[str, Any]:
        matched = [{"invoice": invoice_number, "pdf": pdf_mapping[invoice_number]} for invoice_number in invoice_numbers if invoice_number in pdf_mapping]
        not_found = [invoice_number for invoice_number in invoice_numbers if invoice_number not in pdf_mapping]
        return {"matched": matched, "not_found": not_found}
