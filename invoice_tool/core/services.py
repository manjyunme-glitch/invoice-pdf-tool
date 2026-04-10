from __future__ import annotations

import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from ..infra.logging_setup import logger
from .filtering import InvoiceFilter
from .models import FilterPreviewResult, FilterResultRow, FilterTaskResult, OrganizeTaskResult, PdfScanStats
from .organizer import InvoiceOrganizer
from .report import ReportExporter
from .strategies import (
    FilenameParserStrategy,
    FilterReportExporterStrategy,
    InvoiceColumnResolverStrategy,
)


ProgressCallback = Callable[[int, int], None]
OutputCallback = Callable[[List[Tuple[str, str]]], None]
CancelCallback = Callable[[], bool]


def _parse_conflict_message(conflict: str) -> Tuple[str, str]:
    prefix = "发票号"
    marker = " 重复:"
    if conflict.startswith(prefix) and marker in conflict:
        invoice_number, detail = conflict[len(prefix):].split(marker, 1)
        return invoice_number.strip(), detail.strip()
    return "", conflict


def _build_conflict_rows(conflicts: List[str]) -> List[FilterResultRow]:
    rows: List[FilterResultRow] = []
    for conflict in conflicts:
        invoice_number, detail = _parse_conflict_message(conflict)
        rows.append(
            FilterResultRow(
                status="重复冲突",
                invoice_number=invoice_number,
                detail=detail,
            )
        )
    return rows


class OrganizeService:
    @staticmethod
    def run(
        folder: Path,
        files: List[str],
        preview_data: Dict[str, Dict],
        progress_callback: Optional[ProgressCallback] = None,
        cancel_requested: Optional[CancelCallback] = None,
    ) -> OrganizeTaskResult:
        started = time.time()
        moves: List[Dict[str, str]] = []
        success_count = 0
        fail_count = 0
        cancelled = False
        total = len(files)

        logger.info(f"{'=' * 50}")
        logger.info(f"🚀 开始整理 {total} 个文件")

        for index, filename in enumerate(files):
            if cancel_requested and cancel_requested():
                logger.warning("⏹ 用户取消了操作")
                cancelled = True
                break

            try:
                preview = preview_data.get(filename)
                if not preview or not preview["valid"]:
                    continue
                company = preview["company"]
                source = folder / filename
                target, renamed = InvoiceOrganizer.move_file(source, folder / company, filename)
                if renamed:
                    logger.warning(f"⚠️ 重命名：{renamed}")
                moves.append(
                    {
                        "source": str(source),
                        "target": str(target),
                        "filename": filename,
                        "company": company,
                        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                logger.info(f"✅ {filename} → {company}/")
                success_count += 1
            except PermissionError:
                logger.error(f"❌ {filename}（权限不足）")
                fail_count += 1
            except OSError as exc:
                logger.error(f"❌ {filename}（{exc}）")
                fail_count += 1
            finally:
                if progress_callback:
                    progress_callback(index + 1, total)

        elapsed = time.time() - started
        logger.info(f"{'=' * 50}")
        logger.info(f"📊 整理完成！成功: {success_count} | 失败: {fail_count} | 耗时: {elapsed:.1f}s")
        return OrganizeTaskResult(
            moves=moves,
            success_count=success_count,
            fail_count=fail_count,
            elapsed=elapsed,
            cancelled=cancelled,
        )


class FilterService:
    @staticmethod
    def preview(
        excel_path: Path,
        pdf_folder: Path,
        invoice_index: int,
        recursive: bool = False,
        sheet_name: Optional[str] = None,
        invoice_column_name: Optional[str] = None,
        company_column_name: Optional[str] = None,
        filter_column_name: Optional[str] = None,
        filter_mode: str = "不过滤",
        filter_values: Optional[str] = None,
        company_exclude_keywords: Optional[str] = None,
        extra_aliases: Optional[List[str]] = None,
        exclude_dirs: Optional[List[Path]] = None,
        filename_parser: Optional[FilenameParserStrategy] = None,
        column_resolver: Optional[InvoiceColumnResolverStrategy] = None,
    ) -> FilterPreviewResult:
        excel_result = InvoiceFilter.read_invoice_records(
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
        invoice_numbers = excel_result["invoice_numbers"]
        column_name = excel_result["invoice_column_name"]
        resolved_sheet_name = excel_result["sheet_name"]
        columns = excel_result["columns"]
        mapping, conflicts, stats_raw = InvoiceFilter.build_pdf_mapping(
            pdf_folder,
            invoice_index,
            recursive,
            exclude_dirs=exclude_dirs,
            filename_parser=filename_parser,
        )
        preview = InvoiceFilter.preview_match(invoice_numbers, mapping)
        result_rows = [
            FilterResultRow(
                status="可匹配",
                invoice_number=item["invoice"],
                pdf_name=Path(item["pdf"]).name,
                detail=item["pdf"],
                path=str(pdf_folder / item["pdf"]),
            )
            for item in preview["matched"]
        ]
        result_rows.extend(
            FilterResultRow(
                status="未匹配",
                invoice_number=invoice_number,
                detail="未找到对应PDF",
            )
            for invoice_number in preview["not_found"]
        )
        result_rows.extend(_build_conflict_rows(conflicts))
        return FilterPreviewResult(
            invoice_numbers=invoice_numbers,
            excel_column_name=column_name,
            sheet_name=resolved_sheet_name,
            columns=columns,
            mapping=mapping,
            conflicts=conflicts,
            matched=preview["matched"],
            not_found=preview["not_found"],
            pdf_stats=PdfScanStats(**stats_raw),
            company_column_name=excel_result["company_column_name"],
            filter_column_name=excel_result["filter_column_name"],
            filter_mode=excel_result["filter_mode"],
            filter_values=excel_result["filter_values"],
            source_row_count=excel_result["source_row_count"],
            filtered_out_count=excel_result["filtered_out_count"],
            result_rows=result_rows,
        )

    @staticmethod
    def run(
        excel_path: Path,
        pdf_folder: Path,
        output_dir: Path,
        invoice_index: int,
        recursive: bool = False,
        sheet_name: Optional[str] = None,
        invoice_column_name: Optional[str] = None,
        company_column_name: Optional[str] = None,
        filter_column_name: Optional[str] = None,
        filter_mode: str = "不过滤",
        filter_values: Optional[str] = None,
        company_exclude_keywords: Optional[str] = None,
        extra_aliases: Optional[List[str]] = None,
        exclude_dirs: Optional[List[Path]] = None,
        filename_parser: Optional[FilenameParserStrategy] = None,
        column_resolver: Optional[InvoiceColumnResolverStrategy] = None,
        report_exporter: Optional[FilterReportExporterStrategy] = None,
        progress_callback: Optional[ProgressCallback] = None,
        output_callback: Optional[OutputCallback] = None,
        cancel_requested: Optional[CancelCallback] = None,
    ) -> FilterTaskResult:
        started = time.time()
        output_dir.mkdir(parents=True, exist_ok=True)

        logger.info(f"{'=' * 50}")
        logger.info("🔍 开始筛选发票...")

        excel_result = InvoiceFilter.read_invoice_records(
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
        invoice_numbers = excel_result["invoice_numbers"]
        column_name = excel_result["invoice_column_name"]
        resolved_sheet_name = excel_result["sheet_name"]
        columns = excel_result["columns"]
        logger.info(f"📋 Excel: {len(invoice_numbers)} 个不重复发票（工作表：{resolved_sheet_name} | 列：{column_name}）")
        logger.info(f"📑 当前工作表列：{', '.join(columns[:10])}" + (f" ... 共{len(columns)}列" if len(columns) > 10 else ""))
        if excel_result["filter_mode"] != "不过滤":
            logger.info(
                "🧪 条件筛选：列=%s | 模式=%s | 值=%s",
                excel_result["filter_column_name"],
                excel_result["filter_mode"],
                ", ".join(excel_result["filter_values"]),
            )
        if company_exclude_keywords:
            logger.info("🚫 公司排除关键字：%s", company_exclude_keywords)
        if excel_result["filtered_out_count"]:
            logger.info("🧹 已按条件过滤掉 %s 行", excel_result["filtered_out_count"])

        mapping, conflicts, stats_raw = InvoiceFilter.build_pdf_mapping(
            pdf_folder,
            invoice_index,
            recursive,
            exclude_dirs=exclude_dirs,
            filename_parser=filename_parser,
        )
        pdf_stats = PdfScanStats(**stats_raw)
        logger.info(
            f"📄 PDF扫描: {pdf_stats.scanned} | 命名有效: {pdf_stats.valid_named} | "
            f"命名异常: {pdf_stats.invalid_named} | 重复冲突: {pdf_stats.duplicates} | 唯一映射: {len(mapping)}"
        )
        for conflict in conflicts:
            logger.warning(f"⚠️ {conflict}")

        found_count = 0
        skip_count = 0
        copy_fail_count = 0
        not_found: List[str] = []
        moves: List[Dict[str, str]] = []
        buffer: List[Tuple[str, str]] = []
        cancelled = False
        result_rows: List[FilterResultRow] = []

        total = len(invoice_numbers)
        if progress_callback:
            progress_callback(0, total)

        for index, invoice_number in enumerate(invoice_numbers):
            if cancel_requested and cancel_requested():
                logger.warning("⏹ 取消筛选")
                cancelled = True
                break

            if invoice_number in mapping:
                relative_pdf = mapping[invoice_number]
                source = pdf_folder / relative_pdf
                target = output_dir / Path(relative_pdf).name
                if not target.exists():
                    try:
                        shutil.copy2(str(source), str(target))
                        buffer.append((f"✓ {relative_pdf}\n", "found"))
                        moves.append(
                            {
                                "source": str(source),
                                "target": str(target),
                                "filename": Path(relative_pdf).name,
                                "invoice_number": invoice_number,
                                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            }
                        )
                        result_rows.append(
                            FilterResultRow(
                                status="已导出",
                                invoice_number=invoice_number,
                                pdf_name=Path(relative_pdf).name,
                                detail=f"已导出到 {target}",
                                path=str(target),
                            )
                        )
                        found_count += 1
                    except (PermissionError, OSError) as exc:
                        buffer.append((f"❌ {relative_pdf}（{exc}）\n", "notfound"))
                        result_rows.append(
                            FilterResultRow(
                                status="复制失败",
                                invoice_number=invoice_number,
                                pdf_name=Path(relative_pdf).name,
                                detail=str(exc),
                                path=str(source),
                            )
                        )
                        copy_fail_count += 1
                else:
                    buffer.append((f"⏭ {relative_pdf}（已存在）\n", "skip"))
                    result_rows.append(
                        FilterResultRow(
                            status="已跳过",
                            invoice_number=invoice_number,
                            pdf_name=Path(relative_pdf).name,
                            detail=f"导出目录已存在同名文件：{target.name}",
                            path=str(target),
                        )
                    )
                    skip_count += 1
            else:
                not_found.append(invoice_number)
                result_rows.append(
                    FilterResultRow(
                        status="未匹配",
                        invoice_number=invoice_number,
                        detail="未找到对应PDF",
                    )
                )

            if progress_callback:
                progress_callback(index + 1, total)
            if len(buffer) >= 50 and output_callback:
                output_callback(buffer.copy())
                buffer.clear()

        if buffer and output_callback:
            output_callback(buffer.copy())

        report_path = ReportExporter.export_filter_report(
            output_dir,
            moves,
            not_found,
            column_name,
            exporter=report_exporter,
        )
        elapsed = time.time() - started

        logger.info(f"{'=' * 50}")
        logger.info(
            f"📊 筛选完成！匹配: {found_count} | 跳过: {skip_count} | 复制失败: {copy_fail_count} | "
            f"未找到: {len(not_found)} | {elapsed:.1f}s"
        )
        result_rows.extend(_build_conflict_rows(conflicts))
        return FilterTaskResult(
            found_count=found_count,
            skip_count=skip_count,
            copy_fail_count=copy_fail_count,
            not_found=not_found,
            moves=moves,
            elapsed=elapsed,
            cancelled=cancelled,
            report_path=report_path,
            pdf_stats=pdf_stats,
            sheet_name=resolved_sheet_name,
            excel_column_name=column_name,
            company_column_name=excel_result["company_column_name"],
            filter_column_name=excel_result["filter_column_name"],
            filter_mode=excel_result["filter_mode"],
            filter_values=excel_result["filter_values"],
            source_row_count=excel_result["source_row_count"],
            filtered_out_count=excel_result["filtered_out_count"],
            columns=columns,
            conflicts=conflicts,
            result_rows=result_rows,
        )
