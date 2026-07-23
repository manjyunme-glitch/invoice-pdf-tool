from __future__ import annotations

import time
from concurrent.futures import CancelledError
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..infra.logging_setup import logger
from .file_safety import copy_file_exclusive, fingerprint_file, fingerprint_matches, resolve_contained_path
from .filtering import InvoiceFilter
from .models import (
    FilterPreviewResult,
    FilterResultRow,
    FilterTaskResult,
    OrganizePreviewResult,
    OrganizePreviewRow,
    OrganizeResultRow,
    OrganizeTaskResult,
    PdfScanStats,
)
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
OperationCallback = Callable[[Dict[str, Any]], None]
ReportCallback = Callable[[Path], None]
PauseWaiter = Callable[[], None]


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


def _conflict_invoice_numbers(conflicts: List[str]) -> set[str]:
    return {
        invoice_number
        for invoice_number, _detail in map(_parse_conflict_message, conflicts)
        if invoice_number
    }


class OrganizeService:
    @staticmethod
    def preview(
        folder: Path,
        company_index: int,
        recursive: bool = False,
        filename_parser: Optional[FilenameParserStrategy] = None,
        cancel_requested: Optional[CancelCallback] = None,
    ) -> OrganizePreviewResult:
        pdf_files = InvoiceOrganizer.scan_pdf_files(
            folder,
            recursive,
            cancel_requested=cancel_requested,
        )
        rows: List[OrganizePreviewRow] = []
        selectable_count = 0
        organized_count = 0
        invalid_count = 0
        for relative_file in pdf_files:
            if cancel_requested and cancel_requested():
                raise CancelledError("整理预览已取消")
            relative_name = str(relative_file)
            company, valid = InvoiceOrganizer.parse_filename(
                relative_name,
                company_index,
                filename_parser=filename_parser,
            )
            already_organized = bool(
                valid
                and recursive
                and InvoiceOrganizer.is_already_organized(relative_file, company)
            )
            selectable = valid and not already_organized
            target = "已在目标目录" if already_organized else (company if valid else "-")
            if selectable:
                selectable_count += 1
            elif already_organized:
                organized_count += 1
            else:
                invalid_count += 1
            rows.append(
                OrganizePreviewRow(
                    relative_path=relative_name,
                    company=company,
                    target=target,
                    selectable=selectable,
                    already_organized=already_organized,
                )
            )
        return OrganizePreviewResult(
            rows=rows,
            total_count=len(rows),
            selectable_count=selectable_count,
            organized_count=organized_count,
            invalid_count=invalid_count,
        )

    @staticmethod
    def run(
        folder: Path,
        files: List[str],
        preview_data: Dict[str, Dict],
        progress_callback: Optional[ProgressCallback] = None,
        cancel_requested: Optional[CancelCallback] = None,
        operation_callback: Optional[OperationCallback] = None,
        pause_waiter: Optional[PauseWaiter] = None,
    ) -> OrganizeTaskResult:
        started = time.time()
        moves: List[Dict[str, Any]] = []
        success_count = 0
        fail_count = 0
        skip_count = 0
        cancelled = False
        result_rows: List[OrganizeResultRow] = []
        total = len(files)

        logger.info(f"{'=' * 50}")
        logger.info(f"🚀 开始整理 {total} 个文件")

        for index, filename in enumerate(files):
            if pause_waiter:
                pause_waiter()
            if cancel_requested and cancel_requested():
                logger.warning("⏹ 用户取消了操作")
                cancelled = True
                break

            try:
                preview = preview_data.get(filename)
                if not preview or not preview["valid"]:
                    skip_count += 1
                    result_rows.append(
                        OrganizeResultRow(
                            status="已跳过",
                            filename=filename,
                            detail="预览记录不存在或已标记为不可处理",
                        )
                    )
                    continue
                company = preview["company"]
                raw_source = folder / filename
                if raw_source.is_symlink():
                    raise ValueError(f"不允许处理符号链接：{filename}")
                source = resolve_contained_path(raw_source, folder)
                target_dir = InvoiceOrganizer.resolve_company_target(folder, company)
                if source.parent == target_dir:
                    logger.info(f"⏭ 已在目标目录，跳过：{filename}")
                    skip_count += 1
                    result_rows.append(
                        OrganizeResultRow(
                            status="已跳过",
                            filename=filename,
                            company=company,
                            detail="文件已经位于目标公司目录",
                            source=str(source),
                            target=str(target_dir),
                        )
                    )
                    continue
                fingerprint = fingerprint_file(source)
                target, renamed = InvoiceOrganizer.move_file(
                    source,
                    target_dir,
                    filename,
                    root_dir=folder,
                )
                if renamed:
                    logger.warning(f"⚠️ 重命名：{renamed}")
                move: Dict[str, Any] = {
                    "source": str(source),
                    "target": str(target),
                    "filename": filename,
                    "company": company,
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "operation_root": str(folder.resolve()),
                    "fingerprint": fingerprint,
                }
                if operation_callback:
                    try:
                        operation_callback(dict(move))
                    except Exception as exc:
                        restored, restore_error = InvoiceOrganizer.rollback_single_move(move)
                        if restored:
                            raise OSError("恢复日志写入失败，当前文件已还原到原位置") from exc
                        raise RuntimeError(
                            f"恢复日志写入失败且当前文件无法自动还原：{restore_error}"
                        ) from exc
                moves.append(move)
                logger.info(f"✅ {filename} → {company}/")
                success_count += 1
                result_rows.append(
                    OrganizeResultRow(
                        status="已移动",
                        filename=filename,
                        company=company,
                        detail=(f"目标存在同名文件，已重命名为 {renamed}" if renamed else "已安全移动到目标公司目录"),
                        source=str(source),
                        target=str(target),
                    )
                )
            except PermissionError as exc:
                logger.error(f"❌ {filename}（权限不足）")
                fail_count += 1
                result_rows.append(
                    OrganizeResultRow(
                        status="失败",
                        filename=filename,
                        company=str((preview_data.get(filename) or {}).get("company", "")),
                        detail=f"权限不足或文件正被占用：{exc}",
                        source=str(folder / filename),
                        retryable=True,
                    )
                )
            except OSError as exc:
                logger.error(f"❌ {filename}（{exc}）")
                fail_count += 1
                result_rows.append(
                    OrganizeResultRow(
                        status="失败",
                        filename=filename,
                        company=str((preview_data.get(filename) or {}).get("company", "")),
                        detail=f"文件系统操作失败：{exc}",
                        source=str(folder / filename),
                        retryable=True,
                    )
                )
            except ValueError as exc:
                logger.error(f"❌ {filename}（{exc}）")
                fail_count += 1
                result_rows.append(
                    OrganizeResultRow(
                        status="失败",
                        filename=filename,
                        company=str((preview_data.get(filename) or {}).get("company", "")),
                        detail=str(exc),
                        source=str(folder / filename),
                        retryable=True,
                    )
                )
            finally:
                if progress_callback:
                    progress_callback(index + 1, total)

        elapsed = time.time() - started
        logger.info(f"{'=' * 50}")
        logger.info(
            f"📊 整理完成！成功: {success_count} | 跳过: {skip_count} | "
            f"失败: {fail_count} | 耗时: {elapsed:.1f}s"
        )
        return OrganizeTaskResult(
            moves=moves,
            success_count=success_count,
            fail_count=fail_count,
            skip_count=skip_count,
            elapsed=elapsed,
            cancelled=cancelled,
            result_rows=result_rows,
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
        cancel_requested: Optional[CancelCallback] = None,
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
            cancel_requested=cancel_requested,
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
            cancel_requested=cancel_requested,
        )
        if cancel_requested and cancel_requested():
            raise CancelledError("筛选预览已取消")
        preview = InvoiceFilter.preview_match(invoice_numbers, mapping)
        conflict_invoice_numbers = _conflict_invoice_numbers(conflicts)
        preview_not_found = [
            invoice_number
            for invoice_number in preview["not_found"]
            if invoice_number not in conflict_invoice_numbers
        ]
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
            for invoice_number in preview_not_found
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
            not_found=preview_not_found,
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
        operation_callback: Optional[OperationCallback] = None,
        report_callback: Optional[ReportCallback] = None,
        pause_waiter: Optional[PauseWaiter] = None,
    ) -> FilterTaskResult:
        started = time.time()

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
        conflict_invoice_numbers = _conflict_invoice_numbers(conflicts)

        # Delay the first write until all Excel/PDF inputs have been read successfully.
        output_dir.mkdir(parents=True, exist_ok=True)

        found_count = 0
        skip_count = 0
        copy_fail_count = 0
        target_conflict_count = 0
        not_found: List[str] = []
        moves: List[Dict[str, Any]] = []
        buffer: List[Tuple[str, str]] = []
        cancelled = False
        result_rows: List[FilterResultRow] = []

        total = len(invoice_numbers)
        if progress_callback:
            progress_callback(0, total)

        for index, invoice_number in enumerate(invoice_numbers):
            if pause_waiter:
                pause_waiter()
            if cancel_requested and cancel_requested():
                logger.warning("⏹ 取消筛选")
                cancelled = True
                break

            if invoice_number in mapping:
                relative_pdf = mapping[invoice_number]
                source = pdf_folder / relative_pdf
                target = output_dir / Path(relative_pdf).name
                try:
                    source_fingerprint = fingerprint_file(source)
                    if target.exists():
                        if fingerprint_matches(target, source_fingerprint):
                            buffer.append((f"⏭ {relative_pdf}（已存在且内容一致）\n", "skip"))
                            result_rows.append(
                                FilterResultRow(
                                    status="已跳过",
                                    invoice_number=invoice_number,
                                    pdf_name=Path(relative_pdf).name,
                                    detail=f"导出目录已有内容一致的文件：{target.name}",
                                    path=str(target),
                                )
                            )
                            skip_count += 1
                        else:
                            result_rows.append(
                                FilterResultRow(
                                    status="同名冲突",
                                    invoice_number=invoice_number,
                                    pdf_name=Path(relative_pdf).name,
                                    detail=f"导出目录存在内容不同的同名文件，已保留原文件：{target.name}",
                                    path=str(target),
                                )
                            )
                            target_conflict_count += 1
                    else:
                        try:
                            copy_file_exclusive(source, target, source_fingerprint)
                        except FileExistsError:
                            if fingerprint_matches(target, source_fingerprint):
                                buffer.append((f"⏭ {relative_pdf}（并发创建且内容一致）\n", "skip"))
                                result_rows.append(
                                    FilterResultRow(
                                        status="已跳过",
                                        invoice_number=invoice_number,
                                        pdf_name=Path(relative_pdf).name,
                                        detail=f"导出目录已有内容一致的文件：{target.name}",
                                        path=str(target),
                                    )
                                )
                                skip_count += 1
                                continue
                            result_rows.append(
                                FilterResultRow(
                                    status="同名冲突",
                                    invoice_number=invoice_number,
                                    pdf_name=Path(relative_pdf).name,
                                    detail=f"复制期间出现内容不同的同名文件，已保留原文件：{target.name}",
                                    path=str(target),
                                )
                            )
                            target_conflict_count += 1
                            continue
                        move: Dict[str, Any] = {
                            "source": str(source),
                            "target": str(target),
                            "filename": Path(relative_pdf).name,
                            "invoice_number": invoice_number,
                            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "output_root": str(output_dir.resolve()),
                            "fingerprint": source_fingerprint,
                        }
                        if operation_callback:
                            try:
                                operation_callback(dict(move))
                            except Exception as exc:
                                removed, remove_error = InvoiceOrganizer.delete_recorded_file(move)
                                if removed:
                                    raise OSError("恢复日志写入失败，当前复制文件已安全移除") from exc
                                raise RuntimeError(
                                    f"恢复日志写入失败且当前复制文件无法自动移除：{remove_error}"
                                ) from exc
                        moves.append(move)
                        buffer.append((f"✓ {relative_pdf}\n", "found"))
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
                except (PermissionError, OSError, ValueError) as exc:
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
            elif invoice_number not in conflict_invoice_numbers:
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

        result_rows.extend(_build_conflict_rows(conflicts))
        report_result_rows = [
            {
                "status": row.status,
                "invoice_number": row.invoice_number,
                "pdf_name": row.pdf_name,
                "detail": row.detail,
                "path": row.path,
            }
            for row in result_rows
        ]
        report_path = ReportExporter.export_filter_report(
            output_dir,
            moves,
            not_found,
            column_name,
            exporter=report_exporter,
            result_rows=report_result_rows,
        )
        if report_path and report_callback:
            try:
                report_callback(report_path)
            except Exception as exc:
                report_entry = {
                    "target": str(report_path),
                    "filename": report_path.name,
                    "output_root": str(output_dir.resolve()),
                    "fingerprint": fingerprint_file(report_path),
                }
                removed, remove_error = InvoiceOrganizer.delete_recorded_file(report_entry)
                if removed:
                    raise OSError("报告恢复日志写入失败，未登记的报告已安全移除") from exc
                raise RuntimeError(
                    f"报告恢复日志写入失败且报告无法自动移除：{remove_error}"
                ) from exc
        elapsed = time.time() - started

        logger.info(f"{'=' * 50}")
        logger.info(
            f"📊 筛选完成！匹配: {found_count} | 跳过: {skip_count} | 同名冲突: {target_conflict_count} | "
            f"复制失败: {copy_fail_count} | "
            f"未找到: {len(not_found)} | {elapsed:.1f}s"
        )
        return FilterTaskResult(
            found_count=found_count,
            skip_count=skip_count,
            copy_fail_count=copy_fail_count,
            target_conflict_count=target_conflict_count,
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
