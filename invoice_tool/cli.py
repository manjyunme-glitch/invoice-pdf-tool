from __future__ import annotations

import argparse
import ctypes
import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from .core import (
    DEFAULT_RULE_PRESET_ID,
    FilterService,
    InvoiceFilter,
    InvoiceOrganizer,
    OpenpyxlFilterReportExporter,
    SegmentFilenameParser,
    SmartInvoiceColumnResolver,
    get_rule_preset,
    list_rule_presets,
)
from .infra.paths import is_relative_to


def _get_parent_process_name() -> str:
    if os.name != "nt":
        return ""

    try:
        from ctypes import wintypes

        class PROCESS_BASIC_INFORMATION(ctypes.Structure):
            _fields_ = [
                ("Reserved1", ctypes.c_void_p),
                ("PebBaseAddress", ctypes.c_void_p),
                ("Reserved2_0", ctypes.c_void_p),
                ("Reserved2_1", ctypes.c_void_p),
                ("UniqueProcessId", ctypes.c_size_t),
                ("InheritedFromUniqueProcessId", ctypes.c_size_t),
            ]

        kernel32 = ctypes.windll.kernel32
        ntdll = ctypes.windll.ntdll

        pbi = PROCESS_BASIC_INFORMATION()
        return_length = wintypes.ULONG()
        status = ntdll.NtQueryInformationProcess(
            kernel32.GetCurrentProcess(),
            0,
            ctypes.byref(pbi),
            ctypes.sizeof(pbi),
            ctypes.byref(return_length),
        )
        if status != 0:
            return ""

        parent_pid = int(pbi.InheritedFromUniqueProcessId)
        if parent_pid <= 0:
            return ""

        process_handle = kernel32.OpenProcess(0x1000, False, parent_pid)
        if not process_handle:
            return ""

        try:
            size = wintypes.DWORD(32768)
            buffer = ctypes.create_unicode_buffer(size.value)
            if not kernel32.QueryFullProcessImageNameW(process_handle, 0, buffer, ctypes.byref(size)):
                return ""
            return Path(buffer.value).name.lower()
        finally:
            kernel32.CloseHandle(process_handle)
    except Exception:
        return ""


def _should_hold_console(command: Optional[str], exit_code: int, *, parent_name: Optional[str] = None) -> bool:
    if not getattr(sys, "frozen", False):
        return False
    resolved_parent = (parent_name if parent_name is not None else _get_parent_process_name()).lower()
    if resolved_parent != "explorer.exe":
        return False
    return command is None or exit_code != 0


def _hold_console_if_needed(command: Optional[str], exit_code: int) -> None:
    if not _should_hold_console(command, exit_code):
        return

    if command is None:
        print("\n这是命令行版本，双击打开时不会像 GUI 一样停留在界面。")
        print("请在命令提示符或 PowerShell 中带参数运行，例如：")
        print("  发票处理工具CLI.exe presets --json")
        print("  发票处理工具CLI.exe organize --folder D:\\pdf --dry-run")
        print("  发票处理工具CLI.exe filter --excel D:\\sample.xlsx --pdf-folder D:\\pdf --output-folder D:\\out --dry-run")

    try:
        input("\n按回车键退出...")
    except EOFError:
        pass


def _merge_aliases(*alias_groups: Iterable[str]) -> List[str]:
    merged: List[str] = []
    seen = set()
    for group in alias_groups:
        for alias in group:
            normalized = str(alias).strip()
            if normalized and normalized not in seen:
                seen.add(normalized)
                merged.append(normalized)
    return merged


def _make_filename_parser(preset_id: str, separator_override: Optional[str] = None) -> SegmentFilenameParser:
    preset = get_rule_preset(preset_id)
    return SegmentFilenameParser(separator=separator_override or preset.filename_separator)


def _make_column_resolver(preset_id: str, aliases: List[str]) -> SmartInvoiceColumnResolver:
    preset = get_rule_preset(preset_id)
    exact_names = tuple(
        dict.fromkeys(list(InvoiceFilter.EXACT_COL_NAMES) + list(preset.exact_column_names) + aliases)
    )
    exclude_keywords = tuple(
        dict.fromkeys(list(InvoiceFilter.EXCLUDE_KEYWORDS) + list(preset.exclude_keywords))
    )
    return SmartInvoiceColumnResolver(
        exact_column_names=exact_names,
        exclude_keywords=exclude_keywords,
    )


def _make_report_exporter(_: str) -> OpenpyxlFilterReportExporter:
    return OpenpyxlFilterReportExporter()


def _default_company_index(preset_id: str, override: Optional[int]) -> int:
    return override if override is not None else get_rule_preset(preset_id).company_name_index


def _default_invoice_index(preset_id: str, override: Optional[int]) -> int:
    return override if override is not None else get_rule_preset(preset_id).invoice_number_index


def _print_payload(payload: Dict[str, Any], as_json: bool) -> None:
    if as_json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    for key, value in payload.items():
        print(f"{key}: {value}")


def _list_presets_command(args: argparse.Namespace) -> int:
    payload = {
        "presets": [
            {
                "id": preset.preset_id,
                "name": preset.name,
                "description": preset.description,
                "company_name_index": preset.company_name_index,
                "invoice_number_index": preset.invoice_number_index,
            }
            for preset in list_rule_presets()
        ]
    }
    _print_payload(payload, args.json)
    return 0


def _organize_command(args: argparse.Namespace) -> int:
    folder = Path(args.folder).expanduser().resolve()
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"整理目录无效: {folder}")

    company_index = _default_company_index(args.preset, args.company_index)
    filename_parser = _make_filename_parser(args.preset, args.separator)
    pdf_files = InvoiceOrganizer.scan_pdf_files(folder, recursive=args.recursive)
    preview_data: Dict[str, Dict[str, Any]] = {}
    selected_files: List[str] = []
    requested_files = {item.strip() for item in (args.files or []) if item and item.strip()}

    for pdf_file in pdf_files:
        relative_name = str(pdf_file)
        company, valid = InvoiceOrganizer.parse_filename(
            relative_name,
            company_index,
            filename_parser=filename_parser,
        )
        preview_data[relative_name] = {
            "filename": relative_name,
            "company": company,
            "target": company if valid else "-",
            "valid": valid,
        }
        if requested_files:
            if relative_name in requested_files or Path(relative_name).name in requested_files:
                selected_files.append(relative_name)
        elif valid:
            selected_files.append(relative_name)

    if not selected_files:
        raise ValueError("没有可整理的文件，请检查目录、规则或 --files 参数。")

    total_count = len(pdf_files)
    valid_count = sum(1 for item in preview_data.values() if item["valid"])

    if args.dry_run:
        payload = {
            "mode": "organize-dry-run",
            "preset": args.preset,
            "folder": str(folder),
            "scanned": total_count,
            "valid": valid_count,
            "selected": len(selected_files),
            "selected_files": selected_files[:50],
        }
        _print_payload(payload, args.json)
        return 0

    from .core.services import OrganizeService

    result = OrganizeService.run(
        folder=folder,
        files=selected_files,
        preview_data=preview_data,
    )
    payload = {
        "mode": "organize",
        "preset": args.preset,
        "folder": str(folder),
        "scanned": total_count,
        "valid": valid_count,
        "selected": len(selected_files),
        "success_count": result.success_count,
        "fail_count": result.fail_count,
        "cancelled": result.cancelled,
        "elapsed_seconds": round(result.elapsed, 3),
        "moved_count": len(result.moves),
    }
    _print_payload(payload, args.json)
    return 0


def _filter_command(args: argparse.Namespace) -> int:
    excel_path = Path(args.excel).expanduser().resolve()
    pdf_folder = Path(args.pdf_folder).expanduser().resolve()
    output_folder = Path(args.output_folder).expanduser().resolve()

    if not excel_path.exists():
        raise ValueError(f"Excel 文件不存在: {excel_path}")
    if not pdf_folder.exists() or not pdf_folder.is_dir():
        raise ValueError(f"PDF 目录无效: {pdf_folder}")
    if pdf_folder == output_folder:
        raise ValueError("导出目录不能与 PDF 源文件夹相同。")
    if args.recursive and is_relative_to(output_folder, pdf_folder):
        raise ValueError("递归筛选时，导出目录不能位于 PDF 源文件夹内部。")

    invoice_index = _default_invoice_index(args.preset, args.invoice_index)
    preset = get_rule_preset(args.preset)
    cli_aliases = InvoiceFilter.parse_aliases(args.aliases or "")
    aliases = _merge_aliases(preset.invoice_column_aliases, cli_aliases)
    filename_parser = _make_filename_parser(args.preset, args.separator)
    column_resolver = _make_column_resolver(args.preset, aliases)
    report_exporter = _make_report_exporter(args.preset)

    if args.dry_run:
        preview = FilterService.preview(
            excel_path=excel_path,
            pdf_folder=pdf_folder,
            invoice_index=invoice_index,
            recursive=args.recursive,
            sheet_name=args.sheet,
            extra_aliases=aliases,
            exclude_dirs=[output_folder] if args.recursive else None,
            filename_parser=filename_parser,
            column_resolver=column_resolver,
        )
        payload = {
            "mode": "filter-dry-run",
            "preset": args.preset,
            "excel": str(excel_path),
            "pdf_folder": str(pdf_folder),
            "output_folder": str(output_folder),
            "sheet_name": preview.sheet_name,
            "excel_column": preview.excel_column_name,
            "invoice_count": len(preview.invoice_numbers),
            "matched_count": len(preview.matched),
            "not_found_count": len(preview.not_found),
            "pdf_scanned": preview.pdf_stats.scanned,
            "pdf_duplicates": preview.pdf_stats.duplicates,
        }
        _print_payload(payload, args.json)
        return 0

    result = FilterService.run(
        excel_path=excel_path,
        pdf_folder=pdf_folder,
        output_dir=output_folder,
        invoice_index=invoice_index,
        recursive=args.recursive,
        sheet_name=args.sheet,
        extra_aliases=aliases,
        exclude_dirs=[output_folder] if args.recursive else None,
        filename_parser=filename_parser,
        column_resolver=column_resolver,
        report_exporter=report_exporter,
    )
    payload = {
        "mode": "filter",
        "preset": args.preset,
        "excel": str(excel_path),
        "pdf_folder": str(pdf_folder),
        "output_folder": str(output_folder),
        "sheet_name": result.sheet_name,
        "excel_column": result.excel_column_name,
        "found_count": result.found_count,
        "skip_count": result.skip_count,
        "copy_fail_count": result.copy_fail_count,
        "not_found_count": len(result.not_found),
        "cancelled": result.cancelled,
        "elapsed_seconds": round(result.elapsed, 3),
        "report_path": str(result.report_path) if result.report_path else "",
    }
    _print_payload(payload, args.json)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="发票处理工具v5.2.1",
        description="发票工具 v5.2.1：支持 GUI 与命令行批处理。",
    )
    subparsers = parser.add_subparsers(dest="command")

    gui_parser = subparsers.add_parser("gui", help="启动图形界面")
    gui_parser.set_defaults(handler=None)

    presets_parser = subparsers.add_parser("presets", help="列出内置规则预设")
    presets_parser.add_argument("--json", action="store_true", help="以 JSON 输出")
    presets_parser.set_defaults(handler=_list_presets_command)

    organize_parser = subparsers.add_parser("organize", help="命令行整理发票 PDF")
    organize_parser.add_argument("--folder", required=True, help="待整理的 PDF 目录")
    organize_parser.add_argument("--preset", default=DEFAULT_RULE_PRESET_ID, help="规则预设 ID")
    organize_parser.add_argument("--company-index", type=int, help="公司名所在段位")
    organize_parser.add_argument("--recursive", action="store_true", help="递归扫描子目录")
    organize_parser.add_argument("--separator", help="文件名分隔符，默认使用预设值")
    organize_parser.add_argument("--files", nargs="*", help="仅整理指定文件名或相对路径")
    organize_parser.add_argument("--dry-run", action="store_true", help="只预览，不实际移动文件")
    organize_parser.add_argument("--json", action="store_true", help="以 JSON 输出结果")
    organize_parser.set_defaults(handler=_organize_command)

    filter_parser = subparsers.add_parser("filter", help="命令行筛选并导出发票")
    filter_parser.add_argument("--excel", required=True, help="Excel 文件路径")
    filter_parser.add_argument("--pdf-folder", required=True, help="PDF 目录")
    filter_parser.add_argument("--output-folder", required=True, help="导出目录")
    filter_parser.add_argument("--preset", default=DEFAULT_RULE_PRESET_ID, help="规则预设 ID")
    filter_parser.add_argument("--invoice-index", type=int, help="发票号所在段位")
    filter_parser.add_argument("--sheet", help="Excel 工作表名称")
    filter_parser.add_argument("--aliases", help="自定义列别名，逗号分隔")
    filter_parser.add_argument("--recursive", action="store_true", help="递归扫描子目录")
    filter_parser.add_argument("--separator", help="文件名分隔符，默认使用预设值")
    filter_parser.add_argument("--dry-run", action="store_true", help="只预览，不实际导出")
    filter_parser.add_argument("--json", action="store_true", help="以 JSON 输出结果")
    filter_parser.set_defaults(handler=_filter_command)

    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    handler = getattr(args, "handler", None)

    if getattr(args, "command", None) == "gui":
        from .app import run_gui

        run_gui()
        return 0

    if handler is None:
        parser.print_help()
        _hold_console_if_needed(getattr(args, "command", None), 0)
        return 0

    try:
        return int(handler(args) or 0)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        _hold_console_if_needed(getattr(args, "command", None), 1)
        return 1
