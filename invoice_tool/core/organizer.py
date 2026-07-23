from __future__ import annotations

import os
import shutil
from concurrent.futures import CancelledError
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..infra.paths import is_relative_to
from .file_safety import (
    UnsafeFileOperationError,
    fingerprint_matches,
    has_valid_fingerprint,
    resolve_contained_path,
)
from .strategies import FilenameParserStrategy, SegmentFilenameParser


class InvoiceOrganizer:
    """发票整理相关纯逻辑。"""

    DEFAULT_FILENAME_PARSER = SegmentFilenameParser()
    INVALID_DIRECTORY_CHARACTERS = frozenset('<>:"/\\|?*')
    WINDOWS_RESERVED_NAMES = frozenset(
        {"CON", "PRN", "AUX", "NUL"}
        | {f"COM{index}" for index in range(1, 10)}
        | {f"LPT{index}" for index in range(1, 10)}
    )

    @staticmethod
    def scan_pdf_files(
        folder: Path,
        recursive: bool = False,
        exclude_dirs: Optional[List[Path]] = None,
        cancel_requested: Optional[Callable[[], bool]] = None,
    ) -> List[Path]:
        folder = Path(folder)
        if not folder.exists():
            raise FileNotFoundError(f"PDF目录不存在：{folder}")
        if not folder.is_dir():
            raise NotADirectoryError(f"PDF路径不是文件夹：{folder}")
        root = folder.resolve()
        excluded = [Path(path).resolve() for path in (exclude_dirs or []) if path]

        def is_excluded(path: Path) -> bool:
            resolved = path.resolve()
            return any(resolved == excluded_dir or is_relative_to(resolved, excluded_dir) for excluded_dir in excluded)

        def is_link_like(path: Path) -> bool:
            if path.is_symlink():
                return True
            is_junction = getattr(path, "is_junction", None)
            return bool(callable(is_junction) and is_junction())

        matches: List[Path] = []
        if cancel_requested and cancel_requested():
            raise CancelledError("PDF 扫描已取消")

        if not recursive:
            for path in folder.iterdir():
                if cancel_requested and cancel_requested():
                    raise CancelledError("PDF 扫描已取消")
                if is_link_like(path) or not path.is_file() or path.suffix.lower() != ".pdf":
                    continue
                matches.append(Path(path.name))
        elif not is_excluded(root):
            for current_root, directory_names, file_names in os.walk(
                root,
                topdown=True,
                followlinks=False,
            ):
                if cancel_requested and cancel_requested():
                    raise CancelledError("PDF 扫描已取消")
                current_path = Path(current_root)
                kept_directories: List[str] = []
                for directory_name in directory_names:
                    directory_path = current_path / directory_name
                    if cancel_requested and cancel_requested():
                        raise CancelledError("PDF 扫描已取消")
                    try:
                        if not is_link_like(directory_path) and not is_excluded(directory_path):
                            kept_directories.append(directory_name)
                    except OSError:
                        continue
                directory_names[:] = kept_directories

                for file_name in file_names:
                    if cancel_requested and cancel_requested():
                        raise CancelledError("PDF 扫描已取消")
                    path = current_path / file_name
                    try:
                        if (
                            path.suffix.lower() == ".pdf"
                            and not is_link_like(path)
                            and path.is_file()
                        ):
                            matches.append(path.relative_to(root))
                    except OSError:
                        continue
        if cancel_requested and cancel_requested():
            raise CancelledError("PDF 扫描已取消")
        return sorted(matches)

    @staticmethod
    def parse_filename(
        filename: str,
        company_index: int,
        filename_parser: Optional[FilenameParserStrategy] = None,
    ) -> Tuple[str, bool]:
        parser = filename_parser or InvoiceOrganizer.DEFAULT_FILENAME_PARSER
        company = parser.parse_segment(filename, company_index)
        if company:
            return company, True
        return "格式不符", False

    @staticmethod
    def resolve_company_target(root: Path, company: str) -> Path:
        name = str(company)
        if not name or name != name.strip() or name in {".", ".."}:
            raise UnsafeFileOperationError(f"公司目录名称不安全：{company!r}")
        if any(character in InvoiceOrganizer.INVALID_DIRECTORY_CHARACTERS for character in name):
            raise UnsafeFileOperationError(f"公司目录名称包含 Windows 非法字符：{company!r}")
        if name.endswith((" ", ".")):
            raise UnsafeFileOperationError(f"公司目录名称不能以空格或句点结尾：{company!r}")
        reserved_stem = name.split(".", 1)[0].upper()
        if reserved_stem in InvoiceOrganizer.WINDOWS_RESERVED_NAMES:
            raise UnsafeFileOperationError(f"公司目录名称是 Windows 保留名称：{company!r}")
        return resolve_contained_path(root / name, root)

    @staticmethod
    def is_already_organized(relative_file: Path, company: str) -> bool:
        parent = Path(relative_file).parent
        return (
            len(parent.parts) == 1
            and parent.name.casefold() == str(company).casefold()
        )

    @staticmethod
    def move_file(
        source: Path,
        target_dir: Path,
        filename: str,
        *,
        root_dir: Optional[Path] = None,
    ) -> Tuple[Path, Optional[str]]:
        if source.is_symlink():
            raise UnsafeFileOperationError(f"不允许移动符号链接：{source.name}")
        if root_dir is not None:
            target_dir = resolve_contained_path(target_dir, root_dir)
        target_dir.mkdir(parents=True, exist_ok=True)
        if root_dir is not None:
            target_dir = resolve_contained_path(target_dir, root_dir)
        target = target_dir / Path(filename).name
        renamed: Optional[str] = None
        if target.exists():
            stem = Path(filename).stem
            suffix = Path(filename).suffix
            timestamp = datetime.now().strftime("%H%M%S")
            counter = 1
            while True:
                counter_suffix = "" if counter == 1 else f"_{counter}"
                new_name = f"{stem}_副本{timestamp}{counter_suffix}{suffix}"
                candidate = target_dir / new_name
                if not candidate.exists():
                    target = candidate
                    renamed = new_name
                    break
                counter += 1
        shutil.move(str(source), str(target))
        return target, renamed

    @staticmethod
    def rollback_single_move(move: Dict[str, Any]) -> Tuple[bool, str]:
        filename = str(move.get("filename", "未知文件"))
        root_value = move.get("operation_root")
        fingerprint = move.get("fingerprint")
        if not root_value or not has_valid_fingerprint(fingerprint):
            return False, f"旧历史缺少安全校验信息，已阻止自动回滚：{filename}"
        try:
            root = Path(str(root_value)).resolve()
            raw_target = Path(str(move["target"]))
            raw_source = Path(str(move["source"]))
            if raw_target.is_symlink():
                return False, f"目标是符号链接，已阻止回滚：{filename}"
            target = resolve_contained_path(raw_target, root)
            source = resolve_contained_path(raw_source, root)
            if not target.exists():
                return False, f"文件已不存在：{filename}"
            if source.exists():
                return False, f"原位置已有同名文件，已阻止覆盖：{filename}"
            if not fingerprint_matches(target, fingerprint):
                return False, f"文件内容已变化，已阻止回滚：{filename}"
            source.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(target), str(source))
            if target.parent != root and target.parent.exists() and not any(target.parent.iterdir()):
                target.parent.rmdir()
            return True, ""
        except (KeyError, TypeError, UnsafeFileOperationError) as exc:
            return False, f"回滚记录不安全：{filename} - {exc}"
        except PermissionError:
            return False, f"权限不足：{filename}"
        except OSError as exc:
            return False, f"操作失败：{filename} - {exc}"

    @staticmethod
    def delete_recorded_file(
        record: Dict[str, Any],
        *,
        path_key: str = "target",
        root_key: str = "output_root",
    ) -> Tuple[bool, str]:
        filename = str(record.get("filename") or Path(str(record.get(path_key, "未知文件"))).name)
        root_value = record.get(root_key)
        fingerprint = record.get("fingerprint")
        if not root_value or not has_valid_fingerprint(fingerprint):
            return False, f"旧历史缺少安全校验信息，已阻止自动删除：{filename}"
        try:
            raw_target = Path(str(record[path_key]))
            if raw_target.is_symlink():
                return False, f"目标是符号链接，已阻止删除：{filename}"
            target = resolve_contained_path(raw_target, Path(str(root_value)))
            if not target.exists():
                return False, f"文件已不存在：{filename}"
            if not fingerprint_matches(target, fingerprint):
                return False, f"文件内容已变化，已阻止删除：{filename}"
            target.unlink()
            return True, ""
        except (KeyError, TypeError, UnsafeFileOperationError) as exc:
            return False, f"删除记录不安全：{filename} - {exc}"
        except PermissionError:
            return False, f"权限不足：{filename}"
        except OSError as exc:
            return False, f"删除失败：{filename} - {exc}"
