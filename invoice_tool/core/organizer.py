from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from ..infra.paths import is_relative_to
from .strategies import FilenameParserStrategy, SegmentFilenameParser


class InvoiceOrganizer:
    """发票整理相关纯逻辑。"""

    DEFAULT_FILENAME_PARSER = SegmentFilenameParser()

    @staticmethod
    def scan_pdf_files(
        folder: Path,
        recursive: bool = False,
        exclude_dirs: Optional[List[Path]] = None,
    ) -> List[Path]:
        excluded = [path.resolve() for path in (exclude_dirs or []) if path]

        def is_excluded(path: Path) -> bool:
            resolved = path.resolve()
            return any(resolved == excluded_dir or is_relative_to(resolved, excluded_dir) for excluded_dir in excluded)

        if recursive:
            return sorted(
                path.relative_to(folder)
                for path in folder.rglob("*.pdf")
                if path.is_file() and not is_excluded(path)
            )
        return sorted(
            Path(path.name)
            for path in folder.iterdir()
            if path.is_file() and path.suffix.lower() == ".pdf"
        )

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
    def move_file(source: Path, target_dir: Path, filename: str) -> Tuple[Path, Optional[str]]:
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / Path(filename).name
        renamed: Optional[str] = None
        if target.exists():
            stem = Path(filename).stem
            suffix = Path(filename).suffix
            timestamp = datetime.now().strftime("%H%M%S")
            new_name = f"{stem}_副本{timestamp}{suffix}"
            target = target_dir / new_name
            renamed = new_name
        shutil.move(str(source), str(target))
        return target, renamed

    @staticmethod
    def rollback_single_move(move: Dict[str, str]) -> Tuple[bool, str]:
        target = Path(move["target"])
        source = Path(move["source"])
        try:
            if not target.exists():
                return False, f"文件已不存在：{move['filename']}"
            source.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(target), str(source))
            if target.parent.exists() and not any(target.parent.iterdir()):
                target.parent.rmdir()
            return True, ""
        except PermissionError:
            return False, f"权限不足：{move['filename']}"
        except OSError as exc:
            return False, f"操作失败：{move['filename']} - {exc}"
