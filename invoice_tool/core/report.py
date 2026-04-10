from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from .strategies import FilterReportExporterStrategy, OpenpyxlFilterReportExporter


class ReportExporter:
    """筛选报告导出。"""

    DEFAULT_EXPORTER = OpenpyxlFilterReportExporter()

    @staticmethod
    def export_filter_report(
        output_dir: Path,
        matched: List[Dict[str, str]],
        not_found: List[str],
        excel_col_name: str,
        exporter: Optional[FilterReportExporterStrategy] = None,
    ) -> Optional[Path]:
        strategy = exporter or ReportExporter.DEFAULT_EXPORTER
        return strategy.export_filter_report(output_dir, matched, not_found, excel_col_name)
