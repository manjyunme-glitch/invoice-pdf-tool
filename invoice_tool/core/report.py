from __future__ import annotations

import inspect
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
        result_rows: Optional[List[Dict[str, str]]] = None,
    ) -> Optional[Path]:
        strategy = exporter or ReportExporter.DEFAULT_EXPORTER
        export_method = strategy.export_filter_report
        try:
            parameters = inspect.signature(export_method).parameters.values()
            supports_result_rows = any(
                parameter.name == "result_rows" or parameter.kind == inspect.Parameter.VAR_KEYWORD
                for parameter in parameters
            )
        except (TypeError, ValueError):
            supports_result_rows = False
        if supports_result_rows:
            return export_method(
                output_dir,
                matched,
                not_found,
                excel_col_name,
                result_rows=result_rows,
            )
        return export_method(output_dir, matched, not_found, excel_col_name)
