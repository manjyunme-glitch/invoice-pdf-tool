from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional


def filter_history_records(
    records: List[Dict[str, Any]],
    type_filter: str = "全部",
    date_filter: str = "全部",
    keyword: str = "",
    now: Optional[datetime] = None,
) -> List[int]:
    filtered_indices: List[int] = []
    normalized_keyword = keyword.strip().lower()
    current_time = now or datetime.now()
    cutoff: Optional[datetime] = None

    if date_filter == "最近7天":
        cutoff = current_time - timedelta(days=7)
    elif date_filter == "最近30天":
        cutoff = current_time - timedelta(days=30)

    for index, record in enumerate(records):
        record_type = str(record.get("type", "整理")).strip()
        if type_filter != "全部" and record_type != type_filter:
            continue

        if cutoff is not None:
            record_time_raw = str(record.get("time", "")).strip()
            try:
                record_time = datetime.strptime(record_time_raw, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
            if record_time < cutoff:
                continue

        if normalized_keyword:
            raw_moves = record.get("moves", [])
            moves = raw_moves if isinstance(raw_moves, list) else []
            raw_reports = record.get("report_files", [])
            reports = raw_reports if isinstance(raw_reports, list) else []
            raw_results = record.get("result_rows", [])
            results = raw_results if isinstance(raw_results, list) else []
            names = [str(move.get("filename", "")) for move in moves[:10] if isinstance(move, dict)]
            report_names = [Path(str(path)).name for path in reports[:10]]
            result_text = " ".join(
                " ".join(
                    str(row.get(key, ""))
                    for key in ("status", "filename", "invoice_number", "pdf_name", "detail")
                )
                for row in results[:20]
                if isinstance(row, dict)
            )
            haystack = " ".join(
                part.lower()
                for part in (
                    str(record.get("time", "")),
                    str(record.get("folder", "")),
                    record_type,
                    " ".join(names),
                    " ".join(report_names),
                    result_text,
                )
                if part
            )
            if normalized_keyword not in haystack:
                continue

        filtered_indices.append(index)

    return filtered_indices


def history_record_can_rerun(record: Dict[str, Any]) -> bool:
    rerun = record.get("rerun")
    if not isinstance(rerun, dict):
        return False
    task_type = str(rerun.get("type", record.get("type", "")))
    if task_type == "整理":
        return bool(str(rerun.get("folder", "")).strip()) and isinstance(
            rerun.get("selected_files", []), list
        )
    if task_type == "筛选":
        return all(
            bool(str(rerun.get(key, "")).strip())
            for key in ("excel_path", "pdf_folder", "output_dir")
        )
    return False
