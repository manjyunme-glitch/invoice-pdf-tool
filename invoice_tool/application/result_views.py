from __future__ import annotations

from typing import List

from ..core.models import FilterResultRow


def filter_filter_result_rows(
    rows: List[FilterResultRow],
    status_filter: str = "全部",
    keyword: str = "",
) -> List[FilterResultRow]:
    normalized_keyword = keyword.strip().lower()
    filtered: List[FilterResultRow] = []
    for row in rows:
        if status_filter != "全部" and row.status != status_filter:
            continue
        if normalized_keyword:
            haystack = " ".join(
                part.lower()
                for part in (row.status, row.invoice_number, row.pdf_name, row.detail)
                if part
            )
            if normalized_keyword not in haystack:
                continue
        filtered.append(row)
    return filtered


def sort_filter_result_rows(
    rows: List[FilterResultRow],
    sort_key: str,
    descending: bool = False,
) -> List[FilterResultRow]:
    field_map = {
        "status": "status",
        "invoice": "invoice_number",
        "pdf": "pdf_name",
        "detail": "detail",
    }
    field_name = field_map.get(sort_key, "invoice_number")
    return sorted(
        rows,
        key=lambda row: getattr(row, field_name, "").lower(),
        reverse=descending,
    )
