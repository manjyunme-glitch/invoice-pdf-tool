from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence

from ..infra.logging_setup import logger
from ..runtime import PANDAS_SUPPORT, pd
from .filtering import InvoiceFilter
from .models import SheetColumnCandidate, WorkbookAnalysisResult, WorkbookSheetProfile


DEFAULT_COMPANY_COLUMN_NAMES: Sequence[str] = (
    "公司名称",
    "公司",
    "购方名称",
    "购买方名称",
    "销售方名称",
    "销方名称",
    "客户名称",
    "供应商名称",
    "单位名称",
    "企业名称",
    "名称",
)

DEFAULT_COMPANY_EXCLUDE_KEYWORDS: Sequence[str] = (
    "税额",
    "金额",
    "备注",
    "税率",
    "地址",
    "电话",
    "开户",
    "银行",
    "账号",
    "邮箱",
)

COMPANY_VALUE_HINTS: Sequence[str] = (
    "公司",
    "有限",
    "集团",
    "商贸",
    "贸易",
    "科技",
    "电子",
    "工业",
    "实业",
    "医院",
    "学校",
    "事务所",
    "中心",
    "银行",
    "超市",
    "厂",
    "店",
    "局",
    "院",
)


def _normalize_name(value: str) -> str:
    return re.sub(r"\s+", "", str(value).strip()).lower()


def _clean_cell(value: Any) -> str:
    return InvoiceFilter.normalize_excel_text(value)


def _unique_non_empty(values: Iterable[str], limit: int = 3) -> List[str]:
    unique: List[str] = []
    seen = set()
    for value in values:
        normalized = _clean_cell(value)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        unique.append(normalized)
        if len(unique) >= limit:
            break
    return unique


def _column_pairs(dataframe: Any) -> List[tuple[Any, str]]:
    pairs: List[tuple[Any, str]] = []
    for raw_column in dataframe.columns.tolist():
        display_name = str(raw_column).strip()
        if not display_name:
            display_name = str(raw_column)
        pairs.append((raw_column, display_name))
    return pairs


def _column_lookup(dataframe: Any) -> Dict[str, Any]:
    lookup: Dict[str, Any] = {}
    for raw_column, display_name in _column_pairs(dataframe):
        lookup.setdefault(display_name, raw_column)
    return lookup


def _invoice_value_score(samples: Sequence[str]) -> int:
    if not samples:
        return 0
    numeric_like = 0
    length_fit = 0
    chinese_like = 0
    for sample in samples:
        cleaned = _clean_cell(sample)
        if re.fullmatch(r"[A-Za-z0-9\-_/]+", cleaned or ""):
            numeric_like += 1
        if 5 <= len(cleaned) <= 20:
            length_fit += 1
        if re.search(r"[\u4e00-\u9fff]", cleaned):
            chinese_like += 1
    total = max(len(samples), 1)
    return int(30 * numeric_like / total) + int(15 * length_fit / total) - int(12 * chinese_like / total)


def _company_value_score(samples: Sequence[str]) -> int:
    if not samples:
        return 0
    text_like = 0
    hinted = 0
    numeric_like = 0
    for sample in samples:
        cleaned = _clean_cell(sample)
        if not cleaned:
            continue
        if re.search(r"[A-Za-z\u4e00-\u9fff]", cleaned):
            text_like += 1
        if any(hint in cleaned for hint in COMPANY_VALUE_HINTS):
            hinted += 1
        if re.fullmatch(r"[0-9\-_/]+", cleaned):
            numeric_like += 1
    total = max(len(samples), 1)
    return int(24 * text_like / total) + int(26 * hinted / total) - int(18 * numeric_like / total)


def _build_sample_rows(dataframe: Any, key_columns: Sequence[str], row_limit: int = 3, column_limit: int = 6) -> List[Dict[str, str]]:
    column_lookup = _column_lookup(dataframe)
    display_columns: List[str] = []
    for name in key_columns:
        if name and name in column_lookup and name not in display_columns:
            display_columns.append(name)
    for _, name in _column_pairs(dataframe):
        if name and name not in display_columns:
            display_columns.append(name)
        if len(display_columns) >= column_limit:
            break

    sample_rows: List[Dict[str, str]] = []
    raw_columns = [column_lookup[name] for name in display_columns if name in column_lookup]
    preview_df = dataframe.loc[:, raw_columns].head(row_limit)
    for offset, (_, row) in enumerate(preview_df.iterrows(), start=2):
        current: Dict[str, str] = {"行号": str(offset)}
        for column_name, raw_column in zip(display_columns, raw_columns):
            current[column_name] = _clean_cell(row.get(raw_column, ""))
        sample_rows.append(current)
    return sample_rows


def _build_candidate(
    column_name: str,
    samples: Sequence[str],
    *,
    exact_names: Sequence[str],
    fuzzy_keywords: Sequence[str],
    exclude_keywords: Sequence[str],
    value_score_func,
    generic_bonus: int,
    generic_reason: str,
    value_only_threshold: int,
) -> Optional[SheetColumnCandidate]:
    normalized = _normalize_name(column_name)
    score = 0
    reasons: List[str] = []
    has_name_signal = False

    exact_targets = [_normalize_name(item) for item in exact_names if str(item).strip()]
    fuzzy_targets = [_normalize_name(item) for item in fuzzy_keywords if str(item).strip()]

    if normalized in exact_targets:
        exact_name = next((item for item in exact_names if _normalize_name(item) == normalized), column_name)
        score += 120 if normalized != _normalize_name("名称") else 55
        reasons.append(f"精确匹配 {exact_name}")
        has_name_signal = True

    matched_fuzzy = [item for item in fuzzy_targets if item and item in normalized]
    if matched_fuzzy:
        strongest = max(matched_fuzzy, key=len)
        score += 70 if strongest != _normalize_name("名称") else 20
        reasons.append(f"包含关键字 {strongest}")
        has_name_signal = True

    if generic_reason and generic_bonus and generic_reason in column_name:
        score += generic_bonus
        reasons.append(f"包含 {generic_reason}")
        has_name_signal = True

    if any(keyword and keyword in column_name for keyword in exclude_keywords):
        score -= 85
        reasons.append("命中排除关键字")

    value_score = value_score_func(samples)
    if not has_name_signal and value_score < value_only_threshold:
        return None
    if value_score:
        score += value_score
        reasons.append("样本值特征匹配")

    threshold = 30 if generic_reason == "发票" else 22
    if score < threshold:
        return None
    return SheetColumnCandidate(
        column_name=str(column_name).strip(),
        score=score,
        reason="；".join(reasons[:3]),
        sample_values=list(samples[:3]),
    )


def _rank_invoice_candidates(
    dataframe: Any,
    extra_aliases: Optional[List[str]] = None,
) -> List[SheetColumnCandidate]:
    exact_names = list(
        dict.fromkeys(
            [
                "发票号码",
                "发票号",
                "数电发票号码",
                "数电发票号",
                "电子发票号码",
                "电子发票号",
                *(extra_aliases or []),
            ]
        )
    )
    fuzzy_keywords = list(dict.fromkeys(exact_names + ["票号", "开票号码", "开票号"]))
    exclude_keywords = ("备注", "说明", "原发票", "销方发票代码", "购方发票代码")

    candidates: List[SheetColumnCandidate] = []
    for raw_column, column_name in _column_pairs(dataframe):
        samples = _unique_non_empty(dataframe[raw_column].tolist(), limit=5)
        candidate = _build_candidate(
            column_name,
            samples,
            exact_names=exact_names,
            fuzzy_keywords=fuzzy_keywords,
            exclude_keywords=exclude_keywords,
            value_score_func=_invoice_value_score,
            generic_bonus=35,
            generic_reason="发票",
            value_only_threshold=45,
        )
        if candidate is not None:
            candidates.append(candidate)

    return sorted(candidates, key=lambda item: (-item.score, item.column_name))[:5]


def _rank_company_candidates(
    dataframe: Any,
    extra_aliases: Optional[List[str]] = None,
) -> List[SheetColumnCandidate]:
    exact_names = list(dict.fromkeys([*DEFAULT_COMPANY_COLUMN_NAMES, *(extra_aliases or [])]))
    fuzzy_keywords = list(dict.fromkeys(exact_names + ["公司", "客户", "供应商", "购方", "销方", "单位", "企业"]))

    candidates: List[SheetColumnCandidate] = []
    for raw_column, column_name in _column_pairs(dataframe):
        samples = _unique_non_empty(dataframe[raw_column].tolist(), limit=5)
        candidate = _build_candidate(
            column_name,
            samples,
            exact_names=exact_names,
            fuzzy_keywords=fuzzy_keywords,
            exclude_keywords=DEFAULT_COMPANY_EXCLUDE_KEYWORDS,
            value_score_func=_company_value_score,
            generic_bonus=18,
            generic_reason="名称",
            value_only_threshold=50,
        )
        if candidate is not None:
            candidates.append(candidate)

    return sorted(candidates, key=lambda item: (-item.score, item.column_name))[:5]


class WorkbookAnalyzerService:
    @staticmethod
    def analyze(
        excel_path: Path,
        extra_invoice_aliases: Optional[List[str]] = None,
        extra_company_aliases: Optional[List[str]] = None,
        sample_row_limit: int = 3,
    ) -> WorkbookAnalysisResult:
        if not PANDAS_SUPPORT:
            raise ValueError("当前环境未安装 pandas，无法分析 Excel 工作簿")

        try:
            excel_file = pd.ExcelFile(str(excel_path))
        except FileNotFoundError:
            raise FileNotFoundError(f"Excel 文件不存在：{excel_path}")
        except PermissionError:
            raise PermissionError("无法读取 Excel，可能被其他程序占用")
        except Exception as exc:
            raise ValueError(f"Excel 读取失败：{exc}")

        try:
            sheet_names = [str(name).strip() for name in excel_file.sheet_names]
            if not sheet_names:
                raise ValueError("Excel 中未找到任何工作表")

            profiles: List[WorkbookSheetProfile] = []
            usable_sheet_count = 0
            recommended_sheet_name = ""

            for sheet_name in sheet_names:
                try:
                    dataframe = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=object)
                except Exception as exc:
                    profiles.append(
                        WorkbookSheetProfile(
                            sheet_name=sheet_name,
                            issue=f"读取失败：{exc}",
                        )
                    )
                    continue

                columns = [str(column).strip() for column in dataframe.columns.tolist()]
                invoice_candidates = _rank_invoice_candidates(dataframe, extra_aliases=extra_invoice_aliases)
                company_candidates = _rank_company_candidates(dataframe, extra_aliases=extra_company_aliases)
                selected_invoice_column = invoice_candidates[0].column_name if invoice_candidates else ""
                selected_company_column = company_candidates[0].column_name if company_candidates else ""
                usable = bool(invoice_candidates)
                recommended = bool(invoice_candidates and company_candidates)
                if usable:
                    usable_sheet_count += 1
                if recommended and not recommended_sheet_name:
                    recommended_sheet_name = sheet_name

                issue = ""
                if not invoice_candidates:
                    issue = "未识别到发票列"
                elif not company_candidates:
                    issue = "未识别到公司列"

                profiles.append(
                    WorkbookSheetProfile(
                        sheet_name=sheet_name,
                        row_count=int(len(dataframe.index)),
                        column_count=len(columns),
                        columns=columns,
                        invoice_candidates=invoice_candidates,
                        company_candidates=company_candidates,
                        selected_invoice_column=selected_invoice_column,
                        selected_company_column=selected_company_column,
                        sample_rows=_build_sample_rows(
                            dataframe,
                            key_columns=[selected_invoice_column, selected_company_column],
                            row_limit=sample_row_limit,
                        ),
                        recommended=recommended,
                        usable=usable,
                        issue=issue,
                    )
                )

            if not recommended_sheet_name:
                first_usable = next((profile.sheet_name for profile in profiles if profile.usable), "")
                recommended_sheet_name = first_usable or (profiles[0].sheet_name if profiles else "")

            result = WorkbookAnalysisResult(
                workbook_name=excel_path.name,
                sheet_profiles=profiles,
                recommended_sheet_name=recommended_sheet_name,
                total_sheet_count=len(profiles),
                usable_sheet_count=usable_sheet_count,
            )
            logger.info(
                "工作簿分析完成：%s | 工作表 %s 个 | 可用于筛选 %s 个 | 推荐工作表：%s",
                excel_path.name,
                result.total_sheet_count,
                result.usable_sheet_count,
                result.recommended_sheet_name or "-",
            )
            return result
        finally:
            excel_file.close()
