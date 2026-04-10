#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票处理工具箱 v5.1

当前版本聚焦于：
- 发票整理
- 多 Sheet 发票筛选
- 条件筛选与公司排除
- 白天 / 黑夜双主题
- 更稳定的 GUI 打包交付
"""

from __future__ import annotations

import logging
import platform
import re
import subprocess
import threading
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple
from ..core.filtering import InvoiceFilter
from ..core.models import FilterResultRow, WorkbookAnalysisResult, WorkbookSheetProfile
from ..core.organizer import InvoiceOrganizer
from ..core.presets import DEFAULT_RULE_PRESET_ID, get_rule_preset, list_rule_presets
from ..core.services import FilterService, OrganizeService
from ..core.strategies import OpenpyxlFilterReportExporter, SegmentFilenameParser, SmartInvoiceColumnResolver
from ..core.workbook import WorkbookAnalyzerService
from ..infra.logging_setup import logger
from ..infra.paths import CONFIG_DIR, CONFIG_FILE, HISTORY_FILE, LOG_FILE, is_relative_to
from ..infra.storage import load_json, save_json
from ..runtime import DND_FILES, DND_SUPPORT, MODERN_UI, OPENPYXL_SUPPORT, PANDAS_SUPPORT, ttkb
from .logging_handler import RecentErrorHandler, TkTextHandler


FILTER_RESULT_STATUS_OPTIONS = ("全部", "可匹配", "未匹配", "重复冲突", "已导出", "已跳过", "复制失败")
HISTORY_TYPE_OPTIONS = ("全部", "整理", "筛选")
HISTORY_DATE_OPTIONS = ("全部", "最近7天", "最近30天")
FILTER_RULE_MODE_OPTIONS = ("不过滤", "等于任一", "包含任一", "不等于任一", "不包含任一")
UI_THEME_OPTIONS = ("day", "night")
UI_THEME_LABELS = {"day": "白天", "night": "黑夜"}
APP_VERSION = "v5.1"
APP_TITLE = f"发票处理工具箱 {APP_VERSION}"

UI_THEME_PRESETS: Dict[str, Dict[str, Any]] = {
    "day": {
        "bootstrap_theme": "flatly",
        "root_bg": "#EEF3F8",
        "surface": "#FFFFFF",
        "surface_alt": "#F6F9FC",
        "surface_soft": "#E8EEF5",
        "title_bg": "#103B66",
        "title_fg": "#F8FBFF",
        "title_muted": "#C7D8EA",
        "title_badge_bg": "#1F5A96",
        "title_badge_fg": "#EFF6FF",
        "text": "#16324A",
        "muted": "#66788A",
        "border": "#D4DEE8",
        "entry_bg": "#FFFFFF",
        "entry_fg": "#18324A",
        "button_bg": "#DDE6EF",
        "button_fg": "#18324A",
        "button_hover": "#CFD9E4",
        "primary": "#1C63D5",
        "primary_hover": "#154FB1",
        "success": "#20825B",
        "success_hover": "#176847",
        "warning": "#D38A20",
        "warning_hover": "#B27014",
        "danger": "#D14C5D",
        "danger_hover": "#B93A4E",
        "secondary": "#6D7F92",
        "secondary_hover": "#596B7E",
        "log_bg": "#0F172A",
        "log_fg": "#E2E8F0",
        "log_drawer_bg": "#DCE5EF",
        "status_bg": "#E3EBF4",
        "status_fg": "#35516A",
        "tree_even": "#F8FBFD",
        "tree_odd": "#FFFFFF",
        "tree_selected": "#1D4ED8",
        "detail_bg": "#F7FAFD",
        "detail_fg": "#41566B",
        "card_palette": [
            ("#E0F2FE", "#0C4A6E"),
            ("#E8F8EE", "#14532D"),
            ("#FFF4DE", "#9A3412"),
            ("#FDE7F3", "#9D174D"),
            ("#EEE8FF", "#5B21B6"),
            ("#E9EFF5", "#334155"),
        ],
    },
    "night": {
        "bootstrap_theme": "darkly",
        "root_bg": "#08111F",
        "surface": "#0F1B2D",
        "surface_alt": "#142235",
        "surface_soft": "#18293F",
        "title_bg": "#07101D",
        "title_fg": "#F8FAFC",
        "title_muted": "#91A4BB",
        "title_badge_bg": "#18314F",
        "title_badge_fg": "#DDEAFE",
        "text": "#E2E8F0",
        "muted": "#94A3B8",
        "border": "#24364A",
        "entry_bg": "#0B1626",
        "entry_fg": "#F8FAFC",
        "button_bg": "#1C2B3C",
        "button_fg": "#E2E8F0",
        "button_hover": "#24364A",
        "primary": "#3B82F6",
        "primary_hover": "#2563EB",
        "success": "#1FA971",
        "success_hover": "#19895C",
        "warning": "#C28724",
        "warning_hover": "#A56E17",
        "danger": "#CE506C",
        "danger_hover": "#B73D57",
        "secondary": "#56687E",
        "secondary_hover": "#64778E",
        "log_bg": "#020617",
        "log_fg": "#D8E1EC",
        "log_drawer_bg": "#142235",
        "status_bg": "#0E1A2B",
        "status_fg": "#B8C7D9",
        "tree_even": "#122033",
        "tree_odd": "#0F1B2D",
        "tree_selected": "#2563EB",
        "detail_bg": "#122033",
        "detail_fg": "#D3DEEA",
        "card_palette": [
            ("#133A56", "#D9F1FF"),
            ("#163A2B", "#DDFBEA"),
            ("#4A3012", "#FFF0D5"),
            ("#4A1935", "#FFE2F1"),
            ("#2F1E5C", "#E8DDFF"),
            ("#223042", "#E2E8F0"),
        ],
    },
}


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
    return sorted(rows, key=lambda row: getattr(row, field_name, "").lower(), reverse=descending)


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
            names = [str(move.get("filename", "")) for move in record.get("moves", [])[:10]]
            report_names = [Path(path).name for path in record.get("report_files", [])[:10]]
            haystack = " ".join(
                part.lower()
                for part in (
                    str(record.get("time", "")),
                    str(record.get("folder", "")),
                    record_type,
                    " ".join(names),
                    " ".join(report_names),
                )
                if part
            )
            if normalized_keyword not in haystack:
                continue

        filtered_indices.append(index)

    return filtered_indices


# ==================== GUI 主应用 ====================

class InvoiceToolApp:
    """发票处理工具箱 v5.1"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self._apply_initial_window_geometry()
        self._default_widget_colors = self._capture_default_widget_colors()

        # 加载配置/历史
        self.config: Dict[str, Any] = self._load_json(CONFIG_FILE, {})
        self.all_history: List[Dict] = self._load_json(HISTORY_FILE, [])
        self._auto_clean_old_history()
        self.rule_presets = list_rule_presets()
        self._preset_by_id = {preset.preset_id: preset for preset in self.rule_presets}
        saved_ui_theme = str(self.config.get("ui_theme", "day")).strip().lower()
        if saved_ui_theme not in UI_THEME_PRESETS:
            saved_ui_theme = "day"
        self.ui_theme = tk.StringVar(value=saved_ui_theme)
        self.ui_theme_label = tk.StringVar(value=self._theme_label(saved_ui_theme))
        self.palette = UI_THEME_PRESETS[saved_ui_theme]

        # 线程控制
        self.is_running = False
        self._lock = threading.Lock()
        self._cancel_flag = threading.Event()
        self._start_time: Optional[float] = None

        # ─── 整理变量 ───
        self.organize_folder_path = tk.StringVar()
        self.file_check_vars: Dict[str, tk.BooleanVar] = {}
        self.preview_data: Dict[str, Dict] = {}
        self.current_session_history: List[Dict] = []
        self.organize_recursive = tk.BooleanVar(value=False)

        # ─── 筛选变量 ───
        self.excel_path = tk.StringVar()
        self.excel_sheet_name = tk.StringVar(value=self.config.get("excel_sheet_name", ""))
        self.selected_invoice_column_name = tk.StringVar(value=self.config.get("selected_invoice_column_name", ""))
        self.selected_company_column_name = tk.StringVar(value=self.config.get("selected_company_column_name", ""))
        self.row_filter_column_name = tk.StringVar(value=self.config.get("row_filter_column_name", ""))
        self.row_filter_mode = tk.StringVar(value=self.config.get("row_filter_mode", "不过滤"))
        self.row_filter_values = tk.StringVar(value=self.config.get("row_filter_values", ""))
        self.company_exclude_keywords = tk.StringVar(value=self.config.get("company_exclude_keywords", ""))
        self._active_filter_context: Tuple[str, str] = (
            self.excel_path.get().strip(),
            self.excel_sheet_name.get().strip(),
        )
        self.pdf_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.manual_output_folder = tk.StringVar(value=self.config.get("output_folder", ""))
        self.auto_output_by_sheet = tk.BooleanVar(value=self.config.get("auto_output_by_sheet", True))
        self.filter_recursive = tk.BooleanVar(value=False)
        self.filter_result_status = tk.StringVar(value="全部")
        self.filter_result_keyword = tk.StringVar()
        self.filter_result_rows: List[FilterResultRow] = []
        self.filter_result_sort_key = "invoice"
        self.filter_result_sort_desc = False
        self.filter_result_selection: Dict[str, FilterResultRow] = {}
        self.filter_missing_invoices: List[str] = []
        self.filter_summary_title = tk.StringVar(value="等待预览或筛选")
        self.filter_summary_subtitle = tk.StringVar(value="先选择 Excel、PDF 和导出目录，然后执行预览或筛选。")
        self.filter_detail_var = tk.StringVar(value="提示：结果将显示在下方表格中，可按状态过滤或搜索发票号。")
        self.filter_metric_labels = {
            "metric1": tk.StringVar(value="Excel发票"),
            "metric2": tk.StringVar(value="命中结果"),
            "metric3": tk.StringVar(value="未匹配"),
            "metric4": tk.StringVar(value="异常/冲突"),
            "metric5": tk.StringVar(value="PDF扫描"),
            "metric6": tk.StringVar(value="其他状态"),
        }
        self.filter_metric_values = {
            "metric1": tk.StringVar(value="0"),
            "metric2": tk.StringVar(value="0"),
            "metric3": tk.StringVar(value="0"),
            "metric4": tk.StringVar(value="0"),
            "metric5": tk.StringVar(value="0"),
            "metric6": tk.StringVar(value="0"),
        }
        self.workbook_analysis_summary_var = tk.StringVar(value="打开 Excel 后，会自动分析每个工作表的发票列和公司列候选。")
        self.workbook_sheet_overview_var = tk.StringVar(value="先选择 Excel 文件，再从左侧查看每个 sheet 的识别结果。")
        self.workbook_sheet_sample_var = tk.StringVar(value="样本预览会显示当前工作表前几行数据，便于确认列是否正确。")
        self.workbook_analysis_result: Optional[WorkbookAnalysisResult] = None
        self.workbook_profiles: Dict[str, WorkbookSheetProfile] = {}
        self.workbook_tree_selection: Dict[str, str] = {}
        self.history_type_filter = tk.StringVar(value="全部")
        self.history_date_filter = tk.StringVar(value="全部")
        self.history_keyword = tk.StringVar()
        self.filtered_history_indices: List[int] = []
        self.history_summary_var = tk.StringVar(value="显示 0 / 0 条历史记录")
        self.recent_errors: List[Dict[str, str]] = []
        self.recent_error_limit = 20
        self.recent_error_summary_var = tk.StringVar(value="最近错误 0 条")
        self.recent_error_detail_var = tk.StringVar(value="运行过程中出现的错误会显示在这里，方便快速排查。")

        # ─── 设置变量 ───
        saved_preset_id = self.config.get("rule_preset_id", DEFAULT_RULE_PRESET_ID)
        if saved_preset_id not in self._preset_by_id:
            saved_preset_id = DEFAULT_RULE_PRESET_ID
        self.rule_preset_id = tk.StringVar(value=saved_preset_id)
        self.rule_preset_name = tk.StringVar(value=self._preset_by_id[saved_preset_id].name)
        self.rule_preset_desc = tk.StringVar(value=self._preset_by_id[saved_preset_id].description)
        self.company_name_index = tk.IntVar(value=self.config.get("company_name_index", 2))
        self.invoice_number_index = tk.IntVar(value=self.config.get("invoice_number_index", 1))
        self.invoice_column_aliases = tk.StringVar(value=self.config.get("invoice_column_aliases", ""))
        self.company_column_aliases = tk.StringVar(value=self.config.get("company_column_aliases", ""))

        # 日志抽屉状态
        self._log_visible = tk.BooleanVar(value=False)
        self._gui_log_handler: Optional[TkTextHandler] = None
        self._recent_error_handler: Optional[RecentErrorHandler] = None

        # 构建界面
        self._build_ui()
        self._setup_drag_and_drop()
        self._restore_paths()
        if PANDAS_SUPPORT:
            self._refresh_excel_sheets(silent=True)

        # 关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        logger.info("应用启动 %s", APP_VERSION)

    # ==================== JSON / 配置 ====================

    @staticmethod
    def _load_json(path: Path, default: Any) -> Any:
        return load_json(path, default)

    @staticmethod
    def _save_json(path: Path, data: Any) -> None:
        save_json(path, data)

    def _capture_default_widget_colors(self) -> Dict[str, Dict[str, str]]:
        probes = {
            "Frame": tk.Frame(self.root),
            "Label": tk.Label(self.root),
            "Button": tk.Button(self.root),
            "Entry": tk.Entry(self.root),
            "Checkbutton": tk.Checkbutton(self.root),
            "Listbox": tk.Listbox(self.root),
            "LabelFrame": tk.LabelFrame(self.root),
            "Spinbox": tk.Spinbox(self.root),
        }
        defaults: Dict[str, Dict[str, str]] = {}
        for name, widget in probes.items():
            snapshot: Dict[str, str] = {}
            for option in ("bg", "fg", "activebackground", "activeforeground", "insertbackground"):
                try:
                    snapshot[option] = str(widget.cget(option))
                except tk.TclError:
                    continue
            defaults[name] = snapshot
            widget.destroy()
        return defaults

    def _apply_initial_window_geometry(self) -> None:
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = min(1240, max(980, screen_width - 110))
        height = min(780, max(690, screen_height - 140))
        min_width = min(1080, max(920, screen_width - 140))
        min_height = min(700, max(620, screen_height - 180))
        pos_x = max((screen_width - width) // 2, 24)
        pos_y = max((screen_height - height) // 2 - 24, 24)
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        self.root.minsize(min_width, min_height)

    def _theme_label(self, theme_id: Optional[str] = None) -> str:
        resolved = theme_id or self.ui_theme.get()
        return UI_THEME_LABELS.get(resolved, "白天")

    def _set_ui_theme(self, theme_id: str) -> None:
        theme_id = str(theme_id).strip().lower()
        if theme_id not in UI_THEME_PRESETS:
            return
        if theme_id == self.ui_theme.get():
            return
        self.ui_theme.set(theme_id)
        self.ui_theme_label.set(self._theme_label(theme_id))
        self.palette = UI_THEME_PRESETS[theme_id]
        self._save_config()
        self._rebuild_ui()

    def _toggle_ui_theme(self) -> None:
        next_theme = "night" if self.ui_theme.get() == "day" else "day"
        self._set_ui_theme(next_theme)

    def _on_ui_theme_change(self, event=None) -> None:
        selected_label = self.ui_theme_label.get().strip()
        theme_id = next((key for key, label in UI_THEME_LABELS.items() if label == selected_label), self.ui_theme.get())
        self._set_ui_theme(theme_id)

    def _rebuild_ui(self) -> None:
        selected_tab = 0
        if hasattr(self, "notebook"):
            try:
                selected_tab = self.notebook.index(self.notebook.select())
            except Exception:
                selected_tab = 0

        if self._gui_log_handler is not None:
            logger.removeHandler(self._gui_log_handler)
            self._gui_log_handler = None
        if self._recent_error_handler is not None:
            logger.removeHandler(self._recent_error_handler)
            self._recent_error_handler = None

        for child in self.root.winfo_children():
            child.destroy()

        self._build_ui()
        self._setup_drag_and_drop()
        if PANDAS_SUPPORT:
            self._refresh_excel_sheets(silent=True)
        self._render_organize_preview()
        self._refresh_filter_result_tree()
        self._refresh_history_tree()
        self._refresh_recent_error_list()
        try:
            self.notebook.select(selected_tab)
        except Exception:
            pass

    def _configure_ttk_styles(self) -> None:
        palette = self.palette
        if MODERN_UI and ttkb is not None:
            try:
                ttkb.Style(theme=palette["bootstrap_theme"])
            except Exception:
                pass

        style = ttk.Style()
        try:
            style.theme_use(style.theme_use())
        except Exception:
            pass
        style.configure("TNotebook", background=palette["root_bg"], borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            font=("微软雅黑", 9, "bold"),
            padding=[10, 3],
            background=palette["surface_soft"],
            foreground=palette["muted"],
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", palette["surface"])],
            foreground=[("selected", palette["text"])],
        )
        style.configure(
            "Treeview",
            rowheight=23,
            font=("微软雅黑", 9),
            background=palette["tree_odd"],
            fieldbackground=palette["tree_odd"],
            foreground=palette["text"],
            bordercolor=palette["border"],
        )
        style.configure(
            "Treeview.Heading",
            font=("微软雅黑", 9, "bold"),
            background=palette["surface_soft"],
            foreground=palette["text"],
            relief="flat",
        )
        style.map("Treeview", background=[("selected", palette["tree_selected"])], foreground=[("selected", "#FFFFFF")])
        style.configure(
            "TCombobox",
            fieldbackground=palette["entry_bg"],
            foreground=palette["entry_fg"],
            arrowcolor=palette["text"],
        )

    def _should_apply_default_bg(self, widget: tk.Widget, option: str = "bg") -> bool:
        defaults = self._default_widget_colors.get(widget.winfo_class(), {})
        try:
            current_value = str(widget.cget(option))
        except tk.TclError:
            return False
        default_value = defaults.get(option)
        return default_value is not None and current_value == default_value

    def _apply_theme_to_widget_tree(self, widget: tk.Widget) -> None:
        palette = self.palette
        parent_bg = palette["root_bg"]
        try:
            parent_bg = str(widget.master.cget("bg"))
        except Exception:
            pass

        cls = widget.winfo_class()
        if cls in {"Frame", "Labelframe", "LabelFrame"}:
            if self._should_apply_default_bg(widget):
                widget.configure(bg=parent_bg)
            if cls in {"Labelframe", "LabelFrame"}:
                try:
                    widget.configure(fg=palette["text"])
                except tk.TclError:
                    pass
        elif cls == "Label":
            if self._should_apply_default_bg(widget):
                widget.configure(bg=parent_bg)
            if self._should_apply_default_bg(widget, "fg"):
                widget.configure(fg=palette["text"])
        elif cls in {"Checkbutton", "Radiobutton"}:
            config: Dict[str, Any] = {
                "bg": parent_bg,
                "fg": palette["text"],
                "activebackground": parent_bg,
                "activeforeground": palette["text"],
                "selectcolor": palette["surface"],
            }
            widget.configure(**config)
        elif cls in {"Entry", "Spinbox"}:
            widget.configure(
                bg=palette["entry_bg"],
                fg=palette["entry_fg"],
                insertbackground=palette["entry_fg"],
                highlightbackground=palette["border"],
                highlightcolor=palette["primary"],
                relief="flat",
            )
        elif cls == "Listbox":
            widget.configure(
                bg=palette["entry_bg"],
                fg=palette["entry_fg"],
                selectbackground=palette["tree_selected"],
                selectforeground="#FFFFFF",
                highlightbackground=palette["border"],
                highlightcolor=palette["primary"],
            )
        elif cls == "Button" and self._should_apply_default_bg(widget):
            widget.configure(
                bg=palette["button_bg"],
                fg=palette["button_fg"],
                activebackground=palette["button_hover"],
                activeforeground=palette["button_fg"],
                relief="flat",
                bd=0,
                highlightthickness=0,
            )
        for child in widget.winfo_children():
            self._apply_theme_to_widget_tree(child)

    def _bind_scrollable_canvas(self, canvas: tk.Canvas) -> None:
        def _on_mousewheel(event):
            delta = getattr(event, "delta", 0)
            if delta:
                canvas.yview_scroll(int(-delta / 120), "units")

        def _on_linux_scroll(event):
            if getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")

        def _enter(_event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_linux_scroll)
            canvas.bind_all("<Button-5>", _on_linux_scroll)

        def _leave(_event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _enter)
        canvas.bind("<Leave>", _leave)

    def _create_scrollable_tab_body(self, parent: tk.Widget) -> tk.Frame:
        outer = tk.Frame(parent, bg=self.palette["root_bg"])
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg=self.palette["root_bg"], highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(canvas, bg=self.palette["root_bg"])
        window_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_body_configure(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfigure(window_id, width=event.width)

        body.bind("<Configure>", _on_body_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        self._bind_scrollable_canvas(canvas)
        return body

    def _current_filter_context(self, sheet_name: Optional[str] = None) -> Tuple[str, str]:
        resolved_sheet = self.excel_sheet_name.get().strip() if sheet_name is None else str(sheet_name).strip()
        return (self.excel_path.get().strip(), resolved_sheet)

    def _reset_sheet_row_filters(self) -> None:
        self.row_filter_column_name.set("")
        self.row_filter_mode.set("不过滤")
        self.row_filter_values.set("")
        self.company_exclude_keywords.set("")

    def _sync_filter_context(self, sheet_name: Optional[str] = None) -> bool:
        new_context = self._current_filter_context(sheet_name)
        previous_context = getattr(self, "_active_filter_context", ("", ""))
        changed = new_context != previous_context
        if changed:
            self._reset_sheet_row_filters()
        self._active_filter_context = new_context
        return changed

    @staticmethod
    def _sanitize_output_folder_name(name: str) -> str:
        cleaned = re.sub(r'[<>:"/\\\\|?*]+', "_", str(name).strip())
        cleaned = cleaned.rstrip(". ").strip()
        return cleaned or "筛选结果"

    def _get_effective_output_folder_path(self) -> Optional[Path]:
        if self.auto_output_by_sheet.get():
            excel_text = self.excel_path.get().strip()
            if not excel_text:
                return None
            excel_path = Path(excel_text)
            sheet_name = self.excel_sheet_name.get().strip()
            folder_name = InvoiceToolApp._sanitize_output_folder_name(sheet_name or "筛选结果")
            return excel_path.parent / folder_name

        manual = self.manual_output_folder.get().strip() or self.output_folder.get().strip()
        if not manual:
            return None
        return Path(manual)

    def _sync_output_folder_mode_ui(self) -> None:
        effective_path = self._get_effective_output_folder_path()
        display_value = str(effective_path) if effective_path else ""
        if self.auto_output_by_sheet.get():
            self.output_folder.set(display_value)
        else:
            self.output_folder.set(self.manual_output_folder.get().strip())

        if hasattr(self, "output_folder_entry"):
            if self.auto_output_by_sheet.get():
                self.output_folder_entry.config(
                    state="disabled",
                    disabledbackground=self.palette["surface_soft"],
                    disabledforeground=self.palette["entry_fg"],
                )
            else:
                self.output_folder_entry.config(state="normal")
        if hasattr(self, "output_folder_browse_btn"):
            self.output_folder_browse_btn.config(state="disabled" if self.auto_output_by_sheet.get() else "normal")

    def _on_output_mode_change(self) -> None:
        self._sync_output_folder_mode_ui()
        self._save_config()

    def _save_config(self) -> None:
        self.config["ui_theme"] = self.ui_theme.get().strip()
        self.config["rule_preset_id"] = self.rule_preset_id.get().strip()
        self.config["company_name_index"] = self.company_name_index.get()
        self.config["invoice_number_index"] = self.invoice_number_index.get()
        self.config["excel_sheet_name"] = self.excel_sheet_name.get().strip()
        self.config["selected_invoice_column_name"] = self.selected_invoice_column_name.get().strip()
        self.config["selected_company_column_name"] = self.selected_company_column_name.get().strip()
        self.config["row_filter_column_name"] = self.row_filter_column_name.get().strip()
        self.config["row_filter_mode"] = self.row_filter_mode.get().strip()
        self.config["row_filter_values"] = self.row_filter_values.get().strip()
        self.config["company_exclude_keywords"] = self.company_exclude_keywords.get().strip()
        self.config["auto_output_by_sheet"] = bool(self.auto_output_by_sheet.get())
        self.config["output_folder"] = self.manual_output_folder.get().strip()
        self.config["invoice_column_aliases"] = self.invoice_column_aliases.get().strip()
        self.config["company_column_aliases"] = self.company_column_aliases.get().strip()
        self._save_json(CONFIG_FILE, self.config)

    def _save_history(self) -> None:
        self._save_json(HISTORY_FILE, self.all_history)

    def _restore_paths(self) -> None:
        for key, var in [
            ("organize_folder", self.organize_folder_path),
            ("excel_path", self.excel_path),
            ("pdf_folder", self.pdf_folder),
        ]:
            if key in self.config:
                var.set(self.config[key])
        if "output_folder" in self.config:
            self.manual_output_folder.set(self.config["output_folder"])
        self._sync_output_folder_mode_ui()

    def _on_closing(self) -> None:
        self._save_config()
        if self._gui_log_handler is not None:
            logger.removeHandler(self._gui_log_handler)
            self._gui_log_handler = None
        if self._recent_error_handler is not None:
            logger.removeHandler(self._recent_error_handler)
            self._recent_error_handler = None
        logger.info("应用关闭")
        self.root.destroy()

    def _auto_clean_old_history(self) -> None:
        cutoff = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
        before = len(self.all_history)
        self.all_history = [r for r in self.all_history if r.get("time", "") >= cutoff][:100]
        removed = before - len(self.all_history)
        if removed > 0:
            logger.info(f"自动清理了 {removed} 条过期历史记录")
            self._save_history()

    # ==================== hover 工具 ====================

    @staticmethod
    def _bind_hover(btn: tk.Button, normal_bg: str, hover_bg: str) -> None:
        def enter(e):
            if str(btn["state"]) != "disabled":
                btn.config(bg=hover_bg)
        def leave(e):
            if str(btn["state"]) != "disabled":
                btn.config(bg=normal_bg)
        btn.bind("<Enter>", enter)
        btn.bind("<Leave>", leave)

    def _button_colors(self, role: str) -> Tuple[str, str, str]:
        palette = self.palette
        mapping = {
            "primary": (palette["primary"], palette["primary_hover"], "#FFFFFF"),
            "success": (palette["success"], palette["success_hover"], "#FFFFFF"),
            "warning": (palette["warning"], palette["warning_hover"], "#FFFFFF"),
            "danger": (palette["danger"], palette["danger_hover"], "#FFFFFF"),
            "secondary": (palette["secondary"], palette["secondary_hover"], "#FFFFFF"),
            "neutral": (palette["button_bg"], palette["button_hover"], palette["button_fg"]),
        }
        return mapping.get(role, mapping["neutral"])

    def _style_action_button(self, button: tk.Button, role: str) -> None:
        normal_bg, hover_bg, fg = self._button_colors(role)
        button.config(
            bg=normal_bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            relief="flat",
            bd=0,
            highlightthickness=0,
        )
        self._bind_hover(button, normal_bg, hover_bg)

    # ==================== 界面构建 ====================

    def _build_ui(self) -> None:
        palette = self.palette
        self.root.configure(bg=palette["root_bg"])
        self._configure_ttk_styles()

        # ─── 标题栏 ───
        title_bg = palette["title_bg"]
        title_frame = tk.Frame(self.root, bg=title_bg, pady=5, padx=10)
        title_frame.pack(fill="x")

        top_row = tk.Frame(title_frame, bg=title_bg)
        top_row.pack(fill="x")
        title_left = tk.Frame(top_row, bg=title_bg)
        title_left.pack(side="left", fill="x", expand=True)
        tk.Label(
            title_left,
            text=APP_TITLE,
            font=("微软雅黑", 17, "bold"),
            bg=title_bg,
            fg=palette["title_fg"],
        ).pack(anchor="w")
        tk.Label(
            title_left,
            text="整理、筛选、回滚与诊断都集中在一个更紧凑的工作台里。",
            font=("微软雅黑", 8),
            bg=title_bg,
            fg=palette["title_muted"],
        ).pack(anchor="w", pady=(2, 0))

        title_actions = tk.Frame(top_row, bg=title_bg)
        title_actions.pack(side="right")
        self.theme_badge = tk.Label(
            title_actions,
            text=f"{self._theme_label()} UI",
            font=("微软雅黑", 8, "bold"),
            bg=palette["title_badge_bg"],
            fg=palette["title_badge_fg"],
            padx=10,
            pady=4,
        )
        self.theme_badge.pack(side="left", padx=(0, 8))
        self.theme_toggle_btn = tk.Button(
            title_actions,
            text="切换到黑夜" if self.ui_theme.get() == "day" else "切换到白天",
            font=("微软雅黑", 9),
            bg=palette["button_bg"],
            fg=palette["button_fg"],
            activebackground=palette["button_hover"],
            activeforeground=palette["button_fg"],
            relief="flat",
            bd=0,
            padx=12,
            pady=4,
            command=self._toggle_ui_theme,
        )
        self.theme_toggle_btn.pack(side="left")
        self._style_action_button(self.theme_toggle_btn, "neutral")

        caps = []
        caps.append("✅ 拖拽" if DND_SUPPORT else "❌ 拖拽")
        caps.append("✅ Excel" if PANDAS_SUPPORT else "❌ Excel")
        caps.append("✅ 现代UI" if MODERN_UI else "原生UI")
        caps.append("✅ 报告导出" if OPENPYXL_SUPPORT else "❌ 报告(需openpyxl)")
        tk.Label(
            title_frame, text=" | ".join(caps),
            font=("微软雅黑", 8), bg=title_bg, fg=palette["title_muted"],
        ).pack(anchor="w", pady=(5, 0))

        # ─── 主内容区 (选项卡占满) ───
        content_frame = tk.Frame(self.root, bg=palette["root_bg"])
        content_frame.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        self.notebook = ttk.Notebook(content_frame)
        self.notebook.pack(fill="both", expand=True)

        self.organize_frame = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.organize_frame, text=" 📁 发票整理 ")

        self.filter_frame = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.filter_frame, text=" 🔍 发票筛选 ")

        self.history_frame = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.history_frame, text=" 📚 历史记录 ")

        self.settings_frame = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.settings_frame, text=" ⚙️ 设置 ")

        self._build_organize_tab()
        self._build_filter_tab()
        self._build_history_tab()
        self._build_settings_tab()

        # ─── 日志抽屉（默认折叠） ───
        self._build_log_drawer()

        # ─── 状态栏 ───
        status_frame = tk.Frame(self.root, bg=palette["status_bg"], bd=0)
        status_frame.pack(fill="x", side="bottom")

        self.status_var = tk.StringVar(value="就绪 - 请选择功能开始使用")
        tk.Label(
            status_frame, textvariable=self.status_var,
            font=("微软雅黑", 9), anchor="w", padx=10, pady=5, bg=palette["status_bg"], fg=palette["status_fg"],
        ).pack(side="left", fill="x", expand=True)

        self.progress_label = tk.Label(
            status_frame,
            text="",
            font=("微软雅黑", 9),
            fg=palette["muted"],
            bg=palette["status_bg"],
            padx=10,
        )
        self.progress_label.pack(side="right")
        self._apply_theme_to_widget_tree(self.root)

    # ─────────────── 日志抽屉 ───────────────

    def _build_log_drawer(self) -> None:
        """日志做成可展开/折叠的底部抽屉"""
        palette = self.palette
        self._drawer_frame = tk.Frame(self.root, bg=palette["root_bg"])
        self._drawer_frame.pack(fill="x", padx=12, pady=(0, 4))

        # 抽屉开关栏
        toggle_bar = tk.Frame(self._drawer_frame, bg=palette["log_drawer_bg"], cursor="hand2")
        toggle_bar.pack(fill="x")

        self._log_toggle_label = tk.Label(
            toggle_bar, text="▶ 运行日志（点击展开）",
            font=("微软雅黑", 9, "bold"), bg=palette["log_drawer_bg"], fg=palette["text"],
            padx=10, pady=4, cursor="hand2",
        )
        self._log_toggle_label.pack(side="left")

        # 工具按钮（始终可见）
        tk.Button(
            toggle_bar, text="清空", font=("微软雅黑", 8),
            command=self._clear_log, padx=6, pady=0, bd=0, bg=palette["log_drawer_bg"], fg=palette["text"],
        ).pack(side="right", padx=(0, 6))
        tk.Button(
            toggle_bar, text="导出", font=("微软雅黑", 8),
            command=self._export_log, padx=6, pady=0, bd=0, bg=palette["log_drawer_bg"], fg=palette["text"],
        ).pack(side="right")

        toggle_bar.bind("<Button-1>", lambda e: self._toggle_log_drawer())
        self._log_toggle_label.bind("<Button-1>", lambda e: self._toggle_log_drawer())

        # 日志内容区（初始隐藏）
        self._log_content = tk.Frame(self._drawer_frame, bg=palette["root_bg"])

        log_scroll = tk.Scrollbar(self._log_content)
        log_scroll.pack(side="right", fill="y")

        self.log_text = tk.Text(
            self._log_content, font=("Consolas", 9), wrap="word",
            yscrollcommand=log_scroll.set, bg=palette["log_bg"], fg=palette["log_fg"], height=6,
        )
        self.log_text.pack(fill="both", expand=True)
        log_scroll.config(command=self.log_text.yview)

        self.log_text.tag_config("success", foreground="#4EC9B0")
        self.log_text.tag_config("error", foreground="#FB7185" if self.ui_theme.get() == "night" else "#F44747")
        self.log_text.tag_config("warning", foreground="#FBBF24" if self.ui_theme.get() == "night" else "#DCDCAA")
        self.log_text.tag_config("info", foreground="#7DD3FC" if self.ui_theme.get() == "night" else "#569CD6")
        self.log_text.tag_config("header", foreground="#C4B5FD" if self.ui_theme.get() == "night" else "#C586C0")

        # 注册自定义 GUI handler
        self._gui_log_handler = TkTextHandler(self.log_text, self.root)
        self._gui_log_handler.setFormatter(logging.Formatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S"))
        logger.addHandler(self._gui_log_handler)

        self._recent_error_handler = RecentErrorHandler(self.root, self._append_recent_error)
        self._recent_error_handler.setFormatter(
            logging.Formatter("[%(asctime)s] %(levelname)s %(message)s", datefmt="%H:%M:%S")
        )
        logger.addHandler(self._recent_error_handler)

    def _toggle_log_drawer(self) -> None:
        if self._log_visible.get():
            self._log_content.pack_forget()
            self._log_toggle_label.config(text="▶ 运行日志（点击展开）")
            self._log_visible.set(False)
        else:
            self._log_content.pack(fill="both", expand=False)
            self._log_toggle_label.config(text="▼ 运行日志（点击收起）")
            self._log_visible.set(True)

    def _clear_log(self) -> None:
        self.log_text.delete(1.0, tk.END)

    def _export_log(self) -> None:
        fp = filedialog.asksaveasfilename(
            title="导出日志", defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")],
        )
        if fp:
            try:
                Path(fp).write_text(self.log_text.get(1.0, tk.END), "utf-8")
                logger.info(f"📄 日志已导出到：{fp}")
            except Exception as e:
                logger.error(f"导出失败：{e}")

    def _append_recent_error(self, entry: Dict[str, str]) -> None:
        self.recent_errors.insert(0, entry)
        self.recent_errors = self.recent_errors[: self.recent_error_limit]
        self._refresh_recent_error_list()

    def _refresh_recent_error_list(self) -> None:
        if not hasattr(self, "recent_error_listbox"):
            return
        self.recent_error_listbox.delete(0, tk.END)
        for entry in self.recent_errors:
            summary = entry["summary"]
            if len(summary) > 70:
                summary = summary[:67] + "..."
            self.recent_error_listbox.insert(tk.END, f"[{entry['time']}] {entry['level']} {summary}")
        self.recent_error_summary_var.set(f"最近错误 {len(self.recent_errors)} 条")
        if not self.recent_errors:
            self.recent_error_detail_var.set("运行过程中出现的错误会显示在这里，方便快速排查。")

    def _on_recent_error_select(self, event=None) -> None:
        if not hasattr(self, "recent_error_listbox"):
            return
        selection = self.recent_error_listbox.curselection()
        if not selection:
            self.recent_error_detail_var.set("请选择一条错误查看详情。")
            return
        entry = self.recent_errors[selection[0]]
        self.recent_error_detail_var.set(f"时间：{entry['time']} | 级别：{entry['level']}\n{entry['detail']}")

    def _copy_selected_recent_error(self) -> None:
        if not hasattr(self, "recent_error_listbox"):
            return
        selection = self.recent_error_listbox.curselection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一条错误。")
            return
        entry = self.recent_errors[selection[0]]
        self.root.clipboard_clear()
        self.root.clipboard_append(entry["detail"])
        self.root.update_idletasks()
        messagebox.showinfo("提示", "已复制错误详情到剪贴板。")

    def _clear_recent_errors(self) -> None:
        self.recent_errors.clear()
        self._refresh_recent_error_list()

    @staticmethod
    def _open_path_in_shell(target: Path) -> None:
        try:
            system_name = platform.system()
            if system_name == "Windows":
                __import__("os").startfile(str(target))
            elif system_name == "Darwin":
                subprocess.run(["open", str(target)], check=True)
            else:
                subprocess.run(["xdg-open", str(target)], check=True)
        except Exception:
            messagebox.showwarning("提示", f"无法自动打开，请手动查看：\n{target}")

    def _open_log_file(self) -> None:
        if not LOG_FILE.exists():
            messagebox.showwarning("提示", f"日志文件不存在：\n{LOG_FILE}")
            return
        self._open_path_in_shell(LOG_FILE)

    def _open_config_directory(self) -> None:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        self._open_path_in_shell(CONFIG_DIR)

    # ─────────────── Tab1: 发票整理 ───────────────

    def _build_organize_tab(self) -> None:
        panel = tk.Frame(self.organize_frame)
        panel.pack(fill="both", expand=True)

        # 文件夹选择
        folder_lf = tk.LabelFrame(
            panel,
            text=" 📁 选择发票文件夹" + ("（支持拖拽）" if DND_SUPPORT else ""),
            font=("微软雅黑", 10, "bold"), padx=15, pady=10,
        )
        folder_lf.pack(fill="x")

        row = tk.Frame(folder_lf)
        row.pack(fill="x")

        self.organize_folder_entry = tk.Entry(row, textvariable=self.organize_folder_path, font=("微软雅黑", 11))
        self.organize_folder_entry.pack(side="left", fill="x", expand=True)

        tk.Button(row, text="浏览", font=("微软雅黑", 10), command=self._browse_organize_folder, padx=15).pack(side="right", padx=(10, 0))
        tk.Button(row, text="🔍 扫描", font=("微软雅黑", 10), command=self._scan_files, padx=15).pack(side="right", padx=(5, 0))

        opt = tk.Frame(folder_lf)
        opt.pack(fill="x", pady=(8, 0))
        tk.Checkbutton(opt, text="包含子文件夹", font=("微软雅黑", 9), variable=self.organize_recursive).pack(side="left")
        idx = self.company_name_index.get()
        self.organize_hint = tk.Label(opt, text=f"  💡 公司名在第{idx+1}段（可在设置中修改）", font=("微软雅黑", 9), fg=self.palette["muted"])
        self.organize_hint.pack(side="left", padx=(10, 0))

        # 文件列表
        list_lf = tk.LabelFrame(panel, text=" 📋 文件预览（勾选要处理的文件）", font=("微软雅黑", 10, "bold"), padx=10, pady=8)
        list_lf.pack(fill="both", expand=True, pady=(10, 0))

        sel_bar = tk.Frame(list_lf)
        sel_bar.pack(fill="x", pady=(0, 5))
        tk.Button(sel_bar, text="✅ 全选", font=("微软雅黑", 9), command=self._select_all, padx=8).pack(side="left", padx=(0, 4))
        tk.Button(sel_bar, text="⬜ 取消全选", font=("微软雅黑", 9), command=self._deselect_all, padx=8).pack(side="left")
        self.file_count_label = tk.Label(sel_bar, text="已选择: 0 / 0", font=("微软雅黑", 9), fg=self.palette["muted"])
        self.file_count_label.pack(side="right")

        tree_frame = tk.Frame(list_lf)
        tree_frame.pack(fill="both", expand=True)

        cols = ("select", "filename", "company", "target")
        self.file_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")
        self.file_tree.heading("select", text="✓")
        self.file_tree.heading("filename", text="文件名")
        self.file_tree.heading("company", text="公司名称")
        self.file_tree.heading("target", text="目标文件夹")
        self.file_tree.column("select", width=40, anchor="center")
        self.file_tree.column("filename", width=300)
        self.file_tree.column("company", width=160)
        self.file_tree.column("target", width=160)

        scr = ttk.Scrollbar(tree_frame, orient="vertical", command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scr.set)
        self.file_tree.pack(side="left", fill="both", expand=True)
        scr.pack(side="right", fill="y")
        self.file_tree.bind("<Button-1>", self._on_tree_click)

        self.file_tree.tag_configure("evenrow", background=self.palette["tree_even"])
        self.file_tree.tag_configure("oddrow", background=self.palette["tree_odd"])
        self.file_tree.tag_configure("invalid", foreground=self.palette["muted"])
        self.file_tree.tag_configure("invalid_even", foreground=self.palette["muted"], background=self.palette["tree_even"])

        # 按钮区
        btn_bar = tk.Frame(panel)
        btn_bar.pack(fill="x", pady=10)

        self.start_btn = tk.Button(
            btn_bar, text="🚀 执行整理", font=("微软雅黑", 12, "bold"),
            padx=22, pady=7, cursor="hand2", command=self._execute_organize,
        )
        self.start_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.start_btn, "success")

        self.undo_btn = tk.Button(
            btn_bar, text="↩ 撤销上次", font=("微软雅黑", 10),
            padx=12, pady=5, cursor="hand2",
            command=self._undo_last_move, state="disabled",
        )
        self.undo_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.undo_btn, "warning")

        self.undo_all_btn = tk.Button(
            btn_bar, text="↩ 撤销全部", font=("微软雅黑", 10),
            padx=12, pady=5, cursor="hand2",
            command=self._undo_all_moves, state="disabled",
        )
        self.undo_all_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.undo_all_btn, "danger")

        self.cancel_org_btn = tk.Button(
            btn_bar, text="⏹ 取消", font=("微软雅黑", 10),
            padx=12, pady=5, cursor="hand2",
            command=self._cancel_task, state="disabled",
        )
        self.cancel_org_btn.pack(side="left")
        self._style_action_button(self.cancel_org_btn, "secondary")

        self.organize_progress = ttk.Progressbar(btn_bar, mode="determinate", length=200)
        self.organize_progress.pack(side="right")

    # ─────────────── Tab2: 发票筛选 ───────────────

    def _build_filter_tab(self) -> None:
        panel = self._create_scrollable_tab_body(self.filter_frame)
        if not PANDAS_SUPPORT:
            tk.Label(
                panel, text="⚠️ 此功能需要 pandas\n\n安装命令：python -m pip install pandas openpyxl",
                font=("微软雅黑", 12), fg=self.palette["danger"], justify="center",
            ).pack(pady=40)
            return

        # 帮助
        self.help_visible = tk.BooleanVar(value=False)
        hbf = tk.Frame(panel)
        hbf.pack(fill="x", pady=(0, 6))
        self.help_btn = tk.Button(hbf, text="📖 显示使用说明", font=("微软雅黑", 9), command=self._toggle_help)
        self.help_btn.pack(side="left")

        self.help_content = tk.LabelFrame(panel, text="📋 使用说明", font=("微软雅黑", 10, "bold"), padx=15, pady=10)
        tk.Label(
            self.help_content,
            text=(
                "【Excel格式】 .xlsx/.xls，可选择工作表，需包含发票号列\n"
                "【PDF命名】 dzfp_发票号码_公司名称_时间戳.pdf\n"
                "【步骤】 ① 选Excel/工作表 → ② 选PDF文件夹 → ③ 选导出文件夹 → ④ 预览或筛选\n"
                "【高级】 可在设置里补充自定义发票列别名"
            ),
            font=("微软雅黑", 9), justify="left", anchor="w",
        ).pack(fill="x")

        if not self.config.get("help_seen"):
            self.config["help_seen"] = True

        # 路径
        self.file_path_frame = tk.LabelFrame(
            panel, text=" 📂 文件路径设置 ", font=("微软雅黑", 10, "bold"), padx=12, pady=10,
        )
        self.file_path_frame.pack(fill="x", pady=8)

        path_grid = tk.Frame(self.file_path_frame)
        path_grid.pack(fill="x")
        path_grid.grid_columnconfigure(1, weight=1)
        path_grid.grid_columnconfigure(4, weight=1)

        tk.Label(path_grid, text="Excel文件:", font=("微软雅黑", 9), width=10, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 4), pady=3)
        tk.Entry(path_grid, textvariable=self.excel_path, font=("微软雅黑", 9)).grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=3)
        tk.Button(path_grid, text="浏览", command=self._browse_excel, padx=8).grid(row=0, column=2, sticky="ew", padx=(0, 12), pady=3)

        tk.Label(path_grid, text="工作表:", font=("微软雅黑", 9), width=8, anchor="w").grid(row=0, column=3, sticky="w", padx=(0, 4), pady=3)
        self.excel_sheet_combo = ttk.Combobox(
            path_grid,
            textvariable=self.excel_sheet_name,
            state="readonly",
            font=("微软雅黑", 9),
        )
        self.excel_sheet_combo.grid(row=0, column=4, sticky="ew", padx=(0, 6), pady=3)
        self.excel_sheet_combo.bind("<<ComboboxSelected>>", self._on_excel_sheet_change)
        tk.Button(path_grid, text="刷新", command=self._refresh_excel_sheets, padx=8).grid(row=0, column=5, sticky="ew", pady=3)

        tk.Label(path_grid, text="PDF文件夹:", font=("微软雅黑", 9), width=10, anchor="w").grid(row=1, column=0, sticky="w", padx=(0, 4), pady=3)
        tk.Entry(path_grid, textvariable=self.pdf_folder, font=("微软雅黑", 9)).grid(row=1, column=1, sticky="ew", padx=(0, 6), pady=3)
        tk.Button(path_grid, text="浏览", command=self._browse_pdf_folder, padx=8).grid(row=1, column=2, sticky="ew", padx=(0, 12), pady=3)

        tk.Label(path_grid, text="导出文件夹:", font=("微软雅黑", 9), width=8, anchor="w").grid(row=1, column=3, sticky="w", padx=(0, 4), pady=3)
        self.output_folder_entry = tk.Entry(path_grid, textvariable=self.output_folder, font=("微软雅黑", 9))
        self.output_folder_entry.grid(row=1, column=4, sticky="ew", padx=(0, 6), pady=3)
        self.output_folder_browse_btn = tk.Button(path_grid, text="浏览", command=self._browse_output_folder, padx=8)
        self.output_folder_browse_btn.grid(row=1, column=5, sticky="ew", pady=3)

        auto_output_row = tk.Frame(path_grid)
        auto_output_row.grid(row=2, column=3, columnspan=3, sticky="w", pady=(2, 0))
        tk.Checkbutton(
            auto_output_row,
            text="按 Excel 同目录 + 当前工作表自动建导出文件夹",
            font=("微软雅黑", 9),
            variable=self.auto_output_by_sheet,
            command=self._on_output_mode_change,
        ).pack(side="left")

        fopt = tk.Frame(panel)
        fopt.pack(fill="x", pady=(2, 0))
        tk.Checkbutton(fopt, text="包含子文件夹", font=("微软雅黑", 9), variable=self.filter_recursive).pack(side="left")
        tk.Label(
            fopt,
            text="默认布局已优化为非最大化也能完整查看主信息。",
            font=("微软雅黑", 8),
            fg=self.palette["muted"],
        ).pack(side="right")
        self._sync_output_folder_mode_ui()

        self._build_workbook_analysis_panel(panel)

        # 按钮
        fbtn = tk.Frame(panel)
        fbtn.pack(fill="x", pady=10)

        self.filter_preview_btn = tk.Button(
            fbtn, text="👁 预览匹配", font=("微软雅黑", 10),
            padx=15, pady=6, cursor="hand2", command=self._preview_filter,
        )
        self.filter_preview_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.filter_preview_btn, "secondary")

        self.filter_run_btn = tk.Button(
            fbtn, text="🚀 开始筛选并导出", font=("微软雅黑", 12, "bold"),
            padx=22, pady=7, cursor="hand2", command=self._run_filter,
        )
        self.filter_run_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.filter_run_btn, "primary")

        tk.Button(fbtn, text="📂 打开导出文件夹", font=("微软雅黑", 10), padx=12, pady=5, command=self._open_output_folder).pack(side="left", padx=(0, 8))

        self.cancel_flt_btn = tk.Button(
            fbtn, text="⏹ 取消", font=("微软雅黑", 10),
            padx=12, pady=5, cursor="hand2",
            command=self._cancel_task, state="disabled",
        )
        self.cancel_flt_btn.pack(side="left")
        self._style_action_button(self.cancel_flt_btn, "secondary")

        self.filter_progress = ttk.Progressbar(fbtn, mode="determinate", length=220)
        self.filter_progress.pack(side="right")

        summary_lf = tk.LabelFrame(panel, text=" 📌 本次摘要 ", font=("微软雅黑", 10, "bold"), padx=12, pady=10)
        summary_lf.pack(fill="x", pady=(0, 8))
        tk.Label(summary_lf, textvariable=self.filter_summary_title, font=("微软雅黑", 11, "bold"), fg=self.palette["text"]).pack(anchor="w")
        tk.Label(
            summary_lf,
            textvariable=self.filter_summary_subtitle,
            font=("微软雅黑", 9),
            fg=self.palette["muted"],
            justify="left",
            wraplength=980,
        ).pack(anchor="w", pady=(4, 10))

        cards_row = tk.Frame(summary_lf)
        cards_row.pack(fill="x")
        for index, metric_key in enumerate(("metric1", "metric2", "metric3", "metric4", "metric5", "metric6")):
            bg, fg = self.palette["card_palette"][index]
            self._create_filter_metric_card(cards_row, metric_key, bg, fg)

        # 结果
        res_lf = tk.LabelFrame(panel, text=" 📊 筛选结果 ", font=("微软雅黑", 10, "bold"), padx=10, pady=8)
        res_lf.pack(fill="x", pady=(8, 0))

        tool_row = tk.Frame(res_lf)
        tool_row.pack(fill="x", pady=(0, 8))
        tk.Label(tool_row, text="状态:", font=("微软雅黑", 9)).pack(side="left")
        self.filter_status_combo = ttk.Combobox(
            tool_row,
            textvariable=self.filter_result_status,
            values=FILTER_RESULT_STATUS_OPTIONS,
            state="readonly",
            width=10,
            font=("微软雅黑", 9),
        )
        self.filter_status_combo.pack(side="left", padx=(6, 10))
        self.filter_status_combo.bind("<<ComboboxSelected>>", self._on_filter_result_filters_changed)

        tk.Label(tool_row, text="搜索:", font=("微软雅黑", 9)).pack(side="left")
        self.filter_keyword_entry = tk.Entry(tool_row, textvariable=self.filter_result_keyword, font=("微软雅黑", 9))
        self.filter_keyword_entry.pack(side="left", fill="x", expand=True, padx=(6, 10))
        self.filter_keyword_entry.bind("<KeyRelease>", self._on_filter_result_filters_changed)

        tk.Button(tool_row, text="重置筛选条件", font=("微软雅黑", 9), padx=10, command=self._reset_filter_result_filters).pack(side="left", padx=(0, 6))
        self.copy_missing_btn = tk.Button(tool_row, text="复制未匹配发票号", font=("微软雅黑", 9), padx=10, command=self._copy_missing_invoices, state="disabled")
        self.copy_missing_btn.pack(side="left", padx=(0, 6))
        self.open_result_btn = tk.Button(tool_row, text="打开选中结果", font=("微软雅黑", 9), padx=10, command=self._open_selected_filter_result, state="disabled")
        self.open_result_btn.pack(side="left")
        self.filter_result_count_label = tk.Label(tool_row, text="显示 0 / 0 条", font=("微软雅黑", 9), fg=self.palette["muted"])
        self.filter_result_count_label.pack(side="right")

        tree_frame = tk.Frame(res_lf)
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        cols = ("status", "invoice", "pdf", "detail")
        self.filter_result_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")
        self.filter_result_tree.heading("status", text="状态", command=lambda: self._sort_filter_results("status"))
        self.filter_result_tree.heading("invoice", text="发票号", command=lambda: self._sort_filter_results("invoice"))
        self.filter_result_tree.heading("pdf", text="文件名", command=lambda: self._sort_filter_results("pdf"))
        self.filter_result_tree.heading("detail", text="详情", command=lambda: self._sort_filter_results("detail"))
        self.filter_result_tree.column("status", width=110, anchor="center")
        self.filter_result_tree.column("invoice", width=150, anchor="center")
        self.filter_result_tree.column("pdf", width=260)
        self.filter_result_tree.column("detail", width=420)

        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.filter_result_tree.yview)
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.filter_result_tree.xview)
        self.filter_result_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.filter_result_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        self.filter_result_tree.tag_configure("evenrow", background=self.palette["tree_even"])
        self.filter_result_tree.tag_configure("oddrow", background=self.palette["tree_odd"])
        self.filter_result_tree.tag_configure("success", foreground="#1B5E20")
        self.filter_result_tree.tag_configure("missing", foreground="#C62828")
        self.filter_result_tree.tag_configure("skip", foreground="#EF6C00")
        self.filter_result_tree.tag_configure("error", foreground="#AD1457")
        self.filter_result_tree.tag_configure("conflict", foreground="#6A1B9A")
        self.filter_result_tree.tag_configure("preview", foreground="#1565C0")
        self.filter_result_tree.bind("<<TreeviewSelect>>", self._on_filter_result_select)
        self.filter_result_tree.bind("<Double-1>", self._open_selected_filter_result)

        detail_frame = tk.Frame(res_lf, bg=self.palette["detail_bg"], bd=1, relief="solid")
        detail_frame.pack(fill="x", pady=(8, 0))
        tk.Label(
            detail_frame,
            textvariable=self.filter_detail_var,
            font=("微软雅黑", 9),
            bg=self.palette["detail_bg"],
            fg=self.palette["detail_fg"],
            justify="left",
            wraplength=980,
            anchor="w",
            padx=10,
            pady=8,
        ).pack(fill="x")

    def _build_workbook_analysis_panel(self, parent: tk.Widget) -> None:
        analysis_lf = tk.LabelFrame(
            parent,
            text=" 工作簿分析 ",
            font=("微软雅黑", 10, "bold"),
            padx=10,
            pady=8,
        )
        analysis_lf.pack(fill="x", pady=(8, 0))

        top_row = tk.Frame(analysis_lf)
        top_row.pack(fill="x", pady=(0, 8))
        tk.Label(
            top_row,
            textvariable=self.workbook_analysis_summary_var,
            font=("微软雅黑", 9),
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
        ).pack(side="left", fill="x", expand=True)
        tk.Button(top_row, text="分析工作簿", padx=10, command=self._refresh_workbook_analysis).pack(side="right")

        content = tk.Frame(analysis_lf)
        content.pack(fill="x")

        left_panel = tk.Frame(content)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_panel = tk.Frame(content, bg=self.palette["surface_alt"], bd=1, relief="solid")
        right_panel.pack(side="right", fill="both")

        tree_frame = tk.Frame(left_panel)
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        cols = ("sheet", "shape", "invoice", "company", "status")
        self.workbook_sheet_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=7, selectmode="browse")
        self.workbook_sheet_tree.heading("sheet", text="工作表")
        self.workbook_sheet_tree.heading("shape", text="行/列")
        self.workbook_sheet_tree.heading("invoice", text="发票列")
        self.workbook_sheet_tree.heading("company", text="公司列")
        self.workbook_sheet_tree.heading("status", text="状态")
        self.workbook_sheet_tree.column("sheet", width=150)
        self.workbook_sheet_tree.column("shape", width=80, anchor="center")
        self.workbook_sheet_tree.column("invoice", width=150)
        self.workbook_sheet_tree.column("company", width=150)
        self.workbook_sheet_tree.column("status", width=90, anchor="center")
        self.workbook_sheet_tree.grid(row=0, column=0, sticky="nsew")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.workbook_sheet_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.workbook_sheet_tree.configure(yscrollcommand=tree_scroll.set)
        self.workbook_sheet_tree.tag_configure("evenrow", background=self.palette["tree_even"])
        self.workbook_sheet_tree.tag_configure("oddrow", background=self.palette["tree_odd"])
        self.workbook_sheet_tree.tag_configure("recommended", foreground="#1B5E20")
        self.workbook_sheet_tree.tag_configure("usable", foreground="#1565C0")
        self.workbook_sheet_tree.tag_configure("warning", foreground="#EF6C00")
        self.workbook_sheet_tree.tag_configure("error", foreground="#C62828")
        self.workbook_sheet_tree.bind("<<TreeviewSelect>>", self._on_workbook_sheet_select)
        self.workbook_sheet_tree.bind("<Double-1>", self._on_workbook_sheet_select)

        tk.Label(
            right_panel,
            text="当前 sheet 详情",
            font=("微软雅黑", 10, "bold"),
            bg=self.palette["surface_alt"],
            fg=self.palette["text"],
            anchor="w",
            padx=10,
            pady=8,
        ).pack(fill="x")

        tk.Label(
            right_panel,
            textvariable=self.workbook_sheet_overview_var,
            font=("微软雅黑", 9),
            bg=self.palette["surface_alt"],
            fg=self.palette["text"],
            justify="left",
            wraplength=400,
            anchor="w",
            padx=10,
        ).pack(fill="x")

        picker_frame = tk.Frame(right_panel, bg=self.palette["surface_alt"], padx=10, pady=8)
        picker_frame.pack(fill="x")
        picker_frame.grid_columnconfigure(1, weight=1)

        tk.Label(picker_frame, text="发票列:", font=("微软雅黑", 9), bg=self.palette["surface_alt"], fg=self.palette["text"]).grid(row=0, column=0, sticky="w", pady=3)
        self.analysis_invoice_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.selected_invoice_column_name,
            state="readonly",
            font=("微软雅黑", 9),
        )
        self.analysis_invoice_combo.grid(row=0, column=1, sticky="ew", pady=3)
        self.analysis_invoice_combo.bind("<<ComboboxSelected>>", self._on_analysis_invoice_column_change)

        tk.Label(picker_frame, text="公司列:", font=("微软雅黑", 9), bg=self.palette["surface_alt"], fg=self.palette["text"]).grid(row=1, column=0, sticky="w", pady=3)
        self.analysis_company_combo = ttk.Combobox(
            picker_frame,
            textvariable=self.selected_company_column_name,
            state="readonly",
            font=("微软雅黑", 9),
        )
        self.analysis_company_combo.grid(row=1, column=1, sticky="ew", pady=3)
        self.analysis_company_combo.bind("<<ComboboxSelected>>", self._on_analysis_company_column_change)

        filter_frame = tk.LabelFrame(
            right_panel,
            text="行筛选条件",
            font=("微软雅黑", 9, "bold"),
            bg=self.palette["surface_alt"],
            fg=self.palette["text"],
            padx=8,
            pady=8,
        )
        filter_frame.pack(fill="x", padx=10, pady=(0, 8))
        filter_frame.grid_columnconfigure(1, weight=1)

        tk.Label(filter_frame, text="条件列:", font=("微软雅黑", 9), bg=self.palette["surface_alt"], fg=self.palette["text"]).grid(row=0, column=0, sticky="w", pady=3)
        self.row_filter_column_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.row_filter_column_name,
            state="readonly",
            font=("微软雅黑", 9),
        )
        self.row_filter_column_combo.grid(row=0, column=1, sticky="ew", pady=3)
        self.row_filter_column_combo.bind("<<ComboboxSelected>>", self._on_row_filter_rule_change)

        tk.Label(filter_frame, text="筛选模式:", font=("微软雅黑", 9), bg=self.palette["surface_alt"], fg=self.palette["text"]).grid(row=1, column=0, sticky="w", pady=3)
        self.row_filter_mode_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.row_filter_mode,
            values=FILTER_RULE_MODE_OPTIONS,
            state="readonly",
            font=("微软雅黑", 9),
        )
        self.row_filter_mode_combo.grid(row=1, column=1, sticky="ew", pady=3)
        self.row_filter_mode_combo.bind("<<ComboboxSelected>>", self._on_row_filter_rule_change)

        tk.Label(filter_frame, text="条件值:", font=("微软雅黑", 9), bg=self.palette["surface_alt"], fg=self.palette["text"]).grid(row=2, column=0, sticky="w", pady=3)
        self.row_filter_values_entry = tk.Entry(filter_frame, textvariable=self.row_filter_values, font=("微软雅黑", 9))
        self.row_filter_values_entry.grid(row=2, column=1, sticky="ew", pady=3)
        self.row_filter_values_entry.bind("<FocusOut>", self._on_row_filter_rule_change)

        tk.Label(filter_frame, text="排除公司:", font=("微软雅黑", 9), bg=self.palette["surface_alt"], fg=self.palette["text"]).grid(row=3, column=0, sticky="w", pady=3)
        self.company_exclude_entry = tk.Entry(filter_frame, textvariable=self.company_exclude_keywords, font=("微软雅黑", 9))
        self.company_exclude_entry.grid(row=3, column=1, sticky="ew", pady=3)
        self.company_exclude_entry.bind("<FocusOut>", self._on_row_filter_rule_change)

        tk.Label(
            filter_frame,
            text="示例：条件列=是否抵扣，模式=等于任一，条件值=是；排除公司可填 临时, 乱标记。",
            font=("微软雅黑", 8),
            bg=self.palette["surface_alt"],
            fg=self.palette["muted"],
            justify="left",
            wraplength=380,
            anchor="w",
        ).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        tk.Label(
            right_panel,
            text="样本预览",
            font=("微软雅黑", 9, "bold"),
            bg=self.palette["surface_alt"],
            fg=self.palette["text"],
            anchor="w",
            padx=10,
        ).pack(fill="x", pady=(0, 4))
        tk.Label(
            right_panel,
            textvariable=self.workbook_sheet_sample_var,
            font=("微软雅黑", 8),
            bg=self.palette["surface_alt"],
            fg=self.palette["muted"],
            justify="left",
            wraplength=400,
            anchor="w",
            padx=10,
            pady=8,
        ).pack(fill="x")

    # ─────────────── Tab3: 历史记录 ───────────────

    def _build_history_tab(self) -> None:
        tk.Label(
            self.history_frame,
            text="📚 所有文件操作记录，可选择任意记录回滚（自动保留30天/100条）",
            font=("微软雅黑", 10), fg=self.palette["muted"],
        ).pack(anchor="w", pady=(0, 8))

        filter_bar = tk.Frame(self.history_frame)
        filter_bar.pack(fill="x", pady=(0, 8))
        tk.Label(filter_bar, text="类型:", font=("微软雅黑", 9)).pack(side="left")
        self.history_type_combo = ttk.Combobox(
            filter_bar,
            textvariable=self.history_type_filter,
            values=HISTORY_TYPE_OPTIONS,
            state="readonly",
            width=8,
            font=("微软雅黑", 9),
        )
        self.history_type_combo.pack(side="left", padx=(6, 10))
        self.history_type_combo.bind("<<ComboboxSelected>>", self._on_history_filters_changed)

        tk.Label(filter_bar, text="时间:", font=("微软雅黑", 9)).pack(side="left")
        self.history_date_combo = ttk.Combobox(
            filter_bar,
            textvariable=self.history_date_filter,
            values=HISTORY_DATE_OPTIONS,
            state="readonly",
            width=10,
            font=("微软雅黑", 9),
        )
        self.history_date_combo.pack(side="left", padx=(6, 10))
        self.history_date_combo.bind("<<ComboboxSelected>>", self._on_history_filters_changed)

        tk.Label(filter_bar, text="搜索:", font=("微软雅黑", 9)).pack(side="left")
        self.history_keyword_entry = tk.Entry(filter_bar, textvariable=self.history_keyword, font=("微软雅黑", 9))
        self.history_keyword_entry.pack(side="left", fill="x", expand=True, padx=(6, 10))
        self.history_keyword_entry.bind("<KeyRelease>", self._on_history_filters_changed)

        tk.Button(filter_bar, text="重置", font=("微软雅黑", 9), padx=10, command=self._reset_history_filters).pack(side="left")
        tk.Label(filter_bar, textvariable=self.history_summary_var, font=("微软雅黑", 9), fg=self.palette["muted"]).pack(side="right")

        tf = tk.Frame(self.history_frame)
        tf.pack(fill="both", expand=True)

        cols = ("time", "folder", "count", "type")
        self.history_tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        self.history_tree.heading("time", text="时间")
        self.history_tree.heading("folder", text="操作文件夹")
        self.history_tree.heading("count", text="文件数")
        self.history_tree.heading("type", text="类型")
        self.history_tree.column("time", width=150)
        self.history_tree.column("folder", width=370)
        self.history_tree.column("count", width=80, anchor="center")
        self.history_tree.column("type", width=100, anchor="center")
        self.history_tree.tag_configure("evenrow", background=self.palette["tree_even"])
        self.history_tree.tag_configure("oddrow", background=self.palette["tree_odd"])

        hscr = ttk.Scrollbar(tf, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=hscr.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        hscr.pack(side="right", fill="y")
        self.history_tree.bind("<Double-1>", lambda event: self._view_history_detail())

        hbtn = tk.Frame(self.history_frame)
        hbtn.pack(fill="x", pady=12)

        rb = tk.Button(hbtn, text="🔄 回滚选中", font=("微软雅黑", 10), padx=12, pady=5, command=self._rollback_selected)
        rb.pack(side="left", padx=(0, 8))
        self._style_action_button(rb, "warning")

        tk.Button(hbtn, text="🔍 查看详情", font=("微软雅黑", 10), padx=12, pady=5, command=self._view_history_detail).pack(side="left", padx=(0, 8))
        tk.Button(hbtn, text="📂 打开文件夹", font=("微软雅黑", 10), padx=12, pady=5, command=self._open_history_folder).pack(side="left", padx=(0, 8))

        cb = tk.Button(hbtn, text="🗑️ 清空历史", font=("微软雅黑", 10), padx=12, pady=5, command=self._clear_all_history)
        cb.pack(side="left")
        self._style_action_button(cb, "danger")

        tk.Button(hbtn, text="🔄 刷新", font=("微软雅黑", 10), padx=12, pady=5, command=self._refresh_history_tree).pack(side="right")

        self._refresh_history_tree()

    # ─────────────── Tab4: 设置 ───────────────

    def _get_rule_preset(self):
        return self._preset_by_id.get(self.rule_preset_id.get(), self._preset_by_id[DEFAULT_RULE_PRESET_ID])

    def _sync_rule_preset_ui(self) -> None:
        preset = self._get_rule_preset()
        self.rule_preset_name.set(preset.name)
        self.rule_preset_desc.set(preset.description)

    def _on_rule_preset_change(self, event=None) -> None:
        selected_name = self.rule_preset_name.get().strip()
        preset = next((item for item in self.rule_presets if item.name == selected_name), None)
        if preset is None:
            preset = self._preset_by_id[DEFAULT_RULE_PRESET_ID]
        self.rule_preset_id.set(preset.preset_id)
        self._sync_rule_preset_ui()
        self._save_config()

    def _apply_rule_preset(self) -> None:
        preset = self._get_rule_preset()
        if preset.preset_id == "custom":
            messagebox.showinfo("提示", "“手动配置”预设不会覆盖当前设置。")
            return

        self.company_name_index.set(preset.company_name_index)
        self.invoice_number_index.set(preset.invoice_number_index)
        self.invoice_column_aliases.set(", ".join(preset.invoice_column_aliases))
        self.organize_hint.config(text=f"  💡 公司名在第{preset.company_name_index + 1}段（可在设置中修改）")
        self._save_config()
        logger.info(f"✅ 已应用预设：{preset.name}")
        messagebox.showinfo("提示", f"已应用预设：{preset.name}\n请重新扫描文件使新规则生效。")

    def _get_filename_parser(self) -> SegmentFilenameParser:
        preset = self._get_rule_preset()
        return SegmentFilenameParser(separator=preset.filename_separator)

    def _get_column_resolver(self) -> SmartInvoiceColumnResolver:
        preset = self._get_rule_preset()
        exact_names = tuple(
            dict.fromkeys(list(InvoiceFilter.EXACT_COL_NAMES) + list(preset.exact_column_names))
        )
        exclude_keywords = tuple(
            dict.fromkeys(list(InvoiceFilter.EXCLUDE_KEYWORDS) + list(preset.exclude_keywords))
        )
        return SmartInvoiceColumnResolver(
            exact_column_names=exact_names,
            exclude_keywords=exclude_keywords,
        )

    def _get_report_exporter(self) -> OpenpyxlFilterReportExporter:
        preset = self._get_rule_preset()
        if preset.report_style == "standard":
            return OpenpyxlFilterReportExporter()
        return OpenpyxlFilterReportExporter()

    def _create_filter_metric_card(self, parent: tk.Widget, metric_key: str, bg: str, fg: str) -> None:
        card = tk.Frame(parent, bg=bg, bd=1, relief="flat", padx=10, pady=8)
        card.pack(side="left", fill="x", expand=True, padx=3)
        tk.Label(
            card,
            textvariable=self.filter_metric_labels[metric_key],
            font=("微软雅黑", 8, "bold"),
            bg=bg,
            fg=fg,
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            card,
            textvariable=self.filter_metric_values[metric_key],
            font=("微软雅黑", 15, "bold"),
            bg=bg,
            fg=fg,
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))

    def _update_filter_summary(
        self,
        title: str,
        subtitle: str,
        metrics: List[Tuple[str, str]],
    ) -> None:
        self.filter_summary_title.set(title)
        self.filter_summary_subtitle.set(subtitle)
        for index in range(1, 7):
            key = f"metric{index}"
            if index <= len(metrics):
                label, value = metrics[index - 1]
            else:
                label, value = "-", "-"
            self.filter_metric_labels[key].set(label)
            self.filter_metric_values[key].set(value)

    def _clear_filter_results(self, reset_filters: bool = False) -> None:
        self.filter_result_rows = []
        self.filter_result_selection.clear()
        self.filter_missing_invoices = []
        if reset_filters:
            self.filter_result_status.set("全部")
            self.filter_result_keyword.set("")
            self._update_filter_summary(
                "等待预览或筛选",
                "先选择 Excel、PDF 和导出目录，然后执行预览或筛选。",
                [
                    ("Excel发票", "0"),
                    ("命中结果", "0"),
                    ("未匹配", "0"),
                    ("异常/冲突", "0"),
                    ("PDF扫描", "0"),
                    ("其他状态", "0"),
                ],
            )
        if hasattr(self, "filter_result_tree"):
            self.filter_result_tree.delete(*self.filter_result_tree.get_children())
        if hasattr(self, "filter_result_count_label"):
            self.filter_result_count_label.config(text="显示 0 / 0 条")
        if hasattr(self, "copy_missing_btn"):
            self.copy_missing_btn.config(state="disabled")
        if hasattr(self, "open_result_btn"):
            self.open_result_btn.config(state="disabled")
        self.filter_detail_var.set("提示：结果将显示在下方表格中，可按状态过滤或搜索发票号。")

    def _set_filter_results(
        self,
        rows: List[FilterResultRow],
        missing_invoices: Optional[List[str]] = None,
    ) -> None:
        self.filter_result_rows = list(rows)
        self.filter_missing_invoices = list(missing_invoices or [])
        if hasattr(self, "copy_missing_btn"):
            self.copy_missing_btn.config(state="normal" if self.filter_missing_invoices else "disabled")
        self._refresh_filter_result_tree()

    def _refresh_filter_result_tree(self) -> None:
        if not hasattr(self, "filter_result_tree"):
            return

        self.filter_result_tree.delete(*self.filter_result_tree.get_children())
        self.filter_result_selection.clear()

        filtered_rows = filter_filter_result_rows(
            self.filter_result_rows,
            status_filter=self.filter_result_status.get(),
            keyword=self.filter_result_keyword.get(),
        )
        visible_rows = sort_filter_result_rows(
            filtered_rows,
            sort_key=self.filter_result_sort_key,
            descending=self.filter_result_sort_desc,
        )

        for index, row in enumerate(visible_rows):
            stripe = "evenrow" if index % 2 == 0 else "oddrow"
            status_tag = {
                "未匹配": "missing",
                "复制失败": "error",
                "重复冲突": "conflict",
                "已跳过": "skip",
                "已导出": "success",
                "可匹配": "preview",
            }.get(row.status, "")
            item_id = self.filter_result_tree.insert(
                "",
                "end",
                values=(row.status, row.invoice_number, row.pdf_name, row.detail),
                tags=tuple(tag for tag in (stripe, status_tag) if tag),
            )
            self.filter_result_selection[item_id] = row

        if hasattr(self, "filter_result_count_label"):
            self.filter_result_count_label.config(text=f"显示 {len(visible_rows)} / {len(self.filter_result_rows)} 条")
        if visible_rows:
            self.filter_detail_var.set("提示：双击可打开选中结果对应的文件，或用上方条件继续筛选。")
        else:
            self.filter_detail_var.set("当前没有符合条件的结果，请调整筛选状态或搜索关键字。")

        if hasattr(self, "open_result_btn"):
            self.open_result_btn.config(state="disabled")

    def _on_filter_result_filters_changed(self, event=None) -> None:
        self._refresh_filter_result_tree()

    def _sort_filter_results(self, sort_key: str) -> None:
        if self.filter_result_sort_key == sort_key:
            self.filter_result_sort_desc = not self.filter_result_sort_desc
        else:
            self.filter_result_sort_key = sort_key
            self.filter_result_sort_desc = False
        self._refresh_filter_result_tree()

    def _get_selected_filter_result(self) -> Optional[FilterResultRow]:
        if not hasattr(self, "filter_result_tree"):
            return None
        selection = self.filter_result_tree.selection()
        if not selection:
            return None
        return self.filter_result_selection.get(selection[0])

    def _on_filter_result_select(self, event=None) -> None:
        row = self._get_selected_filter_result()
        if row is None:
            self.filter_detail_var.set("提示：选中某一行后，这里会显示更详细的信息。")
            if hasattr(self, "open_result_btn"):
                self.open_result_btn.config(state="disabled")
            return

        detail_parts = [f"状态：{row.status}"]
        if row.invoice_number:
            detail_parts.append(f"发票号：{row.invoice_number}")
        if row.pdf_name:
            detail_parts.append(f"文件：{row.pdf_name}")
        if row.detail:
            detail_parts.append(f"详情：{row.detail}")
        if row.path:
            detail_parts.append(f"路径：{row.path}")
        self.filter_detail_var.set(" | ".join(detail_parts))
        if hasattr(self, "open_result_btn"):
            self.open_result_btn.config(state="normal" if row.path else "disabled")

    def _open_selected_filter_result(self, event=None) -> None:
        row = self._get_selected_filter_result()
        if row is None or not row.path:
            return

        target = Path(row.path)
        if not target.exists():
            messagebox.showwarning("提示", f"目标不存在：\n{target}")
            return

        try:
            system_name = platform.system()
            if system_name == "Windows":
                __import__("os").startfile(str(target))
            elif system_name == "Darwin":
                subprocess.run(["open", str(target)], check=True)
            else:
                subprocess.run(["xdg-open", str(target)], check=True)
        except Exception:
            messagebox.showwarning("提示", f"无法自动打开，请手动查看：\n{target}")

    def _copy_missing_invoices(self) -> None:
        if not self.filter_missing_invoices:
            messagebox.showinfo("提示", "当前没有未匹配的发票号。")
            return
        payload = "\n".join(self.filter_missing_invoices)
        self.root.clipboard_clear()
        self.root.clipboard_append(payload)
        self.root.update_idletasks()
        messagebox.showinfo("提示", f"已复制 {len(self.filter_missing_invoices)} 个未匹配发票号到剪贴板。")

    def _reset_filter_result_filters(self) -> None:
        self.filter_result_status.set("全部")
        self.filter_result_keyword.set("")
        self._refresh_filter_result_tree()

    def _build_settings_tab(self) -> None:
        panel = self._create_scrollable_tab_body(self.settings_frame)
        container = tk.Frame(panel, bg=self.palette["root_bg"])
        container.pack(fill="both", expand=True)

        left_panel = tk.Frame(container, bg=self.palette["root_bg"])
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_panel = tk.Frame(container, width=360, bg=self.palette["root_bg"])
        right_panel.pack(side="right", fill="both")
        right_panel.pack_propagate(False)

        tk.Label(left_panel, text="⚙️ 设置中心", font=("微软雅黑", 11, "bold"), fg=self.palette["text"]).pack(anchor="w", pady=(0, 12))

        appearance_frame = tk.LabelFrame(left_panel, text="界面外观", padx=12, pady=10, font=("微软雅黑", 10, "bold"))
        appearance_frame.pack(fill="x", pady=(0, 12))
        appearance_row = tk.Frame(appearance_frame)
        appearance_row.pack(fill="x")
        tk.Label(appearance_row, text="主题模式:", font=("微软雅黑", 10), width=12, anchor="w").pack(side="left")
        self.ui_theme_combo = ttk.Combobox(
            appearance_row,
            textvariable=self.ui_theme_label,
            values=[UI_THEME_LABELS[item] for item in UI_THEME_OPTIONS],
            state="readonly",
            width=10,
            font=("微软雅黑", 10),
        )
        self.ui_theme_combo.pack(side="left", padx=5)
        self.ui_theme_combo.bind("<<ComboboxSelected>>", self._on_ui_theme_change)
        tk.Label(
            appearance_frame,
            text=f"当前为 {self._theme_label()} UI。切换后会即时重绘界面，当前数据和记录不会丢失。",
            font=("微软雅黑", 9),
            fg=self.palette["muted"],
            justify="left",
            wraplength=620,
            anchor="w",
        ).pack(anchor="w", pady=(8, 0))

        preset_frame = tk.LabelFrame(left_panel, text="规则预设", padx=12, pady=10, font=("微软雅黑", 10, "bold"))
        preset_frame.pack(fill="x", pady=(0, 12))

        preset_row = tk.Frame(preset_frame)
        preset_row.pack(fill="x")
        tk.Label(preset_row, text="预设方案:", font=("微软雅黑", 10), width=12, anchor="w").pack(side="left")
        self.rule_preset_combo = ttk.Combobox(
            preset_row,
            textvariable=self.rule_preset_name,
            state="readonly",
            values=[preset.name for preset in self.rule_presets],
            font=("微软雅黑", 10),
        )
        self.rule_preset_combo.pack(side="left", fill="x", expand=True, padx=5)
        self.rule_preset_combo.bind("<<ComboboxSelected>>", self._on_rule_preset_change)
        tk.Button(preset_row, text="应用预设", padx=10, command=self._apply_rule_preset).pack(side="right")

        self._sync_rule_preset_ui()
        tk.Label(
            preset_frame,
            textvariable=self.rule_preset_desc,
            font=("微软雅黑", 9),
            fg=self.palette["muted"],
            justify="left",
            wraplength=620,
            anchor="w",
        ).pack(fill="x", pady=(8, 0))

        naming_frame = tk.LabelFrame(left_panel, text="命名规则", padx=12, pady=10, font=("微软雅黑", 10, "bold"))
        naming_frame.pack(fill="x", pady=(0, 12))
        tk.Label(
            naming_frame,
            text="PDF 文件名以下划线 _ 分段（从0开始编号）\n例：dzfp_发票号码_公司名称_时间戳.pdf → 第0段=dzfp, 第1段=发票号码, 第2段=公司名称",
            font=("微软雅黑", 9),
            fg=self.palette["muted"],
            justify="left",
            anchor="w",
            wraplength=620,
        ).pack(anchor="w", pady=(0, 8))

        for label, var in [
            ("公司名称所在段（从0开始）:", self.company_name_index),
            ("发票号码所在段（从0开始）:", self.invoice_number_index),
        ]:
            row = tk.Frame(naming_frame)
            row.pack(fill="x", pady=4)
            tk.Label(row, text=label, font=("微软雅黑", 10), width=25, anchor="w").pack(side="left")
            tk.Spinbox(row, from_=0, to=10, width=5, font=("微软雅黑", 11), textvariable=var).pack(side="left", padx=5)

        excel_frame = tk.LabelFrame(left_panel, text="Excel 识别", padx=12, pady=10, font=("微软雅黑", 10, "bold"))
        excel_frame.pack(fill="x", pady=(0, 12))
        invoice_alias_row = tk.Frame(excel_frame)
        invoice_alias_row.pack(fill="x")
        tk.Label(invoice_alias_row, text="发票列别名（逗号分隔）:", font=("微软雅黑", 10), width=25, anchor="w").pack(side="left")
        tk.Entry(invoice_alias_row, textvariable=self.invoice_column_aliases, font=("微软雅黑", 10)).pack(
            side="left", fill="x", expand=True, padx=5
        )

        company_alias_row = tk.Frame(excel_frame)
        company_alias_row.pack(fill="x", pady=(6, 0))
        tk.Label(company_alias_row, text="公司列别名（逗号分隔）:", font=("微软雅黑", 10), width=25, anchor="w").pack(side="left")
        tk.Entry(company_alias_row, textvariable=self.company_column_aliases, font=("微软雅黑", 10)).pack(
            side="left", fill="x", expand=True, padx=5
        )
        tk.Label(
            excel_frame,
            text="示例：发票列可填 票号, 发票编码, 销项发票号码；公司列可填 客户名称, 购方名称, 单位名称。",
            font=("微软雅黑", 9),
            fg=self.palette["muted"],
            justify="left",
            wraplength=620,
            anchor="w",
        ).pack(anchor="w", pady=(8, 0))

        action_frame = tk.LabelFrame(left_panel, text="保存与环境", padx=12, pady=10, font=("微软雅黑", 10, "bold"))
        action_frame.pack(fill="x")
        button_row = tk.Frame(action_frame)
        button_row.pack(fill="x")
        save_btn = tk.Button(
            button_row,
            text="💾 保存设置",
            font=("微软雅黑", 11, "bold"),
            padx=20,
            pady=6,
            command=self._save_settings,
        )
        save_btn.pack(side="left")
        self.settings_save_btn = save_btn
        self._style_action_button(save_btn, "success")
        tk.Button(button_row, text="打开配置目录", font=("微软雅黑", 9), padx=12, command=self._open_config_directory).pack(side="left", padx=(8, 0))
        tk.Button(button_row, text="打开日志文件", font=("微软雅黑", 9), padx=12, command=self._open_log_file).pack(side="left", padx=(8, 0))

        tk.Label(action_frame, text=f"配置目录：{CONFIG_DIR}", font=("微软雅黑", 8), fg=self.palette["muted"], anchor="w", justify="left").pack(fill="x", pady=(10, 0))
        tk.Label(action_frame, text=f"日志文件：{LOG_FILE}", font=("微软雅黑", 8), fg=self.palette["muted"], anchor="w", justify="left").pack(fill="x", pady=(4, 0))
        tk.Label(
            action_frame,
            text=(
                f"能力状态：拖拽 {'已启用' if DND_SUPPORT else '未启用'} | "
                f"Excel {'已启用' if PANDAS_SUPPORT else '未启用'} | "
                f"报告 {'已启用' if OPENPYXL_SUPPORT else '未启用'}"
            ),
            font=("微软雅黑", 8),
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(4, 0))

        tk.Label(right_panel, text="🩺 诊断中心", font=("微软雅黑", 11, "bold"), fg=self.palette["text"]).pack(anchor="w", pady=(0, 12))

        recent_frame = tk.LabelFrame(right_panel, text="最近错误", padx=10, pady=10, font=("微软雅黑", 10, "bold"))
        recent_frame.pack(fill="both", expand=True)

        top_bar = tk.Frame(recent_frame)
        top_bar.pack(fill="x", pady=(0, 8))
        tk.Label(top_bar, textvariable=self.recent_error_summary_var, font=("微软雅黑", 9), fg=self.palette["muted"]).pack(side="left")
        tk.Button(top_bar, text="复制", font=("微软雅黑", 8), padx=8, command=self._copy_selected_recent_error).pack(side="right")
        tk.Button(top_bar, text="清空", font=("微软雅黑", 8), padx=8, command=self._clear_recent_errors).pack(side="right", padx=(0, 6))

        list_frame = tk.Frame(recent_frame)
        list_frame.pack(fill="both", expand=True)
        recent_scroll = tk.Scrollbar(list_frame)
        recent_scroll.pack(side="right", fill="y")
        self.recent_error_listbox = tk.Listbox(list_frame, font=("Consolas", 9), yscrollcommand=recent_scroll.set, height=10)
        self.recent_error_listbox.pack(side="left", fill="both", expand=True)
        recent_scroll.config(command=self.recent_error_listbox.yview)
        self.recent_error_listbox.bind("<<ListboxSelect>>", self._on_recent_error_select)

        detail_frame = tk.LabelFrame(recent_frame, text="错误详情", padx=8, pady=8, font=("微软雅黑", 9, "bold"))
        detail_frame.pack(fill="x", pady=(8, 0))
        tk.Label(
            detail_frame,
            textvariable=self.recent_error_detail_var,
            font=("微软雅黑", 9),
            fg=self.palette["detail_fg"],
            justify="left",
            wraplength=310,
            anchor="w",
        ).pack(fill="x")

        self._refresh_recent_error_list()

    def _save_settings(self) -> None:
        self._save_config()
        idx = self.company_name_index.get()
        self.organize_hint.config(text=f"  💡 公司名在第{idx+1}段（可在设置中修改）")
        self._refresh_excel_sheets(silent=True)
        logger.info("✅ 设置已保存")
        messagebox.showinfo("提示", "设置已保存，请重新扫描文件使新设置生效。")

    # ─────────────── 拖拽 ───────────────

    def _setup_drag_and_drop(self) -> None:
        if DND_SUPPORT:
            try:
                self.organize_folder_entry.drop_target_register(DND_FILES)
                self.organize_folder_entry.dnd_bind("<<Drop>>", self._on_drop)
                logger.info("✅ 拖拽功能已启用")
            except Exception as e:
                logger.warning(f"拖拽初始化失败：{e}（不影响其他功能）")
        else:
            logger.warning("拖拽未启用（需 tkinterdnd2）")

    def _on_drop(self, event) -> None:
        paths = []
        for m in re.finditer(r"\{([^}]+)\}|(\S+)", event.data):
            p = (m.group(1) or m.group(2)).strip("\"'")
            paths.append(p)
        if not paths:
            return
        p = Path(paths[0])
        folder = p if p.is_dir() else p.parent
        self.organize_folder_path.set(str(folder))
        logger.info(f"📂 拖入文件夹：{folder}")
        self._scan_files()

    # ─────────────── 进度工具 ───────────────

    def _update_progress_info(self, cur: int, total: int) -> None:
        pct = cur * 100 // max(total, 1)
        text = f"{cur}/{total} ({pct}%)"
        def _do():
            self.progress_label.config(text=text)
        if threading.current_thread() is threading.main_thread():
            _do()
        else:
            self.root.after(0, _do)

    def _update_progress(self, bar: ttk.Progressbar, value: int, maximum: Optional[int] = None) -> None:
        def _do():
            if maximum is not None:
                bar["maximum"] = maximum
            bar["value"] = value
        if threading.current_thread() is threading.main_thread():
            _do()
        else:
            self.root.after(0, _do)

    def _try_begin_task(
        self,
        start_btn: tk.Button,
        busy_text: str,
        cancel_btn: tk.Button,
        busy_bg: Optional[str] = None,
    ) -> bool:
        with self._lock:
            if self.is_running:
                return False
            self.is_running = True
        start_btn.config(state="disabled", text=busy_text)
        if busy_bg is not None:
            start_btn.config(bg=busy_bg, activebackground=busy_bg)
        cancel_btn.config(state="normal")
        self.status_var.set("⏳ 任务进行中...")
        return True

    def _finish_task_ui(
        self,
        start_btn: tk.Button,
        idle_text: str,
        cancel_btn: tk.Button,
        progress_bar: ttk.Progressbar,
        idle_bg: Optional[str] = None,
    ) -> None:
        with self._lock:
            self.is_running = False
        start_btn.config(state="normal", text=idle_text)
        if idle_bg is not None:
            start_btn.config(bg=idle_bg, activebackground=idle_bg)
        cancel_btn.config(state="disabled")
        progress_bar["value"] = 0
        self.progress_label.config(text="")
        if self.status_var.get() == "⏳ 任务进行中...":
            self.status_var.set("就绪 - 请选择功能开始使用")

    def _cancel_task(self) -> None:
        self._cancel_flag.set()
        logger.warning("⏹ 用户请求取消任务")

    @staticmethod
    def _open_folder(folder: Path) -> None:
        InvoiceToolApp._open_path_in_shell(folder)

    # ─────────────── 整理功能 ───────────────

    def _browse_organize_folder(self) -> None:
        initial = self.organize_folder_path.get() or ""
        d = filedialog.askdirectory(title="选择发票文件夹", initialdir=initial)
        if d:
            self.organize_folder_path.set(d)
            self.config["organize_folder"] = d
            self._save_config()
            logger.info(f"📂 已选择：{d}")
            self._scan_files()

    def _scan_files(self) -> None:
        folder_str = self.organize_folder_path.get().strip()
        if not folder_str:
            messagebox.showwarning("提示", "请先选择文件夹")
            return
        folder = Path(folder_str)
        if not folder.exists():
            messagebox.showerror("错误", "文件夹不存在")
            return

        self.config["organize_folder"] = folder_str
        self._save_config()

        self.file_tree.delete(*self.file_tree.get_children())
        self.file_check_vars.clear()
        self.preview_data.clear()

        cidx = self.company_name_index.get()
        pdfs = InvoiceOrganizer.scan_pdf_files(folder, self.organize_recursive.get())

        if not pdfs:
            logger.warning("📭 未找到PDF文件")
            self._update_file_count()
            return

        logger.info(f"🔍 扫描到 {len(pdfs)} 个PDF文件")

        for i, f in enumerate(pdfs):
            fname = str(f)
            company, valid = InvoiceOrganizer.parse_filename(
                fname,
                cidx,
                filename_parser=self._get_filename_parser(),
            )
            tgt = company if valid else "-"
            self.preview_data[fname] = {"filename": fname, "company": company, "target": tgt, "valid": valid}
            self.file_check_vars[fname] = tk.BooleanVar(value=valid)
            status = "✓" if valid else "✗"
            tag = ("evenrow" if i % 2 == 0 else "oddrow") if valid else ("invalid_even" if i % 2 == 0 else "invalid")
            self.file_tree.insert("", "end", values=(status, fname, company, tgt), tags=(tag,))

        self._update_file_count()
        self.status_var.set(f"✅ 已扫描 {len(pdfs)} 个文件")

    def _render_organize_preview(self) -> None:
        if not hasattr(self, "file_tree"):
            return
        self.file_tree.delete(*self.file_tree.get_children())
        if not self.preview_data:
            self._update_file_count()
            return
        for index, fname in enumerate(self.preview_data.keys()):
            data = self.preview_data[fname]
            if fname not in self.file_check_vars:
                self.file_check_vars[fname] = tk.BooleanVar(value=bool(data.get("valid")))
            status = "✓" if self.file_check_vars[fname].get() else "✗"
            tag = ("evenrow" if index % 2 == 0 else "oddrow") if data["valid"] else ("invalid_even" if index % 2 == 0 else "invalid")
            self.file_tree.insert("", "end", values=(status, fname, data["company"], data["target"]), tags=(tag,))
        self._update_file_count()

    def _on_tree_click(self, event) -> None:
        if self.file_tree.identify("region", event.x, event.y) == "cell":
            col = self.file_tree.identify_column(event.x)
            item = self.file_tree.identify_row(event.y)
            if item and col == "#1":
                vals = list(self.file_tree.item(item, "values"))
                fn = vals[1]
                if fn in self.file_check_vars:
                    cur = self.file_check_vars[fn].get()
                    self.file_check_vars[fn].set(not cur)
                    vals[0] = "✓" if not cur else "✗"
                    self.file_tree.item(item, values=vals)
                    self._update_file_count()

    def _update_file_count(self) -> None:
        t = len(self.file_check_vars)
        s = sum(1 for v in self.file_check_vars.values() if v.get())
        self.file_count_label.config(text=f"已选择: {s} / {t}")

    def _select_all(self) -> None:
        for item in self.file_tree.get_children():
            vals = list(self.file_tree.item(item, "values"))
            fn = vals[1]
            d = self.preview_data.get(fn)
            if d and d["valid"]:
                self.file_check_vars[fn].set(True)
                vals[0] = "✓"
                self.file_tree.item(item, values=vals)
        self._update_file_count()

    def _deselect_all(self) -> None:
        for item in self.file_tree.get_children():
            vals = list(self.file_tree.item(item, "values"))
            fn = vals[1]
            if fn in self.file_check_vars:
                self.file_check_vars[fn].set(False)
                vals[0] = "✗"
                self.file_tree.item(item, values=vals)
        self._update_file_count()

    def _execute_organize(self) -> None:
        sel = [f for f, v in self.file_check_vars.items() if v.get()]
        if not sel:
            messagebox.showwarning("提示", "请至少选择一个文件")
            return
        if not messagebox.askyesno("确认", f"确定整理 {len(sel)} 个文件？\n文件将被移动到对应公司文件夹。"):
            return
        if not self._try_begin_task(self.start_btn, "⏳ 处理中...", self.cancel_org_btn, busy_bg=self.palette["secondary"]):
            messagebox.showwarning("提示", "任务进行中...")
            return
        self._cancel_flag.clear()
        try:
            threading.Thread(target=self._do_organize, args=(sel,), daemon=True).start()
        except Exception:
            self._finish_task_ui(
                self.start_btn,
                "🚀 执行整理",
                self.cancel_org_btn,
                self.organize_progress,
                idle_bg=self.palette["success"],
            )
            raise

    def _do_organize(self, files: List[str]) -> None:
        folder = Path(self.organize_folder_path.get())

        try:
            def on_progress(current: int, total: int) -> None:
                self._update_progress(self.organize_progress, current, total if current == 0 else None)
                if current == 0:
                    self._update_progress_info(0, total)
                elif current % 5 == 0 or current == total:
                    self._update_progress_info(current, total)

            result = OrganizeService.run(
                folder=folder,
                files=files,
                preview_data=self.preview_data,
                progress_callback=on_progress,
                cancel_requested=self._cancel_flag.is_set,
            )
            final_m = result.moves

            def finish():
                if final_m:
                    self.current_session_history = final_m
                    self._save_to_history(final_m, "整理")
                    self.undo_btn.config(state="normal")
                    self.undo_all_btn.config(state="normal")
                self.status_var.set(f"✅ 成功 {result.success_count} | 失败 {result.fail_count} | {result.elapsed:.1f}秒")
                messagebox.showinfo(
                    "完成",
                    f"整理完成！\n✅ 成功：{result.success_count}\n❌ 失败：{result.fail_count}",
                )
                self._scan_files()
                self._finish_task_ui(
                    self.start_btn,
                    "🚀 执行整理",
                    self.cancel_org_btn,
                    self.organize_progress,
                    idle_bg=self.palette["success"],
                )
            self.root.after(0, finish)

        except Exception as e:
            logger.exception("整理异常")
            msg = str(e)
            def err():
                messagebox.showerror("错误", msg)
                self._finish_task_ui(
                    self.start_btn,
                    "🚀 执行整理",
                    self.cancel_org_btn,
                    self.organize_progress,
                    idle_bg=self.palette["success"],
                )
            self.root.after(0, err)

    # ─── 撤销 ───

    def _undo_last_move(self) -> None:
        if not self.current_session_history:
            messagebox.showinfo("提示", "没有可撤销的操作")
            return
        last = self.current_session_history[-1]
        ok, err = InvoiceOrganizer.rollback_single_move(last)
        if ok:
            self.current_session_history.pop()
            logger.info(f"↩️ 已撤销：{last['filename']}")
            if self.all_history and self.all_history[0].get("type") == "整理":
                rec = self.all_history[0]
                rec["moves"] = [m for m in rec["moves"] if m["filename"] != last["filename"]]
                rec["count"] = len(rec["moves"])
                if rec["count"] == 0:
                    self.all_history.pop(0)
                self._save_history()
                self._refresh_history_tree()
            self._scan_files()
        else:
            logger.warning(err)
        if not self.current_session_history:
            self.undo_btn.config(state="disabled")
            self.undo_all_btn.config(state="disabled")

    def _undo_all_moves(self) -> None:
        if not self.current_session_history:
            messagebox.showinfo("提示", "无可撤销操作")
            return
        if not messagebox.askyesno("确认", f"撤销全部 {len(self.current_session_history)} 个操作？"):
            return
        ok_n = fail_n = 0
        failed: List[Dict] = []
        for m in reversed(self.current_session_history.copy()):
            ok, err = InvoiceOrganizer.rollback_single_move(m)
            if ok:
                ok_n += 1
            else:
                logger.error(err)
                fail_n += 1
                failed.append(m)
        failed.reverse()
        self.current_session_history = failed
        self.undo_btn.config(state="normal" if failed else "disabled")
        self.undo_all_btn.config(state="normal" if failed else "disabled")
        if self.all_history and self.all_history[0].get("type") == "整理":
            if fail_n == 0:
                self.all_history.pop(0)
            else:
                self.all_history[0]["moves"] = failed
                self.all_history[0]["count"] = len(failed)
            self._save_history()
            self._refresh_history_tree()
        logger.info(f"↩️ 批量撤销：成功 {ok_n} 失败 {fail_n}")
        self._scan_files()
        messagebox.showinfo("完成", f"成功 {ok_n} | 失败 {fail_n}" + (f"\n{fail_n}个失败记录已保留" if fail_n else ""))

    # ─────────────── 筛选功能 ───────────────

    def _toggle_help(self) -> None:
        if self.help_visible.get():
            self.help_content.pack_forget()
            self.help_btn.config(text="📖 显示使用说明")
            self.help_visible.set(False)
        else:
            self.help_content.pack(fill="x", pady=8, before=self.file_path_frame)
            self.help_btn.config(text="📖 隐藏使用说明")
            self.help_visible.set(True)

    def _refresh_excel_sheets(self, silent: bool = False) -> None:
        if not PANDAS_SUPPORT or not hasattr(self, "excel_sheet_combo"):
            return
        excel = self.excel_path.get().strip()
        if not excel:
            self.excel_sheet_combo["values"] = ()
            self.excel_sheet_name.set("")
            self._clear_workbook_analysis("打开 Excel 后，会自动分析每个工作表的发票列和公司列候选。")
            self._sync_output_folder_mode_ui()
            return
        excel_path = Path(excel)
        if not excel_path.exists():
            self.excel_sheet_combo["values"] = ()
            self._clear_workbook_analysis("Excel 文件不存在，无法分析工作簿。")
            self._sync_output_folder_mode_ui()
            return

        try:
            sheets = InvoiceFilter.list_excel_sheets(excel_path)
        except (FileNotFoundError, PermissionError, ValueError) as e:
            self.excel_sheet_combo["values"] = ()
            self._clear_workbook_analysis(f"工作簿分析失败：{e}")
            self._sync_output_folder_mode_ui()
            if not silent:
                messagebox.showerror("错误", str(e))
            return

        self.excel_sheet_combo["values"] = sheets
        current = self.excel_sheet_name.get().strip()
        if current not in sheets:
            self.excel_sheet_name.set(sheets[0])
        self._sync_filter_context(self.excel_sheet_name.get().strip())
        self._refresh_workbook_analysis(silent=silent)
        self._sync_output_folder_mode_ui()
        self._save_config()

    def _on_excel_sheet_change(self, event=None) -> None:
        self._sync_filter_context(self.excel_sheet_name.get().strip())
        self._sync_analysis_selection_to_current_sheet()
        self._sync_output_folder_mode_ui()
        self._save_config()

    def _get_invoice_aliases(self) -> List[str]:
        preset = self._get_rule_preset()
        custom_aliases = InvoiceFilter.parse_aliases(self.invoice_column_aliases.get())
        merged = list(dict.fromkeys(list(preset.invoice_column_aliases) + custom_aliases))
        return merged

    def _get_company_aliases(self) -> List[str]:
        return InvoiceFilter.parse_aliases(self.company_column_aliases.get())

    def _get_filter_exclude_dirs(self) -> List[Path]:
        output_path = self._get_effective_output_folder_path()
        if output_path is None:
            return []
        return [output_path]

    def _clear_workbook_analysis(self, message: str) -> None:
        self.workbook_analysis_result = None
        self.workbook_profiles = {}
        self.workbook_tree_selection.clear()
        self._reset_sheet_row_filters()
        self._active_filter_context = ("", "")
        self.workbook_analysis_summary_var.set(message)
        self.workbook_sheet_overview_var.set("先选择 Excel 文件，再从左侧查看每个 sheet 的识别结果。")
        self.workbook_sheet_sample_var.set("样本预览会显示当前工作表前几行数据，便于确认列是否正确。")
        self.selected_invoice_column_name.set("")
        self.selected_company_column_name.set("")
        if hasattr(self, "workbook_sheet_tree"):
            self.workbook_sheet_tree.delete(*self.workbook_sheet_tree.get_children())
        if hasattr(self, "analysis_invoice_combo"):
            self.analysis_invoice_combo["values"] = ()
        if hasattr(self, "analysis_company_combo"):
            self.analysis_company_combo["values"] = ()

    def _render_workbook_analysis(self, result: WorkbookAnalysisResult) -> None:
        if not hasattr(self, "workbook_sheet_tree"):
            return

        self.workbook_sheet_tree.delete(*self.workbook_sheet_tree.get_children())
        self.workbook_tree_selection.clear()
        self.workbook_profiles = {profile.sheet_name: profile for profile in result.sheet_profiles}

        for index, profile in enumerate(result.sheet_profiles):
            invoice_name = profile.selected_invoice_column or "-"
            company_name = profile.selected_company_column or "-"
            if profile.recommended:
                status = "推荐"
                status_tag = "recommended"
            elif profile.usable:
                status = "可用"
                status_tag = "usable"
            elif profile.issue:
                status = profile.issue
                status_tag = "warning" if "公司列" in profile.issue else "error"
            else:
                status = "待确认"
                status_tag = "warning"

            item_id = self.workbook_sheet_tree.insert(
                "",
                "end",
                values=(
                    profile.sheet_name,
                    f"{profile.row_count}/{profile.column_count}",
                    invoice_name,
                    company_name,
                    status,
                ),
                tags=(("evenrow" if index % 2 == 0 else "oddrow"), status_tag),
            )
            self.workbook_tree_selection[item_id] = profile.sheet_name

        summary = (
            f"已分析 {result.total_sheet_count} 个工作表，可用于筛选 {result.usable_sheet_count} 个。"
            f"推荐工作表：{result.recommended_sheet_name or '未识别'}。"
        )
        self.workbook_analysis_summary_var.set(summary)

    def _format_sheet_sample_text(self, profile: WorkbookSheetProfile) -> str:
        if not profile.sample_rows:
            return "当前工作表没有可展示的样本数据。"

        lines: List[str] = []
        for row in profile.sample_rows:
            parts = [f"{key}={value}" for key, value in row.items() if value]
            lines.append(" | ".join(parts) if parts else "（空行）")
        return "\n".join(lines)

    def _populate_workbook_sheet_detail(self, sheet_name: str) -> None:
        profile = self.workbook_profiles.get(sheet_name)
        if profile is None:
            self.workbook_sheet_overview_var.set("先选择 Excel 文件，再从左侧查看每个 sheet 的识别结果。")
            self.workbook_sheet_sample_var.set("样本预览会显示当前工作表前几行数据，便于确认列是否正确。")
            if hasattr(self, "analysis_invoice_combo"):
                self.analysis_invoice_combo["values"] = ()
            if hasattr(self, "analysis_company_combo"):
                self.analysis_company_combo["values"] = ()
            if hasattr(self, "row_filter_column_combo"):
                self.row_filter_column_combo["values"] = ()
            return

        invoice_values = [candidate.column_name for candidate in profile.invoice_candidates]
        company_values = [candidate.column_name for candidate in profile.company_candidates]
        if hasattr(self, "analysis_invoice_combo"):
            self.analysis_invoice_combo["values"] = invoice_values
        if hasattr(self, "analysis_company_combo"):
            self.analysis_company_combo["values"] = company_values
        if hasattr(self, "row_filter_column_combo"):
            self.row_filter_column_combo["values"] = [""] + profile.columns

        invoice_candidate_text = "、".join(invoice_values[:3]) if invoice_values else "未识别到发票列"
        company_candidate_text = "、".join(company_values[:3]) if company_values else "未识别到公司列"
        active_filter_text = self._describe_active_row_filters()
        status = "推荐用于筛选" if profile.recommended else ("可用于筛选" if profile.usable else (profile.issue or "待确认"))
        self.workbook_sheet_overview_var.set(
            f"工作表：{profile.sheet_name}\n"
            f"规模：{profile.row_count} 行 / {profile.column_count} 列\n"
            f"状态：{status}\n"
            f"发票列候选：{invoice_candidate_text}\n"
            f"公司列候选：{company_candidate_text}\n"
            f"当前条件：{active_filter_text}"
        )
        self.workbook_sheet_sample_var.set(self._format_sheet_sample_text(profile))

        self.selected_invoice_column_name.set(profile.selected_invoice_column)
        self.selected_company_column_name.set(profile.selected_company_column)
        if self.row_filter_column_name.get().strip() and self.row_filter_column_name.get().strip() not in profile.columns:
            self.row_filter_column_name.set("")

    def _select_workbook_tree_item(self, sheet_name: str) -> None:
        if not hasattr(self, "workbook_sheet_tree"):
            return
        for item_id, mapped_sheet_name in self.workbook_tree_selection.items():
            if mapped_sheet_name == sheet_name:
                self.workbook_sheet_tree.selection_set(item_id)
                self.workbook_sheet_tree.focus(item_id)
                self.workbook_sheet_tree.see(item_id)
                break

    def _sync_analysis_selection_to_current_sheet(self) -> None:
        current_sheet = self.excel_sheet_name.get().strip()
        if not current_sheet:
            return
        self._select_workbook_tree_item(current_sheet)
        self._populate_workbook_sheet_detail(current_sheet)

    def _refresh_workbook_analysis(self, silent: bool = False) -> None:
        if not PANDAS_SUPPORT:
            return
        excel = self.excel_path.get().strip()
        if not excel:
            self._clear_workbook_analysis("打开 Excel 后，会自动分析每个工作表的发票列和公司列候选。")
            return

        excel_path = Path(excel)
        if not excel_path.exists():
            self._clear_workbook_analysis("Excel 文件不存在，无法分析工作簿。")
            return

        try:
            result = WorkbookAnalyzerService.analyze(
                excel_path,
                extra_invoice_aliases=self._get_invoice_aliases(),
                extra_company_aliases=self._get_company_aliases(),
            )
        except (FileNotFoundError, PermissionError, ValueError) as exc:
            self._clear_workbook_analysis(f"工作簿分析失败：{exc}")
            if not silent:
                messagebox.showerror("错误", str(exc))
            return
        except Exception as exc:
            logger.exception("工作簿分析出错")
            self._clear_workbook_analysis(f"工作簿分析失败：{exc}")
            if not silent:
                messagebox.showerror("错误", str(exc))
            return

        self.workbook_analysis_result = result
        self._render_workbook_analysis(result)

        current_sheet = self.excel_sheet_name.get().strip()
        if current_sheet not in self.workbook_profiles:
            current_sheet = result.recommended_sheet_name
            if current_sheet:
                self.excel_sheet_name.set(current_sheet)
        self._sync_filter_context(current_sheet)
        profile = self.workbook_profiles.get(current_sheet)
        if profile is not None:
            saved_invoice = self.selected_invoice_column_name.get().strip()
            saved_company = self.selected_company_column_name.get().strip()
            if saved_invoice and saved_invoice in [item.column_name for item in profile.invoice_candidates]:
                profile.selected_invoice_column = saved_invoice
            if saved_company and saved_company in [item.column_name for item in profile.company_candidates]:
                profile.selected_company_column = saved_company
        self._sync_analysis_selection_to_current_sheet()
        self._save_config()

    def _on_workbook_sheet_select(self, event=None) -> None:
        if not hasattr(self, "workbook_sheet_tree"):
            return
        selection = self.workbook_sheet_tree.selection()
        if not selection:
            return
        sheet_name = self.workbook_tree_selection.get(selection[0], "")
        if not sheet_name:
            return
        self.excel_sheet_name.set(sheet_name)
        self._sync_filter_context(sheet_name)
        self._populate_workbook_sheet_detail(sheet_name)
        self._sync_output_folder_mode_ui()
        self._save_config()

    def _on_analysis_invoice_column_change(self, event=None) -> None:
        sheet_name = self.excel_sheet_name.get().strip()
        profile = self.workbook_profiles.get(sheet_name)
        if profile is None or self.workbook_analysis_result is None:
            return
        profile.selected_invoice_column = self.selected_invoice_column_name.get().strip()
        self._render_workbook_analysis(self.workbook_analysis_result)
        self._select_workbook_tree_item(sheet_name)
        self._save_config()

    def _on_analysis_company_column_change(self, event=None) -> None:
        sheet_name = self.excel_sheet_name.get().strip()
        profile = self.workbook_profiles.get(sheet_name)
        if profile is None or self.workbook_analysis_result is None:
            return
        profile.selected_company_column = self.selected_company_column_name.get().strip()
        self._render_workbook_analysis(self.workbook_analysis_result)
        self._select_workbook_tree_item(sheet_name)
        self._save_config()

    def _describe_active_row_filters(self) -> str:
        parts: List[str] = []
        filter_column = self.row_filter_column_name.get().strip()
        filter_mode = self.row_filter_mode.get().strip()
        filter_values = self.row_filter_values.get().strip()
        company_excludes = self.company_exclude_keywords.get().strip()
        if filter_column and filter_mode and filter_mode != "不过滤" and filter_values:
            parts.append(f"{filter_column} {filter_mode} {filter_values}")
        if company_excludes:
            parts.append(f"排除公司: {company_excludes}")
        return "；".join(parts) if parts else "不过滤"

    def _on_row_filter_rule_change(self, event=None) -> None:
        current_sheet = self.excel_sheet_name.get().strip()
        if current_sheet:
            self._populate_workbook_sheet_detail(current_sheet)
        self._save_config()

    def _browse_excel(self) -> None:
        ini = str(Path(self.excel_path.get()).parent) if self.excel_path.get() else ""
        fp = filedialog.askopenfilename(title="选择Excel", initialdir=ini, filetypes=[("Excel", "*.xlsx *.xls")])
        if fp:
            self.excel_path.set(fp)
            self.config["excel_path"] = fp
            self._save_config()
            self._refresh_excel_sheets(silent=True)
            self._sync_output_folder_mode_ui()
            logger.info(f"📄 已选择Excel：{fp}")

    def _browse_pdf_folder(self) -> None:
        d = filedialog.askdirectory(title="选择PDF文件夹", initialdir=self.pdf_folder.get() or "")
        if d:
            self.pdf_folder.set(d)
            self.config["pdf_folder"] = d
            self._save_config()
            cnt = sum(1 for _ in Path(d).glob("*.pdf"))
            logger.info(f"📂 已选择PDF文件夹：{d}（{cnt}个PDF）")

    def _browse_output_folder(self) -> None:
        initial_dir = self.manual_output_folder.get().strip() or self.output_folder.get().strip()
        d = filedialog.askdirectory(title="选择导出文件夹", initialdir=initial_dir)
        if d:
            self.manual_output_folder.set(d)
            self.output_folder.set(d)
            self._save_config()
            logger.info(f"📂 已选择导出文件夹：{d}")

    def _open_output_folder(self) -> None:
        target = self._get_effective_output_folder_path()
        if target and target.exists():
            self._open_folder(target)
        else:
            messagebox.showwarning("提示", "请先选择有效的导出文件夹")

    def _validate_filter_paths(self) -> Optional[Tuple[Path, Path, Path]]:
        excel = self.excel_path.get()
        pdf = self.pdf_folder.get()
        if not excel or not Path(excel).exists():
            messagebox.showerror("错误", "请选择有效的Excel文件")
            return None
        if not pdf or not Path(pdf).exists():
            messagebox.showerror("错误", "请选择有效的PDF文件夹")
            return None
        out_path_raw = self._get_effective_output_folder_path()
        if out_path_raw is None:
            messagebox.showerror("错误", "请选择导出文件夹")
            return None
        pdf_path = Path(pdf).resolve()
        out_path = out_path_raw.resolve()
        if pdf_path == out_path:
            messagebox.showerror("错误", "导出文件夹不能与PDF源文件夹相同！")
            return None
        if self.filter_recursive.get() and is_relative_to(out_path, pdf_path):
            messagebox.showerror("错误", "递归筛选时，导出文件夹不能位于PDF源文件夹内部！")
            return None
        return Path(excel), Path(pdf), out_path

    def _preview_filter(self) -> None:
        paths = self._validate_filter_paths()
        if not paths:
            return
        excel_p, pdf_p, out_p = paths
        selected_company_column = self.selected_company_column_name.get().strip()
        active_filter_desc = self._describe_active_row_filters()
        try:
            preview = FilterService.preview(
                excel_p,
                pdf_p,
                self.invoice_number_index.get(),
                recursive=self.filter_recursive.get(),
                sheet_name=self.excel_sheet_name.get(),
                invoice_column_name=self.selected_invoice_column_name.get().strip() or None,
                company_column_name=selected_company_column or None,
                filter_column_name=self.row_filter_column_name.get().strip() or None,
                filter_mode=self.row_filter_mode.get().strip() or "不过滤",
                filter_values=self.row_filter_values.get().strip() or None,
                company_exclude_keywords=self.company_exclude_keywords.get().strip() or None,
                extra_aliases=self._get_invoice_aliases(),
                exclude_dirs=[out_p] if self.filter_recursive.get() else None,
                filename_parser=self._get_filename_parser(),
                column_resolver=self._get_column_resolver(),
            )
            columns_preview = "、".join(preview.columns[:6])
            if len(preview.columns) > 6:
                columns_preview += f" ... 共{len(preview.columns)}列"
            self._update_filter_summary(
                "预览完成",
                f"工作表：{preview.sheet_name} | 发票列：{preview.excel_column_name} | 公司列：{selected_company_column or '未指定'} | 条件：{active_filter_desc} | 可用列：{columns_preview}",
                [
                    ("原始行数", str(preview.source_row_count)),
                    ("筛选后发票", str(len(preview.invoice_numbers))),
                    ("可匹配", str(len(preview.matched))),
                    ("已过滤行", str(preview.filtered_out_count)),
                    ("未匹配", str(len(preview.not_found))),
                    ("PDF扫描", str(preview.pdf_stats.scanned)),
                ],
            )
            self._set_filter_results(preview.result_rows, missing_invoices=preview.not_found)
            if preview.conflicts:
                self.filter_detail_var.set(f"检测到 {len(preview.conflicts)} 个重复冲突，可在表格中按“重复冲突”筛选查看。")
            elif preview.filtered_out_count:
                self.filter_detail_var.set(f"预览完成：已按条件过滤掉 {preview.filtered_out_count} 行，当前保留 {len(preview.invoice_numbers)} 个发票号。")
            elif preview.not_found:
                self.filter_detail_var.set(f"共有 {len(preview.not_found)} 个发票号未匹配，可直接复制未匹配发票号继续跟进。")
            else:
                self.filter_detail_var.set("预览完成：当前发票号均已找到对应 PDF，可直接开始筛选导出。")
            logger.info(
                f"👁 预览：工作表 {preview.sheet_name} | 筛选后 {len(preview.invoice_numbers)} | 过滤掉 {preview.filtered_out_count} | "
                f"匹配 {len(preview.matched)} | 未匹配 {len(preview.not_found)} | PDF扫描 {preview.pdf_stats.scanned}"
            )
        except (FileNotFoundError, PermissionError, ValueError) as e:
            messagebox.showerror("错误", str(e))
        except Exception as e:
            logger.exception("预览出错")
            messagebox.showerror("错误", str(e))

    def _run_filter(self) -> None:
        paths = self._validate_filter_paths()
        if not paths:
            return
        if not self._try_begin_task(self.filter_run_btn, "⏳ 处理中...", self.cancel_flt_btn, busy_bg=self.palette["secondary"]):
            messagebox.showwarning("提示", "任务进行中...")
            return
        self._cancel_flag.clear()
        try:
            threading.Thread(target=self._do_filter, daemon=True).start()
        except Exception:
            self._finish_task_ui(
                self.filter_run_btn,
                "🚀 开始筛选并导出",
                self.cancel_flt_btn,
                self.filter_progress,
                idle_bg=self.palette["primary"],
            )
            raise

    def _do_filter(self) -> None:
        self.root.after(
            0,
            lambda: (
                self._clear_filter_results(),
                self._update_filter_summary(
                    "正在筛选",
                    "正在根据 Excel 发票号匹配 PDF，请稍候。筛选完成后结果会显示在下方表格中。",
                    [
                        ("Excel发票", "-"),
                        ("已导出", "-"),
                        ("未匹配", "-"),
                        ("异常/冲突", "-"),
                        ("PDF扫描", "-"),
                        ("已跳过", "-"),
                    ],
                ),
            ),
        )

        excel_p = Path(self.excel_path.get())
        pdf_p = Path(self.pdf_folder.get())
        out_p = Path(self.output_folder.get())
        selected_company_column = self.selected_company_column_name.get().strip()
        active_filter_desc = self._describe_active_row_filters()

        try:
            def on_progress(current: int, total: int) -> None:
                self._update_progress(self.filter_progress, current, total if current == 0 else None)
                if current == 0:
                    self._update_progress_info(0, total)
                elif current % 10 == 0 or current == total:
                    self._update_progress_info(current, total)

            result = FilterService.run(
                excel_path=excel_p,
                pdf_folder=pdf_p,
                output_dir=out_p,
                invoice_index=self.invoice_number_index.get(),
                recursive=self.filter_recursive.get(),
                sheet_name=self.excel_sheet_name.get(),
                invoice_column_name=self.selected_invoice_column_name.get().strip() or None,
                company_column_name=selected_company_column or None,
                filter_column_name=self.row_filter_column_name.get().strip() or None,
                filter_mode=self.row_filter_mode.get().strip() or "不过滤",
                filter_values=self.row_filter_values.get().strip() or None,
                company_exclude_keywords=self.company_exclude_keywords.get().strip() or None,
                extra_aliases=self._get_invoice_aliases(),
                exclude_dirs=self._get_filter_exclude_dirs() if self.filter_recursive.get() else None,
                filename_parser=self._get_filename_parser(),
                column_resolver=self._get_column_resolver(),
                report_exporter=self._get_report_exporter(),
                progress_callback=on_progress,
                cancel_requested=self._cancel_flag.is_set,
            )
            report_files = [str(result.report_path)] if result.report_path else []

            def finish():
                if result.moves or report_files:
                    self._save_to_history(result.moves, "筛选", {"report_files": report_files})

                columns_preview = "、".join(result.columns[:6])
                if len(result.columns) > 6:
                    columns_preview += f" ... 共{len(result.columns)}列"
                title = "筛选已取消" if result.cancelled else "筛选完成"
                self._update_filter_summary(
                    title,
                    f"工作表：{result.sheet_name} | 发票列：{result.excel_column_name} | 公司列：{selected_company_column or '未指定'} | 条件：{active_filter_desc} | 可用列：{columns_preview}",
                    [
                        ("原始行数", str(result.source_row_count)),
                        ("筛选后发票", str(result.found_count + len(result.not_found) + result.skip_count + result.copy_fail_count)),
                        ("已导出", str(result.found_count)),
                        ("已过滤行", str(result.filtered_out_count)),
                        ("未匹配", str(len(result.not_found))),
                        ("PDF扫描", str(result.pdf_stats.scanned)),
                    ],
                )
                self._set_filter_results(result.result_rows, missing_invoices=result.not_found)
                if result.conflicts:
                    self.filter_detail_var.set(f"本次发现 {len(result.conflicts)} 个重复冲突，已在结果表格中标记为“重复冲突”。")
                elif result.filtered_out_count:
                    self.filter_detail_var.set(f"筛选完成：已按条件过滤掉 {result.filtered_out_count} 行，导出 {result.found_count} 个文件。")
                elif result.not_found:
                    self.filter_detail_var.set(f"本次有 {len(result.not_found)} 个发票号未匹配，可用“复制未匹配发票号”继续处理。")
                elif result.found_count > 0:
                    self.filter_detail_var.set("筛选完成：可双击结果表中的文件直接打开，或点击“打开导出文件夹”查看全部导出结果。")
                else:
                    self.filter_detail_var.set("本次没有匹配到可导出的文件，请检查 Excel 工作表、列名或 PDF 命名规则。")

                self.status_var.set(
                    f"✅ 成功: {result.found_count} | 跳过: {result.skip_count} | 复制失败: {result.copy_fail_count} | "
                    f"未找到: {len(result.not_found)} | {result.elapsed:.1f}秒"
                )

                report_msg = ""
                if report_files:
                    report_msg = f"\n\n📊 筛选报告已保存到导出文件夹"

                if result.cancelled:
                    messagebox.showinfo("已取消", f"已导出 {result.found_count} 个{report_msg}")
                elif result.found_count > 0:
                    if messagebox.askyesno("完成", f"成功导出 {result.found_count} 个文件！{report_msg}\n\n是否打开导出文件夹？"):
                        self._open_folder(out_p)
                else:
                    messagebox.showinfo("完成", f"无匹配文件。未找到: {len(result.not_found)}{report_msg}")

                self._finish_task_ui(
                    self.filter_run_btn,
                    "🚀 开始筛选并导出",
                    self.cancel_flt_btn,
                    self.filter_progress,
                    idle_bg=self.palette["primary"],
                )
            self.root.after(0, finish)

        except (FileNotFoundError, PermissionError, ValueError) as e:
            msg = str(e)
            def err():
                logger.error(msg)
                messagebox.showerror("错误", msg)
                self._finish_task_ui(
                    self.filter_run_btn,
                    "🚀 开始筛选并导出",
                    self.cancel_flt_btn,
                    self.filter_progress,
                    idle_bg=self.palette["primary"],
                )
            self.root.after(0, err)
        except Exception as e:
            logger.exception("筛选异常")
            msg = str(e)
            def err2():
                messagebox.showerror("错误", msg)
                self._finish_task_ui(
                    self.filter_run_btn,
                    "🚀 开始筛选并导出",
                    self.cancel_flt_btn,
                    self.filter_progress,
                    idle_bg=self.palette["primary"],
                )
            self.root.after(0, err2)

    # ─────────────── 历史记录 ───────────────

    def _save_to_history(
        self,
        moves: List[Dict],
        op: str = "整理",
        extra: Optional[Dict[str, Any]] = None,
    ) -> None:
        folder = self.organize_folder_path.get() if op == "整理" else self.pdf_folder.get()
        record = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "folder": folder, "count": len(moves), "type": op, "moves": moves,
        }
        if extra:
            record.update(extra)
        self.all_history.insert(0, record)
        self.all_history = self.all_history[:100]
        self._save_history()
        self._refresh_history_tree()

    def _on_history_filters_changed(self, event=None) -> None:
        self._refresh_history_tree()

    def _reset_history_filters(self) -> None:
        self.history_type_filter.set("全部")
        self.history_date_filter.set("全部")
        self.history_keyword.set("")
        self._refresh_history_tree()

    def _get_selected_history_index(self) -> Optional[int]:
        selection = self.history_tree.selection()
        if not selection:
            return None
        tree_index = self.history_tree.index(selection[0])
        if tree_index >= len(self.filtered_history_indices):
            return None
        return self.filtered_history_indices[tree_index]

    def _get_selected_history_record(self) -> Optional[Dict[str, Any]]:
        history_index = self._get_selected_history_index()
        if history_index is None:
            return None
        return self.all_history[history_index]

    def _refresh_history_tree(self) -> None:
        self.history_tree.delete(*self.history_tree.get_children())
        self.filtered_history_indices = filter_history_records(
            self.all_history,
            type_filter=self.history_type_filter.get(),
            date_filter=self.history_date_filter.get(),
            keyword=self.history_keyword.get(),
        )
        self.history_summary_var.set(f"显示 {len(self.filtered_history_indices)} / {len(self.all_history)} 条历史记录")

        for visible_index, history_index in enumerate(self.filtered_history_indices):
            r = self.all_history[history_index]
            fd = r["folder"]
            if len(fd) > 50:
                fd = "..." + fd[-47:]
            tag = "evenrow" if visible_index % 2 == 0 else "oddrow"
            report_count = len(r.get("report_files", []))
            count_text = f"{r['count']}个"
            if report_count:
                count_text += f" + {report_count}报告"
            self.history_tree.insert("", "end", values=(r["time"], fd, count_text, r.get("type", "整理")), tags=(tag,))

    def _view_history_detail(self) -> None:
        rec = self._get_selected_history_record()
        if rec is None:
            messagebox.showinfo("提示", "请先选择记录")
            return
        win = tk.Toplevel(self.root)
        win.title("历史详情")
        win.geometry("750x500")
        report_count = len(rec.get("report_files", []))
        count_desc = f"{rec['count']}"
        if report_count:
            count_desc += f"（另含 {report_count} 个报告）"
        for t in [f"时间：{rec['time']}", f"类型：{rec.get('type','整理')}", f"文件夹：{rec['folder']}", f"数量：{count_desc}"]:
            tk.Label(win, text=t, font=("微软雅黑", 10), wraplength=700).pack(anchor="w", padx=10)
        lf = tk.LabelFrame(win, text="文件列表", padx=10, pady=10)
        lf.pack(fill="both", expand=True, padx=10, pady=10)
        scr = tk.Scrollbar(lf)
        scr.pack(side="right", fill="y")
        lb = tk.Listbox(lf, font=("Consolas", 9), yscrollcommand=scr.set)
        lb.pack(fill="both", expand=True)
        scr.config(command=lb.yview)
        for m in rec["moves"]:
            lb.insert(tk.END, m["filename"])
        for report_file in rec.get("report_files", []):
            lb.insert(tk.END, f"[报告] {Path(report_file).name}")

    def _open_history_folder(self) -> None:
        rec = self._get_selected_history_record()
        if rec is None:
            messagebox.showinfo("提示", "请先选择记录")
            return
        folder = Path(rec["folder"])
        if not folder.exists():
            messagebox.showwarning("提示", f"文件夹不存在：\n{folder}")
            return
        self._open_folder(folder)

    def _rollback_selected(self) -> None:
        idx = self._get_selected_history_index()
        if idx is None:
            messagebox.showinfo("提示", "请先选择记录")
            return
        rec = self.all_history[idx]
        op = rec.get("type", "整理")

        if op == "筛选":
            report_files = rec.get("report_files", [])
            report_desc = f" 和 {len(report_files)} 个报告" if report_files else ""
            if not messagebox.askyesno("确认", f"回滚筛选？将删除导出的 {rec['count']} 个文件{report_desc}"):
                return
            ok_n = fail_n = 0
            failed = []
            for m in rec["moves"]:
                try:
                    t = Path(m["target"])
                    if t.exists():
                        t.unlink()
                        ok_n += 1
                    else:
                        fail_n += 1
                        failed.append(m)
                except (PermissionError, OSError) as e:
                    logger.error(f"❌ 删除失败：{m['filename']} - {e}")
                    fail_n += 1
                    failed.append(m)
            failed_reports: List[str] = []
            for report_path in report_files:
                try:
                    target = Path(report_path)
                    if target.exists():
                        target.unlink()
                        ok_n += 1
                    else:
                        fail_n += 1
                        failed_reports.append(report_path)
                except (PermissionError, OSError) as e:
                    logger.error(f"❌ 删除报告失败：{Path(report_path).name} - {e}")
                    fail_n += 1
                    failed_reports.append(report_path)
        else:
            if not messagebox.askyesno("确认", f"回滚整理？将移回 {rec['count']} 个文件"):
                return
            ok_n = fail_n = 0
            failed = []
            for m in reversed(rec["moves"]):
                ok, err = InvoiceOrganizer.rollback_single_move(m)
                if ok:
                    ok_n += 1
                else:
                    logger.error(err)
                    fail_n += 1
                    failed.append(m)
            failed_reports = []

        if fail_n == 0:
            self.all_history.pop(idx)
        else:
            rec["moves"] = failed
            rec["count"] = len(failed)
            if failed_reports:
                rec["report_files"] = failed_reports
            elif "report_files" in rec:
                rec.pop("report_files")

        self._save_history()
        self._refresh_history_tree()
        logger.info(f"↩️ 回滚：成功 {ok_n} 失败 {fail_n}")
        if op == "整理":
            self._scan_files()
        messagebox.showinfo("完成", f"成功 {ok_n} | 失败 {fail_n}" + (f"\n{fail_n}个失败记录已保留" if fail_n else ""))

    def _clear_all_history(self) -> None:
        if not self.all_history:
            messagebox.showinfo("提示", "已经是空的")
            return
        if messagebox.askyesno("确认", "清空所有历史？不影响已处理文件。"):
            self.all_history.clear()
            self._save_history()
            self._refresh_history_tree()
            logger.info("🗑️ 历史已清空")
