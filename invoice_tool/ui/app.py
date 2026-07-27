#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票处理工具箱 v6.1.1

当前版本聚焦于：
- 发票整理
- 多 Sheet 发票筛选
- 条件筛选与公司排除
- 白天 / 黑夜双主题
- 更紧凑的筛选页首屏体验
- 更稳定的 GUI 打包交付
"""

from __future__ import annotations

import logging
import platform
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from concurrent.futures import CancelledError
from dataclasses import asdict
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..application import (
    CONFIG_DEFAULTS,
    CONFIG_SCHEMA_VERSION,
    ConfigPlan,
    ConfigurationError,
    FilterExecutionRequest,
    FilterPreviewRequest,
    OrganizeExecutionRequest,
    OrganizePreviewRequest,
    WorkbookAnalysisRequest,
    backup_config,
    build_diagnostic_snapshot,
    create_diagnostic_bundle,
    default_config_plan,
    filter_filter_result_rows,
    filter_history_records,
    history_record_can_rerun,
    load_config_plan,
    normalize_config,
    save_config_export,
    sort_filter_result_rows,
)
from ..core.file_safety import fingerprint_file, has_valid_fingerprint
from ..core.filtering import InvoiceFilter
from ..core.models import (
    FilterPreviewResult,
    FilterResultRow,
    OrganizePreviewResult,
    WorkbookAnalysisResult,
    WorkbookSheetProfile,
)
from ..core.organizer import InvoiceOrganizer
from ..core.presets import DEFAULT_RULE_PRESET_ID, get_rule_preset, list_rule_presets
from ..core.services import FilterService, OrganizeService
from ..core.strategies import OpenpyxlFilterReportExporter, SegmentFilenameParser, SmartInvoiceColumnResolver
from ..core.workbook import WorkbookAnalyzerService
from ..infra.logging_setup import logger
from ..infra.paths import (
    ACTIVE_TASK_FILE,
    CONFIG_DIR,
    CONFIG_DIR_FALLBACK_REASON,
    CONFIG_FILE,
    HISTORY_FILE,
    LOG_FILE,
    is_relative_to,
)
from ..infra.privacy import RedactingFormatter
from ..infra.storage import load_json, save_json
from ..infra.task_journal import TaskJournal
from ..runtime import (
    DND_FILES,
    DND_SUPPORT,
    MODERN_UI,
    OPENPYXL_SUPPORT,
    PANDAS_SUPPORT,
    XLRD_SUPPORT,
    ttkb,
)
from ..version import APP_VERSION, RELEASE_SUMMARY
from .logging_handler import RecentErrorHandler, TkTextHandler


FILTER_RESULT_STATUS_OPTIONS = ("全部", "可匹配", "未匹配", "重复冲突", "同名冲突", "已导出", "已跳过", "复制失败")
HISTORY_TYPE_OPTIONS = ("全部", "整理", "筛选")
HISTORY_DATE_OPTIONS = ("全部", "最近7天", "最近30天")
FILTER_RULE_MODE_OPTIONS = ("不过滤", "等于任一", "包含任一", "不等于任一", "不包含任一")
FILTER_WORKFLOW_STEPS = (
    ("input", "输入", "选择 Excel、PDF 与输出位置"),
    ("rules", "规则", "确认工作表、列映射和筛选条件"),
    ("preview", "预览", "检查匹配、缺失与重复冲突"),
    ("execute", "执行", "安全复制文件并生成报告"),
    ("results", "结果", "查看状态、失败原因和导出位置"),
)

ORGANIZE_WORKFLOW_STEPS = (
    ("input", "目录", "选择发票目录和扫描范围"),
    ("preview", "扫描", "解析文件名并预览目标公司"),
    ("confirm", "确认", "勾选本次要移动的文件"),
    ("execute", "执行", "安全移动并记录恢复信息"),
    ("results", "结果", "查看成功、跳过、失败与撤销入口"),
)
UI_THEME_OPTIONS = ("day", "night")
UI_THEME_LABELS = {"day": "白天", "night": "黑夜"}
APP_TITLE = f"发票处理工具箱 {APP_VERSION}"
APP_ICON_RELATIVE_PATH = Path("assets") / "invoice-pdf-tool-icon.ico"

UI_THEME_PRESETS: Dict[str, Dict[str, Any]] = {
    "day": {
        "bootstrap_theme": "flatly",
        "root_bg": "#F3F7FB",
        "surface": "#FFFFFF",
        "surface_alt": "#F8FAFD",
        "surface_soft": "#EDF3F8",
        "surface_raised": "#FBFDFF",
        "surface_inset": "#EAF1F7",
        "title_bg": "#0F172A",
        "title_fg": "#F8FBFF",
        "title_muted": "#AFC0D4",
        "title_badge_bg": "#17304F",
        "title_badge_fg": "#E7F4FF",
        "hero_card_bg": "#13233B",
        "hero_card_border": "#24415F",
        "hero_accent": "#14B8A6",
        "hero_chip_bg": "#1B3351",
        "hero_chip_fg": "#E2F3FF",
        "text": "#102A43",
        "muted": "#64748B",
        "border": "#D7E0EA",
        "entry_bg": "#FFFFFF",
        "entry_fg": "#17324A",
        "button_bg": "#E8EEF5",
        "button_fg": "#17324A",
        "button_hover": "#D9E3EE",
        "button_disabled_fg": "#475569",
        "button_disabled_accent_fg": "#F8FAFC",
        "primary": "#0F7B78",
        "primary_hover": "#0B6462",
        "success": "#167653",
        "success_hover": "#125E43",
        "warning": "#8A5A12",
        "warning_hover": "#70460D",
        "danger": "#B83C4F",
        "danger_hover": "#9F2F42",
        "secondary": "#5E7388",
        "secondary_hover": "#4D6277",
        "log_bg": "#0F172A",
        "log_fg": "#E2E8F0",
        "log_drawer_bg": "#DCE6EF",
        "status_bg": "#E7EEF6",
        "status_fg": "#35516A",
        "tree_even": "#FAFCFE",
        "tree_odd": "#FFFFFF",
        "tree_selected": "#0F7B78",
        "tree_heading_bg": "#F1F5F9",
        "tree_heading_fg": "#16324A",
        "tab_active_bg": "#FFFFFF",
        "tab_active_fg": "#102A43",
        "tab_idle_bg": "#EAF1F7",
        "tab_idle_fg": "#64748B",
        "detail_bg": "#F7FAFD",
        "detail_fg": "#41566B",
        "status_success": "#166534",
        "status_missing": "#B91C1C",
        "status_skip": "#9A3412",
        "status_error": "#AD1457",
        "status_conflict": "#6B21A8",
        "status_preview": "#0C4A6E",
        "card_palette": [
            ("#DCFCE7", "#166534"),
            ("#DCF4F3", "#0F766E"),
            ("#FFF4DE", "#9A3412"),
            ("#FEE2E2", "#B91C1C"),
            ("#E0F2FE", "#0C4A6E"),
            ("#EEF2F7", "#334155"),
        ],
    },
    "night": {
        "bootstrap_theme": "darkly",
        "root_bg": "#07111D",
        "surface": "#0F1B2D",
        "surface_alt": "#132133",
        "surface_soft": "#18283D",
        "surface_raised": "#162538",
        "surface_inset": "#0C1828",
        "title_bg": "#060D18",
        "title_fg": "#F8FAFC",
        "title_muted": "#8DA3BC",
        "title_badge_bg": "#17304C",
        "title_badge_fg": "#DDEAFE",
        "hero_card_bg": "#0C1828",
        "hero_card_border": "#23415C",
        "hero_accent": "#22C7B8",
        "hero_chip_bg": "#17304C",
        "hero_chip_fg": "#D7F5F2",
        "text": "#E2E8F0",
        "muted": "#94A3B8",
        "border": "#24384D",
        "entry_bg": "#0B1626",
        "entry_fg": "#F8FAFC",
        "button_bg": "#1A2A3D",
        "button_fg": "#E2E8F0",
        "button_hover": "#24364A",
        "button_disabled_fg": "#CBD5E1",
        "button_disabled_accent_fg": "#F8FAFC",
        "primary": "#0F766E",
        "primary_hover": "#115E59",
        "success": "#167653",
        "success_hover": "#125E43",
        "warning": "#8A5A12",
        "warning_hover": "#70460D",
        "danger": "#B83C4F",
        "danger_hover": "#9F2F42",
        "secondary": "#56687E",
        "secondary_hover": "#64778E",
        "log_bg": "#020617",
        "log_fg": "#D8E1EC",
        "log_drawer_bg": "#142235",
        "status_bg": "#0E1A2B",
        "status_fg": "#B8C7D9",
        "tree_even": "#122033",
        "tree_odd": "#0F1B2D",
        "tree_selected": "#0F8F88",
        "tree_heading_bg": "#18283D",
        "tree_heading_fg": "#E2E8F0",
        "tab_active_bg": "#18283D",
        "tab_active_fg": "#F8FAFC",
        "tab_idle_bg": "#122033",
        "tab_idle_fg": "#94A3B8",
        "detail_bg": "#122033",
        "detail_fg": "#D3DEEA",
        "status_success": "#86EFAC",
        "status_missing": "#FCA5A5",
        "status_skip": "#FDBA74",
        "status_error": "#F9A8D4",
        "status_conflict": "#C4B5FD",
        "status_preview": "#93C5FD",
        "card_palette": [
            ("#163A2B", "#DDFBEA"),
            ("#133A56", "#D9F1FF"),
            ("#4A3012", "#FFF0D5"),
            ("#4A1F2A", "#FFE2E7"),
            ("#1C3445", "#D9F1FF"),
            ("#223042", "#E2E8F0"),
        ],
    },
}


# ==================== GUI 主应用 ====================

class InvoiceToolApp:
    """发票处理工具箱 v6.1.1"""

    def __init__(
        self,
        root: tk.Tk,
        *,
        config_file: Path = CONFIG_FILE,
        history_file: Path = HISTORY_FILE,
        active_task_file: Path = ACTIVE_TASK_FILE,
    ) -> None:
        self.root = root
        self._config_file = Path(config_file)
        self._history_file = Path(history_file)
        self._task_journal = TaskJournal(Path(active_task_file))
        self.root.title(APP_TITLE)
        self._apply_window_icon()
        self._apply_initial_window_geometry()
        self._default_widget_colors = self._capture_default_widget_colors()

        # 加载配置/历史
        loaded_config = load_json(
            self._config_file,
            {},
            expected_type=dict,
            quarantine_invalid=True,
        )
        self._config_write_blocked_reason = ""
        self._blocked_config_snapshot: Optional[Dict[str, Any]] = None
        if isinstance(loaded_config, dict):
            try:
                config_plan = normalize_config(loaded_config)
                self.config = config_plan.config
                for warning in config_plan.warnings:
                    logger.warning("配置兼容提示：%s", warning)
            except ConfigurationError as exc:
                self.config = dict(CONFIG_DEFAULTS)
                self._config_write_blocked_reason = str(exc)
                self._blocked_config_snapshot = dict(loaded_config)
                logger.error("配置版本不兼容，已使用只读默认配置：%s", exc)
        else:
            self.config = dict(CONFIG_DEFAULTS)
            logger.error("配置文件根节点必须是对象，已使用默认配置")
        loaded_history = load_json(
            self._history_file,
            [],
            expected_type=list,
            quarantine_invalid=True,
        )
        self.all_history: List[Dict] = (
            [record for record in loaded_history if isinstance(record, dict)]
            if isinstance(loaded_history, list)
            else []
        )
        if not isinstance(loaded_history, list):
            logger.error("历史记录根节点必须是数组，已忽略无效内容")
        self._recovered_task_count = self._recover_interrupted_task()
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
        self._pause_flag = threading.Event()
        self._start_time: Optional[float] = None
        self._worker_thread: Optional[threading.Thread] = None
        self._active_task_id: Optional[str] = None
        self._active_task_kind = ""
        self._readonly_task_name = ""
        self._readonly_task_sequence = 0
        self._readonly_task_controls: List[Tuple[tk.Widget, str, str]] = []
        self._readonly_task_cancel_button: Optional[tk.Button] = None
        self._readonly_task_progress_bar: Optional[ttk.Progressbar] = None
        self._readonly_task_progress_mode = "determinate"
        self._closing_requested = False
        self._close_finalized = False
        self._ui_events: queue.Queue[Callable[[], None]] = queue.Queue()
        self._ui_event_pump_id: Optional[str] = None

        # ─── 整理变量 ───
        self.organize_folder_path = tk.StringVar()
        self.file_check_vars: Dict[str, tk.BooleanVar] = {}
        self.preview_data: Dict[str, Dict] = {}
        self.current_session_history: List[Dict] = []
        self.organize_failed_files: List[str] = []
        self.organize_failure_folder = ""
        self._pending_organize_rerun_files: set[str] = set()
        self.organize_recursive = tk.BooleanVar(value=False)
        self.organize_workflow_stage = tk.StringVar(value="input")
        self.organize_workflow_status_text = tk.StringVar(value="先选择发票目录，再扫描并确认本次要整理的文件。")
        self.organize_result_title = tk.StringVar(value="等待扫描")
        self.organize_result_detail = tk.StringVar(value="扫描后会在这里汇总可处理、已跳过和无效文件。")

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
            self.excel_sheet_name.get(),
        )
        self.pdf_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.manual_output_folder = tk.StringVar(value=self.config.get("output_folder", ""))
        self.auto_output_by_sheet = tk.BooleanVar(value=self.config.get("auto_output_by_sheet", True))
        self.filter_recursive = tk.BooleanVar(value=False)
        self.filter_result_status = tk.StringVar(value="全部")
        self.filter_result_keyword = tk.StringVar()
        self.filter_result_rows: List[FilterResultRow] = []
        self._last_filter_preview_signature: Optional[Tuple[str, ...]] = None
        self._last_filter_preview_result: Optional[FilterPreviewResult] = None
        self.filter_result_sort_key = "invoice"
        self.filter_result_sort_desc = False
        self.filter_result_selection: Dict[str, FilterResultRow] = {}
        self.filter_missing_invoices: List[str] = []
        self.filter_summary_title = tk.StringVar(value="等待预览或筛选")
        self.filter_summary_subtitle = tk.StringVar(value="先选择 Excel、PDF 和导出目录，然后执行预览或筛选。")
        self.filter_detail_var = tk.StringVar(value="提示：结果将显示在下方表格中，可按状态过滤或搜索发票号。")
        self.filter_workflow_stage = tk.StringVar(value="input")
        self.filter_workflow_status_text = tk.StringVar(value="先补齐三个输入位置，再确认工作表和筛选规则。")
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
        self.workbook_analysis_compact_var = tk.StringVar(value="列映射、行筛选和样本预览默认收起，需要调整时展开。")
        self.workbook_sheet_overview_var = tk.StringVar(value="先选择 Excel 文件，再从左侧查看每个 sheet 的识别结果。")
        self.workbook_sheet_sample_var = tk.StringVar(value="样本预览会显示当前工作表前几行数据，便于确认列是否正确。")
        self.workbook_analysis_expanded = tk.BooleanVar(value=False)
        self.workbook_analysis_result: Optional[WorkbookAnalysisResult] = None
        self.workbook_profiles: Dict[str, WorkbookSheetProfile] = {}
        self.workbook_tree_selection: Dict[str, str] = {}
        self.history_type_filter = tk.StringVar(value="全部")
        self.history_date_filter = tk.StringVar(value="全部")
        self.history_keyword = tk.StringVar()
        self.filtered_history_indices: List[int] = []
        self.history_summary_var = tk.StringVar(value="显示 0 / 0 条历史记录")
        self.history_detail_title = tk.StringVar(value="未选择任务")
        self.history_detail_meta = tk.StringVar(value="从左侧选择一条记录查看处理摘要。")
        self.history_detail_folder = tk.StringVar(value="")
        self.history_detail_safety = tk.StringVar(value="尚未选择可评估的历史记录")
        self.history_action_status_text = tk.StringVar(value="选择任务后可查看详情、打开目录或执行安全回滚。")
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
        self.settings_status_text = tk.StringVar(value="配置会自动保存；导入或恢复默认前会先创建本地备份。")

        # 日志抽屉状态
        self._log_visible = tk.BooleanVar(value=False)
        self._gui_log_handler: Optional[TkTextHandler] = None
        self._recent_error_handler: Optional[RecentErrorHandler] = None

        # 构建界面
        self._build_ui()
        self._setup_drag_and_drop()
        self._restore_paths()
        if self._recovered_task_count:
            self.status_var.set(
                f"⚠ 已从上次异常中恢复 {self._recovered_task_count} 个文件操作；请到任务历史检查并按需安全回滚。"
            )
            self.history_action_status_text.set(
                f"发现异常中断恢复记录，共 {self._recovered_task_count} 个项目；已保留校验信息。"
            )
        elif CONFIG_DIR_FALLBACK_REASON:
            self.status_var.set("⚠ 首选配置目录不可用，当前使用临时回退目录；请导出诊断包检查环境权限。")
        if PANDAS_SUPPORT:
            self._refresh_excel_sheets(silent=True)

        # 关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self._ui_event_pump_id = self.root.after(25, self._drain_ui_events)
        logger.info("应用启动 %s", APP_VERSION)
        logger.info(
            "运行能力：pandas=%s | openpyxl=%s | xlrd=%s | ttkbootstrap=%s | tkinterdnd2=%s",
            PANDAS_SUPPORT,
            OPENPYXL_SUPPORT,
            XLRD_SUPPORT,
            MODERN_UI,
            DND_SUPPORT,
        )

    # ==================== JSON / 配置 ====================

    @staticmethod
    def _load_json(path: Path, default: Any) -> Any:
        return load_json(path, default)

    @staticmethod
    def _save_json(path: Path, data: Any) -> bool:
        return save_json(path, data)

    def _recover_interrupted_task(self) -> int:
        journal = self._task_journal.load()
        if not journal:
            return 0

        task_id = str(journal.get("task_id", ""))
        if task_id and any(record.get("task_id") == task_id for record in self.all_history):
            if self._save_history():
                self._task_journal.clear(task_id)
            return 0

        moves = [item for item in journal.get("moves", []) if isinstance(item, dict)]
        report_files = [str(item) for item in journal.get("report_files", []) if str(item)]
        report_entries = [item for item in journal.get("report_entries", []) if isinstance(item, dict)]
        if not moves and not report_files:
            self._task_journal.clear(task_id or None)
            return 0

        operation_type = str(journal.get("type", "整理"))
        if operation_type not in {"整理", "筛选"}:
            operation_type = "整理"
        record: Dict[str, Any] = {
            "time": str(journal.get("updated_at") or journal.get("started_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            "folder": str(journal.get("folder", "")),
            "count": len(moves),
            "type": operation_type,
            "moves": moves,
            "task_id": task_id,
            "recovered": True,
        }
        if report_files:
            record["report_files"] = report_files
        if report_entries:
            record["report_entries"] = report_entries
        metadata = journal.get("metadata", {})
        if isinstance(metadata, dict) and isinstance(metadata.get("rerun"), dict):
            record["rerun"] = dict(metadata["rerun"])
        self.all_history.insert(0, record)
        self.all_history = self.all_history[:100]
        if self._save_history():
            self._task_journal.clear(task_id or None)
        logger.warning("检测到上次未完成任务，已恢复 %s 个文件操作到历史记录", len(moves))
        return len(moves) + len(report_files)

    def _recover_failed_task_into_history(self) -> int:
        recovered = self._recover_interrupted_task()
        if recovered and hasattr(self, "history_tree"):
            self._refresh_history_tree()
        return recovered

    @staticmethod
    def _resource_path(relative_path: Path) -> Path:
        base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
        return base_path / relative_path

    def _apply_window_icon(self) -> None:
        icon_path = self._resource_path(APP_ICON_RELATIVE_PATH)
        if not icon_path.exists():
            return
        try:
            self.root.iconbitmap(str(icon_path))
        except tk.TclError:
            pass

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
        for widget in probes.values():
            snapshot: Dict[str, str] = {}
            for option in ("bg", "fg", "activebackground", "activeforeground", "insertbackground"):
                try:
                    snapshot[option] = str(widget.cget(option))
                except tk.TclError:
                    continue
            # Tk reports ``tk.LabelFrame`` as ``Labelframe`` on Windows.
            # Store the runtime class name so theme traversal can recognise the
            # real system-default background instead of leaving a light panel
            # behind when the user switches to night mode.
            defaults[widget.winfo_class()] = snapshot
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
        if self._task_is_running():
            self.ui_theme_label.set(self._theme_label(self.ui_theme.get()))
            self.status_var.set("⚠ 任务运行期间暂不重建界面，请在任务结束后切换主题。")
            messagebox.showwarning("任务进行中", "请在当前任务结束或取消后再切换主题。")
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
        filter_workflow_stage = "input"
        filter_workflow_status = ""
        organize_workflow_stage = "input"
        organize_workflow_status = ""
        organize_result_title = "等待扫描"
        organize_result_detail = ""
        selected_history_index: Optional[int] = None
        if hasattr(self, "notebook"):
            try:
                selected_tab = self.notebook.index(self.notebook.select())
            except Exception:
                selected_tab = 0
        if hasattr(self, "filter_workflow_stage"):
            filter_workflow_stage = self.filter_workflow_stage.get()
        if hasattr(self, "filter_workflow_status_text"):
            filter_workflow_status = self.filter_workflow_status_text.get()
        if hasattr(self, "organize_workflow_stage"):
            organize_workflow_stage = self.organize_workflow_stage.get()
        if hasattr(self, "organize_workflow_status_text"):
            organize_workflow_status = self.organize_workflow_status_text.get()
        if hasattr(self, "organize_result_title"):
            organize_result_title = self.organize_result_title.get()
        if hasattr(self, "organize_result_detail"):
            organize_result_detail = self.organize_result_detail.get()
        if hasattr(self, "history_tree"):
            try:
                selected_history_index = self._get_selected_history_index()
            except (AttributeError, IndexError, tk.TclError):
                selected_history_index = None

        if self._gui_log_handler is not None:
            logger.removeHandler(self._gui_log_handler)
            self._gui_log_handler.close()
            self._gui_log_handler = None
        if self._recent_error_handler is not None:
            logger.removeHandler(self._recent_error_handler)
            self._recent_error_handler.close()
            self._recent_error_handler = None

        for child in self.root.winfo_children():
            child.destroy()

        self._build_ui()
        self._setup_drag_and_drop()
        if PANDAS_SUPPORT:
            self._refresh_excel_sheets(silent=True)
        self._set_filter_workflow_stage(filter_workflow_stage, filter_workflow_status)
        self._render_organize_preview()
        self._set_organize_workflow_stage(organize_workflow_stage, organize_workflow_status)
        self._update_organize_result(organize_result_title, organize_result_detail)
        self._refresh_filter_result_tree()
        self._refresh_history_tree(preferred_index=selected_history_index)
        self._refresh_recent_error_list()
        try:
            self.notebook.select(selected_tab)
        except Exception:
            pass

    @staticmethod
    def _mix_colors(color1: str, color2: str, weight: float = 0.2) -> str:
        """Mix color2 into color1. weight is the proportion of color2 (0.0 to 1.0)"""
        try:
            c1 = color1.lstrip('#')
            c2 = color2.lstrip('#')
            r1, g1, b1 = int(c1[0:2], 16), int(c1[2:4], 16), int(c1[4:6], 16)
            r2, g2, b2 = int(c2[0:2], 16), int(c2[2:4], 16), int(c2[4:6], 16)
            
            r = int(r1 * (1 - weight) + r2 * weight)
            g = int(g1 * (1 - weight) + g2 * weight)
            b = int(b1 * (1 - weight) + b2 * weight)
            
            return f"#{r:02X}{g:02X}{b:02X}"
        except Exception:
            return color1

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
            font=("微软雅黑", 10, "bold"),
            padding=[16, 6],
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
            rowheight=28,
            font=("微软雅黑", 9),
            background=palette["tree_odd"],
            fieldbackground=palette["tree_odd"],
            foreground=palette["text"],
            bordercolor=palette["border"],
        )
        style.configure(
            "Treeview.Heading",
            font=("微软雅黑", 10, "bold"),
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
                    widget.configure(
                        fg=palette["text"],
                        relief="flat",
                        bd=0,
                        highlightthickness=1,
                        highlightbackground=palette["border"],
                        highlightcolor=palette["border"]
                    )
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
                bd=0,
                highlightthickness=1,
            )
        elif cls == "Listbox":
            widget.configure(
                bg=palette["entry_bg"],
                fg=palette["entry_fg"],
                selectbackground=palette["tree_selected"],
                selectforeground="#FFFFFF",
                highlightbackground=palette["border"],
                highlightcolor=palette["primary"],
                relief="flat",
                bd=0,
                highlightthickness=1,
            )
        elif cls == "Button" and self._should_apply_default_bg(widget):
            widget.configure(
                bg=palette["button_bg"],
                fg=palette["button_fg"],
                activebackground=palette["button_hover"],
                activeforeground=palette["button_fg"],
                disabledforeground=palette["button_disabled_fg"],
                relief="flat",
                bd=0,
                highlightthickness=0,
            )
            self._bind_hover(widget, palette["button_bg"], palette["button_hover"])
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
        resolved_sheet = self.excel_sheet_name.get() if sheet_name is None else str(sheet_name)
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
            sheet_name = self.excel_sheet_name.get()
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

    def _get_safe_company_name_index(self) -> int:
        try:
            val = self.company_name_index.get()
            if val < 0:
                raise ValueError("索引不能为负数")
            return val
        except (tk.TclError, ValueError):
            return self.config.get("company_name_index", 2)

    def _get_safe_invoice_number_index(self) -> int:
        try:
            val = self.invoice_number_index.get()
            if val < 0:
                raise ValueError("索引不能为负数")
            return val
        except (tk.TclError, ValueError):
            return self.config.get("invoice_number_index", 1)

    def _collect_runtime_config(self) -> Dict[str, Any]:
        config = dict(self.config)
        config["config_schema_version"] = CONFIG_SCHEMA_VERSION
        config["ui_theme"] = self.ui_theme.get().strip()
        config["rule_preset_id"] = self.rule_preset_id.get().strip()
        config["company_name_index"] = self._get_safe_company_name_index()
        config["invoice_number_index"] = self._get_safe_invoice_number_index()
        config["excel_sheet_name"] = self.excel_sheet_name.get()
        config["selected_invoice_column_name"] = self.selected_invoice_column_name.get().strip()
        config["selected_company_column_name"] = self.selected_company_column_name.get().strip()
        config["row_filter_column_name"] = self.row_filter_column_name.get().strip()
        config["row_filter_mode"] = self.row_filter_mode.get().strip()
        config["row_filter_values"] = self.row_filter_values.get().strip()
        config["company_exclude_keywords"] = self.company_exclude_keywords.get().strip()
        config["organize_folder"] = self.organize_folder_path.get().strip()
        config["excel_path"] = self.excel_path.get().strip()
        config["pdf_folder"] = self.pdf_folder.get().strip()
        config["auto_output_by_sheet"] = bool(self.auto_output_by_sheet.get())
        config["output_folder"] = self.manual_output_folder.get().strip()
        config["invoice_column_aliases"] = self.invoice_column_aliases.get().strip()
        config["company_column_aliases"] = self.company_column_aliases.get().strip()
        return config

    def _save_config(self) -> bool:
        if self._config_write_blocked_reason:
            logger.error("配置写入已阻止：%s", self._config_write_blocked_reason)
            return False
        self.config = self._collect_runtime_config()
        return self._save_json(self._config_file, self.config)

    def _save_history(self) -> bool:
        return self._save_json(self._history_file, self.all_history)

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
        if self._close_finalized:
            return
        with self._lock:
            running = self.is_running
        if running:
            if self._closing_requested:
                return
            if not messagebox.askyesno(
                "任务仍在运行",
                "当前任务仍在处理文件。\n\n是否请求取消，并在当前文件安全处理完成后自动关闭？",
            ):
                return
            self._closing_requested = True
            self._cancel_flag.set()
            pause_flag = getattr(self, "_pause_flag", None)
            if pause_flag is not None:
                pause_flag.clear()
            self.status_var.set("⏳ 正在安全停止任务，完成后将自动关闭...")
            logger.warning("用户请求关闭，正在等待文件任务安全停止")
            self.root.after(100, self._wait_for_task_before_close)
            return
        self._finalize_close()

    def _wait_for_task_before_close(self) -> None:
        with self._lock:
            running = self.is_running
        if running:
            self.root.after(100, self._wait_for_task_before_close)
            return
        self._finalize_close()

    def _finalize_close(self) -> None:
        if self._close_finalized:
            return
        self._close_finalized = True
        if self._ui_event_pump_id is not None:
            try:
                self.root.after_cancel(self._ui_event_pump_id)
            except tk.TclError:
                pass
            self._ui_event_pump_id = None
        self._save_config()
        if self._gui_log_handler is not None:
            logger.removeHandler(self._gui_log_handler)
            self._gui_log_handler.close()
            self._gui_log_handler = None
        if self._recent_error_handler is not None:
            logger.removeHandler(self._recent_error_handler)
            self._recent_error_handler.close()
            self._recent_error_handler = None
        logger.info("应用关闭")
        self.root.destroy()

    def _auto_clean_old_history(self) -> None:
        cutoff = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
        before = len(self.all_history)
        self.all_history = [
            record
            for record in self.all_history
            if isinstance(record.get("time"), str) and record["time"] >= cutoff
        ][:100]
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
        disabled_fg = (
            self.palette["button_disabled_fg"]
            if role == "neutral"
            else self.palette["button_disabled_accent_fg"]
        )
        button.config(
            bg=normal_bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            disabledforeground=disabled_fg,
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

        # 标题栏底部 1px 主色渐变条（微光分割线）
        title_sep = tk.Frame(self.root, height=1, bg=palette["primary"])
        title_sep.pack(fill="x")

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
        theme_icon = "☀️" if self.ui_theme.get() == "day" else "🌙"
        self.theme_badge = tk.Label(
            title_actions,
            text=f"{theme_icon} {self._theme_label()}模式",
            font=("微软雅黑", 8, "bold"),
            bg=palette["title_badge_bg"],
            fg=palette["title_badge_fg"],
            padx=12,
            pady=3,
            highlightthickness=1,
            highlightbackground=self._mix_colors(palette["title_badge_bg"], palette["title_badge_fg"], 0.2),
            highlightcolor=self._mix_colors(palette["title_badge_bg"], palette["title_badge_fg"], 0.2),
        )
        self.theme_badge.pack(side="left", padx=(0, 8))
        self.theme_toggle_btn = tk.Button(
            title_actions,
            text="🌙 切换到黑夜" if self.ui_theme.get() == "day" else "☀️ 切换到白天",
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
            toggle_bar, text="▲ 运行日志（点击展开）" if not self._log_visible.get() else "▼ 运行日志（点击收起）",
            font=("微软雅黑", 9, "bold"), bg=palette["log_drawer_bg"], fg=palette["text"],
            padx=10, pady=4, cursor="hand2",
        )
        self._log_toggle_label.pack(side="left")

        # 工具按钮（始终可见）
        self.log_clear_btn = tk.Button(
            toggle_bar, text="清空", font=("微软雅黑", 8),
            command=self._clear_log, padx=6, pady=0, bd=0, bg=palette["log_drawer_bg"], fg=palette["text"],
        )
        self.log_clear_btn.pack(side="right", padx=(0, 6))

        self.log_export_btn = tk.Button(
            toggle_bar, text="导出", font=("微软雅黑", 8),
            command=self._export_log, padx=6, pady=0, bd=0, bg=palette["log_drawer_bg"], fg=palette["text"],
        )
        self.log_export_btn.pack(side="right")

        self.log_copy_btn = tk.Button(
            toggle_bar, text="复制", font=("微软雅黑", 8),
            command=self._copy_log, padx=6, pady=0, bd=0, bg=palette["log_drawer_bg"], fg=palette["text"],
        )
        self.log_copy_btn.pack(side="right")

        # Bind hover effects on toggle_bar and its child elements
        normal_bg = palette["log_drawer_bg"]
        hover_bg = self._mix_colors(normal_bg, palette["text"], 0.08)
        btn_hover_bg = self._mix_colors(normal_bg, palette["text"], 0.18)

        def on_enter(e):
            toggle_bar.config(bg=hover_bg)
            self._log_toggle_label.config(bg=hover_bg)
            self.log_clear_btn.config(bg=hover_bg)
            self.log_export_btn.config(bg=hover_bg)
            self.log_copy_btn.config(bg=hover_bg)
        def on_leave(e):
            toggle_bar.config(bg=normal_bg)
            self._log_toggle_label.config(bg=normal_bg)
            self.log_clear_btn.config(bg=normal_bg)
            self.log_export_btn.config(bg=normal_bg)
            self.log_copy_btn.config(bg=normal_bg)

        toggle_bar.bind("<Enter>", on_enter)
        toggle_bar.bind("<Leave>", on_leave)
        self._log_toggle_label.bind("<Enter>", on_enter)
        self._log_toggle_label.bind("<Leave>", on_leave)
        self._bind_hover(self.log_clear_btn, hover_bg, btn_hover_bg)
        self._bind_hover(self.log_export_btn, hover_bg, btn_hover_bg)
        self._bind_hover(self.log_copy_btn, hover_bg, btn_hover_bg)

        toggle_bar.bind("<Button-1>", lambda e: self._toggle_log_drawer())
        self._log_toggle_label.bind("<Button-1>", lambda e: self._toggle_log_drawer())

        # 日志内容区（初始隐藏）
        self._log_content = tk.Frame(self._drawer_frame, bg=palette["root_bg"])

        log_scroll = tk.Scrollbar(self._log_content)
        log_scroll.pack(side="right", fill="y")

        self.log_text = tk.Text(
            self._log_content, font=("Consolas", 9), wrap="word",
            yscrollcommand=log_scroll.set, bg=palette["log_bg"], fg=palette["log_fg"], height=6,
            highlightthickness=1, highlightbackground=palette["border"], highlightcolor=palette["primary"],
            relief="flat", bd=0
        )
        self.log_text.pack(fill="both", expand=True)
        log_scroll.config(command=self.log_text.yview)

        # Make console tag colors theme-aware
        is_night = self.ui_theme.get() == "night"
        self.log_text.tag_config("success", foreground="#34D399" if is_night else "#166534")
        self.log_text.tag_config("error", foreground="#FB7185" if is_night else "#B91C1C")
        self.log_text.tag_config("warning", foreground="#FBBF24" if is_night else "#B45309")
        self.log_text.tag_config("info", foreground="#7DD3FC" if is_night else "#1E40AF")
        self.log_text.tag_config("header", foreground="#C4B5FD" if is_night else "#6D28D9")

        # 注册自定义 GUI handler
        self._gui_log_handler = TkTextHandler(self.log_text, self._post_ui)
        self._gui_log_handler.setFormatter(
            RedactingFormatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
        )
        logger.addHandler(self._gui_log_handler)

        self._recent_error_handler = RecentErrorHandler(self._append_recent_error, self._post_ui)
        self._recent_error_handler.setFormatter(
            RedactingFormatter("[%(asctime)s] %(levelname)s %(message)s", datefmt="%H:%M:%S")
        )
        logger.addHandler(self._recent_error_handler)

    def _toggle_log_drawer(self) -> None:
        if self._log_visible.get():
            self._log_content.pack_forget()
            self._log_toggle_label.config(text="▲ 运行日志（点击展开）")
            self._log_visible.set(False)
        else:
            self._log_content.pack(fill="both", expand=False)
            self._log_toggle_label.config(text="▼ 运行日志（点击收起）")
            self._log_visible.set(True)

    def _clear_log(self) -> None:
        self.log_text.delete(1.0, tk.END)

    def _copy_log(self) -> None:
        content = self.log_text.get(1.0, tk.END).strip()
        if not content:
            self.status_var.set("运行日志为空，无内容可复制。")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.status_var.set("✅ 已复制脱敏后的运行日志")

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
        config_directory = self._config_file.parent
        config_directory.mkdir(parents=True, exist_ok=True)
        self._open_path_in_shell(config_directory)

    # ─────────────── Tab1: 发票整理 ───────────────

    def _build_organize_workflow_rail(self, parent: tk.Widget) -> None:
        rail = tk.Frame(
            parent,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=8,
            pady=8,
        )
        rail.pack(fill="x", pady=(0, 8))
        self.organize_workflow_cards: Dict[str, tk.Button] = {}
        for index, (stage_key, title, _description) in enumerate(ORGANIZE_WORKFLOW_STEPS, 1):
            button = tk.Button(
                rail,
                text=f"{index}  {title}",
                font=("微软雅黑", 9, "bold"),
                relief="flat",
                bd=0,
                padx=8,
                pady=7,
                takefocus=True,
                cursor="hand2",
                command=lambda key=stage_key: self._focus_organize_workflow_step(key),
            )
            button.pack(side="left", fill="x", expand=True, padx=2)
            self.organize_workflow_cards[stage_key] = button
        self._set_organize_workflow_stage(self.organize_workflow_stage.get())

    def _build_organize_section_heading(self, parent: tk.Widget, stage_key: str) -> tk.Frame:
        step_index = next(
            index
            for index, (key, _title, _description) in enumerate(ORGANIZE_WORKFLOW_STEPS, 1)
            if key == stage_key
        )
        _key, title, description = ORGANIZE_WORKFLOW_STEPS[step_index - 1]
        heading = tk.Frame(parent, bg=self.palette["root_bg"])
        heading.pack(fill="x", pady=(5, 3))
        tk.Label(
            heading,
            text=f"{step_index:02d}",
            font=("Segoe UI", 8, "bold"),
            bg=self.palette["primary"],
            fg="#FFFFFF",
            padx=8,
            pady=4,
        ).pack(side="left", padx=(0, 8))
        text_column = tk.Frame(heading, bg=self.palette["root_bg"])
        text_column.pack(side="left", fill="x", expand=True)
        tk.Label(
            text_column,
            text=title,
            font=("微软雅黑", 10, "bold"),
            bg=self.palette["root_bg"],
            fg=self.palette["text"],
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            text_column,
            text=description,
            font=("微软雅黑", 9),
            bg=self.palette["root_bg"],
            fg=self.palette["muted"],
            anchor="w",
        ).pack(fill="x")
        self.organize_workflow_sections[stage_key] = heading
        return heading

    def _set_organize_workflow_stage(self, stage_key: str, status_text: Optional[str] = None) -> None:
        stage_keys = [key for key, _title, _description in ORGANIZE_WORKFLOW_STEPS]
        if stage_key not in stage_keys:
            stage_key = "input"
        stage_var = getattr(self, "organize_workflow_stage", None)
        if stage_var is not None:
            stage_var.set(stage_key)
        status_var = getattr(self, "organize_workflow_status_text", None)
        if status_text is not None and status_var is not None:
            status_var.set(status_text)
        cards = getattr(self, "organize_workflow_cards", {})
        active_index = stage_keys.index(stage_key)
        for index, (key, title, _description) in enumerate(ORGANIZE_WORKFLOW_STEPS):
            button = cards.get(key)
            if button is None:
                continue
            if index < active_index:
                bg = self.palette["surface_soft"]
                fg = self.palette["status_success"]
                text = f"✓  {title}"
            elif index == active_index:
                bg = self.palette["primary"]
                fg = "#FFFFFF"
                text = f"{index + 1}  {title}"
            else:
                bg = self.palette["surface_raised"]
                fg = self.palette["muted"]
                text = f"{index + 1}  {title}"
            button.configure(
                text=text,
                bg=bg,
                fg=fg,
                activebackground=(self.palette["primary_hover"] if index == active_index else self.palette["surface_soft"]),
                activeforeground="#FFFFFF" if index == active_index else self.palette["text"],
                highlightthickness=1,
                highlightbackground=self.palette["border"],
                highlightcolor=self.palette["primary"],
            )

    def _focus_organize_workflow_step(self, stage_key: str) -> str:
        focus_targets = {
            "input": getattr(self, "organize_folder_entry", None),
            "preview": getattr(self, "org_scan_btn", None),
            "confirm": getattr(self, "file_tree", None),
            "execute": getattr(self, "start_btn", None),
            "results": getattr(self, "undo_btn", None),
        }
        target = focus_targets.get(stage_key)
        if target is not None:
            try:
                target.focus_set()
            except tk.TclError:
                pass
        return "break"

    def _update_organize_result(self, title: str, detail: str) -> None:
        title_var = getattr(self, "organize_result_title", None)
        detail_var = getattr(self, "organize_result_detail", None)
        if title_var is not None:
            title_var.set(title)
        if detail_var is not None:
            detail_var.set(detail)

    def _build_organize_tab(self) -> None:
        workflow_shell = tk.Frame(self.organize_frame, bg=self.palette["root_bg"])
        workflow_shell.pack(fill="both", expand=True)
        self.organize_action_bar = tk.Frame(
            workflow_shell,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        self.organize_action_bar.pack(side="bottom", fill="x", pady=(6, 0))
        panel = tk.Frame(workflow_shell, bg=self.palette["root_bg"])
        panel.pack(side="top", fill="both", expand=True)
        self.organize_workflow_sections: Dict[str, tk.Frame] = {}

        self._build_organize_workflow_rail(panel)
        self._build_organize_section_heading(panel, "input")

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

        self.org_browse_btn = tk.Button(row, text="浏览", font=("微软雅黑", 10), command=self._browse_organize_folder, padx=15)
        self.org_browse_btn.pack(side="right", padx=(10, 0))
        self._style_action_button(self.org_browse_btn, "neutral")

        self.org_scan_btn = tk.Button(row, text="🔍 扫描", font=("微软雅黑", 10), command=self._scan_files, padx=15)
        self.org_scan_btn.pack(side="right", padx=(5, 0))
        self._style_action_button(self.org_scan_btn, "secondary")

        opt = tk.Frame(folder_lf)
        opt.pack(fill="x", pady=(8, 0))
        tk.Checkbutton(opt, text="包含子文件夹", font=("微软雅黑", 9), variable=self.organize_recursive).pack(side="left")
        idx = self.company_name_index.get()
        self.organize_hint = tk.Label(opt, text=f"  💡 公司名在第{idx+1}段（可在设置中修改）", font=("微软雅黑", 9), fg=self.palette["muted"])
        self.organize_hint.pack(side="left", padx=(10, 0))

        # 文件列表
        self._build_organize_section_heading(panel, "preview")
        list_lf = tk.LabelFrame(panel, text=" 📋 文件预览（勾选要处理的文件）", font=("微软雅黑", 10, "bold"), padx=10, pady=8)
        list_lf.pack(fill="both", expand=True, pady=(10, 0))

        sel_bar = tk.Frame(list_lf)
        sel_bar.pack(fill="x", pady=(0, 5))
        self.org_sel_all_btn = tk.Button(sel_bar, text="✅ 全选", font=("微软雅黑", 9), command=self._select_all, padx=8)
        self.org_sel_all_btn.pack(side="left", padx=(0, 4))
        self._style_action_button(self.org_sel_all_btn, "neutral")

        self.org_desel_all_btn = tk.Button(sel_bar, text="⬜ 取消全选", font=("微软雅黑", 9), command=self._deselect_all, padx=8)
        self.org_desel_all_btn.pack(side="left")
        self._style_action_button(self.org_desel_all_btn, "neutral")

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
        self.file_tree.tag_configure("failure", foreground=self.palette["status_error"])
        self.file_tree.tag_configure("failure_even", foreground=self.palette["status_error"], background=self.palette["tree_even"])
        self.file_tree.tag_configure("success", foreground=self.palette["status_success"])
        self.file_tree.tag_configure("success_even", foreground=self.palette["status_success"], background=self.palette["tree_even"])

        self._build_organize_section_heading(panel, "results")
        result_card = tk.Frame(
            panel,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=12,
            pady=7,
        )
        result_card.pack(fill="x")
        tk.Label(
            result_card,
            textvariable=self.organize_result_title,
            font=("微软雅黑", 10, "bold"),
            bg=self.palette["surface"],
            fg=self.palette["text"],
            anchor="w",
        ).pack(side="left")
        tk.Label(
            result_card,
            textvariable=self.organize_result_detail,
            font=("微软雅黑", 8),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=(12, 0))

        # 按钮区
        btn_bar = self.organize_action_bar
        tk.Label(
            btn_bar,
            textvariable=self.organize_workflow_status_text,
            font=("微软雅黑", 8),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
            wraplength=250,
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))
        controls = tk.Frame(btn_bar, bg=self.palette["surface"])
        controls.pack(side="right")

        self.start_btn = tk.Button(
            controls, text="🚀 执行整理", font=("微软雅黑", 12, "bold"),
            padx=22, pady=7, cursor="hand2", command=self._execute_organize,
        )
        self.start_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.start_btn, "success")

        self.retry_org_btn = tk.Button(
            controls,
            text="重试失败",
            font=("微软雅黑", 10),
            padx=11,
            pady=5,
            cursor="hand2",
            command=self._retry_failed_organize,
            state="disabled",
        )
        self.retry_org_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.retry_org_btn, "secondary")

        self.pause_org_btn = tk.Button(
            controls,
            text="⏸ 暂停",
            font=("微软雅黑", 10),
            padx=10,
            pady=5,
            cursor="hand2",
            command=self._toggle_pause_task,
            state="disabled",
        )
        self.pause_org_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.pause_org_btn, "secondary")

        self.undo_btn = tk.Button(
            controls, text="↩ 撤销上次", font=("微软雅黑", 10),
            padx=12, pady=5, cursor="hand2",
            command=self._undo_last_move, state="disabled",
        )
        self.undo_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.undo_btn, "warning")

        self.undo_all_btn = tk.Button(
            controls, text="↩ 撤销全部", font=("微软雅黑", 10),
            padx=12, pady=5, cursor="hand2",
            command=self._undo_all_moves, state="disabled",
        )
        self.undo_all_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.undo_all_btn, "danger")

        self.cancel_org_btn = tk.Button(
            controls, text="⏹ 取消", font=("微软雅黑", 10),
            padx=12, pady=5, cursor="hand2",
            command=self._cancel_task, state="disabled",
        )
        self.cancel_org_btn.pack(side="left")
        self._style_action_button(self.cancel_org_btn, "secondary")

        self.organize_progress = ttk.Progressbar(controls, mode="determinate", length=150)
        self.organize_progress.pack(side="right")

    # ─────────────── Tab2: 发票筛选 ───────────────

    def _build_filter_workflow_rail(self, parent: tk.Widget) -> None:
        rail = tk.Frame(
            parent,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=8,
            pady=8,
        )
        rail.pack(fill="x", pady=(0, 8))
        self.filter_workflow_cards: Dict[str, tk.Button] = {}
        for index, (stage_key, title, _description) in enumerate(FILTER_WORKFLOW_STEPS, 1):
            button = tk.Button(
                rail,
                text=f"{index}  {title}",
                font=("微软雅黑", 9, "bold"),
                relief="flat",
                bd=0,
                padx=8,
                pady=7,
                takefocus=True,
                cursor="hand2",
                command=lambda key=stage_key: self._scroll_filter_workflow_to(key),
            )
            button.pack(side="left", fill="x", expand=True, padx=2)
            self.filter_workflow_cards[stage_key] = button
        self._set_filter_workflow_stage(self.filter_workflow_stage.get())

    def _build_filter_section_heading(self, parent: tk.Widget, stage_key: str) -> tk.Frame:
        step_index = next(
            index
            for index, (key, _title, _description) in enumerate(FILTER_WORKFLOW_STEPS, 1)
            if key == stage_key
        )
        _key, title, description = FILTER_WORKFLOW_STEPS[step_index - 1]
        heading = tk.Frame(parent, bg=self.palette["root_bg"])
        heading.pack(fill="x", pady=(8, 3))
        tk.Label(
            heading,
            text=f"{step_index:02d}",
            font=("Segoe UI", 8, "bold"),
            bg=self.palette["primary"],
            fg="#FFFFFF",
            padx=8,
            pady=4,
        ).pack(side="left", padx=(0, 8))
        text_column = tk.Frame(heading, bg=self.palette["root_bg"])
        text_column.pack(side="left", fill="x", expand=True)
        tk.Label(
            text_column,
            text=title,
            font=("微软雅黑", 10, "bold"),
            bg=self.palette["root_bg"],
            fg=self.palette["text"],
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            text_column,
            text=description,
            font=("微软雅黑", 8),
            bg=self.palette["root_bg"],
            fg=self.palette["muted"],
            anchor="w",
        ).pack(fill="x")
        self.filter_workflow_sections[stage_key] = heading
        return heading

    def _set_filter_workflow_stage(self, stage_key: str, status_text: Optional[str] = None) -> None:
        stage_keys = [key for key, _title, _description in FILTER_WORKFLOW_STEPS]
        if stage_key not in stage_keys:
            stage_key = "input"
        stage_var = getattr(self, "filter_workflow_stage", None)
        if stage_var is not None:
            stage_var.set(stage_key)
        status_var = getattr(self, "filter_workflow_status_text", None)
        if status_text is not None and status_var is not None:
            status_var.set(status_text)
        cards = getattr(self, "filter_workflow_cards", {})
        active_index = stage_keys.index(stage_key)
        for index, (key, title, _description) in enumerate(FILTER_WORKFLOW_STEPS):
            button = cards.get(key)
            if button is None:
                continue
            if index < active_index:
                bg = self.palette["surface_soft"]
                fg = self.palette["status_success"]
                text = f"✓  {title}"
            elif index == active_index:
                bg = self.palette["primary"]
                fg = "#FFFFFF"
                text = f"{index + 1}  {title}"
            else:
                bg = self.palette["surface_raised"]
                fg = self.palette["muted"]
                text = f"{index + 1}  {title}"
            button.configure(
                text=text,
                bg=bg,
                fg=fg,
                activebackground=(
                    self.palette["primary_hover"]
                    if index == active_index
                    else self.palette["surface_soft"]
                ),
                activeforeground="#FFFFFF" if index == active_index else self.palette["text"],
                highlightthickness=1,
                highlightbackground=self.palette["border"],
                highlightcolor=self.palette["primary"],
            )

    def _scroll_filter_workflow_to(self, stage_key: str) -> str:
        focus_targets = {
            "input": getattr(self, "excel_path_entry", None),
            "rules": getattr(self, "workbook_analysis_toggle_btn", None),
            "preview": getattr(self, "filter_preview_btn", None),
            "execute": getattr(self, "filter_run_btn", None),
            "results": getattr(self, "filter_result_tree", None),
        }
        target = getattr(self, "filter_workflow_sections", {}).get(stage_key)
        if target is None and stage_key in {"preview", "execute"}:
            # Preview and execution are launched from the fixed action bar; their
            # output is rendered in the results section, so that is their scroll
            # destination while keyboard focus moves to the relevant action.
            target = getattr(self, "filter_workflow_sections", {}).get("results")
        canvas = getattr(self, "filter_scroll_canvas", None)
        panel = getattr(self, "filter_scroll_panel", None)
        if target is not None and canvas is not None and panel is not None:
            panel.update_idletasks()
            canvas.update_idletasks()
            scroll_region = canvas.bbox("all")
            if scroll_region is not None:
                content_top = scroll_region[1]
                content_height = max(scroll_region[3] - content_top, 1)
                viewport_height = max(canvas.winfo_height(), 1)
                max_offset = max(content_height - viewport_height, 0)
                target_offset = min(
                    max(target.winfo_y() - content_top - 8, 0),
                    max_offset,
                )
                # Canvas.yview_moveto expects a fraction of the complete
                # scrollregion, not a fraction of only the scrollable remainder.
                canvas.yview_moveto(target_offset / content_height)

        focus_target = focus_targets.get(stage_key)
        if focus_target is not None:
            try:
                focus_target.focus_set()
            except tk.TclError:
                pass
        return "break"

    def _build_filter_tab(self) -> None:
        workflow_shell = tk.Frame(self.filter_frame, bg=self.palette["root_bg"])
        workflow_shell.pack(fill="both", expand=True)
        self.filter_action_bar = tk.Frame(
            workflow_shell,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        self.filter_action_bar.pack(side="bottom", fill="x", pady=(6, 0))
        scroll_host = tk.Frame(workflow_shell, bg=self.palette["root_bg"])
        scroll_host.pack(side="top", fill="both", expand=True)
        panel = self._create_scrollable_tab_body(scroll_host)
        self.filter_scroll_panel = panel
        self.filter_scroll_canvas = panel.master
        self.filter_workflow_sections: Dict[str, tk.Frame] = {}
        if not PANDAS_SUPPORT:
            tk.Label(
                panel, text="⚠️ 此功能需要 pandas\n\n安装命令：python -m pip install pandas openpyxl",
                font=("微软雅黑", 12), fg=self.palette["danger"], justify="center",
            ).pack(pady=40)
            return

        self._build_filter_workflow_rail(panel)

        # 帮助
        self.help_visible = tk.BooleanVar(value=False)
        hbf = tk.Frame(panel)
        hbf.pack(fill="x", pady=(0, 6))
        self.help_btn = tk.Button(hbf, text="📖 显示使用说明", font=("微软雅黑", 9), command=self._toggle_help)
        self.help_btn.pack(side="left")
        self._style_action_button(self.help_btn, "secondary")

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
        self._build_filter_section_heading(panel, "input")
        self.file_path_frame = tk.LabelFrame(
            panel, text=" 📂 文件路径设置 ", font=("微软雅黑", 10, "bold"), padx=12, pady=10,
        )
        self.file_path_frame.pack(fill="x", pady=8)

        path_grid = tk.Frame(self.file_path_frame)
        path_grid.pack(fill="x")
        path_grid.grid_columnconfigure(1, weight=1)
        path_grid.grid_columnconfigure(4, weight=1)

        tk.Label(path_grid, text="Excel文件:", font=("微软雅黑", 9), width=10, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 4), pady=3)
        self.excel_path_entry = tk.Entry(path_grid, textvariable=self.excel_path, font=("微软雅黑", 9))
        self.excel_path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=3)
        self.excel_browse_btn = tk.Button(path_grid, text="浏览", command=self._browse_excel, padx=8)
        self.excel_browse_btn.grid(row=0, column=2, sticky="ew", padx=(0, 12), pady=3)
        self._style_action_button(self.excel_browse_btn, "neutral")

        tk.Label(path_grid, text="工作表:", font=("微软雅黑", 9), width=8, anchor="w").grid(row=0, column=3, sticky="w", padx=(0, 4), pady=3)
        self.excel_sheet_combo = ttk.Combobox(
            path_grid,
            textvariable=self.excel_sheet_name,
            state="readonly",
            font=("微软雅黑", 9),
        )
        self.excel_sheet_combo.grid(row=0, column=4, sticky="ew", padx=(0, 6), pady=3)
        self.excel_sheet_combo.bind("<<ComboboxSelected>>", self._on_excel_sheet_change)
        self.sheet_refresh_btn = tk.Button(path_grid, text="刷新", command=self._refresh_excel_sheets, padx=8)
        self.sheet_refresh_btn.grid(row=0, column=5, sticky="ew", pady=3)
        self._style_action_button(self.sheet_refresh_btn, "neutral")

        tk.Label(path_grid, text="PDF文件夹:", font=("微软雅黑", 9), width=10, anchor="w").grid(row=1, column=0, sticky="w", padx=(0, 4), pady=3)
        self.pdf_folder_entry = tk.Entry(path_grid, textvariable=self.pdf_folder, font=("微软雅黑", 9))
        self.pdf_folder_entry.grid(row=1, column=1, sticky="ew", padx=(0, 6), pady=3)
        self.pdf_browse_btn = tk.Button(path_grid, text="浏览", command=self._browse_pdf_folder, padx=8)
        self.pdf_browse_btn.grid(row=1, column=2, sticky="ew", padx=(0, 12), pady=3)
        self._style_action_button(self.pdf_browse_btn, "neutral")

        tk.Label(path_grid, text="导出文件夹:", font=("微软雅黑", 9), width=8, anchor="w").grid(row=1, column=3, sticky="w", padx=(0, 4), pady=3)
        self.output_folder_entry = tk.Entry(path_grid, textvariable=self.output_folder, font=("微软雅黑", 9))
        self.output_folder_entry.grid(row=1, column=4, sticky="ew", padx=(0, 6), pady=3)
        self.output_folder_browse_btn = tk.Button(path_grid, text="浏览", command=self._browse_output_folder, padx=8)
        self.output_folder_browse_btn.grid(row=1, column=5, sticky="ew", pady=3)
        self._style_action_button(self.output_folder_browse_btn, "neutral")

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

        self._build_filter_section_heading(panel, "rules")
        self._build_workbook_analysis_panel(panel)

        # 固定操作栏
        fbtn = self.filter_action_bar
        tk.Label(
            fbtn,
            textvariable=self.filter_workflow_status_text,
            font=("微软雅黑", 9),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
            wraplength=260,
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))

        controls = tk.Frame(fbtn, bg=self.palette["surface"])
        controls.pack(side="right")

        self.filter_preview_btn = tk.Button(
            controls, text="👁 预览匹配", font=("微软雅黑", 10),
            padx=15, pady=6, cursor="hand2", command=self._preview_filter,
        )
        self.filter_preview_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.filter_preview_btn, "secondary")

        self.filter_run_btn = tk.Button(
            controls, text="🚀 开始筛选并导出", font=("微软雅黑", 12, "bold"),
            padx=22, pady=7, cursor="hand2", command=self._run_filter,
        )
        self.filter_run_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.filter_run_btn, "primary")

        self.filter_retry_btn = tk.Button(
            controls,
            text="重试失败",
            font=("微软雅黑", 10),
            padx=11,
            pady=5,
            command=self._retry_failed_filter,
            state="disabled",
        )
        self.filter_retry_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.filter_retry_btn, "secondary")

        self.pause_filter_btn = tk.Button(
            controls,
            text="⏸ 暂停",
            font=("微软雅黑", 10),
            padx=10,
            pady=5,
            command=self._toggle_pause_task,
            state="disabled",
        )
        self.pause_filter_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.pause_filter_btn, "secondary")

        self.open_output_btn = tk.Button(controls, text="📂 打开导出文件夹", font=("微软雅黑", 10), padx=12, pady=5, command=self._open_output_folder)
        self.open_output_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.open_output_btn, "secondary")

        self.cancel_flt_btn = tk.Button(
            controls, text="⏹ 取消", font=("微软雅黑", 10),
            padx=12, pady=5, cursor="hand2",
            command=self._cancel_task, state="disabled",
        )
        self.cancel_flt_btn.pack(side="left")
        self._style_action_button(self.cancel_flt_btn, "secondary")

        self.filter_progress = ttk.Progressbar(controls, mode="determinate", length=150)
        self.filter_progress.pack(side="right")

        self._build_filter_section_heading(panel, "results")
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

        self.reset_filter_btn = tk.Button(tool_row, text="重置筛选条件", font=("微软雅黑", 9), padx=10, command=self._reset_filter_result_filters)
        self.reset_filter_btn.pack(side="left", padx=(0, 6))
        self._style_action_button(self.reset_filter_btn, "neutral")
        self.copy_missing_btn = tk.Button(tool_row, text="复制未匹配发票号", font=("微软雅黑", 9), padx=10, command=self._copy_missing_invoices, state="disabled")
        self.copy_missing_btn.pack(side="left", padx=(0, 6))
        self._style_action_button(self.copy_missing_btn, "secondary")
        self.open_result_btn = tk.Button(tool_row, text="打开选中结果", font=("微软雅黑", 9), padx=10, command=self._open_selected_filter_result, state="disabled")
        self.open_result_btn.pack(side="left")
        self._style_action_button(self.open_result_btn, "secondary")
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
        self.filter_result_tree.tag_configure("success", foreground=self.palette["status_success"])
        self.filter_result_tree.tag_configure("missing", foreground=self.palette["status_missing"])
        self.filter_result_tree.tag_configure("skip", foreground=self.palette["status_skip"])
        self.filter_result_tree.tag_configure("error", foreground=self.palette["status_error"])
        self.filter_result_tree.tag_configure("conflict", foreground=self.palette["status_conflict"])
        self.filter_result_tree.tag_configure("preview", foreground=self.palette["status_preview"])
        self.filter_result_tree.bind("<<TreeviewSelect>>", self._on_filter_result_select)
        self.filter_result_tree.bind("<Double-1>", self._open_selected_filter_result)

        detail_frame = tk.Frame(
            res_lf,
            bg=self.palette["detail_bg"],
            highlightthickness=1,
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"]
        )
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
        top_row.pack(fill="x")
        tk.Label(
            top_row,
            textvariable=self.workbook_analysis_summary_var,
            font=("微软雅黑", 9),
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
        ).pack(side="left", fill="x", expand=True)
        self.workbook_analysis_btn = tk.Button(top_row, text="分析工作簿", padx=10, command=self._refresh_workbook_analysis)
        self.workbook_analysis_btn.pack(side="right")
        self._style_action_button(self.workbook_analysis_btn, "secondary")
        self.workbook_analysis_toggle_btn = tk.Button(
            top_row,
            text="展开列映射 / 条件",
            padx=10,
            command=self._toggle_workbook_analysis_panel,
        )
        self.workbook_analysis_toggle_btn.pack(side="right", padx=(0, 8))
        self._style_action_button(self.workbook_analysis_toggle_btn, "neutral")

        compact_row = tk.Frame(analysis_lf)
        compact_row.pack(fill="x", pady=(6, 0))
        tk.Label(
            compact_row,
            textvariable=self.workbook_analysis_compact_var,
            font=("微软雅黑", 8),
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
        ).pack(side="left", fill="x", expand=True)

        self.workbook_analysis_content = tk.Frame(analysis_lf)
        content = self.workbook_analysis_content

        left_panel = tk.Frame(content)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_panel = tk.Frame(
            content,
            bg=self.palette["surface_alt"],
            highlightthickness=1,
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"]
        )
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
        self.workbook_sheet_tree.tag_configure("recommended", foreground=self.palette["status_success"])
        self.workbook_sheet_tree.tag_configure("usable", foreground=self.palette["status_preview"])
        self.workbook_sheet_tree.tag_configure("warning", foreground=self.palette["status_skip"])
        self.workbook_sheet_tree.tag_configure("error", foreground=self.palette["status_missing"])
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
        self._sync_workbook_analysis_panel_visibility()

    def _sync_workbook_analysis_panel_visibility(self) -> None:
        content = getattr(self, "workbook_analysis_content", None)
        if content is None:
            return
        if self.workbook_analysis_expanded.get():
            if not content.winfo_manager():
                content.pack(fill="x", pady=(8, 0))
            button_text = "收起列映射 / 条件"
        else:
            content.pack_forget()
            button_text = "展开列映射 / 条件"
        toggle_btn = getattr(self, "workbook_analysis_toggle_btn", None)
        if toggle_btn is not None:
            toggle_btn.config(text=button_text)

    def _toggle_workbook_analysis_panel(self) -> None:
        self.workbook_analysis_expanded.set(not self.workbook_analysis_expanded.get())
        self._sync_workbook_analysis_panel_visibility()

    # ─────────────── Tab3: 历史记录 ───────────────

    def _build_history_tab(self) -> None:
        workflow_shell = tk.Frame(self.history_frame, bg=self.palette["root_bg"])
        workflow_shell.pack(fill="both", expand=True)
        self.history_action_bar = tk.Frame(
            workflow_shell,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        self.history_action_bar.pack(side="bottom", fill="x", pady=(6, 0))
        body = tk.Frame(workflow_shell, bg=self.palette["root_bg"])
        body.pack(side="top", fill="both", expand=True)

        tk.Label(
            body,
            text="历史记录用于追踪已完成文件操作；回滚前会重新校验文件内容，不会盲目覆盖或删除。",
            font=("微软雅黑", 10), fg=self.palette["muted"],
            bg=self.palette["root_bg"],
        ).pack(anchor="w", pady=(0, 8))

        filter_bar = tk.Frame(
            body,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        filter_bar.pack(fill="x", pady=(0, 8))
        tk.Label(filter_bar, text="类型:", font=("微软雅黑", 9), bg=self.palette["surface"], fg=self.palette["text"]).pack(side="left")
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

        tk.Label(filter_bar, text="时间:", font=("微软雅黑", 9), bg=self.palette["surface"], fg=self.palette["text"]).pack(side="left")
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

        tk.Label(filter_bar, text="搜索:", font=("微软雅黑", 9), bg=self.palette["surface"], fg=self.palette["text"]).pack(side="left")
        self.history_keyword_entry = tk.Entry(filter_bar, textvariable=self.history_keyword, font=("微软雅黑", 9))
        self.history_keyword_entry.pack(side="left", fill="x", expand=True, padx=(6, 10))
        self.history_keyword_entry.bind("<KeyRelease>", self._on_history_filters_changed)

        self.history_reset_btn = tk.Button(filter_bar, text="重置", font=("微软雅黑", 9), padx=10, command=self._reset_history_filters)
        self.history_reset_btn.pack(side="left")
        self._style_action_button(self.history_reset_btn, "neutral")
        tk.Label(
            filter_bar,
            textvariable=self.history_summary_var,
            font=("微软雅黑", 9),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
        ).pack(side="right", padx=(10, 0))

        self.history_split = tk.Frame(body, bg=self.palette["root_bg"])
        self.history_split.pack(fill="both", expand=True)
        self.history_detail_panel = tk.Frame(
            self.history_split,
            width=310,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        self.history_detail_panel.pack(side="right", fill="y", padx=(8, 0))
        self.history_detail_panel.pack_propagate(False)
        tf = tk.Frame(
            self.history_split,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=1,
            pady=1,
        )
        tf.pack(side="left", fill="both", expand=True)

        cols = ("time", "folder", "count", "type")
        self.history_tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        self.history_tree.heading("time", text="时间")
        self.history_tree.heading("folder", text="操作文件夹")
        self.history_tree.heading("count", text="文件数")
        self.history_tree.heading("type", text="类型")
        self.history_tree.column("time", width=150)
        self.history_tree.column("folder", width=300)
        self.history_tree.column("count", width=80, anchor="center")
        self.history_tree.column("type", width=100, anchor="center")
        self.history_tree.tag_configure("evenrow", background=self.palette["tree_even"])
        self.history_tree.tag_configure("oddrow", background=self.palette["tree_odd"])

        hscr = ttk.Scrollbar(tf, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=hscr.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        hscr.pack(side="right", fill="y")
        self.history_tree.bind("<<TreeviewSelect>>", self._on_history_select)
        self.history_tree.bind("<Double-1>", lambda event: self._open_history_folder())
        self.history_tree.bind("<Return>", lambda event: self._view_history_detail())

        tk.Label(
            self.history_detail_panel,
            text="选中任务",
            font=("微软雅黑", 8, "bold"),
            bg=self.palette["surface"],
            fg=self.palette["primary"],
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            self.history_detail_panel,
            textvariable=self.history_detail_title,
            font=("微软雅黑", 13, "bold"),
            bg=self.palette["surface"],
            fg=self.palette["text"],
            anchor="w",
            justify="left",
            wraplength=280,
        ).pack(fill="x", pady=(6, 3))
        tk.Label(
            self.history_detail_panel,
            textvariable=self.history_detail_meta,
            font=("微软雅黑", 9),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
            wraplength=280,
        ).pack(fill="x")
        tk.Label(
            self.history_detail_panel,
            text="操作目录",
            font=("微软雅黑", 8, "bold"),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
            anchor="w",
        ).pack(fill="x", pady=(12, 2))
        tk.Label(
            self.history_detail_panel,
            textvariable=self.history_detail_folder,
            font=("微软雅黑", 8),
            bg=self.palette["surface_soft"],
            fg=self.palette["text"],
            anchor="w",
            justify="left",
            wraplength=270,
            padx=8,
            pady=6,
        ).pack(fill="x")
        self.history_safety_label = tk.Label(
            self.history_detail_panel,
            textvariable=self.history_detail_safety,
            font=("微软雅黑", 9, "bold"),
            bg=self.palette["surface_soft"],
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
            wraplength=270,
            padx=8,
            pady=7,
        )
        self.history_safety_label.pack(fill="x", pady=(8, 10))
        tk.Label(
            self.history_detail_panel,
            text="记录内容",
            font=("微软雅黑", 8, "bold"),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
            anchor="w",
        ).pack(fill="x", pady=(0, 4))
        preview_frame = tk.Frame(self.history_detail_panel, bg=self.palette["surface"])
        preview_frame.pack(fill="both", expand=True)
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical")
        preview_scroll.pack(side="right", fill="y")
        self.history_preview_listbox = tk.Listbox(
            preview_frame,
            font=("微软雅黑", 8),
            bg=self.palette["surface_alt"],
            fg=self.palette["text"],
            relief="flat",
            activestyle="none",
            yscrollcommand=preview_scroll.set,
            takefocus=False,
        )
        self.history_preview_listbox.pack(side="left", fill="both", expand=True)
        preview_scroll.config(command=self.history_preview_listbox.yview)

        hbtn = self.history_action_bar
        tk.Label(
            hbtn,
            textvariable=self.history_action_status_text,
            font=("微软雅黑", 8),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
            wraplength=250,
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))
        controls = tk.Frame(hbtn, bg=self.palette["surface"])
        controls.pack(side="right")

        self.history_rollback_btn = tk.Button(
            controls,
            text="🔄 安全回滚",
            font=("微软雅黑", 10),
            padx=12,
            pady=5,
            command=self._rollback_selected,
            state="disabled",
        )
        self.history_rollback_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.history_rollback_btn, "warning")

        self.history_rerun_btn = tk.Button(
            controls,
            text="再次执行",
            font=("微软雅黑", 10),
            padx=11,
            pady=5,
            command=self._load_history_for_rerun,
            state="disabled",
        )
        self.history_rerun_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.history_rerun_btn, "secondary")

        self.history_view_btn = tk.Button(controls, text="🔍 完整详情", font=("微软雅黑", 10), padx=12, pady=5, command=self._view_history_detail, state="disabled")
        self.history_view_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.history_view_btn, "secondary")

        self.history_open_btn = tk.Button(controls, text="📂 打开目录", font=("微软雅黑", 10), padx=12, pady=5, command=self._open_history_folder, state="disabled")
        self.history_open_btn.pack(side="left", padx=(0, 8))
        self._style_action_button(self.history_open_btn, "secondary")

        self.history_clear_btn = tk.Button(controls, text="清空历史", font=("微软雅黑", 9), padx=10, pady=5, command=self._clear_all_history)
        self.history_clear_btn.pack(side="left")
        self._style_action_button(self.history_clear_btn, "danger")

        self.history_refresh_btn = tk.Button(controls, text="刷新", font=("微软雅黑", 9), padx=10, pady=5, command=self._refresh_history_tree)
        self.history_refresh_btn.pack(side="left", padx=(8, 0))
        self._style_action_button(self.history_refresh_btn, "neutral")

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
        border_color = self._mix_colors(bg, fg, 0.15)
        card = tk.Frame(
            parent,
            bg=bg,
            bd=0,
            highlightthickness=1,
            highlightbackground=border_color,
            highlightcolor=border_color,
            padx=12,
            pady=10,
        )
        card.pack(side="left", fill="x", expand=True, padx=4)
        tk.Label(
            card,
            textvariable=self.filter_metric_labels[metric_key],
            font=("微软雅黑", 8),
            bg=bg,
            fg=self._mix_colors(bg, fg, 0.7),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            card,
            textvariable=self.filter_metric_values[metric_key],
            font=("微软雅黑", 18, "bold"),
            bg=bg,
            fg=fg,
            anchor="w",
        ).pack(anchor="w", pady=(6, 0))

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
        if hasattr(self, "filter_retry_btn"):
            self.filter_retry_btn.config(state="disabled")
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
        if hasattr(self, "filter_retry_btn"):
            retry_count = sum(1 for row in self.filter_result_rows if row.status == "复制失败")
            self.filter_retry_btn.config(state="normal" if retry_count else "disabled")
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
                "同名冲突": "conflict",
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

    @staticmethod
    def _config_change_preview(plan: ConfigPlan, limit: int = 12) -> str:
        labels = {
            "ui_theme": "界面主题",
            "workspace_page": "默认任务页",
            "rule_preset_id": "规则预设",
            "company_name_index": "公司名称段位",
            "invoice_number_index": "发票号码段位",
            "organize_folder": "整理目录",
            "excel_path": "Excel 文件",
            "pdf_folder": "PDF 目录",
            "excel_sheet_name": "工作表",
            "auto_output_by_sheet": "自动输出目录",
            "output_folder": "手动输出目录",
            "invoice_column_aliases": "发票列别名",
            "company_column_aliases": "公司列别名",
        }

        def compact(value: Any) -> str:
            text = repr(value)
            return text if len(text) <= 70 else text[:67] + "..."

        lines = [
            f"• {labels.get(change.key, change.key)}：{compact(change.old_value)} → {compact(change.new_value)}"
            for change in plan.changes[:limit]
        ]
        if len(plan.changes) > limit:
            lines.append(f"• 另有 {len(plan.changes) - limit} 项变更")
        if plan.warnings:
            lines.append("")
            lines.append("兼容提示：")
            lines.extend(f"• {warning}" for warning in plan.warnings[:6])
            if len(plan.warnings) > 6:
                lines.append(f"• 另有 {len(plan.warnings) - 6} 条提示")
        return "\n".join(lines) if lines else "没有配置项发生变化。"

    def _apply_config_to_runtime(self, config: Dict[str, Any]) -> None:
        self.config = dict(config)
        theme_id = str(config.get("ui_theme", "day")).strip().lower()
        if theme_id not in UI_THEME_PRESETS:
            theme_id = "day"
        self.ui_theme.set(theme_id)
        self.ui_theme_label.set(self._theme_label(theme_id))
        self.palette = UI_THEME_PRESETS[theme_id]

        preset_id = str(config.get("rule_preset_id", DEFAULT_RULE_PRESET_ID)).strip()
        if preset_id not in self._preset_by_id:
            preset_id = DEFAULT_RULE_PRESET_ID
        self.rule_preset_id.set(preset_id)
        self._sync_rule_preset_ui()
        self.company_name_index.set(int(config.get("company_name_index", 2)))
        self.invoice_number_index.set(int(config.get("invoice_number_index", 1)))
        self.invoice_column_aliases.set(str(config.get("invoice_column_aliases", "")))
        self.company_column_aliases.set(str(config.get("company_column_aliases", "")))

        self.organize_folder_path.set(str(config.get("organize_folder", "")))
        self.excel_path.set(str(config.get("excel_path", "")))
        self.excel_sheet_name.set(str(config.get("excel_sheet_name", "")))
        self.pdf_folder.set(str(config.get("pdf_folder", "")))
        self.selected_invoice_column_name.set(str(config.get("selected_invoice_column_name", "")))
        self.selected_company_column_name.set(str(config.get("selected_company_column_name", "")))
        self.row_filter_column_name.set(str(config.get("row_filter_column_name", "")))
        self.row_filter_mode.set(str(config.get("row_filter_mode", "不过滤")))
        self.row_filter_values.set(str(config.get("row_filter_values", "")))
        self.company_exclude_keywords.set(str(config.get("company_exclude_keywords", "")))
        self.auto_output_by_sheet.set(bool(config.get("auto_output_by_sheet", True)))
        self.manual_output_folder.set(str(config.get("output_folder", "")))
        self._active_filter_context = (self.excel_path.get().strip(), self.excel_sheet_name.get())
        self._sync_output_folder_mode_ui()

    def _commit_config_plan(self, plan: ConfigPlan, action_name: str) -> Path:
        if self._task_is_running():
            raise RuntimeError(f"当前任务结束后才能执行{action_name}")
        previous_config = self._collect_runtime_config()
        previous_block_reason = self._config_write_blocked_reason
        previous_blocked_snapshot = self._blocked_config_snapshot
        backup_source = self._blocked_config_snapshot or previous_config
        backup_path = backup_config(backup_source, self._config_file.parent / "backups")
        try:
            self._config_write_blocked_reason = ""
            self._blocked_config_snapshot = None
            self._apply_config_to_runtime(plan.config)
            if not self._save_config():
                raise OSError("配置文件写入失败")
        except Exception:
            self._config_write_blocked_reason = previous_block_reason
            self._blocked_config_snapshot = previous_blocked_snapshot
            self._apply_config_to_runtime(previous_config)
            raise

        self._rebuild_ui()
        self.settings_status_text.set(f"{action_name}完成；原配置已备份到 {backup_path.name}")
        self.status_var.set(f"✅ {action_name}完成，配置备份：{backup_path.name}")
        return backup_path

    def _export_settings_config(self) -> None:
        target = filedialog.asksaveasfilename(
            title="导出配置",
            defaultextension=".json",
            initialfile="invoice-tool-config.json",
            filetypes=[("JSON 配置", "*.json")],
        )
        if not target:
            return
        config = self._collect_runtime_config()
        if not save_config_export(Path(target), config):
            messagebox.showerror("错误", "配置导出失败，请检查目标目录权限。")
            return
        self.settings_status_text.set(f"配置已导出：{Path(target).name}")
        logger.info("配置已导出：%s", target)
        messagebox.showinfo("完成", f"配置已导出到：\n{target}")

    def _import_settings_config(self) -> None:
        if not self._require_idle("导入配置"):
            return
        source = filedialog.askopenfilename(
            title="导入配置",
            filetypes=[("JSON 配置", "*.json"), ("所有文件", "*.*")],
        )
        if not source:
            return
        try:
            current = self._collect_runtime_config()
            plan = load_config_plan(
                Path(source),
                current,
                preset_ids=self._preset_by_id.keys(),
            )
        except ConfigurationError as exc:
            messagebox.showerror("配置不可用", str(exc))
            return
        preview = self._config_change_preview(plan)
        if not plan.changes:
            messagebox.showinfo("无需导入", preview)
            return
        if not messagebox.askyesno(
            "确认导入配置",
            f"将应用 {len(plan.changes)} 项配置变更。应用前会自动备份当前配置。\n\n{preview}",
        ):
            return
        try:
            backup_path = self._commit_config_plan(plan, "配置导入")
        except (OSError, RuntimeError) as exc:
            logger.exception("配置导入失败")
            messagebox.showerror("错误", str(exc))
            return
        logger.info("配置导入完成，备份：%s", backup_path)
        messagebox.showinfo("完成", f"配置已导入。\n原配置备份：{backup_path}")

    def _restore_default_settings(self) -> None:
        if not self._require_idle("恢复默认配置"):
            return
        current = self._collect_runtime_config()
        plan = default_config_plan(current)
        if not plan.changes:
            messagebox.showinfo("提示", "当前已是默认配置。")
            return
        preview = self._config_change_preview(plan)
        if not messagebox.askyesno(
            "恢复默认配置",
            f"将恢复 {len(plan.changes)} 项默认设置。当前配置会先自动备份。\n\n{preview}",
        ):
            return
        try:
            backup_path = self._commit_config_plan(plan, "恢复默认配置")
        except (OSError, RuntimeError) as exc:
            logger.exception("恢复默认配置失败")
            messagebox.showerror("错误", str(exc))
            return
        logger.info("默认配置已恢复，备份：%s", backup_path)
        messagebox.showinfo("完成", f"已恢复默认配置。\n原配置备份：{backup_path}")

    def _export_diagnostic_bundle(self) -> None:
        target = filedialog.asksaveasfilename(
            title="导出脱敏诊断包",
            defaultextension=".zip",
            initialfile=f"invoice-tool-diagnostics-{datetime.now():%Y%m%d-%H%M%S}.zip",
            filetypes=[("ZIP 诊断包", "*.zip")],
        )
        if not target:
            return
        snapshot = build_diagnostic_snapshot(
            app_version=APP_VERSION,
            capabilities={
                "pandas": PANDAS_SUPPORT,
                "openpyxl": OPENPYXL_SUPPORT,
                "xlrd": XLRD_SUPPORT,
                "drag_drop": DND_SUPPORT,
                "ttkbootstrap": MODERN_UI,
            },
            config_path=self._config_file,
            history_path=self._history_file,
            log_path=LOG_FILE,
            recent_errors=self.recent_errors,
            config_schema_version=CONFIG_SCHEMA_VERSION,
            config_directory_fallback_reason=CONFIG_DIR_FALLBACK_REASON,
        )
        try:
            bundle_path = create_diagnostic_bundle(Path(target), snapshot, log_path=LOG_FILE)
        except Exception as exc:
            logger.exception("诊断包导出失败")
            messagebox.showerror("导出失败", f"无法写入诊断包：\n{exc}")
            return
        self.settings_status_text.set(f"脱敏诊断包已导出：{bundle_path.name}")
        logger.info("脱敏诊断包已导出")
        messagebox.showinfo(
            "诊断包已导出",
            f"已生成：\n{bundle_path}\n\n"
            "诊断包不包含原始配置、历史记录、Excel 或 PDF；日志已尝试脱敏。"
            "发送给他人前仍请检查其中内容。",
        )

    @staticmethod
    def _show_release_notes() -> None:
        messagebox.showinfo(f"发票处理工具箱 {APP_VERSION}", RELEASE_SUMMARY)

    def _build_settings_tab(self) -> None:
        workflow_shell = tk.Frame(self.settings_frame, bg=self.palette["root_bg"])
        workflow_shell.pack(fill="both", expand=True)
        self.settings_action_bar = tk.Frame(
            workflow_shell,
            bg=self.palette["surface"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        self.settings_action_bar.pack(side="bottom", fill="x", pady=(6, 0))
        scroll_host = tk.Frame(workflow_shell, bg=self.palette["root_bg"])
        scroll_host.pack(side="top", fill="both", expand=True)
        panel = self._create_scrollable_tab_body(scroll_host)
        self.settings_scroll_host = scroll_host
        self.settings_scroll_outer = panel.master.master
        self.settings_scroll_panel = panel
        container = tk.Frame(panel, bg=self.palette["root_bg"])
        container.pack(fill="both", expand=True)

        left_panel = tk.Frame(container, bg=self.palette["root_bg"])
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_panel = tk.Frame(container, width=360, bg=self.palette["root_bg"])
        right_panel.pack(side="right", fill="both")
        right_panel.pack_propagate(False)

        tk.Label(left_panel, text="⚙️ 设置中心", font=("微软雅黑", 11, "bold"), fg=self.palette["text"]).pack(anchor="w", pady=(0, 5))
        tk.Label(
            left_panel,
            text=f"配置格式 v{CONFIG_SCHEMA_VERSION} · 导入和恢复默认前自动备份 · 未识别的新字段会保留",
            font=("微软雅黑", 8),
            fg=self.palette["muted"],
            anchor="w",
        ).pack(fill="x", pady=(0, 12))

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
        self.preset_apply_btn = tk.Button(preset_row, text="应用预设", padx=10, command=self._apply_rule_preset)
        self.preset_apply_btn.pack(side="right")
        self._style_action_button(self.preset_apply_btn, "secondary")

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

        action_frame = tk.LabelFrame(left_panel, text="配置与环境", padx=12, pady=10, font=("微软雅黑", 10, "bold"))
        action_frame.pack(fill="x")
        version_row = tk.Frame(action_frame)
        version_row.pack(fill="x", pady=(0, 6))
        tk.Label(
            version_row,
            text=f"应用版本：{APP_VERSION}",
            font=("微软雅黑", 9, "bold"),
            anchor="w",
        ).pack(side="left")
        self.release_notes_btn = tk.Button(
            version_row,
            text="查看更新说明",
            font=("微软雅黑", 8),
            padx=8,
            command=self._show_release_notes,
        )
        self.release_notes_btn.pack(side="right")
        self._style_action_button(self.release_notes_btn, "secondary")
        tk.Label(action_frame, text=f"配置目录：{self._config_file.parent}", font=("微软雅黑", 8), fg=self.palette["muted"], anchor="w", justify="left").pack(fill="x")
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

        diagnostic_title_bar = tk.Frame(right_panel, bg=self.palette["root_bg"])
        diagnostic_title_bar.pack(fill="x", pady=(0, 12))
        tk.Label(
            diagnostic_title_bar,
            text="🩺 诊断中心",
            font=("微软雅黑", 11, "bold"),
            fg=self.palette["text"],
            bg=self.palette["root_bg"],
        ).pack(side="left")
        self.diagnostic_export_btn = tk.Button(
            diagnostic_title_bar,
            text="导出脱敏诊断包",
            font=("微软雅黑", 8),
            padx=8,
            command=self._export_diagnostic_bundle,
        )
        self.diagnostic_export_btn.pack(side="right")
        self._style_action_button(self.diagnostic_export_btn, "secondary")

        recent_frame = tk.LabelFrame(right_panel, text="最近错误", padx=10, pady=10, font=("微软雅黑", 10, "bold"))
        recent_frame.pack(fill="both", expand=True)

        top_bar = tk.Frame(recent_frame)
        top_bar.pack(fill="x", pady=(0, 8))
        tk.Label(top_bar, textvariable=self.recent_error_summary_var, font=("微软雅黑", 9), fg=self.palette["muted"]).pack(side="left")
        self.error_copy_btn = tk.Button(top_bar, text="复制", font=("微软雅黑", 8), padx=8, command=self._copy_selected_recent_error)
        self.error_copy_btn.pack(side="right")
        self._style_action_button(self.error_copy_btn, "secondary")

        self.error_clear_btn = tk.Button(top_bar, text="清空", font=("微软雅黑", 8), padx=8, command=self._clear_recent_errors)
        self.error_clear_btn.pack(side="right", padx=(0, 6))
        self._style_action_button(self.error_clear_btn, "warning")

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

        action_bar = self.settings_action_bar
        if self._config_write_blocked_reason:
            self.settings_status_text.set(f"配置版本不兼容，写入已保护：{self._config_write_blocked_reason}")
        tk.Label(
            action_bar,
            textvariable=self.settings_status_text,
            font=("微软雅黑", 8),
            bg=self.palette["surface"],
            fg=self.palette["muted"],
            anchor="w",
            justify="left",
            wraplength=250,
        ).pack(side="left", fill="x", expand=True, padx=(0, 10))
        controls = tk.Frame(action_bar, bg=self.palette["surface"])
        controls.pack(side="right")

        self.settings_save_btn = tk.Button(
            controls,
            text="💾 保存",
            font=("微软雅黑", 10, "bold"),
            padx=12,
            pady=5,
            command=self._save_settings,
        )
        self.settings_save_btn.pack(side="left", padx=(0, 6))
        self._style_action_button(self.settings_save_btn, "success")
        self.settings_import_btn = tk.Button(controls, text="导入", font=("微软雅黑", 9), padx=10, pady=5, command=self._import_settings_config)
        self.settings_import_btn.pack(side="left", padx=(0, 6))
        self._style_action_button(self.settings_import_btn, "secondary")
        self.settings_export_btn = tk.Button(controls, text="导出", font=("微软雅黑", 9), padx=10, pady=5, command=self._export_settings_config)
        self.settings_export_btn.pack(side="left", padx=(0, 6))
        self._style_action_button(self.settings_export_btn, "secondary")
        self.settings_reset_btn = tk.Button(controls, text="恢复默认", font=("微软雅黑", 9), padx=10, pady=5, command=self._restore_default_settings)
        self.settings_reset_btn.pack(side="left", padx=(0, 10))
        self._style_action_button(self.settings_reset_btn, "warning")
        self.open_config_btn = tk.Button(controls, text="配置目录", font=("微软雅黑", 9), padx=10, pady=5, command=self._open_config_directory)
        self.open_config_btn.pack(side="left", padx=(0, 6))
        self._style_action_button(self.open_config_btn, "neutral")
        self.open_log_btn = tk.Button(controls, text="日志文件", font=("微软雅黑", 9), padx=10, pady=5, command=self._open_log_file)
        self.open_log_btn.pack(side="left")
        self._style_action_button(self.open_log_btn, "neutral")

    def _save_settings(self) -> None:
        if not self._require_idle("保存设置"):
            return
        try:
            c_idx = self.company_name_index.get()
            i_idx = self.invoice_number_index.get()
            if c_idx < 0 or i_idx < 0:
                raise ValueError("索引不能为负数")
        except (tk.TclError, ValueError):
            messagebox.showerror("错误", "保存失败：段位索引必须是大于或等于 0 的整数！")
            return

        if not self._save_config():
            reason = self._config_write_blocked_reason or "无法写入配置文件"
            self.settings_status_text.set(f"保存失败：{reason}")
            messagebox.showerror("保存失败", f"设置未写入：{reason}\n\n可通过“导入”或“恢复默认”创建兼容配置。")
            return
        idx = self._get_safe_company_name_index()
        self.organize_hint.config(text=f"  💡 公司名在第{idx+1}段（可在设置中修改）")
        self._refresh_excel_sheets(silent=True)
        self.settings_status_text.set("设置已保存；命名和 Excel 识别规则将在下次扫描或分析时生效。")
        logger.info("✅ 设置已保存")
        messagebox.showinfo("提示", "设置已保存，请重新扫描文件使新设置生效。")

    # ─────────────── 拖拽 ───────────────

    def _setup_drag_and_drop(self) -> None:
        if DND_SUPPORT:
            try:
                self.organize_folder_entry.drop_target_register(DND_FILES)
                self.organize_folder_entry.dnd_bind("<<Drop>>", self._on_organize_drop)

                # 筛选 Tab
                self.excel_path_entry.drop_target_register(DND_FILES)
                self.excel_path_entry.dnd_bind("<<Drop>>", self._on_excel_drop)

                self.pdf_folder_entry.drop_target_register(DND_FILES)
                self.pdf_folder_entry.dnd_bind("<<Drop>>", self._on_pdf_drop)

                self.output_folder_entry.drop_target_register(DND_FILES)
                self.output_folder_entry.dnd_bind("<<Drop>>", self._on_output_drop)
                logger.info("✅ 拖拽功能已启用")
            except Exception as e:
                logger.warning(f"拖拽初始化失败：{e}（不影响其他功能）")
        else:
            logger.warning("拖拽未启用（需 tkinterdnd2）")

    def _parse_dnd_event_paths(self, event_data: str) -> List[str]:
        paths = []
        for m in re.finditer(r"\{([^}]+)\}|(\S+)", event_data):
            p = (m.group(1) or m.group(2)).strip("\"'")
            paths.append(p)
        return paths

    def _on_organize_drop(self, event) -> None:
        paths = self._parse_dnd_event_paths(event.data)
        if not paths:
            return
        p = Path(paths[0])
        folder = p if p.is_dir() else p.parent
        self.organize_folder_path.set(str(folder))
        logger.info(f"📂 拖入整理文件夹：{folder}")
        self._scan_files()

    def _on_excel_drop(self, event) -> None:
        paths = self._parse_dnd_event_paths(event.data)
        if not paths:
            return
        p = Path(paths[0])
        if p.is_file() and p.suffix.lower() in (".xlsx", ".xls"):
            self.excel_path.set(str(p))
            logger.info(f"📂 拖入 Excel 文件：{p}")
            self._refresh_excel_sheets()
        else:
            messagebox.showwarning("提示", "请拖入有效的 Excel 文件（.xlsx 或 .xls）")

    def _on_pdf_drop(self, event) -> None:
        paths = self._parse_dnd_event_paths(event.data)
        if not paths:
            return
        p = Path(paths[0])
        folder = p if p.is_dir() else p.parent
        self.pdf_folder.set(str(folder))
        logger.info(f"📂 拖入 PDF 文件夹：{folder}")

    def _on_output_drop(self, event) -> None:
        paths = self._parse_dnd_event_paths(event.data)
        if not paths:
            return
        p = Path(paths[0])
        folder = p if p.is_dir() else p.parent
        if self.auto_output_by_sheet.get():
            self.auto_output_by_sheet.set(False)
            self._sync_output_folder_mode_ui()
        self.manual_output_folder.set(str(folder))
        self.output_folder.set(str(folder))
        logger.info(f"📂 拖入导出文件夹：{folder}")
        self._save_config()

    # ─────────────── 进度工具 ───────────────

    def _post_ui(self, callback: Callable[[], None]) -> None:
        if threading.current_thread() is threading.main_thread():
            callback()
            return
        self._ui_events.put(callback)

    def _drain_ui_events(self) -> None:
        self._ui_event_pump_id = None
        while True:
            try:
                callback = self._ui_events.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception:
                logger.exception("处理后台任务 UI 事件失败")
        if not self._close_finalized:
            self._ui_event_pump_id = self.root.after(25, self._drain_ui_events)

    def _update_progress_info(self, cur: int, total: int) -> None:
        pct = cur * 100 // max(total, 1)
        text = f"{cur}/{total} ({pct}%)"
        def _do():
            self.progress_label.config(text=text)
        self._post_ui(_do)

    def _update_progress(self, bar: ttk.Progressbar, value: int, maximum: Optional[int] = None) -> None:
        def _do():
            if maximum is not None:
                bar["maximum"] = maximum
            bar["value"] = value
        self._post_ui(_do)

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
            self._active_task_kind = "write"
            self._readonly_task_name = ""
        start_btn.config(state="disabled", text=busy_text)
        if busy_bg is not None:
            start_btn.config(bg=busy_bg, activebackground=busy_bg)
        cancel_btn.config(state="normal")
        self._pause_flag.clear()
        self.status_var.set("⏳ 任务进行中...")
        return True

    def _start_worker(self, target, args: Tuple[Any, ...], *, name: str) -> None:
        worker = threading.Thread(target=target, args=args, name=name, daemon=False)
        with self._lock:
            self._worker_thread = worker
        worker.start()

    def _task_is_running(self) -> bool:
        with self._lock:
            return bool(self.is_running)

    def _require_idle(self, action_name: str) -> bool:
        if not self._task_is_running():
            return True
        self.status_var.set(f"⚠ 当前任务结束后才能{action_name}。")
        messagebox.showwarning("任务进行中", f"请先等待当前任务结束或取消，然后再{action_name}。")
        return False

    def _try_begin_readonly_task(
        self,
        *,
        task_name: str,
        controls: List[Tuple[tk.Widget, Optional[str]]],
        cancel_button: tk.Button,
        progress_bar: ttk.Progressbar,
        busy_status: str,
        silent_busy: bool = False,
    ) -> Optional[int]:
        task_already_running = False
        with self._lock:
            if self.is_running:
                task_already_running = True
                token = 0
            else:
                self.is_running = True
                self._active_task_kind = "readonly"
                self._readonly_task_name = task_name
                self._readonly_task_sequence += 1
                token = self._readonly_task_sequence
        if task_already_running:
            if not silent_busy:
                messagebox.showwarning("提示", "已有任务正在运行，请等待完成或先取消当前任务。")
            return None

        snapshots: List[Tuple[tk.Widget, str, str]] = []
        for widget, busy_text in controls:
            try:
                snapshots.append((widget, str(widget.cget("state")), str(widget.cget("text"))))
                config: Dict[str, Any] = {"state": "disabled"}
                if busy_text is not None:
                    config["text"] = busy_text
                widget.config(**config)
            except (AttributeError, tk.TclError):
                continue
        self._readonly_task_controls = snapshots
        self._readonly_task_cancel_button = cancel_button
        self._readonly_task_progress_bar = progress_bar
        try:
            self._readonly_task_progress_mode = str(progress_bar.cget("mode"))
            progress_bar.config(mode="indeterminate")
            progress_bar.start(12)
        except tk.TclError:
            pass
        cancel_button.config(state="normal")
        self._cancel_flag.clear()
        self.progress_label.config(text=task_name)
        self.status_var.set(busy_status)
        return token

    def _finish_readonly_task_ui(self, token: int) -> None:
        with self._lock:
            if token != self._readonly_task_sequence or self._active_task_kind != "readonly":
                return
            controls = self._readonly_task_controls
            cancel_button = self._readonly_task_cancel_button
            progress_bar = self._readonly_task_progress_bar
            progress_mode = self._readonly_task_progress_mode
            self.is_running = False
            self._worker_thread = None
            self._active_task_kind = ""
            self._readonly_task_name = ""
            self._readonly_task_controls = []
            self._readonly_task_cancel_button = None
            self._readonly_task_progress_bar = None

        for widget, state, text in controls:
            try:
                if widget.winfo_exists():
                    widget.config(state=state, text=text)
            except (AttributeError, tk.TclError):
                continue
        if cancel_button is not None:
            try:
                if cancel_button.winfo_exists():
                    cancel_button.config(state="disabled")
            except (AttributeError, tk.TclError):
                pass
        if progress_bar is not None:
            try:
                if progress_bar.winfo_exists():
                    progress_bar.stop()
                    progress_bar.config(mode=progress_mode)
                    progress_bar["value"] = 0
            except (AttributeError, tk.TclError):
                pass
        self.progress_label.config(text="")

    def _start_readonly_task(
        self,
        *,
        task_name: str,
        worker_name: str,
        work: Callable[[], Any],
        on_success: Callable[[Any], None],
        on_error: Callable[[Exception], None],
        on_cancel: Callable[[], None],
        controls: List[Tuple[tk.Widget, Optional[str]]],
        cancel_button: tk.Button,
        progress_bar: ttk.Progressbar,
        busy_status: str,
        silent_busy: bool = False,
    ) -> bool:
        token = self._try_begin_readonly_task(
            task_name=task_name,
            controls=controls,
            cancel_button=cancel_button,
            progress_bar=progress_bar,
            busy_status=busy_status,
            silent_busy=silent_busy,
        )
        if token is None:
            return False

        def complete(*, result: Any = None, error: Optional[Exception] = None, cancelled: bool = False) -> None:
            try:
                if not self._closing_requested:
                    if cancelled:
                        on_cancel()
                    elif error is not None:
                        on_error(error)
                    else:
                        on_success(result)
            finally:
                self._finish_readonly_task_ui(token)

        def run() -> None:
            try:
                result = work()
            except CancelledError:
                self._post_ui(lambda: complete(cancelled=True))
            except Exception as exc:
                logger.exception("%s失败", task_name)
                self._post_ui(lambda error=exc: complete(error=error))
            else:
                if self._cancel_flag.is_set():
                    self._post_ui(lambda: complete(cancelled=True))
                else:
                    self._post_ui(lambda result=result: complete(result=result))

        self._start_worker(run, (), name=worker_name)
        return True

    def _record_active_task_move(self, move: Dict[str, Any]) -> None:
        task_id = self._active_task_id
        if task_id and not self._task_journal.record_move(task_id, move):
            raise OSError("任务恢复日志写入失败，已停止当前文件操作")

    @staticmethod
    def _build_report_entry(report_path: Path, output_root: Path) -> Dict[str, Any]:
        return {
            "path": str(report_path),
            "filename": report_path.name,
            "output_root": str(output_root.resolve()),
            "fingerprint": fingerprint_file(report_path),
        }

    def _complete_active_task_journal(self, history_saved: bool) -> None:
        task_id = self._active_task_id
        if task_id and history_saved:
            self._task_journal.clear(task_id)

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
            self._worker_thread = None
            self._active_task_kind = ""
        start_btn.config(state="normal", text=idle_text)
        if idle_bg is not None:
            start_btn.config(bg=idle_bg, activebackground=idle_bg)
        cancel_btn.config(state="disabled")
        self._pause_flag.clear()
        for pause_button_name in ("pause_org_btn", "pause_filter_btn"):
            pause_button = getattr(self, pause_button_name, None)
            if pause_button is not None:
                try:
                    pause_button.config(state="disabled", text="⏸ 暂停")
                except tk.TclError:
                    pass
        progress_bar["value"] = 0
        self.progress_label.config(text="")
        if self.status_var.get() == "⏳ 任务进行中...":
            self.status_var.set("就绪 - 请选择功能开始使用")
        self._active_task_id = None

    def _cancel_task(self) -> None:
        self._cancel_flag.set()
        self._pause_flag.clear()
        try:
            organize_cancel_active = str(self.cancel_org_btn.cget("state")) != "disabled"
        except (AttributeError, tk.TclError):
            organize_cancel_active = False
        if organize_cancel_active:
            stage = "preview" if self._active_task_kind == "readonly" else "execute"
            self._set_organize_workflow_stage(stage, "正在请求取消；当前读取步骤结束后任务会安全停止。")
        try:
            filter_cancel_active = str(self.cancel_flt_btn.cget("state")) != "disabled"
        except (AttributeError, tk.TclError):
            filter_cancel_active = False
        if filter_cancel_active and self._active_task_kind == "readonly":
            self._set_filter_workflow_stage("preview", "正在请求取消；当前 Excel 或目录读取步骤完成后会停止。")
        logger.warning("⏹ 用户请求取消任务")

    def _wait_if_paused(self) -> None:
        while self._pause_flag.is_set() and not self._cancel_flag.is_set():
            self._cancel_flag.wait(0.1)

    def _toggle_pause_task(self) -> None:
        if not self._task_is_running() or self._active_task_kind != "write":
            return
        if self._pause_flag.is_set():
            self._pause_flag.clear()
            next_text = "⏸ 暂停"
            self.status_var.set("⏳ 任务已继续执行...")
            if str(self.pause_org_btn.cget("state")) != "disabled":
                self._set_organize_workflow_stage("execute", "任务已继续，正在处理剩余文件。")
            if str(self.pause_filter_btn.cget("state")) != "disabled":
                self._set_filter_workflow_stage("execute", "任务已继续，正在处理剩余发票。")
            logger.info("批量任务已继续")
        else:
            self._pause_flag.set()
            next_text = "▶ 继续"
            self.status_var.set("⏸ 已请求暂停；当前文件完成后暂停。")
            if str(self.pause_org_btn.cget("state")) != "disabled":
                self._set_organize_workflow_stage("execute", "已请求暂停；当前文件安全处理完成后等待继续。")
            if str(self.pause_filter_btn.cget("state")) != "disabled":
                self._set_filter_workflow_stage("execute", "已请求暂停；当前文件安全处理完成后等待继续。")
            logger.info("批量任务已请求暂停")
        for pause_button in (self.pause_org_btn, self.pause_filter_btn):
            if str(pause_button.cget("state")) != "disabled":
                pause_button.config(text=next_text)

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

    def _scan_files(self, *, preserve_workflow: bool = False, silent: bool = False) -> bool:
        folder_str = self.organize_folder_path.get().strip()
        if not folder_str:
            self._set_organize_workflow_stage("input", "请先选择或拖入发票文件夹。")
            self._update_organize_result("等待选择目录", "尚未扫描任何文件。")
            if not silent:
                messagebox.showwarning("提示", "请先选择文件夹")
            return False
        folder = Path(folder_str)
        if not folder.exists():
            self._set_organize_workflow_stage("input", "所选文件夹不存在，请重新选择。")
            self._update_organize_result("目录不可用", folder_str)
            if not silent:
                messagebox.showerror("错误", "文件夹不存在")
            return False

        self.config["organize_folder"] = folder_str
        self._save_config()
        request = OrganizePreviewRequest(
            folder=folder,
            company_index=self._get_safe_company_name_index(),
            recursive=bool(self.organize_recursive.get()),
            filename_parser=self._get_filename_parser(),
        )
        signature = (str(folder.resolve()), request.recursive, request.company_index)
        if not preserve_workflow:
            self._set_organize_workflow_stage("preview", "正在扫描 PDF 并解析公司名称……")
            self._update_organize_result("正在扫描", "界面仍可响应；扫描完成后可勾选本次要处理的文件。")

        def work() -> OrganizePreviewResult:
            return OrganizeService.preview(
                request.folder,
                request.company_index,
                recursive=request.recursive,
                filename_parser=request.filename_parser,
                cancel_requested=self._cancel_flag.is_set,
            )

        def success(result: OrganizePreviewResult) -> None:
            current_signature = (
                str(Path(self.organize_folder_path.get().strip()).resolve()),
                bool(self.organize_recursive.get()),
                self._get_safe_company_name_index(),
            )
            if current_signature != signature:
                self.status_var.set("⚠ 扫描期间输入已变化，本次旧结果未应用，请重新扫描。")
                if not preserve_workflow:
                    self._set_organize_workflow_stage("input", "目录或规则已变化，请重新扫描。")
                return
            self._apply_organize_preview(result, preserve_workflow=preserve_workflow)

        def error(exc: Exception) -> None:
            logger.error("整理目录扫描失败：%s", exc)
            if not preserve_workflow:
                self._set_organize_workflow_stage("input", f"目录扫描失败：{exc}")
                self._update_organize_result("扫描失败", str(exc))
            self.status_var.set("❌ PDF 扫描失败")
            if not silent:
                messagebox.showerror("错误", f"无法扫描所选文件夹：\n{exc}")

        def cancelled() -> None:
            self.status_var.set("⏹ PDF 扫描已取消，文件未发生变化")
            if not preserve_workflow:
                self._set_organize_workflow_stage("preview", "扫描已取消；原预览结果保持不变。")
                self._update_organize_result("扫描已取消", "没有移动、复制或删除任何文件。")

        return self._start_readonly_task(
            task_name="扫描 PDF",
            worker_name="invoice-organize-preview",
            work=work,
            on_success=success,
            on_error=error,
            on_cancel=cancelled,
            controls=[
                (self.org_scan_btn, "⏳ 扫描中..."),
                (self.start_btn, None),
            ],
            cancel_button=self.cancel_org_btn,
            progress_bar=self.organize_progress,
            busy_status="⏳ 正在后台扫描 PDF...",
            silent_busy=silent,
        )

    def _apply_organize_preview(
        self,
        result: OrganizePreviewResult,
        *,
        preserve_workflow: bool = False,
    ) -> None:
        self.organize_failed_files = []
        self.organize_failure_folder = ""
        if hasattr(self, "retry_org_btn"):
            self.retry_org_btn.config(state="disabled")
        self.file_check_vars.clear()
        self.preview_data.clear()
        for row in result.rows:
            self.preview_data[row.relative_path] = {
                "filename": row.relative_path,
                "company": row.company,
                "target": row.target,
                "valid": row.selectable,
                "already_organized": row.already_organized,
            }
            self.file_check_vars[row.relative_path] = tk.BooleanVar(value=row.selectable)
        rerun_requested_count = 0
        rerun_selected_count = 0
        if self._pending_organize_rerun_files:
            rerun_requested_count = len(self._pending_organize_rerun_files)
            for filename, selected in self.file_check_vars.items():
                should_select = (
                    filename in self._pending_organize_rerun_files
                    and bool(self.preview_data.get(filename, {}).get("valid"))
                )
                selected.set(should_select)
                rerun_selected_count += int(should_select)
            self._pending_organize_rerun_files.clear()
        self._render_organize_preview(update_workflow=False)
        if preserve_workflow:
            self.status_var.set(f"✅ 已刷新 {result.total_count} 个文件")
            return
        if not result.rows:
            logger.warning("📭 未找到PDF文件")
            self._set_organize_workflow_stage("results", "扫描完成，但当前范围内没有找到 PDF 文件。")
            self._update_organize_result("未找到 PDF", "可检查目录、子文件夹选项或文件扩展名后重新扫描。")
            self.status_var.set("✅ 扫描完成，未找到 PDF")
            return
        logger.info("🔍 扫描到 %s 个PDF文件", result.total_count)
        self._update_organize_result(
            f"扫描到 {result.total_count} 个 PDF",
            f"可处理 {result.selectable_count} · 已在目标目录 {result.organized_count} · 文件名无效 {result.invalid_count}",
        )
        if result.selectable_count:
            if rerun_requested_count:
                unavailable_count = rerun_requested_count - rerun_selected_count
                unavailable_text = f"；{unavailable_count} 个历史文件已不存在或不再可处理" if unavailable_count else ""
                self._set_organize_workflow_stage(
                    "confirm",
                    f"已按历史记录选择 {rerun_selected_count} 个文件{unavailable_text}，请重新确认后执行。",
                )
            else:
                self._set_organize_workflow_stage(
                    "confirm",
                    f"已自动选择 {result.selectable_count} 个可处理文件，请确认后执行整理。",
                )
        else:
            self._set_organize_workflow_stage("results", "扫描完成，但没有可移动的文件。")
        self.status_var.set(f"✅ 已扫描 {result.total_count} 个文件")

    def _render_organize_preview(self, update_workflow: bool = True) -> None:
        if not hasattr(self, "file_tree"):
            return
        self.file_tree.delete(*self.file_tree.get_children())
        if not self.preview_data:
            self._update_file_count(update_workflow=update_workflow)
            return
        for index, fname in enumerate(self.preview_data.keys()):
            data = self.preview_data[fname]
            if fname not in self.file_check_vars:
                self.file_check_vars[fname] = tk.BooleanVar(value=bool(data.get("valid")))
            last_status = str(data.get("last_status", ""))
            if last_status == "失败":
                status = "!"
                tag = "failure_even" if index % 2 == 0 else "failure"
                target = str(data.get("last_detail") or data.get("target", "-"))
            elif last_status in {"已移动", "已跳过"}:
                status = "✓" if last_status == "已移动" else "—"
                tag = "success_even" if index % 2 == 0 else "success"
                target = str(data.get("last_detail") or data.get("target", "-"))
            else:
                status = "—" if data.get("already_organized") else ("✓" if self.file_check_vars[fname].get() else "✗")
                tag = ("evenrow" if index % 2 == 0 else "oddrow") if data["valid"] else ("invalid_even" if index % 2 == 0 else "invalid")
                target = data["target"]
            self.file_tree.insert("", "end", values=(status, fname, data["company"], target), tags=(tag,))
        self._update_file_count(update_workflow=update_workflow)

    def _apply_organize_execution_result(self, result) -> None:
        failed_files: List[str] = []
        failure_details: List[str] = []
        for row in result.result_rows:
            data = self.preview_data.get(row.filename)
            if data is None:
                data = {
                    "filename": row.filename,
                    "company": row.company,
                    "target": row.target or "-",
                    "valid": bool(row.retryable),
                    "already_organized": False,
                }
                self.preview_data[row.filename] = data
            data["last_status"] = row.status
            data["last_detail"] = row.detail
            data["valid"] = bool(row.retryable)
            if row.filename not in self.file_check_vars:
                self.file_check_vars[row.filename] = tk.BooleanVar(value=False)
            self.file_check_vars[row.filename].set(False)
            if row.retryable:
                failed_files.append(row.filename)
                failure_details.append(f"{Path(row.filename).name}：{row.detail}")
        self.organize_failed_files = failed_files
        self.organize_failure_folder = str(self.organize_folder_path.get()).strip() if failed_files else ""
        self.retry_org_btn.config(state="normal" if failed_files else "disabled")
        self._render_organize_preview(update_workflow=False)
        if failure_details:
            preview = "；".join(failure_details[:2])
            if len(failure_details) > 2:
                preview += f"；另有 {len(failure_details) - 2} 项"
            self._update_organize_result(
                "整理部分失败" if result.success_count else "整理失败",
                f"失败 {len(failure_details)} 项：{preview}",
            )

    def _retry_failed_organize(self) -> None:
        if not self.organize_failed_files:
            messagebox.showinfo("提示", "当前没有可重试的整理失败项。")
            return
        current_folder = self.organize_folder_path.get().strip()
        if not current_folder or str(Path(current_folder).resolve()) != str(Path(self.organize_failure_folder).resolve()):
            messagebox.showwarning("输入已变化", "整理目录已变化，请重新扫描后再执行。")
            return
        for filename, selected in self.file_check_vars.items():
            selected.set(filename in self.organize_failed_files)
        self._render_organize_preview(update_workflow=False)
        self._set_organize_workflow_stage(
            "confirm",
            f"已选中 {len(self.organize_failed_files)} 个失败项；确认后仅重试这些文件。",
        )
        self._execute_organize()

    def _on_tree_click(self, event) -> None:
        if self.file_tree.identify("region", event.x, event.y) == "cell":
            col = self.file_tree.identify_column(event.x)
            item = self.file_tree.identify_row(event.y)
            if item and col == "#1":
                vals = list(self.file_tree.item(item, "values"))
                fn = vals[1]
                data = self.preview_data.get(fn)
                if fn in self.file_check_vars and data and data.get("valid"):
                    cur = self.file_check_vars[fn].get()
                    self.file_check_vars[fn].set(not cur)
                    vals[0] = "✓" if not cur else "✗"
                    self.file_tree.item(item, values=vals)
                    self._update_file_count()

    def _update_file_count(self, update_workflow: bool = True) -> None:
        t = len(self.file_check_vars)
        s = sum(1 for v in self.file_check_vars.values() if v.get())
        self.file_count_label.config(text=f"已选择: {s} / {t}")
        if update_workflow and t:
            selectable = sum(1 for data in self.preview_data.values() if data.get("valid"))
            self._set_organize_workflow_stage(
                "confirm",
                f"已选择 {s} / {selectable} 个可处理文件；执行前仍会再次确认。",
            )
            self._update_organize_result(
                "等待确认",
                f"本次选择 {s} 个文件，可处理总数 {selectable}；未勾选文件保持原位。",
            )

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
            self._set_organize_workflow_stage("confirm", "尚未选择文件，请至少勾选一个可处理项目。")
            messagebox.showwarning("提示", "请至少选择一个文件")
            return
        self._set_organize_workflow_stage("confirm", f"准备移动 {len(sel)} 个文件，正在等待最终确认。")
        if not messagebox.askyesno("确认", f"确定整理 {len(sel)} 个文件？\n文件将被移动到对应公司文件夹。"):
            self._set_organize_workflow_stage("confirm", "已取消本次确认，文件尚未发生变化。")
            return
        if not self._try_begin_task(self.start_btn, "⏳ 处理中...", self.cancel_org_btn, busy_bg=self.palette["secondary"]):
            self._set_organize_workflow_stage("execute", "已有任务正在运行，请等待完成或先取消当前任务。")
            messagebox.showwarning("提示", "任务进行中...")
            return
        self.retry_org_btn.config(state="disabled")
        self.pause_org_btn.config(state="normal", text="⏸ 暂停")
        self._cancel_flag.clear()
        try:
            folder = Path(self.organize_folder_path.get()).resolve()
            rerun_payload = {
                "type": "整理",
                "folder": str(folder),
                "recursive": bool(self.organize_recursive.get()),
                "selected_files": list(sel),
            }
            task_id = self._task_journal.begin(
                "整理",
                folder,
                {"selected_count": len(sel), "rerun": rerun_payload},
            )
            self._active_task_id = task_id
            preview_snapshot = {name: dict(value) for name, value in self.preview_data.items()}
            request = OrganizeExecutionRequest(
                folder=folder,
                files=tuple(sel),
                preview_data=preview_snapshot,
                recursive=bool(self.organize_recursive.get()),
            )
            self._set_organize_workflow_stage(
                "execute",
                f"正在安全移动 {len(sel)} 个文件；每个成功操作都会写入恢复记录。",
            )
            self._update_organize_result("整理进行中", f"计划处理 {len(sel)} 个文件，可使用取消按钮安全停止。")
            self._start_worker(
                self._do_organize,
                (request, task_id),
                name="invoice-organize-task",
            )
        except Exception as exc:
            if self._active_task_id:
                self._task_journal.clear(self._active_task_id)
            self._finish_task_ui(
                self.start_btn,
                "🚀 执行整理",
                self.cancel_org_btn,
                self.organize_progress,
                idle_bg=self.palette["success"],
            )
            self.retry_org_btn.config(state="normal" if self.organize_failed_files else "disabled")
            logger.exception("整理任务启动失败")
            self._set_organize_workflow_stage("results", f"整理任务未能启动：{exc}")
            self._update_organize_result("启动失败", str(exc))
            messagebox.showerror("错误", str(exc))

    def _do_organize(
        self,
        request: OrganizeExecutionRequest,
        task_id: str,
    ) -> None:
        try:
            def on_progress(current: int, total: int) -> None:
                self._update_progress(self.organize_progress, current, total if current == 0 else None)
                if current == 0:
                    self._update_progress_info(0, total)
                elif current % 5 == 0 or current == total:
                    self._update_progress_info(current, total)

            result = OrganizeService.run(
                folder=request.folder,
                files=list(request.files),
                preview_data=request.preview_data,
                progress_callback=on_progress,
                cancel_requested=self._cancel_flag.is_set,
                operation_callback=self._record_active_task_move,
                pause_waiter=self._wait_if_paused,
            )
            final_m = result.moves

            def finish():
                history_saved = True
                serialized_results = [asdict(row) for row in result.result_rows]
                if final_m:
                    self.current_session_history = final_m
                if final_m or serialized_results or result.cancelled:
                    history_saved = self._save_to_history(
                        final_m,
                        "整理",
                        {
                            "task_id": task_id,
                            "folder": str(request.folder),
                            "count": len(serialized_results),
                            "result_rows": serialized_results,
                            "failed_count": result.fail_count,
                            "cancelled": result.cancelled,
                            "rerun": {
                                "type": "整理",
                                "folder": str(request.folder),
                                "recursive": request.recursive,
                                "selected_files": list(request.files),
                            },
                        },
                    )
                if final_m:
                    self.undo_btn.config(state="normal")
                    self.undo_all_btn.config(state="normal")
                self._complete_active_task_journal(history_saved)
                self._apply_organize_execution_result(result)
                state_text = "已取消" if result.cancelled else "完成"
                self.status_var.set(
                    f"{state_text} | 成功 {result.success_count} | 跳过 {result.skip_count} | "
                    f"失败 {result.fail_count} | {result.elapsed:.1f}秒"
                )
                outcome = "整理已取消" if result.cancelled else "整理完成"
                self._set_organize_workflow_stage(
                    "results",
                    f"{outcome}：成功 {result.success_count}，跳过 {result.skip_count}，失败 {result.fail_count}。",
                )
                self._update_organize_result(
                    outcome,
                    f"成功 {result.success_count} · 跳过 {result.skip_count} · 失败 {result.fail_count} · {result.elapsed:.1f} 秒"
                    + (
                        "；失败原因已显示在文件预览中，可点击“重试失败”。"
                        if result.fail_count
                        else ("；成功记录可在下方撤销。" if final_m else "")
                    ),
                )
                self._finish_task_ui(
                    self.start_btn,
                    "🚀 执行整理",
                    self.cancel_org_btn,
                    self.organize_progress,
                    idle_bg=self.palette["success"],
                )
                if not self._closing_requested:
                    messagebox.showinfo(
                        "已取消" if result.cancelled else "完成",
                        f"整理{state_text}！\n✅ 成功：{result.success_count}"
                        f"\n⏭ 跳过：{result.skip_count}\n❌ 失败：{result.fail_count}",
                    )
            self._post_ui(finish)

        except Exception as e:
            logger.exception("整理异常")
            msg = str(e)
            def err():
                recovered = self._recover_failed_task_into_history()
                recovery_note = f"\n\n已将 {recovered} 个已完成操作恢复到历史记录，可安全回滚。" if recovered else ""
                self._set_organize_workflow_stage(
                    "results",
                    f"整理异常：{msg}" + (f"；已恢复 {recovered} 个操作记录。" if recovered else ""),
                )
                self._update_organize_result(
                    "整理失败",
                    str(msg) + (f"；{recovered} 个已完成操作已进入历史记录。" if recovered else ""),
                )
                if not self._closing_requested:
                    messagebox.showerror("错误", msg + recovery_note)
                self._finish_task_ui(
                    self.start_btn,
                    "🚀 执行整理",
                    self.cancel_org_btn,
                    self.organize_progress,
                    idle_bg=self.palette["success"],
                )
                self.retry_org_btn.config(state="normal" if self.organize_failed_files else "disabled")
            self._post_ui(err)

    # ─── 撤销 ───

    def _undo_last_move(self) -> None:
        stage_updater = getattr(self, "_set_organize_workflow_stage", None)
        result_updater = getattr(self, "_update_organize_result", None)
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
            self._scan_files(preserve_workflow=True, silent=True)
            if callable(stage_updater):
                stage_updater("results", f"已撤销：{last['filename']}")
            if callable(result_updater):
                result_updater(
                    "撤销完成",
                    f"文件已安全移回原位置；剩余可撤销操作 {len(self.current_session_history)} 个。",
                )
        else:
            logger.warning(err)
            if callable(stage_updater):
                stage_updater("results", f"撤销失败：{err}")
            if callable(result_updater):
                result_updater("撤销失败", err)
        if not self.current_session_history:
            self.undo_btn.config(state="disabled")
            self.undo_all_btn.config(state="disabled")

    def _undo_all_moves(self) -> None:
        stage_updater = getattr(self, "_set_organize_workflow_stage", None)
        result_updater = getattr(self, "_update_organize_result", None)
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
        self._scan_files(preserve_workflow=True, silent=True)
        if callable(stage_updater):
            stage_updater(
                "results",
                f"批量撤销完成：成功 {ok_n}，失败 {fail_n}。" + ("失败记录已保留。" if fail_n else ""),
            )
        if callable(result_updater):
            result_updater(
                "批量撤销完成",
                f"成功 {ok_n} · 失败 {fail_n}" + (f"；{fail_n} 个失败记录仍可重试。" if fail_n else ""),
            )
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

    def _refresh_excel_sheets(self, silent: bool = False) -> bool:
        return self._refresh_workbook_analysis(silent=silent)

    def _on_excel_sheet_change(self, event=None) -> None:
        self._sync_filter_context(self.excel_sheet_name.get())
        self._sync_analysis_selection_to_current_sheet()
        self._sync_output_folder_mode_ui()
        self._save_config()
        self._set_filter_workflow_stage(
            "rules",
            f"当前工作表：{self.excel_sheet_name.get() or '未选择'}。请确认列映射和筛选条件。",
        )

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
        self.workbook_analysis_compact_var.set("列映射、行筛选和样本预览默认收起，需要调整时展开。")
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
        self.workbook_analysis_compact_var.set("列映射、行筛选和样本预览默认收起，需要调整时展开。")

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
            self.workbook_analysis_compact_var.set("列映射、行筛选和样本预览默认收起，需要调整时展开。")
            self.workbook_sheet_overview_var.set("先选择 Excel 文件，再从左侧查看每个 sheet 的识别结果。")
            self.workbook_sheet_sample_var.set("样本预览会显示当前工作表前几行数据，便于确认列是否正确。")
            if hasattr(self, "analysis_invoice_combo"):
                self.analysis_invoice_combo["values"] = ()
            if hasattr(self, "analysis_company_combo"):
                self.analysis_company_combo["values"] = ()
            if hasattr(self, "row_filter_column_combo"):
                self.row_filter_column_combo["values"] = ()
            return

        invoice_values = list(
            dict.fromkeys(
                [candidate.column_name for candidate in profile.invoice_candidates]
                + profile.columns
            )
        )
        company_values = [""] + list(
            dict.fromkeys(
                [candidate.column_name for candidate in profile.company_candidates]
                + profile.columns
            )
        )
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
        self.workbook_analysis_compact_var.set(
            f"当前：{profile.sheet_name} | 发票列：{profile.selected_invoice_column or '-'} | "
            f"公司列：{profile.selected_company_column or '-'} | 条件：{active_filter_text}"
        )
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
        current_sheet = self.excel_sheet_name.get()
        if not current_sheet:
            return
        self._select_workbook_tree_item(current_sheet)
        self._populate_workbook_sheet_detail(current_sheet)

    def _refresh_workbook_analysis(self, silent: bool = False) -> bool:
        if not PANDAS_SUPPORT:
            return False
        excel = self.excel_path.get().strip()
        if not excel:
            if hasattr(self, "excel_sheet_combo"):
                self.excel_sheet_combo["values"] = ()
            self.excel_sheet_name.set("")
            self._clear_workbook_analysis("打开 Excel 后，会自动分析每个工作表的发票列和公司列候选。")
            self._sync_output_folder_mode_ui()
            self._set_filter_workflow_stage("input", "先选择 Excel 文件，再补齐 PDF 来源和导出位置。")
            return False

        excel_path = Path(excel)
        if not excel_path.exists():
            if hasattr(self, "excel_sheet_combo"):
                self.excel_sheet_combo["values"] = ()
            self._clear_workbook_analysis("Excel 文件不存在，无法分析工作簿。")
            self._sync_output_folder_mode_ui()
            self._set_filter_workflow_stage("input", "Excel 文件不存在，请重新选择有效文件。")
            return False

        request = WorkbookAnalysisRequest(
            excel_path=excel_path,
            extra_invoice_aliases=tuple(self._get_invoice_aliases()),
            extra_company_aliases=tuple(self._get_company_aliases()),
            selected_sheet_name=self.excel_sheet_name.get(),
            selected_invoice_column_name=self.selected_invoice_column_name.get().strip(),
            selected_company_column_name=self.selected_company_column_name.get().strip(),
        )
        signature = str(excel_path.resolve())
        self.workbook_analysis_summary_var.set("正在后台分析工作簿，请稍候……")
        self._set_filter_workflow_stage("rules", "正在读取工作表、列候选和样本数据；界面可继续响应。")

        def work() -> WorkbookAnalysisResult:
            return WorkbookAnalyzerService.analyze(
                request.excel_path,
                extra_invoice_aliases=list(request.extra_invoice_aliases),
                extra_company_aliases=list(request.extra_company_aliases),
                cancel_requested=self._cancel_flag.is_set,
            )

        def success(result: WorkbookAnalysisResult) -> None:
            current_excel = self.excel_path.get().strip()
            if not current_excel or str(Path(current_excel).resolve()) != signature:
                self.status_var.set("⚠ 分析期间 Excel 路径已变化，本次旧结果未应用。")
                self._set_filter_workflow_stage("input", "Excel 路径已变化，请重新分析工作簿。")
                return
            self._apply_workbook_analysis_result(request, result)

        def error(exc: Exception) -> None:
            if hasattr(self, "excel_sheet_combo"):
                self.excel_sheet_combo["values"] = ()
            self._clear_workbook_analysis(f"工作簿分析失败：{exc}")
            self._sync_output_folder_mode_ui()
            self._set_filter_workflow_stage("input", f"无法读取 Excel：{exc}")
            self.status_var.set("❌ 工作簿分析失败")
            if not silent:
                messagebox.showerror("错误", str(exc))

        def cancelled() -> None:
            self.workbook_analysis_summary_var.set("工作簿分析已取消；上一次有效结果仍保留。")
            self._set_filter_workflow_stage("rules", "工作簿分析已取消，可调整输入后重试。")
            self.status_var.set("⏹ 工作簿分析已取消")

        controls = [
            (widget, busy_text)
            for widget, busy_text in (
                (getattr(self, "sheet_refresh_btn", None), "读取中..."),
                (getattr(self, "workbook_analysis_btn", None), "分析中..."),
                (getattr(self, "filter_preview_btn", None), None),
                (getattr(self, "filter_run_btn", None), None),
            )
            if widget is not None
        ]
        return self._start_readonly_task(
            task_name="分析工作簿",
            worker_name="invoice-workbook-analysis",
            work=work,
            on_success=success,
            on_error=error,
            on_cancel=cancelled,
            controls=controls,
            cancel_button=self.cancel_flt_btn,
            progress_bar=self.filter_progress,
            busy_status="⏳ 正在后台分析工作簿...",
            silent_busy=silent,
        )

    def _apply_workbook_analysis_result(
        self,
        request: WorkbookAnalysisRequest,
        result: WorkbookAnalysisResult,
    ) -> None:
        sheets = [profile.sheet_name for profile in result.sheet_profiles]
        if hasattr(self, "excel_sheet_combo"):
            self.excel_sheet_combo["values"] = sheets
        self.workbook_analysis_result = result
        self._render_workbook_analysis(result)

        current_sheet = self.excel_sheet_name.get()
        if current_sheet not in self.workbook_profiles:
            current_sheet = (
                request.selected_sheet_name
                if request.selected_sheet_name in self.workbook_profiles
                else result.recommended_sheet_name
            )
            if not current_sheet and sheets:
                current_sheet = sheets[0]
            self.excel_sheet_name.set(current_sheet)
        self._sync_filter_context(current_sheet)
        profile = self.workbook_profiles.get(current_sheet)
        if profile is not None:
            if request.selected_invoice_column_name in profile.columns:
                profile.selected_invoice_column = request.selected_invoice_column_name
            if request.selected_company_column_name in profile.columns:
                profile.selected_company_column = request.selected_company_column_name
        self._sync_analysis_selection_to_current_sheet()
        self._sync_output_folder_mode_ui()
        self._save_config()
        self._set_filter_workflow_stage(
            "rules",
            f"工作簿分析完成，推荐工作表：{result.recommended_sheet_name or current_sheet or '无'}。请确认规则后预览。",
        )
        self.status_var.set(f"✅ 已分析 {len(sheets)} 个工作表")

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
        sheet_name = self.excel_sheet_name.get()
        profile = self.workbook_profiles.get(sheet_name)
        if profile is None or self.workbook_analysis_result is None:
            return
        profile.selected_invoice_column = self.selected_invoice_column_name.get().strip()
        self._render_workbook_analysis(self.workbook_analysis_result)
        self._select_workbook_tree_item(sheet_name)
        self._save_config()
        self._set_filter_workflow_stage("rules", "发票号列已更新，请预览确认匹配结果。")

    def _on_analysis_company_column_change(self, event=None) -> None:
        sheet_name = self.excel_sheet_name.get()
        profile = self.workbook_profiles.get(sheet_name)
        if profile is None or self.workbook_analysis_result is None:
            return
        profile.selected_company_column = self.selected_company_column_name.get().strip()
        self._render_workbook_analysis(self.workbook_analysis_result)
        self._select_workbook_tree_item(sheet_name)
        self._save_config()
        self._set_filter_workflow_stage("rules", "公司列已更新，请预览确认目录和报告信息。")

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
        current_sheet = self.excel_sheet_name.get()
        if current_sheet:
            self._populate_workbook_sheet_detail(current_sheet)
        self._save_config()
        self._set_filter_workflow_stage(
            "rules",
            f"当前筛选条件：{self._describe_active_row_filters()}。建议先预览再执行。",
        )

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
        workflow_updater = getattr(self, "_set_filter_workflow_stage", None)

        def mark_input_error(message: str) -> None:
            if callable(workflow_updater):
                workflow_updater("input", message)

        excel = self.excel_path.get()
        pdf = self.pdf_folder.get()
        if not excel or not Path(excel).exists():
            mark_input_error("Excel 文件无效，请重新选择。")
            messagebox.showerror("错误", "请选择有效的Excel文件")
            return None
        if not pdf or not Path(pdf).exists():
            mark_input_error("PDF 来源文件夹无效，请重新选择。")
            messagebox.showerror("错误", "请选择有效的PDF文件夹")
            return None
        out_path_raw = self._get_effective_output_folder_path()
        if out_path_raw is None:
            mark_input_error("请先设置有效的导出文件夹。")
            messagebox.showerror("错误", "请选择导出文件夹")
            return None
        pdf_path = Path(pdf).resolve()
        out_path = out_path_raw.resolve()
        if pdf_path == out_path:
            mark_input_error("导出位置与 PDF 来源相同，请更换导出位置。")
            messagebox.showerror("错误", "导出文件夹不能与PDF源文件夹相同！")
            return None
        if self.filter_recursive.get() and is_relative_to(out_path, pdf_path):
            mark_input_error("递归模式下，导出位置不能放在 PDF 来源目录内部。")
            messagebox.showerror("错误", "递归筛选时，导出文件夹不能位于PDF源文件夹内部！")
            return None
        return Path(excel), Path(pdf), out_path

    def _filter_preview_context_signature(self, output_dir: Path) -> Tuple[str, ...]:
        excel_text = self.excel_path.get().strip()
        pdf_text = self.pdf_folder.get().strip()
        return (
            str(Path(excel_text).resolve()) if excel_text else "",
            str(Path(pdf_text).resolve()) if pdf_text else "",
            str(output_dir.resolve()),
            str(bool(self.filter_recursive.get())),
            self.excel_sheet_name.get(),
            self.selected_invoice_column_name.get().strip(),
            self.selected_company_column_name.get().strip(),
            self.row_filter_column_name.get().strip(),
            self.row_filter_mode.get().strip() or "不过滤",
            self.row_filter_values.get().strip(),
            self.company_exclude_keywords.get().strip(),
            str(self._get_safe_invoice_number_index()),
            self.rule_preset_id.get().strip(),
            "\x1f".join(self._get_invoice_aliases()),
        )

    def _preview_filter(self) -> bool:
        paths = self._validate_filter_paths()
        if not paths:
            return False
        excel_p, pdf_p, out_p = paths
        selected_company_column = self.selected_company_column_name.get().strip()
        active_filter_desc = self._describe_active_row_filters()
        recursive = bool(self.filter_recursive.get())
        signature = self._filter_preview_context_signature(out_p)
        request = FilterPreviewRequest(
            excel_path=excel_p,
            pdf_folder=pdf_p,
            output_dir=out_p,
            invoice_index=self._get_safe_invoice_number_index(),
            recursive=recursive,
            sheet_name=self.excel_sheet_name.get(),
            invoice_column_name=self.selected_invoice_column_name.get().strip() or None,
            company_column_name=selected_company_column or None,
            filter_column_name=self.row_filter_column_name.get().strip() or None,
            filter_mode=self.row_filter_mode.get().strip() or "不过滤",
            filter_values=self.row_filter_values.get().strip() or None,
            company_exclude_keywords=self.company_exclude_keywords.get().strip() or None,
            extra_aliases=tuple(self._get_invoice_aliases()),
            exclude_dirs=(out_p,) if recursive else (),
            filename_parser=self._get_filename_parser(),
            column_resolver=self._get_column_resolver(),
            active_filter_desc=active_filter_desc,
            context_signature=signature,
        )
        self._set_filter_workflow_stage("preview", "正在检查 Excel 发票号、PDF 匹配和重复冲突……")
        self._update_filter_summary(
            "正在预览",
            "正在后台读取 Excel 并扫描 PDF；当前界面不会冻结，可随时取消。",
            [
                ("原始行数", "-"),
                ("筛选后发票", "-"),
                ("可匹配", "-"),
                ("已过滤行", "-"),
                ("未匹配", "-"),
                ("PDF扫描", "-"),
            ],
        )

        def work() -> FilterPreviewResult:
            return FilterService.preview(
                request.excel_path,
                request.pdf_folder,
                request.invoice_index,
                recursive=request.recursive,
                sheet_name=request.sheet_name,
                invoice_column_name=request.invoice_column_name,
                company_column_name=request.company_column_name,
                filter_column_name=request.filter_column_name,
                filter_mode=request.filter_mode,
                filter_values=request.filter_values,
                company_exclude_keywords=request.company_exclude_keywords,
                extra_aliases=list(request.extra_aliases),
                exclude_dirs=list(request.exclude_dirs) or None,
                filename_parser=request.filename_parser,
                column_resolver=request.column_resolver,
                cancel_requested=self._cancel_flag.is_set,
            )

        def success(preview: FilterPreviewResult) -> None:
            if self._filter_preview_context_signature(out_p) != request.context_signature:
                self.status_var.set("⚠ 预览期间规则或路径已变化，本次旧结果未应用。")
                self._set_filter_workflow_stage("rules", "输入或规则已变化，请重新预览。")
                return
            self._apply_filter_preview_result(request, preview)

        def error(exc: Exception) -> None:
            self._set_filter_workflow_stage("preview", f"预览失败：{exc}")
            self._update_filter_summary("预览失败", str(exc), [])
            self.status_var.set("❌ 筛选预览失败")
            messagebox.showerror("错误", str(exc))

        def cancelled() -> None:
            self._set_filter_workflow_stage("preview", "预览已取消；没有复制、移动或删除任何文件。")
            self._update_filter_summary("预览已取消", "上一次有效结果仍保留，可调整规则后重试。", [])
            self.status_var.set("⏹ 筛选预览已取消")

        return self._start_readonly_task(
            task_name="预览匹配",
            worker_name="invoice-filter-preview",
            work=work,
            on_success=success,
            on_error=error,
            on_cancel=cancelled,
            controls=[
                (self.filter_preview_btn, "⏳ 预览中..."),
                (self.filter_run_btn, None),
                (self.workbook_analysis_btn, None),
                (self.sheet_refresh_btn, None),
            ],
            cancel_button=self.cancel_flt_btn,
            progress_bar=self.filter_progress,
            busy_status="⏳ 正在后台生成匹配预览...",
        )

    def _apply_filter_preview_result(
        self,
        request: FilterPreviewRequest,
        preview: FilterPreviewResult,
    ) -> None:
        self._last_filter_preview_signature = request.context_signature
        self._last_filter_preview_result = preview
        columns_preview = "、".join(preview.columns[:6])
        if len(preview.columns) > 6:
            columns_preview += f" ... 共{len(preview.columns)}列"
        self._update_filter_summary(
            "预览完成",
            f"工作表：{preview.sheet_name} | 发票列：{preview.excel_column_name} | "
            f"公司列：{request.company_column_name or '未指定'} | 条件：{request.active_filter_desc} | 可用列：{columns_preview}",
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
        elif not preview.invoice_numbers:
            self.filter_detail_var.set("当前工作表和筛选条件下没有可处理的发票号；请检查列映射、空值或筛选规则。")
        elif preview.not_found:
            self.filter_detail_var.set(f"共有 {len(preview.not_found)} 个发票号未匹配，可直接复制未匹配发票号继续跟进。")
        else:
            self.filter_detail_var.set("预览完成：当前发票号均已找到对应 PDF，可直接开始筛选导出。")
        issue_count = len(preview.not_found) + len(preview.conflicts)
        self._set_filter_workflow_stage(
            "results",
            f"预览完成：匹配 {len(preview.matched)}，需处理 {issue_count}。确认后可开始导出。",
        )
        self.status_var.set(
            f"✅ 预览完成：匹配 {len(preview.matched)} | 未匹配 {len(preview.not_found)} | 冲突 {len(preview.conflicts)}"
        )
        logger.info(
            "预览完成：工作表 %s | 筛选后 %s | 过滤掉 %s | 匹配 %s | 未匹配 %s | PDF扫描 %s",
            preview.sheet_name,
            len(preview.invoice_numbers),
            preview.filtered_out_count,
            len(preview.matched),
            len(preview.not_found),
            preview.pdf_stats.scanned,
        )

    def _retry_failed_filter(self) -> None:
        retry_count = sum(1 for row in self.filter_result_rows if row.status == "复制失败")
        if not retry_count:
            messagebox.showinfo("提示", "当前没有可重试的复制失败项。")
            return
        if not messagebox.askyesno(
            "重试复制失败项",
            f"将按当前路径和规则重新执行筛选，以重试 {retry_count} 个复制失败项。\n\n"
            "已经成功导出且内容一致的文件会安全跳过，不会覆盖现有文件。是否继续？",
        ):
            return
        self._set_filter_workflow_stage(
            "execute",
            f"准备重试 {retry_count} 个复制失败项；已成功文件会按内容校验后跳过。",
        )
        self._run_filter(skip_confirmation=True)

    def _run_filter(self, *, skip_confirmation: bool = False) -> None:
        paths = self._validate_filter_paths()
        if not paths:
            return
        excel_p, pdf_p, out_p = paths
        current_signature = self._filter_preview_context_signature(out_p)
        preview = self._last_filter_preview_result
        if self._last_filter_preview_signature != current_signature or preview is None:
            self._set_filter_workflow_stage(
                "preview",
                "输入、规则或输出位置尚未完成当前版本的预览，请先预览再执行。",
            )
            messagebox.showwarning(
                "请先预览",
                "为避免按错误的工作表、列或目录导出，请先点击“预览匹配”。\n"
                "预览完成且输入未变化后才能开始筛选。",
            )
            return
        if not preview.invoice_numbers:
            self._set_filter_workflow_stage("preview", "当前预览没有可处理的发票号，未启动文件任务。")
            messagebox.showwarning("没有可处理数据", "当前预览中没有发票号，请调整工作表、列映射或筛选条件。")
            return
        if not skip_confirmation and not messagebox.askyesno(
            "确认筛选并导出",
            f"筛选后发票：{len(preview.invoice_numbers)} 个\n"
            f"可匹配 PDF：{len(preview.matched)} 个\n"
            f"未匹配：{len(preview.not_found)} 个\n"
            f"重复冲突：{len(preview.conflicts)} 个\n\n"
            f"导出位置：{out_p}\n\n"
            "程序不会覆盖内容不同的同名文件。确认开始吗？",
        ):
            self._set_filter_workflow_stage("preview", "已取消最终确认，尚未复制或生成任何文件。")
            return
        if not self._try_begin_task(self.filter_run_btn, "⏳ 处理中...", self.cancel_flt_btn, busy_bg=self.palette["secondary"]):
            messagebox.showwarning("提示", "任务进行中...")
            return
        self.filter_retry_btn.config(state="disabled")
        self.pause_filter_btn.config(state="normal", text="⏸ 暂停")
        self._cancel_flag.clear()
        try:
            recursive = bool(self.filter_recursive.get())
            request = FilterExecutionRequest(
                excel_path=excel_p,
                pdf_folder=pdf_p,
                output_dir=out_p,
                invoice_index=self._get_safe_invoice_number_index(),
                recursive=recursive,
                sheet_name=self.excel_sheet_name.get(),
                invoice_column_name=self.selected_invoice_column_name.get().strip() or None,
                company_column_name=self.selected_company_column_name.get().strip() or None,
                filter_column_name=self.row_filter_column_name.get().strip() or None,
                filter_mode=self.row_filter_mode.get().strip() or "不过滤",
                filter_values=self.row_filter_values.get().strip() or None,
                company_exclude_keywords=self.company_exclude_keywords.get().strip() or None,
                extra_aliases=tuple(self._get_invoice_aliases()),
                exclude_dirs=tuple(self._get_filter_exclude_dirs()) if recursive else (),
                filename_parser=self._get_filename_parser(),
                column_resolver=self._get_column_resolver(),
                report_exporter=self._get_report_exporter(),
                active_filter_desc=self._describe_active_row_filters(),
                rule_preset_id=self.rule_preset_id.get().strip(),
                custom_invoice_aliases=self.invoice_column_aliases.get().strip(),
            )
            rerun_payload = self._filter_rerun_payload(request)
            task_id = self._task_journal.begin(
                "筛选",
                pdf_p,
                {
                    "excel_path": str(excel_p),
                    "output_root": str(out_p),
                    "rerun": rerun_payload,
                },
            )
            self._active_task_id = task_id
            self._set_filter_workflow_stage("execute", "筛选任务正在后台执行，可随时取消；已完成的文件会进入可恢复记录。")
            self._start_worker(
                self._do_filter,
                (request, task_id),
                name="invoice-filter-task",
            )
        except Exception as exc:
            if self._active_task_id:
                self._task_journal.clear(self._active_task_id)
            self._finish_task_ui(
                self.filter_run_btn,
                "🚀 开始筛选并导出",
                self.cancel_flt_btn,
                self.filter_progress,
                idle_bg=self.palette["primary"],
            )
            self.filter_retry_btn.config(
                state="normal" if any(row.status == "复制失败" for row in self.filter_result_rows) else "disabled"
            )
            logger.exception("筛选任务启动失败")
            self._set_filter_workflow_stage("results", f"筛选任务未能启动：{exc}")
            messagebox.showerror("错误", str(exc))

    @staticmethod
    def _filter_rerun_payload(request: FilterExecutionRequest) -> Dict[str, Any]:
        return {
            "type": "筛选",
            "excel_path": str(request.excel_path),
            "pdf_folder": str(request.pdf_folder),
            "output_dir": str(request.output_dir),
            "invoice_index": request.invoice_index,
            "recursive": request.recursive,
            "sheet_name": request.sheet_name,
            "invoice_column_name": request.invoice_column_name or "",
            "company_column_name": request.company_column_name or "",
            "filter_column_name": request.filter_column_name or "",
            "filter_mode": request.filter_mode,
            "filter_values": request.filter_values or "",
            "company_exclude_keywords": request.company_exclude_keywords or "",
            "rule_preset_id": request.rule_preset_id,
            "custom_invoice_aliases": request.custom_invoice_aliases,
        }

    def _do_filter(self, request: FilterExecutionRequest, task_id: str) -> None:
        self._post_ui(
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

        excel_p = request.excel_path
        pdf_p = request.pdf_folder
        out_p = request.output_dir
        selected_company_column = request.company_column_name or ""
        active_filter_desc = request.active_filter_desc

        try:
            report_entries: List[Dict[str, Any]] = []

            def on_report(report_path: Path) -> None:
                try:
                    report_entry = self._build_report_entry(report_path, out_p)
                except (OSError, ValueError) as exc:
                    logger.error("报告安全指纹生成失败：%s", exc)
                    report_entry = {
                        "path": str(report_path),
                        "filename": report_path.name,
                        "output_root": str(out_p.resolve()),
                    }
                if not self._task_journal.record_report(task_id, report_entry):
                    raise OSError("报告恢复日志写入失败")
                report_entries.append(report_entry)

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
                invoice_index=request.invoice_index,
                recursive=request.recursive,
                sheet_name=request.sheet_name,
                invoice_column_name=request.invoice_column_name,
                company_column_name=selected_company_column or None,
                filter_column_name=request.filter_column_name,
                filter_mode=request.filter_mode,
                filter_values=request.filter_values,
                company_exclude_keywords=request.company_exclude_keywords,
                extra_aliases=list(request.extra_aliases),
                exclude_dirs=list(request.exclude_dirs) or None,
                filename_parser=request.filename_parser,
                column_resolver=request.column_resolver,
                report_exporter=request.report_exporter,
                progress_callback=on_progress,
                cancel_requested=self._cancel_flag.is_set,
                operation_callback=self._record_active_task_move,
                report_callback=on_report,
                pause_waiter=self._wait_if_paused,
            )
            report_files = [str(result.report_path)] if result.report_path else []

            def finish():
                history_saved = True
                serialized_results = [asdict(row) for row in result.result_rows]
                if result.moves or report_files or serialized_results or result.cancelled:
                    history_saved = self._save_to_history(
                        result.moves,
                        "筛选",
                        {
                            "report_files": report_files,
                            "report_entries": report_entries,
                            "task_id": task_id,
                            "folder": str(pdf_p),
                            "count": len(serialized_results),
                            "result_rows": serialized_results,
                            "failed_count": result.copy_fail_count + result.target_conflict_count,
                            "cancelled": result.cancelled,
                            "rerun": self._filter_rerun_payload(request),
                        },
                    )
                self._complete_active_task_journal(history_saved)

                columns_preview = "、".join(result.columns[:6])
                if len(result.columns) > 6:
                    columns_preview += f" ... 共{len(result.columns)}列"
                title = "筛选已取消" if result.cancelled else "筛选完成"
                self._update_filter_summary(
                    title,
                    f"工作表：{result.sheet_name} | 发票列：{result.excel_column_name} | 公司列：{selected_company_column or '未指定'} | 条件：{active_filter_desc} | 可用列：{columns_preview}",
                    [
                        ("原始行数", str(result.source_row_count)),
                        (
                            "筛选后发票",
                            str(
                                result.found_count
                                + len(result.not_found)
                                + result.skip_count
                                + result.copy_fail_count
                                + result.target_conflict_count
                                + len(result.conflicts)
                            ),
                        ),
                        ("已导出", str(result.found_count)),
                        ("已过滤行", str(result.filtered_out_count)),
                        ("未匹配", str(len(result.not_found))),
                        ("PDF扫描", str(result.pdf_stats.scanned)),
                    ],
                )
                self._set_filter_results(result.result_rows, missing_invoices=result.not_found)
                if result.conflicts:
                    self.filter_detail_var.set(f"本次发现 {len(result.conflicts)} 个重复冲突，已在结果表格中标记为“重复冲突”。")
                elif result.target_conflict_count:
                    self.filter_detail_var.set(
                        f"本次有 {result.target_conflict_count} 个导出同名冲突；原文件已保留，请在结果表中处理。"
                    )
                elif result.filtered_out_count:
                    self.filter_detail_var.set(f"筛选完成：已按条件过滤掉 {result.filtered_out_count} 行，导出 {result.found_count} 个文件。")
                elif result.not_found:
                    self.filter_detail_var.set(f"本次有 {len(result.not_found)} 个发票号未匹配，可用“复制未匹配发票号”继续处理。")
                elif result.found_count > 0:
                    self.filter_detail_var.set("筛选完成：可双击结果表中的文件直接打开，或点击“打开导出文件夹”查看全部导出结果。")
                else:
                    self.filter_detail_var.set("本次没有匹配到可导出的文件，请检查 Excel 工作表、列名或 PDF 命名规则。")

                self.status_var.set(
                    f"✅ 成功: {result.found_count} | 跳过: {result.skip_count} | "
                    f"同名冲突: {result.target_conflict_count} | 复制失败: {result.copy_fail_count} | "
                    f"未找到: {len(result.not_found)} | {result.elapsed:.1f}秒"
                )
                outcome = "已取消" if result.cancelled else "已完成"
                unresolved = (
                    len(result.not_found)
                    + len(result.conflicts)
                    + result.target_conflict_count
                    + result.copy_fail_count
                )
                self._set_filter_workflow_stage(
                    "results",
                    f"任务{outcome}：导出 {result.found_count}，需处理 {unresolved}，耗时 {result.elapsed:.1f} 秒。",
                )

                report_msg = ""
                if report_files:
                    report_msg = f"\n\n📊 筛选报告已保存到导出文件夹"

                if not self._closing_requested:
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
            self._post_ui(finish)

        except (FileNotFoundError, PermissionError, ValueError) as e:
            msg = str(e)
            def err():
                logger.error(msg)
                recovered = self._recover_failed_task_into_history()
                recovery_note = f"\n\n已将 {recovered} 个已完成操作恢复到历史记录，可安全回滚。" if recovered else ""
                self._set_filter_workflow_stage(
                    "results",
                    f"筛选失败：{msg}" + (f"；已恢复 {recovered} 个操作记录。" if recovered else ""),
                )
                if not self._closing_requested:
                    messagebox.showerror("错误", msg + recovery_note)
                self._finish_task_ui(
                    self.filter_run_btn,
                    "🚀 开始筛选并导出",
                    self.cancel_flt_btn,
                    self.filter_progress,
                    idle_bg=self.palette["primary"],
                )
            self._post_ui(err)
        except Exception as e:
            logger.exception("筛选异常")
            msg = str(e)
            def err2():
                recovered = self._recover_failed_task_into_history()
                recovery_note = f"\n\n已将 {recovered} 个已完成操作恢复到历史记录，可安全回滚。" if recovered else ""
                self._set_filter_workflow_stage(
                    "results",
                    f"筛选异常：{msg}" + (f"；已恢复 {recovered} 个操作记录。" if recovered else ""),
                )
                if not self._closing_requested:
                    messagebox.showerror("错误", msg + recovery_note)
                self._finish_task_ui(
                    self.filter_run_btn,
                    "🚀 开始筛选并导出",
                    self.cancel_flt_btn,
                    self.filter_progress,
                    idle_bg=self.palette["primary"],
                )
            self._post_ui(err2)

    # ─────────────── 历史记录 ───────────────

    def _save_to_history(
        self,
        moves: List[Dict],
        op: str = "整理",
        extra: Optional[Dict[str, Any]] = None,
    ) -> bool:
        folder = self.organize_folder_path.get() if op == "整理" else self.pdf_folder.get()
        record = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "folder": folder, "count": len(moves), "type": op, "moves": moves,
        }
        if extra:
            record.update(extra)
        self.all_history.insert(0, record)
        self.all_history = self.all_history[:100]
        saved = self._save_history()
        self._refresh_history_tree()
        return saved

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

    @staticmethod
    def _history_record_has_safe_rollback(record: Dict[str, Any]) -> bool:
        operation_type = record.get("type", "整理")
        raw_moves = record.get("moves", [])
        raw_report_files = record.get("report_files", [])
        raw_report_entries = record.get("report_entries", [])
        if not isinstance(raw_moves, list) or not isinstance(raw_report_files, list) or not isinstance(raw_report_entries, list):
            return False
        moves = [move for move in raw_moves if isinstance(move, dict)]
        report_files = [str(path) for path in raw_report_files]
        report_entries = [entry for entry in raw_report_entries if isinstance(entry, dict)]
        if not moves and not report_files:
            return False
        root_key = "output_root" if operation_type == "筛选" else "operation_root"
        if any(not move.get(root_key) or not has_valid_fingerprint(move.get("fingerprint")) for move in moves):
            return False
        if report_files:
            entries_by_path = {str(entry.get("path", "")): entry for entry in report_entries}
            for report_path in report_files:
                entry = entries_by_path.get(report_path)
                if not entry or not entry.get("output_root") or not has_valid_fingerprint(entry.get("fingerprint")):
                    return False
        return True

    def _clear_history_detail(self, filtered_empty: bool = False) -> None:
        self.history_detail_title.set("没有匹配记录" if filtered_empty else "未选择任务")
        self.history_detail_meta.set(
            "调整筛选条件后重试。" if filtered_empty else "从左侧选择一条记录查看处理摘要。"
        )
        self.history_detail_folder.set("")
        self.history_detail_safety.set("没有可评估的回滚记录")
        self.history_safety_label.config(fg=self.palette["muted"])
        self.history_preview_listbox.delete(0, tk.END)
        for button in (
            self.history_rollback_btn,
            self.history_rerun_btn,
            self.history_view_btn,
            self.history_open_btn,
        ):
            button.config(state="disabled")
        self.history_action_status_text.set(
            "没有符合当前条件的历史记录。" if filtered_empty else "选择任务后可查看详情、打开目录或执行安全回滚。"
        )

    def _on_history_select(self, event=None) -> None:
        record = self._get_selected_history_record()
        if record is None:
            self._clear_history_detail(filtered_empty=bool(self.all_history and not self.filtered_history_indices))
            return

        operation_type = str(record.get("type", "整理"))
        recovered_record = bool(record.get("recovered"))
        raw_moves = record.get("moves", [])
        raw_report_files = record.get("report_files", [])
        raw_result_rows = record.get("result_rows", [])
        history_structure_valid = (
            isinstance(raw_moves, list)
            and isinstance(raw_report_files, list)
            and isinstance(raw_result_rows, list)
        )
        moves = [move for move in raw_moves if isinstance(move, dict)] if isinstance(raw_moves, list) else []
        report_files = [str(path) for path in raw_report_files] if isinstance(raw_report_files, list) else []
        result_rows = [row for row in raw_result_rows if isinstance(row, dict)] if isinstance(raw_result_rows, list) else []
        move_count = len(moves)
        report_count = len(report_files)
        result_count = len(result_rows)
        self.history_detail_title.set(str(record.get("time", "未知时间")))
        report_text = f" · 报告 {report_count}" if report_count else ""
        result_text = f" · 结果 {result_count}" if result_count else ""
        self.history_detail_meta.set(
            f"{operation_type}任务 · 文件操作 {move_count}{report_text}{result_text}"
            + (" · 异常中断恢复" if recovered_record else "")
        )
        self.history_detail_folder.set(str(record.get("folder", "未记录")))

        rollback_items = move_count + report_count
        safe_rollback = history_structure_valid and self._history_record_has_safe_rollback(record)
        if not history_structure_valid:
            safety_text = "⚠ 历史结构损坏：已禁用自动回滚，可查看原始记录定位问题"
            safety_color = self.palette["status_error"]
        elif not rollback_items:
            safety_text = "此记录没有可回滚的文件操作"
            safety_color = self.palette["muted"]
        elif safe_rollback:
            safety_text = (
                "✓ 异常恢复记录可校验：执行前会验证路径与文件指纹"
                if recovered_record
                else "✓ 可校验回滚：执行前会验证路径与文件指纹"
            )
            safety_color = self.palette["status_success"]
        else:
            safety_text = "⚠ 旧记录或校验信息不完整：无法验证的文件会保留"
            safety_color = self.palette["status_conflict"]
        self.history_detail_safety.set(safety_text)
        self.history_safety_label.config(fg=safety_color)

        self.history_preview_listbox.delete(0, tk.END)
        for move in moves:
            self.history_preview_listbox.insert(tk.END, str(move.get("filename", "未知文件")))
        for report_file in report_files:
            self.history_preview_listbox.insert(tk.END, f"[报告] {Path(report_file).name}")
        for row in result_rows[:100]:
            status = str(row.get("status", "结果"))
            subject = str(
                row.get("filename")
                or row.get("pdf_name")
                or row.get("invoice_number")
                or "未知项目"
            )
            detail = str(row.get("detail", ""))
            text = f"[{status}] {subject}"
            if detail:
                text += f" — {detail}"
            self.history_preview_listbox.insert(tk.END, text)
        if self.history_preview_listbox.size() == 0:
            self.history_preview_listbox.insert(tk.END, "（没有文件明细）")

        self.history_rollback_btn.config(state="normal" if rollback_items and history_structure_valid else "disabled")
        self.history_rerun_btn.config(state="normal" if self._history_record_can_rerun(record) else "disabled")
        self.history_view_btn.config(state="normal")
        self.history_open_btn.config(state="normal")
        self.history_action_status_text.set(
            f"已选择 {operation_type}任务：{rollback_items} 个可评估项目。"
            + (
                "可执行指纹校验回滚。"
                if safe_rollback
                else ("历史结构损坏，自动回滚已禁用。" if not history_structure_valid else "回滚时将保护无法验证的项目。")
            )
        )

    def _refresh_history_tree(self, preferred_index: Optional[int] = None) -> None:
        if preferred_index is None:
            try:
                preferred_index = self._get_selected_history_index()
            except (AttributeError, IndexError, tk.TclError):
                preferred_index = None
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
            fd = str(r.get("folder", "未记录"))
            if len(fd) > 50:
                fd = "..." + fd[-47:]
            tag = "evenrow" if visible_index % 2 == 0 else "oddrow"
            raw_reports = r.get("report_files", [])
            report_count = len(raw_reports) if isinstance(raw_reports, list) else 0
            raw_moves = r.get("moves", [])
            fallback_count = len(raw_moves) if isinstance(raw_moves, list) else 0
            count_text = f"{r.get('count', fallback_count)}个"
            if report_count:
                count_text += f" + {report_count}报告"
            try:
                failed_count = max(0, int(r.get("failed_count", 0) or 0))
            except (TypeError, ValueError):
                failed_count = 0
            if failed_count:
                count_text += f" · {failed_count}失败"
            if r.get("cancelled"):
                count_text += " · 已取消"
            if r.get("recovered"):
                count_text += " · 已恢复"
            self.history_tree.insert(
                "",
                "end",
                values=(str(r.get("time", "未知时间")), fd, count_text, r.get("type", "整理")),
                tags=(tag,),
            )

        self.history_clear_btn.config(state="normal" if self.all_history else "disabled")
        if not self.filtered_history_indices:
            self._clear_history_detail(filtered_empty=bool(self.all_history))
            return
        selected_position = 0
        if preferred_index in self.filtered_history_indices:
            selected_position = self.filtered_history_indices.index(preferred_index)
        children = self.history_tree.get_children()
        if children:
            selected_item = children[selected_position]
            self.history_tree.selection_set(selected_item)
            self.history_tree.focus(selected_item)
            self.history_tree.see(selected_item)
        self._on_history_select()

    @staticmethod
    def _history_record_can_rerun(record: Dict[str, Any]) -> bool:
        return history_record_can_rerun(record)

    def _load_history_for_rerun(self) -> None:
        if not self._require_idle("载入历史任务"):
            return
        record = self._get_selected_history_record()
        if record is None or not self._history_record_can_rerun(record):
            messagebox.showinfo("提示", "该记录来自旧版本或缺少再次执行参数。")
            return
        rerun = dict(record["rerun"])
        task_type = str(rerun.get("type", record.get("type", "")))
        if not messagebox.askyesno(
            "载入历史任务",
            f"将用历史记录中的路径和规则替换当前{task_type}页输入。\n"
            "此操作只载入并预览，不会立即移动或复制文件。是否继续？",
        ):
            return

        if task_type == "整理":
            folder = Path(str(rerun.get("folder", "")))
            if not folder.exists():
                messagebox.showwarning("目录不存在", f"历史整理目录已不存在：\n{folder}")
                return
            selected_files = {
                str(item)
                for item in rerun.get("selected_files", [])
                if isinstance(item, str) and item
            }
            self.organize_folder_path.set(str(folder))
            self.organize_recursive.set(bool(rerun.get("recursive", False)))
            self._pending_organize_rerun_files = selected_files
            self._select_workspace_page("organize")
            self._scan_files()
            self.status_var.set("⏳ 已载入历史整理任务，正在重新扫描；确认后才会移动文件。")
        else:
            self.excel_path.set(str(rerun.get("excel_path", "")))
            self.pdf_folder.set(str(rerun.get("pdf_folder", "")))
            self.auto_output_by_sheet.set(False)
            self.manual_output_folder.set(str(rerun.get("output_dir", "")))
            self.output_folder.set(str(rerun.get("output_dir", "")))
            self._clear_workbook_analysis("已载入历史筛选任务；可重新分析工作簿或直接预览。")
            self.excel_sheet_name.set(str(rerun.get("sheet_name", "")))
            self.selected_invoice_column_name.set(str(rerun.get("invoice_column_name", "")))
            self.selected_company_column_name.set(str(rerun.get("company_column_name", "")))
            self.row_filter_column_name.set(str(rerun.get("filter_column_name", "")))
            self.row_filter_mode.set(str(rerun.get("filter_mode", "不过滤")) or "不过滤")
            self.row_filter_values.set(str(rerun.get("filter_values", "")))
            self.company_exclude_keywords.set(str(rerun.get("company_exclude_keywords", "")))
            try:
                invoice_index = int(rerun.get("invoice_index", self._get_safe_invoice_number_index()))
                if 0 <= invoice_index <= 10:
                    self.invoice_number_index.set(invoice_index)
            except (TypeError, ValueError, tk.TclError):
                pass
            preset_id = str(rerun.get("rule_preset_id", ""))
            if preset_id in self._preset_by_id:
                self.rule_preset_id.set(preset_id)
                self._sync_rule_preset_ui()
            self.invoice_column_aliases.set(str(rerun.get("custom_invoice_aliases", "")))
            self.filter_recursive.set(bool(rerun.get("recursive", False)))
            self._active_filter_context = self._current_filter_context()
            self._sync_output_folder_mode_ui()
            self._save_config()
            self._select_workspace_page("filter")
            self._set_filter_workflow_stage(
                "rules",
                "历史筛选任务已载入；请检查文件是否仍存在，并先预览再执行。",
            )
            self.status_var.set("✅ 已载入历史筛选任务；尚未复制任何文件。")

        self.history_action_status_text.set(
            f"已载入{task_type}任务；为安全起见，不会自动执行文件操作。"
        )

    def _view_history_detail(self) -> None:
        rec = self._get_selected_history_record()
        if rec is None:
            messagebox.showinfo("提示", "请先选择记录")
            return
        win = tk.Toplevel(self.root)
        win.title("历史详情")
        win.geometry("750x500")
        raw_moves = rec.get("moves", [])
        moves = [move for move in raw_moves if isinstance(move, dict)] if isinstance(raw_moves, list) else []
        raw_reports = rec.get("report_files", [])
        report_files = [str(path) for path in raw_reports] if isinstance(raw_reports, list) else []
        raw_results = rec.get("result_rows", [])
        result_rows = [row for row in raw_results if isinstance(row, dict)] if isinstance(raw_results, list) else []
        report_count = len(report_files)
        count_desc = f"{rec.get('count', len(moves))}"
        if report_count:
            count_desc += f"（另含 {report_count} 个报告）"
        for t in [
            f"时间：{rec.get('time', '未知时间')}",
            f"类型：{rec.get('type','整理')}",
            f"文件夹：{rec.get('folder', '未记录')}",
            f"数量：{count_desc}",
        ]:
            tk.Label(win, text=t, font=("微软雅黑", 10), wraplength=700).pack(anchor="w", padx=10)
        lf = tk.LabelFrame(win, text="文件列表 (双击列表项可直接打开文件)", padx=10, pady=10)
        lf.pack(fill="both", expand=True, padx=10, pady=10)
        scr = tk.Scrollbar(lf)
        scr.pack(side="right", fill="y")
        lb = tk.Listbox(lf, font=("Consolas", 9), yscrollcommand=scr.set)
        lb.pack(fill="both", expand=True)
        scr.config(command=lb.yview)

        item_paths = []
        for m in moves:
            lb.insert(tk.END, str(m.get("filename", "未知文件")))
            item_paths.append(Path(str(m.get("target", ""))))
        for report_file in report_files:
            lb.insert(tk.END, f"[报告] {Path(report_file).name}")
            item_paths.append(Path(report_file))
        for row in result_rows:
            status = str(row.get("status", "结果"))
            subject = str(
                row.get("filename")
                or row.get("pdf_name")
                or row.get("invoice_number")
                or "未知项目"
            )
            detail = str(row.get("detail", ""))
            lb.insert(tk.END, f"[{status}] {subject}" + (f" — {detail}" if detail else ""))
        if not isinstance(raw_moves, list) or not isinstance(raw_reports, list) or not isinstance(raw_results, list):
            lb.insert(tk.END, "[警告] 历史结构损坏，自动回滚已禁用")

        def on_item_double_click(event):
            selection = lb.curselection()
            if not selection:
                return
            idx = selection[0]
            if idx < len(item_paths):
                file_path = item_paths[idx]
                if file_path.exists():
                    self._open_path_in_shell(file_path)
                else:
                    messagebox.showwarning("提示", f"该文件已不存在或已被移动：\n{file_path}")

        lb.bind("<Double-1>", on_item_double_click)

    def _open_history_folder(self) -> None:
        rec = self._get_selected_history_record()
        if rec is None:
            messagebox.showinfo("提示", "请先选择记录")
            return
        folder_value = str(rec.get("folder", "")).strip()
        if not folder_value:
            messagebox.showwarning("提示", "该历史记录没有可用的文件夹路径")
            return
        folder = Path(folder_value)
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
        raw_moves = rec.get("moves", [])
        raw_report_files = rec.get("report_files", [])
        raw_report_entries = rec.get("report_entries", [])
        if not isinstance(raw_moves, list) or not isinstance(raw_report_files, list) or not isinstance(raw_report_entries, list):
            self.history_action_status_text.set("历史结构损坏，已阻止自动回滚。")
            messagebox.showerror("错误", "该历史记录结构损坏，已阻止自动回滚。请先导出日志并人工检查。")
            return
        moves = [move for move in raw_moves if isinstance(move, dict)]

        if op == "筛选":
            report_files = [str(path) for path in raw_report_files]
            report_desc = f" 和 {len(report_files)} 个报告" if report_files else ""
            if not messagebox.askyesno(
                "确认",
                f"回滚筛选？将删除仍与历史指纹一致的 {len(moves)} 个导出文件{report_desc}。\n"
                "内容已变化或旧历史缺少校验信息的文件会被保留。",
            ):
                return
            ok_n = fail_n = 0
            failed = []
            for m in moves:
                ok, err = InvoiceOrganizer.delete_recorded_file(m)
                if ok:
                    ok_n += 1
                else:
                    logger.error(err)
                    fail_n += 1
                    failed.append(m)
            report_entries = {
                str(entry.get("path", "")): entry
                for entry in raw_report_entries
                if isinstance(entry, dict)
            }
            failed_reports: List[str] = []
            failed_report_entries: List[Dict[str, Any]] = []
            for report_path in report_files:
                entry = report_entries.get(str(report_path), {"path": str(report_path)})
                ok, err = InvoiceOrganizer.delete_recorded_file(entry, path_key="path")
                if ok:
                    ok_n += 1
                else:
                    logger.error(err)
                    fail_n += 1
                    failed_reports.append(report_path)
                    failed_report_entries.append(entry)
        else:
            if not messagebox.askyesno("确认", f"回滚整理？将移回 {len(moves)} 个文件"):
                return
            ok_n = fail_n = 0
            failed = []
            for m in reversed(moves):
                ok, err = InvoiceOrganizer.rollback_single_move(m)
                if ok:
                    ok_n += 1
                else:
                    logger.error(err)
                    fail_n += 1
                    failed.append(m)
            failed_reports = []
            failed_report_entries = []

        if fail_n == 0:
            self.all_history.pop(idx)
        else:
            rec["moves"] = failed
            rec["count"] = len(failed)
            if failed_reports:
                rec["report_files"] = failed_reports
                rec["report_entries"] = failed_report_entries
            elif "report_files" in rec:
                rec.pop("report_files")
                rec.pop("report_entries", None)

        self._save_history()
        preferred_index = idx if fail_n and idx < len(self.all_history) else None
        self._refresh_history_tree(preferred_index=preferred_index)
        logger.info(f"↩️ 回滚：成功 {ok_n} 失败 {fail_n}")
        self.history_action_status_text.set(
            f"回滚完成：成功 {ok_n}，失败 {fail_n}。"
            + ("失败项目已保留在历史记录中。" if fail_n else "历史记录已同步更新。")
        )
        if op == "整理":
            self._scan_files(preserve_workflow=True, silent=True)
        messagebox.showinfo("完成", f"成功 {ok_n} | 失败 {fail_n}" + (f"\n{fail_n}个失败记录已保留" if fail_n else ""))

    def _clear_all_history(self) -> None:
        if not self.all_history:
            messagebox.showinfo("提示", "已经是空的")
            return
        if messagebox.askyesno("确认", "清空所有历史？不影响已处理文件。"):
            self.all_history.clear()
            self._save_history()
            self._refresh_history_tree()
            self.history_action_status_text.set("历史记录已清空；已处理文件未发生变化。")
            logger.info("🗑️ 历史已清空")
