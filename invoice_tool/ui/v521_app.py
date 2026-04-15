from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from ..runtime import DND_SUPPORT, MODERN_UI, OPENPYXL_SUPPORT, PANDAS_SUPPORT
from .app import APP_TITLE
from .v520_app import InvoiceToolApp as V520InvoiceToolApp


class InvoiceToolApp(V520InvoiceToolApp):
    """v5.2.1 compact workspace refresh."""

    def _create_compact_chip(self, parent: tk.Widget, text: str, *, accent: bool = False) -> None:
        bg = self.palette["hero_chip_bg"] if not accent else self.palette["primary"]
        fg = self.palette["hero_chip_fg"] if not accent else "#FFFFFF"
        tk.Label(
            parent,
            text=text,
            font=("微软雅黑", 8, "bold"),
            bg=bg,
            fg=fg,
            padx=8,
            pady=3,
        ).pack(side="left", padx=(0, 6))

    def _build_ui(self) -> None:
        palette = self.palette
        self.root.configure(bg=palette["root_bg"])
        self._configure_ttk_styles()

        shell = tk.Frame(self.root, bg=palette["root_bg"])
        shell.pack(fill="both", expand=True)

        top_wrap = tk.Frame(shell, bg=palette["root_bg"], padx=14, pady=10)
        top_wrap.pack(fill="x", pady=(0, 8))

        header = tk.Frame(
            top_wrap,
            bg=palette["hero_card_bg"],
            highlightbackground=palette["hero_card_border"],
            highlightcolor=palette["hero_card_border"],
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        header.pack(fill="x")

        top_row = tk.Frame(header, bg=palette["hero_card_bg"])
        top_row.pack(fill="x")

        left_col = tk.Frame(top_row, bg=palette["hero_card_bg"])
        left_col.pack(side="left", fill="x", expand=True)
        tk.Label(
            left_col,
            text=APP_TITLE,
            font=("微软雅黑", 15, "bold"),
            bg=palette["hero_card_bg"],
            fg=palette["title_fg"],
        ).pack(anchor="w")
        tk.Label(
            left_col,
            text="把复杂 Excel、多 Sheet 筛选和 PDF 归档收进更紧凑的桌面工作台。",
            font=("微软雅黑", 9),
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
        ).pack(anchor="w", pady=(4, 0))

        right_col = tk.Frame(top_row, bg=palette["hero_card_bg"])
        right_col.pack(side="right", anchor="ne", padx=(12, 0))
        action_row = tk.Frame(right_col, bg=palette["hero_card_bg"])
        action_row.pack(anchor="e")
        self.theme_badge = tk.Label(
            action_row,
            text=f"{self._theme_label()} UI",
            font=("微软雅黑", 8, "bold"),
            bg=palette["title_badge_bg"],
            fg=palette["title_badge_fg"],
            padx=8,
            pady=4,
        )
        self.theme_badge.pack(side="left", padx=(0, 8))
        self.theme_toggle_btn = tk.Button(
            action_row,
            text="切换到夜间" if self.ui_theme.get() == "day" else "切换到白天",
            command=self._toggle_ui_theme,
        )
        self.theme_toggle_btn.pack(side="left")
        self._style_action_button(self.theme_toggle_btn, "neutral")

        compact_meta = tk.Frame(right_col, bg=palette["hero_card_bg"])
        compact_meta.pack(anchor="e", pady=(8, 0))
        for text, accent in (
            ("多 Sheet", True),
            ("条件筛选", False),
            ("自动归档", False),
            ("单文件 EXE", False),
        ):
            self._create_compact_chip(compact_meta, text, accent=accent)

        bottom_row = tk.Frame(header, bg=palette["hero_card_bg"])
        bottom_row.pack(fill="x", pady=(10, 0))

        quick_flow = tk.Frame(bottom_row, bg=palette["hero_card_bg"])
        quick_flow.pack(side="left", fill="x", expand=True)
        tk.Label(
            quick_flow,
            text="推荐流程",
            font=("微软雅黑", 8, "bold"),
            bg=palette["hero_card_bg"],
            fg=palette["hero_accent"],
        ).pack(side="left", padx=(0, 10))
        tk.Label(
            quick_flow,
            text="1. 分析工作簿   2. 确认列映射   3. 预览匹配   4. 导出结果",
            font=("微软雅黑", 8),
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
        ).pack(side="left")

        capabilities = [
            "拖拽" if DND_SUPPORT else "无拖拽",
            "Excel" if PANDAS_SUPPORT else "无 Excel 支持",
            "现代主题" if MODERN_UI else "原生主题",
            "报告导出" if OPENPYXL_SUPPORT else "无报告导出",
        ]
        tk.Label(
            bottom_row,
            text=" · ".join(capabilities),
            font=("微软雅黑", 8),
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
        ).pack(side="right", padx=(12, 0))

        content_wrap = tk.Frame(shell, bg=palette["root_bg"], padx=14, pady=0)
        content_wrap.pack(fill="both", expand=True)
        content_shell = tk.Frame(
            content_wrap,
            bg=palette["surface_soft"],
            highlightbackground=palette["border"],
            highlightcolor=palette["border"],
            highlightthickness=1,
            padx=8,
            pady=8,
        )
        content_shell.pack(fill="both", expand=True)

        notebook_host = tk.Frame(content_shell, bg=palette["surface_soft"])
        notebook_host.pack(fill="both", expand=True)
        self.notebook = ttk.Notebook(notebook_host)
        self.notebook.pack(fill="both", expand=True)

        tab_padding = 6
        self.organize_frame = ttk.Frame(self.notebook, padding=tab_padding)
        self.notebook.add(self.organize_frame, text="整理")

        self.filter_frame = ttk.Frame(self.notebook, padding=tab_padding)
        self.notebook.add(self.filter_frame, text="筛选")

        self.history_frame = ttk.Frame(self.notebook, padding=tab_padding)
        self.notebook.add(self.history_frame, text="历史")

        self.settings_frame = ttk.Frame(self.notebook, padding=tab_padding)
        self.notebook.add(self.settings_frame, text="设置")

        self._build_organize_tab()
        self._build_filter_tab()
        self._build_history_tab()
        self._build_settings_tab()
        self._polish_action_text()

        self._build_log_drawer()

        status_wrap = tk.Frame(shell, bg=palette["root_bg"], padx=14, pady=6)
        status_wrap.pack(fill="x", side="bottom", pady=(0, 10))
        status_frame = tk.Frame(
            status_wrap,
            bg=palette["status_bg"],
            highlightbackground=palette["border"],
            highlightcolor=palette["border"],
            highlightthickness=1,
        )
        status_frame.pack(fill="x")

        self.status_var = tk.StringVar(value="就绪，可开始整理、筛选或查看历史记录。")
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("微软雅黑", 9),
            anchor="w",
            padx=12,
            pady=6,
            bg=palette["status_bg"],
            fg=palette["status_fg"],
        ).pack(side="left", fill="x", expand=True)

        self.progress_label = tk.Label(
            status_frame,
            text="",
            font=("微软雅黑", 9),
            fg=palette["muted"],
            bg=palette["status_bg"],
            padx=12,
            pady=6,
        )
        self.progress_label.pack(side="right")

        self._apply_theme_to_widget_tree(self.root)
