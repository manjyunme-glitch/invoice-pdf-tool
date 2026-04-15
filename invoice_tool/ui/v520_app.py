from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Any, Dict

from ..runtime import DND_SUPPORT, MODERN_UI, OPENPYXL_SUPPORT, PANDAS_SUPPORT, ttkb
from .app import APP_TITLE, InvoiceToolApp as BaseInvoiceToolApp


class InvoiceToolApp(BaseInvoiceToolApp):
    """v5.2.0 visual refresh with a calmer, more polished workspace shell."""

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

        style.configure("TNotebook", background=palette["root_bg"], borderwidth=0, tabmargins=[0, 8, 0, 0])
        style.configure(
            "TNotebook.Tab",
            font=("微软雅黑", 9, "bold"),
            padding=[16, 9],
            background=palette["tab_idle_bg"],
            foreground=palette["tab_idle_fg"],
            borderwidth=0,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", palette["tab_active_bg"])],
            foreground=[("selected", palette["tab_active_fg"])],
        )

        style.configure(
            "Treeview",
            rowheight=26,
            font=("微软雅黑", 9),
            background=palette["tree_odd"],
            fieldbackground=palette["tree_odd"],
            foreground=palette["text"],
            relief="flat",
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            font=("微软雅黑", 9, "bold"),
            background=palette["tree_heading_bg"],
            foreground=palette["tree_heading_fg"],
            relief="flat",
            padding=[8, 7],
        )
        style.map(
            "Treeview.Heading",
            background=[("active", palette["surface_soft"])],
            foreground=[("active", palette["tree_heading_fg"])],
        )
        style.map("Treeview", background=[("selected", palette["tree_selected"])], foreground=[("selected", "#FFFFFF")])

        style.configure(
            "TCombobox",
            fieldbackground=palette["entry_bg"],
            foreground=palette["entry_fg"],
            background=palette["entry_bg"],
            arrowcolor=palette["text"],
            bordercolor=palette["border"],
            lightcolor=palette["border"],
            darkcolor=palette["border"],
            padding=4,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", palette["entry_bg"])],
            foreground=[("readonly", palette["entry_fg"])],
        )

        for orient in ("Vertical", "Horizontal"):
            style.configure(
                f"{orient}.TScrollbar",
                background=palette["button_bg"],
                troughcolor=palette["surface_soft"],
                bordercolor=palette["border"],
                arrowcolor=palette["muted"],
                relief="flat",
            )
            style.map(
                f"{orient}.TScrollbar",
                background=[("active", palette["button_hover"])],
            )

    def _apply_theme_to_widget_tree(self, widget: tk.Widget) -> None:
        palette = self.palette
        parent_bg = palette["root_bg"]
        try:
            parent_bg = str(widget.master.cget("bg"))
        except Exception:
            pass

        cls = widget.winfo_class()
        if cls in {"Frame", "Labelframe", "LabelFrame"}:
            target_bg = parent_bg
            if cls in {"Labelframe", "LabelFrame"}:
                target_bg = palette["surface_raised"]
            if self._should_apply_default_bg(widget):
                widget.configure(bg=target_bg)
            if cls in {"Labelframe", "LabelFrame"}:
                try:
                    widget.configure(
                        fg=palette["text"],
                        bd=1,
                        relief="flat",
                        highlightbackground=palette["border"],
                        highlightcolor=palette["border"],
                        highlightthickness=1,
                    )
                except tk.TclError:
                    pass
        elif cls == "Label":
            if self._should_apply_default_bg(widget):
                widget.configure(bg=parent_bg)
            if self._should_apply_default_bg(widget, "fg"):
                widget.configure(fg=palette["text"])
        elif cls in {"Checkbutton", "Radiobutton"}:
            widget.configure(
                bg=parent_bg,
                fg=palette["text"],
                activebackground=parent_bg,
                activeforeground=palette["text"],
                selectcolor=palette["surface_raised"],
            )
        elif cls in {"Entry", "Spinbox"}:
            widget.configure(
                bg=palette["entry_bg"],
                fg=palette["entry_fg"],
                insertbackground=palette["entry_fg"],
                highlightbackground=palette["border"],
                highlightcolor=palette["primary"],
                relief="flat",
                bd=1,
            )
        elif cls == "Text" and widget is not getattr(self, "log_text", None):
            widget.configure(
                bg=palette["entry_bg"],
                fg=palette["entry_fg"],
                insertbackground=palette["entry_fg"],
                highlightbackground=palette["border"],
                highlightcolor=palette["primary"],
                relief="flat",
                bd=1,
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
                bd=1,
            )
        elif cls == "Canvas":
            widget.configure(bg=parent_bg, highlightthickness=0, bd=0)
        elif cls == "Button" and self._should_apply_default_bg(widget):
            widget.configure(
                bg=palette["button_bg"],
                fg=palette["button_fg"],
                activebackground=palette["button_hover"],
                activeforeground=palette["button_fg"],
                relief="flat",
                bd=0,
                highlightthickness=0,
                cursor="hand2",
            )

        for child in widget.winfo_children():
            self._apply_theme_to_widget_tree(child)

    def _style_action_button(self, button: tk.Button, role: str) -> None:
        super()._style_action_button(button, role)
        emphasis_roles = {"primary", "success", "warning", "danger"}
        button.configure(
            cursor="hand2",
            font=("微软雅黑", 9, "bold" if role in emphasis_roles else "normal"),
            padx=12,
            pady=7,
        )

    def _create_chip_label(self, parent: tk.Widget, text: str) -> None:
        tk.Label(
            parent,
            text=text,
            font=("微软雅黑", 8, "bold"),
            bg=self.palette["hero_chip_bg"],
            fg=self.palette["hero_chip_fg"],
            padx=10,
            pady=4,
        ).pack(side="left", padx=(0, 8), pady=(0, 4))

    def _create_hero_stat(self, parent: tk.Widget, title: str, value: str, detail: str) -> None:
        card = tk.Frame(
            parent,
            bg=self.palette["surface_raised"],
            highlightbackground=self.palette["hero_card_border"],
            highlightcolor=self.palette["hero_card_border"],
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        card.pack(fill="x", pady=(0, 8))
        tk.Label(
            card,
            text=title,
            font=("微软雅黑", 8, "bold"),
            bg=self.palette["surface_raised"],
            fg=self.palette["muted"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            card,
            text=value,
            font=("微软雅黑", 13, "bold"),
            bg=self.palette["surface_raised"],
            fg=self.palette["text"],
            anchor="w",
        ).pack(anchor="w", pady=(4, 3))
        tk.Label(
            card,
            text=detail,
            font=("微软雅黑", 8),
            bg=self.palette["surface_raised"],
            fg=self.palette["muted"],
            justify="left",
            wraplength=180,
            anchor="w",
        ).pack(anchor="w")

    def _polish_action_text(self) -> None:
        replacements: Dict[str, tuple[str, str]] = {
            "start_btn": ("开始整理", "success"),
            "undo_btn": ("撤销上次", "warning"),
            "undo_all_btn": ("撤销全部", "danger"),
            "cancel_org_btn": ("取消", "secondary"),
            "help_btn": ("查看使用说明", "neutral"),
            "filter_preview_btn": ("预览匹配", "secondary"),
            "filter_run_btn": ("开始筛选并导出", "primary"),
            "cancel_flt_btn": ("取消", "secondary"),
            "copy_missing_btn": ("复制未匹配发票号", "neutral"),
            "open_result_btn": ("打开选中结果", "neutral"),
            "output_folder_browse_btn": ("浏览", "neutral"),
            "theme_toggle_btn": ("切换到夜间" if self.ui_theme.get() == "day" else "切换到白天", "neutral"),
        }
        for attr_name, (text, role) in replacements.items():
            button = getattr(self, attr_name, None)
            if isinstance(button, tk.Button):
                button.configure(text=text)
                self._style_action_button(button, role)

    def _build_ui(self) -> None:
        palette = self.palette
        self.root.configure(bg=palette["root_bg"])
        self._configure_ttk_styles()

        shell = tk.Frame(self.root, bg=palette["root_bg"])
        shell.pack(fill="both", expand=True)

        hero_wrap = tk.Frame(shell, bg=palette["root_bg"], padx=14, pady=12)
        hero_wrap.pack(fill="x")
        hero = tk.Frame(
            hero_wrap,
            bg=palette["hero_card_bg"],
            highlightbackground=palette["hero_card_border"],
            highlightcolor=palette["hero_card_border"],
            highlightthickness=1,
            padx=20,
            pady=18,
        )
        hero.pack(fill="x")

        hero_left = tk.Frame(hero, bg=palette["hero_card_bg"])
        hero_left.pack(side="left", fill="x", expand=True)
        tk.Label(
            hero_left,
            text=APP_TITLE,
            font=("微软雅黑", 18, "bold"),
            bg=palette["hero_card_bg"],
            fg=palette["title_fg"],
        ).pack(anchor="w")
        tk.Label(
            hero_left,
            text="以更安静、更稳定的桌面工作台处理复杂 Excel、多 Sheet 发票筛选与 PDF 归档。",
            font=("微软雅黑", 9),
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
        ).pack(anchor="w", pady=(5, 12))

        chip_row = tk.Frame(hero_left, bg=palette["hero_card_bg"])
        chip_row.pack(anchor="w")
        for chip in ("多 Sheet 分析", "行级条件筛选", "自动归档导出", "白天 / 黑夜主题"):
            self._create_chip_label(chip_row, chip)

        capability_row = tk.Frame(hero_left, bg=palette["hero_card_bg"])
        capability_row.pack(anchor="w", pady=(10, 0))
        capabilities = [
            "拖拽" if DND_SUPPORT else "无拖拽",
            "Excel" if PANDAS_SUPPORT else "无 Excel 支持",
            "现代主题" if MODERN_UI else "原生主题",
            "报告导出" if OPENPYXL_SUPPORT else "无报告导出",
        ]
        tk.Label(
            capability_row,
            text=" · ".join(capabilities),
            font=("微软雅黑", 8),
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
        ).pack(anchor="w")

        hero_right = tk.Frame(hero, bg=palette["hero_card_bg"])
        hero_right.pack(side="right", fill="y", padx=(20, 0))

        title_actions = tk.Frame(hero_right, bg=palette["hero_card_bg"])
        title_actions.pack(anchor="e", fill="x")
        self.theme_badge = tk.Label(
            title_actions,
            text=f"{self._theme_label()} UI",
            font=("微软雅黑", 8, "bold"),
            bg=palette["title_badge_bg"],
            fg=palette["title_badge_fg"],
            padx=10,
            pady=5,
        )
        self.theme_badge.pack(side="left", padx=(0, 8))
        self.theme_toggle_btn = tk.Button(
            title_actions,
            text="切换到夜间" if self.ui_theme.get() == "day" else "切换到白天",
            command=self._toggle_ui_theme,
        )
        self.theme_toggle_btn.pack(side="left")
        self._style_action_button(self.theme_toggle_btn, "neutral")

        stat_row = tk.Frame(hero_right, bg=palette["hero_card_bg"])
        stat_row.pack(fill="x", pady=(14, 0))
        self._create_hero_stat(stat_row, "工作流", "多 Sheet", "先分析工作簿，再确认列映射和样本预览。")
        self._create_hero_stat(stat_row, "筛选器", "按条件", "支持抵扣状态、关键字和排除公司等行级规则。")
        self._create_hero_stat(stat_row, "交付", "单文件 EXE", "适合直接发给同事使用，减少环境依赖。")

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

        tab_padding = 8
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
        status_wrap.pack(fill="x", side="bottom")
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

    def _create_filter_metric_card(self, parent: tk.Widget, metric_key: str, bg: str, fg: str) -> None:
        card = tk.Frame(
            parent,
            bg=self.palette["surface_raised"],
            highlightbackground=self.palette["border"],
            highlightcolor=self.palette["border"],
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        card.pack(side="left", fill="x", expand=True, padx=4)

        badge = tk.Label(
            card,
            textvariable=self.filter_metric_labels[metric_key],
            font=("微软雅黑", 8, "bold"),
            bg=bg,
            fg=fg,
            padx=8,
            pady=3,
        )
        badge.pack(anchor="w")
        tk.Label(
            card,
            textvariable=self.filter_metric_values[metric_key],
            font=("微软雅黑", 16, "bold"),
            bg=self.palette["surface_raised"],
            fg=self.palette["text"],
            anchor="w",
        ).pack(anchor="w", pady=(8, 2))
        tk.Label(
            card,
            text="筛选进度摘要",
            font=("微软雅黑", 8),
            bg=self.palette["surface_raised"],
            fg=self.palette["muted"],
            anchor="w",
        ).pack(anchor="w")
