from __future__ import annotations

from dataclasses import dataclass
import tkinter as tk
from tkinter import ttk
from typing import Dict, Tuple

from ..runtime import DND_SUPPORT, MODERN_UI, OPENPYXL_SUPPORT, PANDAS_SUPPORT
from ..version import APP_VERSION
from .v521_app import InvoiceToolApp as V521InvoiceToolApp


@dataclass(frozen=True)
class WorkspacePage:
    key: str
    number: str
    title: str
    navigation_label: str
    subtitle: str
    workflow: str


WORKSPACE_PAGES: Tuple[WorkspacePage, ...] = (
    WorkspacePage(
        key="filter",
        number="01",
        title="Excel 筛选与 PDF 匹配",
        navigation_label="Excel 筛选",
        subtitle="导入工作簿、确认规则、预览匹配，再安全导出 PDF。",
        workflow="输入  →  规则  →  预览  →  执行  →  结果",
    ),
    WorkspacePage(
        key="organize",
        number="02",
        title="发票整理",
        navigation_label="发票整理",
        subtitle="扫描文件名并预览归档目标，确认后移动到公司目录。",
        workflow="选择目录  →  扫描预览  →  确认  →  整理  →  可撤销",
    ),
    WorkspacePage(
        key="history",
        number="03",
        title="任务历史",
        navigation_label="任务历史",
        subtitle="搜索任务、查看处理详情，并对可信记录执行安全回滚。",
        workflow="搜索筛选  →  查看详情  →  打开目录 / 安全回滚",
    ),
    WorkspacePage(
        key="settings",
        number="04",
        title="设置与问题诊断",
        navigation_label="设置与诊断",
        subtitle="管理识别规则、界面偏好、日志和最近错误。",
        workflow="规则预设  →  界面偏好  →  日志与诊断",
    ),
)


class WorkspacePageStack(tk.Frame):
    """Small tabless page container compatible with the app's select/index calls."""

    def __init__(self, master: tk.Widget, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self._pages: list[tk.Widget] = []
        self._selected_index = 0
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def add(self, child: tk.Widget, **_options) -> None:
        if child in self._pages:
            raise ValueError("页面已经添加到工作台")
        self._pages.append(child)
        child.grid(row=0, column=0, sticky="nsew")
        if len(self._pages) == 1:
            child.tkraise()
        else:
            child.grid_remove()

    def _resolve_index(self, page: object) -> int:
        if isinstance(page, int):
            index = page
        else:
            page_path = str(page)
            index = next(
                (position for position, candidate in enumerate(self._pages) if str(candidate) == page_path),
                -1,
            )
        if index < 0 or index >= len(self._pages):
            raise ValueError(f"未知工作台页面：{page}")
        return index

    def select(self, page: object | None = None) -> str:
        if not self._pages:
            return ""
        if page is not None:
            target_index = self._resolve_index(page)
            if target_index != self._selected_index:
                self._pages[self._selected_index].grid_remove()
                self._selected_index = target_index
                self._pages[self._selected_index].grid()
            self._pages[self._selected_index].tkraise()
        return str(self._pages[self._selected_index])

    def index(self, page: object) -> int:
        if page == "current":
            return self._selected_index
        return self._resolve_index(page)


class InvoiceToolApp(V521InvoiceToolApp):
    """Task-oriented workspace shell for the v5.3 desktop application."""

    BASE_TK_SCALING = 96 / 72
    NAV_BASE_WIDTH = 212

    @staticmethod
    def _dpi_scale(root: tk.Tk) -> float:
        try:
            scaling = float(root.tk.call("tk", "scaling"))
        except (tk.TclError, TypeError, ValueError):
            scaling = InvoiceToolApp.BASE_TK_SCALING
        return min(2.25, max(1.0, scaling / InvoiceToolApp.BASE_TK_SCALING))

    def _scaled(self, value: int) -> int:
        return max(1, round(value * self._workspace_scale))

    def _apply_initial_window_geometry(self) -> None:
        self._workspace_scale = self._dpi_scale(self.root)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        horizontal_margin = self._scaled(56)
        vertical_margin = self._scaled(72)
        available_width = max(1, screen_width - horizontal_margin)
        available_height = max(1, screen_height - vertical_margin)
        width = min(self._scaled(1360), available_width)
        height = min(self._scaled(820), available_height)
        min_width = min(self._scaled(1040), available_width)
        min_height = min(self._scaled(650), available_height)
        pos_x = max((screen_width - width) // 2, 0)
        pos_y = max((screen_height - height) // 2, 0)
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        self.root.minsize(min_width, min_height)

    @property
    def workspace_pages(self) -> Tuple[WorkspacePage, ...]:
        return WORKSPACE_PAGES

    def _build_brand(self, parent: tk.Widget) -> None:
        palette = self.palette
        brand = tk.Frame(parent, bg=palette["hero_card_bg"])
        brand.pack(fill="x", padx=self._scaled(16), pady=(self._scaled(18), self._scaled(20)))
        tk.Label(
            brand,
            text="INVOICE DESK",
            font=("Segoe UI", 8, "bold"),
            bg=palette["hero_card_bg"],
            fg=palette["hero_accent"],
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            brand,
            text="发票工作台",
            font=("微软雅黑", 15, "bold"),
            bg=palette["hero_card_bg"],
            fg=palette["title_fg"],
            anchor="w",
        ).pack(fill="x", pady=(self._scaled(5), 0))
        tk.Label(
            brand,
            text=APP_VERSION,
            font=("Segoe UI", 8),
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
            anchor="w",
        ).pack(fill="x", pady=(self._scaled(2), 0))

    def _build_navigation(self, parent: tk.Widget) -> None:
        palette = self.palette
        tk.Label(
            parent,
            text="任务",
            font=("微软雅黑", 8, "bold"),
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
            anchor="w",
        ).pack(fill="x", padx=self._scaled(18), pady=(0, self._scaled(7)))

        self.workspace_nav_buttons: Dict[str, tk.Button] = {}
        for page in self.workspace_pages:
            button = tk.Button(
                parent,
                text=f"{page.number}   {page.navigation_label}",
                font=("微软雅黑", 10, "bold"),
                anchor="w",
                padx=self._scaled(16),
                pady=self._scaled(10),
                relief="flat",
                bd=0,
                cursor="hand2",
                takefocus=True,
                highlightthickness=self._scaled(2),
                command=lambda page_key=page.key: self._select_workspace_page(page_key),
            )
            button.pack(fill="x", padx=self._scaled(10), pady=self._scaled(2))
            self.workspace_nav_buttons[page.key] = button

        tk.Frame(
            parent,
            height=1,
            bg=palette["hero_card_border"],
        ).pack(fill="x", padx=self._scaled(16), pady=(self._scaled(18), self._scaled(14)))

        tk.Label(
            parent,
            text="快捷键",
            font=("微软雅黑", 8, "bold"),
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
            anchor="w",
        ).pack(fill="x", padx=self._scaled(18))
        tk.Label(
            parent,
            text="Alt + 1…4  切换任务\nTab  在当前页面移动焦点",
            font=("微软雅黑", 8),
            justify="left",
            bg=palette["hero_card_bg"],
            fg=palette["title_muted"],
            anchor="w",
        ).pack(fill="x", padx=self._scaled(18), pady=(self._scaled(5), 0))

    def _build_capability_summary(self, parent: tk.Widget) -> None:
        palette = self.palette
        capability_items = (
            ("Excel", PANDAS_SUPPORT),
            ("拖放", DND_SUPPORT),
            ("报告", OPENPYXL_SUPPORT),
            ("现代主题", MODERN_UI),
        )
        ready_count = sum(1 for _name, available in capability_items if available)
        card = tk.Frame(
            parent,
            bg=palette["hero_chip_bg"],
            highlightbackground=palette["hero_card_border"],
            highlightcolor=palette["hero_card_border"],
            highlightthickness=1,
            padx=self._scaled(12),
            pady=self._scaled(10),
        )
        card.pack(side="bottom", fill="x", padx=self._scaled(12), pady=self._scaled(14))
        tk.Label(
            card,
            text=f"运行能力  {ready_count}/{len(capability_items)}",
            font=("微软雅黑", 8, "bold"),
            bg=palette["hero_chip_bg"],
            fg=palette["hero_chip_fg"],
            anchor="w",
        ).pack(fill="x")
        capability_text = " · ".join(name for name, available in capability_items if available)
        tk.Label(
            card,
            text=capability_text or "核心依赖尚未就绪",
            font=("微软雅黑", 8),
            bg=palette["hero_chip_bg"],
            fg=palette["title_muted"],
            anchor="w",
            wraplength=self._scaled(172),
            justify="left",
        ).pack(fill="x", pady=(self._scaled(4), 0))

    def _build_page_header(self, parent: tk.Widget) -> None:
        palette = self.palette
        header = tk.Frame(
            parent,
            bg=palette["surface"],
            padx=self._scaled(18),
            pady=self._scaled(13),
        )
        header.pack(fill="x")

        title_column = tk.Frame(header, bg=palette["surface"])
        title_column.pack(side="left", fill="x", expand=True)
        self.workspace_page_title = tk.StringVar()
        self.workspace_page_subtitle = tk.StringVar()
        self.workspace_workflow = tk.StringVar()
        tk.Label(
            title_column,
            textvariable=self.workspace_page_title,
            font=("微软雅黑", 15, "bold"),
            bg=palette["surface"],
            fg=palette["text"],
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            title_column,
            textvariable=self.workspace_page_subtitle,
            font=("微软雅黑", 8),
            bg=palette["surface"],
            fg=palette["muted"],
            anchor="w",
        ).pack(fill="x", pady=(self._scaled(3), 0))
        tk.Label(
            title_column,
            textvariable=self.workspace_workflow,
            font=("微软雅黑", 8, "bold"),
            bg=palette["surface"],
            fg=palette["primary"],
            anchor="w",
        ).pack(fill="x", pady=(self._scaled(6), 0))

        actions = tk.Frame(header, bg=palette["surface"])
        actions.pack(side="right", padx=(self._scaled(18), 0), anchor="ne")
        self.workspace_page_counter = tk.Label(
            actions,
            text="",
            font=("Segoe UI", 8, "bold"),
            bg=palette["surface_soft"],
            fg=palette["muted"],
            padx=self._scaled(9),
            pady=self._scaled(5),
        )
        self.workspace_page_counter.pack(side="left", padx=(0, self._scaled(8)))
        self.theme_toggle_btn = tk.Button(
            actions,
            text="夜间" if self.ui_theme.get() == "day" else "白天",
            font=("微软雅黑", 9),
            padx=self._scaled(11),
            pady=self._scaled(5),
            command=self._toggle_ui_theme,
        )
        self.theme_toggle_btn.pack(side="left")
        self._style_action_button(self.theme_toggle_btn, "neutral")

    def _build_page_container(self, parent: tk.Widget) -> None:
        palette = self.palette
        content_border = tk.Frame(
            parent,
            bg=palette["border"],
            padx=1,
            pady=1,
        )
        content_border.pack(fill="both", expand=True, padx=self._scaled(12), pady=(0, self._scaled(8)))
        content = tk.Frame(content_border, bg=palette["surface"])
        content.pack(fill="both", expand=True)

        self.notebook = WorkspacePageStack(
            content,
            bg=palette["surface"],
            takefocus=False,
            highlightthickness=0,
            bd=0,
        )
        self.notebook.pack(fill="both", expand=True)

        tab_padding = self._scaled(6)
        self.filter_frame = ttk.Frame(self.notebook, padding=tab_padding)
        self.notebook.add(self.filter_frame, text="Excel 筛选")
        self.organize_frame = ttk.Frame(self.notebook, padding=tab_padding)
        self.notebook.add(self.organize_frame, text="发票整理")
        self.history_frame = ttk.Frame(self.notebook, padding=tab_padding)
        self.notebook.add(self.history_frame, text="任务历史")
        self.settings_frame = ttk.Frame(self.notebook, padding=tab_padding)
        self.notebook.add(self.settings_frame, text="设置与诊断")

        self._workspace_frames = {
            "organize": self.organize_frame,
            "filter": self.filter_frame,
            "history": self.history_frame,
            "settings": self.settings_frame,
        }

        self._build_organize_tab()
        self._build_filter_tab()
        self._build_history_tab()
        self._build_settings_tab()
        self._polish_action_text()

    def _build_status_bar(self) -> None:
        palette = self.palette
        status_frame = tk.Frame(
            self.root,
            bg=palette["status_bg"],
            highlightbackground=palette["border"],
            highlightcolor=palette["border"],
            highlightthickness=1,
        )
        status_frame.pack(fill="x", side="bottom")
        tk.Label(
            status_frame,
            text="●",
            font=("Segoe UI", 8),
            bg=palette["status_bg"],
            fg=palette["success"],
            padx=self._scaled(10),
        ).pack(side="left")
        self.status_var = tk.StringVar(value="就绪，可从左侧选择任务开始处理。")
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("微软雅黑", 9),
            anchor="w",
            padx=0,
            pady=self._scaled(6),
            bg=palette["status_bg"],
            fg=palette["status_fg"],
        ).pack(side="left", fill="x", expand=True)
        self.progress_label = tk.Label(
            status_frame,
            text="",
            font=("微软雅黑", 9),
            fg=palette["muted"],
            bg=palette["status_bg"],
            padx=self._scaled(12),
            pady=self._scaled(6),
        )
        self.progress_label.pack(side="right")

    def _build_ui(self) -> None:
        palette = self.palette
        self.root.configure(bg=palette["root_bg"])
        self._configure_ttk_styles()

        shell = tk.Frame(self.root, bg=palette["root_bg"])
        shell.pack(fill="both", expand=True)

        self.workspace_sidebar = tk.Frame(
            shell,
            width=self._scaled(self.NAV_BASE_WIDTH),
            bg=palette["hero_card_bg"],
        )
        self.workspace_sidebar.pack(side="left", fill="y")
        self.workspace_sidebar.pack_propagate(False)
        self._build_brand(self.workspace_sidebar)
        self._build_navigation(self.workspace_sidebar)
        self._build_capability_summary(self.workspace_sidebar)

        self.workspace_main = tk.Frame(shell, bg=palette["root_bg"])
        self.workspace_main.pack(side="left", fill="both", expand=True)
        self._build_page_header(self.workspace_main)
        self._build_page_container(self.workspace_main)

        self._build_log_drawer()
        self._build_status_bar()

        saved_page = str(
            getattr(self, "_workspace_page_key", self.config.get("workspace_page", "filter"))
        )
        if saved_page not in self._workspace_frames:
            saved_page = "filter"
        self._select_workspace_page(saved_page, persist=False, focus_navigation=False)
        self._bind_workspace_shortcuts()
        self._apply_theme_to_widget_tree(self.root)

    def _page_metadata(self, page_key: str) -> WorkspacePage:
        return next(page for page in self.workspace_pages if page.key == page_key)

    def _style_navigation_button(self, page_key: str, active: bool) -> None:
        palette = self.palette
        button = self.workspace_nav_buttons[page_key]
        if active:
            normal_bg = palette["primary"]
            hover_bg = palette["primary_hover"]
            foreground = "#FFFFFF"
            focus_color = palette["hero_accent"]
        else:
            normal_bg = palette["hero_card_bg"]
            hover_bg = palette["hero_chip_bg"]
            foreground = palette["title_muted"]
            focus_color = palette["hero_accent"]
        button.configure(
            bg=normal_bg,
            fg=foreground,
            activebackground=hover_bg,
            activeforeground="#FFFFFF" if active else palette["title_fg"],
            highlightbackground=palette["hero_card_bg"],
            highlightcolor=focus_color,
        )

        def enter(_event, key=page_key) -> None:
            if getattr(self, "_workspace_page_key", "") != key:
                self.workspace_nav_buttons[key].configure(
                    bg=palette["hero_chip_bg"],
                    fg=palette["title_fg"],
                )

        def leave(_event, key=page_key) -> None:
            self._style_navigation_button(
                key,
                getattr(self, "_workspace_page_key", "") == key,
            )

        button.bind("<Enter>", enter)
        button.bind("<Leave>", leave)

    def _sync_workspace_navigation(self, page_key: str) -> None:
        page = self._page_metadata(page_key)
        self._workspace_page_key = page_key
        self.workspace_page_title.set(page.title)
        self.workspace_page_subtitle.set(page.subtitle)
        self.workspace_workflow.set(page.workflow)
        page_position = next(index for index, item in enumerate(self.workspace_pages, 1) if item.key == page_key)
        self.workspace_page_counter.configure(text=f"{page_position:02d} / {len(self.workspace_pages):02d}")
        for key in self.workspace_nav_buttons:
            self._style_navigation_button(key, key == page_key)

    def _select_workspace_page(
        self,
        page_key: str,
        *,
        persist: bool = True,
        focus_navigation: bool = True,
    ) -> None:
        if page_key not in self._workspace_frames:
            raise ValueError(f"未知工作台页面：{page_key}")
        target_frame = str(self._workspace_frames[page_key])
        if self.notebook.select() != target_frame:
            self.notebook.select(self._workspace_frames[page_key])
        self._sync_workspace_navigation(page_key)
        if persist:
            self.config["workspace_page"] = page_key
            self._save_config()
        if focus_navigation:
            self.workspace_nav_buttons[page_key].focus_set()

    def _bind_workspace_shortcuts(self) -> None:
        for index, page in enumerate(self.workspace_pages, 1):
            self.root.bind(
                f"<Alt-Key-{index}>",
                lambda _event, page_key=page.key: self._navigate_workspace_shortcut(page_key),
            )

    def _navigate_workspace_shortcut(self, page_key: str) -> str:
        self._select_workspace_page(page_key)
        return "break"
