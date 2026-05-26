from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, VERTICAL, Y, DoubleVar, Toplevel, filedialog, messagebox
from tkinter import Tk, StringVar, Text
from tkinter import ttk

from hr_toolkit import __version__
from hr_toolkit.app_update import (
    UpdateInfo,
    check_for_update,
    download_update_package,
    launch_update_replacement,
    update_check_enabled,
    update_manifest_url,
)
from hr_toolkit.tools.folder_rename import (
    MODE_APPEND,
    MODE_REMOVE,
    MODE_REPLACE,
    rename_person_folders,
)
from hr_toolkit.tools.personnel_change_merge import merge_personnel_changes
from hr_toolkit.tools.salary_merge import merge_monthly_salary
from hr_toolkit.tools.salary_split import split_salary_by_company


RENAME_MODE_LABELS = {
    "追加文字": MODE_APPEND,
    "删除结尾文字": MODE_REMOVE,
    "修改单人名称": MODE_REPLACE,
}

TOOL_NAV_ITEMS = (
    ("salary_split", "需求4  工资表拆分"),
    ("salary_merge", "需求5  工资表合并"),
    ("personnel_change_merge", "需求6  异动表汇总"),
    ("folder_rename", "需求8  文件夹改名"),
    ("archive_import", "需求7  档案入库（待实现）"),
)

COLOR_BG = "#f7f8fa"
COLOR_SIDEBAR = "#f0f2f5"
COLOR_SURFACE = "#ffffff"
COLOR_BORDER = "#e4e7ec"
COLOR_TEXT = "#1a1d23"
COLOR_MUTED = "#8c95a6"
COLOR_PRIMARY = "#2d6ef5"
COLOR_PRIMARY_ACTIVE = "#1a5ae0"
COLOR_NAV_SELECTED = "#ffffff"
COLOR_SUCCESS = "#0a7c4e"
COLOR_WARNING = "#c0392b"
COLOR_TUTORIAL_BG = "#f0f4ff"
COLOR_TUTORIAL_BORDER = "#c7d7fb"


class HRToolkitApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(f"HR工具箱 v{__version__}")
        self.root.geometry("1000x680")
        self.root.minsize(900, 620)
        self.root.configure(bg=COLOR_BG)

        self.current_tool = "salary_split"
        self.nav_buttons: dict[str, ttk.Button] = {}
        self.tool_title = StringVar()
        self.tool_description = StringVar()
        self.input_label = StringVar()
        self.choose_input_text = StringVar()
        self.run_button_text = StringVar()
        self.summary_label = StringVar()
        self.summary_button_text = StringVar()
        self.rename_mode = StringVar(value="追加文字")
        self.rename_target_label = StringVar(value="姓名（可不填）")
        self.rename_text_label = StringVar(value="要追加的文字")
        self.rename_replacement_label = StringVar(value="新名称")
        self.rename_target_name = StringVar()
        self.rename_text = StringVar()
        self.rename_replacement_name = StringVar()
        self.input_path = StringVar()
        self.summary_path = StringVar()
        self.output_dir = StringVar(value=str(default_output_parent_dir(self.current_tool)))
        self.output_dir_user_selected = False
        self.status_queue: queue.Queue[tuple[str, object | None]] = queue.Queue()
        self.update_queue: queue.Queue[tuple[str, object | None]] = queue.Queue()
        self.last_output_dir: Path | None = None
        self.pending_update: UpdateInfo | None = None
        self.update_window: Toplevel | None = None
        self.update_progress_var: DoubleVar | None = None
        self.update_progress_label: ttk.Label | None = None
        self.update_check_in_progress = False
        self.manual_update_check_active = False

        self._configure_style()
        self._set_tool_texts()
        self._build_layout()
        self._poll_status_queue()
        self._poll_update_queue()
        self.root.after(600, self._check_updates_on_startup)

    def _configure_style(self) -> None:
        if sys.platform == "darwin":
            family = "PingFang SC"
            mono_family = "Menlo"
        elif sys.platform.startswith("win"):
            family = "Microsoft YaHei UI"
            mono_family = "Consolas"
        else:
            family = "Arial"
            mono_family = "DejaVu Sans Mono"
        self.base_font = (family, 10)
        self.small_font = (family, 9)
        self.tiny_font = (family, 8)
        self.title_font = (family, 18, "bold")
        self.section_font = (family, 10, "bold")
        self.nav_font = (family, 10)
        self.nav_selected_font = (family, 10, "bold")
        self.mono_font = (mono_family, 10)

        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure(".", font=self.base_font, background=COLOR_BG, foreground=COLOR_TEXT)
        style.configure("App.TFrame", background=COLOR_BG)
        style.configure("Sidebar.TFrame", background=COLOR_SIDEBAR)
        style.configure("Content.TFrame", background=COLOR_BG)
        style.configure("Card.TFrame", background=COLOR_SURFACE)
        style.configure("InputWrap.TFrame", background=COLOR_SURFACE)
        style.configure(
            "Tutorial.TFrame",
            background=COLOR_TUTORIAL_BG,
            bordercolor=COLOR_TUTORIAL_BORDER,
            lightcolor=COLOR_TUTORIAL_BORDER,
            darkcolor=COLOR_TUTORIAL_BORDER,
            borderwidth=1,
            relief="solid",
        )
        style.configure("Tooltip.TFrame", background=COLOR_SURFACE, relief="solid", borderwidth=1, bordercolor=COLOR_BORDER)
        style.configure("NavRow.TFrame", background=COLOR_SIDEBAR)
        style.configure("NavIndicator.TFrame", background=COLOR_SIDEBAR)
        style.configure("NavIndicatorSelected.TFrame", background=COLOR_PRIMARY)
        style.configure("Title.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=self.title_font)
        style.configure("Subtitle.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=self.base_font)
        style.configure("Section.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=self.small_font)
        style.configure("SidebarTitle.TLabel", background=COLOR_SIDEBAR, foreground=COLOR_TEXT, font=(self.base_font[0], 14, "bold"))
        style.configure("SidebarMuted.TLabel", background=COLOR_SIDEBAR, foreground=COLOR_MUTED, font=self.small_font)
        style.configure("Version.TLabel", background=COLOR_SIDEBAR, foreground=COLOR_MUTED, font=self.tiny_font)
        style.configure("TutorialTitle.TLabel", background=COLOR_TUTORIAL_BG, foreground=COLOR_TEXT, font=self.section_font)
        style.configure("Tooltip.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.small_font, padding=(8, 6))
        style.configure(
            "Card.TLabelframe",
            background=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            relief="solid",
        )
        style.configure("Card.TLabelframe.Label", background=COLOR_BG, foreground=COLOR_TEXT, font=self.section_font)
        style.configure("Rename.TLabelframe", background=COLOR_SURFACE, bordercolor=COLOR_BORDER, relief="solid")
        style.configure("Rename.TLabelframe.Label", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.section_font)
        style.configure("App.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=self.base_font)
        style.configure(
            "App.TEntry",
            fieldbackground=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            insertcolor=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            padding=(8, 6),
            relief="solid",
        )
        style.map(
            "App.TEntry",
            bordercolor=[("focus", COLOR_PRIMARY)],
            lightcolor=[("focus", COLOR_PRIMARY)],
            darkcolor=[("focus", COLOR_PRIMARY)],
        )
        style.configure(
            "App.TCombobox",
            fieldbackground=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            arrowcolor=COLOR_MUTED,
            padding=(8, 5),
        )
        style.configure("Nav.TButton", anchor="w", padding=(10, 8), background=COLOR_SIDEBAR, foreground=COLOR_MUTED, borderwidth=0, font=self.nav_font, relief="flat")
        style.configure("NavSelected.TButton", anchor="w", padding=(10, 8), background=COLOR_NAV_SELECTED, foreground=COLOR_TEXT, borderwidth=0, font=self.nav_selected_font, relief="flat")
        style.map("Nav.TButton", background=[("active", "#e8ebf0")], foreground=[("active", COLOR_TEXT)])
        style.map("NavSelected.TButton", background=[("active", COLOR_NAV_SELECTED)])
        style.configure("Primary.TButton", padding=(14, 7), background=COLOR_PRIMARY, foreground="#ffffff", borderwidth=0, font=(self.base_font[0], 10, "bold"), relief="flat")
        style.map("Primary.TButton", background=[("active", COLOR_PRIMARY_ACTIVE), ("disabled", COLOR_BORDER)], foreground=[("disabled", COLOR_MUTED)])
        style.configure("Secondary.TButton", padding=(10, 6), background=COLOR_SURFACE, foreground=COLOR_TEXT, bordercolor=COLOR_BORDER, lightcolor=COLOR_BORDER, darkcolor=COLOR_BORDER, relief="solid")
        style.map("Secondary.TButton", background=[("active", "#e8ebf0")], bordercolor=[("active", COLOR_PRIMARY)])
        style.configure("Icon.TButton", padding=(7, 5), background=COLOR_SURFACE, foreground=COLOR_MUTED, borderwidth=0, relief="flat", font=(self.base_font[0], 10, "bold"))
        style.map("Icon.TButton", background=[("active", "#e8ebf0")], foreground=[("active", COLOR_TEXT)])

    def _build_layout(self) -> None:
        root_frame = ttk.Frame(self.root, padding=0, style="App.TFrame")
        root_frame.pack(fill=BOTH, expand=True)

        left_frame = ttk.Frame(root_frame, width=200, padding=(18, 22, 14, 18), style="Sidebar.TFrame")
        left_frame.pack(side=LEFT, fill=Y)
        left_frame.pack_propagate(False)

        ttk.Label(left_frame, text="HR工具箱", style="SidebarTitle.TLabel").pack(anchor="w")
        self.nav_indicators = {}

        nav_frame = ttk.Frame(left_frame, style="Sidebar.TFrame")
        nav_frame.pack(fill="x", pady=(28, 0))
        for tool_id, label in TOOL_NAV_ITEMS:
            row = ttk.Frame(nav_frame, style="NavRow.TFrame")
            row.pack(fill="x", pady=2)
            indicator = ttk.Frame(
                row,
                width=3,
                style="NavIndicatorSelected.TFrame" if tool_id == self.current_tool else "NavIndicator.TFrame",
            )
            indicator.pack(side=LEFT, fill=Y, padx=(0, 6))
            button = ttk.Button(
                row,
                text=label,
                style="NavSelected.TButton" if tool_id == self.current_tool else "Nav.TButton",
                command=lambda selected=tool_id: self._select_tool(selected),
            )
            button.pack(side=LEFT, fill="x", expand=True)
            self.nav_buttons[tool_id] = button
            self.nav_indicators[tool_id] = indicator

        ttk.Label(left_frame, text=f"v{__version__}", style="Version.TLabel").pack(side="bottom", anchor="w")

        right_frame = ttk.Frame(root_frame, padding=(32, 28, 28, 24), style="Content.TFrame")
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True)

        title_row = ttk.Frame(right_frame, style="Content.TFrame")
        title_row.pack(fill="x")
        title_row.columnconfigure(0, weight=1)
        ttk.Label(title_row, textvariable=self.tool_title, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        title_actions = ttk.Frame(title_row, style="Content.TFrame")
        title_actions.grid(row=0, column=1, sticky="e")
        self.check_update_button = ttk.Button(
            title_actions,
            text="检查更新",
            command=self._check_updates_manually,
            style="Secondary.TButton",
        )
        self.check_update_button.pack(side=LEFT)
        self.tutorial_toggle_button = ttk.Button(title_actions, text="使用教程", style="Secondary.TButton")
        self.tutorial_toggle_button.pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            right_frame,
            textvariable=self.tool_description,
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(6, 18))

        self.tutorial_frame = ttk.Frame(right_frame, padding=12, style="Tutorial.TFrame")
        ttk.Label(self.tutorial_frame, text="使用教程", style="TutorialTitle.TLabel").pack(anchor="w", pady=(0, 6))
        self.tutorial_text = Text(
            self.tutorial_frame,
            height=6,
            wrap="word",
            padx=10,
            pady=8,
            bg=COLOR_TUTORIAL_BG,
            fg=COLOR_TEXT,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLOR_TUTORIAL_BORDER,
            highlightcolor=COLOR_TUTORIAL_BORDER,
            font=self.base_font,
        )
        self.tutorial_text.pack(fill="x")
        self.tutorial_text.tag_configure("strong", font=(self.base_font[0], 10, "bold"))
        self.tutorial_text.tag_configure("warning", foreground=COLOR_WARNING, font=(self.base_font[0], 10, "bold"))
        self._set_tutorial_text()

        form = ttk.Frame(right_frame, style="Content.TFrame")
        form.pack(fill="x")

        def make_input_row(row_index: int, label_text, value_var: StringVar, command) -> tuple[ttk.Label, ttk.Frame]:
            if isinstance(label_text, StringVar):
                label = ttk.Label(form, textvariable=label_text, style="App.TLabel")
            else:
                label = ttk.Label(form, text=label_text, style="App.TLabel")
            label.grid(row=row_index, column=0, sticky="w", pady=4)
            input_frame = ttk.Frame(form, style="InputWrap.TFrame")
            input_frame.grid(row=row_index, column=1, sticky="ew", padx=(10, 0), pady=4)
            entry = ttk.Entry(input_frame, textvariable=value_var, style="App.TEntry")
            entry.pack(side=LEFT, fill=BOTH, expand=True)
            ttk.Button(input_frame, text="...", width=3, command=command, style="Icon.TButton").pack(side=RIGHT, padx=(4, 0))
            return label, input_frame

        make_input_row(0, self.input_label, self.input_path, self._choose_input)
        self.summary_label_widget, self.summary_entry_widget = make_input_row(
            1,
            self.summary_label,
            self.summary_path,
            self._choose_summary,
        )
        self.output_label_widget, self.output_entry_widget = make_input_row(
            2,
            "保存位置",
            self.output_dir,
            self._choose_output,
        )

        self.rename_options_frame = ttk.LabelFrame(form, text="文件夹改名", padding=10, style="Rename.TLabelframe")
        self.rename_options_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Label(self.rename_options_frame, text="操作", style="App.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        self.rename_mode_widget = ttk.Combobox(
            self.rename_options_frame,
            textvariable=self.rename_mode,
            values=list(RENAME_MODE_LABELS.keys()),
            state="readonly",
            width=16,
            style="App.TCombobox",
        )
        self.rename_mode_widget.grid(row=0, column=1, sticky="w", padx=10, pady=4)
        self.rename_mode_widget.bind("<<ComboboxSelected>>", self._on_rename_mode_changed)

        ttk.Label(self.rename_options_frame, textvariable=self.rename_target_label, style="App.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        self.rename_target_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_target_name, style="App.TEntry")
        self.rename_target_widget.grid(row=1, column=1, sticky="ew", padx=10, pady=4)

        self.rename_text_label_widget = ttk.Label(self.rename_options_frame, textvariable=self.rename_text_label, style="App.TLabel")
        self.rename_text_label_widget.grid(row=2, column=0, sticky="w", pady=4)
        self.rename_text_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_text, style="App.TEntry")
        self.rename_text_widget.grid(row=2, column=1, sticky="ew", padx=10, pady=4)

        self.rename_replacement_label_widget = ttk.Label(self.rename_options_frame, textvariable=self.rename_replacement_label, style="App.TLabel")
        self.rename_replacement_label_widget.grid(row=3, column=0, sticky="w", pady=4)
        self.rename_replacement_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_replacement_name, style="App.TEntry")
        self.rename_replacement_widget.grid(row=3, column=1, sticky="ew", padx=10, pady=4)
        self.rename_options_frame.columnconfigure(1, weight=1)
        form.columnconfigure(1, weight=1)
        self._update_summary_controls()
        self._update_output_controls()
        self._update_rename_controls()

        self.tutorial_expanded = False

        def toggle_tutorial() -> None:
            self.tutorial_expanded = not self.tutorial_expanded
            if self.tutorial_expanded:
                self.tutorial_frame.pack(fill="x", pady=(0, 16), before=form)
                self.tutorial_toggle_button.configure(text="收起教程")
            else:
                self.tutorial_frame.pack_forget()
                self.tutorial_toggle_button.configure(text="使用教程")

        self.tutorial_toggle_button.configure(command=toggle_tutorial)

        actions = ttk.Frame(right_frame, style="Content.TFrame")
        actions.pack(fill="x", pady=(12, 14))
        run_button_box = ttk.Frame(actions, width=120, height=36, style="Content.TFrame")
        run_button_box.pack(side=LEFT)
        run_button_box.pack_propagate(False)
        self.run_button = ttk.Button(run_button_box, textvariable=self.run_button_text, command=self._run_current_tool, style="Primary.TButton")
        self.run_button.pack(fill=BOTH, expand=True)
        self.open_button = ttk.Button(actions, text="打开结果目录", command=self._open_output_dir, style="Secondary.TButton")
        self.open_button.pack(side=LEFT, padx=(8, 0))

        ttk.Label(right_frame, text="日志", style="Section.TLabel").pack(anchor="w")
        log_frame = ttk.Frame(right_frame, style="Content.TFrame")
        log_frame.pack(fill=BOTH, expand=True, pady=(6, 0))
        scrollbar = ttk.Scrollbar(log_frame, orient=VERTICAL)
        self.log_text = Text(
            log_frame,
            height=12,
            wrap="word",
            yscrollcommand=scrollbar.set,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_PRIMARY,
            padx=10,
            pady=10,
            font=self.mono_font,
        )
        self.log_text.tag_configure("success", foreground=COLOR_SUCCESS)
        self.log_text.tag_configure("warning", foreground=COLOR_WARNING)
        self.log_text.tag_configure("muted", foreground=COLOR_MUTED)
        scrollbar.config(command=self.log_text.yview)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        self._write_log(self._initial_log_text())

    def _check_updates_on_startup(self) -> None:
        if not update_check_enabled():
            return
        self._start_update_check(manual=False)

    def _check_updates_manually(self) -> None:
        self._start_update_check(manual=True)

    def _start_update_check(self, manual: bool) -> None:
        if self.update_check_in_progress:
            if manual:
                messagebox.showinfo("正在检查更新", "更新检查正在进行，请稍候。", parent=self.root)
            return
        self.update_check_in_progress = True
        self.manual_update_check_active = manual
        manifest_url = update_manifest_url()
        if hasattr(self, "check_update_button"):
            self.check_update_button.config(state="disabled")
        self._write_log("正在检查更新...")
        self._write_log(f"更新配置：{manifest_url}")
        self._show_update_checking_window()
        worker = threading.Thread(target=self._update_check_worker, daemon=True)
        worker.start()

    def _update_check_worker(self) -> None:
        try:
            update = check_for_update(__version__)
        except Exception as exc:
            self.update_queue.put(("check_error", exc))
            return
        if update is None:
            self.update_queue.put(("no_update", None))
        else:
            self.update_queue.put(("available", update))

    def _poll_update_queue(self) -> None:
        try:
            while True:
                status, payload = self.update_queue.get_nowait()
                if status == "no_update":
                    manual = self.manual_update_check_active
                    self._finish_update_check()
                    self._close_update_window()
                    self._write_log("已是最新版本。")
                    if manual:
                        messagebox.showinfo("检查更新", "当前已经是最新版本。", parent=self.root)
                elif status == "check_error":
                    manual = self.manual_update_check_active
                    self._finish_update_check()
                    self._close_update_window()
                    self._write_log(f"更新检查失败，可继续使用：{payload}")
                    if manual:
                        messagebox.showerror("检查更新失败", str(payload), parent=self.root)
                elif status == "available":
                    self._finish_update_check()
                    self._close_update_window()
                    self._show_required_update(payload)
                elif status == "download_progress":
                    downloaded, total = payload
                    self._update_download_progress(downloaded, total)
                elif status == "download_ready":
                    self._finish_update_download(payload)
                elif status == "download_error":
                    self._handle_update_failure(payload)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_update_queue)

    def _finish_update_check(self) -> None:
        self.update_check_in_progress = False
        self.manual_update_check_active = False
        if hasattr(self, "check_update_button"):
            self.check_update_button.config(state="normal")

    def _show_update_checking_window(self) -> None:
        self._close_update_window()
        self.update_window = Toplevel(self.root)
        self.update_window.title("检查更新")
        self._center_window(self.update_window, 360, 120)
        self.update_window.resizable(False, False)
        self.update_window.configure(bg=COLOR_BG)
        self.update_window.transient(self.root)
        self.update_window.grab_set()
        self.update_window.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = ttk.Frame(self.update_window, padding=18, style="Content.TFrame")
        frame.pack(fill=BOTH, expand=True)
        ttk.Label(frame, text="正在检查更新，请稍候...", style="App.TLabel").pack(anchor="w", pady=(2, 10))
        ttk.Progressbar(frame, orient="horizontal", mode="indeterminate").pack(fill="x")
        for child in frame.winfo_children():
            if isinstance(child, ttk.Progressbar):
                child.start(12)

    def _close_update_window(self) -> None:
        if self.update_window is not None and self.update_window.winfo_exists():
            self.update_window.grab_release()
            self.update_window.destroy()
        self.update_window = None
        self.update_progress_var = None
        self.update_progress_label = None

    def _show_required_update(self, update: object | None) -> None:
        if not isinstance(update, UpdateInfo):
            return
        self.pending_update = update
        self._write_log(f"发现新版本：v{update.version}")
        self._write_log(f"下载地址：{update.file_url}")
        notes = "\n".join(f"- {line}" for line in update.notes[:6]) or "- 本次发布未填写更新说明"
        message = (
            f"发现新版本 v{update.version}，必须更新后才能继续使用。\n\n"
            f"更新内容：\n{notes}\n\n"
            "点击“确定”开始更新；点击“取消”将退出程序。"
        )
        if not messagebox.askokcancel("发现新版本", message, parent=self.root):
            self.root.destroy()
            return
        self._start_update_download(update)

    def _start_update_download(self, update: UpdateInfo) -> None:
        self._close_update_window()
        self._write_log(f"开始下载更新包：v{update.version}")
        self._write_log(f"下载地址：{update.file_url}")
        self.update_window = Toplevel(self.root)
        self.update_window.title("正在更新 HR工具箱")
        self._center_window(self.update_window, 420, 160)
        self.update_window.resizable(False, False)
        self.update_window.configure(bg=COLOR_BG)
        self.update_window.transient(self.root)
        self.update_window.grab_set()
        self.update_window.protocol("WM_DELETE_WINDOW", self.root.destroy)

        frame = ttk.Frame(self.update_window, padding=18, style="Content.TFrame")
        frame.pack(fill=BOTH, expand=True)
        ttk.Label(frame, text="正在下载更新，请不要关闭程序", style="App.TLabel").pack(anchor="w")
        self.update_progress_label = ttk.Label(frame, text="准备下载...", style="Section.TLabel")
        self.update_progress_label.pack(anchor="w", pady=(10, 6))
        self.update_progress_var = DoubleVar(value=0)
        ttk.Progressbar(
            frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            variable=self.update_progress_var,
        ).pack(fill="x")
        ttk.Button(frame, text="取消并退出", command=self.root.destroy, style="Secondary.TButton").pack(anchor="e", pady=(16, 0))

        worker = threading.Thread(target=self._download_update_worker, args=(update,), daemon=True)
        worker.start()

    def _download_update_worker(self, update: UpdateInfo) -> None:
        def progress(downloaded: int, total: int) -> None:
            self.update_queue.put(("download_progress", (downloaded, total)))

        try:
            package_path = download_update_package(update, progress_callback=progress)
        except Exception as exc:
            self.update_queue.put(("download_error", exc))
            return
        self.update_queue.put(("download_ready", package_path))

    def _update_download_progress(self, downloaded: int, total: int) -> None:
        downloaded_mb = downloaded / 1024 / 1024
        if total > 0:
            percent = min(downloaded / total * 100, 100)
            total_mb = total / 1024 / 1024
            text = f"已下载 {downloaded_mb:.1f} MB / {total_mb:.1f} MB"
        else:
            percent = 0
            text = f"已下载 {downloaded_mb:.1f} MB"
        if self.update_progress_var is not None:
            self.update_progress_var.set(percent)
        if self.update_progress_label is not None:
            self.update_progress_label.configure(text=text)

    def _finish_update_download(self, package_path: object | None) -> None:
        if not isinstance(package_path, Path):
            return
        if self.update_progress_var is not None:
            self.update_progress_var.set(100)
        if self.update_progress_label is not None:
            self.update_progress_label.configure(text="下载完成，正在准备安装...")
        self._write_log("更新包下载完成，正在启动更新程序...")
        try:
            launch_update_replacement(package_path)
        except Exception as exc:
            self._handle_update_failure(exc)
            return
        if self.update_progress_label is not None:
            self.update_progress_label.configure(text="更新程序已启动，当前程序即将退出。")
        self.root.after(700, self.root.destroy)

    def _handle_update_failure(self, exc: object | None) -> None:
        self._write_log(f"更新失败：{exc}")
        messagebox.showerror(
            "更新失败",
            f"更新没有完成，程序将退出。\n\n原因：{exc}\n\n请联系开发重新处理安装包。",
            parent=self.root,
        )
        self.root.destroy()

    def _center_window(self, window: Toplevel, width: int, height: int) -> None:
        self.root.update_idletasks()
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        if root_width <= 1 or root_height <= 1:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = max((screen_width - width) // 2, 0)
            y = max((screen_height - height) // 2, 0)
        else:
            x = root_x + max((root_width - width) // 2, 0)
            y = root_y + max((root_height - height) // 2, 0)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def _select_tool(self, tool_id: str) -> None:
        if tool_id == self.current_tool:
            return
        self.current_tool = tool_id
        self._refresh_nav_buttons()
        self.last_output_dir = None
        self.input_path.set("")
        self.summary_path.set("")
        self.rename_target_name.set("")
        self.rename_text.set("")
        self.rename_replacement_name.set("")
        if not self.output_dir_user_selected:
            self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
        self._set_tool_texts()
        self._clear_log()
        self._write_log(self._initial_log_text())

    def _refresh_nav_buttons(self) -> None:
        for tool_id, button in self.nav_buttons.items():
            style = "NavSelected.TButton" if tool_id == self.current_tool else "Nav.TButton"
            button.configure(style=style)
            indicator = self.nav_indicators.get(tool_id)
            if indicator is not None:
                indicator_style = "NavIndicatorSelected.TFrame" if tool_id == self.current_tool else "NavIndicator.TFrame"
                indicator.configure(style=indicator_style)

    def _set_tool_texts(self) -> None:
        if self.current_tool == "salary_merge":
            self.tool_title.set("需求5：多月工资合并个人薪资汇总")
            self.tool_description.set("选择工资表文件夹；如已有汇总表，可一并选择后追加新月份。")
            self.input_label.set("工资表文件夹")
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("已有汇总表（可选）")
            self.summary_button_text.set("选择汇总表")
            self.run_button_text.set("开始合并")
        elif self.current_tool == "personnel_change_merge":
            self.tool_title.set("需求6：异动表汇总")
            self.tool_description.set("选择包含各项目异动表的文件夹，工具会汇总增员、减员、转正、调动、奖罚扣补。")
            self.input_label.set("异动表文件夹")
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("异动表模板（可选）")
            self.summary_button_text.set("选择模板")
            self.run_button_text.set("开始汇总")
        elif self.current_tool == "folder_rename":
            self.tool_title.set("需求8：人员资料文件夹改名")
            self.tool_description.set("选择人员资料目录，先预览，再确认改名。")
            self.input_label.set("人员文件夹目录")
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
            self.run_button_text.set("预览")
        elif self.current_tool == "salary_split":
            self.tool_title.set("需求4：工资表按入职公司拆分")
            self.tool_description.set("选择一个包含“汇总表”和“明细表”的工资表，工具会按“入职公司”拆成多个公司文件。")
            self.input_label.set("工资表文件")
            self.choose_input_text.set("选择文件")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
            self.run_button_text.set("开始拆分")
        else:
            self.tool_title.set("该工具暂未实现")
            self.tool_description.set("请选择左侧已经可用的工具。")
            self.input_label.set("输入")
            self.choose_input_text.set("选择")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
            self.run_button_text.set("开始")
        if hasattr(self, "summary_label_widget"):
            self._update_summary_controls()
            self._update_output_controls()
            self._update_rename_controls()
        if hasattr(self, "tutorial_text"):
            self._set_tutorial_text()

    def _set_tutorial_text(self) -> None:
        self.tutorial_text.config(state="normal")
        self.tutorial_text.delete("1.0", END)
        for line, tag in self._tutorial_lines():
            if tag:
                self.tutorial_text.insert(END, line + "\n", tag)
            else:
                self.tutorial_text.insert(END, line + "\n")
        self.tutorial_text.config(state="disabled")

    def _tutorial_lines(self) -> list[tuple[str, str | None]]:
        if self.current_tool == "salary_merge":
            return [
                ("适用：把 1-12 个月工资表合成一张个人应发工资汇总表。", "strong"),
                ("步骤：把月度工资表放进同一个文件夹，选择该文件夹；如已有前几月汇总表，再选择“已有汇总表”。", None),
                ("点击“开始合并”后，结果会生成到保存位置下的新文件夹中。", None),
                ("结果：按姓名、身份证号、月份合并；没有工资的月份填 0；已存在的人员月份不会覆盖。", None),
                ("注意：工资表文件名或表内日期要能识别月份；重复人员或重复月份会在执行结果里提醒。", "warning"),
            ]
        if self.current_tool == "personnel_change_merge":
            return [
                ("适用：把多个项目异动表合成一份异动汇总表。", "strong"),
                ("步骤：把各项目异动表放进同一个文件夹，选择该文件夹；有固定模板时再选择“异动表模板”。", None),
                ("点击“开始汇总”后，结果会生成到保存位置下的新文件夹中。", None),
                ("结果：汇总增员、减员、转正、调动、奖罚扣补，并自动重新编号。", None),
                ("注意：需求6目前暂定，正式使用前要确认各项目异动表格式一致。", "warning"),
            ]
        if self.current_tool == "folder_rename":
            return [
                ("适用：批量修改所选目录下第一层人员文件夹名称。", "strong"),
                ("追加文字：姓名不填就是全部文件夹追加；填姓名就是只处理这个人。输入“劳动合同”会追加为“-劳动合同”。", None),
                ("删除结尾文字：输入“_劳动合同”，可删除“张三_劳动合同 / 张三-劳动合同 / 张三劳动合同”的结尾文字。", None),
                ("修改单人名称：填写原姓名和新名称，例如“张三”改为“章五”。", None),
                ("重要提醒：改名会直接改变真实文件夹名称。必须先看预览，确认无误后再点确认；建议操作前先备份。", "warning"),
            ]
        if self.current_tool == "salary_split":
            return [
                ("适用：一个完整工资表按“入职公司”拆成多个公司工资表。", "strong"),
                ("步骤：选择工资表文件，保存位置默认在桌面“工资表拆分结果”，点击“开始拆分”。", None),
                ("点击“打开所在文件夹”可直接查看本次生成的结果目录。", None),
                ("结果：每个入职公司生成一个 Excel，保留表头、格式、公式、小计和底部总计。", None),
                ("注意：源工资表不会被修改；如果模板列名或表结构变化，先发给开发确认。", "warning"),
            ]
        return [
            ("该工具暂未实现。", "strong"),
            ("请选择左侧已完成的工具：需求4、需求5、需求6、需求8。", None),
        ]

    def _update_summary_controls(self) -> None:
        if self.current_tool in {"salary_merge", "personnel_change_merge"}:
            self.summary_label_widget.grid(row=1, column=0, sticky="w", pady=4)
            self.summary_entry_widget.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=4)
            return
        self.summary_label_widget.grid_remove()
        self.summary_entry_widget.grid_remove()

    def _update_output_controls(self) -> None:
        if self.current_tool == "folder_rename":
            self.output_label_widget.grid_remove()
            self.output_entry_widget.grid_remove()
            return
        self.output_label_widget.grid(row=2, column=0, sticky="w", pady=4)
        self.output_entry_widget.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=4)

    def _update_rename_controls(self) -> None:
        if self.current_tool != "folder_rename":
            self.rename_options_frame.grid_remove()
            return
        self.rename_options_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self._update_rename_mode_controls()

    def _on_rename_mode_changed(self, _event=None) -> None:
        self._update_rename_mode_controls()

    def _update_rename_mode_controls(self) -> None:
        mode = RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND)
        if mode == MODE_APPEND:
            self.rename_target_label.set("姓名（可不填）")
            self.rename_text_label.set("要追加的文字")
            self.rename_text_label_widget.grid(row=2, column=0, sticky="w", pady=4)
            self.rename_text_widget.grid(row=2, column=1, sticky="ew", padx=10, pady=4)
            self.rename_text_widget.config(state="normal")
            self.rename_replacement_label_widget.grid_remove()
            self.rename_replacement_widget.grid_remove()
        elif mode == MODE_REMOVE:
            self.rename_target_label.set("姓名（可不填）")
            self.rename_text_label.set("要删除的结尾文字")
            self.rename_text_label_widget.grid(row=2, column=0, sticky="w", pady=4)
            self.rename_text_widget.grid(row=2, column=1, sticky="ew", padx=10, pady=4)
            self.rename_text_widget.config(state="normal")
            self.rename_replacement_label_widget.grid_remove()
            self.rename_replacement_widget.grid_remove()
        else:
            self.rename_target_label.set("原姓名")
            self.rename_replacement_label.set("新名称")
            self.rename_text_label_widget.grid_remove()
            self.rename_text_widget.grid_remove()
            self.rename_replacement_label_widget.grid(row=2, column=0, sticky="w", pady=4)
            self.rename_replacement_widget.grid(row=2, column=1, sticky="ew", padx=10, pady=4)
            self.rename_replacement_widget.config(state="normal")

    def _initial_log_text(self) -> str:
        if self.current_tool == "salary_merge":
            return "请选择工资表文件夹和保存位置，然后点击“开始合并”。已有汇总表是可选项，用于追加新月份。"
        if self.current_tool == "personnel_change_merge":
            return "请选择异动表文件夹和保存位置，然后点击“开始汇总”。模板是可选项，不选时使用文件夹中的第一份异动表作为模板。"
        if self.current_tool == "folder_rename":
            return "请选择人员文件夹目录，填写改名内容，然后点击“预览”。"
        if self.current_tool == "salary_split":
            return "请选择工资表文件和保存位置，然后点击“开始拆分”。"
        return "该工具暂未实现。"

    def _choose_input(self) -> None:
        if self.current_tool in {"salary_merge", "personnel_change_merge", "folder_rename"}:
            if self.current_tool == "personnel_change_merge":
                title = "选择异动表文件夹"
            elif self.current_tool == "folder_rename":
                title = "选择人员文件夹目录"
            else:
                title = "选择工资表文件夹"
            directory = filedialog.askdirectory(title=title)
            if directory:
                self.input_path.set(directory)
                if not self.output_dir_user_selected:
                    self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
            return

        filename = filedialog.askopenfilename(
            title="选择工资表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if filename:
            self.input_path.set(filename)
            if not self.output_dir_user_selected:
                self.output_dir.set(str(default_output_parent_dir(self.current_tool)))

    def _choose_output(self) -> None:
        directory = filedialog.askdirectory(title="选择保存位置")
        if directory:
            self.output_dir_user_selected = True
            self.output_dir.set(directory)

    def _choose_summary(self) -> None:
        title = "选择异动表模板" if self.current_tool == "personnel_change_merge" else "选择已有汇总表"
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _run_current_tool(self) -> None:
        if self.current_tool == "archive_import":
            messagebox.showinfo("暂未实现", "该工具还在开发中。")
            return
        if self.current_tool == "folder_rename":
            self._run_folder_rename()
            return
        if self.current_tool == "personnel_change_merge":
            self._run_personnel_change_merge()
            return
        if self.current_tool == "salary_merge":
            self._run_salary_merge()
            return
        self._run_salary_split()

    def _run_salary_split(self) -> None:
        input_text = self.input_path.get().strip()
        output_text = self.output_dir.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件", "请先选择工资表文件。")
            return
        input_path = Path(input_text)
        if not input_path.exists():
            messagebox.showwarning("文件不存在", "选择的工资表文件不存在，请重新选择。")
            return
        if input_path.suffix.lower() != ".xlsx":
            messagebox.showwarning("格式不支持", "当前工具只支持 .xlsx 工资表。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = make_result_output_dir(output_parent_dir)
        self.run_button.config(state="disabled")
        self._clear_log()
        self._write_log("开始拆分，请稍候...")

        worker = threading.Thread(
            target=self._salary_split_worker,
            args=(input_path, output_dir),
            daemon=True,
        )
        worker.start()

    def _run_salary_merge(self) -> None:
        input_text = self.input_path.get().strip()
        summary_text = self.summary_path.get().strip()
        summary_path = Path(summary_text) if summary_text else None
        output_text = self.output_dir.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件夹", "请先选择工资表文件夹。")
            return
        input_dir = Path(input_text)
        if not input_dir.exists() or not input_dir.is_dir():
            messagebox.showwarning("文件夹不存在", "选择的工资表文件夹不存在，请重新选择。")
            return
        if summary_path is not None and not summary_path.exists():
            messagebox.showwarning("汇总表不存在", "选择的已有汇总表不存在，请重新选择。")
            return
        if summary_path is not None and summary_path.suffix.lower() != ".xlsx":
            messagebox.showwarning("格式不支持", "已有汇总表只支持 .xlsx 文件。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = make_result_output_dir(output_parent_dir)
        self.run_button.config(state="disabled")
        self._clear_log()
        self._write_log("开始合并，请稍候...")

        worker = threading.Thread(
            target=self._salary_merge_worker,
            args=(input_dir, output_dir, summary_path),
            daemon=True,
        )
        worker.start()

    def _run_personnel_change_merge(self) -> None:
        input_text = self.input_path.get().strip()
        template_text = self.summary_path.get().strip()
        template_path = Path(template_text) if template_text else None
        output_text = self.output_dir.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件夹", "请先选择异动表文件夹。")
            return
        input_dir = Path(input_text)
        if not input_dir.exists() or not input_dir.is_dir():
            messagebox.showwarning("文件夹不存在", "选择的异动表文件夹不存在，请重新选择。")
            return
        if template_path is not None and not template_path.exists():
            messagebox.showwarning("模板不存在", "选择的异动表模板不存在，请重新选择。")
            return
        if template_path is not None and template_path.suffix.lower() != ".xlsx":
            messagebox.showwarning("格式不支持", "异动表模板只支持 .xlsx 文件。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = make_result_output_dir(output_parent_dir)
        self.run_button.config(state="disabled")
        self._clear_log()
        self._write_log("开始汇总，请稍候...")

        worker = threading.Thread(
            target=self._personnel_change_merge_worker,
            args=(input_dir, output_dir, template_path),
            daemon=True,
        )
        worker.start()

    def _run_folder_rename(self) -> None:
        input_text = self.input_path.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件夹", "请先选择人员文件夹目录。")
            return
        root_dir = Path(input_text)
        if not root_dir.exists() or not root_dir.is_dir():
            messagebox.showwarning("文件夹不存在", "选择的人员文件夹目录不存在，请重新选择。")
            return

        mode = RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND)
        try:
            preview = rename_person_folders(
                root_dir=root_dir,
                mode=mode,
                text=self.rename_text.get(),
                target_name=self.rename_target_name.get(),
                replacement_name=self.rename_replacement_name.get(),
                dry_run=True,
            )
        except Exception as exc:
            messagebox.showerror("预览失败", str(exc))
            return

        self._clear_log()
        self._write_log("预览结果：")
        self._write_folder_rename_preview(preview)
        if preview.operation_count == 0:
            messagebox.showinfo("没有可改名文件夹", "没有找到需要改名的文件夹，请检查输入内容。")
            return

        message = self._folder_rename_confirm_message(preview)
        if not messagebox.askyesno("确认改名", message):
            self._write_log("已取消执行。")
            return

        self.run_button.config(state="disabled")
        self._write_log("开始执行改名...")
        worker = threading.Thread(
            target=self._folder_rename_worker,
            args=(root_dir, mode, self.rename_text.get(), self.rename_target_name.get(), self.rename_replacement_name.get()),
            daemon=True,
        )
        worker.start()

    def _salary_split_worker(self, input_path: Path, output_dir: Path) -> None:
        try:
            result = split_salary_by_company(input_path, output_dir)
        except Exception as exc:
            self.status_queue.put(("error", exc))
            return
        self.status_queue.put(("success", result))

    def _salary_merge_worker(self, input_dir: Path, output_dir: Path, summary_path: Path | None) -> None:
        try:
            result = merge_monthly_salary(input_dir, output_dir, existing_summary_path=summary_path)
        except Exception as exc:
            self.status_queue.put(("error", exc))
            return
        self.status_queue.put(("success", result))

    def _personnel_change_merge_worker(self, input_dir: Path, output_dir: Path, template_path: Path | None) -> None:
        try:
            result = merge_personnel_changes(input_dir, output_dir, template_path=template_path)
        except Exception as exc:
            self.status_queue.put(("error", exc))
            return
        self.status_queue.put(("success", result))

    def _folder_rename_worker(
        self,
        root_dir: Path,
        mode: str,
        text: str,
        target_name: str,
        replacement_name: str,
    ) -> None:
        try:
            result = rename_person_folders(
                root_dir=root_dir,
                mode=mode,
                text=text,
                target_name=target_name,
                replacement_name=replacement_name,
            )
        except Exception as exc:
            self.status_queue.put(("error", exc))
            return
        self.status_queue.put(("success", result))

    def _poll_status_queue(self) -> None:
        try:
            while True:
                status, payload = self.status_queue.get_nowait()
                if status == "success":
                    self._handle_success(payload)
                elif status == "error":
                    self._handle_error(payload)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_status_queue)

    def _handle_success(self, result) -> None:
        payload = result.to_dict()
        if self.current_tool == "folder_rename":
            self.last_output_dir = Path(payload["root_dir"])
            self._write_log("改名完成。")
            self._write_folder_rename_preview(result)
            message = f"已完成 {payload['operation_count']} 个文件夹改名。"
        else:
            self.last_output_dir = Path(payload["output_dir"])

            if self.current_tool == "salary_merge":
                self._write_log("合并完成。")
                if payload.get("existing_summary_path"):
                    self._write_log(f"已有汇总表：{payload['existing_summary_path']}")
                self._write_log(f"识别文件数：{payload['source_file_count']}")
                self._write_log(f"识别人员数：{payload['employee_count']}")
                self._write_log(f"工资记录数：{payload['record_count']}")
                self._write_log(f"本次写入记录数：{payload['applied_record_count']}")
                self._write_log(f"已存在未覆盖记录数：{payload['skipped_record_count']}")
                self._write_log(f"输出：{payload['output_file']}")
                for warning in payload["warnings"]:
                    self._write_log(f"提醒：{warning}")
                message = "工资表已合并完成，可以打开结果文件夹查看。"
            elif self.current_tool == "personnel_change_merge":
                self._write_log("汇总完成。")
                self._write_log(f"识别文件数：{payload['source_file_count']}")
                self._write_log(f"异动记录数：{payload['record_count']}")
                for sheet_name, count in payload["sheet_counts"].items():
                    self._write_log(f"- {sheet_name}：{count} 条")
                self._write_log(f"输出：{payload['output_file']}")
                for warning in payload["warnings"]:
                    self._write_log(f"提醒：{warning}")
                message = "异动表已汇总完成，可以打开结果文件夹查看。"
            elif self.current_tool == "salary_split":
                self._write_log("拆分完成。")
                self._write_log(f"识别公司数：{payload['company_count']}")
                self._write_log(f"识别人员数：{payload['employee_count']}")
                for item in payload["outputs"]:
                    self._write_log(f"- {item['company']}：{item['employee_count']} 人")
                    if item.get("file_path"):
                        self._write_log(f"  输出：{item['file_path']}")
                message = "工资表已拆分完成，可以打开结果文件夹查看。"
            else:
                message = "处理完成。"
        self.run_button.config(state="normal")
        messagebox.showinfo("处理完成", message)

    def _handle_error(self, exc: object | None) -> None:
        action = (
            "合并"
            if self.current_tool == "salary_merge"
            else "汇总"
            if self.current_tool == "personnel_change_merge"
            else "改名"
            if self.current_tool == "folder_rename"
            else "拆分"
        )
        self._write_log(f"{action}失败。")
        self._write_log(str(exc))
        self.run_button.config(state="normal")
        messagebox.showerror(f"{action}失败", str(exc))

    def _write_folder_rename_preview(self, result) -> None:
        payload = result.to_dict()
        self._write_log(f"目录：{payload['root_dir']}")
        self._write_log(f"数量：{payload['operation_count']}")
        for operation in payload["operations"][:30]:
            self._write_log(f"- {operation['source_name']} -> {operation['target_name']}")
        remaining = payload["operation_count"] - 30
        if remaining > 0:
            self._write_log(f"... 还有 {remaining} 条")
        for warning in payload["warnings"]:
            self._write_log(f"提醒：{warning}")

    def _folder_rename_confirm_message(self, result) -> str:
        payload = result.to_dict()
        lines = [f"确认改名 {payload['operation_count']} 个文件夹："]
        for operation in payload["operations"][:8]:
            lines.append(f"{operation['source_name']} -> {operation['target_name']}")
        remaining = payload["operation_count"] - 8
        if remaining > 0:
            lines.append(f"... 还有 {remaining} 条")
        lines.append("")
        lines.append("确认后会直接改名，是否继续？")
        return "\n".join(lines)

    def _open_output_dir(self) -> None:
        directory_text = self.output_dir.get().strip()
        directory = self.last_output_dir
        if self.current_tool == "folder_rename" and directory is None:
            input_text = self.input_path.get().strip()
            directory = Path(input_text) if input_text else None
            if directory is None:
                messagebox.showwarning("缺少目录", "请先选择人员文件夹目录。")
                return
        if directory is None and directory_text:
            directory = make_result_output_dir(Path(directory_text))
        if directory is None:
            messagebox.showwarning("缺少目录", "请先选择保存位置。")
            return
        if self.current_tool == "folder_rename" and not directory.exists():
            messagebox.showwarning("目录不存在", "选择的人员文件夹目录不存在。")
            return
        try:
            if self.current_tool != "folder_rename":
                directory.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("无法创建目录", str(exc))
            return
        open_path(directory)

    def _write_log(self, text: str) -> None:
        tag = None
        if any(keyword in text for keyword in ("失败", "错误", "提醒", "不存在", "缺少")):
            tag = "warning"
        elif any(keyword in text for keyword in ("完成", "成功")):
            tag = "success"
        elif text.startswith("- "):
            tag = "muted"
        if tag:
            self.log_text.insert(END, text + "\n", tag)
        else:
            self.log_text.insert(END, text + "\n")
        self.log_text.see(END)

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", END)


def _default_result_dir_name() -> str:
    return "结果_" + datetime.now().strftime("%Y%m%d_%H%M%S")


def default_output_parent_dir(tool: str) -> Path:
    if tool == "salary_merge":
        folder_name = "工资合并结果"
    elif tool == "personnel_change_merge":
        folder_name = "异动表汇总结果"
    else:
        folder_name = "工资表拆分结果"
    return desktop_dir() / folder_name


def make_result_output_dir(parent_dir: Path) -> Path:
    return parent_dir / _default_result_dir_name()


def desktop_dir() -> Path:
    home = Path.home()
    desktop = home / "Desktop"
    if desktop.exists():
        return desktop
    return home / "桌面"


def open_path(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def main() -> None:
    root = Tk()
    HRToolkitApp(root)
    root.mainloop()
