from __future__ import annotations

import calendar
import errno
import importlib
import inspect
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from tkinter import BOTH, BOTTOM, END, LEFT, RIGHT, TOP, VERTICAL, Y, Canvas, Menu, StringVar, BooleanVar, Text, Tk, TkVersion
from tkinter import font as tkfont
from tkinter import ttk

from hr_toolkit import __version__, runlog
from hr_toolkit.app_update import (
    UpdateCancelledError,
    UpdateInfo,
    check_for_update,
    cleanup_stale_update_files,
    download_update_package,
    launch_update_replacement,
    resolve_download_url,
    update_check_enabled,
)
from hr_toolkit.common.inputs import (
    ARCHIVE_FILE_DIALOG_PATTERN,
    ARCHIVE_FORMAT_DESCRIPTION,
    archive_suffix,
    is_supported_archive_file,
)
from hr_toolkit.history_store import (
    HISTORY_PAGE_SIZE,
    HistoryStore,
    HistoryStoreError,
    SourceSpec,
    TaskDetail,
)
from hr_toolkit.material_preferences import MaterialPreferences
from hr_toolkit.tools.folder_rename import (
    MODE_APPEND,
    MODE_EXCEL_BATCH,
    MODE_REMOVE,
    MODE_REPLACE,
    rename_files_by_excel,
    rename_person_folders,
    FILE_TYPE_FOLDER,
    FILE_TYPE_ALL,
    FILE_TYPE_PDF,
    FILE_TYPE_IMAGE,
    FILE_TYPE_DOCUMENT,
)
from hr_toolkit.tools.archive_import import export_company_archive_tables, import_archive_transfers
from hr_toolkit.tools.data_statistics import generate_data_statistics_reports, resolve_month_range, resolve_week_range
from hr_toolkit.tools.personnel_change_merge import merge_personnel_changes, update_roster_from_change_summaries
from hr_toolkit.tools.insurance_ledger import generate_insurance_ledger
from hr_toolkit.tools.material_collector import (
    LIBRARY_MODE_FLAT_OCR,
    LIBRARY_MODE_LABELS,
    LIBRARY_MODE_PERSON_FOLDER,
    MODE_BY_EMPLOYEE,
    MODE_BY_MATERIAL,
    MODE_FLAT,
    MODE_LABELS,
    MATERIAL_SYNONYMS,
    collect_employee_materials,
)
from hr_toolkit.tools.salary_merge import merge_monthly_salary
from hr_toolkit.tools.salary_split import split_salary_by_company
from hr_toolkit.tools.social_security import generate_social_security_reports

from .constants import (
    COLOR_BADGE_DIR_BG,
    COLOR_BADGE_DIR_FG,
    COLOR_BADGE_XLS_BG,
    COLOR_BADGE_XLS_FG,
    COLOR_BADGE_ZIP_BG,
    COLOR_BADGE_ZIP_FG,
    COLOR_BG,
    COLOR_BORDER,
    COLOR_BORDER_FAINT,
    COLOR_DANGER,
    COLOR_DISABLED,
    COLOR_DROP_BG,
    COLOR_DROP_BORDER,
    COLOR_FAINT,
    COLOR_LOG_BG,
    COLOR_LOG_MUTED,
    COLOR_LOG_TEXT,
    COLOR_MUTED,
    COLOR_NAV_HOVER,
    COLOR_NAV_SELECTED,
    COLOR_NAV_TEXT,
    COLOR_NAV_TEXT_SELECTED,
    COLOR_PRIMARY,
    COLOR_PRIMARY_ACTIVE,
    COLOR_PRIMARY_SOFT,
    COLOR_SIDEBAR,
    COLOR_SIDEBAR_BORDER,
    COLOR_SUCCESS,
    COLOR_SUCCESS_DOT,
    COLOR_SURFACE,
    COLOR_SURFACE_ALT,
    COLOR_SURFACE_PRESSED,
    COLOR_TEXT,
    COLOR_WARNING,
    COLOR_WARNING_DOT,
    COLOR_WARNING_SOFT,
    HISTORY_DATE_FILTER_ALL,
    HISTORY_DATE_FILTERS,
    HISTORY_PATH_ARGUMENTS,
    HISTORY_PRIMARY_PATH_ARGUMENTS,
    HISTORY_STATUS_LABELS,
    HISTORY_SUPPORTING_PATH_ARGUMENTS,
    HISTORY_TOOL_FILTER_ALL,
    MULTI_INPUT_TOOLS,
    NAV_GROUPS,
    RENAME_FILE_TYPE_LABELS,
    RENAME_FILE_TYPE_LABELS_REVERSE,
    RENAME_MODE_LABELS,
    TOOL_GROUP_LABELS,
    TOOL_LOG_LABELS,
    TOOL_NAV_ITEMS,
    TOOL_NAV_LABELS,
    UPDATE_DIALOG_BG,
    UPDATE_DIALOG_ICON_BG,
    UPDATE_DIALOG_MUTED,
    UPDATE_DIALOG_NOTES_BG,
    UPDATE_DIALOG_PRIMARY,
    UPDATE_DIALOG_PRIMARY_ACTIVE,
    UPDATE_DIALOG_SECONDARY,
    UPDATE_DIALOG_SECONDARY_ACTIVE,
    UPDATE_DIALOG_TEXT,
    UPDATE_DIALOG_TRACK,
    WORKSPACE_COLLAPSED_WIDTH,
    WORKSPACE_DEFAULT_WIDTH,
    WORKSPACE_DRAWER_BREAKPOINT,
    WORKSPACE_DUMMY_TAG,
    WORKSPACE_HIDDEN_NAMES,
    WORKSPACE_HIDDEN_SUFFIXES,
    WORKSPACE_MAX_WIDTH,
    WORKSPACE_MIN_WIDTH,
    WORKSPACE_SCOPE_ALL,
    WORKSPACE_SCOPE_TOOL,
    WORKSPACE_SEARCH_LIMIT,
    WORKSPACE_TOOL_PATHS,
    APP_DISPLAY_NAME,
    APP_SUBTITLE,
)
from .scaling import (
    LAYOUT_MODE_NARROW,
    LAYOUT_MODE_WIDE,
    _clamp_ui_scale,
    _configure_tk_font_scaling,
    _detect_ui_scale,
    _fit_window_size,
    _font_size,
    _forced_ui_scale,
    _indeterminate_progress_segment,
    _responsive_checkbox_columns,
    _responsive_drawer_width,
    _responsive_layout_mode,
    _scale_float,
    _scale_px,
    _widget_ui_scale,
    _windows_dpi_for_root,
    _windows_work_area_for_root,
)
from .widgets import (
    CodexButton,
    RoundedCard,
    SidebarItem,
    _paint_tool_icon,
    _paint_codex_badge_icon,
    _get_default_font_family,
)
from .helpers import (
    _default_result_dir_name,
    _default_workspace_project_name,
    _enable_high_dpi_rendering,
    _install_crash_logging,
    _set_windows_app_identity,
    _workspace_project_create_error_message,
    _workspace_project_creation_target,
    _workspace_project_name_error,
    _workspace_trash_deleted_text,
    _workspace_trash_dialog_height,
    _workspace_trash_group_tool,
    _workspace_trash_ignore_enter,
    _workspace_trash_matches,
    _workspace_trash_period_label,
    _workspace_trash_restore_location,
    _workspace_trash_title,
    default_output_parent_dir,
    desktop_dir,
    make_result_output_dir,
    open_path,
)
from .task_runner import TaskRunner, TaskToken


EXCEL_ARCHIVE_FILE_DIALOG_PATTERN = f"*.xlsx *.xls {ARCHIVE_FILE_DIALOG_PATTERN}"
EXCEL_ARCHIVE_FILETYPES = (
    ("Excel 或压缩包", EXCEL_ARCHIVE_FILE_DIALOG_PATTERN),
    ("Excel 工作簿", "*.xlsx *.xls"),
    ("常见压缩包", ARCHIVE_FILE_DIALOG_PATTERN),
    ("所有文件", "*.*"),
)
EXCEL_ARCHIVE_FORMAT_TEXT = f".xlsx、.xls 以及 {ARCHIVE_FORMAT_DESCRIPTION} 压缩包"


def _is_excel_or_archive_file(path: Path) -> bool:
    return path.suffix.lower() in {".xlsx", ".xls"} or is_supported_archive_file(path)


class _DynamicProxy:
    def __init__(self, name: str):
        self._name = name

    def __call__(self, *args, **kwargs):
        target = getattr(sys.modules.get('hr_toolkit.gui') or sys.modules[__name__], self._name)
        return target(*args, **kwargs)

    def __getattr__(self, attr: str):
        target = getattr(sys.modules.get('hr_toolkit.gui') or sys.modules[__name__], self._name)
        return getattr(target, attr)

Frame = _DynamicProxy('Frame')
Label = _DynamicProxy('Label')
PhotoImage = _DynamicProxy('PhotoImage')
Toplevel = _DynamicProxy('Toplevel')
filedialog = _DynamicProxy('filedialog')
messagebox = _DynamicProxy('messagebox')
simpledialog = _DynamicProxy('simpledialog')
make_result_output_dir = _DynamicProxy('make_result_output_dir')
runlog = _DynamicProxy('runlog')
_default_result_dir_name = _DynamicProxy('_default_result_dir_name')

class HRToolkitApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.ui_scale = _detect_ui_scale(root)
        setattr(self.root, "_hr_ui_scale", self.ui_scale)
        # Windows 依赖 tk scaling 把“点”字号按 DPI 转像素；macOS 的字体
        # 差异由 _font_size() 按 Tk 版本处理，强制缩放仍需配置 Tk。
        if sys.platform.startswith("win") or _forced_ui_scale() is not None:
            _configure_tk_font_scaling(self.root, self.ui_scale)

        self.root.title(f"{APP_DISPLAY_NAME} v{__version__}")
        initial_width, initial_height = self._window_size(1400, 780)
        min_width, min_height = self._window_size(900, 600)
        self.root.geometry(f"{initial_width}x{initial_height}")
        self.root.minsize(min_width, min_height)
        self.root.configure(bg=COLOR_BG)
        self._loading_overlay = None
        self._startup_loading_timer = None
        self._setup_startup_loading_screen()
        # Window starts hidden and fully transparent so the synchronous widget-creation
        # storm (~70 Canvas widgets on Windows GDI) happens completely off-screen. The
        # user only sees the final, fully-laid-out state when we deiconify and fade back in
        # at the end of __init__.
        self._was_withdrawn = False
        try:
            if self.root.state() == "withdrawn":
                self._was_withdrawn = True
            else:
                self.root.withdraw()
                self._was_withdrawn = True
        except Exception:
            pass

        self._window_faded_out = False
        if sys.platform.startswith("win") or sys.platform == "darwin":
            try:
                self.root.attributes("-alpha", 0.0)
                self._window_faded_out = True
            except Exception:
                self._window_faded_out = False

        self.current_tool = "social_security"
        self.current_view = "tool"
        self.nav_buttons: dict[str, SidebarItem] = {}
        self.tool_title = StringVar()
        self.tool_description = StringVar()
        self.tool_group = StringVar()
        self.input_label = StringVar()
        self.input_hint = StringVar()
        self.choose_input_text = StringVar()
        self.run_button_text = StringVar()
        self.summary_label = StringVar()
        self.summary_button_text = StringVar()
        self.last_run_text = StringVar()
        self.last_run_state = StringVar()
        self.history_search = StringVar()
        self.history_tool_filter = StringVar(value=HISTORY_TOOL_FILTER_ALL)
        self.history_date_filter = StringVar(value=HISTORY_DATE_FILTER_ALL)
        self.history_detail_title = StringVar(value="选择一条记录查看详情")
        self.history_detail_text = StringVar(value="这里用于查看升级前由旧版本保存的处理记录。")
        self.history_page_text = StringVar(value="第 1 页")
        self.history_message = StringVar()
        self.workspace_project_name = StringVar(value="未打开工作项目")
        self.workspace_project_path = StringVar(value="新建或打开项目后，资料会显示在这里")
        self.sidebar_project_summary = StringVar(value="新建或打开项目后开始处理")
        self.workspace_search = StringVar()
        self.workspace_scope = StringVar(value=WORKSPACE_SCOPE_ALL)
        self.workspace_detail_title = StringVar(value="选择文件查看详情")
        self.workspace_detail_text = StringVar(value="项目内的上传资料和处理结果会集中显示在这里。")
        self.workspace_empty_text = StringVar(value="先新建或打开一个工作项目")
        self.current_project_path: Path | None = None
        self._workspace_project_read_only = False
        self._workspace_tree_paths: dict[str, Path] = {}
        self._workspace_search_job: str | None = None
        self._workspace_search_generation = 0
        self._workspace_project_generation = 0
        self._workspace_queue: queue.Queue[tuple[str, int, object]] = queue.Queue()
        self._workspace_write_token = 0
        self._workspace_write_tasks: dict[int, tuple[threading.Event, object]] = {}
        # 复制进度由后台线程覆盖“最新快照”，主线程定时读取；不把每个文件块
        # 都塞进 Tk 队列，避免大文件产生数万条待处理消息。
        self._workspace_write_progress: dict[int, object] = {}
        self._workspace_write_progress_lock = threading.Lock()
        self._workspace_write_callbacks: dict[int, tuple[object | None, object | None, object | None]] = {}
        self._workspace_recovery_blocked = False
        self._workspace_recovery_error: str | None = None
        self._workspace_close_requested = False
        self._workspace_recent_projects: list[tuple[str, Path]] = []
        self._workspace_width_units = WORKSPACE_DEFAULT_WIDTH
        self._workspace_preferred_expanded = True
        self._workspace_small = False
        self._workspace_drawer_open = False
        self._workspace_restore_focus = None
        self._workspace_panel_was_temporary_open = False
        self._workspace_resize_origin: tuple[int, int] | None = None
        self._project_store_error: str | None = None
        self.project_store = self._initialize_project_store()
        self._load_workspace_preferences()
        # 每个工具（含子模式）本次会话内的最近一次运行结果：(时间, 成功/失败)
        self._last_run_results: dict[str, tuple[str, str]] = {}
        # 合并后的上传入口：当前工具可用的文件/文件夹选择动作与提示文案
        self._input_file_cmd = None
        self._input_folder_cmd = None
        self._input_drop_title = ""
        self._input_allow_multi = True
        self._tutorial_window: Toplevel | None = None
        self._project_create_window: Toplevel | None = None
        self._project_create_busy = False
        self._project_create_name_var: StringVar | None = None
        self._project_create_parent_var: StringVar | None = None
        self._project_create_preview_var: StringVar | None = None
        self._project_create_status_var: StringVar | None = None
        self._project_create_trace_ids: list[tuple[StringVar, str]] = []
        self._project_create_name_entry = None
        self._project_create_location_button: CodexButton | None = None
        self._project_create_cancel_button: CodexButton | None = None
        self._project_create_submit_button: CodexButton | None = None
        self._project_create_status_label: Label | None = None
        self._workspace_trash_window: Toplevel | None = None
        self._workspace_trash_search_var: StringVar | None = None
        self._workspace_trash_status_var: StringVar | None = None
        self._workspace_trash_restore_path_var: StringVar | None = None
        self._workspace_trash_project_var: StringVar | None = None
        self._workspace_trash_notice_var: StringVar | None = None
        self._workspace_trash_search_trace: str | None = None
        self._workspace_trash_details: tuple[object, ...] = ()
        self._workspace_trash_selected_id: str | None = None
        self._workspace_trash_restore_in_progress = False
        self._workspace_trash_card_widgets: dict[str, tuple[Frame, tuple[object, ...]]] = {}
        self._workspace_trash_list_body: Frame | None = None
        self._workspace_trash_empty_label: Label | None = None
        self._workspace_trash_detail_title: Label | None = None
        self._workspace_trash_restore_button: CodexButton | None = None
        self._workspace_import_window: Toplevel | None = None
        self._workspace_import_token: int | None = None
        self._workspace_import_started_at: float | None = None
        self._workspace_import_phase = "checking"
        self._workspace_import_animation_job: str | None = None
        self._workspace_import_success_job: str | None = None
        self._workspace_import_animation_offset = 0.0
        self._workspace_import_progress_canvas: Canvas | None = None
        self._workspace_import_progress_width = self._px(548)
        self._workspace_import_title_var: StringVar | None = None
        self._workspace_import_subtitle_var: StringVar | None = None
        self._workspace_import_target_var: StringVar | None = None
        self._workspace_import_state_var: StringVar | None = None
        self._workspace_import_name_var: StringVar | None = None
        self._workspace_import_left_var: StringVar | None = None
        self._workspace_import_middle_var: StringVar | None = None
        self._workspace_import_elapsed_var: StringVar | None = None
        self._workspace_import_safety_title_var: StringVar | None = None
        self._workspace_import_safety_text_var: StringVar | None = None
        self._workspace_import_stage_labels: list[Label] = []
        self._workspace_import_cancel_button: CodexButton | None = None
        self.change_mode = "merge"
        self.change_form_state: dict[str, tuple[str, str, list[Path] | None]] = {
            "merge": ("", "", None),
            "roster": ("", "", None),
        }
        self.archive_mode = "import"
        self.archive_form_state: dict[str, tuple[str, str, list[Path] | None]] = {
            "import": ("", "", None),
            "export": ("", "", None),
        }
        self.rename_mode = StringVar(value="追加文字")
        self.rename_target_label = StringVar(value="姓名（可不填）")
        self.rename_text_label = StringVar(value="要追加的文字")
        self.rename_replacement_label = StringVar(value="新名称")
        self.rename_target_name = StringVar()
        self.rename_text = StringVar()
        self.rename_replacement_name = StringVar()
        self.rename_file_type = StringVar(value="文件夹")
        self.material_mode = StringVar(value="按员工归类（每人一个文件夹）")
        self.material_library_mode = StringVar(value="按人员文件夹查找（原模式）")
        self.material_create_zip = BooleanVar(value=False)
        self.material_collect_all = BooleanVar(value=True)
        # OCR 智能索引缓存：默认开启，二次扫描可秒级命中
        self.material_use_ocr_cache = BooleanVar(value=True)
        self.material_target_input = StringVar(value="")
        self.material_types_selected: dict[str, BooleanVar] = {
            material: BooleanVar(value=True)
            for material in self._material_preferences.available_materials
        }
        initial_material_preset = (
            self._material_preferences.preset_names[0]
            if self._material_preferences.preset_names
            else ""
        )
        self.material_preset_name = StringVar(value=initial_material_preset)
        initial_custom_material = (
            self._material_preferences.custom_materials[0]
            if self._material_preferences.custom_materials
            else ""
        )
        self.material_custom_choice = StringVar(value=initial_custom_material)
        self.input_path = StringVar()
        self.summary_path = StringVar()
        self.stats_week_start = StringVar()
        self.stats_week_end = StringVar()
        self.stats_month_start = StringVar()
        self.stats_month_end = StringVar()
        # 考勤统计表备注中加班/调休的展示单位：day 按天（默认）/ hour 按小时
        self.stats_remark_unit = StringVar(value="day")
        # 是否在考勤统计表新增「公出」列：默认否（不加列）
        self.stats_include_business_trip = BooleanVar(value=False)
        # 是否新增「出差」列：只统计源表中的工作日出差，不含休息日出差天数
        self.stats_include_workday_business_trip = BooleanVar(value=False)
        # 正式结果只写入当前工作项目。这个变量保留给现有工具表单，实际运行时
        # 会被替换为本次批次的“处理结果”目录。
        self.output_dir = StringVar(value="")
        self.output_display_path = StringVar(value="请先新建或打开工作项目")
        self.output_dir_user_selected = True
        self.change_input_paths: list[Path] | None = None
        # (状态, 运行编号, 载荷)；运行编号用于丢弃已停止任务的结果
        self.status_queue: queue.Queue[tuple[str, int, object | None]] = queue.Queue()
        self.update_queue: queue.Queue[tuple[str, object | None]] = queue.Queue()
        self.history_queue: queue.Queue[tuple[str, object | None]] = queue.Queue()
        self.last_output_dir: Path | None = None
        self.pending_update: UpdateInfo | None = None
        self.update_window: Toplevel | None = None
        self.update_progress_label: Label | None = None
        self.update_progress_canvas: Canvas | None = None
        self.update_progress_width = self._px(248)
        self.update_progress_job: str | None = None
        self.update_progress_phase = 0.0
        self.update_progress_last_tick: float | None = None
        self.update_check_in_progress = False
        self.manual_update_check_active = False
        self.update_check_dismissed = False
        self._download_speed_anchor: tuple[float, int] | None = None
        self._update_download_cancel_event: threading.Event | None = None
        self._tool_run_token = 0
        self._tool_running = False
        self._idle_run_button_text = ""
        self._history_page = 0
        self._history_total = 0
        self._history_selected_task: TaskDetail | None = None
        self._history_task_by_token: dict[int, str] = {}
        self._project_batch_by_token: dict[int, str] = {}
        self._run_cancel_events: dict[int, threading.Event] = {}
        self._history_init_error: str | None = None
        try:
            self.history_store: HistoryStore | None = HistoryStore()
        except Exception as exc:
            self.history_store = None
            self._history_init_error = str(exc)
            runlog.log_exception("历史资料库初始化失败", exc)

        self._is_alive = True
        self._poll_update_timer = None
        self._startup_check_timer = None

        try:
            tk_scaling = float(self.root.tk.call("tk", "scaling"))
        except Exception:
            tk_scaling = 0.0
        runlog.log_line(
            f"{APP_DISPLAY_NAME} v{__version__} 启动（{sys.platform}，"
            f"分辨率 {self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}，"
            f"ui_scale {self.ui_scale:.2f}，tk scaling {tk_scaling:.2f}）"
        )
        # 界面回调里的异常默认只打到不存在的控制台，改为写入运行日志
        self.root.report_callback_exception = self._on_tk_callback_exception

        self._apply_app_icon()
        self._configure_style()
        self._set_tool_texts()
        self._build_layout()
        if self._loading_overlay is not None:
            try:
                self._loading_overlay.lift()
            except Exception:
                pass
        self._update_project_output_controls()
        self.root.protocol("WM_DELETE_WINDOW", self._request_close)
        self._poll_status_queue()
        self._poll_update_queue()
        self._poll_history_queue()
        self._poll_workspace_queue()
        # Force one synchronous layout pass while the window is still
        # covered by the loading overlay — every Canvas subclass needs at least one ``update``
        # cycle to realize its backing GDI surface on Windows.
        self.root.update_idletasks()
        if self._window_faded_out:
            try:
                self.root.attributes("-alpha", 1.0)
            except Exception:
                pass
            self._window_faded_out = False
        if self._was_withdrawn:
            try:
                self.root.deiconify()
            except Exception:
                pass
            self._was_withdrawn = False
        if self._loading_overlay is not None:
            try:
                self._loading_overlay.lift()
            except Exception:
                pass
        self.root.after_idle(self._restore_workspace_project)
        self._startup_check_timer = self.root.after(600, self._check_updates_on_startup)
        self._startup_loading_timer = self.root.after(400, self._dismiss_startup_loading_screen)
        # 清理历史更新遗留的临时文件（下载包、解压目录），后台低优先执行
        threading.Thread(target=cleanup_stale_update_files, daemon=True).start()

    def _on_tk_callback_exception(self, exc_type, exc_value, exc_tb) -> None:
        runlog.log_exception("界面异常", exc_value, exc_tb)
        import traceback as traceback_module

        traceback_module.print_exception(exc_type, exc_value, exc_tb)

    def _initialize_project_store(self):
        """Load the optional project backend without making GUI startup depend on it."""
        self._project_store_class = None
        try:
            module = importlib.import_module("hr_toolkit.project_store")
        except ModuleNotFoundError as exc:
            if exc.name == "hr_toolkit.project_store":
                return None
            self._project_store_error = str(exc)
            runlog.log_exception("工作项目模块加载失败", exc)
            return None
        except Exception as exc:
            self._project_store_error = str(exc)
            runlog.log_exception("工作项目模块加载失败", exc)
            return None
        store_class = getattr(module, "ProjectStore", None)
        if store_class is None:
            self._project_store_error = "未找到 ProjectStore。"
            return None
        self._project_store_class = store_class
        return None

    @staticmethod
    def _workspace_settings_path() -> Path:
        if sys.platform.startswith("win"):
            base = Path(os.environ.get("LOCALAPPDATA", "").strip() or (Path.home() / "AppData" / "Local"))
        elif sys.platform == "darwin":
            base = Path.home() / "Library" / "Application Support"
        else:
            base = Path(os.environ.get("XDG_CONFIG_HOME", "").strip() or (Path.home() / ".config"))
        return base / "HRToolkit" / "workspace-ui.json"

    def _load_workspace_preferences(self) -> None:
        self._workspace_last_project_path: Path | None = None
        self._last_selected_dir: Path | None = None
        self._material_preferences = MaterialPreferences()
        settings_path = self._workspace_settings_path()
        try:
            if settings_path.is_symlink() or not settings_path.is_file():
                return
            state = json.loads(settings_path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
            runlog.log_exception("读取项目界面设置失败", exc)
            return
        if not isinstance(state, dict):
            return
        try:
            width = int(state.get("project_panel_width", WORKSPACE_DEFAULT_WIDTH))
        except (AttributeError, TypeError, ValueError):
            width = WORKSPACE_DEFAULT_WIDTH
        self._workspace_width_units = max(WORKSPACE_MIN_WIDTH, min(WORKSPACE_MAX_WIDTH, width))
        try:
            self._workspace_preferred_expanded = bool(state.get("project_panel_expanded", True))
        except AttributeError:
            self._workspace_preferred_expanded = True
        recent: list[tuple[str, Path]] = []
        raw_recent = state.get("recent_projects", [])
        if isinstance(raw_recent, list):
            for raw_path in raw_recent:
                if not isinstance(raw_path, str) or not raw_path.strip():
                    continue
                path = Path(raw_path).expanduser().absolute()
                if any(existing_path == path for _name, existing_path in recent):
                    continue
                recent.append((path.name or "工作项目", path))
                if len(recent) >= 8:
                    break
        self._workspace_recent_projects = recent
        raw_current = state.get("current_project")
        if isinstance(raw_current, str) and raw_current.strip():
            self._workspace_last_project_path = Path(raw_current).expanduser().absolute()
        raw_last_dir = state.get("last_selected_dir")
        if isinstance(raw_last_dir, str) and raw_last_dir.strip():
            try:
                candidate_dir = Path(raw_last_dir).expanduser().absolute()
                if candidate_dir.is_dir():
                    self._last_selected_dir = candidate_dir
            except Exception:
                pass
        self._material_preferences = MaterialPreferences.from_payload(
            state.get("material_preferences")
        )

    def _save_workspace_preferences(self) -> None:
        settings_path = self._workspace_settings_path()
        material_preferences = getattr(self, "_material_preferences", MaterialPreferences())
        payload = {
            "version": 1,
            "project_panel_width": int(self._workspace_width_units),
            "project_panel_expanded": bool(self._workspace_preferred_expanded),
            "current_project": str(self.current_project_path) if self.current_project_path is not None else None,
            "recent_projects": [str(path) for _name, path in self._workspace_recent_projects[:8]],
            "last_selected_dir": str(self._last_selected_dir) if getattr(self, "_last_selected_dir", None) is not None else None,
            "material_preferences": material_preferences.to_payload(),
        }
        temp_path = settings_path.with_name(
            f".{settings_path.name}.{os.getpid()}.{time.time_ns()}.tmp"
        )
        try:
            settings_path.parent.mkdir(parents=True, exist_ok=True)
            if settings_path.parent.is_symlink() or settings_path.is_symlink():
                raise OSError("设置目录不是安全的普通目录。")
            # 独占创建临时文件，避免可预测文件名被链接替换后写到项目外。
            with temp_path.open("x", encoding="utf-8") as temp_file:
                temp_file.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
            try:
                temp_path.chmod(0o600)
            except OSError:
                pass
            temp_path.replace(settings_path)
        except Exception as exc:
            runlog.log_exception("保存项目界面设置失败", exc)
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass

    def _file_dialog_initial_dir(self) -> str:
        last_dir = getattr(self, "_last_selected_dir", None)
        if last_dir is not None:
            try:
                cand = Path(last_dir).expanduser().absolute()
                if cand.is_dir():
                    return str(cand)
            except Exception:
                pass
        project_path = getattr(self, "current_project_path", None)
        if project_path is not None:
            try:
                cand = Path(project_path).expanduser().absolute()
                if cand.is_dir():
                    return str(cand)
            except Exception:
                pass
        try:
            fallback = desktop_dir() or Path.home()
            if fallback and Path(fallback).is_dir():
                return str(Path(fallback).expanduser().absolute())
        except Exception:
            pass
        return str(Path.home())

    def _remember_file_dialog_path(self, selected: str | Path | list[str | Path] | tuple[str | Path, ...] | None) -> None:
        if not selected:
            return
        item = selected[0] if isinstance(selected, (list, tuple)) else selected
        if not item:
            return
        try:
            path = Path(item).expanduser().absolute()
            folder = path if path.is_dir() else path.parent
            if folder.is_dir():
                self._last_selected_dir = folder
                self._save_workspace_preferences()
        except Exception:
            pass

    def _askopenfilename(self, **kwargs) -> str:
        if "initialdir" not in kwargs:
            kwargs["initialdir"] = self._file_dialog_initial_dir()
        if "parent" not in kwargs and getattr(self, "root", None) is not None:
            kwargs["parent"] = self.root
        result = filedialog.askopenfilename(**kwargs)
        if result:
            self._remember_file_dialog_path(result)
        return result

    def _askopenfilenames(self, **kwargs) -> tuple[str, ...]:
        if "initialdir" not in kwargs:
            kwargs["initialdir"] = self._file_dialog_initial_dir()
        if "parent" not in kwargs and getattr(self, "root", None) is not None:
            kwargs["parent"] = self.root
        result = filedialog.askopenfilenames(**kwargs)
        if result:
            self._remember_file_dialog_path(result)
        return result

    def _askdirectory(self, **kwargs) -> str:
        if "initialdir" not in kwargs:
            kwargs["initialdir"] = self._file_dialog_initial_dir()
        if "parent" not in kwargs and getattr(self, "root", None) is not None:
            kwargs["parent"] = self.root
        result = filedialog.askdirectory(**kwargs)
        if result:
            self._remember_file_dialog_path(result)
        return result

    def _apply_app_icon(self) -> None:
        # 替换标题栏/任务栏默认的 Tk 羽毛图标；iconphoto(True, ...) 会同时
        # 应用到之后创建的所有 Toplevel（更新对话框等）
        try:
            from hr_toolkit._icon_data import APP_ICON_PNGS_BASE64

            # 必须从大到小传入：macOS 的 Dock 只用第一张，给小图会被放大成马赛克
            self._app_icon_images = [
                PhotoImage(data=APP_ICON_PNGS_BASE64[size]) for size in sorted(APP_ICON_PNGS_BASE64, reverse=True)
            ]
            self.root.iconphoto(True, *self._app_icon_images)
        except Exception:
            pass

    def _setup_startup_loading_screen(self) -> None:
        """Create the Codex-style startup loading overlay that covers initial widget creation."""
        self._loading_overlay = Frame(self.root, bg=COLOR_BG)
        self._loading_overlay.place(x=0, y=0, relwidth=1.0, relheight=1.0)
        try:
            self._loading_overlay.lift()
        except Exception:
            pass

        center_container = Frame(self._loading_overlay, bg=COLOR_BG)
        center_container.place(relx=0.5, rely=0.5, anchor="center")

        icon_size = self._px(64)
        canvas_padding = self._px(12)
        canvas_dim = icon_size + canvas_padding * 2

        icon_canvas = Canvas(
            center_container,
            width=canvas_dim,
            height=canvas_dim,
            bg=COLOR_BG,
            highlightthickness=0,
            bd=0,
        )
        icon_canvas.pack(side=TOP, pady=(0, self._px(14)))

        _paint_codex_badge_icon(
            icon_canvas,
            canvas_padding,
            canvas_padding,
            icon_size,
            scale=self.ui_scale,
        )

        family = _get_default_font_family(self.root)
        title_font = (family, _font_size(12), "bold")
        title_label = Label(
            center_container,
            text=APP_DISPLAY_NAME,
            font=title_font,
            fg="#4A4845",
            bg=COLOR_BG,
        )
        title_label.pack(side=TOP, pady=(0, self._px(4)))

        sub_font = (family, _font_size(9))
        self._loading_status_label = Label(
            center_container,
            text="正在准备工作区…",
            font=sub_font,
            fg="#8C8A85",
            bg=COLOR_BG,
        )
        self._loading_status_label.pack(side=TOP)

    def _on_startup_rendered(self) -> None:
        """Called once initial project restore and layout passes have settled completely."""
        if not getattr(self, "_is_alive", True):
            return
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        self.root.after(40, self._dismiss_startup_loading_screen)

    def _dismiss_startup_loading_screen(self) -> None:
        """Dismiss the startup loading overlay after all initial UI rendering has settled."""
        overlay = getattr(self, "_loading_overlay", None)
        if overlay is None:
            return
        self._loading_overlay = None
        try:
            overlay.destroy()
        except Exception:
            pass

    def _px(self, value: int | float) -> int:
        return _scale_px(value, self.ui_scale)

    def _pxf(self, value: int | float) -> float:
        return _scale_float(value, self.ui_scale)

    def _pad(self, *values: int | float) -> tuple[int, ...]:
        return tuple(self._px(value) for value in values)

    def _window_size(self, width: int, height: int) -> tuple[int, int]:
        scaled_width = self._px(width)
        scaled_height = self._px(height)
        available_width, available_height, exact_work_area = self._usable_work_area_size()
        if available_width <= 0 or available_height <= 0:
            return scaled_width, scaled_height
        if exact_work_area:
            return _fit_window_size(
                scaled_width,
                scaled_height,
                available_width,
                available_height,
                horizontal_margin=self._px(16),
                vertical_margin=self._px(16),
            )
        return _fit_window_size(
            scaled_width,
            scaled_height,
            available_width,
            available_height,
            horizontal_margin=self._px(48),
            vertical_margin=self._px(72),
        )

    def _usable_work_area_size(self) -> tuple[int, int, bool]:
        work_area = _windows_work_area_for_root(self.root)
        if work_area is not None:
            left, top, right, bottom = work_area
            return max(1, right - left), max(1, bottom - top), True
        try:
            return self.root.winfo_screenwidth(), self.root.winfo_screenheight(), False
        except Exception:
            return 0, 0, False

    def _update_dialog_size(self, width: int, height: int) -> tuple[int, int]:
        scaled_width = self._px(width)
        scaled_height = self._px(height)
        available_width, available_height, exact_work_area = self._usable_work_area_size()
        if available_width <= 0 or available_height <= 0:
            return scaled_width, scaled_height
        return _fit_window_size(
            scaled_width,
            scaled_height,
            available_width,
            available_height,
            horizontal_margin=self._px(16 if exact_work_area else 48),
            vertical_margin=self._px(16 if exact_work_area else 72),
        )

    def _logical_screen_width(self) -> float:
        try:
            return self.root.winfo_screenwidth() / max(self.ui_scale, 1.0)
        except Exception:
            return 1180.0

    def _responsive_content_padding(self, logical_width: float | None = None) -> tuple[int, int, int, int]:
        if logical_width is None:
            logical_width = self._logical_screen_width()
        if logical_width < 560:
            return self._pad(12, 20, 12, 18)
        if logical_width < 820:
            return self._pad(18, 24, 18, 20)
        if logical_width < 1100:
            return self._pad(24, 28, 28, 24)
        return self._pad(42, 34, 58, 28)

    def _responsive_form_padding_units(self, logical_width: float | None = None) -> tuple[int, int, int, int]:
        if logical_width is None:
            logical_width = self._logical_screen_width()
        if logical_width < 560:
            return (10, 14, 10, 14)
        if logical_width < 820:
            return (14, 16, 14, 16)
        if logical_width < 1100:
            return (16, 18, 16, 18)
        return (24, 22, 24, 22)

    def _responsive_sidebar_width(self) -> int:
        logical_width = self._logical_screen_width()
        if self.ui_scale >= 1.75 and logical_width < 900:
            return self._px(200)
        if self.ui_scale >= 1.5 and logical_width < 900:
            return self._px(232)
        if logical_width < 1100:
            return self._px(236)
        return self._px(248)

    def _scrollbar_thumb_image(self, color: str) -> PhotoImage:
        size = self._px(12)
        inset = self._px(2)
        image = PhotoImage(master=self.root, width=size, height=size)
        low = inset
        high = size - inset - 1
        center = (low + high) / 2
        radius = (high - low + 1) / 2 - 0.1
        radius_squared = radius * radius
        for y in range(size):
            for x in range(size):
                if (x - center) ** 2 + (y - center) ** 2 <= radius_squared:
                    image.put(color, (x, y))
        return image

    def _configure_scrollbar_style(self, style: ttk.Style) -> None:
        thumb_element = "HRToolkit.Scrollbar.thumb"
        arrow_elements = {
            "up": "HRToolkit.Vertical.Scrollbar.uparrow",
            "down": "HRToolkit.Vertical.Scrollbar.downarrow",
            "left": "HRToolkit.Horizontal.Scrollbar.leftarrow",
            "right": "HRToolkit.Horizontal.Scrollbar.rightarrow",
        }
        existing_elements = set(style.element_names())
        image_refs: list[PhotoImage] = []
        if thumb_element not in existing_elements:
            thumb_images = (
                self._scrollbar_thumb_image("#EAEAEB"),
                self._scrollbar_thumb_image("#DCDCDD"),
                self._scrollbar_thumb_image("#CCCCCE"),
            )
            image_refs.extend(thumb_images)
            style.element_create(
                thumb_element,
                "image",
                thumb_images[0],
                ("pressed", thumb_images[2]),
                ("active", thumb_images[1]),
                border=self._px(5),
                sticky="nswe",
            )

        missing_arrows = [
            element for element in arrow_elements.values() if element not in existing_elements
        ]
        if missing_arrows:
            transparent_arrow = PhotoImage(
                master=self.root,
                width=self._px(12),
                height=self._px(12),
            )
            image_refs.append(transparent_arrow)
            # 箭头区域继续参与 ttk 原有的逐行滚动绑定，只把图像设为透明。
            for arrow_element in missing_arrows:
                style.element_create(
                    arrow_element,
                    "image",
                    transparent_arrow,
                    sticky="nswe",
                )

        if image_refs:
            stored_images = getattr(self.root, "_hr_scrollbar_images", ())
            if not isinstance(stored_images, tuple):
                stored_images = ()
            images = (*stored_images, *image_refs)
            self._scrollbar_images = images
            setattr(self.root, "_hr_scrollbar_images", images)

        style.layout(
            "Vertical.TScrollbar",
            [
                (
                    "Vertical.Scrollbar.trough",
                    {
                        "sticky": "ns",
                        "children": [
                            (arrow_elements["up"], {"side": "top", "sticky": ""}),
                            (arrow_elements["down"], {"side": "bottom", "sticky": ""}),
                            (thumb_element, {"sticky": "nswe"}),
                        ],
                    },
                )
            ],
        )
        style.layout(
            "Horizontal.TScrollbar",
            [
                (
                    "Horizontal.Scrollbar.trough",
                    {
                        "sticky": "ew",
                        "children": [
                            (arrow_elements["left"], {"side": "left", "sticky": ""}),
                            (arrow_elements["right"], {"side": "right", "sticky": ""}),
                            (thumb_element, {"sticky": "nswe"}),
                        ],
                    },
                )
            ],
        )
        for style_name in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            style.configure(
                style_name,
                arrowsize=self._px(12),
                background=COLOR_SURFACE,
                troughcolor=COLOR_SURFACE,
                bordercolor=COLOR_SURFACE,
                lightcolor=COLOR_SURFACE,
                darkcolor=COLOR_SURFACE,
                relief="flat",
                borderwidth=0,
                gripcount=0,
                gripsize=0,
            )

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
        self.base_font = (family, _font_size(10))
        self.small_font = (family, _font_size(9))
        self.tiny_font = (family, _font_size(8))
        self.title_font = (family, _font_size(18), "bold")
        self.section_font = (family, _font_size(10), "bold")
        self.card_title_font = (family, _font_size(11), "bold")
        self.mono_font = (mono_family, _font_size(10))
        self.root.option_add("*TCombobox*Listbox.font", self.base_font)

        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        self._configure_scrollbar_style(style)

        style.configure(".", font=self.base_font, background=COLOR_BG, foreground=COLOR_TEXT)
        style.configure("App.TFrame", background=COLOR_BG)
        style.configure("Sidebar.TFrame", background=COLOR_SIDEBAR)
        style.configure("Content.TFrame", background=COLOR_BG)
        style.configure(
            "Card.TFrame",
            background=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            relief="solid",
            borderwidth=self._px(1),
        )
        style.configure("InputWrap.TFrame", background=COLOR_SURFACE)
        style.configure("Separator.TFrame", background=COLOR_SIDEBAR_BORDER)
        style.configure("CardSeparator.TFrame", background=COLOR_BORDER_FAINT)
        style.configure("Tooltip.TFrame", background=COLOR_SURFACE, relief="solid", borderwidth=self._px(1), bordercolor=COLOR_BORDER)
        style.configure("Title.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=self.title_font)
        style.configure("Subtitle.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=self.base_font)
        style.configure("Eyebrow.TLabel", background=COLOR_BG, foreground=COLOR_PRIMARY, font=(self.base_font[0], _font_size(9), "bold"))
        style.configure("Section.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=(self.base_font[0], _font_size(10)))
        style.configure("SidebarTitle.TLabel", background=COLOR_SIDEBAR, foreground=COLOR_TEXT, font=(self.base_font[0], _font_size(11), "bold"))
        style.configure("SidebarMuted.TLabel", background=COLOR_SIDEBAR, foreground=COLOR_FAINT, font=self.small_font)
        style.configure("SidebarSection.TLabel", background=COLOR_SIDEBAR, foreground=COLOR_DISABLED, font=(self.base_font[0], _font_size(8), "bold"))
        style.configure("Version.TLabel", background=COLOR_SIDEBAR, foreground=COLOR_DISABLED, font=self.tiny_font)
        style.configure("Tooltip.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.small_font, padding=self._pad(8, 6))
        style.configure("CardTitle.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.card_title_font)
        style.configure("CardHint.TLabel", background=COLOR_SURFACE, foreground=COLOR_FAINT, font=self.small_font)
        style.configure("CardMuted.TLabel", background=COLOR_SURFACE, foreground=COLOR_MUTED, font=self.small_font)
        style.configure(
            "Rename.TLabelframe",
            background=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            relief="solid",
            borderwidth=self._px(1),
        )
        style.configure("Rename.TLabelframe.Label", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.section_font)
        style.configure("App.TLabel", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.base_font)
        style.configure("App.TRadiobutton", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.base_font)
        style.map("App.TRadiobutton", background=[("active", COLOR_SURFACE)])

        # 自定义清晰对勾复选框指示器（彻底解决 clam 默认绘制 "X" 叉号引发的用户困惑）
        try:
            self._img_checkbox_unchecked = PhotoImage(master=self.root, width=16, height=16)
            self._img_checkbox_checked = PhotoImage(master=self.root, width=16, height=16)
            # 未勾选：浅灰圆角边框、纯白背景
            for x in range(16):
                for y in range(16):
                    if (x in (0, 15) and y in (0, 15)):
                        self._img_checkbox_unchecked.put(COLOR_SURFACE, (x, y))
                    elif x in (0, 15) or y in (0, 15):
                        self._img_checkbox_unchecked.put("#CBD5E1", (x, y))
                    else:
                        self._img_checkbox_unchecked.put("#FFFFFF", (x, y))

            # 已勾选：品牌主色实心底色 + 加粗清晰白色对勾 ✓
            for x in range(16):
                for y in range(16):
                    if (x in (0, 15) and y in (0, 15)):
                        self._img_checkbox_checked.put(COLOR_SURFACE, (x, y))
                    else:
                        self._img_checkbox_checked.put(COLOR_PRIMARY, (x, y))

            # 绘制加粗清晰白色对勾 ✓（左短右长）
            check_pixels: set[tuple[int, int]] = set()
            for t in range(3):
                check_pixels.add((3 + t, 8 + t))
                check_pixels.add((3 + t, 9 + t))
            for t in range(7):
                check_pixels.add((6 + t, 9 - t))
                check_pixels.add((6 + t, 10 - t))

            for px, py in check_pixels:
                if 0 <= px < 16 and 0 <= py < 16:
                    self._img_checkbox_checked.put("#FFFFFF", (px, py))

            style.element_create(
                "App.Checkbutton.indicator",
                "image",
                self._img_checkbox_unchecked,
                ("selected", self._img_checkbox_checked),
            )

            for style_name in ("TCheckbutton", "App.TCheckbutton"):
                style.layout(style_name, [
                    ("Checkbutton.padding", {"sticky": "nswe", "children": [
                        ("App.Checkbutton.indicator", {"side": "left", "sticky": ""}),
                        ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [
                            ("Checkbutton.label", {"sticky": "nswe"})
                        ]})
                    ]})
                ])
                style.configure(style_name, background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.base_font, padding=(2, 2))
                style.map(style_name, background=[("active", COLOR_SURFACE)])
        except Exception:
            for style_name in ("TCheckbutton", "App.TCheckbutton"):
                style.configure(style_name, background=COLOR_SURFACE, foreground=COLOR_TEXT, font=self.base_font)
                style.map(style_name, background=[("active", COLOR_SURFACE)])
        style.configure(
            "App.TEntry",
            fieldbackground=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            insertcolor=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            padding=self._pad(12, 8),
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
            padding=self._pad(12, 7),
        )
        style.configure("Change.TNotebook", background=COLOR_BG, borderwidth=0)
        style.configure(
            "Change.TNotebook.Tab",
            padding=self._pad(18, 8),
            background=COLOR_NAV_SELECTED,
            foreground=COLOR_NAV_TEXT,
            bordercolor=COLOR_BORDER,
        )
        style.map(
            "Change.TNotebook.Tab",
            background=[("selected", COLOR_SURFACE)],
            foreground=[("selected", COLOR_PRIMARY)],
        )
        style.configure(
            "History.Treeview",
            background=COLOR_SURFACE,
            fieldbackground=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            rowheight=self._px(34),
            borderwidth=0,
        )
        style.configure(
            "History.Treeview.Heading",
            background=COLOR_SURFACE_ALT,
            foreground=COLOR_MUTED,
            font=(self.base_font[0], _font_size(9), "bold"),
            padding=self._pad(8, 7),
            relief="flat",
        )
        style.map(
            "History.Treeview",
            background=[("selected", COLOR_PRIMARY_SOFT)],
            foreground=[("selected", COLOR_TEXT)],
        )
        style.configure("Workspace.TFrame", background=COLOR_SURFACE)
        style.configure("WorkspaceRail.TFrame", background=COLOR_SURFACE_ALT)
        style.configure(
            "WorkspaceTitle.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            font=(self.base_font[0], _font_size(12), "bold"),
        )
        style.configure(
            "WorkspaceProject.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_PRIMARY,
            font=(self.base_font[0], _font_size(9), "bold"),
        )
        style.configure("WorkspaceMuted.TLabel", background=COLOR_SURFACE, foreground=COLOR_FAINT, font=self.small_font)
        style.configure(
            "Workspace.Treeview",
            background=COLOR_SURFACE,
            fieldbackground=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            rowheight=self._px(29),
            borderwidth=0,
        )
        style.map(
            "Workspace.Treeview",
            background=[("selected", COLOR_PRIMARY_SOFT)],
            foreground=[("selected", COLOR_TEXT)],
        )

    def _build_layout(self) -> None:
        root_frame = ttk.Frame(self.root, padding=0, style="App.TFrame")
        root_frame.pack(fill=BOTH, expand=True)

        def _clear_entry_focus(event) -> None:
            # 点击输入框以外的地方时收起光标（可编辑控件自己会接管焦点）
            try:
                widget_class = event.widget.winfo_class()
            except Exception:
                return
            if widget_class in ("TEntry", "Entry", "Text", "TCombobox", "Listbox"):
                return
            try:
                self.root.focus_set()
            except Exception:
                pass

        self.root.bind("<Button-1>", _clear_entry_focus, add="+")

        left_frame = ttk.Frame(root_frame, width=self._responsive_sidebar_width(), style="Sidebar.TFrame")
        left_frame.pack(side=LEFT, fill=Y)
        left_frame.pack_propagate(False)
        left_frame.grid_propagate(False)
        left_frame.grid_rowconfigure(0, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)

        left_canvas = Canvas(left_frame, width=1, bg=COLOR_SIDEBAR, highlightthickness=0, bd=0)
        left_canvas.grid(row=0, column=0, sticky="nsew")
        left_vscroll = ttk.Scrollbar(left_frame, orient=VERTICAL, command=left_canvas.yview)
        left_vscroll.grid(row=0, column=1, sticky="ns")
        left_vscroll.grid_remove()
        left_canvas.configure(yscrollcommand=left_vscroll.set)

        left_content = ttk.Frame(left_canvas, padding=self._pad(12, 16, 12, 14), style="Sidebar.TFrame")
        left_canvas_window = left_canvas.create_window((0, 0), window=left_content, anchor="nw")

        def _sync_left_canvas(_event=None) -> None:
            canvas_width = max(left_canvas.winfo_width(), 1)
            canvas_height = max(left_canvas.winfo_height(), 1)
            content_height = left_content.winfo_reqheight()
            window_height = max(content_height, canvas_height)
            left_canvas.itemconfig(left_canvas_window, width=canvas_width, height=window_height)
            left_canvas.configure(scrollregion=(0, 0, canvas_width, window_height))
            if content_height > canvas_height:
                if not left_vscroll.winfo_ismapped():
                    left_vscroll.grid(row=0, column=1, sticky="ns")
            else:
                left_canvas.yview_moveto(0)
                if left_vscroll.winfo_ismapped():
                    left_vscroll.grid_remove()

        left_wheel_accumulator = 0.0

        def _left_mousewheel_units(event) -> int:
            nonlocal left_wheel_accumulator
            if sys.platform.startswith("win"):
                left_wheel_accumulator += -float(getattr(event, "delta", 0) or 0) / 120.0
                units = int(left_wheel_accumulator)
                left_wheel_accumulator -= units
                return units
            if sys.platform == "darwin":
                delta = int(getattr(event, "delta", 0) or 0)
                if abs(delta) >= 120:
                    return int(delta / -40)
                return int(-1 * delta)
            if getattr(event, "num", None) == 4:
                return -1
            if getattr(event, "num", None) == 5:
                return 1
            return 0

        def _left_touchpad_deltas(event) -> tuple[int, int]:
            encoded_delta = int(getattr(event, "delta", 0) or 0)
            delta_x = encoded_delta >> 16
            low_word = encoded_delta & 0xFFFF
            delta_y = low_word if low_word < 0x8000 else low_word - 0x10000
            return delta_x, delta_y

        def _left_can_scroll() -> bool:
            first, last = left_canvas.yview()
            return first > 0 or last < 1

        def _on_left_wheel(event):
            delta_units = _left_mousewheel_units(event)
            if delta_units and _left_can_scroll():
                left_canvas.yview_scroll(delta_units, "units")
                return "break"
            return None

        def _on_left_touchpad(event):
            _delta_x, delta_y = _left_touchpad_deltas(event)
            if delta_y and _left_can_scroll():
                canvas_height = max(left_canvas.winfo_height(), 1)
                top = left_canvas.yview()[0]
                left_canvas.yview_moveto(top - (delta_y / canvas_height))
                return "break"
            return None

        def _on_left_linux_up(_event):
            if _left_can_scroll():
                left_canvas.yview_scroll(-1, "units")
                return "break"
            return None

        def _on_left_linux_down(_event):
            if _left_can_scroll():
                left_canvas.yview_scroll(1, "units")
                return "break"
            return None

        LEFT_SCROLL_TAG = "LeftPanelScroll"

        def _safe_bind_left_class(sequence: str, handler) -> None:
            try:
                left_canvas.bind_class(LEFT_SCROLL_TAG, sequence, handler)
            except Exception:
                pass

        _safe_bind_left_class("<MouseWheel>", _on_left_wheel)
        _safe_bind_left_class("<TouchpadScroll>", _on_left_touchpad)
        _safe_bind_left_class("<Button-4>", _on_left_linux_up)
        _safe_bind_left_class("<Button-5>", _on_left_linux_down)

        def _apply_left_scroll_tag(widget) -> None:
            try:
                current = list(widget.bindtags())
                if LEFT_SCROLL_TAG not in current:
                    widget.bindtags([LEFT_SCROLL_TAG] + current)
            except Exception:
                pass
            for child in widget.winfo_children():
                _apply_left_scroll_tag(child)

        brand_row = ttk.Frame(left_content, style="Sidebar.TFrame")
        brand_row.pack(fill="x", padx=self._pad(6), pady=self._pad(2, 16))
        brand_mark = Canvas(brand_row, width=self._px(26), height=self._px(26), bg=COLOR_SIDEBAR, highlightthickness=0, bd=0)
        brand_mark.pack(side=LEFT)
        self._draw_round_rect(brand_mark, self._pxf(0.5), self._pxf(0.5), self._pxf(25.5), self._pxf(25.5), self._pxf(7), fill=COLOR_PRIMARY, outline="")
        brand_mark.create_text(self._pxf(13), self._pxf(13), text="HR", fill="#ffffff", font=(self.base_font[0], _font_size(8), "bold"))
        brand_text = ttk.Frame(brand_row, style="Sidebar.TFrame")
        brand_text.pack(side=LEFT, fill="x", expand=True, padx=self._pad(9, 0))
        ttk.Label(brand_text, text=APP_DISPLAY_NAME, style="SidebarTitle.TLabel").pack(anchor="w")
        ttk.Label(brand_text, text=APP_SUBTITLE, style="SidebarMuted.TLabel").pack(anchor="w")

        ttk.Label(left_content, text="工作项目", style="SidebarSection.TLabel").pack(
            anchor="w", padx=self._pad(9), pady=self._pad(0, 6)
        )
        project_card = ttk.Frame(left_content, padding=self._pad(10, 9), style="Card.TFrame")
        project_card.pack(fill="x", padx=self._pad(3))
        self.sidebar_project_name_label = Label(
            project_card,
            textvariable=self.workspace_project_name,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=self.section_font,
            anchor="w",
            justify="left",
        )
        self.sidebar_project_name_label.pack(fill="x")
        self.sidebar_project_path_label = Label(
            project_card,
            textvariable=self.sidebar_project_summary,
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.tiny_font,
            anchor="w",
            justify="left",
        )
        self.sidebar_project_path_label.pack(fill="x", pady=self._pad(3, 0))
        project_buttons = ttk.Frame(left_content, style="Sidebar.TFrame")
        project_buttons.pack(fill="x", padx=self._pad(3), pady=self._pad(7, 2))
        CodexButton(
            project_buttons,
            text="新建项目",
            command=self._create_workspace_project,
            width=88,
            min_width=78,
            height=30,
        ).pack(side=LEFT)
        CodexButton(
            project_buttons,
            text="打开项目",
            command=self._open_workspace_project,
            width=88,
            min_width=78,
            height=30,
        ).pack(side=LEFT, padx=self._pad(6, 0))
        self.sidebar_recent_project_button = CodexButton(
            left_content,
            text="最近项目",
            command=self._show_recent_projects_menu,
            variant="link",
            width=88,
            min_width=78,
            height=25,
        )
        self.sidebar_recent_project_button.pack(anchor="w", padx=self._pad(4), pady=self._pad(1, 2))

        nav_frame = ttk.Frame(left_content, style="Sidebar.TFrame")
        nav_frame.pack(fill="x")
        for group_label, group_tools in NAV_GROUPS:
            ttk.Label(nav_frame, text=group_label, style="SidebarSection.TLabel").pack(anchor="w", padx=self._pad(9), pady=self._pad(12, 5))
            for tool_id in group_tools:
                item = SidebarItem(
                    nav_frame,
                    text=TOOL_NAV_LABELS[tool_id],
                    icon_id=tool_id,
                    command=lambda selected=tool_id: self._select_tool(selected),
                )
                item.pack(fill="x", pady=self._px(1))
                item.set_selected(tool_id == self.current_tool)
                self.nav_buttons[tool_id] = item

        sidebar_footer = ttk.Frame(left_content, style="Sidebar.TFrame")
        sidebar_footer.pack(side="bottom", fill="x")
        self.history_nav_item = SidebarItem(
            sidebar_footer,
            text="旧版记录",
            icon_id="clock",
            command=self._show_history_view,
        )
        self.history_nav_item.pack(fill="x", pady=self._pad(0, 2))
        tutorial_item = SidebarItem(
            sidebar_footer,
            text="使用教程",
            icon_id="tutorial",
            command=self._open_tutorial_window,
            muted=True,
        )
        tutorial_item.pack(fill="x")
        version_row = ttk.Frame(sidebar_footer, style="Sidebar.TFrame")
        version_row.pack(fill="x", pady=self._pad(8, 0))
        ttk.Frame(version_row, height=self._px(1), style="Separator.TFrame").pack(fill="x", pady=self._pad(0, 9))
        version_line = ttk.Frame(version_row, style="Sidebar.TFrame")
        version_line.pack(fill="x", padx=self._pad(9))
        ttk.Label(version_line, text=f"v{__version__}", style="Version.TLabel").pack(side=LEFT)
        version_dot = Canvas(version_line, width=self._px(9), height=self._px(9), bg=COLOR_SIDEBAR, highlightthickness=0, bd=0)
        version_dot.pack(side=LEFT, padx=self._pad(6, 0), pady=self._pad(1, 0))
        version_dot.create_oval(self._pxf(2), self._pxf(2), self._pxf(7), self._pxf(7), fill=COLOR_SUCCESS_DOT, outline="")
        ttk.Label(version_line, text="本地处理 · 不上传数据", style="SidebarMuted.TLabel").pack(side=RIGHT)

        _apply_left_scroll_tag(left_canvas)
        _apply_left_scroll_tag(left_content)
        left_content.bind("<Configure>", _sync_left_canvas)
        left_canvas.bind("<Configure>", _sync_left_canvas)
        self.root.after_idle(_sync_left_canvas)

        ttk.Frame(root_frame, width=self._px(1), style="Separator.TFrame").pack(side=LEFT, fill=Y)

        self._workspace_main_area = ttk.Frame(root_frame, style="Content.TFrame")
        self._workspace_main_area.pack(side=LEFT, fill=BOTH, expand=True)
        self._workspace_main_area.grid_rowconfigure(0, weight=1)
        self._workspace_main_area.grid_columnconfigure(0, weight=1, minsize=0)
        self._main_view_host = ttk.Frame(self._workspace_main_area, style="Content.TFrame")
        self._main_view_host.grid(row=0, column=0, sticky="nsew")
        self._build_workspace_panel(self._workspace_main_area)

        # Scrollable right panel: Canvas acts as the viewport; right_frame is
        # the inner content frame that all existing children are placed into.
        right_outer = ttk.Frame(self._main_view_host, style="Content.TFrame")
        right_outer.pack(side=RIGHT, fill=BOTH, expand=True)
        self._tool_view = right_outer

        right_vscroll = ttk.Scrollbar(right_outer, orient=VERTICAL)
        right_vscroll.pack(side=RIGHT, fill=Y)

        self._right_canvas = Canvas(
            right_outer,
            bg=COLOR_BG,
            highlightthickness=0,
            bd=0,
            yscrollcommand=right_vscroll.set,
        )
        self._right_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        right_vscroll.config(command=self._right_canvas.yview)

        right_frame = ttk.Frame(self._right_canvas, padding=self._responsive_content_padding(), style="Content.TFrame")
        self._right_canvas_window = self._right_canvas.create_window(
            (0, 0), window=right_frame, anchor="nw"
        )
        self._right_canvas_sync_pending = False
        self._right_canvas_sync_repeat = 0

        def _split_dimension_values(value) -> list[int]:
            try:
                if isinstance(value, tuple):
                    parts = value
                else:
                    parts = self.root.tk.splitlist(value)
                return [int(round(float(part))) for part in parts]
            except Exception:
                try:
                    return [int(round(float(value)))]
                except Exception:
                    return []

        def _frame_vertical_padding_sum(value) -> int:
            parts = _split_dimension_values(value)
            if not parts:
                return 0
            if len(parts) == 1:
                return parts[0] * 2
            if len(parts) == 2:
                return parts[1] * 2
            if len(parts) >= 4:
                return parts[1] + parts[3]
            return parts[1] * 2

        def _pack_vertical_padding_sum(value) -> int:
            parts = _split_dimension_values(value)
            if not parts:
                return 0
            if len(parts) == 1:
                return parts[0] * 2
            return parts[0] + parts[1]

        def _right_frame_natural_height() -> int:
            try:
                height = _frame_vertical_padding_sum(right_frame.cget("padding"))
            except Exception:
                height = 0
            children = []
            try:
                children = list(right_frame.pack_slaves())
            except Exception:
                pass
            if not children:
                return right_frame.winfo_reqheight()
            for child in children:
                try:
                    pack_info = child.pack_info()
                except Exception:
                    continue
                try:
                    height += child.winfo_reqheight()
                except Exception:
                    height += child.winfo_height()
                height += _pack_vertical_padding_sum(pack_info.get("pady", 0))
            return height

        self._last_canvas_window_size = (0, 0)
        self._last_scroll_region = (0, 0, 0, 0)

        def _sync_right_canvas_window(_event=None):
            canvas_width = self._right_canvas.winfo_width()
            canvas_height = self._right_canvas.winfo_height()
            if canvas_width <= 1:
                return
            content_height = _right_frame_natural_height()
            window_height = max(content_height, canvas_height)
            if (canvas_width, window_height) != self._last_canvas_window_size:
                self._last_canvas_window_size = (canvas_width, window_height)
                self._right_canvas.itemconfig(
                    self._right_canvas_window,
                    width=canvas_width,
                    height=window_height,
                )
            region = (0, 0, canvas_width, window_height)
            if region != self._last_scroll_region:
                self._last_scroll_region = region
                self._right_canvas.configure(
                    scrollregion=region
                )
            if window_height <= canvas_height:
                self._right_canvas.yview_moveto(0)

        def _run_right_canvas_sync():
            self._right_canvas_sync_pending = False
            _sync_right_canvas_window()

        def _queue_right_canvas_sync() -> None:
            if self._right_canvas_sync_pending:
                return
            self._right_canvas_sync_pending = True
            self.root.after_idle(_run_right_canvas_sync)

        def _schedule_right_canvas_sync(_event=None):
            _queue_right_canvas_sync()

        self._sync_right_canvas_window = _schedule_right_canvas_sync
        right_frame.bind("<Configure>", _schedule_right_canvas_sync)
        self._right_canvas.bind("<Configure>", _schedule_right_canvas_sync)

        SCROLL_TAG = "RightPanelScroll"

        def _scroll_page(delta_units: int) -> None:
            self._right_canvas.yview_scroll(delta_units, "units")

        def _scroll_page_pixels(delta_y: int) -> None:
            total_height = max(self._right_canvas.winfo_height(), 1)
            try:
                parts = [float(part) for part in self._right_canvas.cget("scrollregion").split()]
                if len(parts) == 4:
                    total_height = max(parts[3] - parts[1], 1)
            except Exception:
                try:
                    bbox = self._right_canvas.bbox("all")
                    if bbox:
                        total_height = max(bbox[3] - bbox[1], 1)
                except Exception:
                    pass
            top = self._right_canvas.yview()[0]
            new_top = max(0.0, min(1.0, top - (delta_y / total_height)))
            self._right_canvas.yview_moveto(new_top)

        def _touchpad_deltas(event) -> tuple[int, int]:
            encoded_delta = int(getattr(event, "delta", 0) or 0)
            delta_x = encoded_delta >> 16
            low_word = encoded_delta & 0xFFFF
            delta_y = low_word if low_word < 0x8000 else low_word - 0x10000
            return delta_x, delta_y

        wheel_accumulator = 0.0

        def _mousewheel_units(event) -> int:
            nonlocal wheel_accumulator
            if sys.platform.startswith("win"):
                wheel_accumulator += -float(getattr(event, "delta", 0) or 0) / 120.0
                units = int(wheel_accumulator)
                wheel_accumulator -= units
                return units
            if sys.platform == "darwin":
                delta = int(getattr(event, "delta", 0) or 0)
                if abs(delta) >= 120:
                    return int(delta / -40)
                return int(-1 * delta)
            if getattr(event, "num", None) == 4:
                return -1
            if getattr(event, "num", None) == 5:
                return 1
            return 0

        def _safe_bind_class(sequence: str, handler) -> None:
            try:
                self.root.bind_class(SCROLL_TAG, sequence, handler)
            except Exception:
                pass

        def _safe_bind_widget(widget, sequence: str, handler) -> None:
            try:
                widget.bind(sequence, handler)
            except Exception:
                pass

        def _on_scroll_tag_wheel(event):
            delta_units = _mousewheel_units(event)
            if delta_units:
                _scroll_page(delta_units)
            return "break"

        def _on_scroll_tag_touchpad(event):
            _delta_x, delta_y = _touchpad_deltas(event)
            if delta_y:
                _scroll_page_pixels(delta_y)
            return "break"

        def _on_scroll_tag_linux_up(event):
            _scroll_page(-1)
            return "break"

        def _on_scroll_tag_linux_down(event):
            _scroll_page(1)
            return "break"

        # Register handlers on the named tag (not on any specific widget)
        _safe_bind_class("<MouseWheel>", _on_scroll_tag_wheel)
        _safe_bind_class("<TouchpadScroll>", _on_scroll_tag_touchpad)
        _safe_bind_class("<Button-4>", _on_scroll_tag_linux_up)
        _safe_bind_class("<Button-5>", _on_scroll_tag_linux_down)

        def _apply_scroll_tag(widget) -> None:
            if hasattr(self, "log_text") and widget is self.log_text:
                return
            try:
                current = list(widget.bindtags())
                if SCROLL_TAG not in current:
                    widget.bindtags([SCROLL_TAG] + current)
            except Exception:
                pass
            for child in widget.winfo_children():
                _apply_scroll_tag(child)

        # Also keep a direct canvas binding as fallback (when cursor is on
        # the canvas background between widgets)
        _safe_bind_widget(self._right_canvas, "<MouseWheel>", _on_scroll_tag_wheel)
        _safe_bind_widget(self._right_canvas, "<TouchpadScroll>", _on_scroll_tag_touchpad)
        _safe_bind_widget(self._right_canvas, "<Button-4>", _on_scroll_tag_linux_up)
        _safe_bind_widget(self._right_canvas, "<Button-5>", _on_scroll_tag_linux_down)

        title_row = ttk.Frame(right_frame, style="Content.TFrame")
        title_row.pack(fill="x")
        title_row.columnconfigure(0, weight=1)
        ttk.Label(title_row, textvariable=self.tool_group, style="Eyebrow.TLabel").grid(row=0, column=0, sticky="w")
        self.title_label = ttk.Label(title_row, textvariable=self.tool_title, style="Title.TLabel", justify="left")
        self.title_label.grid(row=1, column=0, sticky="w", pady=self._pad(5, 0))
        title_actions = ttk.Frame(title_row, style="Content.TFrame")
        title_actions.grid(row=0, column=1, rowspan=2, sticky="ne")
        self.check_update_button = CodexButton(
            title_actions,
            text="检查更新",
            command=self._check_updates_manually,
            icon="↻",
            width=118,
        )
        self.check_update_button.pack(side=LEFT)
        self.workspace_title_button = CodexButton(
            title_actions,
            text="项目文件",
            command=self._toggle_workspace_panel,
            width=92,
            min_width=82,
            variant="tonal",
        )
        self.subtitle_label = Label(
            right_frame,
            textvariable=self.tool_description,
            bg=COLOR_BG,
            fg=COLOR_MUTED,
            font=self.base_font,
            justify="left",
            anchor="w",
        )
        self.subtitle_label.pack(anchor="w", fill="x", pady=self._pad(8, 22))

        self._last_text_wraps = None

        def _update_text_wraps(_event=None) -> None:
            title_row_width = title_row.winfo_width()
            if title_row_width <= 1:
                title_row_width = self._px(240)
            actions_width = max(self._px(118), title_actions.winfo_reqwidth())
            tight_header = title_row_width < actions_width + self._px(360)
            if tight_header:
                title_wrap = title_row_width
            else:
                title_wrap = title_row_width - actions_width - self._px(24)

            title_wrap = max(1, title_wrap)
            subtitle_wrap = max(1, title_row_width - self._px(8))
            wraps_key = (tight_header, title_wrap, subtitle_wrap)
            if wraps_key == self._last_text_wraps:
                return
            self._last_text_wraps = wraps_key

            if tight_header:
                title_actions.grid_configure(row=2, column=0, columnspan=2, rowspan=1, sticky="w", pady=self._pad(12, 0))
            else:
                title_actions.grid_configure(row=0, column=1, columnspan=1, rowspan=2, sticky="ne", pady=0)

            self.title_label.configure(wraplength=title_wrap)
            self.subtitle_label.configure(wraplength=subtitle_wrap)

        self._update_title_text_wraps = _update_text_wraps
        title_row.bind("<Configure>", _update_text_wraps, add="+")
        right_frame.bind("<Configure>", _update_text_wraps, add="+")
        self.root.after_idle(_update_text_wraps)

        self.change_tabs = ttk.Notebook(right_frame, style="Change.TNotebook")
        self.change_tabs.add(ttk.Frame(self.change_tabs, style="Content.TFrame"), text="异动表汇总")
        self.change_tabs.add(ttk.Frame(self.change_tabs, style="Content.TFrame"), text="花名册更新")
        self.change_tabs.bind("<<NotebookTabChanged>>", self._on_change_tab_changed)

        # 合并后的上传入口卡片：空态为虚线拖放区样式，选中后显示文件条目
        self.upload_card = RoundedCard(right_frame, padding=(20, 16, 20, 18))
        self.upload_card.pack(fill="x")
        upload_header = ttk.Frame(self.upload_card.inner, style="InputWrap.TFrame")
        upload_header.pack(fill="x", pady=self._pad(0, 10))
        ttk.Label(upload_header, textvariable=self.input_label, style="CardTitle.TLabel").pack(side=LEFT)
        ttk.Label(upload_header, textvariable=self.input_hint, style="CardHint.TLabel").pack(side=LEFT, padx=self._pad(9, 0))
        self.upload_add_button = CodexButton(
            upload_header,
            text="＋ 添加",
            command=self._show_add_input_menu,
            variant="link",
            width=72,
            min_width=56,
            height=24,
        )
        self.upload_body = ttk.Frame(self.upload_card.inner, style="InputWrap.TFrame")
        self.upload_body.pack(fill="x")

        self.form_card = RoundedCard(right_frame, padding=self._responsive_form_padding_units())
        self.form_card.pack(fill="x", pady=self._pad(14, 0))
        form = self.form_card.inner
        self.form = form
        self._form_layout_mode = LAYOUT_MODE_WIDE
        self._form_compact_layout = False
        self._material_checkbox_columns = 4
        self._summary_row_visible = True
        self._output_row_visible = True
        self._rename_row_visible = True
        self._stats_range_row_visible = True
        self._form_rows = {}

        def make_input_row(row_key: str, row_index: int, label_text, value_var: StringVar, command) -> tuple[ttk.Label, ttk.Frame, CodexButton]:
            if isinstance(label_text, StringVar):
                label = ttk.Label(form, textvariable=label_text, style="App.TLabel")
            else:
                label = ttk.Label(form, text=label_text, style="App.TLabel")
            label.grid(row=row_index, column=0, sticky="w", pady=self._px(7))
            input_frame = ttk.Frame(form, style="InputWrap.TFrame")
            input_frame.grid(row=row_index, column=1, sticky="ew", padx=self._pad(18, 0), pady=self._px(7))
            entry = ttk.Entry(input_frame, textvariable=value_var, style="App.TEntry")
            entry.pack(side=LEFT, fill=BOTH, expand=True)
            button_bar = ttk.Frame(input_frame, style="InputWrap.TFrame")
            button_bar.pack(side=RIGHT)
            button = CodexButton(button_bar, text="选择", command=command, width=64, min_width=56, variant="link")
            setattr(button, "_hr_picker_visible", True)
            button.pack(side=RIGHT, padx=self._pad(10, 0))
            setattr(input_frame, "_hr_entry", entry)
            setattr(input_frame, "_hr_button_bar", button_bar)
            self._form_rows[row_key] = {
                "index": row_index,
                "label": label,
                "frame": input_frame,
                "entry": entry,
                "button_bar": button_bar,
            }
            return label, input_frame, button

        self.summary_label_widget, self.summary_entry_widget, self.summary_choose_button = make_input_row(
            "summary",
            0,
            self.summary_label,
            self.summary_path,
            self._choose_summary,
        )
        self.output_label_widget, self.output_entry_widget, self.output_choose_button = make_input_row(
            "output",
            1,
            "保存位置",
            self.output_dir,
            self._choose_output,
        )
        self.output_choose_button.configure(text="更改")
        self.change_summary_folder_button = CodexButton(
            getattr(self.summary_entry_widget, "_hr_button_bar"),
            text="选择文件夹",
            command=self._choose_change_summary_folder,
            width=96,
            variant="link",
        )
        self.change_summary_file_button = CodexButton(
            getattr(self.summary_entry_widget, "_hr_button_bar"),
            text="选择文件",
            command=self._choose_change_summary_file,
            width=84,
            variant="link",
        )
        self.clear_summary_button = CodexButton(
            getattr(self.summary_entry_widget, "_hr_button_bar"),
            text="✕ 清空",
            command=lambda: self.summary_path.set(""),
            width=64,
            min_width=50,
            variant="secondary",
        )

        # 花名册/汇总表一行只显示文件名，完整路径悬停查看（对应设计稿第二张卡片的行）
        summary_entry = self._form_rows["summary"]["entry"]
        summary_entry.destroy()
        self.summary_display = Label(
            self.summary_entry_widget,
            text="未选择",
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.base_font,
            anchor="w",
        )
        self._form_rows["summary"]["entry"] = self.summary_display
        setattr(self.summary_entry_widget, "_hr_entry", self.summary_display)
        self.summary_path.trace_add("write", lambda *_args: self._update_summary_display())
        self._bind_path_tooltip(self.summary_display, lambda: self.summary_path.get().strip())
        self._update_summary_display()

        self.rename_options_frame = ttk.LabelFrame(form, text="批量改名", padding=self._px(12), style="Rename.TLabelframe")
        self.rename_options_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=self._pad(10, 0))
        ttk.Label(self.rename_options_frame, text="操作", style="App.TLabel").grid(
            row=0, column=0, sticky="w", pady=self._px(5)
        )
        self.rename_mode_widget = ttk.Combobox(
            self.rename_options_frame,
            textvariable=self.rename_mode,
            values=list(RENAME_MODE_LABELS.keys()),
            state="readonly",
            width=28,
            style="App.TCombobox",
        )
        self.rename_mode_widget.grid(row=0, column=1, sticky="w", padx=self._px(12), pady=self._px(5))
        self.rename_mode_widget.bind("<<ComboboxSelected>>", self._on_rename_mode_changed)

        self.rename_target_label_widget = ttk.Label(self.rename_options_frame, textvariable=self.rename_target_label, style="App.TLabel")
        self.rename_target_label_widget.grid(row=1, column=0, sticky="w", pady=self._px(5))
        self.rename_target_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_target_name, style="App.TEntry")
        self.rename_target_widget.grid(row=1, column=1, sticky="ew", padx=self._px(12), pady=self._px(5))

        self.rename_text_label_widget = ttk.Label(self.rename_options_frame, textvariable=self.rename_text_label, style="App.TLabel")
        self.rename_text_label_widget.grid(row=2, column=0, sticky="w", pady=self._px(5))
        self.rename_text_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_text, style="App.TEntry")
        self.rename_text_widget.grid(row=2, column=1, sticky="ew", padx=self._px(12), pady=self._px(5))

        self.rename_replacement_label_widget = ttk.Label(self.rename_options_frame, textvariable=self.rename_replacement_label, style="App.TLabel")
        self.rename_replacement_label_widget.grid(row=3, column=0, sticky="w", pady=self._px(5))
        self.rename_replacement_widget = ttk.Entry(self.rename_options_frame, textvariable=self.rename_replacement_name, style="App.TEntry")
        self.rename_replacement_widget.grid(row=3, column=1, sticky="ew", padx=self._px(12), pady=self._px(5))

        self.rename_file_type_label_widget = ttk.Label(self.rename_options_frame, text="文件类型", style="App.TLabel")
        self.rename_file_type_label_widget.grid(row=4, column=0, sticky="w", pady=self._px(5))
        self.rename_file_type_widget = ttk.Combobox(
            self.rename_options_frame,
            textvariable=self.rename_file_type,
            values=["文件夹", "PDF", "图片（jpg/png/gif等）", "文档（doc/xls/ppt/txt等）", "全部"],
            state="readonly",
            width=22,
        )
        self.rename_file_type_widget.grid(row=4, column=1, sticky="w", padx=self._px(12), pady=self._px(5))

        self.rename_options_frame.columnconfigure(1, weight=1)

        self.material_options_frame = ttk.LabelFrame(form, text="资料检索与打包设置", padding=self._px(12), style="Rename.TLabelframe")
        self.material_options_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=self._pad(10, 0))

        # ── 1. 资料库组织形式（与输出归类模式相互独立） ──
        self.material_library_mode_row = ttk.Frame(self.material_options_frame, style="InputWrap.TFrame")
        self.material_library_mode_row.pack(fill="x", expand=True, pady=(self._px(2), self._px(7)))
        self.material_library_mode_label = ttk.Label(
            self.material_library_mode_row, text="资料库形式", style="App.TLabel", width=8
        )
        self.material_library_mode_label.pack(side=LEFT, anchor="center")
        self.material_library_mode_combo = ttk.Combobox(
            self.material_library_mode_row,
            textvariable=self.material_library_mode,
            values=list(LIBRARY_MODE_LABELS.keys()),
            state="readonly",
            width=28,
            style="App.TCombobox",
        )
        self.material_library_mode_combo.pack(side=LEFT, padx=(self._px(8), self._px(12)))
        self.material_library_mode_combo.bind("<<ComboboxSelected>>", self._on_material_library_mode_changed)
        self.material_library_mode_hint = ttk.Label(
            self.material_library_mode_row,
            text="原模式按姓名文件夹查找",
            style="CardHint.TLabel",
        )
        self.material_library_mode_hint.pack(side=LEFT, fill="x", expand=True)

        # ── 2. 目标人员输入行（全宽拉伸） ──
        self.material_target_row = ttk.Frame(self.material_options_frame, style="InputWrap.TFrame")
        self.material_target_row.pack(fill="x", expand=True, pady=(self._px(2), self._px(4)))

        self.material_target_label = ttk.Label(
            self.material_target_row, text="目标人员", style="App.TLabel", width=8
        )
        self.material_target_label.pack(side=LEFT, anchor="center")

        self.material_target_wrap = ttk.Frame(self.material_target_row, style="InputWrap.TFrame")
        self.material_target_wrap.pack(side=LEFT, fill="x", expand=True, padx=(self._px(8), 0))

        self.material_target_entry = ttk.Entry(
            self.material_target_wrap,
            textvariable=self.material_target_input,
            style="App.TEntry",
        )
        self.material_target_entry.pack(side=LEFT, fill="x", expand=True)

        self.material_target_clear_btn = CodexButton(
            self.material_target_wrap,
            text="✕ 清空",
            command=lambda: self.material_target_input.set(""),
            width=56,
            min_width=48,
            variant="link",
        )
        self.material_target_clear_btn.pack(side=RIGHT, padx=(self._px(8), 0))

        # 目标人员提示语（简短精炼）
        self.material_input_hint = ttk.Label(
            self.material_options_frame,
            text="输入姓名或身份证（多人用逗号隔开，如“张三, 李四”）；留空则按名单表格处理",
            style="CardHint.TLabel",
        )
        self.material_input_hint.pack(fill="x", padx=(self._px(76), 0), pady=(0, self._px(8)))

        # ── 3. 打包与检索设置行 ──
        self.material_opts_row = ttk.Frame(self.material_options_frame, style="InputWrap.TFrame")
        self.material_opts_row.pack(fill="x", pady=(self._px(4), self._px(4)))

        self.material_opts_label = ttk.Label(
            self.material_opts_row, text="打包设置", style="App.TLabel", width=8
        )
        self.material_opts_label.pack(side=LEFT, anchor="center")

        self.material_opts_checks_frame = ttk.Frame(self.material_opts_row, style="InputWrap.TFrame")
        self.material_opts_checks_frame.pack(side=LEFT, fill="x", expand=True, padx=(self._px(8), 0))

        self.material_collect_all_check = ttk.Checkbutton(
            self.material_opts_checks_frame,
            text="全部（直接拷贝匹配到的人员整个文件夹）",
            variable=self.material_collect_all,
            command=self._on_material_collect_all_changed,
            style="App.TCheckbutton",
        )
        self.material_collect_all_check.pack(side=LEFT, padx=(0, self._px(16)))

        self.material_zip_check = ttk.Checkbutton(
            self.material_opts_checks_frame,
            text="生成 ZIP 压缩包",
            variable=self.material_create_zip,
            style="App.TCheckbutton",
        )
        self.material_zip_check.pack(side=LEFT, padx=(0, self._px(16)))

        # OCR 缓存：精简文案并置于右侧不起眼位置
        self.material_use_ocr_cache_check = ttk.Checkbutton(
            self.material_opts_checks_frame,
            text="启用缓存",
            variable=self.material_use_ocr_cache,
            style="App.TCheckbutton",
        )
        self.material_use_ocr_cache_check.pack(side=RIGHT, padx=(0, self._px(4)))

        # 全部模式提示语（简短精炼）
        self.material_types_hint = ttk.Label(
            self.material_options_frame,
            text="取消勾选「全部」后可按需勾选材料类型（如身份证、劳动合同等）",
            style="CardHint.TLabel",
        )
        self.material_types_hint.pack(fill="x", padx=(self._px(76), 0), pady=(0, self._px(6)))

        # ── 4. 指定材料选择区（全部模式下折叠隐藏） ──
        self.material_types_section = ttk.Frame(self.material_options_frame, style="InputWrap.TFrame")
        self.material_types_section.pack(fill="x", pady=(self._px(6), self._px(2)))

        self.material_header_row = ttk.Frame(self.material_types_section, style="InputWrap.TFrame")
        self.material_header_row.pack(side=LEFT, anchor="nw", padx=(0, self._px(8)))

        self.material_types_label = ttk.Label(self.material_header_row, text="指定材料", style="App.TLabel", width=8)
        self.material_types_label.pack(side=TOP, anchor="w")

        mat_actions_frame = ttk.Frame(self.material_header_row, style="InputWrap.TFrame")
        mat_actions_frame.pack(side=TOP, anchor="w", pady=(self._px(4), 0))

        CodexButton(
            mat_actions_frame,
            text="全选",
            command=self._select_all_material_types,
            width=36,
            min_width=32,
            variant="link",
        ).pack(side=LEFT, padx=(0, self._px(4)))

        CodexButton(
            mat_actions_frame,
            text="取消全选",
            command=self._deselect_all_material_types,
            width=56,
            min_width=48,
            variant="link",
        ).pack(side=LEFT)

        self.mat_checks_frame = ttk.Frame(self.material_types_section, style="InputWrap.TFrame")
        self.mat_checks_frame.pack(side=LEFT, fill="x", expand=True)

        preset_row = ttk.Frame(self.mat_checks_frame, style="InputWrap.TFrame")
        preset_row.pack(fill="x", padx=self._px(8), pady=(0, self._px(4)))
        ttk.Label(preset_row, text="常用组合", style="CardHint.TLabel").pack(side=LEFT)
        self.material_preset_combo = ttk.Combobox(
            preset_row,
            textvariable=self.material_preset_name,
            values=self._material_preferences.preset_names,
            state="readonly",
            width=14,
            style="App.TCombobox",
        )
        self.material_preset_combo.pack(side=LEFT, padx=(self._px(8), self._px(6)))
        CodexButton(
            preset_row,
            text="应用",
            command=self._apply_material_preset,
            width=46,
            min_width=42,
            variant="link",
        ).pack(side=LEFT)

        preset_actions_row = ttk.Frame(self.mat_checks_frame, style="InputWrap.TFrame")
        preset_actions_row.pack(fill="x", padx=self._px(8), pady=(0, self._px(4)))
        ttk.Label(preset_actions_row, text="自定义预设", style="CardHint.TLabel").pack(side=LEFT)
        CodexButton(
            preset_actions_row,
            text="保存当前为预设",
            command=self._request_create_material_preset,
            width=98,
            min_width=90,
            variant="link",
        ).pack(side=LEFT, padx=(self._px(8), 0))
        CodexButton(
            preset_actions_row,
            text="更新",
            command=self._request_update_material_preset,
            width=46,
            min_width=42,
            variant="link",
        ).pack(side=LEFT, padx=(self._px(4), 0))
        CodexButton(
            preset_actions_row,
            text="重命名",
            command=self._request_rename_material_preset,
            width=56,
            min_width=52,
            variant="link",
        ).pack(side=LEFT, padx=(self._px(4), 0))
        CodexButton(
            preset_actions_row,
            text="删除",
            command=self._request_delete_material_preset,
            width=46,
            min_width=42,
            variant="link",
        ).pack(side=LEFT, padx=(self._px(4), 0))

        self.material_checks_grid = ttk.Frame(self.mat_checks_frame, style="InputWrap.TFrame")
        self.material_checks_grid.pack(fill="x")
        self._material_check_widgets: list = []
        self._rebuild_material_checkboxes()

        material_catalog_row = ttk.Frame(self.mat_checks_frame, style="InputWrap.TFrame")
        material_catalog_row.pack(fill="x", padx=self._px(8), pady=(self._px(2), 0))
        ttk.Label(material_catalog_row, text="自定义材料", style="CardHint.TLabel").pack(side=LEFT)
        self.material_custom_combo = ttk.Combobox(
            material_catalog_row,
            textvariable=self.material_custom_choice,
            values=self._material_preferences.custom_materials,
            state="readonly",
            width=14,
            style="App.TCombobox",
        )
        self.material_custom_combo.pack(side=LEFT, padx=(self._px(8), self._px(6)))
        CodexButton(
            material_catalog_row,
            text="添加材料",
            command=self._request_add_custom_material,
            width=62,
            min_width=56,
            variant="link",
        ).pack(side=LEFT)
        CodexButton(
            material_catalog_row,
            text="删除材料",
            command=self._request_delete_custom_material,
            width=62,
            min_width=56,
            variant="link",
        ).pack(side=LEFT, padx=(self._px(4), self._px(8)))
        self.material_catalog_hint = ttk.Label(
            self.mat_checks_frame,
            text="自定义材料和预设会保存在本机；应用组合后仍可继续增减勾选。",
            style="CardHint.TLabel",
        )
        self.material_catalog_hint.pack(fill="x", padx=self._px(8), pady=(self._px(2), 0))

        # 初始状态同步
        self._on_material_library_mode_changed()
        self._on_material_collect_all_changed()

        # ──────── 区块 1：统计日期范围 ────────
        # 周报与月报日期范围用对称两行布局；每行 label 放左侧，右侧为「起始日 至 结束日」+ 快捷按钮
        self.stats_range_label = ttk.Label(form, text="周报", style="App.TLabel")
        self.stats_range_frame = ttk.Frame(form, style="InputWrap.TFrame")
        self.stats_week_inputs = ttk.Frame(self.stats_range_frame, style="InputWrap.TFrame")
        self.stats_week_inputs.pack(side="top", fill="x")
        self.stats_week_date_group = ttk.Frame(self.stats_week_inputs, style="InputWrap.TFrame")
        self.stats_week_date_group.pack(side=LEFT)
        self.stats_week_start_entry = ttk.Entry(
            self.stats_week_date_group, textvariable=self.stats_week_start, width=12, style="App.TEntry"
        )
        self.stats_week_start_entry.pack(side=LEFT)
        ttk.Label(self.stats_week_date_group, text="至", style="App.TLabel").pack(side=LEFT, padx=self._px(8))
        self.stats_week_end_entry = ttk.Entry(
            self.stats_week_date_group, textvariable=self.stats_week_end, width=12, style="App.TEntry"
        )
        self.stats_week_end_entry.pack(side=LEFT)
        self.stats_week_hint = ttk.Label(
            self.stats_week_inputs, text="如 2026-06-02，留空按整月统计", style="CardHint.TLabel"
        )
        self.stats_week_hint.pack(
            side=LEFT, padx=self._pad(10, 0)
        )
        self.stats_week_presets = ttk.Frame(self.stats_range_frame, style="InputWrap.TFrame")
        self.stats_week_presets.pack(side="top", fill="x", pady=self._pad(6, 0))
        self.stats_week_preset_buttons = []
        for preset_text, preset_key in (
            ("本月", "this_month"),
            ("上月", "last_month"),
            ("本周", "this_week"),
            ("上周", "last_week"),
            ("清空", "clear"),
        ):
            button = CodexButton(
                self.stats_week_presets,
                text=preset_text,
                command=lambda key=preset_key: self._fill_stats_week_range(key),
                width=56,
                min_width=56,
                height=28,
            )
            button.pack(side=LEFT, padx=self._pad(0, 8))
            self.stats_week_preset_buttons.append(button)

        # —— 月报日期行 ——
        self.stats_month_range_label = ttk.Label(form, text="月报", style="App.TLabel")
        self.stats_month_range_frame = ttk.Frame(form, style="InputWrap.TFrame")
        self.stats_month_inputs = ttk.Frame(self.stats_month_range_frame, style="InputWrap.TFrame")
        self.stats_month_inputs.pack(side="top", fill="x")
        self.stats_month_date_group = ttk.Frame(self.stats_month_inputs, style="InputWrap.TFrame")
        self.stats_month_date_group.pack(side=LEFT)
        self.stats_month_start_entry = ttk.Entry(
            self.stats_month_date_group, textvariable=self.stats_month_start, width=12, style="App.TEntry"
        )
        self.stats_month_start_entry.pack(side=LEFT)
        ttk.Label(self.stats_month_date_group, text="至", style="App.TLabel").pack(side=LEFT, padx=self._px(8))
        self.stats_month_end_entry = ttk.Entry(
            self.stats_month_date_group, textvariable=self.stats_month_end, width=12, style="App.TEntry"
        )
        self.stats_month_end_entry.pack(side=LEFT)
        self.stats_month_hint = ttk.Label(
            self.stats_month_inputs, text="如 2026-06-01，留空不筛选月报", style="CardHint.TLabel"
        )
        self.stats_month_hint.pack(
            side=LEFT, padx=self._pad(10, 0)
        )
        self.stats_month_presets = ttk.Frame(self.stats_month_range_frame, style="InputWrap.TFrame")
        self.stats_month_presets.pack(side="top", fill="x", pady=self._pad(6, 0))
        self.stats_month_preset_buttons = []
        for preset_text, preset_key in (("本月", "this_month"), ("上月", "last_month"), ("清空", "clear")):
            button = CodexButton(
                self.stats_month_presets,
                text=preset_text,
                command=lambda key=preset_key: self._fill_stats_month_range(key),
                width=56,
                min_width=56,
                height=28,
            )
            button.pack(side=LEFT, padx=self._pad(0, 8))
            self.stats_month_preset_buttons.append(button)

        # ──────── 区块 2：输出选项 ────────
        # 与上方日期范围独立分组；三个输出选项并列展示
        self.stats_options_label = ttk.Label(form, text="输出选项", style="App.TLabel")
        self.stats_options_frame = ttk.Frame(form, style="InputWrap.TFrame")
        self.stats_options_grid = ttk.Frame(self.stats_options_frame, style="InputWrap.TFrame")
        self.stats_options_grid.pack(side="top", fill="x")

        # 第 1 列：加班/调休备注单位（单选）
        self.stats_unit_col = ttk.Frame(self.stats_options_grid, style="InputWrap.TFrame")
        self.stats_unit_col.pack(side=LEFT, fill="x", expand=True)
        ttk.Label(self.stats_unit_col, text="加班/调休备注", style="App.TLabel").pack(side=LEFT)
        ttk.Radiobutton(
            self.stats_unit_col, text="按天", value="day",
            variable=self.stats_remark_unit, style="App.TRadiobutton",
        ).pack(side=LEFT, padx=self._pad(12, 0))
        ttk.Radiobutton(
            self.stats_unit_col, text="按小时", value="hour",
            variable=self.stats_remark_unit, style="App.TRadiobutton",
        ).pack(side=LEFT, padx=self._pad(8, 0))
        self._stats_unit_help = ttk.Label(self.stats_unit_col, text=" ⓘ ", style="App.TLabel", cursor="question_arrow")
        self._stats_unit_help.pack(side=LEFT, padx=self._pad(6, 0))
        self._stats_unit_help.bind(
            "<Enter>", lambda _e: self._show_tooltip(
                self._stats_unit_help, "仅影响考勤统计表备注中加班/调休的显示"
            )
        )
        self._stats_unit_help.bind("<Leave>", lambda _e: self._hide_tooltip())

        # 第 2 列：是否新增公出列（勾选）
        self.stats_out_col = ttk.Frame(self.stats_options_grid, style="InputWrap.TFrame")
        self.stats_out_col.pack(side=LEFT, fill="x", expand=True)
        self.stats_business_trip_check = ttk.Checkbutton(
            self.stats_out_col, text="新增「公出」列",
            variable=self.stats_include_business_trip, style="App.TCheckbutton",
        )
        self.stats_business_trip_check.pack(side=LEFT)
        self._stats_out_help = ttk.Label(self.stats_out_col, text=" ⓘ ", style="App.TLabel", cursor="question_arrow")
        self._stats_out_help.pack(side=LEFT, padx=self._pad(6, 0))
        self._stats_out_help.bind(
            "<Enter>", lambda _e: self._show_tooltip(
                self._stats_out_help,
                "默认不勾选；勾选后在调休之后插入「公出（天）」列",
            )
        )
        self._stats_out_help.bind("<Leave>", lambda _e: self._hide_tooltip())

        # 第 3 列：是否新增出差列（勾选）
        self.stats_trip_col = ttk.Frame(self.stats_options_grid, style="InputWrap.TFrame")
        self.stats_trip_col.pack(side=LEFT, fill="x", expand=True)
        self.stats_workday_business_trip_check = ttk.Checkbutton(
            self.stats_trip_col, text="新增「出差」列",
            variable=self.stats_include_workday_business_trip, style="App.TCheckbutton",
        )
        self.stats_workday_business_trip_check.pack(side=LEFT)
        self._stats_trip_help = ttk.Label(self.stats_trip_col, text=" ⓘ ", style="App.TLabel", cursor="question_arrow")
        self._stats_trip_help.pack(side=LEFT, padx=self._pad(6, 0))
        self._stats_trip_help.bind(
            "<Enter>", lambda _e: self._show_tooltip(
                self._stats_trip_help,
                "仅统计源表中的工作日出差，不含休息日出差天数；与公出分别统计",
            )
        )
        self._stats_trip_help.bind("<Leave>", lambda _e: self._hide_tooltip())

        def _refresh_picker_button_bar(button_bar) -> None:
            visible_buttons = [child for child in button_bar.winfo_children() if getattr(child, "_hr_picker_visible", False)]
            state_key = (self._form_compact_layout, tuple(id(b) for b in visible_buttons))
            if getattr(button_bar, "_last_button_bar_state", None) == state_key:
                return
            setattr(button_bar, "_last_button_bar_state", state_key)
            for child in button_bar.winfo_children():
                child.pack_forget()
            if self._form_compact_layout:
                for index, child in enumerate(visible_buttons):
                    pady = self._pad(0, 6) if index < len(visible_buttons) - 1 else 0
                    child.pack(fill="x", pady=pady)
                return
            for child in reversed(visible_buttons):
                child.pack(side=RIGHT, padx=self._pad(4, 0))

        def _set_picker_button_visible(button, visible: bool) -> None:
            setattr(button, "_hr_picker_visible", visible)
            parent = button.master
            if visible:
                _refresh_picker_button_bar(parent)
                return
            button.pack_forget()
            _refresh_picker_button_bar(parent)

        def _layout_input_frame(row_data) -> None:
            entry = row_data["entry"]
            button_bar = row_data["button_bar"]
            last_compact = row_data.get("_last_compact")
            if last_compact == self._form_compact_layout:
                _refresh_picker_button_bar(button_bar)
                return
            row_data["_last_compact"] = self._form_compact_layout
            entry.pack_forget()
            button_bar.pack_forget()
            if self._form_compact_layout:
                entry.pack(side="top", fill="x")
                button_bar.pack(side="top", fill="x", pady=self._pad(8, 0))
                _refresh_picker_button_bar(button_bar)
                return
            entry.pack(side=LEFT, fill=BOTH, expand=True)
            button_bar.pack(side=RIGHT)
            _refresh_picker_button_bar(button_bar)

        self._last_applied_form_layout_key = None

        def _apply_form_layout() -> None:
            visible_keys = []
            if self._summary_row_visible:
                visible_keys.append("summary")
            if self._output_row_visible:
                visible_keys.append("output")

            material_vis = getattr(self, "_material_row_visible", False)
            layout_key = (
                self._form_layout_mode,
                tuple(visible_keys),
                self._rename_row_visible,
                material_vis,
                self._stats_range_row_visible,
            )
            if self._last_applied_form_layout_key == layout_key:
                return
            self._last_applied_form_layout_key = layout_key

            form.columnconfigure(
                0,
                weight=1 if self._form_compact_layout else 0,
                minsize=0 if self._form_compact_layout else self._px(116),
            )
            form.columnconfigure(1, weight=0 if self._form_compact_layout else 1)

            for key, row_data in self._form_rows.items():
                label = row_data["label"]
                frame = row_data["frame"]
                _layout_input_frame(row_data)
                if key not in visible_keys:
                    label.grid_remove()
                    frame.grid_remove()
                    continue
                if self._form_compact_layout:
                    display_index = visible_keys.index(key)
                    label.grid(
                        row=display_index * 2,
                        column=0,
                        columnspan=2,
                        sticky="w",
                        padx=0,
                        pady=self._pad(4, 2),
                    )
                    frame.grid(
                        row=display_index * 2 + 1,
                        column=0,
                        columnspan=2,
                        sticky="ew",
                        padx=0,
                        pady=self._pad(0, 8),
                    )
                    continue
                frame_padx = self._pad(12, 0)
                label.grid(row=row_data["index"], column=0, sticky="w", padx=0, pady=self._px(7))
                frame.grid(row=row_data["index"], column=1, sticky="ew", padx=frame_padx, pady=self._px(7))

            if self._rename_row_visible:
                rename_row = len(visible_keys) * 2 if self._form_compact_layout else 3
                self.rename_options_frame.grid(row=rename_row, column=0, columnspan=2, sticky="ew", pady=self._pad(10, 0))
            else:
                self.rename_options_frame.grid_remove()

            if material_vis:
                material_row = len(visible_keys) * 2 if self._form_compact_layout else 3
                self.material_options_frame.grid(row=material_row, column=0, columnspan=2, sticky="ew", pady=self._pad(10, 0))
            else:
                self.material_options_frame.grid_remove()

            if self._stats_range_row_visible:
                if self._form_compact_layout:
                    base_row = len(visible_keys) * 2
                    self.stats_range_label.grid(
                        row=base_row, column=0, sticky="w", padx=0, pady=self._pad(4, 2)
                    )
                    self.stats_range_frame.grid(
                        row=base_row + 1, column=0, columnspan=2, sticky="ew", padx=0, pady=self._pad(0, 8)
                    )
                    self.stats_month_range_label.grid(
                        row=base_row + 2, column=0, sticky="w", padx=0, pady=self._pad(4, 2)
                    )
                    self.stats_month_range_frame.grid(
                        row=base_row + 3, column=0, columnspan=2, sticky="ew", padx=0, pady=self._pad(0, 8)
                    )
                    self.stats_options_label.grid(
                        row=base_row + 4, column=0, sticky="w", padx=0, pady=self._pad(4, 2)
                    )
                    self.stats_options_frame.grid(
                        row=base_row + 5, column=0, columnspan=2, sticky="ew", padx=0, pady=self._pad(0, 8)
                    )
                else:
                    self.stats_range_label.grid(row=4, column=0, sticky="w", padx=0, pady=self._px(7))
                    self.stats_range_frame.grid(
                        row=4, column=1, sticky="ew", padx=self._pad(12, 0), pady=self._px(7)
                    )
                    self.stats_month_range_label.grid(row=5, column=0, sticky="w", padx=0, pady=self._px(7))
                    self.stats_month_range_frame.grid(
                        row=5, column=1, sticky="ew", padx=self._pad(12, 0), pady=self._px(7)
                    )
                    self.stats_options_label.grid(row=6, column=0, sticky="w", padx=0, pady=self._px(7))
                    self.stats_options_frame.grid(
                        row=6, column=1, sticky="ew", padx=self._pad(12, 0), pady=self._px(7)
                    )
            else:
                self.stats_range_label.grid_remove()
                self.stats_range_frame.grid_remove()
                self.stats_month_range_label.grid_remove()
                self.stats_month_range_frame.grid_remove()
                self.stats_options_label.grid_remove()
                self.stats_options_frame.grid_remove()

            self._apply_tool_specific_responsive_layout()
            if hasattr(self, "_sync_right_canvas_window"):
                self.root.after_idle(self._sync_right_canvas_window)

        def _update_form_responsive_layout(_event=None) -> None:
            canvas_width = self._right_canvas.winfo_width()
            logical_canvas_width = (
                canvas_width / max(self.ui_scale, 1.0)
                if canvas_width > 1
                else self._logical_screen_width()
            )
            content_padding = self._responsive_content_padding(logical_canvas_width)
            form_padding = self._responsive_form_padding_units(logical_canvas_width)
            # 内容列限宽居中（对应设计稿 max-width:780 的主内容列）
            if canvas_width > 1:
                base_left, pad_top, base_right, pad_bottom = content_padding
                extra = max(0, (canvas_width - base_left - base_right - self._px(820)) // 2)
                content_padding = (base_left + extra, pad_top, base_right + extra, pad_bottom)
            layout_changed = False
            if getattr(self, "_right_content_padding", None) != content_padding:
                self._right_content_padding = content_padding
                right_frame.configure(padding=content_padding)
                layout_changed = True
            if getattr(self, "_form_padding", None) != form_padding:
                self._form_padding = form_padding
                self.form_card.set_padding(form_padding)
                layout_changed = True
            if canvas_width > 1:
                usable_width = (
                    canvas_width
                    - content_padding[0]
                    - content_padding[2]
                    - self._px(form_padding[0] + form_padding[2])
                ) / max(self.ui_scale, 1.0)
                layout_mode = _responsive_layout_mode(usable_width)
            else:
                layout_mode = LAYOUT_MODE_WIDE
            compact = layout_mode != LAYOUT_MODE_WIDE
            if layout_mode != self._form_layout_mode:
                self._form_layout_mode = layout_mode
                layout_changed = True
            if compact != self._form_compact_layout:
                self._form_compact_layout = compact
                layout_changed = True
            if layout_changed:
                _apply_form_layout()

        self._apply_form_layout = _apply_form_layout
        self._update_form_responsive_layout = _update_form_responsive_layout
        self._show_picker_button = lambda button: _set_picker_button_visible(button, True)
        self._hide_picker_button = lambda button: _set_picker_button_visible(button, False)
        self._right_content_padding = self._responsive_content_padding()
        self._form_padding = self._responsive_form_padding_units()
        self._right_canvas.bind("<Configure>", _update_form_responsive_layout, add="+")
        self.root.after_idle(_update_form_responsive_layout)
        self._update_change_tabs_visibility()
        self._update_change_picker_buttons()
        self._update_summary_controls(apply_layout=False)
        self._update_output_controls(apply_layout=False)
        self._update_rename_controls(apply_layout=False)
        self._update_material_controls(apply_layout=False)
        self._update_stats_range_controls(apply_layout=False)
        _apply_form_layout()

        actions = ttk.Frame(right_frame, style="Content.TFrame")
        actions.pack(fill="x", pady=self._pad(16, 16))
        run_button_box = ttk.Frame(actions, width=self._px(132), height=self._px(40), style="Content.TFrame")
        run_button_box.pack(side=LEFT)
        run_button_box.pack_propagate(False)
        self.run_button = CodexButton(run_button_box, textvariable=self.run_button_text, command=self._run_current_tool, variant="primary", min_width=132, height=40)
        self.run_button.pack(fill=BOTH, expand=True)
        self.open_button = CodexButton(actions, text="打开结果目录", command=self._open_output_dir, width=138, height=40)
        self.open_button.pack(side=LEFT, padx=self._pad(12, 0))
        last_run_box = ttk.Frame(actions, style="Content.TFrame")
        last_run_box.pack(side=RIGHT)
        self.last_run_state_label = Label(
            last_run_box,
            textvariable=self.last_run_state,
            bg=COLOR_BG,
            fg=COLOR_SUCCESS,
            font=self.small_font,
        )
        self.last_run_state_label.pack(side=RIGHT)
        Label(
            last_run_box,
            textvariable=self.last_run_text,
            bg=COLOR_BG,
            fg=COLOR_FAINT,
            font=self.small_font,
        ).pack(side=RIGHT)

        log_card = RoundedCard(right_frame, padding=(20, 15, 20, 15), fill_height=True, min_height=150)
        log_card.pack(fill=BOTH, expand=True)
        log_header = ttk.Frame(log_card.inner, style="InputWrap.TFrame")
        log_header.pack(fill="x", pady=self._pad(0, 8))
        ttk.Label(log_header, text="运行记录", style="CardTitle.TLabel").pack(side=LEFT)
        log_body = ttk.Frame(log_card.inner, style="InputWrap.TFrame")
        log_body.pack(fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(log_body, orient=VERTICAL)
        self.log_text = Text(
            log_body,
            height=6,
            wrap="word",
            yscrollcommand=scrollbar.set,
            bg=COLOR_LOG_BG,
            fg=COLOR_LOG_TEXT,
            relief="flat",
            bd=0,
            highlightthickness=0,
            insertbackground=COLOR_LOG_TEXT,
            padx=self._px(2),
            pady=self._px(4),
            font=self.base_font,
            spacing3=self._px(8),
        )
        self.log_text.tag_configure("success", foreground=COLOR_SUCCESS)
        self.log_text.tag_configure("warning", foreground=COLOR_WARNING)
        self.log_text.tag_configure("error", foreground=COLOR_DANGER)
        self.log_text.tag_configure("muted", foreground=COLOR_LOG_MUTED)
        self.log_text.tag_configure("timestamp", foreground=COLOR_LOG_MUTED, font=self.small_font)
        self.log_text.tag_configure("dot_success", foreground=COLOR_SUCCESS_DOT)
        self.log_text.tag_configure("dot_warning", foreground=COLOR_WARNING_DOT)
        self.log_text.tag_configure("dot_error", foreground=COLOR_DANGER)
        self.log_text.tag_configure("dot_muted", foreground=COLOR_DROP_BORDER)
        self.log_text.tag_configure("dot_primary", foreground=COLOR_PRIMARY)
        scrollbar.config(command=self.log_text.yview)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # 右下角低调的运行日志入口：图标形式，面向开发排查，不需要 HR 关注
        log_footer = ttk.Frame(right_frame, style="Content.TFrame")
        log_footer.pack(fill="x", pady=self._pad(4, 0))
        log_icon = Canvas(log_footer, width=self._px(26), height=self._px(26), bg=COLOR_BG, highlightthickness=0, bd=0, cursor="hand2")
        log_icon.pack(side=RIGHT)

        def _draw_log_icon(color: str) -> None:
            log_icon.delete("all")
            _paint_tool_icon(log_icon, "social_security", color, self._pxf(5.5), self._pxf(5.5), self._pxf(15), max(1.0, self._pxf(1.3)))

        _draw_log_icon(COLOR_DISABLED)
        log_icon.bind("<Button-1>", lambda _event: self._open_run_log())
        log_icon.bind("<Enter>", lambda _event: _draw_log_icon(COLOR_MUTED), add="+")
        log_icon.bind("<Leave>", lambda _event: _draw_log_icon(COLOR_DISABLED), add="+")
        self._bind_path_tooltip(log_icon, lambda: "查看运行日志")

        # Mousewheel on log_text:
        #  • log has scrollable content in that direction → scroll the log
        #  • log is at top/bottom (or too short) → scroll the outer canvas
        # Note: widget-level bind() takes priority over bind_all(), so we
        # must handle both cases explicitly here.
        def _on_log_mousewheel(event):
            top, bottom = self.log_text.yview()
            delta_units = _mousewheel_units(event)

            can_scroll_up   = (top > 0)
            can_scroll_down = (bottom < 1.0)
            if (delta_units < 0 and can_scroll_up) or (delta_units > 0 and can_scroll_down):
                self.log_text.yview_scroll(delta_units, "units")
            else:
                self._right_canvas.yview_scroll(delta_units, "units")
            return "break"

        def _scroll_log_text_pixels(delta_y: int) -> None:
            line_height = self._px(18)
            try:
                line_info = self.log_text.dlineinfo("@0,0")
                if line_info and line_info[3] > 0:
                    line_height = line_info[3]
            except Exception:
                pass
            units = int(round(-delta_y / max(line_height, 1)))
            if units == 0 and delta_y:
                units = -1 if delta_y > 0 else 1
            if units:
                self.log_text.yview_scroll(units, "units")

        def _on_log_touchpad(event):
            top, bottom = self.log_text.yview()
            _delta_x, delta_y = _touchpad_deltas(event)
            if not delta_y:
                return "break"
            can_scroll_up = top > 0
            can_scroll_down = bottom < 1.0
            if (delta_y > 0 and can_scroll_up) or (delta_y < 0 and can_scroll_down):
                _scroll_log_text_pixels(delta_y)
            else:
                _scroll_page_pixels(delta_y)
            return "break"

        # log_text has its own smart handler — keep it as a widget-level
        # binding so it takes priority over the SCROLL_TAG class binding.
        self.log_text.bind("<MouseWheel>", _on_log_mousewheel)
        _safe_bind_widget(self.log_text, "<TouchpadScroll>", _on_log_touchpad)
        self.log_text.bind("<Button-4>",   _on_log_mousewheel)
        self.log_text.bind("<Button-5>",   _on_log_mousewheel)

        # One-time full scan after all widgets are rendered.
        self._apply_content_scroll_tag = _apply_scroll_tag
        self.root.after_idle(lambda: _apply_scroll_tag(right_frame))
        self.root.after_idle(self._sync_right_canvas_window)

        self._write_log(self._initial_log_text())
        self.root.update_idletasks()
        _sync_right_canvas_window()
        # History view is built lazily on first use so its ~50+ widgets
        # don't participate in the initial layout/render storm.

    def _build_workspace_panel(self, main_area) -> None:
        self._workspace_resize_handle = Canvas(
            main_area,
            width=self._px(4),
            bg=COLOR_SIDEBAR_BORDER,
            highlightthickness=0,
            bd=0,
            cursor="sb_h_double_arrow",
        )
        self._workspace_resize_handle.grid(row=0, column=1, sticky="ns")
        self._workspace_resize_handle.bind("<ButtonPress-1>", self._start_workspace_resize)
        self._workspace_resize_handle.bind("<B1-Motion>", self._resize_workspace_panel)
        self._workspace_resize_handle.bind("<ButtonRelease-1>", self._finish_workspace_resize)

        self._workspace_panel = ttk.Frame(
            main_area,
            width=self._px(self._workspace_width_units),
            style="Workspace.TFrame",
        )
        self._workspace_panel.grid(row=0, column=2, sticky="nsew")
        self._workspace_panel.pack_propagate(False)

        self._workspace_expanded_body = ttk.Frame(
            self._workspace_panel,
            padding=self._pad(16, 18, 16, 14),
            style="Workspace.TFrame",
        )
        self._workspace_expanded_body.pack(fill=BOTH, expand=True)

        header = ttk.Frame(self._workspace_expanded_body, style="Workspace.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text="项目文件", style="WorkspaceTitle.TLabel").pack(side=LEFT)
        self.workspace_collapse_button = CodexButton(
            header,
            text="收起",
            command=self._toggle_workspace_panel,
            variant="link",
            width=56,
            min_width=48,
            height=26,
        )
        self.workspace_collapse_button.pack(side=RIGHT)
        self.workspace_trash_button = CodexButton(
            header,
            text="回收站",
            command=self._open_workspace_trash_dialog,
            variant="link",
            width=62,
            min_width=56,
            height=26,
        )
        self.workspace_trash_button.pack(side=RIGHT, padx=self._pad(0, 3))

        self.workspace_project_name_label = Label(
            self._workspace_expanded_body,
            textvariable=self.workspace_project_name,
            bg=COLOR_SURFACE,
            fg=COLOR_PRIMARY,
            font=(self.base_font[0], _font_size(9), "bold"),
            anchor="w",
            justify="left",
        )
        self.workspace_project_name_label.pack(fill="x", pady=self._pad(11, 0))
        self.workspace_project_path_label = Label(
            self._workspace_expanded_body,
            textvariable=self.workspace_project_path,
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.tiny_font,
            anchor="w",
            justify="left",
        )
        self.workspace_project_path_label.pack(fill="x", pady=self._pad(3, 10))

        project_actions = ttk.Frame(self._workspace_expanded_body, style="Workspace.TFrame")
        project_actions.pack(fill="x", pady=self._pad(0, 12))
        self.workspace_switch_button = CodexButton(
            project_actions,
            text="切换项目",
            command=self._show_workspace_project_menu,
            width=92,
            min_width=82,
            height=31,
        )
        self.workspace_switch_button.pack(side=LEFT)
        self.workspace_open_project_button = CodexButton(
            project_actions,
            text="打开文件夹",
            command=self._open_workspace_root,
            width=100,
            min_width=88,
            height=31,
        )
        self.workspace_open_project_button.pack(side=LEFT, padx=self._pad(7, 0))

        scope_frame = ttk.Frame(self._workspace_expanded_body, style="Workspace.TFrame")
        scope_frame.pack(fill="x", pady=self._pad(0, 9))
        self.workspace_scope_all_button = CodexButton(
            scope_frame,
            text="全部文件",
            command=lambda: self._set_workspace_scope(WORKSPACE_SCOPE_ALL),
            variant="tonal",
            min_width=92,
            height=30,
        )
        self.workspace_scope_all_button.pack(side=LEFT, fill="x", expand=True)
        self.workspace_scope_tool_button = CodexButton(
            scope_frame,
            text="当前功能",
            command=lambda: self._set_workspace_scope(WORKSPACE_SCOPE_TOOL),
            min_width=92,
            height=30,
        )
        self.workspace_scope_tool_button.pack(side=LEFT, fill="x", expand=True, padx=self._pad(6, 0))

        ttk.Label(
            self._workspace_expanded_body,
            text="按文件名查找",
            style="WorkspaceMuted.TLabel",
        ).pack(anchor="w", pady=self._pad(0, 4))
        self.workspace_search_entry = ttk.Entry(
            self._workspace_expanded_body,
            textvariable=self.workspace_search,
            style="App.TEntry",
        )
        self.workspace_search_entry.pack(fill="x")

        file_actions = ttk.Frame(self._workspace_expanded_body, style="Workspace.TFrame")
        file_actions.pack(fill="x", pady=self._pad(10, 9))
        self.workspace_add_button = CodexButton(
            file_actions,
            text="添加",
            command=self._show_workspace_add_menu,
            variant="tonal",
            width=78,
            min_width=70,
            height=31,
        )
        self.workspace_add_button.pack(side=LEFT)
        self.workspace_refresh_button = CodexButton(
            file_actions,
            text="刷新",
            command=self._refresh_workspace_tree,
            variant="link",
            width=58,
            min_width=52,
            height=30,
        )
        self.workspace_refresh_button.pack(side=RIGHT)

        ttk.Frame(self._workspace_expanded_body, height=self._px(1), style="CardSeparator.TFrame").pack(fill="x")

        tree_frame = ttk.Frame(self._workspace_expanded_body, style="Workspace.TFrame")
        tree_frame.pack(fill=BOTH, expand=True, pady=self._pad(9, 8))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.workspace_tree = ttk.Treeview(
            tree_frame,
            show="tree",
            selectmode="browse",
            style="Workspace.Treeview",
        )
        workspace_tree_scroll = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.workspace_tree.yview)
        self.workspace_tree.configure(yscrollcommand=workspace_tree_scroll.set)
        self.workspace_tree.grid(row=0, column=0, sticky="nsew")
        workspace_tree_scroll.grid(row=0, column=1, sticky="ns")
        self.workspace_tree.tag_configure("muted", foreground=COLOR_FAINT)
        self.workspace_tree.tag_configure("result", foreground=COLOR_PRIMARY)
        self.workspace_tree.bind("<<TreeviewOpen>>", self._on_workspace_tree_open)
        self.workspace_tree.bind("<<TreeviewSelect>>", self._on_workspace_tree_selected)
        self.workspace_tree.bind("<Double-1>", self._open_selected_workspace_item)

        self.workspace_empty_label = Label(
            tree_frame,
            textvariable=self.workspace_empty_text,
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.small_font,
            justify="center",
            anchor="center",
        )

        detail = ttk.Frame(
            self._workspace_expanded_body,
            padding=self._pad(10, 9),
            style="Card.TFrame",
        )
        detail.pack(fill="x")
        self.workspace_detail_title_label = Label(
            detail,
            textvariable=self.workspace_detail_title,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=self.section_font,
            anchor="w",
        )
        self.workspace_detail_title_label.pack(fill="x")
        self.workspace_detail_text_label = Label(
            detail,
            textvariable=self.workspace_detail_text,
            bg=COLOR_SURFACE,
            fg=COLOR_MUTED,
            font=self.tiny_font,
            justify="left",
            anchor="w",
        )
        self.workspace_detail_text_label.pack(fill="x", pady=self._pad(3, 7))
        detail_actions = ttk.Frame(detail, style="InputWrap.TFrame")
        detail_actions.pack(fill="x")
        self.workspace_open_item_button = CodexButton(
            detail_actions,
            text="打开",
            command=self._open_selected_workspace_item,
            variant="link",
            width=52,
            min_width=46,
            height=25,
        )
        self.workspace_open_item_button.pack(side=LEFT)
        self.workspace_reveal_item_button = CodexButton(
            detail_actions,
            text="定位",
            command=self._reveal_selected_workspace_item,
            variant="link",
            width=52,
            min_width=46,
            height=25,
        )
        self.workspace_reveal_item_button.pack(side=LEFT, padx=self._pad(4, 0))
        self.workspace_move_to_trash_button = CodexButton(
            detail_actions,
            text="移到回收站",
            command=self._move_selected_workspace_batch_to_trash,
            variant="link",
            width=88,
            min_width=82,
            height=25,
        )

        self._workspace_collapsed_body = ttk.Frame(
            self._workspace_panel,
            style="WorkspaceRail.TFrame",
        )
        workspace_rail_label = Label(
            self._workspace_collapsed_body,
            text="项\n目\n文\n件",
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_PRIMARY,
            font=(self.base_font[0], _font_size(9), "bold"),
            cursor="hand2",
            justify="center",
            takefocus=True,
        )
        workspace_rail_label.pack(fill=BOTH, expand=True, pady=self._pad(17, 0))
        workspace_rail_label.bind("<Button-1>", lambda _event: self._toggle_workspace_panel())
        workspace_rail_label.bind("<Return>", lambda _event: self._toggle_workspace_panel())
        workspace_rail_label.bind("<space>", lambda _event: self._toggle_workspace_panel())

        self.workspace_search.trace_add("write", lambda *_args: self._schedule_workspace_search())
        self._workspace_expanded_body.bind("<Configure>", self._update_workspace_text_wraps, add="+")
        main_area.bind("<Configure>", self._on_workspace_area_resize, add="+")
        self.root.bind("<Escape>", self._close_workspace_drawer, add="+")
        self.root.bind("<Button-1>", self._close_workspace_drawer_on_outside_click, add="+")
        self._set_workspace_detail(None)
        self.root.after_idle(self._apply_workspace_panel_mode)

    def _on_workspace_area_resize(self, _event=None) -> None:
        try:
            logical_width = self._workspace_main_area.winfo_width() / max(self.ui_scale, 1.0)
        except Exception:
            logical_width = WORKSPACE_DRAWER_BREAKPOINT
        small = logical_width < WORKSPACE_DRAWER_BREAKPOINT
        if small != self._workspace_small:
            self._workspace_small = small
            self._workspace_drawer_open = False
            self._apply_workspace_panel_mode()
        elif hasattr(self, "_workspace_panel"):
            self._apply_workspace_panel_mode()

    def _apply_workspace_panel_mode(self) -> None:
        if not hasattr(self, "_workspace_panel"):
            return
        small = self._workspace_small
        expanded = self._workspace_drawer_open if small else self._workspace_preferred_expanded
        try:
            available_units = self._workspace_main_area.winfo_width() / max(self.ui_scale, 1.0)
        except Exception:
            available_units = WORKSPACE_DRAWER_BREAKPOINT
        drawer_width_units = _responsive_drawer_width(
            available_units,
            self._workspace_width_units,
        )
        mode_key = (
            "place" if small else "grid",
            "expanded" if expanded else "collapsed",
            drawer_width_units if small and expanded else 0,
        )
        title_button = getattr(self, "workspace_title_button", None)
        if title_button is not None:
            if small:
                if title_button.winfo_manager() != "pack":
                    title_button.pack(
                        side=LEFT,
                        padx=self._pad(0, 8),
                        before=self.check_update_button,
                    )
            else:
                title_button.pack_forget()
            if hasattr(self, "_update_title_text_wraps"):
                self.root.after_idle(self._update_title_text_wraps)

        if getattr(self, "_workspace_panel_mode_key", None) == mode_key:
            self.root.after_idle(self._update_workspace_text_wraps)
            return
        self._workspace_panel_mode_key = mode_key

        self._workspace_panel.place_forget()
        self._workspace_panel.grid_remove()
        self._workspace_resize_handle.grid_remove()
        self._workspace_expanded_body.pack_forget()
        self._workspace_collapsed_body.pack_forget()

        if self._workspace_small and expanded:
            self.workspace_collapse_button.configure(text="关闭")
            self._workspace_panel.configure(width=self._px(drawer_width_units))
            self._workspace_panel.place(
                in_=self._workspace_main_area,
                relx=1.0,
                y=0,
                relheight=1.0,
                width=self._px(drawer_width_units),
                anchor="ne",
            )
            self._workspace_expanded_body.pack(fill=BOTH, expand=True)
            self._workspace_panel.lift()
        elif self._workspace_small:
            # 小屏关闭状态只保留标题区入口，不再用 46px 竖栏挤占工具内容。
            self.workspace_collapse_button.configure(text="关闭")
        else:
            self.workspace_collapse_button.configure(text="收起")
            width_units = self._workspace_width_units if expanded else WORKSPACE_COLLAPSED_WIDTH
            self._workspace_panel.configure(width=self._px(width_units))
            self._workspace_panel.grid(row=0, column=2, sticky="nsew")
            if expanded:
                self._workspace_resize_handle.grid(row=0, column=1, sticky="ns")
                self._workspace_expanded_body.pack(fill=BOTH, expand=True)
            else:
                self._workspace_collapsed_body.pack(fill=BOTH, expand=True)

        temporary_open = small and expanded
        if temporary_open and not self._workspace_panel_was_temporary_open:
            try:
                self._workspace_restore_focus = self.root.focus_get()
            except Exception:
                self._workspace_restore_focus = None
            self.root.after_idle(self.workspace_search_entry.focus_set)
        elif not temporary_open and self._workspace_panel_was_temporary_open:
            restore_focus = self._workspace_restore_focus
            self._workspace_restore_focus = None
            if restore_focus is not None:
                try:
                    self.root.after_idle(restore_focus.focus_set)
                except Exception:
                    pass
        self._workspace_panel_was_temporary_open = temporary_open
        self.root.after_idle(self._update_workspace_text_wraps)

    def _toggle_workspace_panel(self) -> None:
        if self._workspace_small:
            self._workspace_drawer_open = not self._workspace_drawer_open
        else:
            self._workspace_preferred_expanded = not self._workspace_preferred_expanded
            self._save_workspace_preferences()
        self._apply_workspace_panel_mode()

    def _close_workspace_drawer(self, _event=None) -> None:
        if self._workspace_small and self._workspace_drawer_open:
            self._workspace_drawer_open = False
            self._apply_workspace_panel_mode()

    def _close_workspace_drawer_on_outside_click(self, event=None) -> None:
        if not self._workspace_small or not self._workspace_drawer_open or event is None:
            return
        widget = getattr(event, "widget", None)
        while widget is not None:
            if widget is self._workspace_panel:
                return
            widget = getattr(widget, "master", None)
        self._close_workspace_drawer()

    def _start_workspace_resize(self, event) -> None:
        if self._workspace_small or not self._workspace_preferred_expanded:
            return
        self._workspace_resize_origin = (int(event.x_root), int(self._workspace_width_units))

    def _resize_workspace_panel(self, event) -> None:
        if self._workspace_resize_origin is None:
            return
        start_x, start_width = self._workspace_resize_origin
        logical_delta = (start_x - int(event.x_root)) / max(self.ui_scale, 1.0)
        next_width = int(round(start_width + logical_delta))
        self._workspace_width_units = max(WORKSPACE_MIN_WIDTH, min(WORKSPACE_MAX_WIDTH, next_width))
        self._workspace_panel.configure(width=self._px(self._workspace_width_units))
        self._update_workspace_text_wraps()

    def _finish_workspace_resize(self, _event=None) -> None:
        if self._workspace_resize_origin is None:
            return
        self._workspace_resize_origin = None
        self._save_workspace_preferences()

    def _update_workspace_text_wraps(self, _event=None) -> None:
        if not hasattr(self, "_workspace_panel"):
            return
        width = max(self._workspace_panel.winfo_width() - self._px(36), self._px(190))
        for label in (
            getattr(self, "workspace_project_name_label", None),
            getattr(self, "workspace_project_path_label", None),
            getattr(self, "workspace_detail_text_label", None),
        ):
            if label is not None:
                label.configure(wraplength=width)
        sidebar_width = max(self._responsive_sidebar_width() - self._px(48), self._px(120))
        if hasattr(self, "sidebar_project_name_label"):
            self.sidebar_project_name_label.configure(wraplength=sidebar_width)
        if hasattr(self, "sidebar_project_path_label"):
            self.sidebar_project_path_label.configure(wraplength=sidebar_width)

    @staticmethod
    def _workspace_project_identity(value) -> tuple[str, Path] | None:
        if value is None:
            return None
        if isinstance(value, (str, Path)):
            path = Path(value).expanduser()
            return path.name or "工作项目", path
        if isinstance(value, dict):
            raw_path = value.get("path") or value.get("root") or value.get("root_dir")
            if not raw_path:
                return None
            path = Path(raw_path).expanduser()
            return str(value.get("name") or path.name or "工作项目"), path
        workspace = getattr(value, "workspace", None)
        if workspace is not None:
            raw_path = getattr(workspace, "root", None)
            if raw_path is not None:
                path = Path(raw_path).expanduser()
                return str(getattr(workspace, "name", None) or path.name or "工作项目"), path
        raw_path = getattr(value, "path", None) or getattr(value, "root", None) or getattr(value, "root_dir", None)
        if raw_path is None:
            return None
        path = Path(raw_path).expanduser()
        return str(getattr(value, "name", None) or path.name or "工作项目"), path

    def _restore_workspace_project(self) -> None:
        if not getattr(self, "_is_alive", True):
            return
        # after_idle 可能和自动化测试或极快的用户操作交错；已有项目时不要
        # 再打开第二份只读实例，也不要覆盖刚刚切换完成的界面状态。
        if self.current_project_path is not None:
            self._refresh_workspace_tree()
        else:
            last_path = getattr(self, "_workspace_last_project_path", None)
            if last_path is not None and last_path.is_dir():
                self._open_workspace_project_path(last_path, quiet=True)
            else:
                self._refresh_workspace_tree()
        self.root.after_idle(self._on_startup_rendered)

    def _reload_recent_workspace_projects(self) -> None:
        refreshed: list[tuple[str, Path]] = []
        for name, path in self._workspace_recent_projects:
            if path.is_dir() and all(existing_path != path for _existing_name, existing_path in refreshed):
                refreshed.append((name or path.name or "工作项目", path))
        self._workspace_recent_projects = refreshed[:8]

    def _show_workspace_project_menu(self) -> None:
        menu = Menu(self.root, tearoff=0, font=self.base_font)
        menu.add_command(label="新建工作项目", command=self._create_workspace_project)
        menu.add_command(label="打开已有项目", command=self._open_workspace_project)
        self._reload_recent_workspace_projects()
        if self._workspace_recent_projects:
            menu.add_separator()
            for name, path in self._workspace_recent_projects:
                menu.add_command(
                    label=name,
                    command=lambda project_path=path: self._open_workspace_project_path(project_path),
                )
        self._popup_workspace_menu(menu, self.workspace_switch_button)

    def _show_recent_projects_menu(self) -> None:
        self._reload_recent_workspace_projects()
        menu = Menu(self.root, tearoff=0, font=self.base_font)
        if self._workspace_recent_projects:
            for name, path in self._workspace_recent_projects:
                menu.add_command(
                    label=name,
                    command=lambda project_path=path: self._open_workspace_project_path(project_path),
                )
        else:
            menu.add_command(label="还没有最近项目", state="disabled")
        self._popup_workspace_menu(menu, self.sidebar_recent_project_button)

    @staticmethod
    def _popup_workspace_menu(menu: Menu, anchor_widget) -> None:
        try:
            menu.tk_popup(anchor_widget.winfo_rootx(), anchor_widget.winfo_rooty() + anchor_widget.winfo_height())
        finally:
            menu.grab_release()

    def _workspace_write_in_progress(self) -> bool:
        return bool(getattr(self, "_workspace_write_tasks", {}))

    def _project_change_is_blocked(self) -> bool:
        return bool(
            self._tool_running
            or self._project_batch_by_token
            or self._workspace_write_in_progress()
            or getattr(self, "_workspace_recovery_blocked", False)
        )

    def _create_workspace_project(self) -> None:
        if self._project_change_is_blocked():
            messagebox.showwarning(
                "处理尚未结束",
                "请等待当前处理或资料导入安全结束后，再切换工作项目。",
                parent=self.root,
            )
            return
        store_class = getattr(self, "_project_store_class", None)
        if store_class is None:
            messagebox.showinfo(
                "项目功能准备中",
                "工作项目管理模块尚未连接。当前版本仍可打开已有文件夹查看，但暂不能安全创建项目。",
                parent=self.root,
            )
            return
        self._open_workspace_project_create_dialog()

    def _default_workspace_project_parent(self) -> Path:
        self._reload_recent_workspace_projects()
        if self._workspace_recent_projects:
            recent_parent = self._workspace_recent_projects[0][1].parent
            if recent_parent.is_dir():
                return recent_parent
        documents = Path.home() / "Documents"
        return documents if documents.is_dir() else Path.home()

    def _open_workspace_project_create_dialog(self) -> None:
        existing = getattr(self, "_project_create_window", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                pass

        dialog_width, _ = self._update_dialog_size(600, 0)
        window = Toplevel(self.root)
        self._project_create_window = window
        self._project_create_busy = False
        setattr(window, "_hr_ui_scale", self.ui_scale)
        window.withdraw()
        window.title("新建工作项目")
        window.configure(bg=COLOR_SURFACE)
        window.resizable(False, True)
        window.transient(self.root)

        self._project_create_name_var = StringVar(
            master=window,
            value=_default_workspace_project_name(),
        )
        self._project_create_parent_var = StringVar(
            master=window,
            value=str(self._default_workspace_project_parent()),
        )
        self._project_create_preview_var = StringVar(master=window)
        self._project_create_status_var = StringVar(master=window)
        self._project_create_trace_ids = []

        shell = Frame(window, bg=COLOR_SURFACE)
        shell.pack(fill=BOTH, expand=True)
        scroll = ttk.Scrollbar(shell, orient=VERTICAL)
        scroll.pack(side=RIGHT, fill=Y)
        canvas = Canvas(
            shell,
            bg=COLOR_SURFACE,
            highlightthickness=0,
            bd=0,
            yscrollcommand=scroll.set,
        )
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.configure(command=canvas.yview)

        content = Frame(canvas, bg=COLOR_SURFACE)
        content_window = canvas.create_window(0, 0, window=content, anchor="nw")

        def sync_scroll_region(_event=None) -> None:
            canvas.itemconfigure(content_window, width=max(canvas.winfo_width(), 1))
            bounds = canvas.bbox("all")
            if bounds is not None:
                canvas.configure(scrollregion=bounds)

        content.bind("<Configure>", sync_scroll_region)
        canvas.bind("<Configure>", sync_scroll_region)

        pad = self._px(32)
        wrap_width = max(dialog_width - self._px(76), self._px(320))

        Label(
            content,
            text="新建工作项目",
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=(self.base_font[0], _font_size(16), "bold"),
            anchor="w",
        ).pack(fill="x", padx=pad, pady=self._pad(28, 2))
        Label(
            content,
            text="项目是一套可随时打开、完整留存资料的工作文件夹。",
            bg=COLOR_SURFACE,
            fg=COLOR_MUTED,
            font=self.base_font,
            anchor="w",
        ).pack(fill="x", padx=pad, pady=self._pad(0, 20))

        Label(
            content,
            text="项目名称",
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=self.section_font,
            anchor="w",
        ).pack(fill="x", padx=pad, pady=self._pad(0, 6))
        name_entry = ttk.Entry(
            content,
            textvariable=self._project_create_name_var,
            style="App.TEntry",
        )
        name_entry.pack(fill="x", padx=pad)
        self._project_create_name_entry = name_entry
        Label(
            content,
            text="建议按月份、地区或业务事项命名，方便以后查找。",
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.small_font,
            anchor="w",
        ).pack(fill="x", padx=pad, pady=self._pad(5, 14))

        Label(
            content,
            text="保存位置",
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=self.section_font,
            anchor="w",
        ).pack(fill="x", padx=pad, pady=self._pad(0, 6))
        location_wrap = Frame(
            content,
            bg=COLOR_SURFACE,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_PRIMARY,
            highlightthickness=self._px(1),
            bd=0,
        )
        location_wrap.pack(fill="x", padx=pad)
        location_label = Label(
            location_wrap,
            textvariable=self._project_create_parent_var,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=self.base_font,
            anchor="w",
            justify="left",
            wraplength=max(wrap_width - self._px(150), self._px(160)),
        )
        location_label.pack(side=LEFT, fill="x", expand=True, padx=self._pad(12, 8), pady=self._pad(9))
        location_button = CodexButton(
            location_wrap,
            text="选择其他位置",
            command=self._choose_workspace_project_parent,
            width=116,
            min_width=108,
            height=34,
        )
        location_button.pack(side=RIGHT, padx=self._pad(0, 4), pady=self._pad(4))
        self._project_create_location_button = location_button
        Label(
            content,
            text="不建议选择整个桌面、磁盘根目录或个人主目录。",
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.small_font,
            anchor="w",
        ).pack(fill="x", padx=pad, pady=self._pad(5, 14))

        preview_card = Frame(content, bg=COLOR_PRIMARY_SOFT)
        preview_card.pack(fill="x", padx=pad, pady=self._pad(0, 14))
        Label(
            preview_card,
            text="项目最终位置",
            bg=COLOR_PRIMARY_SOFT,
            fg=COLOR_PRIMARY,
            font=(self.base_font[0], _font_size(9), "bold"),
            anchor="w",
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(11, 2))
        Label(
            preview_card,
            textvariable=self._project_create_preview_var,
            bg=COLOR_PRIMARY_SOFT,
            fg=COLOR_PRIMARY,
            font=self.base_font,
            anchor="w",
            justify="left",
            wraplength=wrap_width,
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(0, 11))

        retention_card = Frame(content, bg=COLOR_SURFACE_ALT)
        retention_card.pack(fill="x", padx=pad, pady=self._pad(0, 14))
        Label(
            retention_card,
            text="建立后会自动做到",
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_TEXT,
            font=self.section_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(11, 4))
        Label(
            retention_card,
            text="上传资料复制进项目，外部原件不动\n处理结果直接保存在项目里，右侧随时可查\n整个项目文件夹可以备份、迁移和交接",
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_MUTED,
            font=self.small_font,
            anchor="w",
            justify="left",
            wraplength=wrap_width,
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(0, 11))

        warning_card = Frame(content, bg=COLOR_WARNING_SOFT)
        warning_card.pack(fill="x", padx=pad)
        Label(
            warning_card,
            text="若保存到网盘或共享盘，项目资料可能被同步或被其他人访问。",
            bg=COLOR_WARNING_SOFT,
            fg=COLOR_WARNING,
            font=self.small_font,
            anchor="w",
            justify="left",
            wraplength=wrap_width,
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(11))

        status_label = Label(
            content,
            textvariable=self._project_create_status_var,
            bg=COLOR_SURFACE,
            fg=COLOR_DANGER,
            font=self.small_font,
            anchor="w",
            justify="left",
            wraplength=wrap_width,
        )
        status_label.pack(fill="x", padx=pad, pady=self._pad(8, 0))
        self._project_create_status_label = status_label

        button_row = Frame(content, bg=COLOR_SURFACE)
        button_row.pack(fill="x", padx=pad, pady=self._pad(14, 24))
        submit_button = CodexButton(
            button_row,
            text="创建并打开",
            command=self._submit_workspace_project_create,
            variant="primary",
            width=106,
            min_width=98,
            height=38,
        )
        submit_button.pack(side=RIGHT)
        self._project_create_submit_button = submit_button
        cancel_button = CodexButton(
            button_row,
            text="取消",
            command=self._close_workspace_project_create_dialog,
            width=70,
            min_width=64,
            height=38,
        )
        cancel_button.pack(side=RIGHT, padx=self._pad(0, 9))
        self._project_create_cancel_button = cancel_button

        def bind_keyboard_button(button: CodexButton) -> None:
            button.configure(takefocus=1)

            def activate(_event=None):
                button._on_click()
                return "break"

            button.bind("<Return>", activate)
            button.bind("<space>", activate)

        for button in (location_button, cancel_button, submit_button):
            bind_keyboard_button(button)

        for variable in (self._project_create_name_var, self._project_create_parent_var):
            trace_id = variable.trace_add("write", lambda *_args: self._refresh_workspace_project_create_form())
            self._project_create_trace_ids.append((variable, trace_id))

        def scroll_dialog(event) -> str | None:
            if content.winfo_reqheight() <= canvas.winfo_height():
                return None
            if getattr(event, "num", None) == 4:
                direction = -1
            elif getattr(event, "num", None) == 5:
                direction = 1
            else:
                delta = getattr(event, "delta", 0)
                direction = -1 if delta > 0 else 1
            canvas.yview_scroll(direction, "units")
            return "break"

        window.bind("<MouseWheel>", scroll_dialog)
        window.bind("<Button-4>", scroll_dialog)
        window.bind("<Button-5>", scroll_dialog)
        window.bind("<Return>", lambda _event: self._submit_workspace_project_create())
        window.bind("<Escape>", lambda _event: self._close_workspace_project_create_dialog())
        window.protocol("WM_DELETE_WINDOW", self._close_workspace_project_create_dialog)

        self._refresh_workspace_project_create_form()
        window.update_idletasks()
        requested_height = content.winfo_reqheight()
        max_height = max(self._px(420), self.root.winfo_screenheight() - self._px(72))
        min_height = min(self._px(560), max_height)
        dialog_height = max(min_height, min(requested_height, max_height))
        self._center_window(window, dialog_width, dialog_height)
        sync_scroll_region()
        window.deiconify()
        try:
            window.grab_set()
        except Exception:
            pass
        name_entry.focus_set()
        name_entry.selection_range(0, END)

    def _choose_workspace_project_parent(self) -> None:
        if getattr(self, "_project_create_busy", False):
            return
        window = getattr(self, "_project_create_window", None)
        parent_var = getattr(self, "_project_create_parent_var", None)
        if window is None or parent_var is None:
            return
        initial = Path(parent_var.get()).expanduser()
        if not initial.is_dir():
            initial = self._default_workspace_project_parent()
        selected = filedialog.askdirectory(
            title="选择工作项目的保存位置",
            initialdir=str(initial),
            mustexist=True,
            parent=window,
        )
        if selected:
            parent_var.set(str(Path(selected).expanduser().absolute()))
            self._remember_file_dialog_path(selected)
        try:
            window.lift()
            window.focus_force()
        except Exception:
            pass

    def _refresh_workspace_project_create_form(self) -> None:
        name_var = getattr(self, "_project_create_name_var", None)
        parent_var = getattr(self, "_project_create_parent_var", None)
        preview_var = getattr(self, "_project_create_preview_var", None)
        status_var = getattr(self, "_project_create_status_var", None)
        if name_var is None or parent_var is None or preview_var is None or status_var is None:
            return
        project_root, error = _workspace_project_creation_target(parent_var.get(), name_var.get())
        preview_var.set(str(project_root) if project_root is not None else "请先填写项目名称并选择保存位置")
        if getattr(self, "_project_create_busy", False):
            return
        status_var.set(error or "")
        status_label = getattr(self, "_project_create_status_label", None)
        if status_label is not None:
            status_label.configure(fg=COLOR_DANGER)
        submit_button = getattr(self, "_project_create_submit_button", None)
        if submit_button is not None:
            submit_button.configure(state="disabled" if error else "normal")

    def _set_workspace_project_create_busy(self, busy: bool) -> None:
        self._project_create_busy = bool(busy)
        state = "disabled" if busy else "normal"
        name_entry = getattr(self, "_project_create_name_entry", None)
        if name_entry is not None:
            name_entry.configure(state=state)
        for attribute in ("_project_create_location_button", "_project_create_cancel_button"):
            button = getattr(self, attribute, None)
            if button is not None:
                button.configure(state=state)
        submit_button = getattr(self, "_project_create_submit_button", None)
        if submit_button is not None:
            submit_button.configure(state="disabled" if busy else "normal")
        if not busy:
            self._refresh_workspace_project_create_form()

    def _set_workspace_project_create_status(self, message: str, *, error: bool) -> None:
        status_var = getattr(self, "_project_create_status_var", None)
        if status_var is not None:
            status_var.set(message)
        status_label = getattr(self, "_project_create_status_label", None)
        if status_label is not None:
            status_label.configure(fg=COLOR_DANGER if error else COLOR_PRIMARY)

    def _submit_workspace_project_create(self) -> None:
        if getattr(self, "_project_create_busy", False):
            return
        name_var = getattr(self, "_project_create_name_var", None)
        parent_var = getattr(self, "_project_create_parent_var", None)
        if name_var is None or parent_var is None:
            return
        project_name = name_var.get().strip()
        project_root, error = _workspace_project_creation_target(parent_var.get(), project_name)
        if error or project_root is None:
            self._refresh_workspace_project_create_form()
            return

        store_class = getattr(self, "_project_store_class", None)
        creator = getattr(store_class, "create", None) if store_class is not None else None
        if not callable(creator):
            self._set_workspace_project_create_status("工作项目管理模块暂时不可用。", error=True)
            return

        self._set_workspace_project_create_busy(True)
        self._set_workspace_project_create_status("正在创建项目，请稍候…", error=False)
        window = getattr(self, "_project_create_window", None)
        if window is not None:
            try:
                window.update_idletasks()
            except Exception:
                pass
        try:
            project = creator(project_root, project_name)
        except Exception as exc:
            runlog.log_exception("新建工作项目失败", exc)
            self._set_workspace_project_create_busy(False)
            self._set_workspace_project_create_status(_workspace_project_create_error_message(exc), error=True)
            return

        identity = self._workspace_project_identity(project) or (project_name, project_root)
        workspace = getattr(project, "workspace", None)
        read_only = not bool(getattr(workspace, "writable", True))
        try:
            self._set_workspace_project(*identity, read_only=read_only, store=project)
        except Exception as exc:
            runlog.log_exception("新建项目后打开失败", exc)
            closer = getattr(project, "close", None)
            if callable(closer):
                try:
                    closer()
                except Exception:
                    pass
            self._set_workspace_project_create_busy(False)
            self._set_workspace_project_create_status(
                "项目已经创建，但暂时无法打开。请取消后使用“打开项目”选择该文件夹。",
                error=True,
            )
            return
        self._close_workspace_project_create_dialog(force=True)

    def _close_workspace_project_create_dialog(self, *, force: bool = False) -> None:
        if getattr(self, "_project_create_busy", False) and not force:
            return
        window = getattr(self, "_project_create_window", None)
        for variable, trace_id in getattr(self, "_project_create_trace_ids", []):
            try:
                variable.trace_remove("write", trace_id)
            except Exception:
                pass
        self._project_create_trace_ids = []
        self._project_create_window = None
        self._project_create_busy = False
        self._project_create_name_var = None
        self._project_create_parent_var = None
        self._project_create_preview_var = None
        self._project_create_status_var = None
        self._project_create_name_entry = None
        self._project_create_location_button = None
        self._project_create_cancel_button = None
        self._project_create_submit_button = None
        self._project_create_status_label = None
        if window is not None:
            try:
                window.grab_release()
            except Exception:
                pass
            try:
                window.destroy()
            except Exception:
                pass

    def _open_workspace_project(self) -> None:
        if self._project_change_is_blocked():
            messagebox.showwarning(
                "处理尚未结束",
                "请等待当前处理或资料导入安全结束后，再切换工作项目。",
                parent=self.root,
            )
            return
        directory = self._askdirectory(title="打开已有工作项目")
        if directory:
            self._open_workspace_project_path(Path(directory))

    def _open_workspace_project_path(self, path: Path, *, quiet: bool = False) -> None:
        if self._project_change_is_blocked():
            if not quiet:
                messagebox.showwarning(
                    "处理尚未结束",
                    "请等待当前处理或资料导入安全结束后，再切换工作项目。",
                    parent=self.root,
                )
            return
        path = Path(path).expanduser().absolute()
        if not path.is_dir():
            if not quiet:
                messagebox.showwarning("项目不存在", "这个工作项目可能已经移动，请重新选择项目文件夹。", parent=self.root)
            return
        if self.current_project_path is not None:
            try:
                is_current = path.samefile(self.current_project_path)
            except OSError:
                is_current = path == self.current_project_path
            if is_current:
                if self.project_store is not None:
                    try:
                        self.project_store.refresh()
                    except Exception as exc:
                        runlog.log_exception("刷新当前工作项目失败", exc)
                self._refresh_workspace_tree()
                self._update_sidebar_project_summary()
                return
        store_class = getattr(self, "_project_store_class", None)
        opener = getattr(store_class, "open", None) if store_class is not None else None
        if callable(opener):
            try:
                project = opener(path)
            except Exception as exc:
                runlog.log_exception("打开工作项目失败", exc)
                if not quiet:
                    messagebox.showerror("无法打开项目", f"这个文件夹无法作为工作项目打开。\n\n原因：{exc}", parent=self.root)
                return
            identity = self._workspace_project_identity(project) or (path.name or "工作项目", path)
            workspace = getattr(project, "workspace", None)
            read_only = not bool(getattr(workspace, "writable", True))
            self._set_workspace_project(*identity, read_only=read_only, store=project)
            return
        # ProjectStore 尚未接入时保持只读，确保工作区外壳仍可启动和浏览。
        self._set_workspace_project(path.name or "工作项目", path, read_only=True, store=None)

    def _set_workspace_project(self, name: str, path: Path, *, read_only: bool, store=None) -> None:
        old_store = self.project_store
        old_path = self.current_project_path
        new_path = Path(path).expanduser().absolute()
        project_changed = old_path is not None and old_path != new_path
        if project_changed:
            self._close_workspace_trash_dialog()
        if old_store is not None and old_store is not store:
            closer = getattr(old_store, "close", None)
            if callable(closer):
                try:
                    closer()
                except Exception as exc:
                    runlog.log_exception("关闭旧工作项目失败", exc)
        self.project_store = store
        self.current_project_path = new_path
        self._workspace_project_generation += 1
        if project_changed:
            self.input_path.set("")
            self.summary_path.set("")
            self.change_input_paths = None
            self.change_form_state = {
                "merge": ("", "", None),
                "roster": ("", "", None),
            }
            self.archive_form_state = {
                "import": ("", "", None),
                "export": ("", "", None),
            }
            self.last_output_dir = None
            if hasattr(self, "upload_body"):
                self._refresh_upload_card()
        self._workspace_project_read_only = bool(read_only)
        display_name = str(name or self.current_project_path.name or "工作项目")
        self.workspace_project_name.set(display_name + (" · 只读" if read_only else ""))
        workspace = getattr(store, "workspace", None)
        read_only_reason = getattr(workspace, "read_only_reason", None) if workspace is not None else None
        path_text = str(self.current_project_path)
        if read_only_reason:
            path_text = f"{path_text}\n{read_only_reason}"
        self.workspace_project_path.set(path_text)
        self.workspace_search.set("")
        self._workspace_search_generation += 1
        self._set_workspace_detail(None)
        self._workspace_recent_projects = [
            (display_name, self.current_project_path),
            *((item_name, item_path) for item_name, item_path in self._workspace_recent_projects if item_path != self.current_project_path),
        ][:8]
        self._workspace_last_project_path = self.current_project_path
        self._update_sidebar_project_summary()
        self._save_workspace_preferences()
        self._refresh_workspace_tree()
        self._update_workspace_action_states()
        self._update_project_output_controls()

    def _update_sidebar_project_summary(self) -> None:
        store = self.project_store
        if self.current_project_path is None or store is None:
            self.sidebar_project_summary.set("新建或打开项目后开始处理")
            return
        try:
            batch_count = len(store.list_batches())
        except Exception:
            self.sidebar_project_summary.set("当前项目 · 资料可在右侧查看")
            return
        if batch_count:
            self.sidebar_project_summary.set(f"当前项目 · {batch_count} 次处理")
        else:
            self.sidebar_project_summary.set("当前项目 · 尚无处理记录")

    def _open_workspace_root(self) -> None:
        project_path = self.current_project_path
        if project_path is None:
            messagebox.showinfo("还没有工作项目", "请先新建或打开一个工作项目。", parent=self.root)
            return
        try:
            open_path(project_path)
        except Exception as exc:
            runlog.log_exception("打开项目文件夹失败", exc)
            messagebox.showerror("无法打开文件夹", f"项目文件夹无法打开。\n\n原因：{exc}", parent=self.root)

    def _show_workspace_backend_unavailable(self, action: str) -> None:
        detail = f"\n\n原因：{self._project_store_error}" if self._project_store_error else ""
        messagebox.showinfo(
            "项目功能准备中",
            f"“{action}”需要工作项目管理模块。模块连接完成后即可使用。{detail}",
            parent=self.root,
        )

    def _set_workspace_scope(self, scope: str) -> None:
        if scope not in {WORKSPACE_SCOPE_ALL, WORKSPACE_SCOPE_TOOL}:
            return
        self.workspace_scope.set(scope)
        self.workspace_scope_all_button.configure(variant="tonal" if scope == WORKSPACE_SCOPE_ALL else "secondary")
        self.workspace_scope_tool_button.configure(variant="tonal" if scope == WORKSPACE_SCOPE_TOOL else "secondary")
        self._refresh_workspace_tree()

    def _project_tool_identity(self) -> tuple[str, str]:
        """Return the stable batch id and HR-facing folder name for the active sub-tool."""

        if self.current_tool == "personnel_change_merge" and self.change_mode == "roster":
            return "roster_update", "花名册更新"
        if self.current_tool == "archive_import" and self.archive_mode == "export":
            return "archive_export", "档案表生成"
        return self.current_tool, TOOL_NAV_LABELS.get(self.current_tool, self._tool_log_label())

    def _workspace_tool_parts(self) -> tuple[str, ...]:
        _tool_id, tool_name = self._project_tool_identity()
        group_name = TOOL_GROUP_LABELS.get(self.current_tool, "人员运营自动化")
        return group_name, tool_name

    def _workspace_scope_root(self) -> Path | None:
        project_path = self.current_project_path
        if project_path is None:
            return None
        if self.workspace_scope.get() != WORKSPACE_SCOPE_TOOL:
            return project_path
        # 项目目录的显示名称来自每次处理时的业务名称，不应只依赖界面里的
        # 固定中文映射。优先从真实批次反查当前工具目录，老项目改名后也能找到。
        store = self.project_store
        if store is not None:
            try:
                project_tool_id, _tool_name = self._project_tool_identity()
                summaries = store.list_batches(tool_id=project_tool_id)
            except Exception as exc:
                runlog.log_exception("读取当前功能项目目录失败", exc)
            else:
                for summary in summaries:
                    try:
                        detail = store.get_batch(summary.id)
                    except Exception:
                        continue
                    if detail is None:
                        continue
                    for category in ("uploads", "supplements", "results"):
                        directory = detail.directories.get(category)
                        if directory is None:
                            continue
                        tool_root = directory.parent.parent
                        try:
                            tool_root.relative_to(project_path)
                        except ValueError:
                            continue
                        if tool_root.is_dir():
                            return tool_root
        return project_path.joinpath(*self._workspace_tool_parts())

    @staticmethod
    def _workspace_should_hide_path(path: Path) -> bool:
        name = path.name
        if name in WORKSPACE_HIDDEN_NAMES or name.startswith("~$"):
            return True
        if name.startswith("."):
            return True
        lower_name = name.lower()
        if any(lower_name.endswith(suffix) for suffix in WORKSPACE_HIDDEN_SUFFIXES):
            return True
        try:
            return path.is_symlink()
        except OSError:
            return True

    def _workspace_visible_children(self, path: Path) -> list[Path]:
        try:
            children = []
            for child in path.iterdir():
                if self._workspace_should_hide_path(child):
                    continue
                if child.name == "上传资料" and child.is_dir():
                    try:
                        if not any(child.iterdir()):
                            continue
                    except OSError:
                        pass
                children.append(child)
        except OSError:
            return []
        return sorted(children, key=lambda child: (not child.is_dir(), child.name.casefold()))

    def _refresh_workspace_tree(self) -> None:
        if not hasattr(self, "workspace_tree"):
            return
        self._workspace_search_generation += 1
        generation = self._workspace_search_generation
        for item in self.workspace_tree.get_children():
            self.workspace_tree.delete(item)
        self._workspace_tree_paths.clear()
        self._set_workspace_detail(None)
        root_path = self._workspace_scope_root()
        if self.current_project_path is None:
            self.workspace_empty_text.set("先新建或打开一个工作项目\n\n资料和结果会集中显示在这里")
            self._show_workspace_empty(True)
            self._update_workspace_action_states()
            return
        if root_path is None or not root_path.is_dir():
            if self.workspace_scope.get() == WORKSPACE_SCOPE_TOOL:
                self.workspace_empty_text.set("当前功能还没有处理记录\n\n请在中间选择资料并开始处理")
            else:
                self.workspace_empty_text.set("项目里还没有资料\n\n点击“添加”导入文件或文件夹")
            self._show_workspace_empty(True)
            self._update_workspace_action_states()
            return
        query = self.workspace_search.get().strip().casefold()
        if query:
            self.workspace_empty_text.set("正在查找项目文件…")
            self._show_workspace_empty(True)
            threading.Thread(
                target=self._search_workspace_files,
                args=(generation, root_path, query),
                daemon=True,
            ).start()
            self._update_workspace_action_states()
            return
        children = self._workspace_visible_children(root_path)
        for child in children:
            self._insert_workspace_tree_path("", child)
        if children:
            self._show_workspace_empty(False)
        else:
            label = "当前功能还没有项目文件" if self.workspace_scope.get() == WORKSPACE_SCOPE_TOOL else "这个项目文件夹目前为空"
            self.workspace_empty_text.set(f"{label}\n\n点击“添加”导入资料")
            self._show_workspace_empty(True)
        self._update_workspace_action_states()

    def _insert_workspace_tree_path(self, parent: str, path: Path, *, search_result: bool = False) -> str:
        text = path.name
        if search_result and self.current_project_path is not None:
            try:
                parent_text = path.parent.relative_to(self.current_project_path).as_posix()
            except ValueError:
                parent_text = path.parent.name
            if parent_text and parent_text != ".":
                text = f"{path.name}   ·   {parent_text}"
        tags: tuple[str, ...] = ()
        try:
            relative_parts = path.relative_to(self.current_project_path).parts if self.current_project_path else ()
        except ValueError:
            relative_parts = ()
        if "处理结果" in relative_parts:
            tags = ("result",)
        item = self.workspace_tree.insert(parent, "end", text=text, open=False, tags=tags)
        self._workspace_tree_paths[item] = path
        if path.is_dir() and not search_result:
            self.workspace_tree.insert(item, "end", text="", tags=(WORKSPACE_DUMMY_TAG,))
        return item

    def _on_workspace_tree_open(self, _event=None) -> None:
        item = self.workspace_tree.focus()
        path = self._workspace_tree_paths.get(item)
        if path is None or not path.is_dir():
            return
        children = self.workspace_tree.get_children(item)
        if len(children) != 1 or WORKSPACE_DUMMY_TAG not in self.workspace_tree.item(children[0], "tags"):
            return
        self.workspace_tree.delete(children[0])
        for child in self._workspace_visible_children(path):
            self._insert_workspace_tree_path(item, child)

    def _on_workspace_tree_selected(self, _event=None) -> None:
        selected = self.workspace_tree.selection()
        path = self._workspace_tree_paths.get(selected[0]) if selected else None
        self._set_workspace_detail(path)

    def _selected_workspace_path(self) -> Path | None:
        if not hasattr(self, "workspace_tree"):
            return None
        selected = self.workspace_tree.selection()
        return self._workspace_tree_paths.get(selected[0]) if selected else None

    def _workspace_batch_root_for_path(self, path: Path | None):
        store = self.project_store
        if path is None or store is None:
            return None
        try:
            summaries = tuple(store.list_batches())
        except Exception:
            return None
        for summary in summaries:
            try:
                detail = store.get_batch(summary.id)
            except Exception:
                continue
            if detail is None:
                continue
            upload_dir = detail.directories.get("uploads")
            if upload_dir is not None and Path(path) == Path(upload_dir).parent:
                return summary, detail
        return None

    def _set_workspace_detail(self, path: Path | None) -> None:
        if path is None or not path.exists():
            self.workspace_detail_title.set("选择文件查看详情")
            self.workspace_detail_text.set("双击可以打开文件；文件夹可展开查看。")
            self._update_workspace_action_states()
            return
        self.workspace_detail_title.set(path.name)
        try:
            relative = path.relative_to(self.current_project_path).as_posix() if self.current_project_path else path.name
        except ValueError:
            relative = path.name
        try:
            modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        except OSError:
            modified = "更新时间未知"
        if path.is_dir():
            detail = f"文件夹 · {relative}\n更新于 {modified}"
        else:
            detail = f"{self._format_file_size(path)} · 更新于 {modified}\n{relative}"
        self.workspace_detail_text.set(detail)
        self._update_workspace_action_states()

    def _update_workspace_action_states(self) -> None:
        if not hasattr(self, "workspace_add_button"):
            return
        has_project = self.current_project_path is not None
        can_write = (
            has_project
            and not self._workspace_project_read_only
            and self.project_store is not None
            and not self._tool_running
            and not self._project_batch_by_token
            and not self._workspace_write_in_progress()
            and not getattr(self, "_workspace_recovery_blocked", False)
        )
        self.workspace_add_button.configure(state="normal" if can_write else "disabled")
        write_busy = self._workspace_write_in_progress()
        self.workspace_refresh_button.configure(state="normal" if has_project and not write_busy else "disabled")
        self.workspace_open_project_button.configure(state="normal" if has_project else "disabled")
        if hasattr(self, "workspace_trash_button"):
            self.workspace_trash_button.configure(
                state="normal" if has_project and not write_busy else "disabled"
            )
        if hasattr(self, "workspace_switch_button"):
            self.workspace_switch_button.configure(state="disabled" if self._project_change_is_blocked() else "normal")
        if hasattr(self, "workspace_search_entry"):
            self.workspace_search_entry.configure(state="disabled" if write_busy else "normal")
        selected = self._selected_workspace_path()
        selected_exists = selected is not None and selected.exists()
        self.workspace_open_item_button.configure(state="normal" if selected_exists else "disabled")
        self.workspace_reveal_item_button.configure(state="normal" if selected_exists else "disabled")
        move_button = getattr(self, "workspace_move_to_trash_button", None)
        if move_button is not None:
            batch_root = self._workspace_batch_root_for_path(selected) if selected_exists else None
            if batch_root is None:
                if move_button.winfo_manager():
                    move_button.pack_forget()
            else:
                if not move_button.winfo_manager():
                    move_button.pack(side=RIGHT)
                summary, _detail = batch_root
                move_allowed = can_write and str(getattr(summary, "status", "")) != "running"
                move_button.configure(state="normal" if move_allowed else "disabled")
        self._update_workspace_trash_restore_state()

    def _show_workspace_empty(self, visible: bool) -> None:
        if visible:
            self.workspace_empty_label.place(relx=0.5, rely=0.36, anchor="center")
        else:
            self.workspace_empty_label.place_forget()

    def _schedule_workspace_search(self) -> None:
        if self._workspace_search_job is not None:
            try:
                self.root.after_cancel(self._workspace_search_job)
            except Exception:
                pass
        self._workspace_search_job = self.root.after(320, self._run_scheduled_workspace_search)

    def _run_scheduled_workspace_search(self) -> None:
        self._workspace_search_job = None
        if not getattr(self, "_is_alive", True):
            return
        try:
            if not self.workspace_tree.winfo_exists():
                return
        except Exception:
            return
        self._refresh_workspace_tree()

    def _search_workspace_files(self, generation: int, root_path: Path, query: str) -> None:
        results: list[Path] = []
        truncated = False
        error: str | None = None
        try:
            for current_root, dir_names, file_names in os.walk(root_path, followlinks=False):
                current = Path(current_root)
                dir_names[:] = [
                    name
                    for name in dir_names
                    if not self._workspace_should_hide_path(current / name)
                ]
                for name in (*dir_names, *file_names):
                    path = current / name
                    if self._workspace_should_hide_path(path):
                        continue
                    if query in name.casefold():
                        results.append(path)
                        if len(results) >= WORKSPACE_SEARCH_LIMIT:
                            truncated = True
                            break
                if truncated:
                    break
        except OSError as exc:
            error = str(exc)
        self._workspace_queue.put(("search", generation, (results, truncated, error)))

    def _render_workspace_search_results(self, results: list[Path], truncated: bool, error: str | None) -> None:
        for item in self.workspace_tree.get_children():
            self.workspace_tree.delete(item)
        self._workspace_tree_paths.clear()
        if error:
            self.workspace_empty_text.set(f"项目文件暂时无法读取\n\n{error}")
            self._show_workspace_empty(True)
            return
        for path in sorted(results, key=lambda item: (not item.is_dir(), item.name.casefold())):
            self._insert_workspace_tree_path("", path, search_result=True)
        if results:
            self._show_workspace_empty(False)
            if truncated:
                self.workspace_detail_title.set("搜索结果较多")
                self.workspace_detail_text.set(f"已显示前 {WORKSPACE_SEARCH_LIMIT} 项，请输入更完整的文件名。")
        else:
            self.workspace_empty_text.set("没有找到匹配的文件\n\n换一个文件名，或查看“全部文件”")
            self._show_workspace_empty(True)

    def _workspace_latest_progress(self, token: int):
        """Return one worker-owned progress snapshot while holding its lock."""

        progress_store = getattr(self, "_workspace_write_progress", {})
        progress_lock = getattr(self, "_workspace_write_progress_lock", None)
        if progress_lock is None:
            return progress_store.get(token)
        with progress_lock:
            return progress_store.get(token)

    def _poll_workspace_queue(self) -> None:
        progress_store = getattr(self, "_workspace_write_progress", {})
        progress_lock = getattr(self, "_workspace_write_progress_lock", None)
        if progress_lock is None:
            progress_items = tuple(progress_store.items())
        else:
            with progress_lock:
                progress_items = tuple(progress_store.items())
        for token, snapshot in progress_items:
            self._render_workspace_import_progress(token, snapshot)
        try:
            while True:
                status, generation, payload = self._workspace_queue.get_nowait()
                if status == "search":
                    if generation != self._workspace_search_generation:
                        continue
                    results, truncated, error = payload
                    self._render_workspace_search_results(results, truncated, error)
                elif status in {
                    "write_changed",
                    "write_cancelled",
                    "write_error",
                    "write_recovered",
                    "write_recovery_blocked",
                }:
                    token, store, *details = payload
                    tracked = self._workspace_write_tasks.get(token)
                    if tracked is None or tracked[1] is not store:
                        continue
                    # The terminal queue item and the latest progress snapshot are
                    # stored separately.  Re-read under the progress lock so a fast
                    # import cannot jump from "copying" straight to a closed dialog.
                    latest = self._workspace_latest_progress(token)
                    if latest is not None:
                        self._render_workspace_import_progress(token, latest)
                    self._workspace_write_tasks.pop(token, None)
                    callbacks = getattr(self, "_workspace_write_callbacks", {}).pop(
                        token,
                        (None, None, None),
                    )
                    if progress_lock is None:
                        progress_store.pop(token, None)
                    else:
                        with progress_lock:
                            progress_store.pop(token, None)
                    is_current = (
                        store is self.project_store
                        and generation == self._workspace_project_generation
                    )
                    if status == "write_recovery_blocked":
                        action, exc, recovery_exc = details
                        self._workspace_recovery_blocked = True
                        self._workspace_recovery_error = str(recovery_exc)
                        self._close_workspace_import_progress(token=token, force=True)
                        runlog.log_exception(f"{action}失败", exc)
                        runlog.log_exception("项目自动恢复失败", recovery_exc)
                        if is_current and not self._workspace_close_requested:
                            if callable(callbacks[2]):
                                callbacks[2](exc)
                            messagebox.showerror(
                                "项目需要恢复",
                                "资料在最后保存时没有完成，项目已暂停继续写入，以避免覆盖或遗漏资料。\n\n"
                                "请关闭工具后重新打开当前项目。程序会再次尝试安全恢复。\n\n"
                                f"原因：{recovery_exc}",
                                parent=self.root,
                            )
                    elif status == "write_error":
                        action, exc = details
                        self._close_workspace_import_progress(token=token, force=True)
                        runlog.log_exception(f"{action}失败", exc)
                        if is_current and not self._workspace_close_requested:
                            if callable(callbacks[2]):
                                callbacks[2](exc)
                            else:
                                messagebox.showerror(
                                    f"无法{action}",
                                    f"操作没有完成。\n\n原因：{exc}",
                                    parent=self.root,
                                )
                    elif status == "write_recovered":
                        action, exc = details
                        self._close_workspace_import_progress(token=token, force=True)
                        runlog.log_exception(f"{action}最后保存时出现异常", exc)
                        runlog.log_line(f"{action}已自动恢复到安全状态")
                        if is_current and not self._workspace_close_requested:
                            self._refresh_workspace_tree()
                            self._update_sidebar_project_summary()
                            messagebox.showwarning(
                                "项目已恢复到安全状态",
                                "最后保存时出现异常，程序已完成安全恢复。资料可能已经保存成功。\n\n"
                                "请先检查右侧是否已出现资料，再决定是否重试。\n\n"
                                f"原始原因：{exc}",
                                parent=self.root,
                            )
                    elif status == "write_cancelled":
                        self._close_workspace_import_progress(token=token, force=True)
                        runlog.log_line(f"{details[0]}已由用户安全停止")
                        if is_current and not self._workspace_close_requested and callable(callbacks[1]):
                            callbacks[1]()
                    elif is_current and not self._workspace_close_requested:
                        self._show_workspace_import_success(token)
                        self._refresh_workspace_tree()
                        self._update_sidebar_project_summary()
                        if callable(callbacks[0]):
                            result = details[0] if details else None
                            callbacks[0](result)
                    else:
                        self._close_workspace_import_progress(token=token, force=True)
                    self._update_workspace_action_states()
        except queue.Empty:
            pass
        if self._workspace_close_requested and not self._workspace_write_in_progress():
            self._finish_app_close()
            return
        self.root.after(180, self._poll_workspace_queue)

    def _open_selected_workspace_item(self, event=None) -> None:
        path = self._selected_workspace_path()
        if path is None or not path.exists():
            return
        if event is not None and path.is_dir():
            item = self.workspace_tree.selection()[0]
            self.workspace_tree.item(item, open=not bool(self.workspace_tree.item(item, "open")))
            if self.workspace_tree.item(item, "open"):
                self._on_workspace_tree_open()
            return
        try:
            open_path(path)
            self._close_workspace_drawer()
        except Exception as exc:
            runlog.log_exception("打开项目文件失败", exc)
            messagebox.showerror("无法打开", f"这个项目文件无法打开。\n\n原因：{exc}", parent=self.root)

    def _reveal_selected_workspace_item(self) -> None:
        path = self._selected_workspace_path()
        if path is None or not path.exists():
            return
        try:
            if path.is_dir():
                open_path(path)
            elif sys.platform.startswith("win"):
                subprocess.Popen(["explorer", "/select,", str(path)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", str(path)])
            else:
                open_path(path.parent)
            self._close_workspace_drawer()
        except Exception as exc:
            runlog.log_exception("定位项目文件失败", exc)
            messagebox.showerror("无法定位", f"暂时无法在文件夹中定位这个文件。\n\n原因：{exc}", parent=self.root)

    def _open_workspace_trash_dialog(self) -> None:
        if self.current_project_path is None or self.project_store is None:
            messagebox.showinfo("还没有工作项目", "请先新建或打开一个工作项目。", parent=self.root)
            return
        existing = getattr(self, "_workspace_trash_window", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                pass

        dialog_width, preferred_dialog_height = self._update_dialog_size(760, 487)
        window = Toplevel(self.root)
        self._workspace_trash_window = window
        self._workspace_trash_restore_in_progress = False
        setattr(window, "_hr_ui_scale", self.ui_scale)
        window.withdraw()
        window.title("项目回收站")
        window.configure(bg=COLOR_SURFACE)
        window.resizable(False, False)
        window.transient(self.root)

        self._workspace_trash_search_var = StringVar(master=window)
        self._workspace_trash_status_var = StringVar(master=window)
        self._workspace_trash_restore_path_var = StringVar(master=window, value="请选择左侧处理记录")
        self._workspace_trash_project_var = StringVar(master=window)
        self._workspace_trash_notice_var = StringVar(
            master=window,
            value="恢复时不会覆盖项目中的现有资料。",
        )
        self._workspace_trash_card_widgets = {}

        body = Frame(window, bg=COLOR_SURFACE)
        body.pack(fill=BOTH, expand=True, padx=self._pad(28), pady=self._pad(22, 20))
        Label(
            body,
            text="项目回收站",
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=(self.base_font[0], _font_size(16), "bold"),
            anchor="w",
        ).pack(fill="x")
        Label(
            body,
            text="这里保存从当前项目移走的处理批次，可以恢复，不会立即永久删除。",
            bg=COLOR_SURFACE,
            fg=COLOR_MUTED,
            font=self.base_font,
            anchor="w",
        ).pack(fill="x", pady=self._pad(2, 12))

        search_shell = Frame(body, bg=COLOR_SURFACE)
        search_shell.pack(fill="x", pady=self._pad(0, 12))
        search_entry = ttk.Entry(
            search_shell,
            textvariable=self._workspace_trash_search_var,
            style="App.TEntry",
        )
        search_entry.pack(fill="x")
        search_placeholder = Label(
            search_shell,
            text="查找已移除的批次",
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.base_font,
            anchor="w",
            cursor="xterm",
            takefocus=0,
        )

        def sync_search_placeholder() -> None:
            if self._workspace_trash_search_var is not None and self._workspace_trash_search_var.get():
                search_placeholder.place_forget()
                return
            search_placeholder.place(x=self._px(13), rely=0.5, anchor="w")

        search_placeholder.bind("<Button-1>", lambda _event: search_entry.focus_set())
        search_entry.bind("<Return>", _workspace_trash_ignore_enter)
        search_entry.bind("<KP_Enter>", _workspace_trash_ignore_enter)
        sync_search_placeholder()

        # 底部操作区先占位，再让中间双栏填充剩余空间。这样即使屏幕高度不足，
        # 也只会缩小可滚动的中间区，不会把“关闭/恢复”按钮裁掉。
        button_row = Frame(body, bg=COLOR_SURFACE)
        button_row.pack(side="bottom", fill="x", pady=self._pad(14, 0))
        restore_button = CodexButton(
            button_row,
            text="恢复到项目",
            command=self._restore_selected_workspace_trash,
            variant="primary",
            width=104,
            min_width=96,
            height=36,
        )
        restore_button.pack(side=RIGHT)
        self._workspace_trash_restore_button = restore_button
        close_button = CodexButton(
            button_row,
            text="关闭",
            command=self._close_workspace_trash_dialog,
            width=70,
            min_width=64,
            height=36,
        )
        close_button.pack(side=RIGHT, padx=self._pad(0, 9))
        for button in (close_button, restore_button):
            button.configure(takefocus=1)

            def activate(_event=None, target=button):
                target._on_click()
                return "break"

            button.bind("<Return>", activate)
            button.bind("<space>", activate)

        columns = Frame(body, bg=COLOR_SURFACE)
        columns.pack(fill=BOTH, expand=True)
        left = Frame(
            columns,
            width=self._px(322),
            bg=COLOR_SURFACE_ALT,
            highlightbackground=COLOR_BORDER,
            highlightthickness=self._px(1),
            bd=0,
        )
        left.pack(side=LEFT, fill=BOTH)
        left.pack_propagate(False)
        Label(
            left,
            text="已移入的处理记录",
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_MUTED,
            font=self.small_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(12), pady=self._pad(10, 6))
        list_shell = Frame(left, bg=COLOR_SURFACE_ALT)
        list_shell.pack(fill=BOTH, expand=True, padx=self._pad(6), pady=self._pad(0, 6))
        list_scroll = ttk.Scrollbar(list_shell, orient=VERTICAL)
        list_scroll.pack(side=RIGHT, fill=Y)
        list_canvas = Canvas(
            list_shell,
            bg=COLOR_SURFACE_ALT,
            highlightthickness=0,
            bd=0,
            yscrollcommand=list_scroll.set,
        )
        list_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        list_scroll.configure(command=list_canvas.yview)
        list_body = Frame(list_canvas, bg=COLOR_SURFACE_ALT)
        self._workspace_trash_list_body = list_body
        list_window = list_canvas.create_window(0, 0, window=list_body, anchor="nw")

        def sync_trash_list(_event=None) -> None:
            list_canvas.itemconfigure(list_window, width=max(list_canvas.winfo_width(), 1))
            bounds = list_canvas.bbox("all")
            if bounds is not None:
                list_canvas.configure(scrollregion=bounds)

        list_body.bind("<Configure>", sync_trash_list)
        list_canvas.bind("<Configure>", sync_trash_list)

        def scroll_trash_list(event) -> str | None:
            if list_body.winfo_reqheight() <= list_canvas.winfo_height():
                return None
            if getattr(event, "num", None) == 4:
                direction = -1
            elif getattr(event, "num", None) == 5:
                direction = 1
            else:
                direction = -1 if getattr(event, "delta", 0) > 0 else 1
            list_canvas.yview_scroll(direction, "units")
            return "break"

        list_canvas.bind("<MouseWheel>", scroll_trash_list)
        list_canvas.bind("<Button-4>", scroll_trash_list)
        list_canvas.bind("<Button-5>", scroll_trash_list)
        window.bind("<MouseWheel>", scroll_trash_list)
        window.bind("<Button-4>", scroll_trash_list)
        window.bind("<Button-5>", scroll_trash_list)
        self._workspace_trash_empty_label = None

        right = Frame(
            columns,
            bg=COLOR_SURFACE,
            highlightbackground=COLOR_BORDER,
            highlightthickness=self._px(1),
            bd=0,
        )
        right.pack(side=LEFT, fill=BOTH, expand=True, padx=self._pad(12, 0))
        detail_title = Label(
            right,
            text="选择一条记录查看恢复位置",
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=self.card_title_font,
            anchor="w",
            justify="left",
            wraplength=self._px(320),
        )
        detail_title.pack(fill="x", padx=self._pad(16), pady=self._pad(14, 10))
        self._workspace_trash_detail_title = detail_title
        Label(
            right,
            text="恢复位置",
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.small_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(16))
        Label(
            right,
            textvariable=self._workspace_trash_restore_path_var,
            bg=COLOR_SURFACE,
            fg=COLOR_PRIMARY,
            font=self.section_font,
            anchor="w",
            justify="left",
            wraplength=self._px(320),
        ).pack(fill="x", padx=self._pad(16), pady=self._pad(3, 10))
        Label(
            right,
            text="当前项目",
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.small_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(16))
        Label(
            right,
            textvariable=self._workspace_trash_project_var,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=self.base_font,
            anchor="w",
            justify="left",
            wraplength=self._px(320),
        ).pack(fill="x", padx=self._pad(16), pady=self._pad(3, 10))
        notice_card = Frame(right, bg=COLOR_WARNING_SOFT)
        notice_card.pack(fill="x", padx=self._pad(16), pady=self._pad(0, 8))
        Label(
            notice_card,
            textvariable=self._workspace_trash_notice_var,
            bg=COLOR_WARNING_SOFT,
            fg=COLOR_WARNING,
            font=self.small_font,
            anchor="w",
            justify="left",
            wraplength=self._px(288),
        ).pack(fill="x", padx=self._pad(12), pady=self._pad(10))
        Label(
            right,
            textvariable=self._workspace_trash_status_var,
            bg=COLOR_SURFACE,
            fg=COLOR_DANGER,
            font=self.small_font,
            anchor="w",
            justify="left",
            wraplength=self._px(320),
        ).pack(fill="x", padx=self._pad(16), pady=self._pad(0, 6))

        def search_changed(*_args) -> None:
            sync_search_placeholder()
            self._render_workspace_trash_cards()

        self._workspace_trash_search_trace = self._workspace_trash_search_var.trace_add(
            "write",
            search_changed,
        )
        window.protocol("WM_DELETE_WINDOW", self._close_workspace_trash_dialog)
        window.bind("<Escape>", lambda _event: self._close_workspace_trash_dialog())

        self._reload_workspace_trash_details()
        window.update_idletasks()
        try:
            available_height = max(1, window.winfo_screenheight() - self._px(72))
        except Exception:
            available_height = max(preferred_dialog_height, window.winfo_reqheight())
        dialog_height = _workspace_trash_dialog_height(
            preferred_dialog_height,
            window.winfo_reqheight(),
            available_height,
        )
        self._center_window(window, dialog_width, dialog_height)
        window.deiconify()
        try:
            window.grab_set()
        except Exception:
            pass
        search_entry.focus_set()

    def _reload_workspace_trash_details(self) -> None:
        store = self.project_store
        if store is None:
            return
        loader = getattr(store, "list_trash_details", None)
        if not callable(loader):
            status_var = getattr(self, "_workspace_trash_status_var", None)
            if status_var is not None:
                status_var.set("当前版本暂时无法读取项目回收站。")
            self._render_workspace_trash_cards()
            return
        try:
            details = tuple(loader())
        except Exception as exc:
            runlog.log_exception("读取项目回收站失败", exc)
            if self._workspace_trash_status_var is not None:
                self._workspace_trash_status_var.set(f"回收站暂时无法读取：{exc}")
            self._render_workspace_trash_cards()
            return
        self._workspace_trash_details = details
        available_ids = {str(getattr(item.summary, "id", "")) for item in details}
        if self._workspace_trash_selected_id not in available_ids:
            self._workspace_trash_selected_id = str(details[0].summary.id) if details else None
        if self._workspace_trash_status_var is not None:
            self._workspace_trash_status_var.set("")
        self._render_workspace_trash_cards()

    def _render_workspace_trash_cards(self) -> None:
        list_body = getattr(self, "_workspace_trash_list_body", None)
        if list_body is None:
            return
        for child in list_body.winfo_children():
            child.destroy()
        self._workspace_trash_card_widgets = {}
        query_var = getattr(self, "_workspace_trash_search_var", None)
        query = query_var.get() if query_var is not None else ""
        filtered = [item for item in self._workspace_trash_details if _workspace_trash_matches(item, query)]
        if not filtered:
            empty = Label(
                list_body,
                text=(
                    "没有找到匹配的处理记录。"
                    if self._workspace_trash_details
                    else "回收站是空的。\n\n从项目中移走的处理批次会暂时保存在这里。"
                ),
                bg=COLOR_SURFACE_ALT,
                fg=COLOR_FAINT,
                font=self.small_font,
                justify="center",
                anchor="center",
                wraplength=self._px(250),
            )
            self._workspace_trash_empty_label = empty
            empty.pack(fill=BOTH, expand=True, padx=self._pad(16), pady=self._pad(42))
            self._render_workspace_trash_detail(None)
            return
        self._workspace_trash_empty_label = None

        filtered_ids = {str(item.summary.id) for item in filtered}
        if self._workspace_trash_selected_id not in filtered_ids:
            self._workspace_trash_selected_id = str(filtered[0].summary.id)
        for detail in filtered:
            batch_id = str(detail.summary.id)
            selected = batch_id == self._workspace_trash_selected_id
            background = COLOR_PRIMARY_SOFT if selected else COLOR_SURFACE
            card = Frame(
                list_body,
                bg=background,
                highlightbackground=COLOR_PRIMARY if selected else COLOR_BORDER,
                highlightthickness=self._px(1),
                bd=0,
                takefocus=True,
                cursor="hand2",
            )
            card.pack(fill="x", padx=self._pad(4), pady=self._pad(0, 7))
            title_row = Frame(card, bg=background)
            title_row.pack(fill="x", padx=self._pad(11), pady=self._pad(9, 2))
            title = Label(
                title_row,
                text=_workspace_trash_title(detail),
                bg=background,
                fg=COLOR_TEXT,
                font=self.section_font,
                anchor="w",
                justify="left",
                wraplength=self._px(195),
            )
            title.pack(side=LEFT, fill="x", expand=True)
            status = HISTORY_STATUS_LABELS.get(str(detail.summary.status), "未开始")
            status_label = Label(
                title_row,
                text=status,
                bg=COLOR_PRIMARY_SOFT,
                fg=COLOR_PRIMARY,
                font=(self.base_font[0], _font_size(8), "bold"),
                highlightbackground=COLOR_PRIMARY,
                highlightthickness=self._px(1),
                anchor="center",
                padx=self._px(5),
                pady=self._px(2),
            )
            status_label.pack(side=RIGHT, padx=self._pad(8, 0))
            group_tool = Label(
                card,
                text=_workspace_trash_group_tool(detail),
                bg=background,
                fg=COLOR_MUTED,
                font=self.tiny_font,
                anchor="w",
                justify="left",
                wraplength=self._px(260),
            )
            group_tool.pack(fill="x", padx=self._pad(11), pady=self._pad(0, 3))
            meta = Label(
                card,
                text=f"移入时间：{_workspace_trash_deleted_text(detail.summary.deleted_at)}",
                bg=background,
                fg=COLOR_MUTED,
                font=self.tiny_font,
                anchor="w",
            )
            meta.pack(fill="x", padx=self._pad(11))
            counts = (
                f"上传资料 {detail.upload_count} 个 · 处理结果 {detail.result_count} 个 · "
                f"补充资料 {detail.supplement_count} 个 · {self._format_history_size(detail.total_size_bytes)}"
            )
            count_label = Label(
                card,
                text=counts,
                bg=background,
                fg=COLOR_FAINT,
                font=self.tiny_font,
                anchor="w",
                justify="left",
                wraplength=self._px(260),
            )
            count_label.pack(fill="x", padx=self._pad(11), pady=self._pad(3, 9))
            background_widgets = (title_row, title, group_tool, meta, count_label)
            self._workspace_trash_card_widgets[batch_id] = (card, background_widgets)

            def select(_event=None, selected_id=batch_id):
                self._select_workspace_trash_detail(selected_id)
                return "break"

            for widget in (card, *background_widgets, status_label):
                widget.bind("<Button-1>", select)
            card.bind("<Return>", select)
            card.bind("<space>", select)

        self._select_workspace_trash_detail(self._workspace_trash_selected_id)

    def _workspace_trash_current_detail(self):
        selected_id = self._workspace_trash_selected_id
        if selected_id is None:
            return None
        query_var = getattr(self, "_workspace_trash_search_var", None)
        query = query_var.get() if query_var is not None else ""
        return next(
            (
                item
                for item in self._workspace_trash_details
                if str(item.summary.id) == selected_id and _workspace_trash_matches(item, query)
            ),
            None,
        )

    def _select_workspace_trash_detail(self, batch_id: str | None) -> None:
        self._workspace_trash_selected_id = str(batch_id) if batch_id else None
        for item_id, (card, background_widgets) in self._workspace_trash_card_widgets.items():
            selected = item_id == self._workspace_trash_selected_id
            background = COLOR_PRIMARY_SOFT if selected else COLOR_SURFACE
            card.configure(
                bg=background,
                highlightbackground=COLOR_PRIMARY if selected else COLOR_BORDER,
            )
            for widget in background_widgets:
                widget.configure(bg=background)
        self._render_workspace_trash_detail(self._workspace_trash_current_detail())

    def _render_workspace_trash_detail(self, detail) -> None:
        title_label = getattr(self, "_workspace_trash_detail_title", None)
        restore_path_var = getattr(self, "_workspace_trash_restore_path_var", None)
        project_var = getattr(self, "_workspace_trash_project_var", None)
        notice_var = getattr(self, "_workspace_trash_notice_var", None)
        if detail is None:
            if title_label is not None:
                title_label.configure(text="选择一条记录查看恢复位置")
            if restore_path_var is not None:
                restore_path_var.set("请选择左侧处理记录")
            if project_var is not None:
                project_var.set(self.workspace_project_name.get())
            if notice_var is not None:
                notice_var.set("恢复时不会覆盖项目中的现有资料。")
            self._update_workspace_trash_restore_state()
            return

        if title_label is not None:
            title_label.configure(text=_workspace_trash_title(detail))
        if restore_path_var is not None:
            restore_path_var.set(_workspace_trash_restore_location(detail))
        workspace = getattr(self.project_store, "workspace", None)
        project_name = str(getattr(workspace, "name", "") or self.workspace_project_name.get())
        if project_var is not None:
            project_var.set(project_name + (" · 只读" if self._workspace_project_read_only else ""))
        if notice_var is not None:
            notice_var.set("不会覆盖现有资料。若已有同名批次，系统会自动使用新名称恢复。")
        if self._workspace_trash_status_var is not None and not self._workspace_trash_restore_in_progress:
            self._workspace_trash_status_var.set(
                "当前项目为只读，只能查看回收站。" if self._workspace_project_read_only else ""
            )
        self._update_workspace_trash_restore_state()

    def _update_workspace_trash_restore_state(self) -> None:
        button = getattr(self, "_workspace_trash_restore_button", None)
        if button is None:
            return
        detail = self._workspace_trash_current_detail()
        busy = self._workspace_write_in_progress()
        can_restore = bool(
            detail is not None
            and self.current_project_path is not None
            and self.project_store is not None
            and not self._workspace_project_read_only
            and not self._tool_running
            and not self._project_batch_by_token
            and not self._workspace_recovery_blocked
            and not busy
            and not self._workspace_trash_restore_in_progress
        )
        button.configure(
            text="正在恢复…" if self._workspace_trash_restore_in_progress else "恢复到项目",
            state="normal" if can_restore else "disabled",
        )

    def _restore_selected_workspace_trash(self) -> None:
        detail = self._workspace_trash_current_detail()
        store = self.project_store
        if detail is None or store is None or self._workspace_trash_restore_in_progress:
            return
        if self._workspace_project_read_only:
            if self._workspace_trash_status_var is not None:
                self._workspace_trash_status_var.set("当前项目为只读，只能查看回收站。")
            self._update_workspace_trash_restore_state()
            return
        if self._workspace_recovery_blocked:
            if self._workspace_trash_status_var is not None:
                self._workspace_trash_status_var.set("当前项目需要重新打开并完成恢复后，才能恢复回收站资料。")
            self._update_workspace_trash_restore_state()
            return
        if self._tool_running or self._project_batch_by_token or self._workspace_write_in_progress():
            if self._workspace_trash_status_var is not None:
                self._workspace_trash_status_var.set("请等待当前处理或资料保存完成后再恢复。")
            self._update_workspace_trash_restore_state()
            return

        batch_id = str(detail.summary.id)
        self._workspace_trash_restore_in_progress = True
        if self._workspace_trash_status_var is not None:
            self._workspace_trash_status_var.set("正在恢复到当前项目，请稍候…")
        self._update_workspace_trash_restore_state()

        def restored(_result) -> None:
            self._workspace_trash_restore_in_progress = False
            restore_location = _workspace_trash_restore_location(detail)
            self._close_workspace_trash_dialog(force=True)
            messagebox.showinfo(
                "已恢复到项目",
                f"处理批次已恢复到：\n当前项目 / {restore_location}\n\n现有资料没有被覆盖。",
                parent=self.root,
            )

        def restore_failed(exc: Exception) -> None:
            self._workspace_trash_restore_in_progress = False
            if self._workspace_trash_status_var is not None:
                self._workspace_trash_status_var.set(f"恢复没有完成：{exc}")
            self._update_workspace_trash_restore_state()

        self._workspace_run_write(
            "恢复项目资料",
            lambda _cancelled: store.restore_from_trash(batch_id),
            on_success=restored,
            on_error=restore_failed,
        )

    def _close_workspace_trash_dialog(self, *, force: bool = False) -> None:
        if getattr(self, "_workspace_trash_restore_in_progress", False) and not force:
            return
        search_var = getattr(self, "_workspace_trash_search_var", None)
        trace_id = getattr(self, "_workspace_trash_search_trace", None)
        if search_var is not None and trace_id:
            try:
                search_var.trace_remove("write", trace_id)
            except Exception:
                pass
        window = getattr(self, "_workspace_trash_window", None)
        self._workspace_trash_window = None
        self._workspace_trash_search_var = None
        self._workspace_trash_status_var = None
        self._workspace_trash_restore_path_var = None
        self._workspace_trash_project_var = None
        self._workspace_trash_notice_var = None
        self._workspace_trash_search_trace = None
        self._workspace_trash_details = ()
        self._workspace_trash_selected_id = None
        self._workspace_trash_restore_in_progress = False
        self._workspace_trash_card_widgets = {}
        self._workspace_trash_list_body = None
        self._workspace_trash_empty_label = None
        self._workspace_trash_detail_title = None
        self._workspace_trash_restore_button = None
        if window is not None:
            try:
                window.grab_release()
            except Exception:
                pass
            try:
                window.destroy()
            except Exception:
                pass

    def _move_selected_workspace_batch_to_trash(self) -> None:
        selected = self._selected_workspace_path()
        batch_root = self._workspace_batch_root_for_path(selected)
        store = self.project_store
        if batch_root is None or store is None:
            return
        summary, _detail = batch_root
        if (
            self._workspace_project_read_only
            or self._tool_running
            or self._project_batch_by_token
            or self._workspace_write_in_progress()
            or str(summary.status) == "running"
        ):
            messagebox.showinfo(
                "暂时不能移到回收站",
                "请等待当前处理或资料保存完成，并确认项目不是只读状态。",
                parent=self.root,
            )
            return
        display_name = str(summary.business_description or summary.business_period or summary.directory_name)
        if not messagebox.askyesno(
            "移到项目回收站",
            f"“{display_name}”的上传资料、处理结果和补充资料会一起移到当前项目回收站。\n\n"
            "不会永久删除，之后可以从“回收站”恢复。是否继续？",
            parent=self.root,
        ):
            return
        batch_id = str(summary.id)

        def moved(_result) -> None:
            messagebox.showinfo(
                "已移到项目回收站",
                "完整处理批次已移到当前项目回收站，之后可以恢复。",
                parent=self.root,
            )

        def move_failed(exc: Exception) -> None:
            messagebox.showerror(
                "无法移到回收站",
                f"处理批次没有被移走。\n\n原因：{exc}",
                parent=self.root,
            )

        self._workspace_run_write(
            "移到项目回收站",
            lambda _cancelled: store.move_to_trash(batch_id),
            on_success=moved,
            on_error=move_failed,
        )

    def _show_workspace_add_menu(self) -> None:
        can_write = (
            self.current_project_path is not None
            and not self._workspace_project_read_only
            and self.project_store is not None
            and not self._tool_running
            and not self._project_batch_by_token
            and not self._workspace_write_in_progress()
        )
        state = "normal" if can_write else "disabled"
        menu = Menu(self.root, tearoff=0, font=self.base_font)
        menu.add_command(label="导入文件", command=self._import_workspace_files, state=state)
        menu.add_command(label="导入文件夹", command=self._import_workspace_folder, state=state)
        menu.add_separator()
        menu.add_command(label="新建文件夹", command=self._create_workspace_folder, state=state)
        self._popup_workspace_menu(menu, self.workspace_add_button)

    def _workspace_selected_target(self) -> Path | None:
        selected = self._selected_workspace_path()
        if selected is None and self.workspace_scope.get() == WORKSPACE_SCOPE_TOOL:
            messagebox.showinfo(
                "请先选择保存位置",
                "请先展开并选择某次处理的“上传资料”或“补充资料”。\n\n"
                "如果只是提前存放资料，请切换到“全部文件”，再放入“共用资料”。",
                parent=self.root,
            )
            return None
        if selected is not None and selected.exists():
            target = selected if selected.is_dir() else selected.parent
        else:
            target = self._workspace_scope_root() or self.current_project_path
        if target is None:
            return None
        store = self.project_store
        workspace = getattr(store, "workspace", None) if store is not None else None
        common_root = getattr(workspace, "common_root", None)
        if target == self.current_project_path or not target.is_dir():
            if isinstance(common_root, Path) and common_root.is_dir():
                target = common_root
        relative_parts: tuple[str, ...] = ()
        try:
            relative_parts = target.relative_to(self.current_project_path).parts
        except ValueError:
            pass
        if target.name == "处理结果" or "处理结果" in relative_parts:
            result_parent = target
            while result_parent != self.current_project_path and result_parent.name != "处理结果":
                result_parent = result_parent.parent
            supplement_target = result_parent.parent / "补充资料"
            if result_parent.name == "处理结果":
                if not messagebox.askyesno(
                    "改存到补充资料",
                    "“处理结果”只保存工具生成的正式文件。是否改为添加到同一批次的“补充资料”？",
                    parent=self.root,
                ):
                    return None
                target = supplement_target
        elif (target / "上传资料").is_dir():
            target = target / "上传资料"
        else:
            # 业务分组和工具目录由系统维护，人工资料只进入“共用资料”或
            # 某次批次的上传/补充资料，避免出现无法追溯的散落文件。
            batch_target = self._workspace_batch_for_target(target)
            in_common = False
            if isinstance(common_root, Path):
                try:
                    target.relative_to(common_root)
                    in_common = True
                except ValueError:
                    pass
            if batch_target is None and not in_common and isinstance(common_root, Path):
                target = common_root
        batch_target = self._workspace_batch_for_target(target)
        if batch_target is not None and batch_target[1] == "uploads":
            batch_id, _category = batch_target
            try:
                detail = self.project_store.get_batch(batch_id)
            except Exception:
                detail = None
            if detail is not None and detail.summary.status != "draft":
                if not messagebox.askyesno(
                    "改存到补充资料",
                    "已开始或已完成的“上传资料”需要保持原样，方便追溯。是否改为添加到本批次的“补充资料”？",
                    parent=self.root,
                ):
                    return None
                target = detail.directories["supplements"]
        return target

    def _workspace_batch_for_target(self, target: Path) -> tuple[str, str] | None:
        store = self.project_store
        if store is None:
            return None
        try:
            batches = store.list_batches()
        except Exception as exc:
            runlog.log_exception("读取项目批次失败", exc)
            return None
        for summary in batches:
            try:
                detail = store.get_batch(summary.id)
            except Exception:
                continue
            if detail is None:
                continue
            for category in ("uploads", "supplements"):
                directory = detail.directories.get(category)
                if directory is None:
                    continue
                try:
                    target.relative_to(directory)
                except ValueError:
                    continue
                return summary.id, category
        return None

    def _workspace_import_target_text(self, target: Path) -> str:
        try:
            relative = Path(target).relative_to(self.current_project_path)
        except (TypeError, ValueError):
            return "当前项目"
        parts = ["当前项目", *relative.parts]
        return " / ".join(parts)

    def _open_workspace_import_progress(self, token: int, target: Path) -> None:
        self._close_workspace_import_progress(force=True)
        dialog_width, dialog_height = self._update_dialog_size(640, 510)
        window = Toplevel(self.root)
        self._workspace_import_window = window
        self._workspace_import_token = token
        self._workspace_import_started_at = time.monotonic()
        self._workspace_import_phase = "checking"
        setattr(window, "_hr_ui_scale", self.ui_scale)
        window.withdraw()
        window.title("正在把资料保存到项目")
        window.configure(bg=COLOR_SURFACE)
        window.resizable(False, False)
        window.transient(self.root)

        self._workspace_import_title_var = StringVar(master=window, value="正在检查所选资料…")
        self._workspace_import_subtitle_var = StringVar(
            master=window,
            value="正在确认文件数量、大小和可用空间，资料还没有开始保存。",
        )
        self._workspace_import_target_var = StringVar(
            master=window,
            value=self._workspace_import_target_text(target),
        )
        self._workspace_import_state_var = StringVar(master=window, value="正在扫描")
        self._workspace_import_name_var = StringVar(master=window, value="正在读取所选资料…")
        self._workspace_import_left_var = StringVar(master=window, value="已发现 0 个文件")
        self._workspace_import_middle_var = StringVar(master=window, value="正在计算总大小")
        self._workspace_import_elapsed_var = StringVar(master=window, value="已用时 00:00")
        self._workspace_import_safety_title_var = StringVar(master=window, value="检查完成后开始保存")
        self._workspace_import_safety_text_var = StringVar(
            master=window,
            value="现在取消不会写入项目；外部原文件不会被移动或修改。",
        )

        body = Frame(window, bg=COLOR_SURFACE)
        body.pack(fill=BOTH, expand=True, padx=self._pad(32), pady=self._pad(28, 24))
        Label(
            body,
            textvariable=self._workspace_import_title_var,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=(self.base_font[0], _font_size(16), "bold"),
            anchor="w",
        ).pack(fill="x")
        Label(
            body,
            textvariable=self._workspace_import_subtitle_var,
            bg=COLOR_SURFACE,
            fg=COLOR_MUTED,
            font=self.base_font,
            anchor="w",
        ).pack(fill="x", pady=self._pad(3, 14))

        stage_row = Frame(body, bg=COLOR_SURFACE)
        stage_row.pack(fill="x", pady=self._pad(0, 16))
        self._workspace_import_stage_labels = []
        for text_value in ("1  检查资料", "2  复制并校验", "3  完成保存"):
            stage = Label(
                stage_row,
                text=text_value,
                bg=COLOR_SURFACE_PRESSED,
                fg=COLOR_MUTED,
                font=self.small_font,
                padx=self._px(12),
                pady=self._px(6),
            )
            stage.pack(side=LEFT, padx=self._pad(0, 8))
            self._workspace_import_stage_labels.append(stage)

        target_card = Frame(
            body,
            bg=COLOR_SURFACE,
            highlightbackground=COLOR_BORDER,
            highlightthickness=self._px(1),
            bd=0,
        )
        target_card.pack(fill="x")
        Label(
            target_card,
            text="保存到",
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.small_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(10, 2))
        Label(
            target_card,
            textvariable=self._workspace_import_target_var,
            bg=COLOR_SURFACE,
            fg=COLOR_PRIMARY,
            font=self.section_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(0, 10))

        progress_card = Frame(body, bg=COLOR_SURFACE_ALT)
        progress_card.pack(fill="x", pady=self._pad(16, 0))
        Label(
            progress_card,
            textvariable=self._workspace_import_state_var,
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_FAINT,
            font=self.small_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(11, 3))
        Label(
            progress_card,
            textvariable=self._workspace_import_name_var,
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_TEXT,
            font=self.section_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(14))
        self._workspace_import_progress_width = max(self._px(220), dialog_width - self._px(92))
        self._workspace_import_progress_canvas = Canvas(
            progress_card,
            width=self._workspace_import_progress_width,
            height=self._px(10),
            bg=COLOR_SURFACE_ALT,
            highlightthickness=0,
            bd=0,
        )
        self._workspace_import_progress_canvas.pack(fill="x", padx=self._pad(14), pady=self._pad(7, 5))
        stats = Frame(progress_card, bg=COLOR_SURFACE_ALT)
        stats.pack(fill="x", padx=self._pad(14), pady=self._pad(0, 11))
        Label(
            stats,
            textvariable=self._workspace_import_left_var,
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_TEXT,
            font=self.small_font,
        ).pack(side=LEFT)
        Label(
            stats,
            textvariable=self._workspace_import_elapsed_var,
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_MUTED,
            font=self.small_font,
        ).pack(side=RIGHT)
        Label(
            stats,
            textvariable=self._workspace_import_middle_var,
            bg=COLOR_SURFACE_ALT,
            fg=COLOR_MUTED,
            font=self.small_font,
        ).pack(side=RIGHT, padx=self._pad(0, 14))

        safety = Frame(body, bg=COLOR_PRIMARY_SOFT)
        safety.pack(fill="x", pady=self._pad(16, 0))
        Label(
            safety,
            textvariable=self._workspace_import_safety_title_var,
            bg=COLOR_PRIMARY_SOFT,
            fg=COLOR_PRIMARY,
            font=self.section_font,
            anchor="w",
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(11, 2))
        Label(
            safety,
            textvariable=self._workspace_import_safety_text_var,
            bg=COLOR_PRIMARY_SOFT,
            fg=COLOR_PRIMARY,
            font=self.small_font,
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=self._pad(14), pady=self._pad(0, 11))

        button_row = Frame(body, bg=COLOR_SURFACE)
        button_row.pack(fill="x", pady=self._pad(17, 0))
        cancel_button = CodexButton(
            button_row,
            text="取消导入",
            command=self._cancel_workspace_import,
            width=92,
            min_width=86,
            height=38,
        )
        cancel_button.pack(side=RIGHT)
        self._workspace_import_cancel_button = cancel_button
        window.protocol("WM_DELETE_WINDOW", self._cancel_workspace_import)
        window.bind("<Escape>", lambda _event: self._cancel_workspace_import())

        self._set_workspace_import_stage("checking")
        window.update_idletasks()
        self._center_window(window, dialog_width, dialog_height)
        window.deiconify()
        try:
            window.grab_set()
        except Exception:
            pass
        window.focus_set()

    def _set_workspace_import_stage(self, phase: str) -> None:
        self._workspace_import_phase = phase
        active_index = {"checking": 0, "copying": 1, "finalizing": 2}.get(phase, 0)
        for index, label in enumerate(getattr(self, "_workspace_import_stage_labels", [])):
            if index < active_index:
                label.configure(bg=COLOR_PRIMARY_SOFT, fg=COLOR_PRIMARY)
            elif index == active_index:
                label.configure(bg=COLOR_PRIMARY, fg="#ffffff")
            else:
                label.configure(bg=COLOR_SURFACE_PRESSED, fg=COLOR_MUTED)
        cancel_button = getattr(self, "_workspace_import_cancel_button", None)
        if cancel_button is not None and phase == "finalizing":
            cancel_button.configure(text="正在完成…", state="disabled")

    def _draw_workspace_import_progress(self, fraction: float | None) -> None:
        canvas = getattr(self, "_workspace_import_progress_canvas", None)
        if canvas is None:
            return
        try:
            width = max(canvas.winfo_width(), self._workspace_import_progress_width)
        except Exception:
            width = self._workspace_import_progress_width
        height = self._pxf(8)
        canvas.delete("all")
        CodexButton._draw_round_rect(
            canvas,
            0,
            self._pxf(1),
            width,
            height,
            self._pxf(4),
            fill="#E2E0DA",
            outline="",
        )
        if fraction is None:
            segment = min(self._pxf(164), width * 0.30)
            visible = _indeterminate_progress_segment(width, self._workspace_import_animation_offset, segment)
            if visible is None:
                return
            start, end = visible
        else:
            start, end = 0.0, width * max(0.0, min(float(fraction), 1.0))
            if end <= 0:
                return
        CodexButton._draw_round_rect(
            canvas,
            start,
            self._pxf(1),
            end,
            height,
            self._pxf(4),
            fill=COLOR_PRIMARY if self._workspace_import_phase != "checking" else "#87B7A8",
            outline="",
        )

    def _tick_workspace_import_animation(self) -> None:
        if self._workspace_import_window is None or self._workspace_import_phase != "checking":
            self._workspace_import_animation_job = None
            return
        width = max(float(self._workspace_import_progress_width), 1.0)
        self._workspace_import_animation_offset = (self._workspace_import_animation_offset + self._pxf(9)) % (
            width + self._pxf(164)
        )
        self._draw_workspace_import_progress(None)
        self._workspace_import_animation_job = self.root.after(40, self._tick_workspace_import_animation)

    @staticmethod
    def _workspace_progress_field(progress, name: str, default=None):
        return getattr(progress, name, default)

    def _render_workspace_import_progress(self, token: int, progress) -> None:
        if token != getattr(self, "_workspace_import_token", None) or progress is None:
            return
        phase = str(self._workspace_progress_field(progress, "phase", "checking"))
        if phase not in {"checking", "copying", "finalizing"}:
            return
        if phase != self._workspace_import_phase:
            self._set_workspace_import_stage(phase)
        current_name = str(self._workspace_progress_field(progress, "current_name", "") or "")
        scanned = max(0, int(self._workspace_progress_field(progress, "files_scanned", 0) or 0))
        completed = max(0, int(self._workspace_progress_field(progress, "files_completed", 0) or 0))
        total = self._workspace_progress_field(progress, "files_total", None)
        total = None if total is None else max(0, int(total))
        copied = max(0, int(self._workspace_progress_field(progress, "bytes_copied", 0) or 0))
        total_bytes = self._workspace_progress_field(progress, "bytes_total", None)
        total_bytes = None if total_bytes is None else max(0, int(total_bytes))

        if self._workspace_import_name_var is not None:
            self._workspace_import_name_var.set(current_name or ("正在核对项目清单" if phase == "finalizing" else "正在读取所选资料…"))
        if phase == "checking":
            self._workspace_import_title_var.set("正在检查所选资料…")
            self._workspace_import_subtitle_var.set("正在确认文件数量、大小和可用空间，资料还没有开始保存。")
            self._workspace_import_state_var.set("正在扫描")
            self._workspace_import_left_var.set(f"已发现 {scanned} 个文件")
            self._workspace_import_middle_var.set(
                f"共 {self._format_history_size(total_bytes)}" if total_bytes is not None else "正在计算总大小"
            )
            self._workspace_import_safety_title_var.set("检查完成后开始保存")
            self._workspace_import_safety_text_var.set("现在取消不会写入项目；外部原文件不会被移动或修改。")
            if self._workspace_import_animation_job is None:
                self._tick_workspace_import_animation()
        elif phase == "copying":
            self._workspace_import_title_var.set("正在把资料保存到项目")
            self._workspace_import_subtitle_var.set("外部原文件不会被移动或修改。")
            self._workspace_import_state_var.set("正在处理")
            total_text = str(total) if total is not None else "?"
            self._workspace_import_left_var.set(f"已完成 {completed} / {total_text} 个文件")
            if total_bytes is not None:
                self._workspace_import_middle_var.set(
                    f"{self._format_history_size(copied)} / {self._format_history_size(total_bytes)}"
                )
            else:
                self._workspace_import_middle_var.set(self._format_history_size(copied))
            self._workspace_import_safety_title_var.set("完成后才会显示在项目中")
            self._workspace_import_safety_text_var.set(
                "全部复制并校验完成后，资料才会加入项目。取消后，本次尚未完成的资料不会加入项目。"
            )
            fraction = copied / total_bytes if total_bytes else (completed / total if total else 0.0)
            self._draw_workspace_import_progress(fraction)
        else:
            self._workspace_import_title_var.set("正在完成保存")
            self._workspace_import_subtitle_var.set("正在完成安全保存，暂时无法取消。")
            self._workspace_import_state_var.set("正在完成")
            total_value = total if total is not None else completed
            self._workspace_import_left_var.set(f"{completed} / {total_value} 个文件")
            if total_bytes is not None:
                self._workspace_import_middle_var.set(
                    f"{self._format_history_size(copied)} / {self._format_history_size(total_bytes)}"
                )
            else:
                self._workspace_import_middle_var.set("正在登记")
            self._workspace_import_safety_title_var.set("正在进行最后的安全登记")
            self._workspace_import_safety_text_var.set(
                "请保持工具打开；如电脑意外中断，下次打开项目时会自动恢复。"
            )
            self._draw_workspace_import_progress(1.0)
        started = self._workspace_import_started_at
        if started is not None and self._workspace_import_elapsed_var is not None:
            elapsed = max(0, int(time.monotonic() - started))
            self._workspace_import_elapsed_var.set(f"已用时 {elapsed // 60:02d}:{elapsed % 60:02d}")

    def _cancel_workspace_import(self) -> None:
        token = getattr(self, "_workspace_import_token", None)
        if token is None:
            return
        latest = self._workspace_latest_progress(token)
        latest_phase = str(self._workspace_progress_field(latest, "phase", "")) if latest is not None else ""
        if latest_phase == "finalizing":
            # The worker may have crossed the non-cancellable boundary since
            # Tk last painted the dialog.  Reflect that boundary before return.
            if getattr(self, "_workspace_import_phase", "") != "finalizing":
                self._set_workspace_import_stage("finalizing")
            return
        if getattr(self, "_workspace_import_phase", "") == "finalizing":
            return
        tracked = getattr(self, "_workspace_write_tasks", {}).get(token)
        if tracked is None:
            return
        tracked[0].set()
        if self._workspace_import_title_var is not None:
            self._workspace_import_title_var.set("正在安全停止…")
        if self._workspace_import_subtitle_var is not None:
            self._workspace_import_subtitle_var.set("正在清理尚未加入项目的临时文件，请稍候。")
        cancel_button = getattr(self, "_workspace_import_cancel_button", None)
        if cancel_button is not None:
            cancel_button.configure(text="正在停止…", state="disabled")

    def _show_workspace_import_success(self, token: int) -> None:
        if token != getattr(self, "_workspace_import_token", None):
            return
        self._set_workspace_import_stage("finalizing")
        values = (
            ("_workspace_import_title_var", "完成保存"),
            ("_workspace_import_subtitle_var", "资料已安全保存到项目。"),
            ("_workspace_import_state_var", "已完成"),
            ("_workspace_import_name_var", "项目文件已更新"),
            ("_workspace_import_safety_title_var", "资料已安全保存"),
            ("_workspace_import_safety_text_var", "现在可以继续使用项目中的资料。"),
        )
        for attribute, value in values:
            variable = getattr(self, attribute, None)
            if variable is not None:
                variable.set(value)
        cancel_button = getattr(self, "_workspace_import_cancel_button", None)
        if cancel_button is not None:
            cancel_button.configure(text="已完成", state="disabled")
        self._draw_workspace_import_progress(1.0)

        previous_job = getattr(self, "_workspace_import_success_job", None)
        if previous_job is not None:
            try:
                self.root.after_cancel(previous_job)
            except Exception:
                pass

        def close_after_success() -> None:
            self._workspace_import_success_job = None
            self._close_workspace_import_progress(token=token, force=True)

        self._workspace_import_success_job = self.root.after(700, close_after_success)

    def _close_workspace_import_progress(self, *, token: int | None = None, force: bool = False) -> None:
        if token is not None and token != getattr(self, "_workspace_import_token", None):
            return
        if not force and getattr(self, "_workspace_import_phase", "") == "finalizing":
            return
        success_job = getattr(self, "_workspace_import_success_job", None)
        if success_job is not None:
            try:
                self.root.after_cancel(success_job)
            except Exception:
                pass
        self._workspace_import_success_job = None
        job = getattr(self, "_workspace_import_animation_job", None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
        self._workspace_import_animation_job = None
        window = getattr(self, "_workspace_import_window", None)
        self._workspace_import_window = None
        self._workspace_import_token = None
        self._workspace_import_started_at = None
        self._workspace_import_progress_canvas = None
        self._workspace_import_stage_labels = []
        self._workspace_import_cancel_button = None
        if window is not None:
            try:
                window.grab_release()
            except Exception:
                pass
            try:
                window.destroy()
            except Exception:
                pass

    @staticmethod
    def _workspace_import_cancelled_exception(exc: Exception) -> bool:
        """Match the backend cancellation type without making it a GUI import dependency."""

        try:
            module = importlib.import_module("hr_toolkit.project_store")
        except Exception:
            return False
        cancelled_type = getattr(module, "ImportCancelled", None)
        return isinstance(cancelled_type, type) and isinstance(exc, cancelled_type)

    def _workspace_run_write(
        self,
        action: str,
        callback,
        *,
        progress_target: Path | None = None,
        on_success=None,
        on_cancelled=None,
        on_error=None,
    ) -> None:
        if getattr(self, "_workspace_recovery_blocked", False):
            messagebox.showerror(
                "项目需要恢复",
                "上次资料保存没有完成安全恢复。为避免覆盖或遗漏资料，请关闭工具后重新打开当前项目，"
                "再继续操作。",
                parent=self.root,
            )
            return
        if self._workspace_write_in_progress():
            messagebox.showinfo(
                "资料正在保存",
                "请等待当前资料保存完成后再继续。",
                parent=self.root,
            )
            return
        store = self.project_store
        if store is None:
            return
        self._workspace_write_token += 1
        token = self._workspace_write_token
        generation = self._workspace_project_generation
        cancel_event = threading.Event()
        self._workspace_write_tasks[token] = (cancel_event, store)
        if not hasattr(self, "_workspace_write_progress"):
            self._workspace_write_progress = {}
        if not hasattr(self, "_workspace_write_callbacks"):
            self._workspace_write_callbacks = {}
        self._workspace_write_callbacks[token] = (on_success, on_cancelled, on_error)
        self._update_workspace_action_states()
        if progress_target is not None:
            self._open_workspace_import_progress(token, progress_target)

        def report_progress(snapshot) -> None:
            tracked = self._workspace_write_tasks.get(token)
            if tracked is None or tracked[1] is not store:
                return
            lock = getattr(self, "_workspace_write_progress_lock", None)
            if lock is None:
                self._workspace_write_progress[token] = snapshot
                return
            with lock:
                self._workspace_write_progress[token] = snapshot

        def worker() -> None:
            try:
                if progress_target is None:
                    result = callback(cancel_event.is_set)
                else:
                    result = callback(cancel_event.is_set, report_progress)
            except Exception as exc:
                if self._workspace_import_cancelled_exception(exc):
                    self._workspace_queue.put(
                        ("write_cancelled", generation, (token, store, action, exc))
                    )
                    return
                latest = self._workspace_latest_progress(token)
                phase = str(self._workspace_progress_field(latest, "phase", "")) if latest is not None else ""
                if phase == "finalizing":
                    try:
                        store.refresh()
                    except Exception as recovery_exc:
                        self._workspace_queue.put(
                            (
                                "write_recovery_blocked",
                                generation,
                                (token, store, action, exc, recovery_exc),
                            )
                        )
                        return
                    self._workspace_queue.put(
                        ("write_recovered", generation, (token, store, action, exc))
                    )
                    return
                self._workspace_queue.put(
                    ("write_error", generation, (token, store, action, exc))
                )
                return
            self._workspace_queue.put(("write_changed", generation, (token, store, result)))

        threading.Thread(target=worker, daemon=True).start()

    def _import_workspace_files(self) -> None:
        target = self._workspace_selected_target()
        store = self.project_store
        if target is None or self.current_project_path is None or store is None:
            return
        batch_target = self._workspace_batch_for_target(target)
        selected = self._askopenfilenames(title="选择要导入项目的文件")
        if not selected:
            return
        paths = [Path(path) for path in selected]
        if batch_target is not None:
            batch_id, category = batch_target
            callback = lambda cancelled, on_progress: store.import_sources(
                batch_id,
                paths,
                category=category,
                role="workspace",
                cancelled=cancelled,
                on_progress=on_progress,
            )
        else:
            callback = lambda cancelled, on_progress: store.import_to_directory(
                target,
                paths,
                cancelled=cancelled,
                on_progress=on_progress,
            )
        self._workspace_run_write("导入文件", callback, progress_target=target)

    def _import_workspace_folder(self) -> None:
        target = self._workspace_selected_target()
        store = self.project_store
        if target is None or self.current_project_path is None or store is None:
            return
        batch_target = self._workspace_batch_for_target(target)
        selected = self._askdirectory(title="选择要导入项目的文件夹")
        if not selected:
            return
        source_dir = Path(selected)
        if batch_target is not None:
            batch_id, category = batch_target
            callback = lambda cancelled, on_progress: store.import_sources(
                batch_id,
                [source_dir],
                category=category,
                role="workspace",
                cancelled=cancelled,
                on_progress=on_progress,
            )
        else:
            callback = lambda cancelled, on_progress: store.import_to_directory(
                target,
                [source_dir],
                cancelled=cancelled,
                on_progress=on_progress,
            )
        self._workspace_run_write("导入文件夹", callback, progress_target=target)

    def _create_workspace_folder(self) -> None:
        target = self._workspace_selected_target()
        store = self.project_store
        if target is None or store is None:
            return
        name = simpledialog.askstring(
            "新建文件夹",
            "文件夹名称",
            parent=self.root,
        )
        if name is None or not name.strip():
            return
        self._workspace_run_write(
            "新建文件夹",
            lambda _cancelled: store.new_folder(target, name.strip()),
        )

    def _build_history_view(self, root_frame) -> None:
        self._history_view = ttk.Frame(root_frame, style="Content.TFrame")
        history_vscroll = ttk.Scrollbar(self._history_view, orient=VERTICAL)
        history_vscroll.pack(side=RIGHT, fill=Y)
        self._history_canvas = Canvas(
            self._history_view,
            bg=COLOR_BG,
            highlightthickness=0,
            bd=0,
            yscrollcommand=history_vscroll.set,
        )
        self._history_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        history_vscroll.configure(command=self._history_canvas.yview)
        content = ttk.Frame(
            self._history_canvas,
            padding=self._responsive_content_padding(),
            style="Content.TFrame",
        )
        self._history_canvas_window = self._history_canvas.create_window(
            (0, 0),
            window=content,
            anchor="nw",
        )

        def _sync_history_canvas(_event=None) -> None:
            self._history_canvas.configure(scrollregion=self._history_canvas.bbox("all"))
            viewport_width = self._history_canvas.winfo_width()
            if viewport_width > 1:
                self._history_canvas.itemconfigure(self._history_canvas_window, width=viewport_width)

        content.bind("<Configure>", _sync_history_canvas, add="+")
        self._history_canvas.bind("<Configure>", _sync_history_canvas, add="+")

        header = ttk.Frame(content, style="Content.TFrame")
        header.pack(fill="x")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="资料追溯", style="Eyebrow.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="旧版记录", style="Title.TLabel").grid(row=1, column=0, sticky="w", pady=self._pad(5, 0))
        header_actions = ttk.Frame(header, style="Content.TFrame")
        header_actions.grid(row=0, column=1, rowspan=2, sticky="ne")
        self.history_rebuild_button = CodexButton(
            header_actions,
            text="重新整理记录",
            command=self._rebuild_history_index,
            variant="link",
            width=112,
        )
        self.history_rebuild_button.pack(side=LEFT, padx=self._pad(0, 6))
        CodexButton(
            header_actions,
            text="打开全部归档资料",
            command=self._open_history_root,
            width=150,
        ).pack(side=LEFT)
        Label(
            content,
            text="这里保留升级前的历史资料；新处理的资料和结果请直接在右侧“项目文件”中查看。",
            bg=COLOR_BG,
            fg=COLOR_MUTED,
            font=self.base_font,
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=self._pad(8, 16))

        filter_card = RoundedCard(content, padding=(18, 14, 18, 14))
        filter_card.pack(fill="x")
        filters = filter_card.inner
        filters.columnconfigure(1, weight=1)
        ttk.Label(filters, text="查找文件名", style="App.TLabel").grid(row=0, column=0, sticky="w")
        search_entry = ttk.Entry(filters, textvariable=self.history_search, style="App.TEntry")
        search_entry.grid(row=0, column=1, sticky="ew", padx=self._pad(10, 12))
        search_entry.bind("<Return>", lambda _event: self._apply_history_filters())
        self.history_tool_combo = ttk.Combobox(
            filters,
            textvariable=self.history_tool_filter,
            values=(HISTORY_TOOL_FILTER_ALL, *(label for _tool_id, label in TOOL_NAV_ITEMS)),
            state="readonly",
            width=16,
            style="App.TCombobox",
        )
        self.history_tool_combo.grid(row=0, column=2, padx=self._pad(0, 10))
        self.history_tool_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_history_filters())
        self.history_date_combo = ttk.Combobox(
            filters,
            textvariable=self.history_date_filter,
            values=HISTORY_DATE_FILTERS,
            state="readonly",
            width=10,
            style="App.TCombobox",
        )
        self.history_date_combo.grid(row=0, column=3, padx=self._pad(0, 10))
        self.history_date_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_history_filters())
        CodexButton(
            filters,
            text="查找",
            command=self._apply_history_filters,
            variant="tonal",
            width=70,
            min_width=64,
        ).grid(row=0, column=4)

        list_card = RoundedCard(content, padding=(14, 12, 14, 10), fill_height=True, min_height=285)
        list_card.pack(fill=BOTH, expand=True, pady=self._pad(14, 0))
        list_body = list_card.inner
        list_body.grid_rowconfigure(1, weight=1)
        list_body.grid_columnconfigure(0, weight=1)
        self.history_message_label = Label(
            list_body,
            textvariable=self.history_message,
            bg=COLOR_SURFACE,
            fg=COLOR_MUTED,
            font=self.small_font,
            anchor="w",
        )
        self.history_message_label.grid(row=0, column=0, sticky="ew", pady=self._pad(0, 8))

        tree_frame = ttk.Frame(list_body, style="InputWrap.TFrame")
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        columns = ("time", "tool", "status", "inputs", "outputs")
        self.history_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
            style="History.Treeview",
            height=9,
        )
        headings = {
            "time": "处理时间",
            "tool": "功能",
            "status": "状态",
            "inputs": "上传资料",
            "outputs": "结果",
        }
        for column, title in headings.items():
            self.history_tree.heading(column, text=title, anchor="w")
        self.history_tree.column("time", width=148, minwidth=138, stretch=False, anchor="w")
        self.history_tree.column("tool", width=142, minwidth=110, stretch=False, anchor="w")
        self.history_tree.column("status", width=82, minwidth=72, stretch=False, anchor="w")
        self.history_tree.column("inputs", width=220, minwidth=135, stretch=True, anchor="w")
        self.history_tree.column("outputs", width=190, minwidth=135, stretch=True, anchor="w")
        self.history_tree.tag_configure("success", foreground=COLOR_SUCCESS)
        self.history_tree.tag_configure("failed", foreground=COLOR_DANGER)
        self.history_tree.tag_configure("stopped", foreground=COLOR_WARNING)
        self.history_tree.tag_configure("running", foreground=COLOR_PRIMARY)
        tree_scroll = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.history_tree.yview)
        tree_hscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=tree_scroll.set, xscrollcommand=tree_hscroll.set)
        self.history_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.grid(row=0, column=1, sticky="ns")
        tree_hscroll.grid(row=1, column=0, sticky="ew")
        self.history_tree.bind("<<TreeviewSelect>>", self._on_history_selected)
        self.history_tree.bind("<Double-1>", lambda _event: self._open_selected_history_output())

        pager = ttk.Frame(list_body, style="InputWrap.TFrame")
        pager.grid(row=2, column=0, sticky="ew", pady=self._pad(9, 0))
        self.history_previous_button = CodexButton(
            pager,
            text="上一页",
            command=lambda: self._change_history_page(-1),
            variant="link",
            width=64,
            min_width=56,
            height=26,
        )
        self.history_previous_button.pack(side=LEFT)
        Label(
            pager,
            textvariable=self.history_page_text,
            bg=COLOR_SURFACE,
            fg=COLOR_FAINT,
            font=self.small_font,
        ).pack(side=LEFT, padx=self._pad(8, 8))
        self.history_next_button = CodexButton(
            pager,
            text="下一页",
            command=lambda: self._change_history_page(1),
            variant="link",
            width=64,
            min_width=56,
            height=26,
        )
        self.history_next_button.pack(side=LEFT)
        CodexButton(
            pager,
            text="打开回收站",
            command=self._open_history_trash,
            variant="link",
            width=92,
            min_width=82,
            height=26,
        ).pack(side=RIGHT)

        detail_card = RoundedCard(content, padding=(18, 13, 18, 14))
        detail_card.pack(fill="x", pady=self._pad(14, 0))
        detail_header = ttk.Frame(detail_card.inner, style="InputWrap.TFrame")
        detail_header.pack(fill="x")
        ttk.Label(detail_header, textvariable=self.history_detail_title, style="CardTitle.TLabel").pack(side=LEFT)
        self.history_delete_button = CodexButton(
            detail_header,
            text="移到回收站",
            command=self._move_selected_history_to_trash,
            variant="link",
            width=92,
            min_width=82,
            height=26,
        )
        self.history_delete_button.pack(side=RIGHT)
        self.history_detail_label = Label(
            detail_card.inner,
            textvariable=self.history_detail_text,
            bg=COLOR_SURFACE,
            fg=COLOR_MUTED,
            font=self.small_font,
            anchor="w",
            justify="left",
            wraplength=self._px(720),
        )
        self.history_detail_label.pack(fill="x", pady=self._pad(7, 10))
        detail_actions = ttk.Frame(detail_card.inner, style="InputWrap.TFrame")
        detail_actions.pack(fill="x")
        self.history_open_output_button = CodexButton(
            detail_actions,
            text="打开结果文件夹",
            command=self._open_selected_history_output,
            variant="primary",
            width=138,
        )
        self.history_open_output_button.pack(side=LEFT)
        self.history_open_input_button = CodexButton(
            detail_actions,
            text="查看上传资料",
            command=self._open_selected_history_input,
            width=124,
        )
        self.history_open_input_button.pack(side=LEFT, padx=self._pad(10, 0))
        self.history_reuse_button = CodexButton(
            detail_actions,
            text="再次使用这些资料",
            command=self._reuse_selected_history,
            width=150,
        )
        self.history_reuse_button.pack(side=LEFT, padx=self._pad(10, 0))
        self._set_history_detail_buttons(False, False, False, False)

        def _resize_history(_event=None) -> None:
            width = max(content.winfo_width(), self._px(640))
            self.history_detail_label.configure(wraplength=max(self._px(320), width - self._px(80)))

        content.bind("<Configure>", _resize_history, add="+")

        history_scroll_tag = f"HRToolkitHistoryScroll{id(self)}"

        def _history_scroll_units(event) -> int:
            if getattr(event, "num", None) == 4:
                return -3
            if getattr(event, "num", None) == 5:
                return 3
            delta = getattr(event, "delta", 0)
            if not delta:
                return 0
            return -max(1, int(abs(delta) / 120)) if delta > 0 else max(1, int(abs(delta) / 120))

        def _on_history_scroll(event):
            units = _history_scroll_units(event)
            if not units:
                return None
            if event.widget is self.history_tree:
                top, bottom = self.history_tree.yview()
                can_scroll_tree = (units < 0 and top > 0) or (units > 0 and bottom < 1.0)
                if can_scroll_tree:
                    self.history_tree.yview_scroll(units, "units")
                    return "break"
            self._history_canvas.yview_scroll(units, "units")
            return "break"

        for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            try:
                self._history_canvas.bind_class(history_scroll_tag, sequence, _on_history_scroll)
            except Exception:
                pass

        def _apply_history_scroll_tag(widget) -> None:
            try:
                tags = list(widget.bindtags())
                if history_scroll_tag not in tags:
                    if widget is self.history_tree:
                        widget.bindtags((history_scroll_tag, *tags))
                    else:
                        widget.bindtags((*tags, history_scroll_tag))
            except Exception:
                pass
            for child in widget.winfo_children():
                _apply_history_scroll_tag(child)

        _apply_history_scroll_tag(self._history_canvas)
        _apply_history_scroll_tag(content)
        self.root.after_idle(_sync_history_canvas)

    def _show_history_view(self) -> None:
        if not hasattr(self, "_history_view"):
            self._build_history_view(self._main_view_host)
        if self.current_view == "history":
            self._refresh_history()
            return
        self.current_view = "history"
        self._tool_view.pack_forget()
        self._history_view.pack(side=RIGHT, fill=BOTH, expand=True)
        self._refresh_nav_buttons()
        self._refresh_history()

    def _show_tool_view(self) -> None:
        if self.current_view == "tool":
            return
        self.current_view = "tool"
        self._history_view.pack_forget()
        self._tool_view.pack(side=RIGHT, fill=BOTH, expand=True)
        self._refresh_nav_buttons()
        if hasattr(self, "_sync_right_canvas_window"):
            self.root.after_idle(self._sync_right_canvas_window)

    def _apply_history_filters(self) -> None:
        self._history_page = 0
        self._refresh_history()

    def _change_history_page(self, delta: int) -> None:
        next_page = max(0, self._history_page + delta)
        if next_page * HISTORY_PAGE_SIZE >= self._history_total and delta > 0:
            return
        self._history_page = next_page
        self._refresh_history()

    def _refresh_history(self) -> None:
        if not hasattr(self, "history_tree"):
            return
        selected_id = self._history_selected_task.summary.id if self._history_selected_task is not None else None
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        self._history_selected_task = None
        self._reset_history_detail()
        store = self.history_store
        if store is None:
            self.history_message.set("历史记录暂时无法读取，您的原文件和已有结果不会被删除。")
            if self._history_init_error:
                self.history_detail_text.set(f"请联系管理员检查资料库位置。原因：{self._history_init_error}")
            self._history_total = 0
            self._update_history_pager()
            return
        try:
            tool_label = self.history_tool_filter.get().strip()
            tool_id = next((key for key, label in TOOL_NAV_ITEMS if label == tool_label), None)
            tasks, total = store.list_tasks(
                search=self.history_search.get(),
                tool_id=tool_id,
                started_after=self._history_started_after(),
                limit=HISTORY_PAGE_SIZE,
                offset=self._history_page * HISTORY_PAGE_SIZE,
            )
        except Exception as exc:
            runlog.log_exception("读取历史记录失败", exc)
            self.history_message.set("历史记录暂时无法读取，资料仍保存在本机。可以点击“重新整理记录”。")
            self.history_detail_text.set(str(exc))
            self._history_total = 0
            self._update_history_pager()
            return
        self._history_total = total
        if not tasks:
            if self.history_search.get().strip() or self.history_tool_filter.get() != HISTORY_TOOL_FILTER_ALL or self.history_date_filter.get() != HISTORY_DATE_FILTER_ALL:
                self.history_message.set("没有找到相关记录，可以清除查找内容或更换时间范围。")
            else:
                self.history_message.set("没有找到旧版记录。新处理的资料和结果请在右侧“项目文件”中查看。")
            self._update_history_pager()
            return
        try:
            storage = store.storage_stats()
            storage_text = (
                f" · 已留存 {self._format_history_size(storage['total_bytes'])}"
                f" · 磁盘可用 {self._format_history_size(storage['free_bytes'])}"
            )
            if storage["trash_bytes"]:
                storage_text += f"（回收站 {self._format_history_size(storage['trash_bytes'])}）"
        except Exception:
            storage_text = ""
        self.history_message.set(f"共找到 {total} 次处理记录{storage_text}")
        for task in tasks:
            self.history_tree.insert(
                "",
                END,
                iid=task.id,
                values=(
                    self._format_history_list_time(task.started_at),
                    task.tool_name,
                    HISTORY_STATUS_LABELS.get(task.status, task.status),
                    self._history_names_text(task.input_names, empty="未归档上传资料"),
                    self._history_names_text(task.output_names, empty="暂无完整结果"),
                ),
                tags=(task.status,),
            )
        children = self.history_tree.get_children()
        target = selected_id if selected_id in children else children[0]
        self.history_tree.selection_set(target)
        self.history_tree.focus(target)
        self.history_tree.see(target)
        self._load_history_detail(target)
        self._update_history_pager()

    def _update_history_pager(self) -> None:
        pages = max(1, (self._history_total + HISTORY_PAGE_SIZE - 1) // HISTORY_PAGE_SIZE)
        self._history_page = min(self._history_page, pages - 1)
        self.history_page_text.set(f"第 {self._history_page + 1} / {pages} 页")
        self.history_previous_button.config(state="normal" if self._history_page > 0 else "disabled")
        has_next = (self._history_page + 1) * HISTORY_PAGE_SIZE < self._history_total
        self.history_next_button.config(state="normal" if has_next else "disabled")

    def _on_history_selected(self, _event=None) -> None:
        selection = self.history_tree.selection()
        if selection:
            self._load_history_detail(selection[0])

    def _load_history_detail(self, task_id: str) -> None:
        if self.history_store is None:
            return
        try:
            detail = self.history_store.get_task(task_id)
        except Exception as exc:
            self.history_detail_title.set("这条记录暂时无法读取")
            self.history_detail_text.set(str(exc))
            self._set_history_detail_buttons(False, False, False, False)
            return
        if detail is None:
            self._reset_history_detail()
            return
        self._history_selected_task = detail
        status = HISTORY_STATUS_LABELS.get(detail.summary.status, detail.summary.status)
        self.history_detail_title.set(f"{detail.summary.tool_name} · {status}")
        input_text = self._history_names_text(detail.summary.input_names, empty="未归档上传资料", full=True)
        output_text = self._history_names_text(detail.summary.output_names, empty="未生成完整结果", full=True)
        lines = [
            f"处理时间：{self._format_history_time(detail.summary.started_at)}",
            f"上传资料：{input_text}",
            f"处理结果：{output_text}",
        ]
        if detail.summary.status == "failed":
            lines.append("说明：上传资料已保存，但本次没有正常生成完整结果。")
        elif detail.summary.status == "stopped":
            lines.append("说明：这次处理没有正常完成，可以再次使用已保存的资料。")
        if detail.summary.error_message and detail.summary.status in {"failed", "stopped"}:
            lines.append(f"原因：{detail.summary.error_message}")
        self.history_detail_text.set("\n".join(lines))
        can_reuse = detail.summary.tool_id != "folder_rename" and bool(detail.inputs)
        still_finishing = detail.summary.id in self._history_task_by_token.values()
        self._set_history_detail_buttons(
            bool(detail.outputs) and not still_finishing,
            bool(detail.inputs) and detail.summary.status != "running" and not still_finishing,
            can_reuse and detail.summary.status != "running" and not still_finishing,
            detail.summary.status != "running" and not still_finishing,
        )

    def _reset_history_detail(self) -> None:
        self.history_detail_title.set("选择一条记录查看详情")
        self.history_detail_text.set("这里用于查看升级前由旧版本保存的处理记录。")
        if hasattr(self, "history_open_output_button"):
            self._set_history_detail_buttons(False, False, False, False)

    def _set_history_detail_buttons(self, output: bool, inputs: bool, reuse: bool, delete: bool) -> None:
        self.history_open_output_button.config(state="normal" if output else "disabled")
        self.history_open_input_button.config(state="normal" if inputs else "disabled")
        self.history_reuse_button.config(state="normal" if reuse else "disabled")
        self.history_delete_button.config(state="normal" if delete else "disabled")

    def _open_selected_history_output(self) -> None:
        detail = self._history_selected_task
        if detail is None or not detail.outputs:
            return
        if not detail.output_dir.is_dir() or not any(item.archived_path.is_file() for item in detail.outputs):
            messagebox.showwarning(
                "结果资料已被移动",
                "这次处理的结果资料已被移动或删除，请联系管理员检查历史资料库。",
                parent=self.root,
            )
            return
        try:
            open_path(detail.output_dir)
        except Exception as exc:
            messagebox.showerror("无法打开结果", str(exc), parent=self.root)

    def _open_selected_history_input(self) -> None:
        detail = self._history_selected_task
        if detail is None or not detail.inputs:
            return
        if not detail.input_dir.is_dir() or not any(item.archived_path.is_file() for item in detail.inputs):
            messagebox.showwarning(
                "上传资料已被移动",
                "这次处理的上传资料已被移动或删除，请联系管理员检查历史资料库。",
                parent=self.root,
            )
            return
        try:
            open_path(detail.input_dir)
        except Exception as exc:
            messagebox.showerror("无法打开资料", str(exc), parent=self.root)

    def _open_history_root(self) -> None:
        if self.history_store is None:
            messagebox.showwarning("历史记录不可用", "资料库暂时无法打开，请联系管理员检查保存位置。")
            return
        try:
            open_path(self.history_store.records_dir)
        except Exception as exc:
            messagebox.showerror("无法打开归档资料", str(exc), parent=self.root)

    def _open_history_trash(self) -> None:
        if self.history_store is None:
            messagebox.showwarning("历史记录不可用", "资料库暂时无法打开，请联系管理员检查保存位置。")
            return
        try:
            open_path(self.history_store.trash_dir)
        except Exception as exc:
            messagebox.showerror("无法打开回收站", str(exc), parent=self.root)

    def _move_selected_history_to_trash(self) -> None:
        detail = self._history_selected_task
        if detail is None or self.history_store is None:
            return
        if detail.summary.id in self._history_task_by_token.values():
            messagebox.showwarning(
                "正在安全结束",
                "这次处理的后台工作还在安全结束，请稍等片刻再移动。",
                parent=self.root,
            )
            return
        if not messagebox.askyesno(
            "移到回收站",
            "这次处理的上传资料和结果会移到 HRToolkit 回收站，不会立即永久删除；如需恢复请联系管理员。是否继续？",
            parent=self.root,
        ):
            return
        try:
            self.history_store.move_to_trash(detail.summary.id)
        except Exception as exc:
            messagebox.showerror("无法移动", str(exc), parent=self.root)
            return
        self._history_selected_task = None
        self._refresh_history()

    def _rebuild_history_index(self) -> None:
        if self.history_store is None:
            messagebox.showwarning("历史记录不可用", "资料库暂时无法读取，请联系管理员检查保存位置。")
            return
        self.history_rebuild_button.config(state="disabled")
        self.history_message.set("正在整理历史记录，请稍候…")

        def worker() -> None:
            try:
                count = self.history_store.rebuild_index_from_manifests()
            except Exception as exc:
                self.history_queue.put(("rebuild_error", exc))
                return
            self.history_queue.put(("rebuild_done", count))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_history_queue(self) -> None:
        try:
            while True:
                status, payload = self.history_queue.get_nowait()
                if hasattr(self, "history_rebuild_button"):
                    self.history_rebuild_button.config(state="normal")
                if status == "rebuild_done":
                    self._refresh_history()
                    messagebox.showinfo(
                        "整理完成",
                        f"历史记录已经整理完成，恢复或修复了 {int(payload or 0)} 条记录。",
                        parent=self.root,
                    )
                elif status == "rebuild_error":
                    self.history_message.set("整理没有完成，资料仍保存在本机。")
                    messagebox.showerror("整理失败", str(payload), parent=self.root)
        except queue.Empty:
            pass
        self.root.after(200, self._poll_history_queue)

    def _reuse_selected_history(self) -> None:
        detail = self._history_selected_task
        if detail is None:
            return
        if self._tool_running:
            messagebox.showwarning("正在处理", "请先等待当前处理完成，再使用历史资料。", parent=self.root)
            return
        tool_id = detail.summary.tool_id
        if tool_id not in TOOL_NAV_LABELS:
            messagebox.showwarning("暂不支持", "这条旧记录暂时不能直接再次使用，可以先打开上传资料。", parent=self.root)
            return
        if tool_id == "folder_rename":
            messagebox.showinfo(
                "请先复制资料",
                "为了保护历史原件，文件夹改名记录不能直接再次处理。请先打开上传资料并复制到新的文件夹。",
                parent=self.root,
            )
            return
        self._show_tool_view()
        self._select_tool(tool_id)
        if tool_id == "personnel_change_merge" and detail.summary.mode in {"merge", "roster"}:
            target_index = 1 if detail.summary.mode == "roster" else 0
            self.change_tabs.select(target_index)
            self._on_change_tab_changed()
        elif tool_id == "archive_import" and detail.summary.mode in {"import", "export"}:
            target_index = 1 if detail.summary.mode == "export" else 0
            self.change_tabs.select(target_index)
            self._on_change_tab_changed()

        main_inputs = [item.archived_path for item in detail.inputs if item.role in {"input_path", "input_paths"} and item.archived_path.exists()]
        secondary: dict[str, list[Path]] = {}
        for item in detail.inputs:
            if item.role not in {"input_path", "input_paths"} and item.archived_path.exists():
                secondary.setdefault(item.role, []).append(item.archived_path)
        self.change_input_paths = None
        self.input_path.set("")
        self.summary_path.set("")
        if tool_id == "salary_split":
            if main_inputs:
                self.input_path.set(str(main_inputs[0]))
        elif main_inputs:
            self._set_change_input_paths(main_inputs)
        if secondary:
            role_paths = next(iter(secondary.values()))
            summary_value = role_paths[0] if len(role_paths) == 1 else Path(os.path.commonpath([str(path) for path in role_paths]))
            self.summary_path.set(str(summary_value))
        self._refresh_upload_card()
        self._update_summary_display()
        messagebox.showinfo("资料已带入", "以前保存的资料已经放回当前功能，请确认后重新处理。", parent=self.root)

    def _history_started_after(self) -> str | None:
        selected = self.history_date_filter.get()
        today = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
        if selected == "今天":
            start = today
        elif selected == "最近7天":
            start = today - timedelta(days=6)
        elif selected == "最近30天":
            start = today - timedelta(days=29)
        elif selected == "今年":
            start = today.replace(month=1, day=1)
        else:
            return None
        return start.astimezone(timezone.utc).isoformat(timespec="seconds")

    @staticmethod
    def _format_history_time(value: str) -> str:
        try:
            return datetime.fromisoformat(value).astimezone().strftime("%Y-%m-%d %H:%M")
        except (TypeError, ValueError):
            return value

    @staticmethod
    def _format_history_list_time(value: str) -> str:
        try:
            return datetime.fromisoformat(value).astimezone().strftime("%m-%d %H:%M")
        except (TypeError, ValueError):
            return value

    @staticmethod
    def _format_history_size(value: int) -> str:
        size = max(0, int(value))
        if size >= 1024**3:
            return f"{size / 1024**3:.1f} GB"
        if size >= 1024**2:
            return f"{size / 1024**2:.1f} MB"
        if size >= 1024:
            return f"{size / 1024:.1f} KB"
        return f"{size} B"

    @staticmethod
    def _history_names_text(names: tuple[str, ...], *, empty: str, full: bool = False) -> str:
        if not names:
            return empty
        if full:
            visible = "、".join(names[:5])
            return visible if len(names) <= 5 else f"{visible} 等 {len(names)} 个文件"
        return names[0] if len(names) == 1 else f"{names[0]} 等 {len(names)} 个"

    def _check_updates_on_startup(self) -> None:
        if not getattr(self, "_is_alive", True):
            return
        if not update_check_enabled():
            return
        self._start_update_check(manual=False)

    def _check_updates_manually(self) -> None:
        if not getattr(self, "_is_alive", True):
            return
        self._start_update_check(manual=True)

    def _start_update_check(self, manual: bool) -> None:
        if not getattr(self, "_is_alive", True):
            return
        if self.update_check_in_progress:
            if manual:
                # 静默检查进行中时用户点了“检查更新”：升级为手动检查并给出可见反馈
                self.manual_update_check_active = True
                self.update_check_dismissed = False
                if self.update_window is not None and self.update_window.winfo_exists():
                    self._focus_update_window()
                else:
                    self._show_update_checking_window()
            return
        self.update_check_in_progress = True
        self.manual_update_check_active = manual
        self.update_check_dismissed = False
        if hasattr(self, "check_update_button"):
            try:
                if self.check_update_button.winfo_exists():
                    self.check_update_button.config(state="disabled")
            except Exception:
                pass
        self._write_log("正在检查更新...")
        # 启动时的自动检查静默进行，只有确实存在新版本才打扰用户
        if manual:
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
        if not getattr(self, "_is_alive", True):
            return
        try:
            if not self.root.winfo_exists():
                return
        except Exception:
            return
        try:
            while True:
                status, payload = self.update_queue.get_nowait()
                if status == "no_update":
                    interactive = self.manual_update_check_active and not self.update_check_dismissed
                    self._finish_update_check()
                    self._write_log("已是最新版本。")
                    if interactive:
                        self._show_update_done_window()
                    else:
                        self._close_update_window()
                elif status == "check_error":
                    interactive = self.manual_update_check_active and not self.update_check_dismissed
                    self._finish_update_check()
                    self._write_log(f"更新检查失败，可继续使用：{payload}")
                    if interactive:
                        self._show_update_failure_window("检查更新失败", str(payload), exit_after=False)
                    else:
                        self._close_update_window()
                elif status == "available":
                    self._finish_update_check()
                    self._show_update_prompt(payload)
                elif status == "download_progress":
                    downloaded, total = payload
                    self._update_download_progress(downloaded, total)
                elif status == "download_ready":
                    self._finish_update_download(payload)
                elif status == "download_error":
                    self._handle_update_failure(payload)
                elif status == "manual_download_ready":
                    self._finish_manual_update(payload)
                elif status == "manual_download_error":
                    self._handle_manual_update_failure(payload)
        except queue.Empty:
            pass
        if getattr(self, "_is_alive", True):
            try:
                if self.root.winfo_exists():
                    self._poll_update_timer = self.root.after(150, self._poll_update_queue)
            except Exception:
                pass

    def _finish_update_check(self) -> None:
        self.update_check_in_progress = False
        self.manual_update_check_active = False
        if hasattr(self, "check_update_button"):
            try:
                if self.check_update_button.winfo_exists():
                    self.check_update_button.config(state="normal")
            except Exception:
                pass

    def _show_update_checking_window(self) -> None:
        self._show_update_progress_window(
            title="正在检查更新",
            detail="请稍候，正在确认是否有新版本。",
            indeterminate=True,
            close_command=self._dismiss_update_check,
        )

    def _dismiss_update_check(self) -> None:
        # 用户关闭“正在检查”窗口：检查在后台继续，结果只写日志，
        # 除非发现了新版本（那仍需要提示）
        self.update_check_dismissed = True
        self._close_update_window()

    def _close_update_window(self) -> None:
        if self.update_progress_job is not None:
            try:
                self.root.after_cancel(self.update_progress_job)
            except Exception:
                pass
            self.update_progress_job = None
        if self.update_window is not None and self.update_window.winfo_exists():
            self.update_window.grab_release()
            self.update_window.destroy()
        self.update_window = None
        self.update_progress_label = None
        self.update_progress_canvas = None
        self.update_progress_phase = 0.0
        self.update_progress_last_tick = None

    def _show_update_prompt(self, update: object | None) -> None:
        if not isinstance(update, UpdateInfo):
            return
        self.pending_update = update
        self._write_log(f"发现新版本：v{update.version}")
        self._write_log(f"下载地址：{update.file_url}")
        notes = list(update.notes) or ["本次发布未填写更新说明。"]
        if update.update_mode == "manual":
            self._show_update_message_window(
                title=f"发现新版本 v{update.version}",
                detail=(
                    "macOS 当前使用标准 DMG 手动更新。点击“下载 DMG”后，"
                    "请打开下载的文件并把 HRToolkit 拖到 Applications；本程序不会自动替换 .app。"
                ),
                notes=notes,
                primary_text="下载 DMG",
                primary_command=lambda: self._open_manual_update(update),
                secondary_text="稍后再说",
                secondary_command=self._defer_update,
                close_command=self._defer_update,
            )
            return
        if update.mandatory:
            self._show_update_message_window(
                title=f"发现新版本 v{update.version}",
                detail="这是必须安装的更新，更新完成后程序会自动重新打开。",
                notes=notes,
                primary_text="立即更新",
                primary_command=lambda: self._start_update_download(update),
                secondary_text="退出程序",
                secondary_command=self.root.destroy,
                close_command=self.root.destroy,
                escape_closes=False,
            )
        else:
            self._show_update_message_window(
                title=f"发现新版本 v{update.version}",
                detail="建议尽快更新。选择“稍后再说”可以继续使用当前版本，下次启动时会再次提醒。",
                notes=notes,
                primary_text="立即更新",
                primary_command=lambda: self._start_update_download(update),
                secondary_text="稍后再说",
                secondary_command=self._defer_update,
                close_command=self._defer_update,
            )

    def _open_manual_update(self, update: UpdateInfo) -> None:
        self._write_log("正在选择可用的 DMG 下载地址，优先使用 Gitee 国内源...")
        self._show_update_progress_window(
            title=f"准备下载 v{update.version}",
            detail="正在连接国内下载源；不可用时会自动尝试 GitHub 备用源。",
            indeterminate=True,
            close_command=self._close_update_window,
        )
        worker = threading.Thread(target=self._manual_update_worker, args=(update,), daemon=True)
        worker.start()

    def _manual_update_worker(self, update: UpdateInfo) -> None:
        try:
            download_url = resolve_download_url(update)
        except Exception as exc:
            self.update_queue.put(("manual_download_error", exc))
            return
        self.update_queue.put(("manual_download_ready", download_url))

    def _finish_manual_update(self, download_url: object | None) -> None:
        if not isinstance(download_url, str):
            return
        self._write_log(f"正在打开手动更新下载地址：{download_url}")
        try:
            open_path(download_url)
        except Exception as exc:
            self._handle_manual_update_failure(exc)
            return
        self._close_update_window()

    def _handle_manual_update_failure(self, exc: object | None) -> None:
        self._write_log(f"无法打开 DMG 下载地址：{exc}")
        self._show_update_failure_window("无法打开下载地址", str(exc), exit_after=False)

    def _defer_update(self) -> None:
        self._write_log("已选择稍后更新，下次启动时会再次提醒。")
        self._close_update_window()

    def _cancel_update_download(self) -> None:
        if self._update_download_cancel_event is not None:
            self._update_download_cancel_event.set()
        self._write_log("已取消更新包下载。")
        self._close_update_window()

    def _start_update_download(self, update: UpdateInfo) -> None:
        if self._tool_running or self._history_task_by_token:
            messagebox.showwarning(
                "请先完成当前处理",
                "当前资料还在处理或留存中。请等待完成后再更新，避免结果文件被中断。",
                parent=self.root,
            )
            return
        self._write_log(f"开始下载更新包：v{update.version}")
        self._write_log(f"下载地址：{update.file_url}")
        self._download_speed_anchor = None
        self._update_download_cancel_event = threading.Event()
        self._show_update_progress_window(
            title=f"正在下载 v{update.version}",
            detail="请不要关闭程序，下载完成后会自动开始安装。",
            indeterminate=False,
            close_command=self._cancel_update_download,
        )

        worker = threading.Thread(target=self._download_update_worker, args=(update,), daemon=True)
        worker.start()

    def _download_update_worker(self, update: UpdateInfo) -> None:
        def progress(downloaded: int, total: int) -> None:
            self.update_queue.put(("download_progress", (downloaded, total)))

        try:
            package_path = download_update_package(
                update,
                progress_callback=progress,
                cancel_event=self._update_download_cancel_event,
            )
        except UpdateCancelledError:
            return
        except Exception as exc:
            if self._update_download_cancel_event is not None and self._update_download_cancel_event.is_set():
                return
            self.update_queue.put(("download_error", exc))
            return
        if self._update_download_cancel_event is not None and self._update_download_cancel_event.is_set():
            return
        self.update_queue.put(("download_ready", package_path))

    def _update_download_progress(self, downloaded: int, total: int) -> None:
        now = time.monotonic()
        if self._download_speed_anchor is None:
            self._download_speed_anchor = (now, downloaded)
        anchor_time, anchor_bytes = self._download_speed_anchor
        elapsed = now - anchor_time
        speed_text = ""
        if elapsed > 0.8 and downloaded > anchor_bytes:
            speed_mb = (downloaded - anchor_bytes) / elapsed / 1024 / 1024
            speed_text = f"，{speed_mb:.1f} MB/s"

        downloaded_mb = downloaded / 1024 / 1024
        if total > 0:
            percent = min(downloaded / total * 100, 100)
            total_mb = total / 1024 / 1024
            text = f"已下载 {percent:.0f}%（{downloaded_mb:.1f}/{total_mb:.1f} MB{speed_text}）"
        else:
            percent = 0
            text = f"已下载 {downloaded_mb:.1f} MB{speed_text}"
        if self.update_progress_label is not None:
            self.update_progress_label.configure(text=text)
        self._set_update_progress(percent)

    def _finish_update_download(self, package_path: object | None) -> None:
        if not isinstance(package_path, Path):
            return
        if self._tool_running or self._history_task_by_token:
            if self.update_progress_label is not None:
                self.update_progress_label.configure(text="更新已下载，正在等待当前处理安全结束...")
            self.root.after(1000, lambda: self._finish_update_download(package_path))
            return
        if self.update_progress_label is not None:
            self.update_progress_label.configure(text="下载完成，正在启动安装程序...")
        self._set_update_progress(100)
        self._write_log("更新包下载完成，正在启动更新程序...")
        try:
            launch_update_replacement(package_path)
        except Exception as exc:
            self._handle_update_failure(exc)
            return
        if self.update_progress_label is not None:
            self.update_progress_label.configure(text="安装程序已启动，本窗口即将关闭。")
        self.root.after(700, self.root.destroy)

    def _handle_update_failure(self, exc: object | None) -> None:
        self._write_log(f"更新失败：{exc}")
        update = self.pending_update
        if not isinstance(update, UpdateInfo):
            self._show_update_failure_window(
                "更新失败",
                f"更新没有完成，程序将退出。\n\n原因：{exc}",
                exit_after=True,
            )
            return
        detail = f"更新没有完成，可以点击“重试”重新下载。\n\n原因：{exc}\n\n如果多次失败，请联系管理员。"
        retry = lambda: self._start_update_download(update)  # noqa: E731
        if update.mandatory:
            self._show_update_message_window(
                title="更新失败",
                detail=detail,
                primary_text="重试",
                primary_command=retry,
                secondary_text="退出程序",
                secondary_command=self.root.destroy,
                close_command=self.root.destroy,
                escape_closes=False,
            )
        else:
            self._show_update_message_window(
                title="更新失败",
                detail=detail,
                primary_text="重试",
                primary_command=retry,
                secondary_text="稍后再说",
                secondary_command=self._defer_update,
                close_command=self._defer_update,
            )

    def _show_update_done_window(self) -> None:
        self._show_update_message_window(
            title="已经是最新版本",
            detail=f"{APP_DISPLAY_NAME} v{__version__} 已经是最新版本，无需更新。",
            primary_text="确定",
            primary_command=self._close_update_window,
            width=380,
            close_command=self._close_update_window,
        )

    def _show_update_failure_window(self, title: str, detail: str, *, exit_after: bool) -> None:
        close_command = self.root.destroy if exit_after else self._close_update_window
        self._show_update_message_window(
            title=title,
            detail=detail,
            primary_text="退出程序" if exit_after else "知道了",
            primary_command=close_command,
            close_command=close_command,
            escape_closes=not exit_after,
        )

    def _show_update_progress_window(
        self,
        *,
        title: str,
        detail: str,
        indeterminate: bool,
        close_command,
    ) -> None:
        window, body, dialog_width = self._build_update_window(width=420, close_command=close_command)
        pad = self._px(24)
        content_width = dialog_width - pad * 2

        self._build_update_header(body, title=title, pad=pad)
        self._create_update_progress_bar(body, width=content_width, padx=pad, pady=self._pad(18, 0))
        self.update_progress_label = Label(
            body,
            text=detail,
            bg=UPDATE_DIALOG_BG,
            fg=UPDATE_DIALOG_MUTED,
            font=self.small_font,
            justify="left",
            wraplength=content_width,
        )
        self.update_progress_label.pack(anchor="w", padx=pad, pady=self._pad(10, 24))
        if indeterminate:
            self._start_indeterminate_update_progress()
        else:
            self._set_update_progress(0)
        self._finalize_update_window(window, dialog_width, close_command=close_command)

    def _show_update_message_window(
        self,
        *,
        title: str,
        detail: str,
        primary_text: str,
        primary_command,
        secondary_text: str | None = None,
        secondary_command=None,
        notes: list[str] | None = None,
        width: int = 420,
        close_command=None,
        escape_closes: bool = True,
    ) -> None:
        close_command = close_command or self._close_update_window
        window, body, dialog_width = self._build_update_window(width=width, close_command=close_command)
        pad = self._px(24)
        text_wrap_width = dialog_width - pad * 2

        self._build_update_header(body, title=title, pad=pad)
        Label(
            body,
            text=detail,
            bg=UPDATE_DIALOG_BG,
            fg=UPDATE_DIALOG_MUTED,
            font=self.base_font,
            justify="left",
            wraplength=text_wrap_width,
        ).pack(anchor="w", padx=pad, pady=self._pad(12, 0))
        if notes:
            self._build_update_notes(body, notes, pad=pad)

        button_row = Frame(body, bg=UPDATE_DIALOG_BG)
        button_row.pack(fill="x", padx=pad, pady=self._pad(22, 20))
        self._create_update_button(
            button_row,
            text=primary_text,
            command=primary_command,
            primary=True,
        ).pack(side=RIGHT)
        if secondary_text and secondary_command:
            self._create_update_button(
                button_row,
                text=secondary_text,
                command=secondary_command,
                primary=False,
            ).pack(side=RIGHT, padx=self._pad(0, 10))

        self._finalize_update_window(
            window,
            dialog_width,
            primary_command=primary_command,
            close_command=close_command if escape_closes else None,
        )

    def _build_update_header(self, body: Frame, *, title: str, pad: int) -> None:
        header = Frame(body, bg=UPDATE_DIALOG_BG)
        header.pack(fill="x", padx=pad, pady=self._pad(22, 0))
        icon = Canvas(header, width=self._px(44), height=self._px(44), bg=UPDATE_DIALOG_BG, highlightthickness=0)
        icon.pack(side=LEFT)
        self._draw_update_icon(icon)
        Label(
            header,
            text=title,
            bg=UPDATE_DIALOG_BG,
            fg=UPDATE_DIALOG_TEXT,
            font=(self.base_font[0], _font_size(13), "bold"),
        ).pack(side=LEFT, padx=self._pad(14, 0))

    def _build_update_notes(self, body: Frame, notes: list[str], *, pad: int) -> None:
        Label(
            body,
            text="更新内容",
            bg=UPDATE_DIALOG_BG,
            fg=UPDATE_DIALOG_MUTED,
            font=self.small_font,
        ).pack(anchor="w", padx=pad, pady=self._pad(14, 4))
        notes_frame = Frame(body, bg=UPDATE_DIALOG_NOTES_BG)
        notes_frame.pack(fill="x", padx=pad)
        text = Text(
            notes_frame,
            height=min(max(len(notes), 2), 6),
            wrap="word",
            bg=UPDATE_DIALOG_NOTES_BG,
            fg=UPDATE_DIALOG_TEXT,
            relief="flat",
            bd=0,
            padx=self._px(12),
            pady=self._px(10),
            font=self.base_font,
            highlightthickness=0,
        )
        if len(notes) > 6:
            scrollbar = ttk.Scrollbar(notes_frame, orient=VERTICAL, command=text.yview)
            scrollbar.pack(side=RIGHT, fill=Y)
            text.configure(yscrollcommand=scrollbar.set)
        text.pack(side=LEFT, fill=BOTH, expand=True)
        text.insert("1.0", "\n".join(f"· {line}" for line in notes))
        text.config(state="disabled")

    def _build_update_window(self, *, width: int, close_command) -> tuple[Toplevel, Frame, int]:
        self._close_update_window()
        scaled_width, _ = self._update_dialog_size(width, 0)
        window = Toplevel(self.root)
        self.update_window = window
        # 传递缩放系数：CodexButton 等自绘控件按所在顶层窗口取缩放，
        # 不设置的话高缩放屏上弹窗里的按钮会偏小
        setattr(window, "_hr_ui_scale", self.ui_scale)
        window.withdraw()
        window.title("软件更新")
        window.resizable(False, False)
        window.configure(bg=UPDATE_DIALOG_BG)
        window.transient(self.root)
        window.protocol("WM_DELETE_WINDOW", close_command or (lambda: None))
        body = Frame(window, bg=UPDATE_DIALOG_BG, width=scaled_width)
        body.pack(fill=BOTH, expand=True)
        return window, body, scaled_width

    def _finalize_update_window(
        self,
        window: Toplevel,
        width: int,
        *,
        primary_command=None,
        close_command=None,
    ) -> None:
        # 高度由内容决定，避免固定尺寸裁掉换行后的中文文本
        window.update_idletasks()
        height = min(window.winfo_reqheight(), max(1, self.root.winfo_screenheight() - self._px(72)))
        self._center_window(window, width, height)
        window.deiconify()
        try:
            window.grab_set()
        except Exception:
            pass
        window.focus_set()
        if primary_command is not None:
            window.bind("<Return>", lambda _event: primary_command())
        if close_command is not None:
            window.bind("<Escape>", lambda _event: close_command())

    def _focus_update_window(self) -> None:
        if self.update_window is not None and self.update_window.winfo_exists():
            self.update_window.lift()
            self.update_window.focus_force()

    def _create_update_progress_bar(self, parent: Frame, *, width: int, padx, pady) -> None:
        self.update_progress_width = width
        self.update_progress_canvas = Canvas(
            parent,
            width=self.update_progress_width,
            height=self._px(7),
            bg=UPDATE_DIALOG_BG,
            highlightthickness=0,
        )
        self.update_progress_canvas.pack(anchor="w", padx=padx, pady=pady)
        self._draw_round_rect(
            self.update_progress_canvas,
            0,
            self._pxf(1),
            self.update_progress_width,
            self._pxf(6),
            self._pxf(2.5),
            fill=UPDATE_DIALOG_TRACK,
        )

    def _set_update_progress(self, percent: float) -> None:
        canvas = self.update_progress_canvas
        if canvas is None:
            return
        canvas.delete("fill")
        width = max(0, min(self.update_progress_width * percent / 100, self.update_progress_width))
        if width <= 0:
            return
        self._draw_round_rect(
            canvas,
            0,
            self._pxf(1),
            width,
            self._pxf(6),
            self._pxf(2.5),
            fill=UPDATE_DIALOG_PRIMARY,
            tags=("fill",),
        )

    def _start_indeterminate_update_progress(self) -> None:
        if self.update_progress_job is not None:
            try:
                self.root.after_cancel(self.update_progress_job)
            except Exception:
                pass
            self.update_progress_job = None

        segment = min(self._pxf(76), self.update_progress_width * 0.22)
        gap = self._pxf(26)
        span = self.update_progress_width + segment + gap * 2
        speed = self._pxf(360)

        def tick() -> None:
            canvas = self.update_progress_canvas
            if canvas is None:
                return
            now = time.monotonic()
            previous_tick = self.update_progress_last_tick
            elapsed = 1 / 60 if previous_tick is None else max(0.0, min(now - previous_tick, 0.05))
            self.update_progress_last_tick = now
            self.update_progress_phase = (self.update_progress_phase + speed * elapsed) % span

            canvas.delete("fill")
            sweep_head = self.update_progress_phase - gap
            visible_segment = _indeterminate_progress_segment(self.update_progress_width, sweep_head, segment)
            if visible_segment is not None:
                x1, x2 = visible_segment
                self._draw_round_rect(
                    canvas,
                    x1,
                    self._pxf(1),
                    x2,
                    self._pxf(6),
                    self._pxf(2.5),
                    fill=UPDATE_DIALOG_PRIMARY,
                    tags=("fill",),
                )
            self.update_progress_job = self.root.after(16, tick)

        self.update_progress_phase = gap
        self.update_progress_last_tick = time.monotonic()
        tick()

    def _create_update_button(
        self,
        parent: Frame,
        *,
        text: str,
        command,
        primary: bool,
    ) -> Canvas:
        fill = UPDATE_DIALOG_PRIMARY if primary else UPDATE_DIALOG_SECONDARY
        active_fill = UPDATE_DIALOG_PRIMARY_ACTIVE if primary else UPDATE_DIALOG_SECONDARY_ACTIVE
        foreground = "#ffffff" if primary else UPDATE_DIALOG_TEXT
        font_spec = (self.base_font[0], _font_size(10), "bold")
        width = max(self._px(92), tkfont.Font(font=font_spec).measure(text) + self._px(40))
        height = self._px(32)
        button = Canvas(parent, width=width, height=height, bg=UPDATE_DIALOG_BG, highlightthickness=0, cursor="hand2")

        def paint(color: str) -> None:
            button.delete("all")
            self._draw_round_rect(button, 0, 0, width, height, self._pxf(10), fill=color)
            button.create_text(width / 2, height / 2, text=text, fill=foreground, font=font_spec)

        button.bind("<Enter>", lambda _event: paint(active_fill))
        button.bind("<Leave>", lambda _event: paint(fill))
        button.bind("<Button-1>", lambda _event: command())
        paint(fill)
        return button

    def _draw_update_icon(self, canvas: Canvas) -> None:
        # 与侧栏导航一致的线性图标风格：淡色圆底 + 下载箭头
        p = self._pxf
        canvas.create_oval(p(2), p(2), p(42), p(42), fill=UPDATE_DIALOG_ICON_BG, outline="")
        line = {"fill": COLOR_PRIMARY, "width": max(1.0, p(2.4)), "capstyle": "round"}
        canvas.create_line(p(22), p(12), p(22), p(26), **line)
        canvas.create_line(p(16), p(20.5), p(22), p(27), **line)
        canvas.create_line(p(28), p(20.5), p(22), p(27), **line)
        canvas.create_line(p(14), p(32), p(30), p(32), **line)

    def _draw_round_rect(self, canvas: Canvas, x1: float, y1: float, x2: float, y2: float, radius: float, **kwargs) -> None:
        radius = max(0, min(radius, (x2 - x1) / 2, (y2 - y1) / 2))
        tags = kwargs.pop("tags", ())
        fill = kwargs.pop("fill", "")
        outline = kwargs.pop("outline", "")
        canvas.create_rectangle(x1 + radius, y1, x2 - radius, y2, fill=fill, outline=outline, tags=tags, **kwargs)
        canvas.create_rectangle(x1, y1 + radius, x2, y2 - radius, fill=fill, outline=outline, tags=tags, **kwargs)
        canvas.create_oval(x1, y1, x1 + 2 * radius, y1 + 2 * radius, fill=fill, outline=outline, tags=tags, **kwargs)
        canvas.create_oval(x2 - 2 * radius, y1, x2, y1 + 2 * radius, fill=fill, outline=outline, tags=tags, **kwargs)
        canvas.create_oval(x1, y2 - 2 * radius, x1 + 2 * radius, y2, fill=fill, outline=outline, tags=tags, **kwargs)
        canvas.create_oval(x2 - 2 * radius, y2 - 2 * radius, x2, y2, fill=fill, outline=outline, tags=tags, **kwargs)

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
            self._show_tool_view()
            return
        self._show_tool_view()
        if self._tool_running:
            self._stop_tool_run()
        if self.current_tool == "personnel_change_merge":
            self._save_change_form_state(self.change_mode)
        if self.current_tool == "archive_import":
            self._save_archive_form_state(self.archive_mode)
        self.current_tool = tool_id
        if tool_id == "personnel_change_merge":
            self.change_mode = "merge"
            self._load_change_form_state("merge")
            if hasattr(self, "change_tabs"):
                self.change_tabs.select(0)
        if tool_id == "archive_import":
            self.archive_mode = "import"
            self._load_archive_form_state("import")
            if hasattr(self, "change_tabs"):
                self.change_tabs.select(0)
        self._refresh_nav_buttons()
        self.last_output_dir = None
        if tool_id not in {"personnel_change_merge", "archive_import"}:
            self.input_path.set("")
            self.summary_path.set("")
            self.change_input_paths = None
        self.rename_target_name.set("")
        self.rename_text.set("")
        self.rename_replacement_name.set("")
        if not self.output_dir_user_selected:
            self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
        self._set_tool_texts()
        self._update_project_output_controls()
        self._clear_log()
        self._write_log(self._initial_log_text())
        if hasattr(self, "workspace_tree") and self.workspace_scope.get() == WORKSPACE_SCOPE_TOOL:
            self._refresh_workspace_tree()

    def _save_change_form_state(self, mode: str) -> None:
        self.change_form_state[mode] = (self.input_path.get(), self.summary_path.get(), self.change_input_paths)

    def _load_change_form_state(self, mode: str) -> None:
        input_text, summary_text, input_paths = self.change_form_state.get(mode, ("", "", None))
        self.input_path.set(input_text)
        self.summary_path.set(summary_text)
        self.change_input_paths = input_paths

    def _save_archive_form_state(self, mode: str) -> None:
        self.archive_form_state[mode] = (self.input_path.get(), self.summary_path.get(), self.change_input_paths)

    def _load_archive_form_state(self, mode: str) -> None:
        input_text, summary_text, input_paths = self.archive_form_state.get(mode, ("", "", None))
        self.input_path.set(input_text)
        self.summary_path.set(summary_text)
        self.change_input_paths = input_paths

    def _on_change_tab_changed(self, _event=None) -> None:
        if self.current_tool not in {"personnel_change_merge", "archive_import"}:
            return
        if self.current_tool == "archive_import":
            selected_mode = "export" if self.change_tabs.index("current") == 1 else "import"
            if selected_mode == self.archive_mode:
                return
            self._save_archive_form_state(self.archive_mode)
            self.archive_mode = selected_mode
            self._load_archive_form_state(selected_mode)
        else:
            selected_mode = "roster" if self.change_tabs.index("current") == 1 else "merge"
            if selected_mode == self.change_mode:
                return
            self._save_change_form_state(self.change_mode)
            self.change_mode = selected_mode
            self._load_change_form_state(selected_mode)
        self._set_tool_texts()
        self.last_output_dir = None
        self._update_project_output_controls()
        self._clear_log()
        self._write_log(self._initial_log_text())
        if hasattr(self, "workspace_tree") and self.workspace_scope.get() == WORKSPACE_SCOPE_TOOL:
            self._refresh_workspace_tree()

    def _refresh_nav_buttons(self) -> None:
        for tool_id, item in self.nav_buttons.items():
            item.set_selected(self.current_view == "tool" and tool_id == self.current_tool)
        if hasattr(self, "history_nav_item"):
            self.history_nav_item.set_selected(self.current_view == "history")

    def _set_tool_texts(self) -> None:
        self.tool_group.set(TOOL_GROUP_LABELS.get(self.current_tool, "人员运营自动化"))
        multi_hint = "支持 .xlsx / .xls / ZIP / RAR / 7Z / TAR / 文件夹 · 可多选"
        if self.current_tool == "social_security":
            self.tool_title.set("社保明细与汇总")
            self.tool_description.set("选择社保缴费清单、压缩包或文件夹，再选择参保人员花名册，自动生成明细和汇总。")
            self.input_label.set("社保缴费清单")
            self.input_hint.set(multi_hint)
            self._input_drop_title = "选择缴费清单、压缩包或文件夹"
            self.choose_input_text.set("选择")
            self.summary_label.set("参保人员花名册")
            self.summary_button_text.set("选择花名册")
            self.run_button_text.set("生成报表")
        elif self.current_tool == "data_statistics":
            self.tool_title.set("考勤与周月报统计")
            self.tool_description.set("选择考勤结果、周报记录、月报记录，或包含这些文件的文件夹/压缩包，自动生成统计表和异常明细。")
            self.input_label.set("考勤与周月报数据")
            self.input_hint.set(multi_hint)
            self._input_drop_title = "选择考勤 / 周报 / 月报文件、压缩包或文件夹"
            self.choose_input_text.set("选择")
            self.summary_label.set("应汇报人员名单（可选）")
            self.summary_button_text.set("选择名单")
            self.run_button_text.set("生成统计")
        elif self.current_tool == "insurance_ledger":
            self.tool_title.set("保险台账与增减预警")
            self.tool_description.set("选择各保单人员清单、压缩包或文件夹，再选择需求6的人力资源分析表，自动生成保险台账。")
            self.input_label.set("保单人员清单")
            self.input_hint.set(multi_hint)
            self._input_drop_title = "选择保单清单、压缩包或文件夹"
            self.choose_input_text.set("选择")
            self.summary_label.set("人力资源分析表")
            self.summary_button_text.set("选择分析表")
            self.run_button_text.set("生成台账")
        elif self.current_tool == "salary_merge":
            self.tool_title.set("多月工资合并")
            self.tool_description.set("选择工资表文件、压缩包或文件夹；如已有汇总表，可一并选择后追加新月份。")
            self.input_label.set("工资表文件")
            self.input_hint.set(multi_hint)
            self._input_drop_title = "选择工资表、压缩包或文件夹"
            self.choose_input_text.set("选择")
            self.summary_label.set("已有汇总表（可选）")
            self.summary_button_text.set("选择汇总表")
            self.run_button_text.set("开始合并")
        elif self.current_tool == "personnel_change_merge":
            self.tool_title.set("异动表汇总与花名册")
            if self.change_mode == "roster":
                self.tool_description.set("选择异动汇总表和人力资源花名册，单独更新花名册。")
                self.input_label.set("异动汇总表")
                self.input_hint.set(multi_hint)
                self._input_drop_title = "选择异动汇总表、压缩包或文件夹"
                self.choose_input_text.set("选择汇总表")
                self.summary_label.set("人力资源花名册")
                self.summary_button_text.set("选择花名册")
                self.run_button_text.set("更新花名册")
            else:
                self.tool_description.set("选择异动表、压缩包或文件夹；如已有月度汇总表，可选择后按月份追加。")
                self.input_label.set("异动表文件")
                self.input_hint.set(multi_hint)
                self._input_drop_title = "选择异动表、压缩包或文件夹"
                self.choose_input_text.set("选择")
                self.summary_label.set("已有汇总表/文件夹（可选）")
                self.summary_button_text.set("选择汇总表")
                self.run_button_text.set("开始汇总")
        elif self.current_tool == "folder_rename":
            self.tool_title.set("人员资料文件夹改名")
            rename_mode = RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND)
            if rename_mode == MODE_EXCEL_BATCH:
                self.tool_description.set("选择人员资料目录和人员名单 Excel，按名单与目录顺序预览后确认改名。")
            else:
                self.tool_description.set("选择人员资料目录，先预览，再确认改名。")
            self.input_label.set("人员文件夹目录")
            self.input_hint.set("只处理所选目录第一层，原目录不会被修改")
            self._input_drop_title = "选择人员文件夹目录"
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("人员名单 Excel" if rename_mode == MODE_EXCEL_BATCH else "")
            self.summary_button_text.set("选择名单" if rename_mode == MODE_EXCEL_BATCH else "选择")
            self.run_button_text.set("预览")
        elif self.current_tool == "archive_import":
            self.tool_title.set("档案入库与档案表")
            if self.archive_mode == "export":
                self.tool_description.set("选择档案汇总表、压缩包或文件夹，按公司写入已有档案表；没有已有表时自动新建。")
                self.input_label.set("档案汇总表")
                self.input_hint.set(multi_hint)
                self._input_drop_title = "选择档案汇总表、压缩包或文件夹"
                self.choose_input_text.set("选择汇总表")
                self.summary_label.set("已有公司档案表（可选）")
                self.summary_button_text.set("选择档案表")
                self.run_button_text.set("生成档案表")
            else:
                self.tool_description.set("选择项目档案移交表、压缩包或文件夹；可选已有档案汇总表，不选则新建。")
                self.input_label.set("档案移交表")
                self.input_hint.set(multi_hint)
                self._input_drop_title = "选择移交表、压缩包或文件夹"
                self.choose_input_text.set("选择")
                self.summary_label.set("已有档案汇总表（可选）")
                self.summary_button_text.set("选择汇总表")
                self.run_button_text.set("开始入库")
        elif self.current_tool == "salary_split":
            self.tool_title.set("工资表按入职公司拆分")
            self.tool_description.set("选择一个包含“汇总表”和“明细表”的工资表，工具会按“入职公司”拆成多个公司文件。")
            self.input_label.set("工资表文件")
            self.input_hint.set("支持 .xlsx / .xls · 单个文件")
            self._input_drop_title = "选择工资表文件"
            self.choose_input_text.set("选择文件")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
            self.run_button_text.set("开始拆分")
        elif self.current_tool == "material_collector":
            self.tool_title.set("员工资料智能检索与打包")
            self.tool_description.set("支持按人员文件夹查找，也支持从无序平铺资料库建立 OCR 索引后按人员检索。")
            self.input_label.set("员工资料库路径（只读检索）")
            self.input_hint.set("仅做本地只读扫描，不复制原资料库，支持上万人超大资料库")
            self._input_drop_title = "选择员工资料库路径（只读扫描）"
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("员工名单文件（Excel）")
            self.summary_button_text.set("选择名单")
            self.run_button_text.set("开始打包")
        else:
            self.tool_title.set("该工具暂未实现")
            self.tool_description.set("请选择左侧已经可用的工具。")
            self.input_label.set("输入")
            self.input_hint.set("")
            self._input_drop_title = "选择输入文件"
            self.choose_input_text.set("选择")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
            self.run_button_text.set("开始")
        if hasattr(self, "summary_label_widget"):
            self._update_change_tabs_visibility()
            self._update_change_picker_buttons()
            self._update_summary_controls(apply_layout=False)
            self._update_output_controls(apply_layout=False)
            self._update_rename_controls(apply_layout=False)
            self._update_material_controls(apply_layout=False)
            self._update_stats_range_controls(apply_layout=False)
            if hasattr(self, "_apply_form_layout"):
                self._apply_form_layout()
        self._refresh_last_run_status()
        if hasattr(self, "_sync_right_canvas_window"):
            self.root.after_idle(self._sync_right_canvas_window)

    # ---------- 运行状态 ----------

    def _run_state_key(self) -> str:
        if self.current_tool == "personnel_change_merge":
            return f"personnel_change_merge:{self.change_mode}"
        if self.current_tool == "archive_import":
            return f"archive_import:{self.archive_mode}"
        return self.current_tool

    def _record_last_run(self, success: bool) -> None:
        stamp = datetime.now().strftime("%H:%M")
        self._last_run_results[self._run_state_key()] = (stamp, "成功" if success else "失败")
        self._refresh_last_run_status()

    def _refresh_last_run_status(self) -> None:
        record = self._last_run_results.get(self._run_state_key())
        if record is None:
            self.last_run_text.set("")
            self.last_run_state.set("")
            return
        stamp, state = record
        self.last_run_text.set(f"上次运行 {stamp} · ")
        self.last_run_state.set(state)
        if hasattr(self, "last_run_state_label"):
            self.last_run_state_label.configure(fg=COLOR_SUCCESS if state == "成功" else COLOR_DANGER)

    def _open_run_log(self) -> None:
        try:
            log_path = runlog.run_log_path()
        except Exception:
            return
        if not log_path.exists():
            messagebox.showinfo("暂无日志", "运行日志文件还未生成。", parent=self.root)
            return
        open_path(log_path)

    # ---------- 路径悬浮提示 ----------

    def _cancel_path_tooltip_job(self) -> None:
        job = getattr(self, "_path_tooltip_job", None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
        self._path_tooltip_job = None

    def _hide_path_tooltip(self) -> None:
        self._cancel_path_tooltip_job()
        tip = getattr(self, "_path_tooltip", None)
        if tip is not None:
            try:
                tip.destroy()
            except Exception:
                pass
        self._path_tooltip = None

    def _show_path_tooltip(self, widget, text: str) -> None:
        # macOS Aqua 对 overrideredirect 顶层窗口不渲染，改为窗口内浮层，跨平台可靠
        self._hide_path_tooltip()
        if not text:
            return
        try:
            if not widget.winfo_exists():
                return
            anchor_widget = widget.winfo_toplevel()
            x = widget.winfo_rootx() - anchor_widget.winfo_rootx() + self._px(12)
            y = widget.winfo_rooty() - anchor_widget.winfo_rooty() + widget.winfo_height() + self._px(4)
        except Exception:
            return
        tip = Label(
            anchor_widget,
            text=text,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=self.small_font,
            bd=0,
            padx=self._px(9),
            pady=self._px(5),
            highlightthickness=self._px(1),
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_BORDER,
        )
        tip.update_idletasks()
        window_width = max(anchor_widget.winfo_width(), 1)
        window_height = max(anchor_widget.winfo_height(), 1)
        tip_width = tip.winfo_reqwidth()
        tip_height = tip.winfo_reqheight()
        x = max(self._px(4), min(x, window_width - tip_width - self._px(8)))
        if y + tip_height > window_height - self._px(4):
            y = widget.winfo_rooty() - anchor_widget.winfo_rooty() - tip_height - self._px(4)
        tip.place(x=x, y=y)
        tip.lift()
        self._path_tooltip = tip

    def _bind_path_tooltip(self, widget, text_getter) -> None:
        """悬停约半秒后显示完整路径，移开或点击即收起。"""

        def on_enter(_event=None):
            self._cancel_path_tooltip_job()
            self._path_tooltip_job = self.root.after(
                450, lambda: self._show_path_tooltip(widget, text_getter())
            )

        def on_leave(_event=None):
            self._hide_path_tooltip()

        widget.bind("<Enter>", on_enter, add="+")
        widget.bind("<Leave>", on_leave, add="+")
        widget.bind("<Button-1>", on_leave, add="+")

    def _ellipsize(self, text: str, font_spec, max_width: float) -> str:
        if max_width <= 0:
            return ""
        cache = getattr(self, "_ellipsize_fonts", None)
        if cache is None:
            cache = self._ellipsize_fonts = {}
        key = tuple(font_spec)
        font = cache.get(key)
        if font is None:
            font = cache[key] = tkfont.Font(root=self.root, font=font_spec)
        if font.measure(text) <= max_width:
            return text
        while text and font.measure(text + "…") > max_width:
            text = text[:-1]
        return text + "…"

    def _update_summary_display(self) -> None:
        if not hasattr(self, "summary_display"):
            return
        text = self.summary_path.get().strip()
        if not text:
            self.summary_display.configure(text="未选择", fg=COLOR_FAINT, font=self.base_font)
            if hasattr(self, "clear_summary_button"):
                self.clear_summary_button.pack_forget()
            return
        name = Path(text).name or text
        self.summary_display.configure(text=name, fg=COLOR_TEXT, font=(self.base_font[0], _font_size(10), "bold"))
        if hasattr(self, "clear_summary_button"):
            self.clear_summary_button.pack(side=RIGHT, padx=self._pad(6, 0))

    # ---------- 合并后的上传入口 ----------

    def _upload_items(self) -> list[Path]:
        if self._input_allow_multi:
            return list(self.change_input_paths or [])
        text = self.input_path.get().strip()
        if text and not text.startswith("已选择 "):
            return [Path(text)]
        return []

    def _remove_upload_item(self, index: int) -> None:
        if self._input_allow_multi:
            paths = list(self.change_input_paths or [])
            if 0 <= index < len(paths):
                del paths[index]
            self.change_input_paths = paths or None
            self._sync_input_path_text()
        else:
            self.input_path.set("")
        self._refresh_upload_card()

    def _show_add_input_menu(self, _event=None) -> None:
        commands = []
        if self._input_file_cmd is not None:
            commands.append(("添加文件 / 压缩包", self._input_file_cmd))
        if self._input_folder_cmd is not None:
            commands.append(("添加文件夹", self._input_folder_cmd))
        if not commands:
            return
        if len(commands) == 1:
            commands[0][1]()
            return
        menu = Menu(
            self.root,
            tearoff=0,
            font=self.base_font,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            activebackground=COLOR_SURFACE_PRESSED,
            activeforeground=COLOR_TEXT,
        )
        for label, command in commands:
            menu.add_command(label=label, command=command)
        try:
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            menu.grab_release()

    @staticmethod
    def _format_file_size(path: Path) -> str:
        try:
            size = path.stat().st_size
        except Exception:
            return ""
        if size >= 1024 * 1024:
            return f"{size / (1024 * 1024):.1f} MB"
        return f"{max(1, size // 1024)} KB"

    def _upload_item_meta(self, path: Path) -> tuple[str, str, str, str]:
        """返回 (徽标文字, 徽标底色, 徽标前景, 说明文字)。"""
        if not path.exists():
            return "！", COLOR_BADGE_ZIP_BG, COLOR_BADGE_ZIP_FG, "文件不存在"
        if path.is_dir():
            try:
                child_count = sum(1 for _ in path.iterdir())
                detail = f"文件夹 · {child_count} 个项目"
            except Exception:
                detail = "文件夹"
            return "", COLOR_BADGE_DIR_BG, COLOR_BADGE_DIR_FG, detail
        suffix = path.suffix.lower()
        size_text = self._format_file_size(path)
        archive_type = archive_suffix(path)
        if archive_type:
            archive_badges = {
                ".zip": "ZIP",
                ".rar": "RAR",
                ".7z": "7Z",
                ".tar": "TAR",
                ".tar.gz": "TGZ",
                ".tgz": "TGZ",
                ".tar.bz2": "TBZ",
                ".tbz2": "TBZ",
                ".tar.xz": "TXZ",
                ".txz": "TXZ",
            }
            return archive_badges[archive_type], COLOR_BADGE_ZIP_BG, COLOR_BADGE_ZIP_FG, size_text
        if suffix in {".xlsx", ".xls"}:
            return "XLS", COLOR_BADGE_XLS_BG, COLOR_BADGE_XLS_FG, size_text
        return suffix.lstrip(".").upper()[:3] or "?", COLOR_BADGE_DIR_BG, COLOR_BADGE_DIR_FG, size_text

    def _refresh_upload_card(self) -> None:
        if not hasattr(self, "upload_body"):
            return
        for child in self.upload_body.winfo_children():
            child.destroy()
        items = self._upload_items()
        if items and self._input_allow_multi:
            self.upload_add_button.pack(side=RIGHT)
        else:
            self.upload_add_button.pack_forget()
        if items:
            self._render_upload_items(items)
        else:
            self._render_upload_drop_zone()
        if hasattr(self, "_apply_content_scroll_tag"):
            self._apply_content_scroll_tag(self.upload_body)
        if hasattr(self, "_sync_right_canvas_window"):
            self.root.after_idle(self._sync_right_canvas_window)

    def _render_upload_items(self, items: list[Path]) -> None:
        for index, path in enumerate(items):
            self._render_upload_chip(index, path, last=index == len(items) - 1)

    def _render_upload_chip(self, index: int, path: Path, *, last: bool) -> None:
        """圆角文件条目：类型徽标 + 文件名（超长省略，悬停显示完整路径）+ 说明 + ✕。"""
        badge_text, badge_bg, badge_fg, detail = self._upload_item_meta(path)
        chip = Canvas(self.upload_body, height=self._px(44), bg=COLOR_SURFACE, highlightthickness=0, bd=0)
        chip.pack(fill="x", pady=0 if last else self._pad(0, 8))
        name_font = (self.base_font[0], _font_size(10), "bold")
        last_chip_size = (0, 0)

        def redraw(_event=None) -> None:
            nonlocal last_chip_size
            width = max(chip.winfo_width(), 1)
            height = max(chip.winfo_height(), 1)
            if (width, height) == last_chip_size:
                return
            last_chip_size = (width, height)
            chip.delete("all")
            CodexButton._draw_round_rect(
                chip,
                self._pxf(0.5),
                self._pxf(0.5),
                width - self._pxf(0.5),
                height - self._pxf(0.5),
                self._pxf(10),
                fill=COLOR_SURFACE_ALT,
                outline=COLOR_BORDER_FAINT,
                width=max(1.0, self._pxf(1)),
            )
            badge_x = self._pxf(13)
            badge_size = self._pxf(26)
            badge_y = (height - badge_size) / 2
            CodexButton._draw_round_rect(chip, badge_x, badge_y, badge_x + badge_size, badge_y + badge_size, self._pxf(7), fill=badge_bg, outline="")
            if badge_text:
                chip.create_text(badge_x + badge_size / 2, height / 2, text=badge_text, fill=badge_fg, font=(self.base_font[0], _font_size(7), "bold"))
            else:
                _paint_tool_icon(chip, "folder_rename", badge_fg, badge_x + badge_size * 0.25, badge_y + badge_size * 0.25, badge_size * 0.5, max(1.0, self._pxf(1.3)))
            close_x = width - self._pxf(20)
            chip.create_text(close_x, height / 2, text="✕", fill="#C4C1B7", font=self.base_font, tags="chip_close")
            right_edge = close_x - self._pxf(16)
            if detail:
                meta_item = chip.create_text(right_edge, height / 2, text=detail, fill=COLOR_FAINT, font=self.small_font, anchor="e")
                meta_bbox = chip.bbox(meta_item)
                if meta_bbox:
                    right_edge = meta_bbox[0] - self._pxf(12)
            name_x = badge_x + badge_size + self._pxf(11)
            display_name = self._ellipsize(path.name or str(path), name_font, right_edge - name_x)
            chip.create_text(name_x, height / 2, text=display_name, fill=COLOR_TEXT, font=name_font, anchor="w")

        chip.bind("<Configure>", redraw)
        chip.tag_bind("chip_close", "<Button-1>", lambda _event, item_index=index: self._remove_upload_item(item_index))
        chip.tag_bind(
            "chip_close",
            "<Enter>",
            lambda _event: (chip.itemconfigure("chip_close", fill=COLOR_DANGER), chip.configure(cursor="hand2")),
        )
        chip.tag_bind(
            "chip_close",
            "<Leave>",
            lambda _event: (chip.itemconfigure("chip_close", fill="#C4C1B7"), chip.configure(cursor="")),
        )
        self._bind_path_tooltip(chip, lambda chip_path=path: str(chip_path))
        redraw()

    def _render_upload_drop_zone(self) -> None:
        zone = Canvas(self.upload_body, height=self._px(118), bg=COLOR_SURFACE, highlightthickness=0, bd=0)
        zone.pack(fill="x")
        last_zone_size = (0, 0)

        def redraw(_event=None) -> None:
            nonlocal last_zone_size
            width = max(zone.winfo_width(), 1)
            height = max(zone.winfo_height(), 1)
            if (width, height) == last_zone_size:
                return
            last_zone_size = (width, height)
            zone.delete("all")
            x1, y1 = self._pxf(1), self._pxf(1)
            x2, y2 = width - self._pxf(1), height - self._pxf(1)
            radius = self._pxf(12)
            self._draw_round_rect(zone, x1, y1, x2, y2, radius, fill=COLOR_SURFACE, outline="")
            # 虚线圆角边框：平滑多边形加 dash 会在角上留下墨点，改用直线 + 圆弧拼接
            dash = (5, 4)
            border = {"fill": COLOR_DROP_BORDER, "width": max(1.0, self._pxf(1.5)), "dash": dash}
            arc = {"outline": COLOR_DROP_BORDER, "width": max(1.0, self._pxf(1.5)), "dash": dash, "style": "arc"}
            zone.create_line(x1 + radius, y1, x2 - radius, y1, **border)
            zone.create_line(x2, y1 + radius, x2, y2 - radius, **border)
            zone.create_line(x2 - radius, y2, x1 + radius, y2, **border)
            zone.create_line(x1, y2 - radius, x1, y1 + radius, **border)
            zone.create_arc(x1, y1, x1 + 2 * radius, y1 + 2 * radius, start=90, extent=90, **arc)
            zone.create_arc(x2 - 2 * radius, y1, x2, y1 + 2 * radius, start=0, extent=90, **arc)
            zone.create_arc(x2 - 2 * radius, y2 - 2 * radius, x2, y2, start=270, extent=90, **arc)
            zone.create_arc(x1, y2 - 2 * radius, x1 + 2 * radius, y2, start=180, extent=90, **arc)
            center_x = width / 2
            icon_size = self._pxf(34)
            icon_top = self._pxf(16)
            self._draw_round_rect(
                zone,
                center_x - icon_size / 2,
                icon_top,
                center_x + icon_size / 2,
                icon_top + icon_size,
                self._pxf(10),
                fill=COLOR_PRIMARY_SOFT,
                outline="",
            )
            if self.current_tool in {"material_collector", "folder_rename"}:
                _paint_tool_icon(
                    zone,
                    "folder_rename",
                    COLOR_PRIMARY,
                    center_x - icon_size * 0.25,
                    icon_top + icon_size * 0.25,
                    icon_size * 0.5,
                    max(1.0, self._pxf(1.6)),
                )
            else:
                arrow = {"fill": COLOR_PRIMARY, "width": max(1.0, self._pxf(1.6)), "capstyle": "round", "joinstyle": "round"}
                icon_cx = center_x
                icon_cy = icon_top + icon_size / 2
                zone.create_line(icon_cx, icon_cy + self._pxf(5), icon_cx, icon_cy - self._pxf(7), **arrow)
                zone.create_line(icon_cx - self._pxf(5), icon_cy - self._pxf(2), icon_cx, icon_cy - self._pxf(7), icon_cx + self._pxf(5), icon_cy - self._pxf(2), **arrow)
                zone.create_line(icon_cx - self._pxf(7), icon_cy + self._pxf(9), icon_cx + self._pxf(7), icon_cy + self._pxf(9), **arrow)
            title_y = icon_top + icon_size + self._pxf(18)
            zone.create_text(
                center_x,
                title_y,
                text=self._input_drop_title,
                fill=COLOR_TEXT,
                font=(self.base_font[0], _font_size(10), "bold"),
            )
            links = []
            if self._input_file_cmd is not None:
                links.append(("浏览文件", self._input_file_cmd))
            if self._input_folder_cmd is not None:
                folder_label = "点击浏览文件夹路径" if self._input_file_cmd is None else "选择文件夹"
                links.append((folder_label, self._input_folder_cmd))
            link_y = title_y + self._pxf(21)
            segments: list[tuple[str, str, object | None]] = []
            if links and self._input_file_cmd is not None:
                segments.append(("或 ", COLOR_FAINT, None))
            for link_index, (label, command) in enumerate(links):
                if link_index > 0:
                    segments.append((" · ", COLOR_FAINT, None))
                segments.append((label, COLOR_PRIMARY, command))
            font_plain = self.small_font
            font_link = (self.small_font[0], self.small_font[1], "bold")
            total_width = 0.0
            measured = []
            for text, color, command in segments:
                font = font_link if command else font_plain
                segment_width = tkfont.Font(root=self.root, font=font).measure(text)
                measured.append((text, color, command, font, segment_width))
                total_width += segment_width
            cursor_x = center_x - total_width / 2
            for text, color, command, font, segment_width in measured:
                item = zone.create_text(cursor_x, link_y, text=text, fill=color, font=font, anchor="w")
                if command is not None:
                    zone.addtag_withtag("link", item)
                    zone.tag_bind(item, "<Button-1>", lambda _event, cmd=command: cmd())
                    zone.tag_bind(item, "<Enter>", lambda _event: zone.configure(cursor="hand2"))
                    zone.tag_bind(item, "<Leave>", lambda _event: zone.configure(cursor=""))
                cursor_x += segment_width

        def on_zone_click(_event=None):
            # 链接文字有自己的点击动作，避免和整块区域的默认动作重复触发
            current = zone.find_withtag("current")
            if current and "link" in zone.gettags(current[0]):
                return
            self._on_drop_zone_click()

        zone.bind("<Configure>", redraw)
        zone.bind("<Button-1>", on_zone_click)
        redraw()

    def _on_drop_zone_click(self) -> None:
        # 点击空态区域时弹出“文件/压缩包 或 文件夹”的选择菜单，
        # 避免首次上传只能进文件对话框、无法直接选文件夹
        self._show_add_input_menu()

    # ---------- 使用教程 ----------

    def _tutorial_entries(self) -> list[tuple[str, str | None, str]]:
        entries: list[tuple[str, str | None, str]] = []
        for tool_id, label in TOOL_NAV_ITEMS:
            if tool_id == "personnel_change_merge":
                entries.append((tool_id, "merge", "异动表汇总"))
                entries.append((tool_id, "roster", "花名册更新"))
            elif tool_id == "archive_import":
                entries.append((tool_id, "import", "档案入库"))
                entries.append((tool_id, "export", "档案表生成"))
            else:
                entries.append((tool_id, None, label))
        return entries

    def _current_tutorial_selection(self) -> tuple[str, str | None]:
        if self.current_tool == "personnel_change_merge":
            return self.current_tool, self.change_mode
        if self.current_tool == "archive_import":
            return self.current_tool, self.archive_mode
        return self.current_tool, None

    def _open_tutorial_window(self) -> None:
        if self._tutorial_window is not None and self._tutorial_window.winfo_exists():
            self._tutorial_window.lift()
            self._tutorial_window.focus_force()
            return
        window = Toplevel(self.root)
        self._tutorial_window = window
        window.title("使用教程")
        window.configure(bg=COLOR_BG)
        width, height = self._update_dialog_size(860, 620)
        self._center_window(window, width, height)
        window.minsize(self._px(640), self._px(420))

        body = ttk.Frame(window, padding=self._pad(16, 16, 20, 16), style="App.TFrame")
        body.pack(fill=BOTH, expand=True)

        nav = ttk.Frame(body, width=self._px(190), style="Sidebar.TFrame")
        nav.pack(side=LEFT, fill=Y)
        nav.pack_propagate(False)

        content_card = RoundedCard(body, padding=(22, 18, 22, 18), fill_height=True, min_height=320)
        content_card.pack(side=LEFT, fill=BOTH, expand=True, padx=self._pad(16, 0))
        content_title = ttk.Label(content_card.inner, text="", style="CardTitle.TLabel")
        content_title.pack(anchor="w", pady=self._pad(0, 10))
        text_wrap = ttk.Frame(content_card.inner, style="InputWrap.TFrame")
        text_wrap.pack(fill=BOTH, expand=True)
        text_scroll = ttk.Scrollbar(text_wrap, orient=VERTICAL)
        content_text = Text(
            text_wrap,
            wrap="word",
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=self.base_font,
            padx=self._px(2),
            pady=self._px(2),
            spacing3=self._px(9),
            yscrollcommand=text_scroll.set,
        )
        text_scroll.config(command=content_text.yview)
        content_text.pack(side=LEFT, fill=BOTH, expand=True)
        text_scroll.pack(side=RIGHT, fill=Y)
        content_text.tag_configure("strong", font=(self.base_font[0], _font_size(10), "bold"))
        content_text.tag_configure("warning", foreground=COLOR_WARNING, font=(self.base_font[0], _font_size(10), "bold"))

        nav_items: dict[tuple[str, str | None], SidebarItem] = {}

        def render(tool_id: str, mode: str | None) -> None:
            for entry_key, entry_item in nav_items.items():
                entry_item.set_selected(entry_key == (tool_id, mode))
            label = next(
                (entry_label for entry_tool, entry_mode, entry_label in self._tutorial_entries() if (entry_tool, entry_mode) == (tool_id, mode)),
                TOOL_NAV_LABELS.get(tool_id, ""),
            )
            content_title.configure(text=label)
            content_text.config(state="normal")
            content_text.delete("1.0", END)
            for line, tag in self._tutorial_lines(tool_id, mode):
                if tag:
                    content_text.insert(END, line + "\n", tag)
                else:
                    content_text.insert(END, line + "\n")
            content_text.config(state="disabled")

        previous_group: str | None = None
        for entry_tool, entry_mode, entry_label in self._tutorial_entries():
            group = TOOL_GROUP_LABELS.get(entry_tool, "")
            if group and group != previous_group:
                ttk.Label(nav, text=group, style="SidebarSection.TLabel").pack(anchor="w", padx=self._pad(9), pady=self._pad(10, 4))
                previous_group = group
            item = SidebarItem(
                nav,
                text=entry_label,
                icon_id=entry_tool,
                command=lambda tool=entry_tool, mode=entry_mode: render(tool, mode),
                height=30,
            )
            item.pack(fill="x", pady=self._px(1))
            nav_items[(entry_tool, entry_mode)] = item
            previous_tool = entry_tool

        selection = self._current_tutorial_selection()
        if selection not in nav_items:
            selection = next(iter(nav_items))
        render(*selection)

        def on_close() -> None:
            self._tutorial_window = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", on_close)
        window.transient(self.root)
        window.focus_force()

    def _tutorial_lines(self, tool_id: str, mode: str | None = None) -> list[tuple[str, str | None]]:
        if tool_id == "social_security":
            return [
                ("适用：把各社保账户缴费清单整理成社保明细表和社保汇总表。", "strong"),
                ("步骤：选择单个缴费清单、多个清单、常见压缩包，或包含清单的文件夹；再选择参保人员花名册。", None),
                ("结果：生成“社保明细表.xlsx”和“社保汇总表.xlsx”，汇总表里含基础数据分析和异常提醒。", None),
                ("目前规则：按身份证关联花名册；优先按账单文件夹或文件名识别账单月份、缴纳地和缴纳单位。", None),
                ("注意：公积金、残保金、管理费暂无数据时留空；账单识别结果与花名册不一致时会提醒。", "warning"),
            ]
        if tool_id == "data_statistics":
            return [
                ("适用：把 HR 系统导出的考勤结果、周报记录、月报记录自动整理成统计表。", "strong"),
                ("步骤：选择单个文件、多个文件、常见压缩包，或包含这些文件的文件夹。", None),
                ("如需统计未写周报/月报，请选择“应汇报人员名单”；不选时只能按文件中出现过的人推断。", None),
                ("周报统计日期（可选）：填写如 2026-06-02 至 2026-06-30，只统计范围内周一截止的周报；留空按整月统计。适合 1 号正好是周一的月份，避免把上月最后一周重复统计。", None),
                ("结果：生成“考勤周月报汇总表.xlsx”，包含考勤统计、周月报统计、考勤异常明细、周月报异常明细。", None),
                ("当前规则：考勤公司默认“总部”；周报截止次周一17:00，周二至周四补交算上一期超时（备注写明提交时间），周五起交的算下一期；月报按次月2日17:01及以后算超时。", None),
                ("容易疑惑1：如果某人上一期已经交过周报，周二到周四又交了一份，这份算他提前交的下一期，不记超时，下一期也不会记未写。", None),
                ("容易疑惑2：选了统计日期时，归属期超出范围的周报本次不统计、留给下一次。比如范围选到6.24，6.26（周五）交的属于6.29截止那期，本次不会出现。", None),
                ("注意：周月报异常只统计次数和明细，不计算扣款金额。", "warning"),
            ]
        if tool_id == "insurance_ledger":
            return [
                ("适用：把各保单人员清单整理成保险台账，并根据需求6的人力资源分析表做增减预警。", "strong"),
                ("步骤：选择单个保单清单、多个清单、常见压缩包，或包含清单的文件夹；再选择人力资源分析表。", None),
                ("结果：生成“保险台账.xlsx”，包含保险台账和人员增减预警两个工作表。", None),
                ("当前规则：PZDX保额取“每人伤残死亡限额”，按万元显示；PEAC保额固定按60万元。", None),
                ("注意：人力资源分析表需包含“花名册”工作表；花名册在职但保单没有会提示需加保，保单有但花名册没有或已标记离职会提示需减保。", "warning"),
            ]
        if tool_id == "salary_merge":
            return [
                ("适用：把 1-12 个月工资表合成一张个人应发工资汇总表。", "strong"),
                ("步骤：可选择单个月度工资表、多个工资表、常见压缩包，或包含这些文件的文件夹。", None),
                ("如已有前几月汇总表，再选择“已有汇总表”；不选则新建一张汇总表。", None),
                ("点击“开始合并”后，上传资料和结果会自动保存到当前工作项目。", None),
                ("结果：按姓名、身份证号、月份合并；没有工资的月份填 0；已存在的人员月份不会覆盖。", None),
                ("注意：工资表文件名或表内日期要能识别月份；重复人员或重复月份会在执行结果里提醒。", "warning"),
            ]
        if tool_id == "personnel_change_merge":
            if mode == "roster":
                return [
                    ("适用：已有月度异动汇总表时，单独更新人力资源花名册。", "strong"),
                    ("步骤：选择单个异动汇总表、多个汇总表，或包含汇总表的文件夹；再选择人力资源花名册。", None),
                    ("点击“更新花名册”后，上传资料和结果会自动保存到当前工作项目。", None),
                    ("结果：根据汇总表里的增员写入花名册，根据减员在花名册中标记离职。", None),
                    ("注意：不会清空原花名册；身份证已存在的增员不会重复写入，找不到的减员会在日志提醒。", "warning"),
                ]
            return [
                ("适用：把项目异动表按记录日期分到对应月份汇总表。", "strong"),
                ("步骤：可选择单个异动表、多个异动表、常见压缩包，或包含这些文件的文件夹。", None),
                ("如已有月度汇总表，可选择单个汇总表或包含多个汇总表的文件夹；工具会按月份追加，原有记录不会清空。", None),
                ("不选择已有汇总表时，工具会按月份新建干净汇总表。缺少某个月份汇总表时也会自动创建。", None),
                ("如果同一文件夹里放了人力资源分析表，工具会自动更新其中的花名册。", None),
                ("点击“开始汇总”后，上传资料和结果会自动保存到当前工作项目。", None),
                ("月份规则：增员看入职日期，减员看离职日期，转正看转正日期，调动看调整日期。", None),
                ("注意：只处理增补表、离职、转正、调整；薪酬、产值和同行对比分析暂不处理。", "warning"),
            ]
        if tool_id == "archive_import":
            if mode == "export":
                return [
                    ("适用：把一个或多个档案汇总表写入各公司独立档案表。", "strong"),
                    ("步骤：选择档案汇总表文件、多个文件、常见压缩包，或包含汇总表的文件夹。", None),
                    ("如已有某个公司的档案表，可选择文件、常见压缩包或文件夹；不选或没匹配到时会按内置干净模板新建。", None),
                    ("结果：按公司生成独立 Excel；已有身份证不重复新增，只补充空白字段。", None),
                    ("注意：公司档案表会自动改公司名，新增行会补边框、居中和公式。", "warning"),
                ]
            return [
                ("适用：把项目部提交的人事档案移交表写入公司档案汇总表。", "strong"),
                ("步骤：可选择单个移交表、多个移交表、常见压缩包，或包含这些文件的文件夹。", None),
                ("已有档案汇总表可不选；不选时工具会用内置空模板新建一份汇总表。", None),
                ("结果：按“公司”写入对应工作表；身份证已存在时不重复新增，只补充空白材料字段。", None),
                ("注意：编号会从文件名或表头标题识别项目地区，如“茂名项目部”自动填 11；识别不到会留空并提醒。", "warning"),
            ]
        if tool_id == "folder_rename":
            return [
                ("适用：批量修改所选目录下第一层文件夹或文件名称。", "strong"),
                ("按 Excel 人名顺序批量重命名：名单按姓名行顺序，项目按文件名顺序一一对应；文件保留原扩展名。", None),
                ("数量不一致、姓名无效、目标重名或目标已存在时会在预览中明确提醒；未配对或冲突项目不会改名，也不会覆盖。", "warning"),
                ("追加文字：姓名不填就是全部项目追加；填姓名就是只处理这个人。输入内容会原样追加，需要分隔符时请一并输入。", None),
                ("删除结尾文字：输入“_劳动合同”，可删除“张三_劳动合同 / 张三-劳动合同 / 张三劳动合同”的结尾文字。", None),
                ("修改单人名称：填写原姓名和新名称，例如“张三”改为“章五”。", None),
                ("安全说明：确认后会先把所选文件夹复制进当前项目，再在“处理结果”的副本上改名；电脑上的原文件夹不会被修改。", "warning"),
            ]
        if tool_id == "salary_split":
            return [
                ("适用：一个完整工资表按“入职公司”拆成多个公司工资表。", "strong"),
                ("步骤：先打开工作项目，选择工资表文件，再点击“开始拆分”。", None),
                ("点击“打开所在文件夹”可直接查看本次生成的结果目录。", None),
                ("结果：每个入职公司生成一个 Excel，保留表头、格式、公式、小计和底部总计。", None),
                ("注意：源工资表不会被修改；如果模板列名或表结构变化，先发给开发确认。", "warning"),
            ]
        if tool_id == "material_collector":
            return [
                ("适用：根据员工名单从资料库批量提取特定材料（身份证、合同、学历等）并自动打包。", "strong"),
                ("步骤：选择员工资料库根目录，在第二行选择员工名单表格（Excel），勾选需要的材料类型后点击“开始打包”。", None),
                ("资料库形式：已有姓名文件夹选“原模式”；文件无序混放时选“OCR 索引”，首次建立索引后会复用未变化文件。", None),
                ("归类方式：支持“按员工归类”（每人建一个文件夹）、“按材料归类”或“平铺输出”，可选自动生成 ZIP 压缩包。", None),
                ("结果：按指定结构导出文件，并自动生成《员工资料提取汇总与缺失清单.xlsx》。", None),
                ("安全说明：纯本地处理、不上传外网；源文件不会修改，OCR 索引模式只会新增隐藏缓存文件。", "warning"),
            ]
        return [
            ("该工具暂未实现。", "strong"),
            ("请选择左侧已完成的工具：需求1、需求2、需求4、需求5、需求6、需求7、需求8、需求9。", None),
        ]

    def _update_change_tabs_visibility(self) -> None:
        if self.current_tool in {"personnel_change_merge", "archive_import"}:
            if self.current_tool == "archive_import":
                self.change_tabs.tab(0, text="档案入库")
                self.change_tabs.tab(1, text="档案表生成")
                target_index = 1 if self.archive_mode == "export" else 0
            else:
                self.change_tabs.tab(0, text="异动表汇总")
                self.change_tabs.tab(1, text="花名册更新")
                target_index = 1 if self.change_mode == "roster" else 0
            if not self.change_tabs.winfo_ismapped():
                self.change_tabs.pack(fill="x", pady=self._pad(0, 16), before=self.upload_card)
            if self.change_tabs.index("current") != target_index:
                self.change_tabs.select(target_index)
            return
        self.change_tabs.pack_forget()

    def _update_summary_controls(self, apply_layout: bool = True) -> None:
        excel_rename = (
            self.current_tool == "folder_rename"
            and RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND) == MODE_EXCEL_BATCH
        )
        self._summary_row_visible = self.current_tool in {"social_security", "data_statistics", "insurance_ledger", "salary_merge", "personnel_change_merge", "archive_import", "material_collector"} or excel_rename
        if apply_layout and hasattr(self, "_apply_form_layout"):
            self._apply_form_layout()

    def _update_material_controls(self, apply_layout: bool = True) -> None:
        self._material_row_visible = self.current_tool == "material_collector"
        if apply_layout and hasattr(self, "_apply_form_layout"):
            self._apply_form_layout()

    def _apply_tool_specific_responsive_layout(self) -> None:
        mode = getattr(self, "_form_layout_mode", LAYOUT_MODE_WIDE)
        stacked = mode != LAYOUT_MODE_WIDE
        narrow = mode == LAYOUT_MODE_NARROW

        if hasattr(self, "material_library_mode_row"):
            for widget in (
                self.material_library_mode_label,
                self.material_library_mode_combo,
                self.material_library_mode_hint,
            ):
                widget.pack_forget()
            self.material_library_mode_label.configure(width=0 if stacked else 8)
            self.material_library_mode_combo.configure(width=1 if stacked else 28)
            if stacked:
                self.material_library_mode_label.pack(side=TOP, anchor="w")
                self.material_library_mode_combo.pack(side=TOP, fill="x", pady=self._pad(5, 0))
                self.material_library_mode_hint.pack(side=TOP, fill="x", pady=self._pad(4, 0))
            else:
                self.material_library_mode_label.pack(side=LEFT, anchor="center")
                self.material_library_mode_combo.pack(side=LEFT, padx=self._pad(8, 12))
                self.material_library_mode_hint.pack(side=LEFT, fill="x", expand=True)

            for widget in (self.material_target_label, self.material_target_wrap):
                widget.pack_forget()
            self.material_target_label.configure(width=0 if stacked else 8)
            if stacked:
                self.material_target_label.pack(side=TOP, anchor="w")
                self.material_target_wrap.pack(side=TOP, fill="x", pady=self._pad(5, 0))
            else:
                self.material_target_label.pack(side=LEFT, anchor="center")
                self.material_target_wrap.pack(side=LEFT, fill="x", expand=True, padx=self._pad(8, 0))

            for widget in (self.material_opts_label, self.material_opts_checks_frame):
                widget.pack_forget()
            self.material_opts_label.configure(width=0 if stacked else 8)
            if stacked:
                self.material_opts_label.pack(side=TOP, anchor="w")
                self.material_opts_checks_frame.pack(side=TOP, fill="x", pady=self._pad(4, 0))
            else:
                self.material_opts_label.pack(side=LEFT, anchor="center")
                self.material_opts_checks_frame.pack(side=LEFT, fill="x", expand=True, padx=self._pad(8, 0))

            for widget in (
                self.material_collect_all_check,
                self.material_zip_check,
                self.material_use_ocr_cache_check,
            ):
                widget.pack_forget()
            if stacked:
                self.material_collect_all_check.pack(side=TOP, anchor="w")
                self.material_zip_check.pack(side=TOP, anchor="w", pady=self._pad(4, 0))
                self.material_use_ocr_cache_check.pack(side=TOP, anchor="w", pady=self._pad(4, 0))
            else:
                self.material_collect_all_check.pack(side=LEFT, padx=self._pad(0, 16))
                self.material_zip_check.pack(side=LEFT, padx=self._pad(0, 16))
                self.material_use_ocr_cache_check.pack(side=RIGHT, padx=self._pad(0, 4))

            hint_pad = self._pad(0, 0) if stacked else self._pad(76, 0)
            self.material_input_hint.pack_configure(padx=hint_pad)
            self.material_types_hint.pack_configure(padx=hint_pad)

            self.material_header_row.pack_forget()
            self.mat_checks_frame.pack_forget()
            if stacked:
                self.material_header_row.pack(side=TOP, fill="x", anchor="w", padx=0)
                self.mat_checks_frame.pack(side=TOP, fill="x", expand=True, pady=self._pad(6, 0))
            else:
                self.material_header_row.pack(side=LEFT, anchor="nw", padx=self._pad(0, 8))
                self.mat_checks_frame.pack(side=LEFT, fill="x", expand=True)

            checkbox_columns = _responsive_checkbox_columns(mode)
            if checkbox_columns != self._material_checkbox_columns:
                self._material_checkbox_columns = checkbox_columns
                self._rebuild_material_checkboxes()

        if hasattr(self, "stats_week_inputs"):
            for date_group, hint in (
                (self.stats_week_date_group, self.stats_week_hint),
                (self.stats_month_date_group, self.stats_month_hint),
            ):
                date_group.pack_forget()
                hint.pack_forget()
                if stacked:
                    date_group.pack(side=TOP, anchor="w")
                    hint.pack(side=TOP, fill="x", anchor="w", pady=self._pad(4, 0))
                else:
                    date_group.pack(side=LEFT)
                    hint.pack(side=LEFT, padx=self._pad(10, 0))

            for column in (self.stats_unit_col, self.stats_out_col, self.stats_trip_col):
                column.pack_forget()
            if stacked:
                self.stats_unit_col.pack(side=TOP, fill="x")
                self.stats_out_col.pack(side=TOP, fill="x", pady=self._pad(7, 0))
                self.stats_trip_col.pack(side=TOP, fill="x", pady=self._pad(7, 0))
            else:
                self.stats_unit_col.pack(side=LEFT, fill="x", expand=True)
                self.stats_out_col.pack(side=LEFT, fill="x", expand=True)
                self.stats_trip_col.pack(side=LEFT, fill="x", expand=True)

            for frame, buttons in (
                (self.stats_week_presets, self.stats_week_preset_buttons),
                (self.stats_month_presets, self.stats_month_preset_buttons),
            ):
                for button in buttons:
                    button.pack_forget()
                    try:
                        button.grid_forget()
                    except Exception:
                        pass
                if narrow:
                    for column in range(2):
                        frame.columnconfigure(column, weight=1)
                    for index, button in enumerate(buttons):
                        button.grid(
                            row=index // 2,
                            column=index % 2,
                            sticky="ew",
                            padx=self._pad(0, 8),
                            pady=self._pad(0, 6),
                        )
                else:
                    for column in range(2):
                        frame.columnconfigure(column, weight=0)
                    for button in buttons:
                        button.pack(side=LEFT, padx=self._pad(0, 8))

    def _selected_material_names(self) -> list[str]:
        return [
            material
            for material, variable in self.material_types_selected.items()
            if variable.get()
        ]

    def _rebuild_material_checkboxes(self) -> None:
        grid = getattr(self, "material_checks_grid", None)
        if grid is None:
            return
        for widget in getattr(self, "_material_check_widgets", []):
            try:
                widget.destroy()
            except Exception:
                pass
        self._material_check_widgets = []
        columns = max(1, int(getattr(self, "_material_checkbox_columns", 4)))
        for column in range(4):
            grid.columnconfigure(column, weight=1 if column < columns else 0)
        for idx, (material, variable) in enumerate(self.material_types_selected.items()):
            row = idx // columns
            column = idx % columns
            checkbox = ttk.Checkbutton(
                grid,
                text=material,
                variable=variable,
                style="App.TCheckbutton",
            )
            checkbox.grid(
                row=row,
                column=column,
                sticky="w",
                padx=self._px(8),
                pady=self._px(4),
            )
            self._material_check_widgets.append(checkbox)

    def _refresh_material_preset_combo(self, preferred: str | None = None) -> None:
        names = self._material_preferences.preset_names
        combo = getattr(self, "material_preset_combo", None)
        if combo is not None:
            combo.configure(values=names)
        current = preferred or self.material_preset_name.get().strip()
        if current in names:
            self.material_preset_name.set(current)
        else:
            self.material_preset_name.set(names[0] if names else "")

    def _refresh_material_catalog_combo(self, preferred: str | None = None) -> None:
        names = self._material_preferences.custom_materials
        combo = getattr(self, "material_custom_combo", None)
        if combo is not None:
            combo.configure(values=names)
        current = preferred or self.material_custom_choice.get().strip()
        if current in names:
            self.material_custom_choice.set(current)
        else:
            self.material_custom_choice.set(names[0] if names else "")

    def _select_all_material_types(self) -> None:
        for var in self.material_types_selected.values():
            var.set(True)

    def _deselect_all_material_types(self) -> None:
        for var in self.material_types_selected.values():
            var.set(False)

    def _request_add_custom_material(self) -> None:
        raw_name = simpledialog.askstring(
            "添加自定义材料",
            "输入材料名称（例如：户口本、体检报告）",
            parent=self.root,
        )
        if raw_name is None:
            return
        try:
            material = self._material_preferences.add_material(raw_name)
        except ValueError as exc:
            messagebox.showwarning("无法添加材料", str(exc), parent=self.root)
            return
        self.material_types_selected[material] = BooleanVar(value=True)
        self._rebuild_material_checkboxes()
        self._refresh_material_catalog_combo(material)
        self._save_workspace_preferences()

    def _request_delete_custom_material(self) -> None:
        custom_materials = self._material_preferences.custom_materials
        if not custom_materials:
            messagebox.showinfo(
                "没有自定义材料",
                "内置材料不能删除；当前还没有可删除的自定义材料。",
                parent=self.root,
            )
            return
        chosen = self.material_custom_choice.get().strip()
        material = next((item for item in custom_materials if item == chosen), None)
        if material is None:
            messagebox.showwarning(
                "请选择材料",
                "请先从“自定义材料”列表中选择要删除的材料。",
                parent=self.root,
            )
            return

        referenced_by = [
            preset_name
            for preset_name, materials in self._material_preferences.custom_presets.items()
            if material in materials
        ]
        reference_text = ""
        if referenced_by:
            reference_text = (
                "\n\n该材料还被以下预设引用："
                + "、".join(referenced_by)
                + "。删除后会自动清理这些引用；没有剩余材料的预设也会一并删除。"
            )
        if not messagebox.askyesno(
            "确认删除材料",
            f"确定删除自定义材料“{material}”吗？{reference_text}",
            parent=self.root,
        ):
            return

        result = self._material_preferences.remove_material(material)
        self.material_types_selected.pop(material, None)
        self._rebuild_material_checkboxes()
        self._refresh_material_catalog_combo()
        self._refresh_material_preset_combo()
        self._save_workspace_preferences()

        changed: list[str] = []
        if result.updated_presets:
            changed.append("已更新预设：" + "、".join(result.updated_presets))
        if result.removed_presets:
            changed.append("已删除空预设：" + "、".join(result.removed_presets))
        message = f"已删除自定义材料“{material}”。"
        if changed:
            message += "\n" + "\n".join(changed)
        messagebox.showinfo("材料已删除", message, parent=self.root)

    def _apply_material_preset(self, preset_name: str | None = None) -> None:
        name = (preset_name or self.material_preset_name.get()).strip()
        materials = self._material_preferences.get_preset(name)
        if materials is None:
            messagebox.showwarning(
                "预设不可用",
                "这个预设不存在或已经被删除，请重新选择。",
                parent=self.root,
            )
            self._refresh_material_preset_combo()
            return
        selected = set(materials)
        for material, variable in self.material_types_selected.items():
            variable.set(material in selected)
        self.material_preset_name.set(name)
        self.material_collect_all.set(False)
        self._on_material_collect_all_changed()

    def _save_current_material_preset(
        self,
        name: str,
        *,
        replacing: str | None = None,
    ) -> str:
        saved_name = self._material_preferences.save_preset(
            name,
            self._selected_material_names(),
            replacing=replacing,
        )
        self._refresh_material_preset_combo(saved_name)
        self._save_workspace_preferences()
        return saved_name

    def _request_create_material_preset(self) -> None:
        if not self._selected_material_names():
            messagebox.showwarning(
                "没有选择材料",
                "请先勾选至少一种材料，再保存为预设。",
                parent=self.root,
            )
            return
        raw_name = simpledialog.askstring(
            "保存自定义预设",
            "输入预设名称",
            parent=self.root,
        )
        if raw_name is None:
            return
        try:
            saved_name = self._save_current_material_preset(raw_name)
        except ValueError as exc:
            messagebox.showwarning("无法保存预设", str(exc), parent=self.root)
            return
        messagebox.showinfo(
            "预设已保存",
            f"已保存“{saved_name}”，以后可从常用组合中直接应用。",
            parent=self.root,
        )

    def _request_update_material_preset(self) -> None:
        name = self.material_preset_name.get().strip()
        if self._material_preferences.is_builtin_preset(name):
            messagebox.showinfo(
                "内置预设不能修改",
                "如需调整，请先勾选需要的材料，再保存为新的自定义预设。",
                parent=self.root,
            )
            return
        try:
            self._save_current_material_preset(name, replacing=name)
        except ValueError as exc:
            messagebox.showwarning("无法更新预设", str(exc), parent=self.root)
            return
        messagebox.showinfo("预设已更新", f"“{name}”已按当前勾选更新。", parent=self.root)

    def _request_rename_material_preset(self) -> None:
        current_name = self.material_preset_name.get().strip()
        if self._material_preferences.is_builtin_preset(current_name):
            messagebox.showinfo(
                "内置预设不能重命名",
                "内置预设会一直保留；自定义预设可以重命名。",
                parent=self.root,
            )
            return
        raw_name = simpledialog.askstring(
            "重命名预设",
            "输入新的预设名称",
            initialvalue=current_name,
            parent=self.root,
        )
        if raw_name is None:
            return
        try:
            saved_name = self._material_preferences.rename_preset(current_name, raw_name)
        except ValueError as exc:
            messagebox.showwarning("无法重命名预设", str(exc), parent=self.root)
            return
        self._refresh_material_preset_combo(saved_name)
        self._save_workspace_preferences()

    def _request_delete_material_preset(self) -> None:
        name = self.material_preset_name.get().strip()
        if self._material_preferences.is_builtin_preset(name):
            messagebox.showinfo(
                "内置预设不能删除",
                "内置预设会一直保留；自定义预设可以删除。",
                parent=self.root,
            )
            return
        if not messagebox.askyesno(
            "确认删除预设",
            f"确定删除自定义预设“{name}”吗？\n\n材料本身不会被删除。",
            parent=self.root,
        ):
            return
        try:
            self._material_preferences.delete_preset(name)
        except ValueError as exc:
            messagebox.showwarning("无法删除预设", str(exc), parent=self.root)
            return
        self._refresh_material_preset_combo()
        self._save_workspace_preferences()

    def _on_material_library_mode_changed(self, _event=None) -> None:
        """切换资料库组织形式，不改变现有输出归类模式。"""
        if not hasattr(self, "material_library_mode"):
            return
        selected = LIBRARY_MODE_LABELS.get(
            self.material_library_mode.get(),
            LIBRARY_MODE_PERSON_FOLDER,
        )
        is_flat_ocr = selected == LIBRARY_MODE_FLAT_OCR
        previous_mode = getattr(
            self,
            "_last_material_library_mode",
            LIBRARY_MODE_PERSON_FOLDER,
        )
        if is_flat_ocr:
            if previous_mode != LIBRARY_MODE_FLAT_OCR:
                self._material_person_mode_cache_preference = bool(
                    self.material_use_ocr_cache.get()
                )
            self.material_use_ocr_cache.set(True)
            if hasattr(self, "material_use_ocr_cache_check"):
                self.material_use_ocr_cache_check.configure(state="disabled")
            if hasattr(self, "material_collect_all_check"):
                self.material_collect_all_check.configure(
                    text="全部（提取 OCR 识别到的该人员全部材料）"
                )
            if hasattr(self, "material_library_mode_hint"):
                self.material_library_mode_hint.configure(
                    text="源文件不改；首次建立隐藏索引，未变化文件直接复用"
                )
            if hasattr(self, "material_types_hint"):
                self.material_types_hint.configure(
                    text="取消勾选「全部」后，可只提取指定材料；索引仍会覆盖整个资料库"
                )
        else:
            if previous_mode == LIBRARY_MODE_FLAT_OCR and hasattr(
                self, "_material_person_mode_cache_preference"
            ):
                self.material_use_ocr_cache.set(
                    self._material_person_mode_cache_preference
                )
            if hasattr(self, "material_use_ocr_cache_check"):
                self.material_use_ocr_cache_check.configure(state="normal")
            if hasattr(self, "material_collect_all_check"):
                self.material_collect_all_check.configure(
                    text="全部（直接拷贝匹配到的人员整个文件夹）"
                )
            if hasattr(self, "material_library_mode_hint"):
                self.material_library_mode_hint.configure(text="原模式按姓名文件夹查找")
            if hasattr(self, "material_types_hint"):
                self.material_types_hint.configure(
                    text="取消勾选「全部」后可按需勾选材料类型（如身份证、劳动合同等）"
                )
        self._last_material_library_mode = selected

    def _on_material_collect_all_changed(self) -> None:
        """全部模式切换：勾选时隐藏材料类型选择，取消时显示。"""
        if not hasattr(self, "material_types_section"):
            return
        if self.material_collect_all.get():
            self.material_types_section.pack_forget()
            self.material_types_hint.pack(fill="x", padx=(self._px(76), 0), pady=(0, self._px(6)))
        else:
            self.material_types_hint.pack_forget()
            self.material_types_section.pack(fill="x", pady=(self._px(6), self._px(2)))

    def _update_change_picker_buttons(self) -> None:
        """配置合并后的上传入口动作，以及第二行（花名册/汇总表）的选择链接。"""

        def hide(*buttons) -> None:
            for button in buttons:
                self._hide_picker_button(button)

        def show(*buttons) -> None:
            for button in buttons:
                self._show_picker_button(button)

        tool = self.current_tool
        self._input_allow_multi = tool in MULTI_INPUT_TOOLS

        if tool == "social_security":
            self._input_file_cmd = self._choose_social_security_files_or_zip
            self._input_folder_cmd = self._choose_social_security_folder
        elif tool == "data_statistics":
            self._input_file_cmd = self._choose_data_statistics_files_or_zip
            self._input_folder_cmd = self._choose_data_statistics_folder
        elif tool == "insurance_ledger":
            self._input_file_cmd = self._choose_insurance_files_or_zip
            self._input_folder_cmd = self._choose_insurance_folder
        elif tool == "salary_merge":
            self._input_file_cmd = self._choose_salary_files_or_zip
            self._input_folder_cmd = self._choose_salary_folder
        elif tool == "personnel_change_merge":
            if self.change_mode == "roster":
                self._input_file_cmd = self._choose_roster_summary_files
                self._input_folder_cmd = self._choose_roster_summary_folder
            else:
                self._input_file_cmd = self._choose_change_files_or_zip
                self._input_folder_cmd = self._choose_change_folder
        elif tool == "archive_import":
            if self.archive_mode == "export":
                self._input_file_cmd = self._choose_archive_export_summary_files_or_zip
                self._input_folder_cmd = self._choose_archive_export_summary_folder
            else:
                self._input_file_cmd = self._choose_archive_files_or_zip
                self._input_folder_cmd = self._choose_archive_folder
        elif tool in {"folder_rename", "material_collector"}:
            self._input_file_cmd = None
            self._input_folder_cmd = self._choose_input
        else:  # salary_split 及未实现工具：单个文件
            self._input_file_cmd = self._choose_input
            self._input_folder_cmd = None

        if tool == "social_security":
            hide(self.summary_choose_button, self.change_summary_folder_button)
            self.change_summary_file_button.configure(text="选择文件", command=self._choose_social_security_roster_file)
            show(self.change_summary_file_button)
        elif tool == "data_statistics":
            hide(self.summary_choose_button, self.change_summary_folder_button)
            self.change_summary_file_button.configure(text="选择文件", command=self._choose_data_statistics_staff_file)
            show(self.change_summary_file_button)
        elif tool == "insurance_ledger":
            hide(self.summary_choose_button, self.change_summary_folder_button)
            self.change_summary_file_button.configure(text="选择文件", command=self._choose_insurance_roster_file)
            show(self.change_summary_file_button)
        elif tool == "salary_merge":
            hide(self.change_summary_folder_button, self.change_summary_file_button)
            self.summary_choose_button.configure(text="选择文件")
            show(self.summary_choose_button)
        elif tool == "personnel_change_merge":
            hide(self.summary_choose_button)
            if self.change_mode == "roster":
                hide(self.change_summary_folder_button)
                self.change_summary_file_button.configure(text="选择文件", command=self._choose_roster_analysis_file)
                show(self.change_summary_file_button)
            else:
                self.change_summary_folder_button.configure(text="选择文件夹", command=self._choose_change_summary_folder)
                self.change_summary_file_button.configure(text="选择文件", command=self._choose_change_summary_file)
                show(self.change_summary_file_button, self.change_summary_folder_button)
        elif tool == "archive_import":
            hide(self.summary_choose_button)
            if self.archive_mode == "export":
                self.change_summary_folder_button.configure(text="选择文件夹", command=self._choose_archive_export_existing_folder)
                self.change_summary_file_button.configure(text="选择文件", command=self._choose_archive_export_existing_file_or_zip)
                show(self.change_summary_file_button, self.change_summary_folder_button)
            else:
                hide(self.change_summary_folder_button)
                self.change_summary_file_button.configure(text="选择文件", command=self._choose_archive_summary_file)
                show(self.change_summary_file_button)
        elif tool == "folder_rename":
            hide(self.summary_choose_button, self.change_summary_folder_button, self.change_summary_file_button)
            if RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND) == MODE_EXCEL_BATCH:
                self.change_summary_file_button.configure(text="选择名单", command=self._choose_folder_rename_excel_file)
                show(self.change_summary_file_button)
        elif tool == "material_collector":
            hide(self.summary_choose_button, self.change_summary_folder_button)
            self.change_summary_file_button.configure(text="选择文件", command=self._choose_material_roster_file)
            show(self.change_summary_file_button)
        else:
            hide(self.change_summary_folder_button, self.change_summary_file_button)
            show(self.summary_choose_button)

        self._refresh_upload_card()

    def _update_output_controls(self, apply_layout: bool = True) -> None:
        self._output_row_visible = self.current_tool != "folder_rename"
        if apply_layout and hasattr(self, "_apply_form_layout"):
            self._apply_form_layout()

    def _update_rename_controls(self, apply_layout: bool = True) -> None:
        self._rename_row_visible = self.current_tool == "folder_rename"
        if apply_layout and hasattr(self, "_apply_form_layout"):
            self._apply_form_layout()
        if self._rename_row_visible:
            self._update_rename_mode_controls()

    def _update_stats_range_controls(self, apply_layout: bool = True) -> None:
        self._stats_range_row_visible = self.current_tool == "data_statistics"
        if apply_layout and hasattr(self, "_apply_form_layout"):
            self._apply_form_layout()

    def _fill_stats_week_range(self, preset: str) -> None:
        if preset == "clear":
            self.stats_week_start.set("")
            self.stats_week_end.set("")
            return
        today = date.today()
        monday = today - timedelta(days=today.weekday())
        if preset == "this_month":
            start = today.replace(day=1)
            end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
        elif preset == "last_month":
            end = today.replace(day=1) - timedelta(days=1)
            start = end.replace(day=1)
        elif preset == "this_week":
            start = monday
            end = monday + timedelta(days=6)
        else:  # last_week
            start = monday - timedelta(days=7)
            end = monday - timedelta(days=1)
        self.stats_week_start.set(start.isoformat())
        self.stats_week_end.set(end.isoformat())

    def _fill_stats_month_range(self, preset: str) -> None:
        if preset == "clear":
            self.stats_month_start.set("")
            self.stats_month_end.set("")
            return
        today = date.today()
        if preset == "this_month":
            start = today.replace(day=1)
            end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
        elif preset == "last_month":
            end = today.replace(day=1) - timedelta(days=1)
            start = end.replace(day=1)
        else:
            return
        self.stats_month_start.set(start.isoformat())
        self.stats_month_end.set(end.isoformat())

    def _show_tooltip(self, widget, text: str) -> None:
        """显示一个轻量的气泡提示，用于解释控件语义。"""
        self._hide_tooltip()
        try:
            x = widget.winfo_rootx() + widget.winfo_width() + 6
            y = widget.winfo_rooty() - 2
        except Exception:
            x = y = 0
        tip = Toplevel(widget)
        tip.wm_overrideredirect(True)
        tip.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(
            tip,
            text=text,
            style="App.TLabel",
            background="#FFF8E1",
            foreground="#5C4400",
            borderwidth=1,
            relief="solid",
            padding=(8, 6),
            wraplength=280,
            justify="left",
        )
        label.pack()
        self._tooltip_window = tip

    def _hide_tooltip(self) -> None:
        tip = getattr(self, "_tooltip_window", None)
        if tip is not None:
            try:
                tip.destroy()
            except Exception:
                pass
            self._tooltip_window = None

    def _on_rename_mode_changed(self, _event=None) -> None:
        self._set_tool_texts()
        self._clear_log()
        self._write_log(self._initial_log_text())

    def _update_rename_mode_controls(self) -> None:
        mode = RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND)
        # 文件类型选择器始终显示
        self.rename_file_type_label_widget.grid(row=4, column=0, sticky="w", pady=self._px(5))
        self.rename_file_type_widget.grid(row=4, column=1, sticky="w", padx=self._px(12), pady=self._px(5))
        if mode == MODE_EXCEL_BATCH:
            self.rename_target_label_widget.grid_remove()
            self.rename_target_widget.grid_remove()
            self.rename_text_label_widget.grid_remove()
            self.rename_text_widget.grid_remove()
            self.rename_replacement_label_widget.grid_remove()
            self.rename_replacement_widget.grid_remove()
            return

        self.rename_target_label_widget.grid(row=1, column=0, sticky="w", pady=self._px(5))
        self.rename_target_widget.grid(row=1, column=1, sticky="ew", padx=self._px(12), pady=self._px(5))
        if mode == MODE_APPEND:
            self.rename_target_label.set("姓名（可不填）")
            self.rename_text_label.set("要追加的文字")
            self.rename_text_label_widget.grid(row=2, column=0, sticky="w", pady=self._px(5))
            self.rename_text_widget.grid(row=2, column=1, sticky="ew", padx=self._px(12), pady=self._px(5))
            self.rename_text_widget.config(state="normal")
            self.rename_replacement_label_widget.grid_remove()
            self.rename_replacement_widget.grid_remove()
        elif mode == MODE_REMOVE:
            self.rename_target_label.set("姓名（可不填）")
            self.rename_text_label.set("要删除的结尾文字")
            self.rename_text_label_widget.grid(row=2, column=0, sticky="w", pady=self._px(5))
            self.rename_text_widget.grid(row=2, column=1, sticky="ew", padx=self._px(12), pady=self._px(5))
            self.rename_text_widget.config(state="normal")
            self.rename_replacement_label_widget.grid_remove()
            self.rename_replacement_widget.grid_remove()
        else:
            self.rename_target_label.set("原名称")
            self.rename_replacement_label.set("新名称")
            self.rename_text_label_widget.grid_remove()
            self.rename_text_widget.grid_remove()
            self.rename_replacement_label_widget.grid(row=2, column=0, sticky="w", pady=self._px(5))
            self.rename_replacement_widget.grid(row=2, column=1, sticky="ew", padx=self._px(12), pady=self._px(5))
            self.rename_replacement_widget.config(state="normal")

    def _initial_log_text(self) -> str:
        if self.current_tool == "social_security":
            return "请选择社保缴费清单和参保人员花名册，然后点击“生成报表”。资料和结果会自动留存在当前项目。"
        if self.current_tool == "data_statistics":
            return "请选择考勤结果、周报记录和月报记录，然后点击“生成统计”。应汇报人员名单是可选项，资料和结果会自动留存在当前项目。"
        if self.current_tool == "insurance_ledger":
            return "请选择保单人员清单和人力资源分析表，然后点击“生成台账”。资料和结果会自动留存在当前项目。"
        if self.current_tool == "salary_merge":
            return "请选择工资表文件、压缩包或文件夹，然后点击“开始合并”。已有汇总表是可选项，资料和结果会自动留存在当前项目。"
        if self.current_tool == "personnel_change_merge":
            if self.change_mode == "roster":
                return "请选择异动汇总表和人力资源花名册，然后点击“更新花名册”。资料和结果会自动留存在当前项目。"
            return "请选择异动表文件或文件夹，然后点击“开始汇总”。已有汇总表是可选项，资料和结果会自动留存在当前项目。"
        if self.current_tool == "archive_import":
            if self.archive_mode == "export":
                return "请选择档案汇总表、压缩包或文件夹，然后点击“生成档案表”。已有公司档案表是可选项，结果会自动留存在当前项目。"
            return "请选择移交表文件、压缩包或文件夹，然后点击“开始入库”。已有档案汇总表是可选项，结果会自动留存在当前项目。"
        if self.current_tool == "folder_rename":
            if RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND) == MODE_EXCEL_BATCH:
                return "请选择人员资料目录和人员名单 Excel，选择文件类型后点击“预览”。名单和原目录都不会在预览时改变。"
            return "请选择人员文件夹目录，填写改名内容，然后点击“预览”。"
        if self.current_tool == "salary_split":
            return "请选择工资表文件，然后点击“开始拆分”。资料和结果会自动留存在当前项目。"
        if self.current_tool == "material_collector":
            return "请选择员工资料库根目录和员工名单 Excel 文件，勾选所需材料类型，然后点击“开始打包”。"
        return "该工具暂未实现。"

    def _choose_input(self) -> None:
        if self.current_tool in {"salary_merge", "personnel_change_merge", "folder_rename", "archive_import", "material_collector"}:
            if self.current_tool == "personnel_change_merge":
                if self.change_mode == "roster":
                    self._choose_roster_summary_files()
                else:
                    self._choose_change_files_or_zip()
                return
            elif self.current_tool == "archive_import":
                title = "选择档案移交表文件夹"
            elif self.current_tool == "folder_rename":
                title = "选择人员文件夹目录"
            elif self.current_tool == "material_collector":
                title = "选择员工资料库根目录"
            else:
                title = "选择工资表文件夹"
            directory = self._askdirectory(title=title)
            if directory:
                self.input_path.set(directory)
                if not self.output_dir_user_selected:
                    self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
                self._refresh_upload_card()
            return

        filename = self._askopenfilename(
            title="选择工资表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.input_path.set(filename)
            if not self.output_dir_user_selected:
                self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
            self._refresh_upload_card()

    def _choose_change_folder(self) -> None:
        directory = self._askdirectory(title="选择异动表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_change_files_or_zip(self) -> None:
        filenames = self._askopenfilenames(
            title="选择异动表文件或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_salary_folder(self) -> None:
        directory = self._askdirectory(title="选择工资表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_salary_files_or_zip(self) -> None:
        filenames = self._askopenfilenames(
            title="选择工资表文件或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_social_security_folder(self) -> None:
        directory = self._askdirectory(title="选择社保缴费清单文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_social_security_files_or_zip(self) -> None:
        filenames = self._askopenfilenames(
            title="选择社保缴费清单或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_social_security_roster_file(self) -> None:
        filename = self._askopenfilename(
            title="选择参保人员花名册",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_material_roster_file(self) -> None:
        filename = self._askopenfilename(
            title="选择员工名单 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_folder_rename_excel_file(self) -> None:
        filename = self._askopenfilename(
            title="选择人员名单 Excel",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_data_statistics_folder(self) -> None:
        directory = self._askdirectory(title="选择考勤周月报数据文件夹")
        if directory:
            # 替换式：每次重新选择都覆盖旧路径，避免 chip 残留导致重复上传时统计报错
            self.change_input_paths = [Path(directory)]
            self._sync_input_path_text()
            self._refresh_upload_card()
            if not self.output_dir_user_selected:
                self.output_dir.set(str(default_output_parent_dir(self.current_tool)))

    def _choose_data_statistics_files_or_zip(self) -> None:
        filenames = self._askopenfilenames(
            title="选择考勤周月报文件或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filenames:
            # 替换式：每次重新选择都覆盖旧路径，避免 chip 残留导致重复上传时统计报错
            self.change_input_paths = [Path(filename) for filename in filenames]
            self._sync_input_path_text()
            self._refresh_upload_card()
            if not self.output_dir_user_selected:
                self.output_dir.set(str(default_output_parent_dir(self.current_tool)))

    def _choose_data_statistics_staff_file(self) -> None:
        filename = self._askopenfilename(
            title="选择应汇报人员名单",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_insurance_folder(self) -> None:
        directory = self._askdirectory(title="选择保单人员清单文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_insurance_files_or_zip(self) -> None:
        filenames = self._askopenfilenames(
            title="选择保单人员清单或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_insurance_roster_file(self) -> None:
        filename = self._askopenfilename(
            title="选择人力资源分析表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_archive_folder(self) -> None:
        directory = self._askdirectory(title="选择档案移交表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_archive_files_or_zip(self) -> None:
        filenames = self._askopenfilenames(
            title="选择档案移交表文件或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_archive_summary_file(self) -> None:
        filename = self._askopenfilename(
            title="选择已有档案汇总表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_archive_export_summary_files_or_zip(self) -> None:
        filenames = self._askopenfilenames(
            title="选择档案汇总表或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_archive_export_summary_folder(self) -> None:
        directory = self._askdirectory(title="选择档案汇总表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_archive_export_existing_file_or_zip(self) -> None:
        filename = self._askopenfilename(
            title="选择已有公司档案表或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_archive_export_existing_folder(self) -> None:
        directory = self._askdirectory(title="选择已有公司档案表文件夹")
        if directory:
            self.summary_path.set(directory)

    def _choose_roster_summary_folder(self) -> None:
        directory = self._askdirectory(title="选择异动汇总表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_roster_summary_files(self) -> None:
        filenames = self._askopenfilenames(
            title="选择异动汇总表文件或压缩包",
            filetypes=EXCEL_ARCHIVE_FILETYPES,
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _set_change_input_paths(self, paths: list[Path]) -> None:
        # 合并后的上传入口支持“＋ 添加”累加选择，重复项按路径去重
        current = list(self.change_input_paths or [])
        for path in paths:
            if path not in current:
                current.append(path)
        self.change_input_paths = current or None
        self._sync_input_path_text()
        if paths:
            self._remember_file_dialog_path(paths)
        if not self.output_dir_user_selected:
            self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
        self._refresh_upload_card()

    def _sync_input_path_text(self) -> None:
        paths = self.change_input_paths or []
        if not paths:
            self.input_path.set("")
        elif len(paths) == 1:
            self.input_path.set(str(paths[0]))
        else:
            self.input_path.set(f"已选择 {len(paths)} 个文件")

    def _choose_change_summary_folder(self) -> None:
        directory = self._askdirectory(title="选择已有异动汇总表文件夹")
        if directory:
            self.summary_path.set(directory)

    def _choose_change_summary_file(self) -> None:
        filename = self._askopenfilename(
            title="选择已有异动汇总表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_roster_analysis_file(self) -> None:
        filename = self._askopenfilename(
            title="选择人力资源花名册",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_output(self) -> None:
        self._open_workspace_root()

    def _update_project_output_controls(self) -> None:
        """Keep the legacy form row, but make the project the only formal output target."""

        project_path = self.current_project_path
        self.output_dir_user_selected = True
        self.output_dir.set(str(project_path) if project_path is not None else "")
        self.output_display_path.set(
            "当前项目 / 本次处理结果" if project_path is not None else "请先新建或打开工作项目"
        )
        if not hasattr(self, "output_label_widget"):
            return
        self.output_label_widget.configure(text="结果位置")
        entry = self._form_rows.get("output", {}).get("entry")
        if entry is not None:
            try:
                entry.configure(textvariable=self.output_display_path, state="disabled")
            except Exception:
                pass
        self.output_choose_button.configure(
            text="打开项目",
            command=self._open_workspace_root,
            state="normal" if project_path is not None else "disabled",
        )
        if hasattr(self, "open_button"):
            can_open_result = self.last_output_dir is not None and self.last_output_dir.exists()
            self.open_button.configure(state="normal" if can_open_result else "disabled")

    def _choose_summary(self) -> None:
        if self.current_tool == "social_security":
            self._choose_social_security_roster_file()
            return
        if self.current_tool == "data_statistics":
            self._choose_data_statistics_staff_file()
            return
        if self.current_tool == "insurance_ledger":
            self._choose_insurance_roster_file()
            return
        if self.current_tool == "personnel_change_merge":
            if self.change_mode == "roster":
                self._choose_roster_analysis_file()
            else:
                self._choose_change_summary_file()
            return
        elif self.current_tool == "archive_import":
            title = "选择档案汇总表"
        else:
            title = "选择已有汇总表"
        filename = self._askopenfilename(
            title=title,
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _run_current_tool(self) -> None:
        if self._tool_running:
            self._stop_tool_run()
            return
        if getattr(self, "_workspace_recovery_blocked", False):
            messagebox.showerror(
                "项目需要恢复",
                "上次资料保存没有完成安全恢复。为避免覆盖或遗漏资料，请关闭工具后重新打开当前项目，"
                "再开始新的处理。",
                parent=self.root,
            )
            return
        if self._history_task_by_token or self._project_batch_by_token or self._workspace_write_in_progress():
            messagebox.showwarning(
                "正在安全结束",
                "上一项处理或资料保存还在安全结束，请稍等片刻再开始新的处理。",
                parent=self.root,
            )
            return
        if self.current_project_path is None or self.project_store is None:
            messagebox.showerror(
                "请先打开工作项目",
                "请先在左侧“工作项目”中新建或打开一个项目。\n\n"
                "上传资料和处理结果会自动保存在项目中，方便以后查找和追溯。",
                parent=self.root,
            )
            return
        if self._workspace_project_read_only or not bool(getattr(self.project_store, "writable", False)):
            reason = getattr(getattr(self.project_store, "workspace", None), "read_only_reason", None)
            detail = f"\n\n原因：{reason}" if reason else ""
            messagebox.showerror(
                "当前项目只能查看",
                "这个项目目前是只读状态，不能新增处理批次。请关闭其他正在使用该项目的窗口后重试。" + detail,
                parent=self.root,
            )
            return
        if self.current_tool == "folder_rename":
            self._run_folder_rename()
            return
        if self.current_tool == "social_security":
            self._run_social_security()
            return
        if self.current_tool == "data_statistics":
            self._run_data_statistics()
            return
        if self.current_tool == "insurance_ledger":
            self._run_insurance_ledger()
            return
        if self.current_tool == "archive_import":
            if self.archive_mode == "export":
                self._run_archive_export()
            else:
                self._run_archive_import()
            return
        if self.current_tool == "personnel_change_merge":
            if self.change_mode == "roster":
                self._run_roster_update()
            else:
                self._run_personnel_change_merge()
            return
        if self.current_tool == "salary_merge":
            self._run_salary_merge()
            return
        if self.current_tool == "material_collector":
            self._run_material_collector()
            return
        self._run_salary_split()

    def _prepare_result_output_dir(self, parent_dir: Path) -> Path | None:
        # 工具表单仍会在开始前调用这个方法；项目模式下只返回占位路径，
        # 真正的批次结果目录由 _start_tool_worker 在输入快照完成后创建。
        if getattr(self, "project_store", None) is not None and getattr(self, "current_project_path", None) is not None:
            return Path(self.current_project_path)
        try:
            return make_result_output_dir(parent_dir)
        except (OSError, RuntimeError) as exc:
            runlog.log_exception("创建结果保存目录失败", exc)
            messagebox.showerror(
                "无法创建保存目录",
                "无法在所选位置创建结果文件夹。\n\n"
                "请检查该位置是否有写入权限、磁盘空间是否充足，或重新选择保存位置。\n\n"
                f"原因：{exc}",
                parent=self.root,
            )
            return None

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
        if input_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "当前工具只支持 .xlsx 或 .xls 工资表。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = self._prepare_result_output_dir(output_parent_dir)
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始拆分，请稍候...")

        self._start_tool_worker(split_salary_by_company, input_path, output_dir)

    def _run_salary_merge(self) -> None:
        input_text = self.input_path.get().strip()
        summary_text = self.summary_path.get().strip()
        summary_path = Path(summary_text) if summary_text else None
        output_text = self.output_dir.get().strip()
        input_paths = self.change_input_paths
        if not input_paths and input_text and not input_text.startswith("已选择 "):
            input_paths = [Path(input_text)]
        if not input_paths:
            messagebox.showwarning("缺少输入", "请先选择工资表文件、压缩包或文件夹。")
            return
        for input_path in input_paths:
            if not input_path.exists():
                messagebox.showwarning("输入不存在", "选择的工资表文件、压缩包或文件夹不存在，请重新选择。")
                return
            if input_path.is_file() and not _is_excel_or_archive_file(input_path):
                messagebox.showwarning("格式不支持", f"工资表文件只支持 {EXCEL_ARCHIVE_FORMAT_TEXT}。")
                return
        if summary_path is not None and not summary_path.exists():
            messagebox.showwarning("汇总表不存在", "选择的已有汇总表不存在，请重新选择。")
            return
        if summary_path is not None and summary_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "已有汇总表只支持 .xlsx 或 .xls 文件。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = self._prepare_result_output_dir(output_parent_dir)
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始合并，请稍候...")

        self._start_tool_worker(merge_monthly_salary, input_paths, output_dir, existing_summary_path=summary_path)

    def _run_social_security(self) -> None:
        input_text = self.input_path.get().strip()
        roster_text = self.summary_path.get().strip()
        output_text = self.output_dir.get().strip()
        input_paths = self.change_input_paths
        if not input_paths and input_text and not input_text.startswith("已选择 "):
            input_paths = [Path(input_text)]
        if not input_paths:
            messagebox.showwarning("缺少输入", "请先选择社保缴费清单文件、压缩包或文件夹。")
            return
        for input_path in input_paths:
            if not input_path.exists():
                messagebox.showwarning("输入不存在", "选择的社保缴费清单文件、压缩包或文件夹不存在，请重新选择。")
                return
            if input_path.is_file() and not _is_excel_or_archive_file(input_path):
                messagebox.showwarning("格式不支持", f"社保缴费清单只支持 {EXCEL_ARCHIVE_FORMAT_TEXT}。")
                return
        if not roster_text:
            messagebox.showwarning("缺少花名册", "请先选择参保人员花名册。")
            return
        roster_path = Path(roster_text)
        if not roster_path.exists() or not roster_path.is_file():
            messagebox.showwarning("花名册不存在", "选择的参保人员花名册不存在，请重新选择。")
            return
        if roster_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "参保人员花名册只支持 .xlsx 或 .xls。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return

        output_dir = self._prepare_result_output_dir(Path(output_text))
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始生成社保报表，请稍候...")

        self._start_tool_worker(generate_social_security_reports, input_paths, roster_path, output_dir)

    def _run_data_statistics(self) -> None:
        input_text = self.input_path.get().strip()
        staff_text = self.summary_path.get().strip()
        staff_path = Path(staff_text) if staff_text else None
        output_text = self.output_dir.get().strip()
        input_paths = self.change_input_paths
        if not input_paths and input_text and not input_text.startswith("已选择 "):
            input_paths = [Path(input_text)]
        if not input_paths:
            messagebox.showwarning("缺少输入", "请先选择考勤结果、周报记录、月报记录文件、压缩包或文件夹。")
            return
        for input_path in input_paths:
            if not input_path.exists():
                messagebox.showwarning("输入不存在", "选择的数据文件、压缩包或文件夹不存在，请重新选择。")
                return
            if input_path.is_file() and not _is_excel_or_archive_file(input_path):
                messagebox.showwarning("格式不支持", f"数据文件只支持 {EXCEL_ARCHIVE_FORMAT_TEXT}。")
                return
        if staff_path is not None and (not staff_path.exists() or not staff_path.is_file()):
            messagebox.showwarning("名单不存在", "选择的应汇报人员名单不存在，请重新选择。")
            return
        if staff_path is not None and staff_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "应汇报人员名单只支持 .xlsx 或 .xls。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        try:
            week_range = resolve_week_range(
                self.stats_week_start.get().strip() or None,
                self.stats_week_end.get().strip() or None,
            )
        except ValueError as exc:
            messagebox.showwarning("日期填写有误", str(exc))
            return
        try:
            month_range = resolve_month_range(
                self.stats_month_start.get().strip() or None,
                self.stats_month_end.get().strip() or None,
            )
        except ValueError as exc:
            messagebox.showwarning("日期填写有误", str(exc))
            return

        output_dir = self._prepare_result_output_dir(Path(output_text))
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始生成统计表，请稍候...")

        self._start_tool_worker(
            generate_data_statistics_reports,
            input_paths,
            output_dir,
            report_staff_path=staff_path,
            week_start=None if week_range is None else week_range[0],
            week_end=None if week_range is None else week_range[1],
            month_start=None if month_range is None else month_range[0],
            month_end=None if month_range is None else month_range[1],
            remark_unit=self.stats_remark_unit.get() or "day",
            include_business_trip=bool(self.stats_include_business_trip.get()),
            include_workday_business_trip=bool(self.stats_include_workday_business_trip.get()),
        )

    def _run_insurance_ledger(self) -> None:
        input_text = self.input_path.get().strip()
        roster_text = self.summary_path.get().strip()
        output_text = self.output_dir.get().strip()
        input_paths = self.change_input_paths
        if not input_paths and input_text and not input_text.startswith("已选择 "):
            input_paths = [Path(input_text)]
        if not input_paths:
            messagebox.showwarning("缺少输入", "请先选择保单人员清单文件、压缩包或文件夹。")
            return
        for input_path in input_paths:
            if not input_path.exists():
                messagebox.showwarning("输入不存在", "选择的保单人员清单文件、压缩包或文件夹不存在，请重新选择。")
                return
            if input_path.is_file() and not _is_excel_or_archive_file(input_path):
                messagebox.showwarning("格式不支持", f"保单人员清单只支持 {EXCEL_ARCHIVE_FORMAT_TEXT}。")
                return
        if not roster_text:
            messagebox.showwarning("缺少分析表", "请先选择人力资源分析表。")
            return
        roster_path = Path(roster_text)
        if not roster_path.exists() or not roster_path.is_file():
            messagebox.showwarning("分析表不存在", "选择的人力资源分析表不存在，请重新选择。")
            return
        if roster_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "人力资源分析表只支持 .xlsx 或 .xls。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return

        output_dir = self._prepare_result_output_dir(Path(output_text))
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始生成保险台账，请稍候...")

        self._start_tool_worker(generate_insurance_ledger, input_paths, roster_path, output_dir)

    def _run_personnel_change_merge(self) -> None:
        input_text = self.input_path.get().strip()
        summary_text = self.summary_path.get().strip()
        summary_path = Path(summary_text) if summary_text else None
        output_text = self.output_dir.get().strip()
        input_paths = self.change_input_paths
        if not input_paths and input_text and not input_text.startswith("已选择 "):
            input_paths = [Path(input_text)]
        if not input_paths:
            messagebox.showwarning("缺少输入", "请先选择异动表文件或文件夹。")
            return
        for input_path in input_paths:
            if not input_path.exists():
                messagebox.showwarning("输入不存在", "选择的异动表文件、压缩包或文件夹不存在，请重新选择。")
                return
            if input_path.is_file() and not _is_excel_or_archive_file(input_path):
                messagebox.showwarning("格式不支持", f"异动表文件只支持 {EXCEL_ARCHIVE_FORMAT_TEXT}。")
                return
        if summary_path is not None and not summary_path.exists():
            messagebox.showwarning("汇总表不存在", "选择的已有异动汇总表不存在，请重新选择。")
            return
        if summary_path is not None and summary_path.is_file() and summary_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "已有异动汇总表只支持 .xlsx、.xls 文件或文件夹。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = self._prepare_result_output_dir(output_parent_dir)
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始汇总，请稍候...")

        self._start_tool_worker(merge_personnel_changes, input_paths, output_dir, template_path=summary_path)

    def _run_roster_update(self) -> None:
        input_text = self.input_path.get().strip()
        roster_text = self.summary_path.get().strip()
        output_text = self.output_dir.get().strip()
        input_paths = self.change_input_paths
        if not input_paths and input_text and not input_text.startswith("已选择 "):
            input_paths = [Path(input_text)]
        if not input_paths:
            messagebox.showwarning("缺少汇总表", "请先选择异动汇总表文件、压缩包或文件夹。")
            return
        for input_path in input_paths:
            if not input_path.exists():
                messagebox.showwarning("汇总表不存在", "选择的异动汇总表文件、压缩包或文件夹不存在，请重新选择。")
                return
            if input_path.is_file() and not _is_excel_or_archive_file(input_path):
                messagebox.showwarning("格式不支持", f"异动汇总表只支持 {EXCEL_ARCHIVE_FORMAT_TEXT} 或文件夹。")
                return
        if not roster_text:
            messagebox.showwarning("缺少花名册", "请先选择人力资源花名册。")
            return
        roster_path = Path(roster_text)
        if not roster_path.exists() or not roster_path.is_file():
            messagebox.showwarning("花名册不存在", "选择的人力资源花名册不存在，请重新选择。")
            return
        if roster_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "人力资源花名册目前只支持 .xlsx 或 .xls 文件。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_dir = self._prepare_result_output_dir(Path(output_text))
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始更新花名册，请稍候...")

        self._start_tool_worker(update_roster_from_change_summaries, input_paths, roster_path, output_dir)

    def _run_archive_import(self) -> None:
        input_text = self.input_path.get().strip()
        target_text = self.summary_path.get().strip()
        target_path = Path(target_text) if target_text else None
        output_text = self.output_dir.get().strip()
        input_paths = self.change_input_paths
        if not input_paths and input_text and not input_text.startswith("已选择 "):
            input_paths = [Path(input_text)]
        if not input_paths:
            messagebox.showwarning("缺少输入", "请先选择档案移交表文件、压缩包或文件夹。")
            return
        for input_path in input_paths:
            if not input_path.exists():
                messagebox.showwarning("输入不存在", "选择的档案移交表文件、压缩包或文件夹不存在，请重新选择。")
                return
            if input_path.is_file() and not _is_excel_or_archive_file(input_path):
                messagebox.showwarning("格式不支持", f"档案移交表文件只支持 {EXCEL_ARCHIVE_FORMAT_TEXT}。")
                return
        if target_path is not None and (not target_path.exists() or not target_path.is_file()):
            messagebox.showwarning("汇总表不存在", "选择的档案汇总表不存在，请重新选择。")
            return
        if target_path is not None and target_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "档案汇总表目前只支持 .xlsx 或 .xls 文件。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_dir = self._prepare_result_output_dir(Path(output_text))
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始入库，请稍候...")

        self._start_tool_worker(import_archive_transfers, input_paths, target_path, output_dir)

    def _run_archive_export(self) -> None:
        input_text = self.input_path.get().strip()
        existing_text = self.summary_path.get().strip()
        output_text = self.output_dir.get().strip()
        input_paths = self.change_input_paths
        if not input_paths and input_text and not input_text.startswith("已选择 "):
            input_paths = [Path(input_text)]
        if not input_paths:
            messagebox.showwarning("缺少汇总表", "请先选择档案汇总表。")
            return
        for summary_path in input_paths:
            if not summary_path.exists():
                messagebox.showwarning("汇总表不存在", "选择的档案汇总表不存在，请重新选择。")
                return
            if summary_path.is_file() and not _is_excel_or_archive_file(summary_path):
                messagebox.showwarning("格式不支持", f"档案汇总表目前只支持 {EXCEL_ARCHIVE_FORMAT_TEXT}。")
                return
        existing_path = Path(existing_text) if existing_text else None
        if existing_path is not None and not existing_path.exists():
            messagebox.showwarning("档案表不存在", "选择的已有公司档案表不存在，请重新选择。")
            return
        if existing_path is not None and existing_path.is_file() and not _is_excel_or_archive_file(existing_path):
            messagebox.showwarning("格式不支持", f"已有公司档案表目前只支持 {EXCEL_ARCHIVE_FORMAT_TEXT}。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_dir = self._prepare_result_output_dir(Path(output_text))
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        self._write_log("开始生成档案表，请稍候...")

        self._start_tool_worker(export_company_archive_tables, input_paths, output_dir, existing_archive_path=existing_path)

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
        file_type_label = self.rename_file_type.get()
        file_type = RENAME_FILE_TYPE_LABELS.get(file_type_label, FILE_TYPE_FOLDER)
        excel_path: Path | None = None
        if mode == MODE_EXCEL_BATCH:
            excel_text = self.summary_path.get().strip()
            if not excel_text:
                messagebox.showwarning("缺少人员名单", "请先选择包含姓名列的 Excel 名单。")
                return
            excel_path = Path(excel_text)
            if not excel_path.exists() or not excel_path.is_file():
                messagebox.showwarning("名单不存在", "选择的人员名单不存在，请重新选择。")
                return
            if excel_path.suffix.lower() not in {".xlsx", ".xls"}:
                messagebox.showwarning("格式不支持", "人员名单只支持 .xlsx 或 .xls 文件。")
                return

        try:
            if mode == MODE_EXCEL_BATCH:
                assert excel_path is not None
                preview = rename_files_by_excel(
                    root_dir=root_dir,
                    excel_path=excel_path,
                    file_type=file_type,
                    dry_run=True,
                )
            else:
                preview = rename_person_folders(
                    root_dir=root_dir,
                    mode=mode,
                    text=self.rename_text.get(),
                    target_name=self.rename_target_name.get(),
                    replacement_name=self.rename_replacement_name.get(),
                    file_type=file_type,
                    dry_run=True,
                )
        except Exception as exc:
            messagebox.showerror("预览失败", str(exc))
            return

        self._clear_log()
        self._write_log("预览结果：")
        self._write_folder_rename_preview(preview)
        if preview.operation_count == 0:
            warning_text = "\n".join(preview.warnings[:8])
            message = "没有找到可以安全改名的项目，请检查目录、名单、文件类型和预览提醒。"
            if warning_text:
                message += f"\n\n提醒：\n{warning_text}"
            messagebox.showinfo("没有可改名项目", message)
            return

        message = self._folder_rename_confirm_message(preview)
        if not messagebox.askyesno("确认改名", message):
            self._write_log("已取消执行。")
            return

        self._begin_tool_run()
        self._write_log("开始执行改名...")
        if mode == MODE_EXCEL_BATCH:
            assert excel_path is not None
            self._start_tool_worker(
                rename_files_by_excel,
                root_dir=root_dir,
                excel_path=excel_path,
                file_type=file_type,
                expected_operations=[
                    (operation.source.name, operation.target.name)
                    for operation in preview.operations
                ],
                expected_warnings=list(preview.warnings),
            )
        else:
            self._start_tool_worker(
                rename_person_folders,
                root_dir=root_dir,
                mode=mode,
                text=self.rename_text.get(),
                target_name=self.rename_target_name.get(),
                replacement_name=self.rename_replacement_name.get(),
                file_type=file_type,
            )

    def _run_material_collector(self) -> None:
        input_text = self.input_path.get().strip()
        direct_target_text = self.material_target_input.get().strip()
        roster_text = self.summary_path.get().strip()
        output_text = self.output_dir.get().strip()

        if not input_text:
            messagebox.showwarning("缺少资料库", "请先选择员工资料库根目录。")
            return
        lib_path = Path(input_text)
        if not lib_path.exists() or not lib_path.is_dir():
            messagebox.showwarning("资料库不存在", "选择的资料库根目录不存在，请重新选择。")
            return

        # 确定名单来源：输入框直接指定优先，其次为上传的 Excel 名单
        roster_source: str | Path | None = None
        if direct_target_text:
            roster_source = direct_target_text
        elif roster_text:
            roster_path = Path(roster_text)
            if not roster_path.exists() or not roster_path.is_file():
                messagebox.showwarning("名单文件不存在", "选择的员工名单文件不存在，请重新选择。")
                return
            if roster_path.suffix.lower() not in {".xlsx", ".xls"}:
                messagebox.showwarning("格式不支持", "员工名单只支持 .xlsx 或 .xls 文件。")
                return
            roster_source = roster_path
        else:
            messagebox.showwarning(
                "缺少员工信息",
                "请在上方输入框直接输入员工姓名/身份证（如“张三”），或在下方选择员工名单 Excel 表格。",
            )
            return

        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        out_candidate = Path(output_text).resolve()
        try:
            if out_candidate == lib_path.resolve() or out_candidate.is_relative_to(lib_path.resolve()):
                messagebox.showwarning(
                    "保存目录无效",
                    "保存目录不能设在资料库目录内部（会导致循环嵌套复制）。\n请选择一个位于资料库外部的独立保存文件夹。",
                )
                return
        except Exception:
            pass

        is_collect_all = self.material_collect_all.get()

        selected_materials: list[str] | None = None
        if not is_collect_all:
            selected_materials = self._selected_material_names()
            if not selected_materials:
                messagebox.showwarning(
                    "未选择材料",
                    "请至少勾选一种需要提取的材料，或者勾选「全部」直接拷贝整个文件夹。",
                )
                return

        create_zip_val = self.material_create_zip.get()
        library_mode_label = (
            self.material_library_mode.get()
            if hasattr(self, "material_library_mode")
            else "按人员文件夹查找（原模式）"
        )
        library_mode_val = LIBRARY_MODE_LABELS.get(
            library_mode_label,
            LIBRARY_MODE_PERSON_FOLDER,
        )
        use_ocr_cache_val = (
            True
            if library_mode_val == LIBRARY_MODE_FLAT_OCR
            else self.material_use_ocr_cache.get()
        )

        output_dir = self._prepare_result_output_dir(Path(output_text))
        if output_dir is None:
            return
        self._begin_tool_run()
        self._clear_log()
        if library_mode_val == LIBRARY_MODE_FLAT_OCR:
            self._write_log("开始扫描无序资料库并建立/复用人员材料索引，请稍候...")
        elif is_collect_all:
            self._write_log("开始检索并拷贝员工整个资料文件夹，请稍候...")
        else:
            self._write_log("开始检索并打包指定材料，请稍候...")
        if use_ocr_cache_val:
            self._write_log("OCR 智能索引缓存已启用：首次扫描会建立资料库缓存；二次扫描将秒级命中。")

        run_token = getattr(self, "_tool_run_token", 0)
        cancel_event = getattr(self, "_run_cancel_events", {}).get(run_token)
        progress_queue = getattr(self, "status_queue", None)
        last_progress: dict[str, tuple[int, int]] = {}

        def material_progress(current: int, total: int, message: str) -> None:
            if cancel_event is not None and cancel_event.is_set():
                raise RuntimeError("本次处理已停止。")
            phase = "index" if "索引" in message else "employee"
            interval = max(1, total // 20) if phase == "index" else max(1, total // 10)
            marker = (current, total)
            previous = last_progress.get(phase)
            should_report = current in {0, 1, total} or current % interval == 0
            if should_report and marker != previous:
                last_progress[phase] = marker
                if progress_queue is not None:
                    progress_queue.put(("progress", run_token, message))

        self._start_tool_worker(
            collect_employee_materials,
            lib_path,
            output_dir,
            roster_source=roster_source,
            material_types=selected_materials,
            mode=MODE_BY_EMPLOYEE,
            library_mode=library_mode_val,
            create_zip=create_zip_val,
            generate_report=True,
            collect_all=is_collect_all,
            use_ocr_cache=use_ocr_cache_val,
            progress_callback=material_progress,
        )

    def _begin_tool_run(self) -> None:
        """进入运行状态：主按钮变为“停止”，并为本次运行分配编号。"""
        self._tool_run_token += 1
        self._tool_running = True
        self._run_cancel_events[self._tool_run_token] = threading.Event()
        self._idle_run_button_text = self.run_button_text.get()
        self.run_button_text.set("停止")

    def _finish_tool_run(self) -> None:
        self._tool_running = False
        if self._idle_run_button_text:
            self.run_button_text.set(self._idle_run_button_text)

    def _stop_tool_run(self) -> None:
        token = self._tool_run_token
        cancel_event = self._run_cancel_events.get(token)
        if cancel_event is not None:
            cancel_event.set()
        task_id = self._history_task_by_token.get(token)
        if task_id is not None and self.history_store is not None:
            try:
                self.history_store.mark_stopped(task_id)
            except Exception as exc:
                runlog.log_exception("保存停止状态失败", exc)
        # 项目批次不能在工具仍可能写文件时提前结案。后台线程会在安全点将其
        # 标记为“未完成”，并把未登记的半成品移入隐藏隔离区。
        self._tool_run_token += 1
        self._finish_tool_run()
        self._write_log("已停止本次处理，后台正在安全结束。")
        runlog.log_line(f"用户停止了 {self._tool_log_label()}。")

    def _tool_log_label(self) -> str:
        if self.current_tool == "personnel_change_merge" and self.change_mode == "roster":
            return "花名册更新"
        if self.current_tool == "archive_import" and self.archive_mode == "export":
            return "档案表生成"
        return TOOL_LOG_LABELS.get(self.current_tool, self.current_tool)

    def _history_mode_for_current_tool(self) -> str | None:
        if self.current_tool == "personnel_change_merge":
            return self.change_mode
        if self.current_tool == "archive_import":
            return self.archive_mode
        if self.current_tool == "folder_rename":
            return RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND)
        return None

    @staticmethod
    def _history_serializable(value):
        if isinstance(value, Path):
            return str(value)
        if isinstance(value, (date, datetime)):
            return value.isoformat()
        if isinstance(value, (list, tuple)):
            return [HRToolkitApp._history_serializable(item) for item in value]
        if isinstance(value, dict):
            return {str(key): HRToolkitApp._history_serializable(item) for key, item in value.items()}
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        return str(value)

    def _history_context_from_call(self, tool_func, args, kwargs):
        bound = inspect.signature(tool_func).bind_partial(*args, **kwargs)
        bound.apply_defaults()
        parameters: dict[str, object] = {}
        sources: list[SourceSpec] = []
        output_dir: Path | None = None
        for name, value in bound.arguments.items():
            if name == "progress_callback":
                continue
            if name == "output_dir" and value is not None:
                output_dir = Path(value).expanduser()
                parameters[name] = output_dir.name
                continue
            if name == "root_dir" and value is not None:
                root_dir = Path(value).expanduser()
                parameters[name] = root_dir.name
                if root_dir.exists():
                    sources.append(
                        SourceSpec(
                            path=root_dir,
                            role="input_path",
                            suffixes=None,
                            preserve_directories=True,
                        )
                    )
                    output_dir = root_dir
                continue
            if name == "library_dir" and value is not None:
                parameters[name] = str(value)
                continue
            if name == "roster_source" and not (isinstance(value, Path) or (isinstance(value, str) and Path(value).expanduser().is_file())):
                parameters[name] = self._history_serializable(value)
                continue
            if name not in HISTORY_PATH_ARGUMENTS or value is None:
                parameters[name] = self._history_serializable(value)
                continue
            raw_paths = value if isinstance(value, (list, tuple)) else [value]
            parameters[name] = [Path(item).name for item in raw_paths if item is not None]
            if not isinstance(value, (list, tuple)):
                parameters[name] = parameters[name][0] if parameters[name] else None
            role = "input_path" if name in HISTORY_PRIMARY_PATH_ARGUMENTS else name
            for raw_path in raw_paths:
                if raw_path is None:
                    continue
                path = Path(raw_path).expanduser()
                if path.exists():
                    sources.append(SourceSpec(path=path, role=role))
        return sources, parameters, output_dir

    @classmethod
    def _history_result_for_storage(cls, value, *, key: str = ""):
        if isinstance(value, dict):
            return {
                str(item_key): cls._history_result_for_storage(item, key=str(item_key))
                for item_key, item in value.items()
            }
        if isinstance(value, (list, tuple)):
            return [cls._history_result_for_storage(item, key=key) for item in value]
        if isinstance(value, Path):
            return value.name
        if isinstance(value, str):
            normalized_key = key.lower()
            if any(token in normalized_key for token in ("path", "file", "dir")):
                return Path(value).name
            return value
        return cls._history_serializable(value)

    def _call_with_archived_inputs(self, tool_func, args, kwargs, archived_records, task_id: str):
        if self.history_store is None:
            raise HistoryStoreError("资料留存功能不可用。")
        bound = inspect.signature(tool_func).bind_partial(*args, **kwargs)
        bound.apply_defaults()
        detail = self.history_store.get_task(task_id)
        if detail is None:
            raise HistoryStoreError("本次历史记录无法读取。")
        self.history_store.verify_task_files(task_id, kind="input")
        records_by_role: dict[str, list[Path]] = {}
        for record in archived_records:
            records_by_role.setdefault(record.role, []).append(record.archived_path)

        for name, value in tuple(bound.arguments.items()):
            if name == "library_dir":
                continue
            if name == "roster_source" and not (isinstance(value, Path) or (isinstance(value, str) and Path(value).expanduser().is_file())):
                continue
            if name not in HISTORY_PATH_ARGUMENTS or value is None:
                continue
            role = "input_path" if name in HISTORY_PRIMARY_PATH_ARGUMENTS else name
            archived_paths = records_by_role.get(role, [])
            if not archived_paths:
                raise HistoryStoreError(f"没有完整保存 {name} 对应的原始资料。")
            original_values = value if isinstance(value, (list, tuple)) else [value]
            original_paths = [Path(item).expanduser() for item in original_values if item is not None]
            original_was_directory = len(original_paths) == 1 and original_paths[0].is_dir()

            if name == "template_path" and original_was_directory:
                source_name = original_paths[0].name
                replacement = next(
                    (
                        parent
                        for parent in (archived_paths[0].parent, *archived_paths[0].parents)
                        if parent != detail.input_dir
                        and parent.is_relative_to(detail.input_dir)
                        and parent.name == source_name
                    ),
                    None,
                )
                if replacement is None:
                    common_path = Path(os.path.commonpath([str(path) for path in archived_paths]))
                    replacement = common_path if common_path.is_dir() else common_path.parent
            elif isinstance(value, (list, tuple)) or original_was_directory:
                replacement = archived_paths
            else:
                replacement = archived_paths[0]
            bound.arguments[name] = replacement
        return bound.args, bound.kwargs

    def _request_close(self) -> None:
        has_background_work = (
            self._tool_running
            or bool(self._history_task_by_token)
            or bool(self._project_batch_by_token)
            or self._workspace_write_in_progress()
        )
        if has_background_work and not messagebox.askyesno(
            "处理尚未结束",
            "当前处理或资料保存还没有完全结束。现在退出会先安全停止正在保存的资料，"
            "并把未完成的处理留待下次打开时恢复。是否仍要退出？",
            parent=self.root,
        ):
            return
        for cancel_event in self._run_cancel_events.values():
            cancel_event.set()
        for cancel_event, _store in self._workspace_write_tasks.values():
            cancel_event.set()
        if self.history_store is not None:
            for task_id in set(self._history_task_by_token.values()):
                try:
                    self.history_store.mark_stopped(task_id)
                except Exception as exc:
                    runlog.log_exception("退出时保存任务状态失败", exc)
        if self._workspace_write_in_progress():
            self._workspace_close_requested = True
            self._update_workspace_action_states()
            return
        self._finish_app_close()

    def _finish_app_close(self) -> None:
        if not getattr(self, "_workspace_close_requested", False):
            self._workspace_close_requested = True
        self._save_workspace_preferences()
        # 正在运行时保留项目写锁直到进程结束，避免后台线程与另一个窗口同时
        # 写入；下次打开项目时会把残留 running 批次恢复为“未完成”。
        if self.project_store is not None and not self._project_batch_by_token:
            try:
                self.project_store.close()
            except Exception as exc:
                runlog.log_exception("关闭工作项目失败", exc)
        self.root.destroy()

    def _project_source_replacement(
        self,
        store,
        batch_id: str,
        source: SourceSpec,
        records,
        *,
        source_was_file: bool,
    ) -> Path:
        """Resolve one imported source back to the file/folder shape expected by a tool."""

        if not records:
            raise RuntimeError(f"没有可留存的资料：{source.path.name}")
        paths = [record.path(store.workspace) for record in records]
        if source_was_file:
            if len(paths) != 1:
                raise RuntimeError(f"资料快照不完整：{source.path.name}")
            return paths[0]
        detail = store.get_batch(batch_id)
        if detail is None:
            raise RuntimeError("本次处理批次无法读取。")
        upload_root = detail.directories["uploads"]
        top_names: set[str] = set()
        for path in paths:
            relative = path.relative_to(upload_root)
            if not relative.parts:
                raise RuntimeError(f"资料快照位置无效：{source.path.name}")
            top_names.add(relative.parts[0])
        if len(top_names) != 1:
            raise RuntimeError(f"文件夹快照不完整：{source.path.name}")
        return upload_root / next(iter(top_names))

    def _import_project_run_sources(self, store, batch_id: str, sources: list[SourceSpec], cancel_event) -> dict[str, list[Path]]:
        replacements: dict[str, list[Path]] = {}
        project_root = Path(store.root).absolute()
        for source in sources:
            # 保留用户实际选择的词法路径，不能先 resolve；否则文件/目录链接会
            # 被替换成目标路径，绕过 ProjectStore 的链接与越界检查。
            source_path = Path(source.path).expanduser().absolute()
            source_was_file = source_path.is_file()
            try:
                source_path.relative_to(project_root)
                is_project_source = True
            except ValueError:
                is_project_source = False
            if source.preserve_directories:
                if source_was_file:
                    raise RuntimeError(f"需要文件夹资料：{source.path.name}")
                method_name = (
                    "copy_project_directory_snapshot"
                    if is_project_source
                    else "import_directory_snapshot"
                )
                snapshotter = getattr(store, method_name, None)
                if not callable(snapshotter):
                    raise RuntimeError("当前版本无法安全留存文件夹结构。")
                replacement = snapshotter(
                    batch_id,
                    source_path,
                    category="uploads",
                    role=source.role,
                    cancelled=cancel_event.is_set,
                )
            else:
                if is_project_source:
                    copier = getattr(store, "copy_project_sources", None)
                    if not callable(copier):
                        raise RuntimeError(
                            "当前版本暂时不能复用项目内资料，请从原文件位置重新选择。"
                        )
                else:
                    copier = store.import_sources
                records = copier(
                    batch_id,
                    [source_path],
                    category="uploads",
                    role=source.role,
                    cancelled=cancel_event.is_set,
                )
                replacement = self._project_source_replacement(
                    store,
                    batch_id,
                    source,
                    records,
                    source_was_file=source_was_file,
                )
            replacements.setdefault(source.role, []).append(replacement)
        return replacements

    @staticmethod
    def _copy_project_directory_for_result(store, batch_id: str, source: Path) -> Path:
        copier = getattr(store, "create_result_working_copy", None)
        if not callable(copier):
            raise RuntimeError("当前版本无法安全建立文件夹处理副本。")
        return copier(batch_id, source)

    @staticmethod
    def _rebase_project_replacements(
        replacements: dict[str, list[Path]],
        old_upload_root: Path,
        new_upload_root: Path,
    ) -> dict[str, list[Path]]:
        return {
            role: [new_upload_root / path.relative_to(old_upload_root) for path in paths]
            for role, paths in replacements.items()
        }

    def _call_with_project_inputs(
        self,
        tool_func,
        args,
        kwargs,
        replacements: dict[str, list[Path]],
        result_dir: Path,
        store,
        batch_id: str,
    ):
        bound = inspect.signature(tool_func).bind_partial(*args, **kwargs)
        bound.apply_defaults()
        for name, value in tuple(bound.arguments.items()):
            if name == "output_dir":
                bound.arguments[name] = result_dir
                continue
            if name == "library_dir":
                continue
            if name == "roster_source" and not (isinstance(value, Path) or (isinstance(value, str) and Path(value).expanduser().is_file())):
                continue
            if value is None or name not in HISTORY_PATH_ARGUMENTS | {"root_dir"}:
                continue
            role = "input_path" if name in HISTORY_PRIMARY_PATH_ARGUMENTS or name == "root_dir" else name
            copied_paths = replacements.get(role, [])
            if not copied_paths:
                raise RuntimeError(f"没有完整保存 {name} 对应的原始资料。")
            original_values = value if isinstance(value, (list, tuple)) else [value]
            original_paths = [Path(item).expanduser() for item in original_values if item is not None]
            original_was_directory = len(original_paths) == 1 and original_paths[0].is_dir()
            if name == "root_dir":
                if len(copied_paths) != 1 or not copied_paths[0].is_dir():
                    raise RuntimeError("人员资料文件夹快照不完整。")
                replacement = self._copy_project_directory_for_result(store, batch_id, copied_paths[0])
            elif name == "template_path" and original_was_directory:
                replacement = copied_paths[0]
            elif isinstance(value, (list, tuple)) or original_was_directory:
                replacement = copied_paths
            else:
                replacement = copied_paths[0]
            bound.arguments[name] = replacement
        return bound.args, bound.kwargs

    @staticmethod
    def _project_batch_is_closed(store, batch_id: str) -> bool:
        detail = store.get_batch(batch_id)
        if detail is not None:
            return detail.summary.status in {"success", "failed", "stopped"}
        return any(summary.id == batch_id for summary in store.list_trash())

    def _start_tool_worker(self, tool_func, /, *args, **kwargs) -> None:
        token = self._tool_run_token
        label = self._tool_log_label()
        cancel_event = self._run_cancel_events.setdefault(token, threading.Event())
        store = self.project_store
        if store is None or self.current_project_path is None or not bool(getattr(store, "writable", False)):
            self._finish_tool_run()
            self._run_cancel_events.pop(token, None)
            self._show_error_after_log("无法开始处理", "请先打开一个可写的工作项目。")
            return

        try:
            sources, _parameters, _legacy_output = self._history_context_from_call(tool_func, args, kwargs)
            project_tool_id, project_tool_name = self._project_tool_identity()
            description = project_tool_name
            if self.current_tool == "folder_rename":
                description = f"{project_tool_name}-{self.rename_mode.get()}"
            period = date.today().strftime("%Y-%m-%d")
            draft = store.create_draft(
                group_name=TOOL_GROUP_LABELS.get(self.current_tool, "人员运营自动化"),
                tool_id=project_tool_id,
                tool_name=project_tool_name,
                business_description=description,
                business_period=period,
            )
            batch_id = draft.summary.id
            runlog.log_line(f"开始 {label}（{len(sources)} 个资料来源，自动留存在当前项目）")
        except Exception as exc:
            self._finish_tool_run()
            self._run_cancel_events.pop(token, None)
            runlog.log_exception("创建项目批次失败", exc)
            self._write_log(f"无法建立本次项目记录：{exc}")
            self._show_error_after_log(
                "无法开始处理",
                f"为了避免资料无法追溯，本次处理没有开始。\n\n原因：{exc}",
            )
            return

        self._project_batch_by_token[token] = batch_id
        self._update_workspace_action_states()

        def worker() -> None:
            start = time.monotonic()
            result = None
            started = False
            try:
                replacements = self._import_project_run_sources(store, batch_id, sources, cancel_event)
                if cancel_event.is_set():
                    raise RuntimeError("本次处理已停止。")
                old_upload_root = draft.directories["uploads"]
                running = store.start_batch(batch_id)
                started = True
                replacements = self._rebase_project_replacements(
                    replacements,
                    old_upload_root,
                    running.directories["uploads"],
                )
                result_dir = store.result_directory(batch_id)
                call_args, call_kwargs = self._call_with_project_inputs(
                    tool_func,
                    args,
                    kwargs,
                    replacements,
                    result_dir,
                    store,
                    batch_id,
                )
                result = tool_func(*call_args, **call_kwargs)
                if cancel_event.is_set():
                    raise RuntimeError("本次处理已停止。")
                store.register_results(batch_id, result_dir)
                if cancel_event.is_set():
                    raise RuntimeError("本次处理已停止。")
                store.mark_success(batch_id)
                try:
                    upload_path = Path(store.root) / running.directories["uploads"]
                    if upload_path.is_dir() and not any(upload_path.iterdir()):
                        upload_path.rmdir()
                except Exception:
                    pass
            except Exception as exc:
                stopped = cancel_event.is_set()
                finalization_error = None
                try:
                    if started:
                        if stopped:
                            store.mark_stopped(batch_id)
                        else:
                            store.mark_failed(batch_id, str(exc))
                    else:
                        store.move_to_trash(batch_id)
                    if not self._project_batch_is_closed(store, batch_id):
                        raise RuntimeError("项目批次仍未进入安全结束状态。")
                except Exception as project_exc:
                    finalization_error = project_exc
                    runlog.log_exception("保存项目批次状态失败", project_exc)
                if finalization_error is not None:
                    self.status_queue.put(
                        (
                            "project_finalize_error",
                            token,
                            (exc, finalization_error),
                        )
                    )
                    return
                if stopped:
                    runlog.log_line(f"{label} 已停止，后台耗时 {time.monotonic() - start:.1f} 秒")
                    self.status_queue.put(("stopped", token, None))
                else:
                    runlog.log_exception(f"{label} 失败，耗时 {time.monotonic() - start:.1f} 秒", exc)
                    self.status_queue.put(("error", token, exc))
                return
            warnings = getattr(result, "warnings", None)
            warn_text = f"，提醒 {len(warnings)} 条" if warnings else ""
            runlog.log_line(f"完成 {label}，耗时 {time.monotonic() - start:.1f} 秒{warn_text}")
            self.status_queue.put(("success", token, result))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_status_queue(self) -> None:
        try:
            while True:
                status, token, payload = self.status_queue.get_nowait()
                if status == "progress":
                    if token == self._tool_run_token and payload:
                        self._write_log(str(payload))
                    continue
                keep_project_guard = status == "project_finalize_error"
                if not keep_project_guard:
                    self._history_task_by_token.pop(token, None)
                    self._project_batch_by_token.pop(token, None)
                    self._run_cancel_events.pop(token, None)
                if token != self._tool_run_token:
                    if status == "project_finalize_error":
                        self._write_log("（后台处理未能安全结案，请退出并重新打开工具恢复项目。）")
                    elif status == "stopped":
                        self._write_log("（已停止的后台处理已安全结束。）")
                    else:
                        self._write_log("（已停止的任务在后台结束，结果已忽略。）")
                    if self.current_view == "history":
                        self._refresh_history()
                    if self.current_project_path is not None:
                        self._refresh_workspace_tree()
                        self._update_workspace_action_states()
                        self._update_sidebar_project_summary()
                    continue
                self._finish_tool_run()
                if status == "success":
                    self._record_last_run(True)
                    self._handle_success(payload)
                elif status == "history_warning" and isinstance(payload, tuple):
                    self._record_last_run(True)
                    result, exc = payload
                    self._handle_success(result, history_warning=str(exc))
                elif status == "error":
                    self._record_last_run(False)
                    self._handle_error(payload)
                elif status == "stopped":
                    self._write_log("本次处理已停止。")
                elif status == "project_finalize_error":
                    self._record_last_run(False)
                    original_exc, finalization_exc = payload
                    runlog.log_exception("项目批次未能安全结案", finalization_exc)
                    self._write_log(f"处理没有完成：{original_exc}")
                    self._write_log("项目仍处于保护状态。请退出并重新打开工具，系统会恢复这次记录。")
                    self._show_error_after_log(
                        "项目需要重新打开",
                        "本次处理没有正常结束，项目已保持锁定，不能继续写入。\n\n"
                        "请退出并重新打开工具，系统会自动恢复未完成记录。",
                    )
                if self.current_view == "history":
                    self._refresh_history()
                if self.current_project_path is not None:
                    self._refresh_workspace_tree()
                    self._update_workspace_action_states()
                    self._update_sidebar_project_summary()
        except queue.Empty:
            pass
        self.root.after(150, self._poll_status_queue)

    def _handle_success(self, result, *, history_warning: str | None = None) -> None:
        payload = result.to_dict()
        if self.current_tool == "folder_rename":
            self.last_output_dir = Path(payload["root_dir"])
            self._write_log("改名完成。")
            self._write_folder_rename_preview(result)
            message = f"已完成 {payload['operation_count']} 个项目改名。"
        else:
            self.last_output_dir = Path(payload["output_dir"])
            if self.current_tool == "social_security":
                self._write_log("社保报表生成完成。")
                self._write_log(f"识别文件数：{payload['source_file_count']}")
                self._write_log(f"识别缴费记录数：{payload['source_record_count']}")
                self._write_log(f"生成明细行数：{payload['detail_record_count']}")
                self._write_log(f"识别人员数：{payload['employee_count']}")
                for account, count in payload["account_counts"].items():
                    self._write_log(f"- {account}：{count} 人")
                for period, count in payload["period_counts"].items():
                    self._write_log(f"- {period}：{count} 行")
                self._write_log(f"明细输出：{payload['detail_output_file']}")
                if payload.get("detail_output_files"):
                    self._write_log("按参保单位/参保地拆分明细：")
                    for output_file in payload["detail_output_files"]:
                        self._write_log(f"- {output_file}")
                self._write_log(f"汇总输出：{payload['summary_output_file']}")
                for warning in payload["warnings"]:
                    self._write_log(f"提醒：{warning}")
                message = "社保报表已生成完成，可以打开结果文件夹查看。"
            elif self.current_tool == "data_statistics":
                self._write_log("数据统计生成完成。")
                self._write_log(f"识别文件数：{payload['source_file_count']}")
                self._write_log(f"考勤原始记录数：{payload['attendance_source_count']}")
                self._write_log(f"考勤统计人数：{payload['attendance_person_count']}")
                self._write_log(f"考勤异常明细数：{payload['attendance_exception_count']}")
                self._write_log(f"周报记录数：{payload['weekly_record_count']}")
                self._write_log(f"月报记录数：{payload['monthly_record_count']}")
                if payload.get("report_staff_path"):
                    self._write_log(f"应汇报人员名单：{payload['report_staff_path']}")
                    self._write_log(f"应汇报人数：{payload['expected_reporter_count']}")
                self._write_log(f"周月报异常人数：{payload['report_person_count']}")
                self._write_log(f"周月报异常明细数：{payload['report_exception_count']}")
                self._write_log(f"输出：{payload['output_file']}")
                for warning in payload["warnings"]:
                    self._write_log(f"提醒：{warning}")
                message = "考勤周月报统计已生成完成，可以打开结果文件夹查看。"
            elif self.current_tool == "insurance_ledger":
                self._write_log("保险台账生成完成。")
                self._write_log(f"识别文件数：{payload['source_file_count']}")
                self._write_log(f"识别保单数：{payload['policy_count']}")
                self._write_log(f"保单人员数：{payload['insured_person_count']}")
                self._write_log(f"花名册在职人数：{payload['roster_person_count']}")
                self._write_log(f"需加保预警：{payload['add_warning_count']}")
                self._write_log(f"需减保预警：{payload['reduce_warning_count']}")
                self._write_log(f"输出：{payload['output_file']}")
                if payload.get("roster_warning_file"):
                    self._write_log(f"花名册预警输出：{payload['roster_warning_file']}")
                for warning in payload["warnings"]:
                    self._write_log(f"提醒：{warning}")
                message = "保险台账已生成完成，可以打开结果文件夹查看。"
            elif self.current_tool == "salary_merge":
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
                if payload.get("tool_name") == "需求6-花名册更新":
                    self._write_log("花名册更新完成。")
                    self._write_log(f"识别汇总表数：{payload['source_file_count']}")
                    self._write_log(f"识别异动记录数：{payload['record_count']}")
                    self._write_log(f"花名册新增：{payload['roster_added_count']} 人")
                    self._write_log(f"花名册标记离职：{payload['roster_marked_count']} 人")
                    for sheet_name, count in payload["sheet_counts"].items():
                        self._write_log(f"- {sheet_name}：{count} 条")
                    if payload.get("output_file"):
                        self._write_log(f"输出：{payload['output_file']}")
                    for warning in payload["warnings"]:
                        self._write_log(f"提醒：{warning}")
                    message = "花名册已更新完成，可以打开结果文件夹查看。"
                else:
                    self._write_log("汇总完成。")
                    self._write_log(f"识别文件数：{payload['source_file_count']}")
                    self._write_log(f"异动记录数：{payload['record_count']}")
                    self._write_log(f"写入模式：{'追加到已有汇总表' if payload.get('append_mode') else '新建干净汇总表'}")
                    self._write_log(f"新增记录数：{payload['inserted_count']}")
                    self._write_log(f"补充已有记录数：{payload['updated_count']}")
                    self._write_log(f"已存在未修改记录数：{payload['skipped_count']}")
                    for sheet_name, count in payload["sheet_counts"].items():
                        self._write_log(f"- {sheet_name}：{count} 条")
                    for period, counts in payload.get("period_counts", {}).items():
                        month_total = sum(counts.values())
                        self._write_log(f"- {period}：{month_total} 条")
                    if payload.get("output_files"):
                        for output_file in payload["output_files"]:
                            self._write_log(f"输出：{output_file}")
                    elif payload.get("output_file"):
                        self._write_log(f"输出：{payload['output_file']}")
                    if payload.get("roster_output_file"):
                        self._write_log(f"花名册输出：{payload['roster_output_file']}")
                        self._write_log(f"花名册新增：{payload['roster_added_count']} 人")
                        self._write_log(f"花名册标记离职：{payload['roster_marked_count']} 人")
                    for warning in payload["warnings"]:
                        self._write_log(f"提醒：{warning}")
                    message = "异动表已汇总完成，可以打开结果文件夹查看。"
            elif self.current_tool == "archive_import":
                if payload.get("tool_name") == "需求7-档案表生成":
                    self._write_log("档案表生成完成。")
                    self._write_log(f"识别公司数：{len(payload['company_counts'])}")
                    self._write_log(f"新建公司档案表数：{payload['created_count']}")
                    self._write_log(f"新增记录数：{payload['inserted_count']}")
                    self._write_log(f"补充已有记录数：{payload['updated_count']}")
                    self._write_log(f"已存在未修改记录数：{payload['skipped_count']}")
                    for company, count in payload["company_counts"].items():
                        self._write_log(f"- {company}：{count} 条")
                    for output_file in payload.get("output_files", []):
                        self._write_log(f"输出：{output_file}")
                    for warning in payload["warnings"]:
                        self._write_log(f"提醒：{warning}")
                    message = "档案表已生成完成，可以打开结果文件夹查看。"
                else:
                    self._write_log("入库完成。")
                    self._write_log("汇总表来源：{}".format(payload["target_path"] or "内置空模板"))
                    self._write_log(f"识别文件数：{payload['source_file_count']}")
                    self._write_log(f"识别记录数：{payload['source_record_count']}")
                    self._write_log(f"新增记录数：{payload['inserted_count']}")
                    self._write_log(f"补充已有记录数：{payload['updated_count']}")
                    self._write_log(f"已存在未修改记录数：{payload['skipped_count']}")
                    for company, count in payload["company_counts"].items():
                        self._write_log(f"- {company}：{count} 条")
                    self._write_log(f"输出：{payload['output_file']}")
                    for warning in payload["warnings"]:
                        self._write_log(f"提醒：{warning}")
                    message = "档案入库已完成，可以打开结果文件夹查看。"
            elif self.current_tool == "salary_split":
                self._write_log("拆分完成。")
                self._write_log(f"识别公司数：{payload['company_count']}")
                self._write_log(f"识别人员数：{payload['employee_count']}")
                for item in payload["outputs"]:
                    self._write_log(f"- {item['company']}：{item['employee_count']} 人")
                    if item.get("file_path"):
                        self._write_log(f"  输出：{item['file_path']}")
                message = "工资表已拆分完成，可以打开结果文件夹查看。"
            elif self.current_tool == "material_collector":
                self._write_log("员工资料检索与打包完成。")
                self._write_log(f"目标员工数：{payload['total_employees']} 人")
                self._write_log(f"材料齐全人数：{payload['complete_employee_count']} 人")
                self._write_log(f"提取文件总数：{payload['matched_file_count']} 个")
                if payload.get("report_path"):
                    self._write_log(f"汇总报告：{payload['report_path']}")
                if payload.get("zip_path"):
                    self._write_log(f"压缩包输出：{payload['zip_path']}")
                if payload.get("missing_records"):
                    self._write_log("存在缺件的员工：")
                    for emp, missing in payload["missing_records"].items():
                        self._write_log(f"- {emp} 缺少：{', '.join(missing)}")
                mismatches = [
                    f"- {m['employee_name']}（{m['material_type']}）：{m['mismatch_warning']}"
                    for m in payload.get("matches", [])
                    if m.get("mismatch_warning")
                ]
                if mismatches:
                    self._write_log("⚠️【信息核对预警】以下提取的资料与目标人员信息不一致：")
                    for warn_text in mismatches:
                        self._write_log(warn_text)
                for warning in payload.get("warnings", []):
                    self._write_log(f"提醒：{warning}")
                if mismatches:
                    message = "员工资料已打包完成，但存在【姓名/号码不一致】的预警，详情请查看汇总 Excel 或运行日志。"
                else:
                    message = "员工资料已检索打包完成，可以打开结果文件夹查看。"
            else:
                message = "处理完成。"
        self._update_project_output_controls()
        if hasattr(self, "workspace_tree") and self.current_project_path is not None:
            self._refresh_workspace_tree()
        if history_warning:
            self._write_log(f"提醒：结果已经生成，但没有完整保存到历史记录：{history_warning}")
            self._show_warning_after_log(
                "结果已生成，留存未完成",
                f"{message}\n\n但本次资料没有完整保存到“历史记录”。请不要删除原文件，并联系管理员检查保存空间。",
            )
        else:
            self._show_success_after_log("处理完成", message)

    def _handle_error(self, exc: object | None) -> None:
        action = (
            "生成"
            if self.current_tool == "social_security"
            else "生成"
            if self.current_tool == "data_statistics"
            else "生成"
            if self.current_tool == "insurance_ledger"
            else
            "合并"
            if self.current_tool == "salary_merge"
            else "更新"
            if self.current_tool == "personnel_change_merge" and self.change_mode == "roster"
            else "汇总"
            if self.current_tool == "personnel_change_merge"
            else "入库"
            if self.current_tool == "archive_import" and self.archive_mode == "import"
            else "生成"
            if self.current_tool == "archive_import"
            else "改名"
            if self.current_tool == "folder_rename"
            else "拆分"
        )
        self._write_log(f"{action}失败。")
        self._write_log(str(exc))
        self._show_error_after_log(f"{action}失败", str(exc))

    def _show_success_after_log(self, title: str, message: str) -> None:
        self._flush_log_view()
        self.root.after(80, lambda: messagebox.showinfo(title, message, parent=self.root))

    def _show_error_after_log(self, title: str, message: str) -> None:
        self._flush_log_view()
        self.root.after(80, lambda: messagebox.showerror(title, message, parent=self.root))

    def _show_warning_after_log(self, title: str, message: str) -> None:
        self._flush_log_view()
        self.root.after(80, lambda: messagebox.showwarning(title, message, parent=self.root))

    def _write_folder_rename_preview(self, result) -> None:
        payload = result.to_dict()
        self._write_log(f"目录：{payload['root_dir']}")
        self._write_log(f"数量：{payload['operation_count']}")
        preview_limit = payload["operation_count"] if payload["mode"] == MODE_EXCEL_BATCH else 30
        for operation in payload["operations"][:preview_limit]:
            self._write_log(f"- {operation['source_name']} -> {operation['target_name']}")
        remaining = payload["operation_count"] - preview_limit
        if remaining > 0:
            self._write_log(f"... 还有 {remaining} 条")
        for warning in payload["warnings"]:
            self._write_log(f"提醒：{warning}")

    def _folder_rename_confirm_message(self, result) -> str:
        payload = result.to_dict()
        lines = [f"确认改名 {payload['operation_count']} 个项目："]
        for operation in payload["operations"][:8]:
            lines.append(f"{operation['source_name']} -> {operation['target_name']}")
        remaining = payload["operation_count"] - 8
        if remaining > 0:
            lines.append(f"... 还有 {remaining} 条")
        if payload["warnings"]:
            lines.append("")
            lines.append("预览提醒：")
            for warning in payload["warnings"][:6]:
                lines.append(f"- {warning}")
            warning_remaining = len(payload["warnings"]) - 6
            if warning_remaining > 0:
                lines.append(f"- ... 还有 {warning_remaining} 条提醒，请查看运行记录")
        lines.append("")
        lines.append("确认后会复制到当前项目，并在处理结果副本上改名；原文件夹不会改变。是否继续？")
        return "\n".join(lines)

    def _open_output_dir(self) -> None:
        directory = self.last_output_dir
        if directory is None:
            messagebox.showwarning("暂无处理结果", "请先完成一次处理，再打开结果目录。")
            return
        if not directory.exists():
            messagebox.showwarning("结果目录不存在", "本次结果目录可能已被移动，请在右侧“项目文件”中查找。")
            return
        open_path(directory)

    def _write_log(self, text: str) -> None:
        if not getattr(self, "_is_alive", True):
            return
        if not hasattr(self, "log_text"):
            return
        try:
            if not self.log_text.winfo_exists():
                return
        except Exception:
            return
        # 时间线式日志：彩色圆点 + 时间戳 + 内容（对应设计稿“运行记录”）
        tag = None
        if any(keyword in text for keyword in ("失败", "错误")):
            tag = "error"
        elif any(keyword in text for keyword in ("提醒", "不存在", "缺少")):
            tag = "warning"
        elif any(keyword in text for keyword in ("完成", "成功")):
            tag = "success"
        elif text.startswith(("- ", "  ", "（")):
            tag = "muted"
        dot_tag = {
            "error": "dot_error",
            "warning": "dot_warning",
            "success": "dot_success",
            "muted": "dot_muted",
        }.get(tag or "", "dot_success")
        timestamp = datetime.now().strftime("%H:%M:%S")
        try:
            if tag == "muted":
                self.log_text.insert(END, "   ", "muted")
            else:
                self.log_text.insert(END, "● ", dot_tag)
                self.log_text.insert(END, f"{timestamp}  ", "timestamp")
            if tag:
                self.log_text.insert(END, text + "\n", tag)
            else:
                self.log_text.insert(END, text + "\n")
            self.log_text.see(END)
        except Exception:
            pass

    def _flush_log_view(self) -> None:
        if not getattr(self, "_is_alive", True):
            return
        try:
            if hasattr(self, "log_text") and self.log_text.winfo_exists():
                self.log_text.see(END)
                self.log_text.update_idletasks()
            if hasattr(self, "root") and self.root.winfo_exists():
                self.root.update_idletasks()
        except Exception:
            pass

    def _clear_log(self) -> None:
        if not getattr(self, "_is_alive", True):
            return
        try:
            if hasattr(self, "log_text") and self.log_text.winfo_exists():
                self.log_text.delete("1.0", END)
        except Exception:
            pass

    def destroy(self) -> None:
        self._is_alive = False
        if getattr(self, "_workspace_search_job", None) is not None:
            try:
                self.root.after_cancel(self._workspace_search_job)
            except Exception:
                pass
            self._workspace_search_job = None
        if hasattr(self, "_startup_loading_timer") and self._startup_loading_timer:
            try:
                self.root.after_cancel(self._startup_loading_timer)
            except Exception:
                pass
            self._startup_loading_timer = None
        self._dismiss_startup_loading_screen()
        if hasattr(self, "_startup_check_timer") and self._startup_check_timer:
            try:
                self.root.after_cancel(self._startup_check_timer)
            except Exception:
                pass
            self._startup_check_timer = None
        if hasattr(self, "_poll_update_timer") and self._poll_update_timer:
            try:
                self.root.after_cancel(self._poll_update_timer)
            except Exception:
                pass
            self._poll_update_timer = None




def main() -> None:
    _install_crash_logging()
    _set_windows_app_identity()
    _enable_high_dpi_rendering()
    root = Tk()
    root.withdraw()
    HRToolkitApp(root)
    root.mainloop()

def __getattr__(name: str):
    import hr_toolkit.gui as _pkg
    try:
        return getattr(_pkg, name)
    except AttributeError:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
