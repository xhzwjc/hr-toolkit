"""UI constants, color palettes, and dimensions for HR Toolkit."""

from __future__ import annotations

from hr_toolkit.tools.folder_rename import (
    MODE_APPEND,
    MODE_EXCEL_BATCH,
    MODE_REMOVE,
    MODE_REPLACE,
    FILE_TYPE_FOLDER,
    FILE_TYPE_ALL,
    FILE_TYPE_PDF,
    FILE_TYPE_IMAGE,
    FILE_TYPE_DOCUMENT,
)
from hr_toolkit.desktop_contract import NAV_GROUPS, TOOL_NAV_ITEMS

RENAME_MODE_LABELS = {
    "追加文字": MODE_APPEND,
    "删除结尾文字": MODE_REMOVE,
    "修改单人名称": MODE_REPLACE,
    "按 Excel 人名顺序批量重命名": MODE_EXCEL_BATCH,
}

RENAME_FILE_TYPE_LABELS = {
    "文件夹": FILE_TYPE_FOLDER,
    "PDF": FILE_TYPE_PDF,
    "图片（jpg/png/gif等）": FILE_TYPE_IMAGE,
    "文档（doc/xls/ppt/txt等）": FILE_TYPE_DOCUMENT,
    "全部": FILE_TYPE_ALL,
}
RENAME_FILE_TYPE_LABELS_REVERSE = {v: k for k, v in RENAME_FILE_TYPE_LABELS.items()}

TOOL_NAV_LABELS = dict(TOOL_NAV_ITEMS)

TOOL_LOG_LABELS = {
    "social_security": "社保汇总",
    "data_statistics": "数据统计",
    "insurance_ledger": "保险台账",
    "salary_split": "工资拆分",
    "salary_merge": "工资合并",
    "personnel_change_merge": "异动汇总",
    "archive_import": "档案入库",
    "material_collector": "资料打包",
    "folder_rename": "文件夹改名",
}

TOOL_GROUP_LABELS = {tool_id: group for group, tools in NAV_GROUPS for tool_id in tools}

MULTI_INPUT_TOOLS = {
    "social_security",
    "data_statistics",
    "insurance_ledger",
    "salary_merge",
    "personnel_change_merge",
    "archive_import",
}

HISTORY_STATUS_LABELS = {
    "running": "处理中",
    "success": "已完成",
    "failed": "处理失败",
    "stopped": "未完成",
}

HISTORY_TOOL_FILTER_ALL = "全部功能"
HISTORY_DATE_FILTER_ALL = "全部时间"
HISTORY_DATE_FILTERS = (HISTORY_DATE_FILTER_ALL, "今天", "最近7天", "最近30天", "今年")

HISTORY_PRIMARY_PATH_ARGUMENTS = {"input_path", "input_dir", "summary_input", "summary_path"}
HISTORY_SUPPORTING_PATH_ARGUMENTS = {
    "roster_path",
    "roster_source",
    "report_staff_path",
    "existing_summary_path",
    "template_path",
    "analysis_template_path",
    "excel_path",
    "target_path",
    "existing_archive_path",
}
HISTORY_PATH_ARGUMENTS = HISTORY_PRIMARY_PATH_ARGUMENTS | HISTORY_SUPPORTING_PATH_ARGUMENTS

WORKSPACE_DEFAULT_WIDTH = 320
WORKSPACE_MIN_WIDTH = 270
WORKSPACE_MAX_WIDTH = 430
WORKSPACE_COLLAPSED_WIDTH = 46
WORKSPACE_DRAWER_BREAKPOINT = 980
WORKSPACE_SEARCH_LIMIT = 500
WORKSPACE_DUMMY_TAG = "__workspace_dummy__"
WORKSPACE_SCOPE_ALL = "all"
WORKSPACE_SCOPE_TOOL = "tool"

WORKSPACE_TOOL_PATHS = {
    "social_security": ("社保与保险", "社保明细与汇总"),
    "insurance_ledger": ("社保与保险", "保险台账与预警"),
    "data_statistics": ("考勤与统计", "考勤与周月报"),
    "salary_split": ("薪酬管理", "工资表拆分"),
    "salary_merge": ("薪酬管理", "多月工资合并"),
    "personnel_change_merge": ("人员与档案", "异动汇总"),
    "archive_import": ("人员与档案", "档案入库"),
    "material_collector": ("人员与档案", "员工资料打包"),
    "folder_rename": ("人员与档案", "资料文件夹改名"),
}

WORKSPACE_HIDDEN_NAMES = {
    ".hrtoolkit",
    ".DS_Store",
    "Thumbs.db",
    "desktop.ini",
}
WORKSPACE_HIDDEN_SUFFIXES = (".partial", ".tmp", ".temp", ".lock")

# Color Palette
COLOR_BG = "#F7F5F1"
COLOR_SIDEBAR = "#F7F5F1"
COLOR_SIDEBAR_BORDER = "#EBE9E4"
COLOR_SURFACE = "#ffffff"
COLOR_SURFACE_ALT = "#FAF9F6"
COLOR_SURFACE_PRESSED = "#F2F0EA"
COLOR_BORDER = "#ECEAE4"
COLOR_BORDER_FAINT = "#F1EFE9"
COLOR_TEXT = "#292825"
COLOR_MUTED = "#78766E"
COLOR_FAINT = "#98958C"
COLOR_DISABLED = "#B3B0A6"
COLOR_PRIMARY = "#17715B"
COLOR_PRIMARY_ACTIVE = "#125E4B"
COLOR_PRIMARY_SOFT = "#E4EFEA"
COLOR_NAV_SELECTED = "#EBE8E1"
COLOR_NAV_HOVER = "#F0EEE8"
COLOR_NAV_TEXT = "#55534C"
COLOR_NAV_TEXT_SELECTED = "#17715B"
COLOR_SUCCESS = "#1F7A52"
COLOR_SUCCESS_DOT = "#2E9E6B"
COLOR_WARNING = "#A05E12"
COLOR_WARNING_DOT = "#D9A441"
COLOR_WARNING_SOFT = "#F8EBD2"
COLOR_DANGER = "#B0352B"
COLOR_LOG_BG = "#ffffff"
COLOR_LOG_TEXT = "#292825"
COLOR_LOG_MUTED = "#98958C"
COLOR_DROP_BORDER = "#D8D5CB"
COLOR_DROP_BG = "#FBFAF7"
COLOR_BADGE_ZIP_BG = "#F6E8D4"
COLOR_BADGE_ZIP_FG = "#A05E12"
COLOR_BADGE_XLS_BG = "#DFEFE7"
COLOR_BADGE_XLS_FG = "#1F7A52"
COLOR_BADGE_DIR_BG = "#EBE8E1"
COLOR_BADGE_DIR_FG = "#78766E"

APP_DISPLAY_NAME = "HR Workbench"
APP_SUBTITLE = "人员运营自动化"

UPDATE_DIALOG_BG = COLOR_SURFACE
UPDATE_DIALOG_TEXT = COLOR_TEXT
UPDATE_DIALOG_MUTED = COLOR_MUTED
UPDATE_DIALOG_TRACK = "#EFEDE7"
UPDATE_DIALOG_PRIMARY = COLOR_PRIMARY
UPDATE_DIALOG_PRIMARY_ACTIVE = COLOR_PRIMARY_ACTIVE
UPDATE_DIALOG_SECONDARY = "#F2F0EA"
UPDATE_DIALOG_SECONDARY_ACTIVE = "#EBE8E1"
UPDATE_DIALOG_ICON_BG = COLOR_PRIMARY_SOFT
UPDATE_DIALOG_NOTES_BG = "#FAF9F6"

BASE_WINDOWS_DPI = 96
TK_POINTS_PER_INCH = 72
FORCE_UI_SCALE_ENV = "HR_TOOLKIT_FORCE_UI_SCALE"
