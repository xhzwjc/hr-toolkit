from __future__ import annotations

import calendar
import os
import queue
import subprocess
import sys
import threading
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, VERTICAL, Y, Canvas, Frame, Label, Menu, PhotoImage, Toplevel, filedialog, messagebox
from tkinter import Tk, StringVar, Text
from tkinter import font as tkfont
from tkinter import ttk

from hr_toolkit import __version__, runlog
from hr_toolkit.app_update import (
    UpdateInfo,
    check_for_update,
    cleanup_stale_update_files,
    download_update_package,
    launch_update_replacement,
    resolve_download_url,
    update_check_enabled,
)
from hr_toolkit.tools.folder_rename import (
    MODE_APPEND,
    MODE_REMOVE,
    MODE_REPLACE,
    rename_person_folders,
    FILE_TYPE_FOLDER,
    FILE_TYPE_ALL,
    FILE_TYPE_PDF,
    FILE_TYPE_IMAGE,
    FILE_TYPE_DOCUMENT,
)
from hr_toolkit.tools.archive_import import export_company_archive_tables, import_archive_transfers
from hr_toolkit.tools.data_statistics import generate_data_statistics_reports, resolve_week_range
from hr_toolkit.tools.personnel_change_merge import merge_personnel_changes, update_roster_from_change_summaries
from hr_toolkit.tools.insurance_ledger import generate_insurance_ledger
from hr_toolkit.tools.salary_merge import merge_monthly_salary
from hr_toolkit.tools.salary_split import split_salary_by_company
from hr_toolkit.tools.social_security import generate_social_security_reports


RENAME_MODE_LABELS = {
    "追加文字": MODE_APPEND,
    "删除结尾文字": MODE_REMOVE,
    "修改单人名称": MODE_REPLACE,
}

RENAME_FILE_TYPE_LABELS = {
    "文件夹": FILE_TYPE_FOLDER,
    "PDF": FILE_TYPE_PDF,
    "图片（jpg/png/gif等）": FILE_TYPE_IMAGE,
    "文档（doc/xls/ppt/txt等）": FILE_TYPE_DOCUMENT,
    "全部": FILE_TYPE_ALL,
}
RENAME_FILE_TYPE_LABELS_REVERSE = {v: k for k, v in RENAME_FILE_TYPE_LABELS.items()}

TOOL_NAV_ITEMS = (
    ("social_security", "社保明细与汇总"),
    ("insurance_ledger", "保险台账与预警"),
    ("data_statistics", "考勤与周月报"),
    ("salary_split", "工资表拆分"),
    ("salary_merge", "多月工资合并"),
    ("personnel_change_merge", "异动汇总"),
    ("archive_import", "档案入库"),
    ("folder_rename", "资料文件夹改名"),
)
TOOL_NAV_LABELS = dict(TOOL_NAV_ITEMS)
# 运行日志沿用旧的简短名称，保证历史日志可以按同一关键字检索
TOOL_LOG_LABELS = {
    "social_security": "社保汇总",
    "data_statistics": "数据统计",
    "insurance_ledger": "保险台账",
    "salary_split": "工资拆分",
    "salary_merge": "工资合并",
    "personnel_change_merge": "异动汇总",
    "archive_import": "档案入库",
    "folder_rename": "文件夹改名",
}
NAV_GROUPS = (
    ("社保与保险", ("social_security", "insurance_ledger")),
    ("考勤与统计", ("data_statistics",)),
    ("薪酬管理", ("salary_split", "salary_merge")),
    ("人员与档案", ("personnel_change_merge", "archive_import", "folder_rename")),
)
TOOL_GROUP_LABELS = {tool_id: group for group, tools in NAV_GROUPS for tool_id in tools}
# 支持一次选择多个文件/压缩包/文件夹作为输入的工具
MULTI_INPUT_TOOLS = {
    "social_security",
    "data_statistics",
    "insurance_ledger",
    "salary_merge",
    "personnel_change_merge",
    "archive_import",
}

# “从容舒适 · Notion / macOS 风”暖纸色系：暖纸底色、深青主色、
# 白色卡片分区、时间线式日志（对应设计稿方案 1b）。
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


def _scale_px(value: int | float, scale: float) -> int:
    if value == 0:
        return 0
    scaled = int(round(value * scale))
    if value > 0:
        return max(1, scaled)
    return min(-1, scaled)


def _scale_float(value: int | float, scale: float) -> float:
    return float(value) * scale


def _clamp_ui_scale(scale: float) -> float:
    return max(1.0, min(scale, 3.0))


def _indeterminate_progress_segment(
    track_width: float,
    sweep_head: float,
    segment_width: float,
) -> tuple[float, float] | None:
    """Return the visible part of a left-to-right indeterminate sweep."""
    if track_width <= 0 or segment_width <= 0:
        return None
    visible_start = max(0.0, sweep_head - segment_width)
    visible_end = min(track_width, sweep_head)
    if visible_end <= visible_start:
        return None
    return visible_start, visible_end


def _forced_ui_scale() -> float | None:
    raw_value = os.environ.get(FORCE_UI_SCALE_ENV, "").strip()
    if not raw_value:
        return None
    try:
        return _clamp_ui_scale(float(raw_value))
    except ValueError:
        return None


def _windows_dpi_for_root(root: Tk) -> float | None:
    if not sys.platform.startswith("win"):
        return None
    try:
        import ctypes
    except Exception:
        return None

    try:
        hwnd = int(root.winfo_id())
        get_dpi_for_window = getattr(ctypes.windll.user32, "GetDpiForWindow", None)
        if get_dpi_for_window is not None and hwnd:
            dpi = int(get_dpi_for_window(hwnd))
            if dpi > 0:
                return float(dpi)
    except Exception:
        pass

    try:
        get_dpi_for_system = getattr(ctypes.windll.user32, "GetDpiForSystem", None)
        if get_dpi_for_system is not None:
            dpi = int(get_dpi_for_system())
            if dpi > 0:
                return float(dpi)
    except Exception:
        pass

    hdc = None
    try:
        hdc = ctypes.windll.user32.GetDC(None)
        if hdc:
            dpi = int(ctypes.windll.gdi32.GetDeviceCaps(hdc, 88))
            if dpi > 0:
                return float(dpi)
    except Exception:
        pass
    finally:
        if hdc:
            try:
                ctypes.windll.user32.ReleaseDC(None, hdc)
            except Exception:
                pass
    return None


def _detect_ui_scale(root: Tk) -> float:
    forced = _forced_ui_scale()
    if forced is not None:
        return forced
    if not sys.platform.startswith("win"):
        return 1.0
    dpi = _windows_dpi_for_root(root)
    if dpi is None:
        try:
            dpi = float(root.winfo_fpixels("1i"))
        except Exception:
            dpi = BASE_WINDOWS_DPI
    return _clamp_ui_scale(dpi / BASE_WINDOWS_DPI)


def _configure_tk_font_scaling(root: Tk, ui_scale: float) -> None:
    try:
        root.tk.call("tk", "scaling", (BASE_WINDOWS_DPI * ui_scale) / TK_POINTS_PER_INCH)
    except Exception:
        pass


def _font_size(size: int) -> int:
    """把设计字号（Windows 96dpi 基准）换算成当前平台的 Tk 字号。

    macOS 的 aqua 后端把“点”直接按像素渲染（72dpi 假设），且完全忽略
    tk scaling 设置，同样的数值在 Mac 上只有 Windows 上的 3/4 大，
    造成“布局一致、字明显偏小”。这里放大 4/3 对齐设计基准；
    Windows/Linux 走 tk scaling 机制，原值返回。
    """
    if sys.platform == "darwin":
        return max(1, round(size * 4 / 3))
    return size


def _widget_ui_scale(widget) -> float:
    try:
        return float(getattr(widget.winfo_toplevel(), "_hr_ui_scale", 1.0))
    except Exception:
        return 1.0


class CodexButton(Canvas):
    def __init__(
        self,
        master,
        *,
        text: str = "",
        command=None,
        textvariable: StringVar | None = None,
        icon: str = "",
        variant: str = "secondary",
        width: int | None = None,
        height: int = 34,
        min_width: int = 92,
    ) -> None:
        self._scale = _widget_ui_scale(master)
        self._text = text
        self._command = command
        self._textvariable = textvariable
        self._icon = icon
        self._variant = variant
        self._state = "normal"
        self._hover = False
        self._height = self._px(height)
        self._min_width = self._px(min_width)
        self._variable_trace: str | None = None
        display_text = self._display_text()
        initial_width = self._px(width) if width is not None else self._measure_width(display_text, icon, self._min_width)
        self._canvas_bg = self._resolve_parent_bg(master)
        super().__init__(
            master,
            width=initial_width,
            height=self._height,
            bg=self._canvas_bg,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        if textvariable is not None:
            self._variable_trace = textvariable.trace_add("write", lambda *_args: self._redraw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Configure>", lambda _event: self._redraw())
        self._redraw()

    def configure(self, cnf=None, **kwargs):  # type: ignore[override]
        if cnf:
            kwargs.update(cnf)
        if "text" in kwargs:
            self._text = kwargs.pop("text")
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
            super().configure(cursor="" if self._state == "disabled" else "hand2")
        if "textvariable" in kwargs:
            self._textvariable = kwargs.pop("textvariable")
        if "icon" in kwargs:
            self._icon = kwargs.pop("icon")
        if "variant" in kwargs:
            self._variant = kwargs.pop("variant")
        if kwargs:
            super().configure(**kwargs)
        self._redraw()

    config = configure

    @staticmethod
    def _resolve_parent_bg(master) -> str:
        try:
            background = master.cget("background")
            if background:
                return background
        except Exception:
            pass
        try:
            style_name = master.cget("style")
            if style_name:
                background = ttk.Style(master).lookup(style_name, "background")
                if background:
                    return background
        except Exception:
            pass
        return COLOR_BG

    def _px(self, value: int | float) -> int:
        return _scale_px(value, self._scale)

    def _pxf(self, value: int | float) -> float:
        return _scale_float(value, self._scale)

    def _display_text(self) -> str:
        if self._textvariable is not None:
            return self._textvariable.get()
        return self._text

    def _measure_width(self, text: str, icon: str, min_width: int) -> int:
        text_units = sum(2 if ord(char) > 127 else 1 for char in text)
        width = self._px(28) + text_units * self._px(7)
        if icon:
            width += self._px(20)
        return max(min_width, width)

    def _palette(self) -> tuple[str, str, str, str]:
        if self._state == "disabled":
            return COLOR_SURFACE_PRESSED, COLOR_SURFACE_PRESSED, COLOR_DISABLED, COLOR_BORDER
        if self._variant == "primary":
            return COLOR_PRIMARY, COLOR_PRIMARY_ACTIVE, "#ffffff", COLOR_PRIMARY
        if self._variant == "tonal":
            return COLOR_SURFACE_PRESSED, COLOR_NAV_SELECTED, COLOR_NAV_TEXT, COLOR_SURFACE_PRESSED
        return COLOR_SURFACE, COLOR_SURFACE_ALT, COLOR_NAV_TEXT, COLOR_BORDER

    def _on_enter(self, _event=None) -> None:
        self._hover = True
        self._redraw()

    def _on_leave(self, _event=None) -> None:
        self._hover = False
        self._redraw()

    def _on_click(self, _event=None) -> None:
        if self._state == "disabled" or self._command is None:
            return
        self._command()

    def _redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), int(float(self.cget("width"))))
        height = max(self.winfo_height(), self._height)
        text = self._display_text()
        family = self.master.winfo_toplevel().tk.call("font", "actual", "TkDefaultFont", "-family")
        if self._variant == "link":
            # 文字链接样式：无底色无边框，仅深青文字，悬停加深
            if self._state == "disabled":
                foreground = COLOR_DISABLED
            else:
                foreground = COLOR_PRIMARY_ACTIVE if self._hover else COLOR_PRIMARY
            font = (family, _font_size(10))
            content = f"{self._icon} {text}".strip() if self._icon else text
            self.create_text(width / 2, height / 2, text=content, fill=foreground, font=font)
            return
        normal, active, foreground, border = self._palette()
        fill = active if self._hover and self._state != "disabled" else normal
        inset = self._pxf(1)
        self._draw_round_rect(inset, inset, width - inset, height - inset, self._pxf(9), fill=fill, outline=border, width=self._pxf(1))
        font = (family, _font_size(10), "bold") if self._variant == "primary" else (family, _font_size(10))
        if self._icon:
            content_width = self._measure_width(text, self._icon, 0) - self._px(28)
            start_x = max((width - content_width) / 2, self._pxf(12))
            self.create_text(start_x + self._pxf(7), height / 2, text=self._icon, fill=foreground, font=font, anchor="center")
            self.create_text(start_x + self._pxf(22), height / 2, text=text, fill=foreground, font=font, anchor="w")
        else:
            self.create_text(width / 2, height / 2, text=text, fill=foreground, font=font)

    def _draw_round_rect(self, x1: float, y1: float, x2: float, y2: float, radius: float, **kwargs) -> None:
        radius = max(0, min(radius, (x2 - x1) / 2, (y2 - y1) / 2))
        self.create_polygon(
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1,
            smooth=True,
            splinesteps=16,
            **kwargs,
        )


def _paint_tool_icon(canvas: Canvas, icon_id: str, color: str, x: float, y: float, size: float, line_width: float) -> None:
    """在 size×size 的方框内绘制工具线性图标（坐标按设计稿 14×14 视图换算）。"""

    def px(value: float) -> float:
        return x + value * size / 14.0

    def py(value: float) -> float:
        return y + value * size / 14.0

    line = {"fill": color, "width": line_width, "capstyle": "round", "joinstyle": "round"}
    if icon_id == "social_security":
        canvas.create_rectangle(px(2.5), py(1.5), px(11.5), py(12.5), outline=color, width=line_width)
        canvas.create_line(px(5), py(5), px(9), py(5), **line)
        canvas.create_line(px(5), py(8), px(9), py(8), **line)
    elif icon_id == "insurance_ledger":
        canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width)
        canvas.create_line(px(4.6), py(7), px(6.3), py(8.7), px(9.6), py(5.4), **line)
    elif icon_id == "data_statistics":
        canvas.create_line(px(2.5), py(12), px(2.5), py(7), **line)
        canvas.create_line(px(7), py(12), px(7), py(2.5), **line)
        canvas.create_line(px(11.5), py(12), px(11.5), py(5), **line)
    elif icon_id == "salary_split":
        canvas.create_line(px(7), py(2), px(7), py(6), **line)
        canvas.create_line(px(7), py(6), px(3), py(11), **line)
        canvas.create_line(px(7), py(6), px(11), py(11), **line)
    elif icon_id == "salary_merge":
        canvas.create_line(px(3), py(3), px(7), py(8), **line)
        canvas.create_line(px(11), py(3), px(7), py(8), **line)
        canvas.create_line(px(7), py(8), px(7), py(12), **line)
    elif icon_id == "personnel_change_merge":
        canvas.create_line(px(2), py(4.5), px(10), py(4.5), **line)
        canvas.create_line(px(8), py(2), px(10.5), py(4.5), px(8), py(7), **line)
        canvas.create_line(px(12), py(9.5), px(4), py(9.5), **line)
        canvas.create_line(px(6), py(7), px(3.5), py(9.5), px(6), py(12), **line)
    elif icon_id == "archive_import":
        canvas.create_rectangle(px(2), py(4.5), px(12), py(12), outline=color, width=line_width)
        canvas.create_line(px(2), py(7), px(12), py(7), **line)
        canvas.create_line(px(7), py(4.5), px(7), py(2.5), **line)
    elif icon_id == "folder_rename":
        canvas.create_line(
            px(2), py(10.5), px(2), py(4), px(3.5), py(2.5), px(6), py(2.5), px(7.2), py(4),
            px(10.5), py(4), px(12), py(5.5), px(12), py(10.5), px(10.5), py(12), px(3.5), py(12), px(2), py(10.5),
            **line,
        )
    elif icon_id == "tutorial":
        canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width)
        canvas.create_line(px(7), py(6.5), px(7), py(10), **line)
        canvas.create_line(px(7), py(4), px(7), py(4.45), **line)
    elif icon_id == "clock":
        canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width)
        canvas.create_line(px(7), py(4), px(7), py(7.2), px(9), py(8.6), **line)
    else:
        canvas.create_oval(px(3), py(3), px(11), py(11), outline=color, width=line_width)


class SidebarItem(Canvas):
    """侧边栏导航条目：圆角底 + 线性图标 + 文字（对应设计稿导航行）。"""

    def __init__(
        self,
        master,
        *,
        text: str,
        icon_id: str,
        command=None,
        height: int = 32,
        muted: bool = False,
    ) -> None:
        self._scale = _widget_ui_scale(master)
        self._text = text
        self._icon_id = icon_id
        self._command = command
        self._muted = muted
        self._selected = False
        self._hover = False
        super().__init__(
            master,
            height=_scale_px(height, self._scale),
            bg=COLOR_SIDEBAR,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Configure>", lambda _event: self._redraw())
        self._redraw()

    def _px(self, value: int | float) -> int:
        return _scale_px(value, self._scale)

    def _pxf(self, value: int | float) -> float:
        return _scale_float(value, self._scale)

    def set_selected(self, selected: bool) -> None:
        if self._selected != selected:
            self._selected = selected
            self._redraw()

    def _on_enter(self, _event=None) -> None:
        self._hover = True
        self._redraw()

    def _on_leave(self, _event=None) -> None:
        self._hover = False
        self._redraw()

    def _on_click(self, _event=None) -> None:
        if self._command is not None:
            self._command()

    def _redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), 1)
        height = max(self.winfo_height(), 1)
        if self._selected:
            fill = COLOR_NAV_SELECTED
        elif self._hover:
            fill = COLOR_NAV_HOVER
        else:
            fill = COLOR_SIDEBAR
        if fill != COLOR_SIDEBAR:
            radius = self._pxf(7)
            CodexButton._draw_round_rect(self, 0, 0, width, height, radius, fill=fill, outline="")
        if self._selected:
            foreground = COLOR_NAV_TEXT_SELECTED
        elif self._muted:
            foreground = COLOR_MUTED
        else:
            foreground = COLOR_NAV_TEXT
        icon_size = self._pxf(15)
        icon_x = self._pxf(9)
        icon_y = (height - icon_size) / 2
        _paint_tool_icon(self, self._icon_id, foreground, icon_x, icon_y, icon_size, max(1.0, self._pxf(1.4)))
        family = self.winfo_toplevel().tk.call("font", "actual", "TkDefaultFont", "-family")
        font = (family, _font_size(10), "bold") if self._selected else (family, _font_size(10))
        self.create_text(icon_x + icon_size + self._pxf(9), height / 2, text=self._text, fill=foreground, font=font, anchor="w")


class RoundedCard(Canvas):
    """白色圆角卡片容器（设计稿 border-radius:14px 的卡片分区）。

    Canvas 负责画圆角底和淡投影，内容放进 ``inner``（ttk.Frame）；
    fill_height=False 时卡片高度跟随内容，True 时内容撑满分配到的高度。
    """

    def __init__(
        self,
        master,
        *,
        padding: tuple[int, int, int, int] = (20, 16, 20, 18),
        radius: int = 14,
        fill_height: bool = False,
        min_height: int = 0,
    ) -> None:
        self._scale = _widget_ui_scale(master)
        self._radius = _scale_float(radius, self._scale)
        self._fill_height = fill_height
        page_bg = CodexButton._resolve_parent_bg(master)
        super().__init__(
            master,
            bg=page_bg,
            highlightthickness=0,
            bd=0,
            height=_scale_px(min_height, self._scale) if min_height else 1,
        )
        self.inner = ttk.Frame(self, style="InputWrap.TFrame")
        self._pads = (0, 0, 0, 0)
        self._window = self.create_window(0, 0, window=self.inner, anchor="nw")
        self._last_bg_size = (0, 0)
        self.set_padding(padding, sync=False)
        self.inner.bind("<Configure>", self._sync)
        self.bind("<Configure>", self._sync)

    def set_padding(self, padding: tuple[int, int, int, int], *, sync: bool = True) -> None:
        # 内容窗口向内缩进，让方角的内容框始终落在圆角轮廓之内
        self._pads = tuple(_scale_px(value, self._scale) for value in padding)
        self.coords(self._window, self._pads[0], self._pads[1])
        if sync:
            self._sync()

    def _sync(self, _event=None) -> None:
        left, top, right, bottom = self._pads
        width = max(self.winfo_width(), 1)
        inner_width = max(width - left - right, 1)
        if self._fill_height:
            height = max(self.winfo_height(), 1)
            self.itemconfigure(self._window, width=inner_width, height=max(height - top - bottom, 1))
        else:
            self.itemconfigure(self._window, width=inner_width)
            height = self.inner.winfo_reqheight() + top + bottom
            if int(float(self.cget("height"))) != height:
                self.configure(height=height)
        if (width, height) != self._last_bg_size:
            self._last_bg_size = (width, height)
            self._redraw_bg(width, height)

    def _redraw_bg(self, width: int, height: int) -> None:
        self.delete("card_bg")
        offset = max(1.0, _scale_float(1.5, self._scale))
        # 无边框卡片：底下垫一层淡色圆角模拟设计稿的轻投影
        CodexButton._draw_round_rect(
            self, offset, offset * 1.6, width - offset * 0.4, height, self._radius, fill="#ECE9E2", outline="", tags="card_bg"
        )
        CodexButton._draw_round_rect(
            self, 0, 0, width - offset, height - offset, self._radius, fill=COLOR_SURFACE, outline="", tags="card_bg"
        )
        self.tag_lower("card_bg")


class HRToolkitApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.ui_scale = _detect_ui_scale(root)
        setattr(self.root, "_hr_ui_scale", self.ui_scale)
        # Windows 依赖 tk scaling 把“点”字号按 DPI 转像素；
        # macOS 的 aqua 后端忽略 tk scaling，字号统一走 _font_size() 换算
        if sys.platform.startswith("win") or _forced_ui_scale() is not None:
            _configure_tk_font_scaling(self.root, self.ui_scale)

        self.root.title(f"{APP_DISPLAY_NAME} v{__version__}")
        initial_width, initial_height = self._window_size(1180, 760)
        min_width, min_height = self._window_size(1020, 680)
        self.root.geometry(f"{initial_width}x{initial_height}")
        self.root.minsize(min_width, min_height)
        self.root.configure(bg=COLOR_BG)

        self.current_tool = "social_security"
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
        # 每个工具（含子模式）本次会话内的最近一次运行结果：(时间, 成功/失败)
        self._last_run_results: dict[str, tuple[str, str]] = {}
        # 合并后的上传入口：当前工具可用的文件/文件夹选择动作与提示文案
        self._input_file_cmd = None
        self._input_folder_cmd = None
        self._input_drop_title = ""
        self._input_allow_multi = True
        self._tutorial_window: Toplevel | None = None
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
        self.input_path = StringVar()
        self.summary_path = StringVar()
        self.stats_week_start = StringVar()
        self.stats_week_end = StringVar()
        # 考勤统计表备注中加班/调休的展示单位：day 按天（默认）/ hour 按小时
        self.stats_remark_unit = StringVar(value="day")
        self.output_dir = StringVar(value=str(default_output_parent_dir(self.current_tool)))
        self.output_dir_user_selected = False
        self.change_input_paths: list[Path] | None = None
        # (状态, 运行编号, 载荷)；运行编号用于丢弃已停止任务的结果
        self.status_queue: queue.Queue[tuple[str, int, object | None]] = queue.Queue()
        self.update_queue: queue.Queue[tuple[str, object | None]] = queue.Queue()
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
        self._tool_run_token = 0
        self._tool_running = False
        self._idle_run_button_text = ""

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
        self._poll_status_queue()
        self._poll_update_queue()
        self.root.after(600, self._check_updates_on_startup)
        # 清理历史更新遗留的临时文件（下载包、解压目录），后台低优先执行
        threading.Thread(target=cleanup_stale_update_files, daemon=True).start()

    def _on_tk_callback_exception(self, exc_type, exc_value, exc_tb) -> None:
        runlog.log_exception("界面异常", exc_value, exc_tb)
        import traceback as traceback_module

        traceback_module.print_exception(exc_type, exc_value, exc_tb)

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

    def _px(self, value: int | float) -> int:
        return _scale_px(value, self.ui_scale)

    def _pxf(self, value: int | float) -> float:
        return _scale_float(value, self.ui_scale)

    def _pad(self, *values: int | float) -> tuple[int, ...]:
        return tuple(self._px(value) for value in values)

    def _window_size(self, width: int, height: int) -> tuple[int, int]:
        scaled_width = self._px(width)
        scaled_height = self._px(height)
        if self.ui_scale <= 1.0:
            return scaled_width, scaled_height
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            max_width = max(1, screen_width - self._px(48))
            max_height = max(1, screen_height - self._px(72))
        except Exception:
            return scaled_width, scaled_height
        return min(scaled_width, max_width), min(scaled_height, max_height)

    def _update_dialog_size(self, width: int, height: int) -> tuple[int, int]:
        scaled_width = self._px(width)
        scaled_height = self._px(height)
        try:
            max_width = max(1, self.root.winfo_screenwidth() - self._px(48))
            max_height = max(1, self.root.winfo_screenheight() - self._px(72))
        except Exception:
            return scaled_width, scaled_height
        return min(scaled_width, max_width), min(scaled_height, max_height)

    def _logical_screen_width(self) -> float:
        try:
            return self.root.winfo_screenwidth() / max(self.ui_scale, 1.0)
        except Exception:
            return 1180.0

    def _responsive_content_padding(self) -> tuple[int, int, int, int]:
        logical_width = self._logical_screen_width()
        if self.ui_scale >= 1.75 and logical_width < 900:
            return self._pad(16, 24, 16, 20)
        if logical_width < 1100:
            return self._pad(24, 28, 28, 24)
        return self._pad(42, 34, 58, 28)

    def _responsive_form_padding_units(self) -> tuple[int, int, int, int]:
        logical_width = self._logical_screen_width()
        if self.ui_scale >= 1.75 and logical_width < 900:
            return (12, 16, 12, 16)
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

        # Scrollable right panel: Canvas acts as the viewport; right_frame is
        # the inner content frame that all existing children are placed into.
        right_outer = ttk.Frame(root_frame, style="Content.TFrame")
        right_outer.pack(side=RIGHT, fill=BOTH, expand=True)

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

        def _sync_right_canvas_window(_event=None):
            canvas_width = max(self._right_canvas.winfo_width(), 1)
            canvas_height = max(self._right_canvas.winfo_height(), 1)
            content_height = _right_frame_natural_height()
            window_height = max(content_height, canvas_height)
            self._right_canvas.itemconfig(
                self._right_canvas_window,
                width=canvas_width,
                height=window_height,
            )
            self._right_canvas.configure(
                scrollregion=(0, 0, canvas_width, window_height)
            )
            if window_height <= canvas_height:
                self._right_canvas.yview_moveto(0)

        def _run_right_canvas_sync():
            self._right_canvas_sync_pending = False
            _sync_right_canvas_window()
            if self._right_canvas_sync_repeat > 0:
                self._right_canvas_sync_repeat -= 1
                _queue_right_canvas_sync()

        def _queue_right_canvas_sync() -> None:
            if self._right_canvas_sync_pending:
                return
            self._right_canvas_sync_pending = True
            self.root.after_idle(_run_right_canvas_sync)

        def _schedule_right_canvas_sync(_event=None):
            self._right_canvas_sync_repeat = max(self._right_canvas_sync_repeat, 2)
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

        def _update_text_wraps(_event=None) -> None:
            title_row_width = title_row.winfo_width()
            if title_row_width <= 1:
                title_row_width = self._px(240)
            actions_width = self._px(118)
            tight_header = title_row_width < self._px(480)
            if tight_header:
                title_actions.grid_configure(row=2, column=0, columnspan=2, rowspan=1, sticky="w", pady=self._pad(12, 0))
                title_wrap = title_row_width
            else:
                title_actions.grid_configure(row=0, column=1, columnspan=1, rowspan=2, sticky="ne", pady=0)
                title_wrap = title_row_width - actions_width - self._px(24)

            title_wrap = max(1, title_wrap)
            subtitle_wrap = max(1, title_row_width - self._px(8))
            self.title_label.configure(wraplength=title_wrap)
            self.subtitle_label.configure(wraplength=subtitle_wrap)

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
        self._form_compact_layout = False
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

        self.rename_options_frame = ttk.LabelFrame(form, text="文件夹改名", padding=self._px(12), style="Rename.TLabelframe")
        self.rename_options_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=self._pad(10, 0))
        ttk.Label(self.rename_options_frame, text="操作", style="App.TLabel").grid(row=0, column=0, sticky="w", pady=self._px(5))
        self.rename_mode_widget = ttk.Combobox(
            self.rename_options_frame,
            textvariable=self.rename_mode,
            values=list(RENAME_MODE_LABELS.keys()),
            state="readonly",
            width=16,
            style="App.TCombobox",
        )
        self.rename_mode_widget.grid(row=0, column=1, sticky="w", padx=self._px(12), pady=self._px(5))
        self.rename_mode_widget.bind("<<ComboboxSelected>>", self._on_rename_mode_changed)

        ttk.Label(self.rename_options_frame, textvariable=self.rename_target_label, style="App.TLabel").grid(row=1, column=0, sticky="w", pady=self._px(5))
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

        self.stats_range_label = ttk.Label(form, text="周报统计日期（可选）", style="App.TLabel")
        self.stats_range_frame = ttk.Frame(form, style="InputWrap.TFrame")
        stats_range_inputs = ttk.Frame(self.stats_range_frame, style="InputWrap.TFrame")
        stats_range_inputs.pack(side="top", fill="x")
        self.stats_week_start_entry = ttk.Entry(stats_range_inputs, textvariable=self.stats_week_start, width=12, style="App.TEntry")
        self.stats_week_start_entry.pack(side=LEFT)
        ttk.Label(stats_range_inputs, text="至", style="App.TLabel").pack(side=LEFT, padx=self._px(8))
        self.stats_week_end_entry = ttk.Entry(stats_range_inputs, textvariable=self.stats_week_end, width=12, style="App.TEntry")
        self.stats_week_end_entry.pack(side=LEFT)
        ttk.Label(stats_range_inputs, text="如 2026-06-02，留空按整月统计", style="App.TLabel").pack(side=LEFT, padx=self._pad(10, 0))
        stats_range_presets = ttk.Frame(self.stats_range_frame, style="InputWrap.TFrame")
        stats_range_presets.pack(side="top", fill="x", pady=self._pad(6, 0))
        for preset_text, preset_key in (("本月", "this_month"), ("上月", "last_month"), ("本周", "this_week"), ("上周", "last_week"), ("清空", "clear")):
            button = CodexButton(
                stats_range_presets,
                text=preset_text,
                command=lambda key=preset_key: self._fill_stats_week_range(key),
                width=56,
                min_width=56,
                height=28,
            )
            button.pack(side=LEFT, padx=self._pad(0, 8))
        stats_remark_unit_row = ttk.Frame(self.stats_range_frame, style="InputWrap.TFrame")
        stats_remark_unit_row.pack(side="top", fill="x", pady=self._pad(8, 0))
        ttk.Label(stats_remark_unit_row, text="备注加班/调休单位", style="App.TLabel").pack(side=LEFT)
        ttk.Radiobutton(
            stats_remark_unit_row,
            text="按天",
            value="day",
            variable=self.stats_remark_unit,
            style="App.TRadiobutton",
        ).pack(side=LEFT, padx=self._pad(12, 0))
        ttk.Radiobutton(
            stats_remark_unit_row,
            text="按小时",
            value="hour",
            variable=self.stats_remark_unit,
            style="App.TRadiobutton",
        ).pack(side=LEFT, padx=self._pad(8, 0))
        ttk.Label(
            stats_remark_unit_row,
            text="仅影响考勤统计表备注中加班/调休的显示",
            style="App.TLabel",
        ).pack(side=LEFT, padx=self._pad(10, 0))

        def _refresh_picker_button_bar(button_bar) -> None:
            visible_buttons = [child for child in button_bar.winfo_children() if getattr(child, "_hr_picker_visible", False)]
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

        def _layout_input_frame(row_data) -> None:
            entry = row_data["entry"]
            button_bar = row_data["button_bar"]
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

        def _apply_form_layout() -> None:
            form.columnconfigure(0, weight=1 if self._form_compact_layout else 0)
            form.columnconfigure(1, weight=0 if self._form_compact_layout else 1)
            visible_keys = []
            if self._summary_row_visible:
                visible_keys.append("summary")
            if self._output_row_visible:
                visible_keys.append("output")

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

            if self._stats_range_row_visible:
                if self._form_compact_layout:
                    base_row = len(visible_keys) * 2
                    self.stats_range_label.grid(row=base_row, column=0, columnspan=2, sticky="w", padx=0, pady=self._pad(4, 2))
                    self.stats_range_frame.grid(row=base_row + 1, column=0, columnspan=2, sticky="ew", padx=0, pady=self._pad(0, 8))
                else:
                    self.stats_range_label.grid(row=4, column=0, sticky="w", padx=0, pady=self._px(7))
                    self.stats_range_frame.grid(row=4, column=1, sticky="ew", padx=self._pad(12, 0), pady=self._px(7))
            else:
                self.stats_range_label.grid_remove()
                self.stats_range_frame.grid_remove()

            if hasattr(self, "_sync_right_canvas_window"):
                self.root.after_idle(self._sync_right_canvas_window)

        def _update_form_responsive_layout(_event=None) -> None:
            content_padding = self._responsive_content_padding()
            form_padding = self._responsive_form_padding_units()
            # 内容列限宽居中（对应设计稿 max-width:780 的主内容列）
            canvas_width = self._right_canvas.winfo_width()
            if canvas_width > 1:
                base_left, pad_top, base_right, pad_bottom = content_padding
                extra = max(0, (canvas_width - base_left - base_right - self._px(820)) // 2)
                content_padding = (base_left + extra, pad_top, base_right + extra, pad_bottom)
            if getattr(self, "_right_content_padding", None) != content_padding:
                self._right_content_padding = content_padding
                right_frame.configure(padding=content_padding)
            if getattr(self, "_form_padding", None) != form_padding:
                self._form_padding = form_padding
                self.form_card.set_padding(form_padding)
            canvas_width = self._right_canvas.winfo_width()
            compact = canvas_width > 1 and (canvas_width / max(self.ui_scale, 1.0)) < 560
            if compact != self._form_compact_layout:
                self._form_compact_layout = compact
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
        self._update_summary_controls()
        self._update_output_controls()
        self._update_rename_controls()
        self._update_stats_range_controls()

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
        self.root.update_idletasks()
        _sync_right_canvas_window()

    def _check_updates_on_startup(self) -> None:
        if not update_check_enabled():
            return
        self._start_update_check(manual=False)

    def _check_updates_manually(self) -> None:
        self._start_update_check(manual=True)

    def _start_update_check(self, manual: bool) -> None:
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
            self.check_update_button.config(state="disabled")
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
        self.root.after(150, self._poll_update_queue)

    def _finish_update_check(self) -> None:
        self.update_check_in_progress = False
        self.manual_update_check_active = False
        if hasattr(self, "check_update_button"):
            self.check_update_button.config(state="normal")

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
            close_command=None,
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

    def _start_update_download(self, update: UpdateInfo) -> None:
        self._write_log(f"开始下载更新包：v{update.version}")
        self._write_log(f"下载地址：{update.file_url}")
        self._download_speed_anchor = None
        self._show_update_progress_window(
            title=f"正在下载 v{update.version}",
            detail="请不要关闭程序，下载完成后会自动开始安装。",
            indeterminate=False,
            close_command=None,
        )

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
            return
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
        self._clear_log()
        self._write_log(self._initial_log_text())

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
        self._clear_log()
        self._write_log(self._initial_log_text())

    def _refresh_nav_buttons(self) -> None:
        for tool_id, item in self.nav_buttons.items():
            item.set_selected(tool_id == self.current_tool)

    def _set_tool_texts(self) -> None:
        self.tool_group.set(TOOL_GROUP_LABELS.get(self.current_tool, "人员运营自动化"))
        multi_hint = "支持 .xlsx / .xls / .zip / 文件夹 · 可多选"
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
            self.tool_description.set("选择人员资料目录，先预览，再确认改名。")
            self.input_label.set("人员文件夹目录")
            self.input_hint.set("选择人员资料所在目录")
            self._input_drop_title = "选择人员文件夹目录"
            self.choose_input_text.set("选择文件夹")
            self.summary_label.set("")
            self.summary_button_text.set("选择")
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
            self._update_summary_controls()
            self._update_output_controls()
            self._update_rename_controls()
            self._update_stats_range_controls()
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
            return
        name = Path(text).name or text
        self.summary_display.configure(text=name, fg=COLOR_TEXT, font=(self.base_font[0], _font_size(10), "bold"))

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
        if suffix == ".zip":
            return "ZIP", COLOR_BADGE_ZIP_BG, COLOR_BADGE_ZIP_FG, size_text
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

        def redraw(_event=None) -> None:
            chip.delete("all")
            width = max(chip.winfo_width(), 1)
            height = max(chip.winfo_height(), 1)
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

        def redraw(_event=None) -> None:
            zone.delete("all")
            width = max(zone.winfo_width(), 1)
            height = max(zone.winfo_height(), 1)
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
                links.append(("选择文件夹", self._input_folder_cmd))
            link_y = title_y + self._pxf(21)
            segments: list[tuple[str, str, object | None]] = []
            if links:
                segments.append(("或 " if len(links) > 0 else "", COLOR_FAINT, None))
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
                item = zone.create_text(0, -100, text=text, font=font, anchor="w")
                bbox = zone.bbox(item)
                segment_width = (bbox[2] - bbox[0]) if bbox else 0
                zone.delete(item)
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
                ("步骤：选择单个缴费清单、多个清单、zip压缩包，或包含清单的文件夹；再选择参保人员花名册。", None),
                ("结果：生成“社保明细表.xlsx”和“社保汇总表.xlsx”，汇总表里含基础数据分析和异常提醒。", None),
                ("目前规则：按身份证关联花名册；优先按账单文件夹或文件名识别账单月份、缴纳地和缴纳单位。", None),
                ("注意：公积金、残保金、管理费暂无数据时留空；账单识别结果与花名册不一致时会提醒。", "warning"),
            ]
        if tool_id == "data_statistics":
            return [
                ("适用：把 HR 系统导出的考勤结果、周报记录、月报记录自动整理成统计表。", "strong"),
                ("步骤：选择单个文件、多个文件、zip压缩包，或包含这些文件的文件夹。", None),
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
                ("步骤：选择单个保单清单、多个清单、zip压缩包，或包含清单的文件夹；再选择人力资源分析表。", None),
                ("结果：生成“保险台账.xlsx”，包含保险台账和人员增减预警两个工作表。", None),
                ("当前规则：PZDX保额取“每人伤残死亡限额”，按万元显示；PEAC保额固定按60万元。", None),
                ("注意：人力资源分析表需包含“花名册”工作表；花名册在职但保单没有会提示需加保，保单有但花名册没有或已标记离职会提示需减保。", "warning"),
            ]
        if tool_id == "salary_merge":
            return [
                ("适用：把 1-12 个月工资表合成一张个人应发工资汇总表。", "strong"),
                ("步骤：可选择单个月度工资表、多个工资表、zip压缩包，或包含这些文件的文件夹。", None),
                ("如已有前几月汇总表，再选择“已有汇总表”；不选则新建一张汇总表。", None),
                ("点击“开始合并”后，结果会生成到保存位置下的新文件夹中。", None),
                ("结果：按姓名、身份证号、月份合并；没有工资的月份填 0；已存在的人员月份不会覆盖。", None),
                ("注意：工资表文件名或表内日期要能识别月份；重复人员或重复月份会在执行结果里提醒。", "warning"),
            ]
        if tool_id == "personnel_change_merge":
            if mode == "roster":
                return [
                    ("适用：已有月度异动汇总表时，单独更新人力资源花名册。", "strong"),
                    ("步骤：选择单个异动汇总表、多个汇总表，或包含汇总表的文件夹；再选择人力资源花名册。", None),
                    ("点击“更新花名册”后，结果会生成到保存位置下的新文件夹中。", None),
                    ("结果：根据汇总表里的增员写入花名册，根据减员在花名册中标记离职。", None),
                    ("注意：不会清空原花名册；身份证已存在的增员不会重复写入，找不到的减员会在日志提醒。", "warning"),
                ]
            return [
                ("适用：把项目异动表按记录日期分到对应月份汇总表。", "strong"),
                ("步骤：可选择单个异动表、多个异动表、zip压缩包，或包含这些文件的文件夹。", None),
                ("如已有月度汇总表，可选择单个汇总表或包含多个汇总表的文件夹；工具会按月份追加，原有记录不会清空。", None),
                ("不选择已有汇总表时，工具会按月份新建干净汇总表。缺少某个月份汇总表时也会自动创建。", None),
                ("如果同一文件夹里放了人力资源分析表，工具会自动更新其中的花名册。", None),
                ("点击“开始汇总”后，结果会生成到保存位置下的新文件夹中。", None),
                ("月份规则：增员看入职日期，减员看离职日期，转正看转正日期，调动看调整日期。", None),
                ("注意：只处理增补表、离职、转正、调整；薪酬、产值和同行对比分析暂不处理。", "warning"),
            ]
        if tool_id == "archive_import":
            if mode == "export":
                return [
                    ("适用：把一个或多个档案汇总表写入各公司独立档案表。", "strong"),
                    ("步骤：选择档案汇总表文件、多个文件、zip压缩包，或包含汇总表的文件夹。", None),
                    ("如已有某个公司的档案表，可选择文件、zip压缩包或文件夹；不选或没匹配到时会按内置干净模板新建。", None),
                    ("结果：按公司生成独立 Excel；已有身份证不重复新增，只补充空白字段。", None),
                    ("注意：公司档案表会自动改公司名，新增行会补边框、居中和公式。", "warning"),
                ]
            return [
                ("适用：把项目部提交的人事档案移交表写入公司档案汇总表。", "strong"),
                ("步骤：可选择单个移交表、多个移交表、zip压缩包，或包含这些文件的文件夹。", None),
                ("已有档案汇总表可不选；不选时工具会用内置空模板新建一份汇总表。", None),
                ("结果：按“公司”写入对应工作表；身份证已存在时不重复新增，只补充空白材料字段。", None),
                ("注意：编号会从文件名或表头标题识别项目地区，如“茂名项目部”自动填 11；识别不到会留空并提醒。", "warning"),
            ]
        if tool_id == "folder_rename":
            return [
                ("适用：批量修改所选目录下第一层人员文件夹名称。", "strong"),
                ("追加文字：姓名不填就是全部文件夹追加；填姓名就是只处理这个人。输入“劳动合同”会追加为“-劳动合同”。", None),
                ("删除结尾文字：输入“_劳动合同”，可删除“张三_劳动合同 / 张三-劳动合同 / 张三劳动合同”的结尾文字。", None),
                ("修改单人名称：填写原姓名和新名称，例如“张三”改为“章五”。", None),
                ("重要提醒：改名会直接改变真实文件夹名称。必须先看预览，确认无误后再点确认；建议操作前先备份。", "warning"),
            ]
        if tool_id == "salary_split":
            return [
                ("适用：一个完整工资表按“入职公司”拆成多个公司工资表。", "strong"),
                ("步骤：选择工资表文件，保存位置默认在桌面“工资表拆分结果”，点击“开始拆分”。", None),
                ("点击“打开所在文件夹”可直接查看本次生成的结果目录。", None),
                ("结果：每个入职公司生成一个 Excel，保留表头、格式、公式、小计和底部总计。", None),
                ("注意：源工资表不会被修改；如果模板列名或表结构变化，先发给开发确认。", "warning"),
            ]
        return [
            ("该工具暂未实现。", "strong"),
            ("请选择左侧已完成的工具：需求1、需求2、需求4、需求5、需求6、需求7、需求8。", None),
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

    def _update_summary_controls(self) -> None:
        self._summary_row_visible = self.current_tool in {"social_security", "data_statistics", "insurance_ledger", "salary_merge", "personnel_change_merge", "archive_import"}
        if hasattr(self, "_apply_form_layout"):
            self._apply_form_layout()

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
        elif tool == "folder_rename":
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
        else:
            hide(self.change_summary_folder_button, self.change_summary_file_button)
            show(self.summary_choose_button)

        self._refresh_upload_card()

    def _update_output_controls(self) -> None:
        self._output_row_visible = self.current_tool != "folder_rename"
        if hasattr(self, "_apply_form_layout"):
            self._apply_form_layout()

    def _update_rename_controls(self) -> None:
        self._rename_row_visible = self.current_tool == "folder_rename"
        if hasattr(self, "_apply_form_layout"):
            self._apply_form_layout()
        if self._rename_row_visible:
            self._update_rename_mode_controls()

    def _update_stats_range_controls(self) -> None:
        self._stats_range_row_visible = self.current_tool == "data_statistics"
        if hasattr(self, "_apply_form_layout"):
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

    def _on_rename_mode_changed(self, _event=None) -> None:
        self._update_rename_mode_controls()
        if hasattr(self, "_sync_right_canvas_window"):
            self.root.after_idle(self._sync_right_canvas_window)

    def _update_rename_mode_controls(self) -> None:
        mode = RENAME_MODE_LABELS.get(self.rename_mode.get(), MODE_APPEND)
        # 文件类型选择器始终显示
        self.rename_file_type_label_widget.grid(row=4, column=0, sticky="w", pady=self._px(5))
        self.rename_file_type_widget.grid(row=4, column=1, sticky="w", padx=self._px(12), pady=self._px(5))
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
            return "请选择社保缴费清单、参保人员花名册和保存位置，然后点击“生成报表”。"
        if self.current_tool == "data_statistics":
            return "请选择考勤结果、周报记录、月报记录文件或文件夹和保存位置，然后点击“生成统计”。应汇报人员名单是可选项。"
        if self.current_tool == "insurance_ledger":
            return "请选择保单人员清单、人力资源分析表和保存位置，然后点击“生成台账”。"
        if self.current_tool == "salary_merge":
            return "请选择工资表文件、压缩包或文件夹和保存位置，然后点击“开始合并”。已有汇总表是可选项，用于追加新月份。"
        if self.current_tool == "personnel_change_merge":
            if self.change_mode == "roster":
                return "请选择异动汇总表、人力资源花名册和保存位置，然后点击“更新花名册”。"
            return "请选择异动表文件或文件夹和保存位置，然后点击“开始汇总”。已有汇总表是可选项，用于追加新记录。"
        if self.current_tool == "archive_import":
            if self.archive_mode == "export":
                return "请选择档案汇总表、压缩包或文件夹和保存位置，然后点击“生成档案表”。已有公司档案表是可选项，用于追加。"
            return "请选择移交表文件、压缩包或文件夹和保存位置，然后点击“开始入库”。已有档案汇总表是可选项。"
        if self.current_tool == "folder_rename":
            return "请选择人员文件夹目录，填写改名内容，然后点击“预览”。"
        if self.current_tool == "salary_split":
            return "请选择工资表文件和保存位置，然后点击“开始拆分”。"
        return "该工具暂未实现。"

    def _choose_input(self) -> None:
        if self.current_tool in {"salary_merge", "personnel_change_merge", "folder_rename", "archive_import"}:
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
            else:
                title = "选择工资表文件夹"
            directory = filedialog.askdirectory(title=title)
            if directory:
                self.input_path.set(directory)
                if not self.output_dir_user_selected:
                    self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
                self._refresh_upload_card()
            return

        filename = filedialog.askopenfilename(
            title="选择工资表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.input_path.set(filename)
            if not self.output_dir_user_selected:
                self.output_dir.set(str(default_output_parent_dir(self.current_tool)))
            self._refresh_upload_card()

    def _choose_change_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择异动表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_change_files_or_zip(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="选择异动表文件或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_salary_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择工资表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_salary_files_or_zip(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="选择工资表文件或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_social_security_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择社保缴费清单文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_social_security_files_or_zip(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="选择社保缴费清单或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_social_security_roster_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择参保人员花名册",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_data_statistics_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择考勤周月报数据文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_data_statistics_files_or_zip(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="选择考勤周月报文件或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_data_statistics_staff_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择应汇报人员名单",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_insurance_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择保单人员清单文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_insurance_files_or_zip(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="选择保单人员清单或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_insurance_roster_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择人力资源分析表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_archive_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择档案移交表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_archive_files_or_zip(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="选择档案移交表文件或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_archive_summary_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择已有档案汇总表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_archive_export_summary_files_or_zip(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="选择档案汇总表或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
        )
        if filenames:
            self._set_change_input_paths([Path(filename) for filename in filenames])

    def _choose_archive_export_summary_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择档案汇总表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_archive_export_existing_file_or_zip(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择已有公司档案表或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_archive_export_existing_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择已有公司档案表文件夹")
        if directory:
            self.summary_path.set(directory)

    def _choose_roster_summary_folder(self) -> None:
        directory = filedialog.askdirectory(title="选择异动汇总表文件夹")
        if directory:
            self._set_change_input_paths([Path(directory)])

    def _choose_roster_summary_files(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="选择异动汇总表文件或压缩包",
            filetypes=[("Excel 或 ZIP", "*.xlsx *.xls *.zip"), ("Excel 工作簿", "*.xlsx *.xls"), ("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
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
        directory = filedialog.askdirectory(title="选择已有异动汇总表文件夹")
        if directory:
            self.summary_path.set(directory)

    def _choose_change_summary_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择已有异动汇总表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_roster_analysis_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择人力资源花名册",
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _choose_output(self) -> None:
        directory = filedialog.askdirectory(title="选择保存位置")
        if directory:
            self.output_dir_user_selected = True
            self.output_dir.set(directory)

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
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel 工作簿", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.summary_path.set(filename)

    def _run_current_tool(self) -> None:
        if self._tool_running:
            self._stop_tool_run()
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
        if input_path.suffix.lower() not in {".xlsx", ".xls"}:
            messagebox.showwarning("格式不支持", "当前工具只支持 .xlsx 或 .xls 工资表。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_parent_dir = Path(output_text)

        output_dir = make_result_output_dir(output_parent_dir)
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
            if input_path.is_file() and input_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
                messagebox.showwarning("格式不支持", "工资表文件只支持 .xlsx、.xls 或 .zip。")
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

        output_dir = make_result_output_dir(output_parent_dir)
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
            if input_path.is_file() and input_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
                messagebox.showwarning("格式不支持", "社保缴费清单只支持 .xlsx、.xls 或 .zip。")
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

        output_dir = make_result_output_dir(Path(output_text))
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
            if input_path.is_file() and input_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
                messagebox.showwarning("格式不支持", "数据文件只支持 .xlsx、.xls 或 .zip。")
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

        output_dir = make_result_output_dir(Path(output_text))
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
            remark_unit=self.stats_remark_unit.get() or "day",
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
            if input_path.is_file() and input_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
                messagebox.showwarning("格式不支持", "保单人员清单只支持 .xlsx、.xls 或 .zip。")
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

        output_dir = make_result_output_dir(Path(output_text))
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
            if input_path.is_file() and input_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
                messagebox.showwarning("格式不支持", "异动表文件只支持 .xlsx、.xls 或 .zip。")
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

        output_dir = make_result_output_dir(output_parent_dir)
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
            if input_path.is_file() and input_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
                messagebox.showwarning("格式不支持", "异动汇总表只支持 .xlsx、.xls、.zip 文件或文件夹。")
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
        output_dir = make_result_output_dir(Path(output_text))
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
            if input_path.is_file() and input_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
                messagebox.showwarning("格式不支持", "档案移交表文件只支持 .xlsx、.xls 或 .zip。")
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
        output_dir = make_result_output_dir(Path(output_text))
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
            if summary_path.is_file() and summary_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
                messagebox.showwarning("格式不支持", "档案汇总表目前只支持 .xlsx、.xls 或 .zip。")
                return
        existing_path = Path(existing_text) if existing_text else None
        if existing_path is not None and not existing_path.exists():
            messagebox.showwarning("档案表不存在", "选择的已有公司档案表不存在，请重新选择。")
            return
        if existing_path is not None and existing_path.is_file() and existing_path.suffix.lower() not in {".xlsx", ".xls", ".zip"}:
            messagebox.showwarning("格式不支持", "已有公司档案表目前只支持 .xlsx、.xls 或 .zip。")
            return
        if not output_text:
            messagebox.showwarning("缺少目录", "请选择保存位置。")
            return
        output_dir = make_result_output_dir(Path(output_text))
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
        # 获取文件类型
        file_type_label = self.rename_file_type.get()
        file_type = RENAME_FILE_TYPE_LABELS.get(file_type_label, FILE_TYPE_FOLDER)

        try:
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
            messagebox.showinfo("没有可改名项目", "没有找到需要改名的项目，请检查输入内容。")
            return

        message = self._folder_rename_confirm_message(preview)
        if not messagebox.askyesno("确认改名", message):
            self._write_log("已取消执行。")
            return

        self._begin_tool_run()
        self._write_log("开始执行改名...")
        self._start_tool_worker(
            rename_person_folders,
            root_dir=root_dir,
            mode=mode,
            text=self.rename_text.get(),
            target_name=self.rename_target_name.get(),
            replacement_name=self.rename_replacement_name.get(),
            file_type=file_type,
        )

    def _begin_tool_run(self) -> None:
        """进入运行状态：主按钮变为“停止”，并为本次运行分配编号。"""
        self._tool_run_token += 1
        self._tool_running = True
        self._idle_run_button_text = self.run_button_text.get()
        self.run_button_text.set("停止")

    def _finish_tool_run(self) -> None:
        self._tool_running = False
        if self._idle_run_button_text:
            self.run_button_text.set(self._idle_run_button_text)

    def _stop_tool_run(self) -> None:
        # 工具函数本身不可中断：递增运行编号让后台任务的结果在返回时被丢弃，
        # 界面立即恢复可用。被停止的任务如已写出文件，直接忽略即可。
        self._tool_run_token += 1
        self._finish_tool_run()
        self._write_log("已停止本次生成。")
        runlog.log_line(f"用户停止了 {self._tool_log_label()}。")

    def _tool_log_label(self) -> str:
        if self.current_tool == "personnel_change_merge" and self.change_mode == "roster":
            return "花名册更新"
        if self.current_tool == "archive_import" and self.archive_mode == "export":
            return "档案表生成"
        return TOOL_LOG_LABELS.get(self.current_tool, self.current_tool)

    def _start_tool_worker(self, tool_func, /, *args, **kwargs) -> None:
        token = self._tool_run_token
        label = self._tool_log_label()
        details = runlog.describe_call(args, kwargs)
        runlog.log_line(f"开始 {label}：{details}" if details else f"开始 {label}")

        def worker() -> None:
            start = time.monotonic()
            try:
                result = tool_func(*args, **kwargs)
            except Exception as exc:
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
                if token != self._tool_run_token:
                    self._write_log("（已停止的任务在后台结束，结果已忽略。）")
                    continue
                self._finish_tool_run()
                if status == "success":
                    self._record_last_run(True)
                    self._handle_success(payload)
                elif status == "error":
                    self._record_last_run(False)
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
            else:
                message = "处理完成。"
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

    def _flush_log_view(self) -> None:
        self.log_text.see(END)
        self.log_text.update_idletasks()
        self.root.update_idletasks()

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", END)


def _default_result_dir_name() -> str:
    return "结果_" + datetime.now().strftime("%Y%m%d_%H%M%S")


def default_output_parent_dir(tool: str) -> Path:
    if tool == "social_security":
        folder_name = "社保汇总结果"
    elif tool == "data_statistics":
        folder_name = "数据统计结果"
    elif tool == "insurance_ledger":
        folder_name = "保险台账结果"
    elif tool == "salary_merge":
        folder_name = "工资合并结果"
    elif tool == "personnel_change_merge":
        folder_name = "异动表汇总结果"
    elif tool == "archive_import":
        folder_name = "档案处理结果"
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


def open_path(path: Path | str) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def _enable_high_dpi_rendering() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
    except Exception:
        return

    # Tk widgets are scaled once at startup; System-aware avoids runtime DPI
    # changes that can leave fixed Canvas geometry out of sync.
    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-2)):
            return
    except Exception:
        pass
    try:
        if ctypes.windll.shcore.SetProcessDpiAwareness(1) == 0:
            return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def main() -> None:
    _install_crash_logging()
    _set_windows_app_identity()
    _enable_high_dpi_rendering()
    root = Tk()
    HRToolkitApp(root)
    root.mainloop()


def _set_windows_app_identity() -> None:
    """给进程声明独立的应用身份（AppUserModelID）。

    Windows 任务栏按 AppUserModelID 归组图标：用 python.exe 直接运行时，
    窗口会被归到“Python”名下，任务栏显示 python 的图标而不是 iconphoto
    设置的应用图标。显式声明后，任务栏改用窗口自己的图标；对打包 exe
    也顺带让固定到任务栏的身份保持稳定。"""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("HRWorkbench.HRToolkit")
    except Exception:
        pass


def _install_crash_logging() -> None:
    """把没被捕获的异常写进运行日志。

    打包后的 --windowed 程序没有控制台，未捕获异常会无声消失，
    这是 HR 电脑上"程序突然不见了"却查无线索的根源。"""
    default_excepthook = sys.excepthook

    def log_and_delegate(exc_type, exc_value, exc_tb):
        runlog.log_exception("程序异常退出", exc_value, exc_tb)
        default_excepthook(exc_type, exc_value, exc_tb)

    sys.excepthook = log_and_delegate

    default_thread_hook = threading.excepthook

    def log_thread_exception(args) -> None:
        if args.exc_value is not None:
            runlog.log_exception(f"后台线程异常（{args.thread.name if args.thread else '未知'}）", args.exc_value, args.exc_traceback)
        default_thread_hook(args)

    threading.excepthook = log_thread_exception
