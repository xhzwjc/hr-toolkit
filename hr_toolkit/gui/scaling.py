"""DPI and scaling detection, coordinate calculation, and font scaling helpers."""

from __future__ import annotations

import os
import sys
from tkinter import Tk, TkVersion

from .constants import BASE_WINDOWS_DPI, FORCE_UI_SCALE_ENV, TK_POINTS_PER_INCH


LAYOUT_MODE_WIDE = "wide"
LAYOUT_MODE_COMPACT = "compact"
LAYOUT_MODE_NARROW = "narrow"


def _fit_window_size(
    width: int | float,
    height: int | float,
    available_width: int | float,
    available_height: int | float,
    *,
    horizontal_margin: int | float = 0,
    vertical_margin: int | float = 0,
) -> tuple[int, int]:
    """Fit a requested window inside the usable desktop rectangle."""
    requested_width = max(1, int(round(width)))
    requested_height = max(1, int(round(height)))
    max_width = max(1, int(round(available_width - max(0, horizontal_margin))))
    max_height = max(1, int(round(available_height - max(0, vertical_margin))))
    return min(requested_width, max_width), min(requested_height, max_height)


def _responsive_layout_mode(viewport_width: int | float) -> str:
    """Return the form layout mode for the actual usable logical width."""
    width = max(0.0, float(viewport_width))
    if width < 520:
        return LAYOUT_MODE_NARROW
    if width < 760:
        return LAYOUT_MODE_COMPACT
    return LAYOUT_MODE_WIDE


def _responsive_checkbox_columns(layout_mode: str) -> int:
    if layout_mode == LAYOUT_MODE_NARROW:
        return 1
    if layout_mode == LAYOUT_MODE_COMPACT:
        return 2
    return 4


def _responsive_drawer_width(
    available_width: int | float,
    preferred_width: int | float,
    *,
    edge_gap: int | float = 12,
) -> int:
    """Keep a temporary drawer fully inside its current parent viewport."""
    available = max(1, int(round(available_width)))
    preferred = max(1, int(round(preferred_width)))
    maximum = max(1, available - max(0, int(round(edge_gap))))
    return min(preferred, maximum)


def _windows_work_area_for_root(root: Tk) -> tuple[int, int, int, int] | None:
    """Return the nearest Windows monitor work area, excluding the taskbar."""
    if not sys.platform.startswith("win"):
        return None
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return None

    class _MonitorInfo(ctypes.Structure):
        _fields_ = (
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", wintypes.RECT),
            ("rcWork", wintypes.RECT),
            ("dwFlags", wintypes.DWORD),
        )

    try:
        hwnd = int(root.winfo_id())
        monitor_from_window = ctypes.windll.user32.MonitorFromWindow
        monitor_from_window.argtypes = (wintypes.HWND, wintypes.DWORD)
        monitor_from_window.restype = wintypes.HANDLE
        get_monitor_info = ctypes.windll.user32.GetMonitorInfoW
        get_monitor_info.argtypes = (wintypes.HANDLE, ctypes.POINTER(_MonitorInfo))
        get_monitor_info.restype = wintypes.BOOL
        monitor = monitor_from_window(wintypes.HWND(hwnd), 2)
        if monitor:
            info = _MonitorInfo()
            info.cbSize = ctypes.sizeof(_MonitorInfo)
            if get_monitor_info(monitor, ctypes.byref(info)):
                work = info.rcWork
                if work.right > work.left and work.bottom > work.top:
                    return int(work.left), int(work.top), int(work.right), int(work.bottom)
    except Exception:
        pass

    try:
        rect = wintypes.RECT()
        if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
            if rect.right > rect.left and rect.bottom > rect.top:
                return int(rect.left), int(rect.top), int(rect.right), int(rect.bottom)
    except Exception:
        pass
    return None


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

    Tk 8.6 的 macOS aqua 后端把“点”直接按像素渲染，同样的数值只有
    Windows 96dpi 基准的 3/4 大，因此需要放大 4/3。Tk 8.7 起已修复
    这个历史行为，继续补偿会造成二次放大；Windows/Linux 原值返回。
    """
    import hr_toolkit.gui as _gui

    tk_ver = getattr(_gui, "TkVersion", TkVersion)
    if sys.platform == "darwin" and tk_ver < 8.7:
        return max(1, round(size * 4 / 3))
    return size


def _widget_ui_scale(widget) -> float:
    try:
        return float(getattr(widget.winfo_toplevel(), "_hr_ui_scale", 1.0))
    except Exception:
        return 1.0


__all__ = [
    "LAYOUT_MODE_WIDE",
    "LAYOUT_MODE_COMPACT",
    "LAYOUT_MODE_NARROW",
    "_fit_window_size",
    "_responsive_layout_mode",
    "_responsive_checkbox_columns",
    "_responsive_drawer_width",
    "_windows_work_area_for_root",
    "_scale_px",
    "_scale_float",
    "_clamp_ui_scale",
    "_indeterminate_progress_segment",
    "_forced_ui_scale",
    "_windows_dpi_for_root",
    "_detect_ui_scale",
    "_configure_tk_font_scaling",
    "_font_size",
    "_widget_ui_scale",
]
