"""Low-risk live-resize helpers for Qt Quick desktop windows.

The Windows compositor can expose a newly enlarged part of a Qt Quick window
before the GPU swap chain has presented a frame at the new size.  Keeping the
native class background aligned with the QML clear color avoids a black flash,
while explicit QQuickWindow updates reduce the time until the new frame lands.
"""

from __future__ import annotations

import os
import sys
from typing import Callable


WINDOW_BACKGROUND_RGB = (0xF7, 0xF5, 0xF1)


def _windows_colorref(red: int, green: int, blue: int) -> int:
    """Pack RGB bytes into the COLORREF layout used by Win32 GDI."""

    return (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16)


class LiveResizeUpdater:
    """Request a fresh Qt Quick frame after every native size change.

    QQuickWindow.update() coalesces repeated calls, so a diagonal resize that
    emits both widthChanged and heightChanged still schedules only one frame.
    This stays inside Qt's supported signal/render path and does not subclass
    the Win32 window procedure.
    """

    def __init__(self, window) -> None:
        self._window = window
        self._connections: list[tuple[object, Callable[..., None]]] = []
        update = getattr(window, "update", None)
        if not callable(update):
            return

        for setter_name in ("setPersistentGraphics", "setPersistentSceneGraph"):
            setter = getattr(window, setter_name, None)
            if callable(setter):
                setter(True)

        for signal_name in ("widthChanged", "heightChanged"):
            signal = getattr(window, signal_name, None)
            connect = getattr(signal, "connect", None)
            if not callable(connect):
                continue
            connect(self._request_frame)
            self._connections.append((signal, self._request_frame))

    def _request_frame(self, *_args) -> None:
        window = self._window
        update = getattr(window, "update", None) if window is not None else None
        if callable(update):
            update()

    def close(self) -> None:
        connections = self._connections
        self._connections = []
        for signal, callback in connections:
            disconnect = getattr(signal, "disconnect", None)
            if callable(disconnect):
                try:
                    disconnect(callback)
                except (RuntimeError, TypeError):
                    pass
        self._window = None


class WindowsResizeBackdrop:
    """Give freshly exposed Win32 client pixels the application's background.

    This changes only the window class background brush.  It never replaces or
    subclasses Qt's WndProc, and it restores the previous brush before exit.
    """

    _GCLP_HBRBACKGROUND = -10

    def __init__(self) -> None:
        self._active = False
        self._hwnd = 0
        self._brush = 0
        self._previous = 0
        self._getter = None
        self._setter = None
        self._delete_object = None

    @property
    def active(self) -> bool:
        return self._active

    @classmethod
    def install(cls, window) -> "WindowsResizeBackdrop":
        backdrop = cls()
        if not sys.platform.startswith("win"):
            return backdrop
        qpa_platform = (
            os.environ.get("QT_QPA_PLATFORM", "").split(":", 1)[0].casefold()
        )
        if qpa_platform in {"minimal", "offscreen"}:
            # Headless Qt platform plugins do not own a Win32 HWND.  Avoid
            # treating their synthetic winId as a native window handle.
            return backdrop

        try:
            import ctypes
            from ctypes import wintypes

            hwnd = int(window.winId())
            if not hwnd:
                return backdrop

            user32 = ctypes.WinDLL("user32", use_last_error=True)
            gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
            pointer_sized = ctypes.sizeof(ctypes.c_void_p) == 8
            if pointer_sized:
                getter = user32.GetClassLongPtrW
                setter = user32.SetClassLongPtrW
                value_type = ctypes.c_void_p
            else:
                getter = user32.GetClassLongW
                setter = user32.SetClassLongW
                value_type = wintypes.DWORD

            getter.argtypes = (wintypes.HWND, ctypes.c_int)
            getter.restype = value_type
            setter.argtypes = (wintypes.HWND, ctypes.c_int, value_type)
            setter.restype = value_type
            gdi32.CreateSolidBrush.argtypes = (wintypes.DWORD,)
            gdi32.CreateSolidBrush.restype = wintypes.HANDLE
            gdi32.DeleteObject.argtypes = (wintypes.HANDLE,)
            gdi32.DeleteObject.restype = wintypes.BOOL

            color_ref = _windows_colorref(*WINDOW_BACKGROUND_RGB)
            brush = int(gdi32.CreateSolidBrush(color_ref) or 0)
            if not brush:
                return backdrop
            # Store the owned GDI handle immediately so any later setup error
            # still releases it through close().
            backdrop._active = True
            backdrop._brush = brush
            backdrop._delete_object = gdi32.DeleteObject

            ctypes.set_last_error(0)
            previous_value = setter(
                hwnd,
                cls._GCLP_HBRBACKGROUND,
                value_type(brush),
            )
            last_error = ctypes.get_last_error()
            previous = int(previous_value or 0)
            if previous == 0 and last_error:
                backdrop.close()
                return backdrop

            backdrop._hwnd = hwnd
            backdrop._previous = previous
            backdrop._getter = getter
            backdrop._setter = setter
        except (AttributeError, OSError, TypeError, ValueError):
            backdrop.close()
        return backdrop

    def close(self) -> None:
        if not self._active:
            return
        self._active = False
        delete_object = self._delete_object
        brush = self._brush
        try:
            import ctypes
            from ctypes import wintypes

            getter = self._getter
            setter = self._setter
            if callable(getter) and callable(setter) and self._hwnd:
                current = int(
                    getter(self._hwnd, self._GCLP_HBRBACKGROUND) or 0
                )
                if current == self._brush:
                    value_type = (
                        ctypes.c_void_p
                        if ctypes.sizeof(ctypes.c_void_p) == 8
                        else wintypes.DWORD
                    )
                    setter(
                        self._hwnd,
                        self._GCLP_HBRBACKGROUND,
                        value_type(self._previous),
                    )
        except (AttributeError, OSError, TypeError, ValueError):
            pass
        finally:
            # Restoration can fail when Windows has already destroyed the
            # native window.  The brush is still ours and must always be
            # released independently to avoid leaking a GDI handle.
            if callable(delete_object) and brush:
                try:
                    from ctypes import wintypes

                    delete_object(wintypes.HANDLE(brush))
                except (AttributeError, OSError, TypeError, ValueError):
                    pass
            self._hwnd = 0
            self._brush = 0
            self._previous = 0
            self._getter = None
            self._setter = None
            self._delete_object = None


__all__ = [
    "LiveResizeUpdater",
    "WINDOW_BACKGROUND_RGB",
    "WindowsResizeBackdrop",
]
