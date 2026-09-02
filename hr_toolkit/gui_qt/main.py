"""Qt Quick application bootstrap."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from hr_toolkit import __version__, runlog
from hr_toolkit.desktop_helpers import install_crash_logging, set_windows_app_identity


class _WindowsSizingHook:
    """Keep Qt Quick rendering alive during the Windows modal sizing loop.

    When a user drags a window border on Windows, the OS enters a modal
    message loop (WM_SYSCOMMAND SC_SIZE) that blocks Qt's GUI thread.
    The render thread starves because scene-graph updates are triggered
    from the blocked GUI thread, causing the window content to freeze
    until the user releases the border.

    This hook subclasses the native window procedure via SetWindowLongPtrW
    to intercept WM_ENTERSIZEMOVE / WM_EXITSIZEMOVE directly.  During the
    modal loop it runs a 16 ms Win32 timer that posts WM_PAINT messages,
    which unblocks the DWM present path and lets the user see live content
    while dragging — matching the behaviour of Chromium-based desktop apps
    such as Codex, Claude Desktop and GitHub Desktop.
    """

    _WM_ENTERSIZEMOVE = 0x0231
    _WM_EXITSIZEMOVE = 0x0232
    _WM_PAINT = 0x000F
    _GWLP_WNDPROC = -4

    def __init__(self) -> None:
        self._hwnd: int | None = None
        self._timer_id: int | None = None
        self._sizing = False
        self._original_wndproc = None
        # Prevent GC of ctypes callbacks while the hook is alive.
        self._c_wndproc = None
        self._c_timer_proc = None

    def install(self, app, hwnd: int) -> None:
        try:
            import ctypes
            import ctypes.wintypes
        except ImportError:
            return

        self._hwnd = hwnd
        user32 = ctypes.windll.user32

        user32.SetWindowLongPtrW.argtypes = (
            ctypes.wintypes.HWND, ctypes.c_int, ctypes.c_void_p,
        )
        user32.SetWindowLongPtrW.restype = ctypes.c_void_p
        user32.GetWindowLongPtrW.argtypes = (
            ctypes.wintypes.HWND, ctypes.c_int,
        )
        user32.GetWindowLongPtrW.restype = ctypes.c_void_p
        user32.CallWindowProcW.argtypes = (
            ctypes.c_void_p, ctypes.wintypes.HWND, ctypes.wintypes.UINT,
            ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM,
        )
        user32.CallWindowProcW.restype = ctypes.c_long

        WNDPROC = ctypes.WINFUNCTYPE(
            ctypes.c_long, ctypes.wintypes.HWND, ctypes.wintypes.UINT,
            ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM,
        )
        TIMERPROC = ctypes.WINFUNCTYPE(
            None, ctypes.wintypes.HWND, ctypes.wintypes.UINT,
            ctypes.POINTER(ctypes.wintypes.UINT), ctypes.c_ulong,
        )

        hook = self

        def _timer_proc(_hwnd, _msg, _tid, _dw_time):
            if hook._sizing:
                user32.PostMessageW(hwnd, hook._WM_PAINT, 0, 0)

        self._c_timer_proc = TIMERPROC(_timer_proc)

        def _wndproc(hwnd_arg, msg, wparam, lparam):
            if msg == hook._WM_ENTERSIZEMOVE:
                hook._sizing = True
                hook._timer_id = user32.SetTimer(
                    hwnd_arg, 42, 16, hook._c_timer_proc,
                )
            elif msg == hook._WM_EXITSIZEMOVE:
                hook._sizing = False
                if hook._timer_id is not None:
                    user32.KillTimer(hwnd_arg, hook._timer_id)
                    hook._timer_id = None
            return user32.CallWindowProcW(
                hook._original_wndproc, hwnd_arg, msg, wparam, lparam,
            )

        self._c_wndproc = WNDPROC(_wndproc)
        self._original_wndproc = user32.GetWindowLongPtrW(
            hwnd, self._GWLP_WNDPROC,
        )
        user32.SetWindowLongPtrW(
            hwnd, self._GWLP_WNDPROC,
            ctypes.cast(self._c_wndproc, ctypes.c_void_p),
        )

    def uninstall(self) -> None:
        self._sizing = False
        if self._hwnd is None:
            return
        try:
            import ctypes
            user32 = ctypes.windll.user32
            if self._timer_id is not None:
                user32.KillTimer(self._hwnd, self._timer_id)
                self._timer_id = None
            if self._original_wndproc is not None:
                user32.SetWindowLongPtrW(
                    self._hwnd, self._GWLP_WNDPROC,
                    self._original_wndproc,
                )
                self._original_wndproc = None
        except Exception:
            pass


def _prepare_environment() -> None:
    # Business work runs outside the GUI thread.  Using Qt Quick's basic render
    # loop therefore avoids the GUI/render-thread synchronization stalls that
    # otherwise become visible while native windows are resized on Windows and
    # macOS.  Keep the graphics backend native (D3D/ANGLE/Metal); only the
    # scheduling model is made deterministic.  An explicit environment value
    # remains available as a diagnostics escape hatch.
    os.environ.setdefault("QSG_RENDER_LOOP", "basic")
    os.environ.setdefault("QML_DISABLE_DISK_CACHE", "0")
    # On Windows, prefer D3D11 over D3D12 for the RHI backend.  D3D12 can
    # cause micro-stutter during live resize when the DWM compositor and the
    # render thread compete for the same swap-chain.  D3D11 presents through
    # the legacy flip model which is more predictable under windowed resize.
    if sys.platform == "win32":
        os.environ.setdefault("QSG_RHI_BACKEND", "d3d11")


def main() -> int:
    _prepare_environment()
    install_crash_logging()
    set_windows_app_identity()

    from .compat import (
        QApplication,
        QCoreApplication,
        QFontDatabase,
        QQmlApplicationEngine,
        QTimer,
        QUrl,
        QT_API,
        QT_MAJOR,
        application_attribute,
        delete_qobject,
    )
    from .controller import AppController

    if QT_MAJOR == 5:
        for name in ("AA_EnableHighDpiScaling", "AA_UseHighDpiPixmaps"):
            attribute = application_attribute(name)
            if attribute is not None:
                QCoreApplication.setAttribute(attribute, True)
    share_contexts = application_attribute("AA_ShareOpenGLContexts")
    if share_contexts is not None:
        QCoreApplication.setAttribute(share_contexts, True)

    # Qt 5.15 calls its minimal Controls 2 style "Default"; Qt 6 renamed it
    # "Basic".  Both avoid native-widget relayout during a live resize.
    controls_style = "Default" if QT_MAJOR == 5 else "Basic"
    os.environ.setdefault("QT_QUICK_CONTROLS_STYLE", controls_style)
    app = QApplication(sys.argv)
    app.setApplicationName("HRToolkit")
    app.setApplicationDisplayName("HR Workbench")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("HRToolkit")
    system_font_enum = getattr(QFontDatabase, "SystemFont", None)
    general_font = (
        system_font_enum.GeneralFont
        if system_font_enum is not None
        else QFontDatabase.GeneralFont
    )
    app.setFont(QFontDatabase.systemFont(general_font))

    controller = AppController()
    engine = QQmlApplicationEngine()
    engine.rootContext().setContextProperty("controller", controller)
    qml_path = Path(__file__).resolve().parent / "qml" / "Main.qml"
    engine.load(QUrl.fromLocalFile(str(qml_path)))
    if not engine.rootObjects():
        runlog.log_line(f"Qt Quick 主界面加载失败：{qml_path}")
        controller.close()
        return 1
    # On Windows, install a native event hook that keeps the render pipeline
    # alive during the modal sizing loop.  Without this, dragging the window
    # border freezes all QML content until the user releases the mouse.
    sizing_hook = _WindowsSizingHook()
    if sys.platform == "win32":
        try:
            root_win_id = int(engine.rootObjects()[0].winId())
            sizing_hook.install(app, root_win_id)
            runlog.log_line("Windows native resize hook installed")
        except Exception as exc:
            runlog.log_line(f"Windows resize hook unavailable: {exc}")
    runlog.log_line(
        f"HR Workbench v{__version__} Qt Quick 启动（{sys.platform}，{QT_API}，"
        f"渲染循环 {os.environ.get('QSG_RENDER_LOOP', 'auto')}）"
    )
    smoke_tool = str(os.environ.get("HR_TOOLKIT_QT_SMOKE_TOOL", "")).strip()
    if smoke_tool:
        controller.selectTool(smoke_tool)
    smoke_screenshot = str(
        os.environ.get("HR_TOOLKIT_QT_SMOKE_SCREENSHOT", "")
    ).strip()
    smoke_size = str(os.environ.get("HR_TOOLKIT_QT_SMOKE_SIZE", "")).strip()
    smoke_exit = str(os.environ.get("HR_TOOLKIT_QT_SMOKE_EXIT_MS", "")).strip()
    if smoke_size:
        try:
            width_text, height_text = smoke_size.lower().split("x", 1)
            smoke_width = max(760, int(width_text))
            smoke_height = max(600, int(height_text))
            root_window = engine.rootObjects()[0]
            root_window.setWidth(smoke_width)
            root_window.setHeight(smoke_height)
        except (AttributeError, TypeError, ValueError):
            runlog.log_line(f"忽略无效 Qt smoke 尺寸：{smoke_size}")
    if smoke_screenshot:
        def capture_and_quit() -> None:
            try:
                roots = engine.rootObjects()
                if roots:
                    root_window = roots[0]
                    direct_grab = getattr(root_window, "grabWindow", None)
                    if callable(direct_grab):
                        image = direct_grab()
                    else:
                        # PySide6 may expose ApplicationWindow as QWindow rather
                        # than QQuickWindow. QScreen works for both Qt 5 and 6.
                        image = root_window.screen().grabWindow(
                            int(root_window.winId())
                        )
                    if not image.save(smoke_screenshot):
                        runlog.log_line(
                            f"Qt smoke 截图保存失败：{smoke_screenshot}"
                        )
            except Exception as exc:
                runlog.log_line(f"Qt smoke 截图失败：{exc}")
            finally:
                app.quit()

        QTimer.singleShot(max(250, int(smoke_exit or "600")), capture_and_quit)
    elif smoke_exit.isdigit():
        QTimer.singleShot(max(1, int(smoke_exit)), app.quit)
    app.aboutToQuit.connect(controller.close)
    app.aboutToQuit.connect(sizing_hook.uninstall)
    execute = getattr(app, "exec", None) or app.exec_
    exit_code = int(execute())
    sizing_hook.uninstall()
    # Destroy QML roots before their context is cleared.  Besides avoiding
    # noisy null-binding warnings, this deterministically releases scene-graph
    # textures and GPU resources during repeated packaged-app smoke tests.
    for root_object in engine.rootObjects():
        delete_qobject(root_object)
    return exit_code


__all__ = ["main"]
