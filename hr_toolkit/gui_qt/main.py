"""Qt Quick application bootstrap."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from hr_toolkit import __version__, runlog
from hr_toolkit.desktop_helpers import install_crash_logging, set_windows_app_identity

from .live_resize import LiveResizeUpdater, WindowsResizeBackdrop
from .smoke import mark_stage


def _prepare_environment() -> None:
    software_requested = (
        os.environ.get("HR_TOOLKIT_QT_SOFTWARE_RENDER", "").strip().lower()
        in {"1", "true", "yes", "on"}
        or os.environ.get("HR_TOOLKIT_SOFTWARE_RENDER", "").strip().lower()
        in {"1", "true", "yes", "on"}
    )
    if software_requested:
        os.environ.setdefault("QT_QUICK_BACKEND", "software")
        os.environ.setdefault("QSG_RENDER_LOOP", "basic")

    # Windows must retain Qt's platform/driver selection.  Modern Qt normally
    # chooses its threaded scene-graph renderer there, while the Win7 Qt 5 lane
    # can select a compatible path for the actual GPU.  The production Qt
    # 6.6.3 macOS benchmark has measurably steadier frame pacing with the basic
    # loop, so keep that verified platform-specific policy.  Explicit values
    # remain available for diagnostics on every platform.
    if sys.platform == "darwin":
        os.environ.setdefault("QSG_RENDER_LOOP", "basic")
    elif sys.platform.startswith("win"):
        # QWindow::requestUpdate() otherwise allows up to roughly 5 ms of idle
        # time before delivery.  A zero delay reduces the exposed-background
        # interval during native edge/corner resizing without changing the
        # graphics backend or disabling GPU acceleration.
        os.environ.setdefault("QT_QPA_UPDATE_IDLE_TIME", "0")
    os.environ.setdefault("QML_DISABLE_DISK_CACHE", "0")


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
    mark_stage("qt-application")
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
    application_font = QFontDatabase.systemFont(general_font)
    # Match the original Tk interface's platform typography.  The locked
    # Win7/PySide2 lane uses the font shipped with Windows 7, while modern
    # Windows uses the UI-tuned family.  Missing families still fall back
    # through Qt normally, so this does not add a font dependency.
    if sys.platform == "darwin":
        application_font.setFamily("PingFang SC")
    elif sys.platform.startswith("win"):
        application_font.setFamily(
            "Microsoft YaHei" if QT_MAJOR == 5 else "Microsoft YaHei UI"
        )
    app.setFont(application_font)

    mark_stage("qt-controller")
    controller = AppController()
    engine = QQmlApplicationEngine()
    engine.rootContext().setContextProperty("controller", controller)
    qml_path = Path(__file__).resolve().parent / "qml" / "Main.qml"
    mark_stage("qt-qml-load")
    engine.load(QUrl.fromLocalFile(str(qml_path)))
    mark_stage("qt-qml-loaded")
    if not engine.rootObjects():
        runlog.log_line(f"Qt Quick 主界面加载失败：{qml_path}")
        controller.close()
        return 1
    root_window = engine.rootObjects()[0]
    live_resize_updater = LiveResizeUpdater(
        root_window if sys.platform.startswith("win") else None
    )
    resize_backdrop = WindowsResizeBackdrop.install(root_window)
    resize_helpers_closed = False

    def close_resize_helpers() -> None:
        nonlocal resize_helpers_closed
        if resize_helpers_closed:
            return
        resize_helpers_closed = True
        live_resize_updater.close()
        resize_backdrop.close()

    runlog.log_line(
        f"HR Workbench v{__version__} Qt Quick 启动（{sys.platform}，{QT_API}，"
        f"后端 {os.environ.get('QT_QUICK_BACKEND', 'hardware')}，"
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
    app.aboutToQuit.connect(close_resize_helpers)
    execute = getattr(app, "exec", None) or app.exec_
    mark_stage("qt-event-loop")
    exit_code = int(execute())
    mark_stage("qt-cleanup")
    close_resize_helpers()
    # Destroy QML roots before their context is cleared.  Besides avoiding
    # noisy null-binding warnings, this deterministically releases scene-graph
    # textures and GPU resources during repeated packaged-app smoke tests.
    for root_object in engine.rootObjects():
        delete_qobject(root_object)
    mark_stage("qt-cleanup-complete")
    return exit_code


__all__ = ["main"]
