"""Qt Quick application bootstrap."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from hr_toolkit import __version__, runlog
from hr_toolkit.desktop_helpers import install_crash_logging, set_windows_app_identity


def _prepare_environment() -> None:
    # Keep native GPU selection.  Qt/ANGLE can choose the supported backend on
    # Win7 while Qt 6 can use the modern RHI backend on Win11 and macOS.
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
    runlog.log_line(
        f"HR Workbench v{__version__} Qt Quick 启动（{sys.platform}，{QT_API}）"
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
            roots = engine.rootObjects()
            if roots:
                image = roots[0].grabWindow()
                image.save(smoke_screenshot)
            app.quit()

        QTimer.singleShot(max(250, int(smoke_exit or "600")), capture_and_quit)
    elif smoke_exit.isdigit():
        QTimer.singleShot(max(1, int(smoke_exit)), app.quit)
    app.aboutToQuit.connect(controller.close)
    execute = getattr(app, "exec", None) or app.exec_
    exit_code = int(execute())
    # Destroy QML roots before their context is cleared.  Besides avoiding
    # noisy null-binding warnings, this deterministically releases scene-graph
    # textures and GPU resources during repeated packaged-app smoke tests.
    for root_object in engine.rootObjects():
        delete_qobject(root_object)
    return exit_code


__all__ = ["main"]
