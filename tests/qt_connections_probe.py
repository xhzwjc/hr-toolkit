"""Exercise the production QML signal handlers in an isolated process.

This intentionally uses the real AppController and the Connections block read
from Main.qml. Pure Python controller tests do not execute Qt's QML connection
compiler, where the frozen PySide2/Qt 5.15.2 stack previously crashed.
"""

from __future__ import annotations

import faulthandler
from pathlib import Path
import sys


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def main() -> int:
    faulthandler.enable()
    from hr_toolkit.gui_qt.compat import (
        QCoreApplication,
        QQmlApplicationEngine,
        QT_API,
        delete_qobject,
    )
    from hr_toolkit.gui_qt.controller import AppController

    app = QCoreApplication([])
    controller = AppController()
    controller._save_workspace_preferences = lambda: None
    source = (REPO_ROOT / "hr_toolkit/gui_qt/qml/Main.qml").read_text(encoding="utf-8")
    # Use the actual final Connections block, not a hand-maintained copy.
    start = source.rindex("\n    Connections {")
    handlers = source[start:source.rindex("\n}")].strip()
    qml = """
import QtQml 2.15
QtObject {
    id: probe
    property var calls: []
    function record(kind, values) { calls.push([kind, values]) }
    property QtObject notificationDialog: QtObject {
        function showMessage(title, message, level) {
            probe.record("notification", [title, message, level])
        }
    }
    property QtObject confirmationDialog: QtObject {
        property string title
        property string bodyText
        property string actionToken
        function open() { probe.record("confirmation", [title, bodyText, actionToken]) }
    }
    property QtObject createProjectDialog: QtObject {
        property string projectName
        property string projectParent
        function open() { probe.record("project", [projectName, projectParent]) }
    }
    property QtObject textInputDialog: QtObject {
        function request(title, prompt, initialValue, token) {
            probe.record("text", [title, prompt, initialValue, token])
        }
    }
    property QtObject updateProgressDialog: QtObject {
        property bool opened: false
        function open() { opened = true; probe.record("update-open", []) }
        function close() { opened = false; probe.record("update-close", []) }
    }
    property Connections controllerHandlers: HANDLERS
}
""".replace("HANDLERS", handlers)
    expected = [
        ["notification", ["提醒标题", "完整消息", "warning"]],
        ["confirmation", ["确认标题", "确认内容", "opaque-token:123"]],
        ["project", ["工作项目", "E:\\资料\\工作目录"]],
        ["text", ["输入标题", "输入提示", "初始值", "text-token:456"]],
        ["update-open", []],
        ["update-close", []],
    ]
    # Rebuilding connections repeatedly exposes nondeterministic native crashes
    # without depending on a visible desktop or a user's saved project.
    for iteration in range(20):
        engine = QQmlApplicationEngine()
        engine.rootContext().setContextProperty("controller", controller)
        try:
            engine.loadData(qml.encode("utf-8"))
            roots = engine.rootObjects()
            if len(roots) != 1:
                raise AssertionError("Production Connections probe did not load")
            controller.notificationRequested.emit(*expected[0][1])
            controller.confirmationRequested.emit(*expected[1][1])
            controller.projectCreationRequested.emit(*expected[2][1])
            controller.textInputRequested.emit(*expected[3][1])
            controller._update_busy = True
            controller._update_status = "正在下载更新"
            controller.updateChanged.emit()
            controller._update_busy = False
            controller.updateChanged.emit()
            app.processEvents()
            observed = roots[0].property("calls")
            if hasattr(observed, "toVariant"):
                observed = observed.toVariant()
            if observed != expected:
                raise AssertionError(
                    f"Signal arguments changed on iteration {iteration}: {observed!r}"
                )
        finally:
            delete_qobject(engine)
    controller.close()
    print(f"{QT_API} production QML connections: 20/20 passed", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
