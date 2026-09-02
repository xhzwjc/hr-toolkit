from __future__ import annotations

import os
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import hr_toolkit_app
from hr_toolkit.gui_qt.live_resize import (
    LiveResizeUpdater,
    WindowsResizeBackdrop,
    _windows_colorref,
)
from hr_toolkit.gui_qt.main import _prepare_environment


class QtEntrypointTests(unittest.TestCase):
    def test_windows_qt_environment_keeps_platform_render_defaults(self) -> None:
        with patch.object(sys, "platform", "win32"):
            with patch.dict(os.environ, {}, clear=True):
                _prepare_environment()
                self.assertNotIn("QSG_RENDER_LOOP", os.environ)
                self.assertNotIn("QSG_RHI_BACKEND", os.environ)
                self.assertEqual(os.environ["QML_DISABLE_DISK_CACHE"], "0")
                self.assertEqual(os.environ["QT_QPA_UPDATE_IDLE_TIME"], "0")

    def test_windows_qt_environment_preserves_explicit_update_delay(self) -> None:
        with patch.object(sys, "platform", "win32"):
            with patch.dict(
                os.environ,
                {"QT_QPA_UPDATE_IDLE_TIME": "4"},
                clear=True,
            ):
                _prepare_environment()
                self.assertEqual(os.environ["QT_QPA_UPDATE_IDLE_TIME"], "4")

    def test_live_resize_requests_coalescible_quick_window_updates(self) -> None:
        class FakeSignal:
            def __init__(self) -> None:
                self.callback = None

            def connect(self, callback) -> None:
                self.callback = callback

            def disconnect(self, callback) -> None:
                if self.callback == callback:
                    self.callback = None

            def emit(self) -> None:
                if self.callback is not None:
                    self.callback()

        class FakeWindow:
            def __init__(self) -> None:
                self.widthChanged = FakeSignal()
                self.heightChanged = FakeSignal()
                self.update_calls = 0
                self.persistent = []

            def setPersistentGraphics(self, value) -> None:
                self.persistent.append(("graphics", value))

            def setPersistentSceneGraph(self, value) -> None:
                self.persistent.append(("scene", value))

            def update(self) -> None:
                self.update_calls += 1

        window = FakeWindow()
        updater = LiveResizeUpdater(window)
        window.widthChanged.emit()
        window.heightChanged.emit()

        self.assertEqual(window.update_calls, 2)
        self.assertEqual(
            window.persistent,
            [("graphics", True), ("scene", True)],
        )
        updater.close()
        window.widthChanged.emit()
        self.assertEqual(window.update_calls, 2)

    def test_windows_resize_backdrop_is_inert_off_windows(self) -> None:
        with patch.object(sys, "platform", "darwin"):
            backdrop = WindowsResizeBackdrop.install(object())

        self.assertFalse(backdrop.active)
        self.assertEqual(_windows_colorref(0xF7, 0xF5, 0xF1), 0xF1F5F7)

    def test_macos_keeps_the_benchmarked_basic_render_loop(self) -> None:
        with patch.object(sys, "platform", "darwin"):
            with patch.dict(os.environ, {}, clear=True):
                _prepare_environment()
                self.assertEqual(os.environ["QSG_RENDER_LOOP"], "basic")
                self.assertNotIn("QSG_RHI_BACKEND", os.environ)

    def test_qt_environment_preserves_explicit_graphics_overrides(self) -> None:
        with patch.dict(
            os.environ,
            {
                "QSG_RENDER_LOOP": "threaded",
                "QSG_RHI_BACKEND": "opengl",
                "QML_DISABLE_DISK_CACHE": "1",
            },
            clear=True,
        ):
            with patch.object(sys, "platform", "darwin"):
                _prepare_environment()
                self.assertEqual(os.environ["QSG_RENDER_LOOP"], "threaded")
                self.assertEqual(os.environ["QSG_RHI_BACKEND"], "opengl")
                self.assertEqual(os.environ["QML_DISABLE_DISK_CACHE"], "1")

    def test_qt_entrypoint_does_not_subclass_the_native_window(self) -> None:
        source_path = (
            Path(__file__).resolve().parents[1]
            / "hr_toolkit"
            / "gui_qt"
            / "main.py"
        )
        source = source_path.read_text(encoding="utf-8")

        self.assertNotIn("_WindowsSizingHook", source)
        self.assertNotIn("SetWindowLongPtrW", source)
        self.assertNotIn("WM_PAINT", source)

    def test_live_resize_keeps_layout_breakpoints_and_workspace_stable(self) -> None:
        qml_path = (
            Path(__file__).resolve().parents[1]
            / "hr_toolkit"
            / "gui_qt"
            / "qml"
            / "Main.qml"
        )
        source = qml_path.read_text(encoding="utf-8")

        # Responsive breakpoints use settled width to avoid layout churn
        self.assertIn("settledWidth <= 860", source)
        self.assertIn("settledWidth >= 980", source)
        self.assertNotIn("Behavior on Layout.preferredWidth", source)
        self.assertIn(
            "anchors.leftMargin: root.compactSidebar ? 12 : 28",
            source,
        )
        self.assertIn("readonly property int contentMaxWidth: 820", source)
        self.assertIn('objectName: "workspaceButton"', source)
        self.assertIn('objectName: "workspaceButtonMouse"', source)
        self.assertIn('objectName: "runButton"', source)
        self.assertIn('text: "项\\n目\\n文\\n件"', source)
        self.assertIn("enabled: controller.hasProject\n                    cursorShape", source)
        self.assertIn("enabled: controller.busy || !controller.workspaceBusy", source)
        self.assertIn("readonly property int preferredWindowWidth: 1400", source)
        self.assertIn("readonly property int preferredWindowHeight: 780", source)
        self.assertIn("Math.min(Screen.width, Screen.desktopAvailableWidth)", source)
        self.assertIn("Math.min(Screen.height, Screen.desktopAvailableHeight)", source)
        self.assertIn("currentScreenAvailableWidth - initialWindowMargin", source)
        self.assertIn("currentScreenAvailableHeight - initialWindowMargin", source)
        # The open project-files panel must follow every native resize frame;
        # dialogs that are not being dragged can keep settled dimensions.
        self.assertIn("x: root.width - width", source)
        self.assertIn("width: Math.min(340, root.width - 24)", source)
        self.assertIn("height: root.height", source)
        self.assertNotIn("width: Math.min(340, root.settledWidth - 24)", source)
        self.assertNotIn("Math.max(360, root.width * 0.38)", source)
        self.assertIn("enter: Transition {}", source)
        self.assertIn("exit: Transition {}", source)
        # Settled-width timer defers layout during native resize
        self.assertIn("settledWidth", source)
        self.assertIn("settleTimer", source)
        self.assertIn('objectName: "sidebarProjectCard"', source)
        self.assertIn('readonly property bool showLegacyHistoryEntry: false', source)
        self.assertIn('objectName: "runLogIconButton"', source)
        self.assertIn('ToolTip.text: "打开运行日志"', source)
        self.assertIn("model: controller.tutorialGroups", source)
        self.assertIn("workspaceList.positionViewAtIndex(safeRow, ListView.Contain)", source)

    def test_qt_is_the_default_desktop_renderer(self) -> None:
        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("HR_TOOLKIT_RENDERER", None)
            with patch("hr_toolkit.gui_qt.main", return_value=0) as qt_main:
                self.assertEqual(hr_toolkit_app._run_desktop(), 0)
        qt_main.assert_called_once_with()

    def test_legacy_renderer_remains_an_explicit_recovery_option(self) -> None:
        legacy_main = Mock()
        legacy_module = types.ModuleType("hr_toolkit.gui")
        legacy_module.main = legacy_main
        with patch.dict(os.environ, {"HR_TOOLKIT_RENDERER": "legacy-tk"}):
            with patch.dict(sys.modules, {"hr_toolkit.gui": legacy_module}):
                self.assertEqual(hr_toolkit_app._run_desktop(), 0)
        legacy_main.assert_called_once_with()

    def test_missing_qt_runtime_never_silently_starts_legacy_renderer(self) -> None:
        legacy_main = Mock()
        legacy_module = types.ModuleType("hr_toolkit.gui")
        legacy_module.main = legacy_main
        missing_qt = ImportError("No module named 'shiboken6'", name="shiboken6")

        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("HR_TOOLKIT_RENDERER", None)
            with patch.dict(sys.modules, {"hr_toolkit.gui": legacy_module}):
                with patch("hr_toolkit.gui_qt.main", side_effect=missing_qt):
                    with self.assertRaises(
                        hr_toolkit_app.DesktopRuntimeUnavailable
                    ) as raised:
                        hr_toolkit_app._run_desktop()

        self.assertIn("shiboken6", str(raised.exception))
        self.assertIn("requirements-gui.txt", str(raised.exception))
        legacy_main.assert_not_called()

    def test_non_qt_import_failure_is_not_misreported_as_missing_qt(self) -> None:
        missing_core = ImportError("No module named 'openpyxl'", name="openpyxl")

        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("HR_TOOLKIT_RENDERER", None)
            with patch("hr_toolkit.gui_qt.main", side_effect=missing_core):
                with self.assertRaises(ImportError) as raised:
                    hr_toolkit_app._run_desktop()

        self.assertIs(raised.exception, missing_core)

    def test_win7_source_install_uses_the_locked_qt5_constraints(self) -> None:
        with patch.object(hr_toolkit_app.sys, "platform", "win32"):
            with patch.object(hr_toolkit_app.sys, "version_info", (3, 8, 10)):
                command = hr_toolkit_app._qt_install_command()

        self.assertIn("requirements-gui.txt", command)
        self.assertIn("constraints/python38-win7.txt", command)


if __name__ == "__main__":
    unittest.main()
