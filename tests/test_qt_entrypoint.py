from __future__ import annotations

import os
import unittest
from pathlib import Path
from unittest.mock import patch

import hr_toolkit_app
from hr_toolkit.gui_qt.main import _prepare_environment


class QtEntrypointTests(unittest.TestCase):
    def test_qt_environment_uses_basic_render_loop_by_default(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            _prepare_environment()
            self.assertEqual(os.environ["QSG_RENDER_LOOP"], "basic")
            self.assertEqual(os.environ["QML_DISABLE_DISK_CACHE"], "0")

    def test_qt_environment_preserves_explicit_render_loop_override(self) -> None:
        with patch.dict(
            os.environ,
            {"QSG_RENDER_LOOP": "threaded", "QML_DISABLE_DISK_CACHE": "1"},
            clear=True,
        ):
            _prepare_environment()
            self.assertEqual(os.environ["QSG_RENDER_LOOP"], "threaded")
            self.assertEqual(os.environ["QML_DISABLE_DISK_CACHE"], "1")

    def test_live_resize_keeps_layout_breakpoints_and_workspace_stable(self) -> None:
        qml_path = (
            Path(__file__).resolve().parents[1]
            / "hr_toolkit"
            / "gui_qt"
            / "qml"
            / "Main.qml"
        )
        source = qml_path.read_text(encoding="utf-8")

        self.assertIn("width <= 860", source)
        self.assertIn("width >= 980", source)
        self.assertNotIn("Behavior on Layout.preferredWidth", source)
        self.assertIn("anchors.leftMargin: 28", source)
        self.assertIn("width: Math.min(440, root.width - 24)", source)
        self.assertNotIn("Math.max(360, root.width * 0.38)", source)
        self.assertIn("enter: Transition {}", source)
        self.assertIn("exit: Transition {}", source)

    def test_qt_is_the_default_desktop_renderer(self) -> None:
        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop("HR_TOOLKIT_RENDERER", None)
            with patch("hr_toolkit.gui_qt.main", return_value=0) as qt_main:
                self.assertEqual(hr_toolkit_app._run_desktop(), 0)
        qt_main.assert_called_once_with()

    def test_legacy_renderer_remains_an_explicit_recovery_option(self) -> None:
        with patch.dict(os.environ, {"HR_TOOLKIT_RENDERER": "legacy-tk"}):
            with patch("hr_toolkit.gui.main", return_value=None) as legacy_main:
                self.assertEqual(hr_toolkit_app._run_desktop(), 0)
        legacy_main.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
