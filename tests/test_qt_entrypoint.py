from __future__ import annotations

import os
import unittest
from unittest.mock import patch

import hr_toolkit_app


class QtEntrypointTests(unittest.TestCase):
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
