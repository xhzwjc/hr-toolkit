"""Check actual Qt startup icons without touching the user's saved project."""
from __future__ import annotations

import base64
from pathlib import Path
import sys
import traceback
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def main() -> int:
    from hr_toolkit._icon_data import APP_ICON_PNGS_BASE64
    from hr_toolkit.gui_qt.compat import QApplication, QFileDialog, QPixmap, QTimer, delete_qobject
    from hr_toolkit.gui_qt.controller import AppController
    from hr_toolkit.gui_qt.main import main as desktop_main

    def verify():
        app = QApplication.instance()
        try:
            icon = app.windowIcon()
            assert not icon.isNull(), "Application icon is empty"
            for size, encoded in APP_ICON_PNGS_BASE64.items():
                expected = QPixmap()
                assert expected.loadFromData(base64.b64decode(encoded), "PNG")
                assert icon.pixmap(size, size).toImage() == expected.toImage(), size
            windows = [window for window in app.topLevelWindows() if window.isVisible()]
            assert windows, "Production window did not open"
            for window in windows:
                assert window.icon().cacheKey() == icon.cacheKey(), "Window lost app icon"
            dialog = QFileDialog()
            try:
                assert dialog.windowIcon().cacheKey() == icon.cacheKey(), "Dialog lost app icon"
            finally:
                delete_qobject(dialog)
            print("brand icon: 6 sizes, main window and dialog OK", flush=True)
            app.exit(0)
        except BaseException:
            traceback.print_exc()
            app.exit(1)

    def start(_controller):
        QTimer.singleShot(50, verify)

    with patch.object(AppController, "start", start), patch.object(
        AppController, "_save_workspace_preferences", lambda self: None
    ):
        return desktop_main()


if __name__ == "__main__":
    raise SystemExit(main())
