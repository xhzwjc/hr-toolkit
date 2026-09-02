from __future__ import annotations

import multiprocessing
import os
import sys

from hr_toolkit.runtime_checks import run_headless_command


class DesktopRuntimeUnavailable(RuntimeError):
    """The production Qt desktop runtime could not be imported."""


def _qt_install_command() -> str:
    if sys.platform == "win32" and sys.version_info < (3, 9):
        return (
            "python -m pip install -r requirements-gui.txt "
            "-c constraints/python38-win7.txt"
        )
    if sys.version_info[:2] == (3, 12):
        return (
            "python -m pip install -r requirements-gui.txt "
            "-c constraints/python312-production.txt"
        )
    return "python -m pip install -r requirements-gui.txt"


def _is_qt_runtime_import_error(exc: ImportError) -> bool:
    module_name = str(getattr(exc, "name", "") or "").split(".", 1)[0]
    return module_name.casefold() in {
        "pyside2",
        "pyside6",
        "shiboken2",
        "shiboken6",
    }


def _run_desktop() -> int:
    renderer = os.environ.get("HR_TOOLKIT_RENDERER", "qt").strip().casefold()
    if renderer in {"tk", "legacy", "legacy-tk"}:
        from hr_toolkit.gui import main as legacy_gui_main

        legacy_gui_main()
        return 0
    try:
        from hr_toolkit.gui_qt import main as qt_gui_main

        return int(qt_gui_main())
    except ImportError as exc:
        if not _is_qt_runtime_import_error(exc):
            raise
        # Never disguise a missing or broken Qt runtime as a slow production
        # desktop by silently launching the legacy Tk renderer.  Packaged apps
        # contain Qt and fail their smoke gate if it is missing; source users
        # receive the exact installation command.  Tk remains an explicit
        # diagnostics/recovery option through HR_TOOLKIT_RENDERER=legacy-tk.
        from hr_toolkit import runlog

        missing = getattr(exc, "name", "") or str(exc)
        message = (
            "Qt 桌面运行库不可用，未启动旧版兼容界面。"
            f"缺少或无法导入：{missing}。请执行：{_qt_install_command()}"
        )
        runlog.log_line(message)
        raise DesktopRuntimeUnavailable(message) from exc


if __name__ == "__main__":
    multiprocessing.freeze_support()
    if sys.argv[1:] == ["--qt-smoke-test"]:
        from hr_toolkit.gui_qt.smoke import run as run_qt_smoke

        raise SystemExit(run_qt_smoke())
    headless_result = run_headless_command(sys.argv[1:])
    if headless_result is not None:
        raise SystemExit(headless_result)
    if len(sys.argv) > 1:
        from hr_toolkit.cli import main as cli_main

        raise SystemExit(cli_main())
    try:
        desktop_result = _run_desktop()
    except DesktopRuntimeUnavailable as exc:
        if sys.stderr is not None:
            print(str(exc), file=sys.stderr, flush=True)
        raise SystemExit(2)
    raise SystemExit(desktop_result)
