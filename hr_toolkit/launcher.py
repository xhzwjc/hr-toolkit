"""Unified application launcher for GUI desktop and CLI commands."""

from __future__ import annotations

import multiprocessing
import os
import sys
from typing import Sequence

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
    try:
        from hr_toolkit.gui_qt import main as qt_gui_main

        return int(qt_gui_main())
    except ImportError as exc:
        if not _is_qt_runtime_import_error(exc):
            raise
        from hr_toolkit import runlog

        missing = getattr(exc, "name", "") or str(exc)
        message = (
            "Qt 桌面运行库不可用。"
            f"缺少或无法导入：{missing}。请执行：{_qt_install_command()}"
        )
        runlog.log_line(message)
        raise DesktopRuntimeUnavailable(message) from exc


def _apply_software_rendering_flags(args: Sequence[str]) -> list[str]:
    software_flags = {"--software-rendering", "--software-render"}
    if any(arg.casefold() in software_flags for arg in args):
        os.environ["HR_TOOLKIT_QT_SOFTWARE_RENDER"] = "1"
        return [arg for arg in args if arg.casefold() not in software_flags]
    return list(args)


def main(argv: Sequence[str] | None = None) -> int:
    multiprocessing.freeze_support()
    raw_args = list(sys.argv[1:] if argv is None else argv)
    filtered_args = _apply_software_rendering_flags(raw_args)

    if filtered_args == ["--qt-smoke-test"]:
        from hr_toolkit.gui_qt.smoke import run as run_qt_smoke

        return int(run_qt_smoke())

    headless_result = run_headless_command(filtered_args)
    if headless_result is not None:
        return int(headless_result)

    if len(filtered_args) > 0:
        from hr_toolkit.cli import main as cli_main

        old_argv = sys.argv
        try:
            sys.argv = [old_argv[0]] + filtered_args
            return int(cli_main())
        finally:
            sys.argv = old_argv

    try:
        return _run_desktop()
    except DesktopRuntimeUnavailable as exc:
        if sys.stderr is not None:
            print(str(exc), file=sys.stderr, flush=True)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
