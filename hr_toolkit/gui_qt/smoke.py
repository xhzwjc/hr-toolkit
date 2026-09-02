"""Qt startup check with diagnostics that survive a native crash."""

from __future__ import annotations

from contextlib import ExitStack
import faulthandler
import os
from pathlib import Path
import sys
import traceback

from hr_toolkit.runtime_checks import CHECK_OUTPUT_ENV, _emit


def mark_stage(stage: str) -> None:
    if os.environ.get("HR_TOOLKIT_QT_SMOKE_EXIT_MS", "").strip():
        _emit(f"HRToolkit Qt smoke-test RUNNING: {stage}")


def run() -> int:
    os.environ.setdefault("HR_TOOLKIT_SKIP_UPDATE", "1")
    os.environ.setdefault("HR_TOOLKIT_QT_SMOKE_EXIT_MS", "800")
    sys.argv = [sys.argv[0]]
    with ExitStack() as cleanup:
        if not faulthandler.is_enabled():
            output = os.environ.get(CHECK_OUTPUT_ENV, "").strip()
            if output:
                # A windowed PyInstaller EXE has no stderr. Keep an independent
                # file open throughout Qt startup AND shutdown so Windows
                # access violations leave a Python/native-boundary traceback.
                stream = cleanup.enter_context(
                    Path(output + ".native.log").open("w", encoding="utf-8")
                )
            else:
                stream = sys.stderr
            if stream is not None:
                faulthandler.enable(file=stream, all_threads=True)
                cleanup.callback(faulthandler.disable)
        try:
            mark_stage("qt-import")
            from .main import main

            result = int(main())
        except Exception:
            _emit("HRToolkit Qt smoke-test FAILED\n" + traceback.format_exc().rstrip())
            return 1
        message = (
            "HRToolkit Qt smoke-test OK"
            if result == 0
            else f"HRToolkit Qt smoke-test FAILED: exit={result}"
        )
        _emit(message)
        return result
