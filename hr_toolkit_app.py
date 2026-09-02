from __future__ import annotations

import multiprocessing
import os
import sys
from pathlib import Path

from hr_toolkit.runtime_checks import run_headless_command


def _run_desktop() -> int:
    renderer = os.environ.get("HR_TOOLKIT_RENDERER", "qt").strip().casefold()
    if renderer in {"tk", "legacy", "legacy-tk"}:
        from hr_toolkit.gui import main as legacy_gui_main

        legacy_gui_main()
        return 0
    try:
        from hr_toolkit.gui_qt import main as qt_gui_main

        return int(qt_gui_main())
    except ImportError:
        # Keep source checkouts usable when only the core dependencies were
        # installed.  Release builds include Qt and their packaged smoke test
        # fails if that runtime is missing, so this is not a silent build gate.
        from hr_toolkit import runlog
        from hr_toolkit.gui import main as legacy_gui_main

        runlog.log_line("Qt 桌面运行库未安装，临时使用兼容界面。")
        legacy_gui_main()
        return 0


if __name__ == "__main__":
    multiprocessing.freeze_support()
    if sys.argv[1:] == ["--qt-smoke-test"]:
        os.environ.setdefault("HR_TOOLKIT_SKIP_UPDATE", "1")
        os.environ.setdefault("HR_TOOLKIT_QT_SMOKE_EXIT_MS", "800")
        sys.argv = [sys.argv[0]]
        try:
            from hr_toolkit.gui_qt import main as qt_main

            result = int(qt_main())
        except Exception as exc:
            message = f"HRToolkit Qt smoke-test FAILED: {exc}"
            result = 1
        else:
            message = "HRToolkit Qt smoke-test OK" if result == 0 else f"HRToolkit Qt smoke-test FAILED: exit={result}"
        if sys.stdout is not None:
            print(message, flush=True)
        output = os.environ.get("HR_TOOLKIT_CHECK_OUTPUT", "").strip()
        if output:
            Path(output).write_text(message + "\n", encoding="utf-8")
        raise SystemExit(result)
    headless_result = run_headless_command(sys.argv[1:])
    if headless_result is not None:
        raise SystemExit(headless_result)
    if len(sys.argv) > 1:
        from hr_toolkit.cli import main as cli_main

        raise SystemExit(cli_main())
    raise SystemExit(_run_desktop())
