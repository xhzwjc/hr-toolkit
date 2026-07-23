from __future__ import annotations

import sys

from hr_toolkit.runtime_checks import run_headless_command


if __name__ == "__main__":
    headless_result = run_headless_command(sys.argv[1:])
    if headless_result is not None:
        raise SystemExit(headless_result)
    if len(sys.argv) > 1:
        from hr_toolkit.cli import main as cli_main

        raise SystemExit(cli_main())
    from hr_toolkit.gui import main as gui_main

    gui_main()
