from __future__ import annotations

import sys

from hr_toolkit.cli import main as cli_main
from hr_toolkit.gui import main as gui_main


if __name__ == "__main__":
    if len(sys.argv) > 1:
        raise SystemExit(cli_main())
    gui_main()
