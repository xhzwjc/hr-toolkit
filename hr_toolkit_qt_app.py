from __future__ import annotations

import multiprocessing


if __name__ == "__main__":
    multiprocessing.freeze_support()
    from hr_toolkit.gui_qt import main

    raise SystemExit(main())
