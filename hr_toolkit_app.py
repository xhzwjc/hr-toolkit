from __future__ import annotations

import sys
from hr_toolkit.launcher import (
    DesktopRuntimeUnavailable,
    _apply_software_rendering_flags,
    _is_qt_runtime_import_error,
    _qt_install_command,
    _run_desktop,
    main,
)

if __name__ == "__main__":
    raise SystemExit(main())
