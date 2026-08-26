from __future__ import annotations

from pathlib import Path


def path_is_relative_to(path: Path, root: Path) -> bool:
    """Python 3.8 compatible equivalent of ``Path.is_relative_to``."""
    try:
        path.relative_to(root)
    except ValueError:
        return False
    return True
