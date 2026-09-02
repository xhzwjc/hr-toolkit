"""Qt Quick desktop front end for HR Toolkit.

The package intentionally imports no Qt modules at package-import time.  Core
tests and command-line tools therefore remain usable when a desktop Qt runtime
is not installed.
"""

from __future__ import annotations


def main() -> int:
    from .main import main as qt_main

    return qt_main()


__all__ = ["main"]
