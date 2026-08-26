"""Keep modern-host API Set forwarders out of the frozen Win7 payload."""

from __future__ import annotations

import os

from PyInstaller.depend import dylib


_original_include_library = dylib.include_library


def _include_win7_library(libname) -> bool:
    name = os.path.basename(os.fspath(libname)).casefold()
    if name.startswith(("api-ms-win-", "ext-ms-win-")):
        # Pinned down-level UCRT files are explicit --add-binary inputs. This
        # filter applies only to recursively discovered host dependencies.
        return False
    return _original_include_library(libname)


dylib.include_library = _include_win7_library
