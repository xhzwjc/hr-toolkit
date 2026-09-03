"""Unified application launcher for CLI commands and WPF IPC server."""

from __future__ import annotations

import multiprocessing
import sys
from typing import Sequence

from hr_toolkit.runtime_checks import run_headless_command


def _run_ipc() -> int:
    from hr_toolkit.ipc_server import main as ipc_main

    return int(ipc_main())


def main(argv: Sequence[str] | None = None) -> int:
    multiprocessing.freeze_support()
    args = list(sys.argv[1:] if argv is None else argv)

    if args == ["--ipc"] or args == ["ipc"]:
        return _run_ipc()

    headless_result = run_headless_command(args)
    if headless_result is not None:
        return int(headless_result)

    from hr_toolkit.cli import main as cli_main

    if len(args) > 0:
        old_argv = sys.argv
        try:
            sys.argv = [old_argv[0]] + args
            return int(cli_main())
        finally:
            sys.argv = old_argv

    # When launched with no arguments, display friendly guidance and CLI help
    print(
        "HR Toolkit 业务引擎 (Python Core)\n"
        "桌面客户端由原生 WPF 前端 (src/HRToolkit.Wpf) 驱动。\n"
        "如需启动 IPC 守护进程供 WPF 客户端调用，请传入: --ipc\n"
        "如需在命令行直接使用工具，请参考下方帮助：\n",
        file=sys.stderr,
    )
    old_argv = sys.argv
    try:
        sys.argv = [old_argv[0], "--help"]
        return int(cli_main())
    finally:
        sys.argv = old_argv


if __name__ == "__main__":
    raise SystemExit(main())

