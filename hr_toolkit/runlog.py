"""运行日志：记录每次工具运行的关键节点和异常，便于远程排查问题。

设计约定：
- 只记录文件名、文件大小、耗时、统计数字和异常堆栈，绝不记录表格内容
  （HR 文件包含身份证号、工资等敏感数据，日志必须可以放心外发）。
- 日志失败绝不能影响业务，所有写入都是尽力而为。
- 与更新日志 HRToolkit_update.log 放在同一位置，方便一次性收集。
"""

from __future__ import annotations

import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from types import TracebackType
from typing import Any

from hr_toolkit.app_update import current_app_dir, trim_log_file

RUN_LOG_FILE = "HRToolkit_app.log"
RUN_LOG_ENV = "HR_TOOLKIT_APP_LOG"

_write_lock = threading.Lock()


def run_log_path() -> Path:
    env_path = os.environ.get(RUN_LOG_ENV, "").strip()
    if env_path:
        return Path(env_path)
    if getattr(sys, "frozen", False):
        # 与 HRToolkit_update.log 同级：HRToolkit 程序目录的上一级
        return current_app_dir().parent / RUN_LOG_FILE
    return Path.cwd() / RUN_LOG_FILE


def log_line(text: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        path = run_log_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        with _write_lock:
            trim_log_file(path)
            with path.open("a", encoding="utf-8") as handle:
                handle.write(f"[{timestamp}] {text}\n")
    except Exception:
        pass


def log_exception(context: str, exc: BaseException, tb: TracebackType | None = None) -> None:
    detail = "".join(traceback.format_exception(type(exc), exc, tb or exc.__traceback__)).rstrip()
    log_line(f"{context}\n{detail}")


def describe_value(value: Any) -> str:
    """把运行参数渲染成日志可读的摘要：路径只记名字和大小，不展开内容。"""
    if isinstance(value, Path):
        return _describe_path(value)
    if isinstance(value, (list, tuple)):
        names = [describe_value(item) for item in list(value)[:5]]
        if len(value) > 5:
            names.append(f"等共{len(value)}项")
        return "、".join(names) if names else "空"
    if value is None:
        return "无"
    return str(value)


def describe_call(args: tuple[Any, ...], kwargs: dict[str, Any]) -> str:
    parts = [describe_value(value) for value in args if value is not None]
    parts.extend(f"{key}={describe_value(value)}" for key, value in kwargs.items() if value is not None)
    return "；".join(parts)


def _describe_path(value: Path) -> str:
    try:
        if value.is_file():
            size_bytes = value.stat().st_size
            if size_bytes >= 1024 * 1024:
                size = f"{size_bytes / 1024 / 1024:.1f}MB"
            else:
                size = f"{max(size_bytes, 1) / 1024:.0f}KB"
            return f"{value.name}({size})"
        if value.is_dir():
            return f"{value.name}/"
    except OSError:
        pass
    return value.name or str(value)
