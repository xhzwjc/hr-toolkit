"""运行日志：记录每次工具运行的关键节点和异常，便于远程排查问题。

设计约定：
- 只记录文件名、文件大小、耗时、统计数字和异常堆栈，绝不记录表格内容
  （HR 文件包含身份证号、工资等敏感数据，日志必须可以放心外发）。
- 日志失败绝不能影响业务，所有写入都是尽力而为。
- 冻结程序写入用户日志目录，不向只读安装目录或 .app Bundle 写文件。
- 支持纯文本与结构化 JSON (通过环境变量 HR_TOOLKIT_LOG_JSON=1 开启) 两种模式。
"""

from __future__ import annotations

import json
import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from types import TracebackType
from typing import Any

RUN_LOG_FILE = "HRToolkit_app.log"
RUN_LOG_ENV = "HR_TOOLKIT_APP_LOG"
RUN_LOG_JSON_ENV = "HR_TOOLKIT_LOG_JSON"
LOG_MAX_BYTES = 1024 * 1024
LOG_KEEP_BYTES = 256 * 1024

_write_lock = threading.Lock()


def current_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd().resolve()


def user_log_dir() -> Path:
    """Return a writable per-user log directory on every supported desktop."""

    if sys.platform.startswith("win"):
        base_text = os.environ.get("LOCALAPPDATA", "").strip()
        base = Path(base_text) if base_text else Path.home() / "AppData" / "Local"
        return base / "HRToolkit" / "logs"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Logs" / "HRToolkit"
    state_text = os.environ.get("XDG_STATE_HOME", "").strip()
    base = Path(state_text) if state_text else Path.home() / ".local" / "state"
    return base / "HRToolkit" / "logs"


def trim_log_file(log_file: Path, max_bytes: int = LOG_MAX_BYTES, keep_bytes: int = LOG_KEEP_BYTES) -> None:
    """日志超限时只保留末尾内容，避免日志文件无限增长。"""
    try:
        if not log_file.exists():
            return
        size = log_file.stat().st_size
        if size <= max_bytes:
            return
        with log_file.open("rb") as source:
            source.seek(max(size - keep_bytes, 0))
            data = source.read()
        newline = data.find(b"\n")
        if newline >= 0:
            data = data[newline + 1 :]
        log_file.write_bytes(b"(...earlier log trimmed...)\n" + data)
    except OSError:
        pass


def is_json_log_enabled() -> bool:
    val = os.environ.get(RUN_LOG_JSON_ENV, "").strip().lower()
    return val in {"1", "true", "yes", "on"}


def run_log_path() -> Path:
    env_path = os.environ.get(RUN_LOG_ENV, "").strip()
    if env_path:
        return Path(env_path)
    if getattr(sys, "frozen", False):
        return user_log_dir() / RUN_LOG_FILE
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


def log_event(event_name: str, **kwargs: Any) -> None:
    """记录结构化运行事件。"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sanitized_kwargs = {
        key: describe_value(value) if isinstance(value, (Path, list, tuple)) else value
        for key, value in kwargs.items()
        if value is not None
    }
    if is_json_log_enabled():
        record = {
            "timestamp": timestamp,
            "event": event_name,
            **sanitized_kwargs,
        }
        try:
            line = json.dumps(record, ensure_ascii=False, default=str)
            path = run_log_path()
            path.parent.mkdir(parents=True, exist_ok=True)
            with _write_lock:
                trim_log_file(path)
                with path.open("a", encoding="utf-8") as handle:
                    handle.write(line + "\n")
        except Exception:
            pass
    else:
        parts = [f"{k}={v}" for k, v in sanitized_kwargs.items()]
        body = f"[EVENT: {event_name}] " + "；".join(parts) if parts else f"[EVENT: {event_name}]"
        log_line(body)


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
