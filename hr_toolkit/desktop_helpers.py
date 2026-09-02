"""GUI-framework-neutral desktop helpers shared by Qt and legacy Tk."""

from __future__ import annotations

import errno
import importlib
import os
import subprocess
import sys
import threading
from datetime import date
from pathlib import Path
from typing import Optional, Tuple, Union

from hr_toolkit import runlog


def default_workspace_project_name(today_value: Optional[date] = None) -> str:
    current = today_value or date.today()
    return f"{current.year}年{current.month}月人事月度工作"


def workspace_project_name_error(value: str) -> Optional[str]:
    try:
        module = importlib.import_module("hr_toolkit.project_store")
        validator = getattr(module, "validate_project_name", None)
        validation_error = getattr(module, "ProjectStoreError", None)
    except Exception:
        validator = None
        validation_error = None
    if callable(validator) and isinstance(validation_error, type):
        try:
            validator(value)
        except validation_error as exc:
            return str(exc)
        return None

    project_name = str(value).strip()
    if not project_name:
        return "项目名称不能为空。"
    if len(project_name) > 120:
        return "项目名称不能超过 120 个字。"
    if project_name in {".", ".."}:
        return "项目名称不能使用英文句点。"
    if project_name.endswith("."):
        return "项目名称末尾不能使用句点。"
    if any(ord(character) < 32 for character in project_name):
        return "项目名称不能包含换行或控制字符。"
    if any(character in '<>:"/\\|?*' for character in project_name):
        return '项目名称不能包含 \\ / : * ? " < > |。'
    if project_name.casefold() == ".hrtoolkit":
        return "该名称是项目保留名称，请换一个名称。"
    portable_base = project_name.split(".", 1)[0].casefold()
    windows_reserved = {
        "con",
        "prn",
        "aux",
        "nul",
        *(f"com{index}" for index in range(1, 10)),
        *(f"lpt{index}" for index in range(1, 10)),
    }
    if portable_base in windows_reserved:
        return "该名称是 Windows 系统保留名称，请换一个名称。"
    return None


def workspace_project_creation_target(
    parent_value: str,
    name_value: str,
) -> Tuple[Optional[Path], Optional[str]]:
    name_error = workspace_project_name_error(name_value)
    parent_text = str(parent_value).strip()
    if not parent_text:
        return None, name_error or "请选择保存位置。"

    parent_dir = Path(parent_text).expanduser().absolute()
    project_name = str(name_value).strip()
    project_root = parent_dir / project_name if project_name else None
    if name_error:
        return project_root, name_error
    try:
        if not parent_dir.is_dir():
            return project_root, "保存位置已不可用，请重新选择。"
        if project_root is not None and project_root.exists():
            if not project_root.is_dir():
                return project_root, "同名位置已被文件占用，请修改项目名称。"
            if next(project_root.iterdir(), None) is not None:
                return project_root, "这里已有同名文件夹，请修改项目名称；如需继续以前的工作，请打开已有项目。"
    except OSError:
        return project_root, "暂时无法读取这个保存位置，请重新选择。"
    return project_root, None


def workspace_project_create_error_message(exc: Exception) -> str:
    if isinstance(exc, PermissionError):
        return "没有权限在这里创建项目，请选择“文档”等本机文件夹。"
    if isinstance(exc, FileNotFoundError):
        return "保存位置已不可用，请重新选择。"
    if isinstance(exc, OSError) and getattr(exc, "errno", None) == errno.ENOSPC:
        return "磁盘空间不足，无法创建项目。"
    detail = str(exc).strip()
    return detail or "暂时无法在这里创建项目，请检查磁盘连接或写入权限。"


def open_path(path: Union[Path, str]) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def set_windows_app_identity() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("HRWorkbench.HRToolkit")
    except Exception:
        pass


def install_crash_logging() -> None:
    default_excepthook = sys.excepthook

    def log_and_delegate(exc_type, exc_value, exc_tb):
        runlog.log_exception("程序异常退出", exc_value, exc_tb)
        default_excepthook(exc_type, exc_value, exc_tb)

    sys.excepthook = log_and_delegate
    default_thread_hook = threading.excepthook

    def log_thread_exception(args) -> None:
        if args.exc_value is not None:
            thread_name = args.thread.name if args.thread else "未知"
            runlog.log_exception(
                f"后台线程异常（{thread_name}）",
                args.exc_value,
                args.exc_traceback,
            )
        default_thread_hook(args)

    threading.excepthook = log_thread_exception


__all__ = [
    "default_workspace_project_name",
    "workspace_project_name_error",
    "workspace_project_creation_target",
    "workspace_project_create_error_message",
    "open_path",
    "set_windows_app_identity",
    "install_crash_logging",
]
