"""Desktop, path, project and workspace utility helpers."""

from __future__ import annotations

import errno
import importlib
import os
import subprocess
import sys
import threading
from datetime import date, datetime, timezone
from pathlib import Path

from hr_toolkit import runlog


def _default_workspace_project_name(today_value: date | None = None) -> str:
    current = today_value or date.today()
    return f"{current.year}年{current.month}月人事月度工作"


def _workspace_project_name_error(value: str) -> str | None:
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


def _workspace_project_creation_target(parent_value: str, name_value: str) -> tuple[Path | None, str | None]:
    name_error = _workspace_project_name_error(name_value)
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


def _workspace_project_create_error_message(exc: Exception) -> str:
    if isinstance(exc, PermissionError):
        return "没有权限在这里创建项目，请选择“文档”等本机文件夹。"
    if isinstance(exc, FileNotFoundError):
        return "保存位置已不可用，请重新选择。"
    if isinstance(exc, OSError) and getattr(exc, "errno", None) == errno.ENOSPC:
        return "磁盘空间不足，无法创建项目。"
    detail = str(exc).strip()
    return detail or "暂时无法在这里创建项目，请检查磁盘连接或写入权限。"


def _workspace_trash_period_label(value: str) -> str:
    text = str(value or "").strip()
    try:
        year_text, month_text = text.split("-", 1)
        year = int(year_text)
        month = int(month_text)
    except (TypeError, ValueError):
        return text
    if 1 <= month <= 12 and 1900 <= year <= 9999:
        return f"{year}年{month}月"
    return text


def _workspace_trash_title(detail) -> str:
    summary = getattr(detail, "summary", None)
    period = _workspace_trash_period_label(getattr(summary, "business_period", ""))
    description = str(getattr(summary, "business_description", "") or "").strip()
    directory_name = str(getattr(summary, "directory_name", "") or "").strip()
    return description or period or directory_name or "处理记录"


def _workspace_trash_group_tool(detail, *, separator: str = " · ") -> str:
    summary = getattr(detail, "summary", None)
    group_name = str(getattr(summary, "group_name", "") or "").strip()
    tool_name = str(getattr(summary, "tool_name", "") or "").strip()
    return separator.join(value for value in (group_name, tool_name) if value) or "功能信息待确认"


def _workspace_trash_restore_location(detail) -> str:
    summary = getattr(detail, "summary", None)
    group_name = str(getattr(summary, "group_name", "") or "").strip()
    tool_name = str(getattr(summary, "tool_name", "") or "").strip()
    original_parts = [
        part.strip()
        for part in str(getattr(detail, "original_relative_path", "") or "").replace("\\", "/").split("/")
        if part.strip()
    ]
    if not group_name and original_parts:
        group_name = original_parts[0]
    if not tool_name and len(original_parts) >= 2:
        tool_name = original_parts[1]
    return " / ".join(value for value in (group_name, tool_name) if value) or "原业务位置"


def _workspace_trash_dialog_height(preferred: int, required: int, maximum: int) -> int:
    preferred_height = max(1, int(preferred))
    required_height = max(1, int(required))
    maximum_height = max(1, int(maximum))
    return min(max(preferred_height, required_height), maximum_height)


def _workspace_trash_ignore_enter(_event=None) -> str:
    return "break"


def _workspace_trash_matches(detail, query: str) -> bool:
    normalized = str(query or "").strip().casefold()
    if not normalized:
        return True
    summary = getattr(detail, "summary", None)
    values = (
        getattr(summary, "business_period", ""),
        getattr(summary, "business_description", ""),
        getattr(summary, "group_name", ""),
        getattr(summary, "tool_name", ""),
        getattr(summary, "directory_name", ""),
        getattr(detail, "original_relative_path", ""),
        _workspace_trash_title(detail),
        _workspace_trash_group_tool(detail),
    )
    return any(normalized in str(value or "").casefold() for value in values)


def _workspace_trash_deleted_text(value: str | None) -> str:
    raw = str(value or "").strip()
    if not raw:
        return "移入时间未知"
    try:
        parsed = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone().strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return raw


def _default_result_dir_name() -> str:
    return "结果_" + datetime.now().strftime("%Y%m%d_%H%M%S")


def default_output_parent_dir(tool: str) -> Path:
    if tool == "social_security":
        folder_name = "社保汇总结果"
    elif tool == "data_statistics":
        folder_name = "数据统计结果"
    elif tool == "insurance_ledger":
        folder_name = "保险台账结果"
    elif tool == "salary_merge":
        folder_name = "工资合并结果"
    elif tool == "personnel_change_merge":
        folder_name = "异动表汇总结果"
    elif tool == "archive_import":
        folder_name = "档案处理结果"
    else:
        folder_name = "工资表拆分结果"
    return desktop_dir() / folder_name


def make_result_output_dir(parent_dir: Path) -> Path:
    parent_dir = parent_dir.expanduser()
    parent_dir.mkdir(parents=True, exist_ok=True)
    base_name = _default_result_dir_name()
    for index in range(1, 10_000):
        name = base_name if index == 1 else f"{base_name}_{index}"
        candidate = parent_dir / name
        try:
            candidate.mkdir(exist_ok=False)
        except FileExistsError:
            continue
        return candidate
    raise RuntimeError("同一保存位置的结果文件夹过多，请更换保存位置。")


def desktop_dir() -> Path:
    home = Path.home()
    desktop = home / "Desktop"
    if desktop.exists():
        return desktop
    return home / "桌面"


def open_path(path: Path | str) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def _enable_high_dpi_rendering() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
    except Exception:
        return

    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-2)):
            return
    except Exception:
        pass
    try:
        if ctypes.windll.shcore.SetProcessDpiAwareness(1) == 0:
            return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _set_windows_app_identity() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("HRWorkbench.HRToolkit")
    except Exception:
        pass


def _install_crash_logging() -> None:
    default_excepthook = sys.excepthook

    def log_and_delegate(exc_type, exc_value, exc_tb):
        runlog.log_exception("程序异常退出", exc_value, exc_tb)
        default_excepthook(exc_type, exc_value, exc_tb)

    sys.excepthook = log_and_delegate

    default_thread_hook = threading.excepthook

    def log_thread_exception(args) -> None:
        if args.exc_value is not None:
            runlog.log_exception(f"后台线程异常（{args.thread.name if args.thread else '未知'}）", args.exc_value, args.exc_traceback)
        default_thread_hook(args)

    threading.excepthook = log_thread_exception


__all__ = [
    "_default_workspace_project_name",
    "_workspace_project_name_error",
    "_workspace_project_creation_target",
    "_workspace_project_create_error_message",
    "_workspace_trash_period_label",
    "_workspace_trash_title",
    "_workspace_trash_group_tool",
    "_workspace_trash_restore_location",
    "_workspace_trash_dialog_height",
    "_workspace_trash_ignore_enter",
    "_workspace_trash_matches",
    "_workspace_trash_deleted_text",
    "_default_result_dir_name",
    "default_output_parent_dir",
    "make_result_output_dir",
    "desktop_dir",
    "open_path",
    "_enable_high_dpi_rendering",
    "_set_windows_app_identity",
    "_install_crash_logging",
]
