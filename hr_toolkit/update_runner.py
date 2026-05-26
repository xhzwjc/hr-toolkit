from __future__ import annotations

import argparse
import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import datetime
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HR工具箱独立更新程序")
    parser.add_argument("--zip", required=True, type=Path, help="更新包 zip 路径")
    parser.add_argument("--app-dir", required=True, type=Path, help="当前程序目录")
    parser.add_argument("--launcher", required=True, help="主程序文件名")
    parser.add_argument("--wait-pid", type=int, help="需要等待退出的主程序 PID")
    parser.add_argument("--relaunch", action="store_true", help="更新完成后重新打开主程序")
    args = parser.parse_args(argv)

    log_file = Path(tempfile.gettempdir()) / "hr_toolkit_update.log"
    try:
        _run_update(args)
    except Exception as exc:
        _append_log(log_file, f"更新失败：{exc}")
        return 1
    _append_log(log_file, "更新完成。")
    return 0


def _run_update(args: argparse.Namespace) -> None:
    package_path = args.zip.resolve()
    app_dir = args.app_dir.resolve()
    if args.wait_pid:
        _wait_for_process(args.wait_pid, timeout_seconds=120)
    if not package_path.exists():
        raise RuntimeError(f"更新包不存在：{package_path}")
    if not app_dir.exists():
        raise RuntimeError(f"程序目录不存在：{app_dir}")

    extract_dir = Path(tempfile.mkdtemp(prefix="hr_toolkit_extract_"))
    _safe_extract_zip(package_path, extract_dir)
    payload_root = _find_payload_root(extract_dir)
    _validate_payload_root(payload_root, args.launcher)
    _replace_app_dir(payload_root, app_dir)

    if args.relaunch:
        launcher = app_dir / args.launcher
        if launcher.exists():
            subprocess.Popen([str(launcher)], close_fds=True)


def _safe_extract_zip(package_path: Path, extract_dir: Path) -> None:
    extract_root = extract_dir.resolve()
    with zipfile.ZipFile(package_path) as archive:
        for member in archive.infolist():
            target = (extract_dir / member.filename).resolve()
            if extract_root not in target.parents and target != extract_root:
                raise RuntimeError("更新包包含非法路径。")
        archive.extractall(extract_dir)


def _find_payload_root(extract_dir: Path) -> Path:
    entries = [item for item in extract_dir.iterdir() if item.name != "__MACOSX"]
    candidates = [extract_dir]
    if len(entries) == 1 and entries[0].is_dir():
        candidates.insert(0, entries[0])
    for candidate in candidates:
        if (candidate / "_internal").exists() or (candidate / "HRToolkit.exe").exists() or (candidate / "HRToolkit").exists():
            return candidate
    return candidates[0]


def _replace_app_dir(payload_root: Path, app_dir: Path) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = app_dir.parent / f"{app_dir.name}_backup_{timestamp}"
    new_dir = app_dir.parent / f"{app_dir.name}_new_{timestamp}"
    try:
        shutil.copytree(payload_root, new_dir)
        _validate_copied_app_dir(new_dir)
        shutil.move(str(app_dir), str(backup_dir))
        shutil.move(str(new_dir), str(app_dir))
    except Exception:
        if new_dir.exists():
            shutil.rmtree(new_dir, ignore_errors=True)
        if not app_dir.exists() and backup_dir.exists():
            shutil.move(str(backup_dir), str(app_dir))
        elif app_dir.exists() and not any(app_dir.iterdir()) and backup_dir.exists():
            shutil.rmtree(app_dir, ignore_errors=True)
            shutil.move(str(backup_dir), str(app_dir))
        raise
    if _validate_copied_app_dir(app_dir):
        shutil.rmtree(backup_dir, ignore_errors=True)


def _validate_payload_root(payload_root: Path, launcher: str) -> None:
    launcher_path = payload_root / launcher
    if not launcher_path.exists():
        raise RuntimeError(f"更新包缺少主程序：{launcher}")
    if not (payload_root / "_internal").exists():
        raise RuntimeError("更新包缺少 _internal 目录。")


def _validate_copied_app_dir(app_dir: Path) -> bool:
    launchers = ("HRToolkit.exe", "HRToolkit")
    if not any((app_dir / launcher).exists() for launcher in launchers):
        raise RuntimeError("更新后目录缺少 HRToolkit 主程序。")
    if not (app_dir / "_internal").exists():
        raise RuntimeError("更新后目录缺少 _internal 目录。")
    return True


def _wait_for_process(pid: int, timeout_seconds: int) -> None:
    if sys.platform.startswith("win"):
        _wait_for_windows_process(pid, timeout_seconds)
        return
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        try:
            os.kill(pid, 0)
        except OSError:
            return
        time.sleep(0.3)
    raise RuntimeError("等待主程序退出超时。")


def _wait_for_windows_process(pid: int, timeout_seconds: int) -> None:
    synchronize = 0x00100000
    wait_timeout = 0x00000102
    wait_failed = 0xFFFFFFFF
    handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, pid)
    if not handle:
        return
    try:
        result = ctypes.windll.kernel32.WaitForSingleObject(handle, timeout_seconds * 1000)
        if result in (wait_timeout, wait_failed):
            raise RuntimeError("等待主程序退出超时。")
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def _append_log(log_file: Path, text: str) -> None:
    log_file.write_text(
        (log_file.read_text(encoding="utf-8") if log_file.exists() else "") + text + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    raise SystemExit(main())
