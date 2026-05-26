from __future__ import annotations

import argparse
import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import zipfile
from datetime import datetime
from pathlib import Path


UPDATE_LOG_FILE = "HRToolkit_update.log"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HR工具箱独立更新程序")
    parser.add_argument("--zip", required=True, type=Path, help="更新包 zip 路径")
    parser.add_argument("--app-dir", required=True, type=Path, help="当前程序目录")
    parser.add_argument("--launcher", required=True, help="主程序文件名")
    parser.add_argument("--wait-pid", type=int, help="需要等待退出的主程序 PID")
    parser.add_argument("--log-file", type=Path, help="更新详细日志路径")
    parser.add_argument("--relaunch", action="store_true", help="更新完成后重新打开主程序")
    args = parser.parse_args(argv)

    log_file = _resolve_log_file(args)
    _append_log(log_file, "")
    _append_log(log_file, f"===== {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 更新开始 =====")
    try:
        _run_update(args, log_file)
    except Exception as exc:
        _append_log(log_file, f"更新失败：{exc}")
        _append_log(log_file, traceback.format_exc().rstrip())
        return 1
    _append_log(log_file, "更新完成。")
    _append_log(log_file, f"===== {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 更新结束 =====")
    return 0


def _run_update(args: argparse.Namespace, log_file: Path) -> None:
    package_path = args.zip.resolve()
    app_dir = args.app_dir.resolve()
    _append_log(log_file, f"系统平台：{sys.platform}")
    _append_log(log_file, f"更新程序：{Path(sys.argv[0]).resolve()}")
    _append_log(log_file, f"更新包路径：{package_path}")
    _append_log(log_file, f"程序目录：{app_dir}")
    _append_log(log_file, f"主程序文件：{args.launcher}")
    _switch_working_dir(app_dir.parent, log_file)
    if args.wait_pid:
        _append_log(log_file, f"等待主程序退出，PID：{args.wait_pid}")
        _wait_for_process(args.wait_pid, timeout_seconds=120)
        if sys.platform.startswith("win"):
            time.sleep(1)
        _append_log(log_file, "主程序已退出。")
    if not package_path.exists():
        raise RuntimeError(f"更新包不存在：{package_path}")
    if not app_dir.exists():
        raise RuntimeError(f"程序目录不存在：{app_dir}")

    extract_dir = Path(tempfile.mkdtemp(prefix="hr_toolkit_extract_"))
    _append_log(log_file, f"解压目录：{extract_dir}")
    _safe_extract_zip(package_path, extract_dir, log_file)
    payload_root = _find_payload_root(extract_dir, log_file)
    _append_log(log_file, f"更新包根目录：{payload_root}")
    _validate_payload_root(payload_root, args.launcher)
    _append_log(log_file, "更新包校验通过。")
    _replace_app_dir(payload_root, app_dir, log_file)

    if args.relaunch:
        launcher = app_dir / args.launcher
        if launcher.exists():
            _append_log(log_file, f"重新打开主程序：{launcher}")
            subprocess.Popen([str(launcher)], cwd=str(app_dir.parent), close_fds=True)
        else:
            _append_log(log_file, f"跳过重新打开，未找到主程序：{launcher}")


def _safe_extract_zip(package_path: Path, extract_dir: Path, log_file: Path) -> None:
    extract_root = extract_dir.resolve()
    with zipfile.ZipFile(package_path) as archive:
        members = archive.infolist()
        _append_log(log_file, f"zip 文件数量：{len(members)}")
        for member in members[:20]:
            _append_log(log_file, f"zip 条目：{member.filename}")
        for member in archive.infolist():
            target = (extract_dir / member.filename).resolve()
            if extract_root not in target.parents and target != extract_root:
                raise RuntimeError("更新包包含非法路径。")
        archive.extractall(extract_dir)
    _append_log(log_file, "更新包解压完成。")


def _find_payload_root(extract_dir: Path, log_file: Path) -> Path:
    entries = [item for item in extract_dir.iterdir() if item.name != "__MACOSX"]
    _append_log(log_file, "解压顶层内容：" + ", ".join(item.name for item in entries[:30]))
    candidates = [extract_dir]
    if len(entries) == 1 and entries[0].is_dir():
        candidates.insert(0, entries[0])
    for candidate in candidates:
        if (candidate / "_internal").exists() or (candidate / "HRToolkit.exe").exists() or (candidate / "HRToolkit").exists():
            return candidate
    return candidates[0]


def _replace_app_dir(payload_root: Path, app_dir: Path, log_file: Path) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = app_dir.parent / f"{app_dir.name}_backup_{timestamp}"
    new_dir = app_dir.parent / f"{app_dir.name}_new_{timestamp}"
    try:
        _append_log(log_file, f"复制新版本到临时目录：{new_dir}")
        _copytree_with_retry(payload_root, new_dir, log_file)
        _validate_copied_app_dir(new_dir)
        _append_log(log_file, "新版本临时目录校验通过。")
        _append_log(log_file, f"备份旧目录：{app_dir} -> {backup_dir}")
        _move_to_exact_path(app_dir, backup_dir, log_file)
        _append_log(log_file, f"启用新目录：{new_dir} -> {app_dir}")
        _move_to_exact_path(new_dir, app_dir, log_file)
    except Exception:
        _append_log(log_file, "替换失败，开始回滚。")
        if new_dir.exists():
            _rmtree_with_retry(new_dir, log_file, ignore_errors=True)
        if backup_dir.exists():
            _prepare_restore_target(app_dir, log_file, timestamp)
            _append_log(log_file, f"恢复旧目录：{backup_dir} -> {app_dir}")
            _move_to_exact_path(backup_dir, app_dir, log_file)
        raise
    if _validate_copied_app_dir(app_dir):
        _append_log(log_file, "最终目录校验通过。")
        _append_log(log_file, f"删除备份目录：{backup_dir}")
        _rmtree_with_retry(backup_dir, log_file, ignore_errors=True)


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


def _switch_working_dir(target_dir: Path, log_file: Path) -> None:
    try:
        os.chdir(target_dir)
    except OSError as exc:
        _append_log(log_file, f"切换工作目录失败：{target_dir}，原因：{exc}")
        raise
    _append_log(log_file, f"工作目录已切换到：{Path.cwd()}")


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


def _copytree_with_retry(source: Path, target: Path, log_file: Path, attempts: int = 5) -> None:
    def copy_action() -> None:
        if target.exists():
            shutil.rmtree(target, ignore_errors=True)
        shutil.copytree(source, target)

    _run_with_retry(copy_action, f"复制目录 {source} -> {target}", log_file, attempts=attempts)


def _rename_with_retry(source: Path, target: Path, log_file: Path, attempts: int = 10) -> None:
    _run_with_retry(lambda: os.rename(source, target), f"重命名目录 {source} -> {target}", log_file, attempts=attempts)


def _move_to_exact_path(source: Path, target: Path, log_file: Path) -> None:
    _ensure_target_missing(target, log_file)
    _rename_with_retry(source, target, log_file)


def _ensure_target_missing(target: Path, log_file: Path) -> None:
    if not target.exists():
        return
    if target.is_dir() and not any(target.iterdir()):
        _append_log(log_file, f"目标目录已存在但为空，先删除：{target}")
        _rmtree_with_retry(target, log_file, ignore_errors=False)
        return
    raise RuntimeError(f"目标目录已存在，不能覆盖：{target}")


def _prepare_restore_target(app_dir: Path, log_file: Path, timestamp: str) -> None:
    if not app_dir.exists():
        return
    if app_dir.is_dir() and not any(app_dir.iterdir()):
        _append_log(log_file, f"发现空程序目录，删除后恢复：{app_dir}")
        _rmtree_with_retry(app_dir, log_file, ignore_errors=False)
        return
    failed_dir = app_dir.parent / f"{app_dir.name}_failed_{timestamp}"
    _append_log(log_file, f"发现未完成的新目录，先保留到：{failed_dir}")
    _move_to_exact_path(app_dir, failed_dir, log_file)


def _rmtree_with_retry(path: Path, log_file: Path, attempts: int = 10, ignore_errors: bool = False) -> None:
    try:
        _run_with_retry(lambda: shutil.rmtree(path), f"删除目录 {path}", log_file, attempts=attempts)
    except Exception:
        if not ignore_errors:
            raise
        _append_log(log_file, f"删除目录失败但已忽略：{path}")


def _run_with_retry(action, description: str, log_file: Path, attempts: int) -> None:
    last_error: Exception | None = None
    for index in range(1, attempts + 1):
        try:
            action()
            if index > 1:
                _append_log(log_file, f"{description} 第 {index} 次尝试成功。")
            return
        except Exception as exc:
            last_error = exc
            _append_log(log_file, f"{description} 第 {index} 次失败：{exc}")
            time.sleep(min(0.5 * index, 3))
    assert last_error is not None
    raise last_error


def _resolve_log_file(args: argparse.Namespace) -> Path:
    if args.log_file is not None:
        return args.log_file.resolve()
    try:
        return args.app_dir.resolve().parent / UPDATE_LOG_FILE
    except Exception:
        return Path(tempfile.gettempdir()) / UPDATE_LOG_FILE


def _append_log(log_file: Path, text: str) -> None:
    try:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        with log_file.open("a", encoding="utf-8") as handle:
            handle.write(text + "\n")
    except OSError:
        fallback = Path(tempfile.gettempdir()) / UPDATE_LOG_FILE
        if log_file != fallback:
            with fallback.open("a", encoding="utf-8") as handle:
                handle.write(text + "\n")


if __name__ == "__main__":
    raise SystemExit(main())
