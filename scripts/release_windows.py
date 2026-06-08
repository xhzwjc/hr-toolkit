from __future__ import annotations

import argparse
import importlib.util
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from prepare_gitee_release import main as prepare_gitee_release
from versioning import REPO_ROOT, bump_project_version


WINDOWS_BUILD_MODULES = {
    "PyInstaller": "pyinstaller",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd",
    "pythoncom": "pywin32",
    "pywintypes": "pywin32",
    "win32com.client": "pywin32",
    "win32timezone": "pywin32",
}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Windows 一键发布 HR工具箱")
    parser.add_argument("--bump", choices=["patch", "minor", "major"], default="patch")
    parser.add_argument("--notes", nargs="*", default=["更新 HR工具箱"], help="更新说明，可写多条")
    parser.add_argument("--publish-dir", type=Path, help="可选：直接复制到 ScriptHub 的 fastApiProject/static/hr-toolkit 目录")
    args = parser.parse_args(argv)

    _ensure_windows_build_dependencies()
    new_version = bump_project_version(args.bump)
    print(f"发布版本：{new_version}")

    _run([
        sys.executable,
        "-m",
        "PyInstaller",
        "--name",
        "HRToolkit",
        "--onedir",
        "--windowed",
        "--clean",
        "--add-data",
        "README.md;.",
        "--add-data",
        "hr_toolkit/templates;hr_toolkit/templates",
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "pywintypes",
        "--hidden-import",
        "win32com.client",
        "--hidden-import",
        "win32timezone",
        "--hidden-import",
        "xlrd",
        "hr_toolkit_app.py",
    ])
    _run([
        sys.executable,
        "-m",
        "PyInstaller",
        "--name",
        "HRToolkitUpdater",
        "--onefile",
        "--windowed",
        "--clean",
        "hr_toolkit_updater.py",
    ])
    prepare_gitee_release([
        "--platform",
        "windows",
        "--version",
        new_version,
        "--notes",
        *args.notes,
        *([] if args.publish_dir is None else ["--publish-dir", str(args.publish_dir)]),
    ])
    print("Windows 发布文件已生成。提交并推送后，旧版客户端即可检查到更新。")
    return 0


def _run(command: list[str]) -> None:
    print("执行：" + " ".join(command))
    subprocess.run(command, cwd=REPO_ROOT, check=True)


def _ensure_windows_build_dependencies() -> None:
    if not sys.platform.startswith("win"):
        raise RuntimeError("Windows 发布包必须在 Windows 电脑上打包，不能在 Mac 上直接生成 Windows exe。")

    missing = [module for module in WINDOWS_BUILD_MODULES if not _module_exists(module)]
    if not missing:
        return

    packages = sorted({WINDOWS_BUILD_MODULES[module] for module in missing})
    raise RuntimeError(
        "Windows 打包环境缺少依赖模块："
        + ", ".join(missing)
        + "。请在 Windows 虚拟环境中执行："
        + f"{sys.executable} -m pip install -r requirements.txt；"
        + f"{sys.executable} -m pip install "
        + " ".join(packages)
        + "。安装后重新执行发布命令。"
    )


def _module_exists(module: str) -> bool:
    try:
        return importlib.util.find_spec(module) is not None
    except ModuleNotFoundError:
        return False


if __name__ == "__main__":
    raise SystemExit(main())
