from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from prepare_gitee_release import main as prepare_gitee_release
from versioning import REPO_ROOT, bump_project_version


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Windows 一键发布 HR工具箱")
    parser.add_argument("--bump", choices=["patch", "minor", "major"], default="patch")
    parser.add_argument("--notes", nargs="*", default=["更新 HR工具箱"], help="更新说明，可写多条")
    parser.add_argument("--publish-dir", type=Path, help="可选：直接复制到 ScriptHub 的 fastApiProject/static/hr-toolkit 目录")
    args = parser.parse_args(argv)

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


if __name__ == "__main__":
    raise SystemExit(main())
