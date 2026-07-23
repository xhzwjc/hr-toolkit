from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from build_update_assets import LEGACY_MANIFEST_NAME, update_zip_name
from build_windows import APP_NAME, UPDATER_NAME, validate_build_version
from build_windows_installers import installer_asset_names
from versioning import read_project_version


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Windows 三阶段发布产物编排器：纯构建、纯安装器、纯更新资产。"
            "不修改版本、不提交、不创建 Tag、不推送。"
        )
    )
    parser.add_argument("--version", default=read_project_version(), help="必须与 hr_toolkit.__version__ 一致")
    parser.add_argument(
        "--build-dir",
        type=Path,
        default=REPO_ROOT / "dist" / "windows",
        help="PyInstaller 二进制输出目录",
    )
    parser.add_argument(
        "--work-dir",
        type=Path,
        default=REPO_ROOT / "build" / "windows",
        help="PyInstaller 临时工作目录",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=REPO_ROOT / "dist" / "release-windows",
        help="EXE/MSI/ZIP/桥接清单输出目录",
    )
    parser.add_argument("--notes", nargs="*", default=None, help="旧服务器桥接更新说明")
    parser.add_argument("--optional", action="store_true", help="旧服务器桥接清单标记为可选更新")
    parser.add_argument("--inno-compiler", help="ISCC.exe 路径或命令名")
    parser.add_argument("--wix-executable", help="WiX v4 wix.exe 路径或命令名")
    parser.add_argument(
        "--skip-install-smoke",
        action="store_true",
        help="仅供诊断；跳过安装器静默安装、运行和卸载验证",
    )
    args = parser.parse_args(argv)

    version = validate_build_version(args.version)
    commands = stage_commands(
        version=version,
        build_dir=args.build_dir.resolve(),
        work_dir=args.work_dir.resolve(),
        output_dir=args.output_dir.resolve(),
        notes=args.notes,
        optional=args.optional,
        inno_compiler=args.inno_compiler,
        wix_executable=args.wix_executable,
        skip_install_smoke=args.skip_install_smoke,
    )
    for label, command in commands:
        print(f"\n=== {label} ===")
        _run(command)

    exe_name, msi_name = installer_asset_names(version)
    expected = (
        args.output_dir.resolve() / exe_name,
        args.output_dir.resolve() / msi_name,
        args.output_dir.resolve() / update_zip_name(version),
        args.output_dir.resolve() / LEGACY_MANIFEST_NAME,
    )
    missing = [path for path in expected if not path.is_file()]
    if missing:
        raise RuntimeError(f"Windows 三阶段完成后缺少产物：{missing}")
    print("\nWindows 发布资产已完成：")
    for path in expected:
        print(f"- {path}")
    return 0


def stage_commands(
    *,
    version: str,
    build_dir: Path,
    work_dir: Path,
    output_dir: Path,
    notes: list[str] | None = None,
    optional: bool = False,
    inno_compiler: str | None = None,
    wix_executable: str | None = None,
    skip_install_smoke: bool = False,
) -> tuple[tuple[str, list[str]], ...]:
    validate_build_version(version)
    python = sys.executable
    app_dir = build_dir / APP_NAME
    updater = build_dir / f"{UPDATER_NAME}.exe"

    build = [
        python,
        str(SCRIPT_DIR / "build_windows.py"),
        "--version",
        version,
        "--output-dir",
        str(build_dir),
        "--work-dir",
        str(work_dir),
    ]
    installers = [
        python,
        str(SCRIPT_DIR / "build_windows_installers.py"),
        "--version",
        version,
        "--app-dir",
        str(app_dir),
        "--updater",
        str(updater),
        "--output-dir",
        str(output_dir),
    ]
    if inno_compiler:
        installers.extend(["--inno-compiler", inno_compiler])
    if wix_executable:
        installers.extend(["--wix-executable", wix_executable])
    if skip_install_smoke:
        installers.append("--skip-install-smoke")

    update_assets = [
        python,
        str(SCRIPT_DIR / "build_update_assets.py"),
        "--version",
        version,
        "--app-dir",
        str(app_dir),
        "--updater",
        str(updater),
        "--output-dir",
        str(output_dir),
    ]
    if notes:
        update_assets.extend(["--notes", *notes])
    if optional:
        update_assets.append("--optional")

    return (
        ("1/3 PyInstaller 纯构建", build),
        ("2/3 EXE/MSI 安装器", installers),
        ("3/3 Windows ZIP 更新资产", update_assets),
    )


def _run(command: list[str]) -> None:
    print("执行：" + subprocess.list2cmdline(command))
    subprocess.run(command, cwd=REPO_ROOT, check=True)


if __name__ == "__main__":
    raise SystemExit(main())
