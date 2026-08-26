from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from build_update_assets import legacy_manifest_name
from build_windows import (
    APP_NAME,
    UPDATER_NAME,
    WINDOWS_TARGET_MODERN,
    WINDOWS_TARGETS,
    validate_build_version,
    validate_windows_target,
)
from build_windows_installers import installer_asset_names
from generate_release_metadata import require_release_assets_under_limit
from versioning import read_project_version


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Windows 三阶段发布产物编排器：纯构建、纯安装器、桥接更新清单。"
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
        help="EXE/MSI/桥接清单输出目录",
    )
    parser.add_argument("--notes", nargs="*", default=None, help="旧服务器桥接更新说明")
    parser.add_argument("--optional", action="store_true", help="旧服务器桥接清单标记为可选更新")
    parser.add_argument("--inno-compiler", help="ISCC.exe 路径或命令名")
    parser.add_argument("--wix-executable", help="WiX v4 wix.exe 路径或命令名")
    parser.add_argument(
        "--target",
        choices=WINDOWS_TARGETS,
        default=WINDOWS_TARGET_MODERN,
        help="modern 保持现有发布资产；win7 生成独立兼容资产",
    )
    parser.add_argument("--seven-zip-dir", type=Path, help="Win7 构建的 7-Zip x64 运行时目录")
    parser.add_argument("--ucrt-dir", type=Path, help="Win7 构建的 app-local UCRT x64 目录")
    parser.add_argument("--vc-runtime-dir", type=Path, help="Win7 构建的 Visual C++ x64 目录")
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
        target=args.target,
        seven_zip_dir=args.seven_zip_dir,
        ucrt_dir=args.ucrt_dir,
        vc_runtime_dir=args.vc_runtime_dir,
    )
    for label, command in commands:
        started_at = time.perf_counter()
        print(f"\n=== {label} ===", flush=True)
        _run(command)
        elapsed = time.perf_counter() - started_at
        print(f"=== {label} 完成，用时 {elapsed:.1f} 秒 ===", flush=True)

    exe_name, msi_name = installer_asset_names(version, args.target)
    expected = (
        args.output_dir.resolve() / exe_name,
        args.output_dir.resolve() / msi_name,
        args.output_dir.resolve() / legacy_manifest_name(args.target),
    )
    missing = [path for path in expected if not path.is_file()]
    if missing:
        raise RuntimeError(f"Windows 三阶段完成后缺少产物：{missing}")
    require_release_assets_under_limit(expected[:2])
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
    target: str = WINDOWS_TARGET_MODERN,
    seven_zip_dir: Path | None = None,
    ucrt_dir: Path | None = None,
    vc_runtime_dir: Path | None = None,
) -> tuple[tuple[str, list[str]], ...]:
    validate_build_version(version)
    target = validate_windows_target(target)
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
        "--target",
        target,
    ]
    if seven_zip_dir is not None:
        build.extend(["--seven-zip-dir", str(seven_zip_dir.resolve())])
    if ucrt_dir is not None:
        build.extend(["--ucrt-dir", str(ucrt_dir.resolve())])
    if vc_runtime_dir is not None:
        build.extend(["--vc-runtime-dir", str(vc_runtime_dir.resolve())])
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
        "--target",
        target,
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
        "--skip-runtime-smoke",
        "--target",
        target,
    ]
    if notes:
        update_assets.extend(["--notes", *notes])
    if optional:
        update_assets.append("--optional")

    return (
        ("1/3 PyInstaller 纯构建", build),
        ("2/3 EXE/MSI 安装器", installers),
        ("3/3 旧服务器桥接更新资产", update_assets),
    )


def _run(command: list[str]) -> None:
    print("执行：" + subprocess.list2cmdline(command), flush=True)
    subprocess.run(command, cwd=REPO_ROOT, check=True)


if __name__ == "__main__":
    raise SystemExit(main())
