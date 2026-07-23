from __future__ import annotations

import argparse
import hashlib
import os
import plistlib
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Sequence


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE_DIR = REPO_ROOT / "hr_toolkit" / "templates"
DEFAULT_README = REPO_ROOT / "README.md"
EXPECTED_BUNDLE_IDENTIFIER = "com.xhzwjc.hrtoolkit"
ARCHITECTURES = {"universal2", "x86_64", "arm64"}
SPREADSHEET_SUFFIXES = {".csv", ".tsv", ".xls", ".xlsb", ".xlsm", ".xlsx"}
PROHIBITED_PATH_PARTS = {
    ".git",
    ".github",
    "__pycache__",
    "attachments",
    "outputs",
    "tests",
    "testdata",
    "test_data",
    "附件",
    "二期新增的附件",
}


class MacBundleVerificationError(RuntimeError):
    """Raised when an app bundle or DMG is mislabeled or incomplete."""


def _run(command: Sequence[str], *, capture: bool = False) -> subprocess.CompletedProcess:
    result = subprocess.run(
        list(command),
        check=False,
        text=True,
        stdout=subprocess.PIPE if capture else None,
        stderr=subprocess.PIPE if capture else None,
    )
    if result.returncode != 0:
        detail = ""
        if capture:
            detail = (result.stderr or result.stdout or "").strip()
        raise MacBundleVerificationError(
            f"命令失败（{result.returncode}）：{' '.join(command)}" + (f"\n{detail}" if detail else "")
        )
    return result


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _expected_dmg_name(version: str, architecture: str) -> str:
    suffix = {"universal2": "universal", "x86_64": "x64", "arm64": "arm64"}[architecture]
    return f"HRToolkit_{version}_{suffix}.dmg"


def _mach_o_architectures(path: Path) -> Optional[set[str]]:
    file_result = _run(["file", "-b", str(path)], capture=True)
    if "Mach-O" not in file_result.stdout:
        return None
    lipo_result = _run(["lipo", "-archs", str(path)], capture=True)
    architectures = {part.strip() for part in lipo_result.stdout.split() if part.strip()}
    if not architectures:
        raise MacBundleVerificationError(f"lipo 未返回架构：{path}")
    return architectures


def verify_mach_o_architectures(app_path: Path, architecture: str) -> int:
    if architecture not in ARCHITECTURES:
        raise MacBundleVerificationError(f"未知架构：{architecture}")
    mach_o_files = []
    for path in sorted(app_path.rglob("*")):
        if path.is_symlink() or not path.is_file():
            continue
        architectures = _mach_o_architectures(path)
        if architectures is None:
            continue
        mach_o_files.append((path, architectures))
        if architecture == "universal2":
            missing = {"arm64", "x86_64"} - architectures
            if missing:
                raise MacBundleVerificationError(
                    f"universal2 Bundle 中存在缺失架构的 Mach-O：{path}，"
                    f"实际 {sorted(architectures)}，缺少 {sorted(missing)}"
                )
        elif architecture not in architectures:
            raise MacBundleVerificationError(
                f"{architecture} Bundle 中存在不兼容 Mach-O：{path}，实际 {sorted(architectures)}"
            )

    if not mach_o_files:
        raise MacBundleVerificationError(f"Bundle 中没有找到 Mach-O 文件：{app_path}")

    launcher = app_path / "Contents" / "MacOS" / "HRToolkit"
    launcher_architectures = _mach_o_architectures(launcher)
    if launcher_architectures is None:
        raise MacBundleVerificationError(f"主程序不是 Mach-O：{launcher}")
    if architecture == "universal2" and {"arm64", "x86_64"} - launcher_architectures:
        raise MacBundleVerificationError(f"主程序不是有效 universal2：{launcher_architectures}")
    if architecture != "universal2" and architecture not in launcher_architectures:
        raise MacBundleVerificationError(f"主程序不包含 {architecture}：{launcher_architectures}")
    return len(mach_o_files)


def verify_packaged_resources(
    app_path: Path,
    *,
    template_dir: Path = DEFAULT_TEMPLATE_DIR,
    readme_path: Path = DEFAULT_README,
) -> None:
    expected_templates = {
        path.name: _sha256(path) for path in template_dir.glob("*.xlsx") if path.is_file()
    }
    if not expected_templates:
        raise MacBundleVerificationError(f"源码模板白名单为空：{template_dir}")

    packaged_spreadsheets = []
    matching_readmes = []
    for path in sorted(app_path.rglob("*")):
        if not path.is_file():
            continue
        lowered_parts = {part.lower() for part in path.relative_to(app_path).parts}
        prohibited = lowered_parts & {part.lower() for part in PROHIBITED_PATH_PARTS}
        if prohibited:
            raise MacBundleVerificationError(f"Bundle 包含禁止目录 {sorted(prohibited)}：{path}")
        if path.suffix.lower() == ".log":
            raise MacBundleVerificationError(f"Bundle 包含日志：{path}")
        if path.suffix.lower() in SPREADSHEET_SUFFIXES:
            packaged_spreadsheets.append(path)
        if path.name == "README.md" and _sha256(path) == _sha256(readme_path):
            matching_readmes.append(path)

    if not matching_readmes:
        raise MacBundleVerificationError("Bundle 缺少与仓库一致的 README.md")
    packaged_template_names = {path.name for path in packaged_spreadsheets}
    expected_template_names = set(expected_templates)
    if packaged_template_names != expected_template_names:
        missing = expected_template_names - packaged_template_names
        extra = packaged_template_names - expected_template_names
        raise MacBundleVerificationError(
            f"Bundle 模板白名单不一致；缺少 {sorted(missing)}，多出 {sorted(extra)}"
        )
    for path in packaged_spreadsheets:
        parts = path.relative_to(app_path).parts
        if not any(parts[index : index + 2] == ("hr_toolkit", "templates") for index in range(len(parts) - 1)):
            raise MacBundleVerificationError(f"Bundle 中存在模板目录外的表格：{path}")
        if _sha256(path) != expected_templates[path.name]:
            raise MacBundleVerificationError(f"Bundle 模板内容与源码不一致：{path}")


def verify_info_plist(app_path: Path, version: str) -> None:
    info_plist = app_path / "Contents" / "Info.plist"
    if not info_plist.is_file():
        raise MacBundleVerificationError(f"缺少 Info.plist：{info_plist}")
    with info_plist.open("rb") as handle:
        info = plistlib.load(handle)
    expected = {
        "CFBundleIdentifier": EXPECTED_BUNDLE_IDENTIFIER,
        "CFBundlePackageType": "APPL",
        "CFBundleShortVersionString": version,
        "CFBundleVersion": version,
    }
    for key, value in expected.items():
        if info.get(key) != value:
            raise MacBundleVerificationError(
                f"Info.plist {key} 不一致：期望 {value!r}，实际 {info.get(key)!r}"
            )


def run_headless_smoke_tests(app_path: Path, version: str) -> None:
    launcher = app_path / "Contents" / "MacOS" / "HRToolkit"
    if not launcher.is_file() or not os.access(launcher, os.X_OK):
        raise MacBundleVerificationError(f"主程序不存在或不可执行：{launcher}")
    environment = dict(os.environ)
    environment["HR_TOOLKIT_SKIP_UPDATE"] = "1"
    # PyInstaller windowed bootloader may deliberately set sys.stdout to None.
    # The application writes the same result to this CI-only path, which lets us
    # verify headless commands without weakening the GUI build.
    with tempfile.TemporaryDirectory(prefix="hr_toolkit_smoke_") as temporary:
        version_output = Path(temporary) / "version.txt"
        environment["HR_TOOLKIT_CHECK_OUTPUT"] = str(version_output)
        version_result = subprocess.run(
            [str(launcher), "--version"],
            check=False,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=environment,
            timeout=30,
        )
        if version_result.returncode != 0:
            raise MacBundleVerificationError(
                f"打包程序 --version 失败（{version_result.returncode}）：{version_result.stderr.strip()}"
            )
        if not version_output.is_file() or version_output.read_text(encoding="utf-8").strip() != version:
            fallback = version_result.stdout.strip()
            raise MacBundleVerificationError(
                f"打包程序 --version 结果不一致：文件={version_output!s}，stdout={fallback!r}"
            )

        smoke_output = Path(temporary) / "smoke.txt"
        environment["HR_TOOLKIT_CHECK_OUTPUT"] = str(smoke_output)
        smoke_result = subprocess.run(
            [str(launcher), "--smoke-test"],
            check=False,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=environment,
            timeout=60,
        )
        if smoke_result.returncode != 0:
            raise MacBundleVerificationError(
                f"打包程序 --smoke-test 失败（{smoke_result.returncode}）：{smoke_result.stderr.strip()}"
            )
        expected_smoke = f"HRToolkit {version} smoke-test OK"
        if not smoke_output.is_file() or smoke_output.read_text(encoding="utf-8").strip() != expected_smoke:
            fallback = smoke_result.stdout.strip()
            raise MacBundleVerificationError(
                f"打包程序 --smoke-test 结果不一致：文件={smoke_output!s}，stdout={fallback!r}"
            )

        update_smoke_output = Path(temporary) / "update-smoke.txt"
        environment["HR_TOOLKIT_CHECK_OUTPUT"] = str(update_smoke_output)
        update_smoke_result = subprocess.run(
            [str(launcher), "--update-smoke-test"],
            check=False,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=environment,
            timeout=90,
        )
        if update_smoke_result.returncode != 0:
            raise MacBundleVerificationError(
                "打包程序 --update-smoke-test 失败"
                f"（{update_smoke_result.returncode}）：{update_smoke_result.stderr.strip()}"
            )
        expected_update_prefix = f"HRToolkit {version} update-smoke-test OK; latest="
        if (
            not update_smoke_output.is_file()
            or not update_smoke_output.read_text(encoding="utf-8").strip().startswith(expected_update_prefix)
        ):
            fallback = update_smoke_result.stdout.strip()
            raise MacBundleVerificationError(
                "打包程序 --update-smoke-test 结果不一致："
                f"文件={update_smoke_output!s}，stdout={fallback!r}"
            )


def verify_app_bundle(
    app_path: Path,
    *,
    version: str,
    architecture: str,
    smoke_test: bool = True,
) -> int:
    if not app_path.is_dir() or app_path.suffix != ".app":
        raise MacBundleVerificationError(f"不是标准 .app Bundle：{app_path}")
    verify_info_plist(app_path, version)
    verify_packaged_resources(app_path)
    mach_o_count = verify_mach_o_architectures(app_path, architecture)
    if smoke_test:
        run_headless_smoke_tests(app_path, version)
    return mach_o_count


def verify_dmg(
    dmg_path: Path,
    *,
    version: str,
    architecture: str,
    smoke_test: bool = True,
) -> int:
    expected_name = _expected_dmg_name(version, architecture)
    if dmg_path.name != expected_name:
        raise MacBundleVerificationError(
            f"DMG 名称与真实架构不一致：期望 {expected_name}，实际 {dmg_path.name}"
        )
    if not dmg_path.is_file() or dmg_path.stat().st_size <= 0:
        raise MacBundleVerificationError(f"DMG 不存在或为空：{dmg_path}")
    _run(["hdiutil", "verify", str(dmg_path)])

    with tempfile.TemporaryDirectory(prefix="hr_toolkit_dmg_verify_") as temporary:
        mount_point = Path(temporary) / "mounted"
        mount_point.mkdir()
        attached = False
        try:
            _run(
                [
                    "hdiutil",
                    "attach",
                    "-readonly",
                    "-nobrowse",
                    "-mountpoint",
                    str(mount_point),
                    str(dmg_path),
                ]
            )
            attached = True
            app_path = mount_point / "HRToolkit.app"
            applications_link = mount_point / "Applications"
            if not applications_link.is_symlink() or os.readlink(applications_link) != "/Applications":
                raise MacBundleVerificationError("DMG 缺少指向 /Applications 的快捷方式")
            return verify_app_bundle(
                app_path,
                version=version,
                architecture=architecture,
                smoke_test=smoke_test,
            )
        finally:
            if attached:
                detach = subprocess.run(
                    ["hdiutil", "detach", str(mount_point)],
                    check=False,
                    text=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                )
                if detach.returncode != 0:
                    subprocess.run(
                        ["hdiutil", "detach", "-force", str(mount_point)],
                        check=False,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="验证 HRToolkit macOS .app 与 DMG")
    parser.add_argument("--app", type=Path, help="待验证 HRToolkit.app")
    parser.add_argument("--dmg", type=Path, help="待挂载验证的 DMG")
    parser.add_argument("--version", required=True)
    parser.add_argument("--architecture", required=True, choices=sorted(ARCHITECTURES))
    parser.add_argument("--skip-smoke-test", action="store_true", help="仅用于无法运行目标架构的交叉检查")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if bool(args.app) == bool(args.dmg):
        raise SystemExit("必须且只能指定 --app 或 --dmg")
    if args.app:
        count = verify_app_bundle(
            args.app,
            version=args.version,
            architecture=args.architecture,
            smoke_test=not args.skip_smoke_test,
        )
    else:
        count = verify_dmg(
            args.dmg,
            version=args.version,
            architecture=args.architecture,
            smoke_test=not args.skip_smoke_test,
        )
    print(f"macOS 验证通过：{count} 个 Mach-O，架构 {args.architecture}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
