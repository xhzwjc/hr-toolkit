from __future__ import annotations

import argparse
import importlib.util
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional, Sequence


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from generate_app_icons import encode_png, render_icon  # noqa: E402
from generate_release_metadata import (  # noqa: E402
    read_project_version,
    require_release_asset_under_limit,
    validate_version,
)
from verify_macos_bundle import (  # noqa: E402
    EXPECTED_DMG_FORMAT,
    QT_IMAGE_FORMAT_PLUGIN_NAMES,
    verify_app_bundle,
    verify_dmg,
)


ARCHITECTURES = ("universal2", "x86_64", "arm64")
ARCH_SUFFIXES = {"universal2": "universal", "x86_64": "x64", "arm64": "arm64"}
DEFAULT_ENTITLEMENTS = REPO_ROOT / "packaging" / "macos" / "entitlements.plist"
QML_DIR = REPO_ROOT / "hr_toolkit" / "gui_qt" / "qml"
QT_NOTICE = REPO_ROOT / "packaging" / "qt" / "THIRD-PARTY-NOTICES.txt"
QT_HOOKS_DIR = REPO_ROOT / "packaging" / "qt" / "hooks"
QT_TRANSLATION_SUFFIXES = ("_en.qm", "_zh_CN.qm", "_zh_TW.qm")
HDIUTIL_BUSY_RETRY_DELAYS = (2.0, 5.0)
MACOS_BUILD_MODULES = {
    "PyInstaller": "pyinstaller",
    "PySide6.QtQml": "PySide6_Essentials",
    "PySide6.QtQuick": "PySide6_Essentials",
    "PySide6.QtQuickControls2": "PySide6_Essentials",
    "cv2": "opencv-python",
    "onnxruntime": "onnxruntime",
    "PIL.Image": "Pillow",
    "rapidocr_onnxruntime": "rapidocr-onnxruntime",
}


class MacBuildError(RuntimeError):
    """Raised when the native macOS build cannot be produced safely."""


def ensure_build_dependencies() -> None:
    missing = []
    for module in MACOS_BUILD_MODULES:
        try:
            available = importlib.util.find_spec(module) is not None
        except ModuleNotFoundError:
            available = False
        if not available:
            missing.append(module)
    if missing:
        packages = sorted({MACOS_BUILD_MODULES[module] for module in missing})
        raise MacBuildError(
            "macOS 打包环境缺少依赖模块："
            + ", ".join(missing)
            + "。请安装："
            + " ".join(packages)
        )


def _run(command: Sequence[str], *, cwd: Path = REPO_ROOT, capture: bool = False) -> subprocess.CompletedProcess:
    print("执行：" + " ".join(str(part) for part in command))
    result = subprocess.run(
        list(command),
        cwd=str(cwd),
        check=False,
        text=True,
        stdout=subprocess.PIPE if capture else None,
        stderr=subprocess.PIPE if capture else None,
    )
    if result.returncode != 0:
        detail = ""
        if capture:
            detail = (result.stderr or result.stdout or "").strip()
        raise MacBuildError(
            f"命令失败（{result.returncode}）：{' '.join(str(part) for part in command)}"
            + (f"\n{detail}" if detail else "")
        )
    return result


def _run_hdiutil_with_busy_retry(
    command: Sequence[str],
    *,
    cleanup_paths: Sequence[Path],
) -> subprocess.CompletedProcess:
    """Retry only hdiutil's transient disk-image service contention."""

    for attempt in range(len(HDIUTIL_BUSY_RETRY_DELAYS) + 1):
        for path in cleanup_paths:
            path.unlink(missing_ok=True)
        try:
            return _run(command, capture=True)
        except MacBuildError as exc:
            is_resource_busy = "resource busy" in str(exc).casefold()
            if not is_resource_busy or attempt >= len(HDIUTIL_BUSY_RETRY_DELAYS):
                raise
            delay = HDIUTIL_BUSY_RETRY_DELAYS[attempt]
            print(
                "hdiutil 磁盘镜像服务暂时繁忙，"
                f"{delay:g} 秒后重试（{attempt + 2}/{len(HDIUTIL_BUSY_RETRY_DELAYS) + 1}）"
            )
            time.sleep(delay)

    raise AssertionError("unreachable")


def _safe_clean_directory(path: Path) -> None:
    resolved = path.resolve()
    allowed_root = (REPO_ROOT / "build").resolve()
    if resolved == allowed_root or allowed_root not in resolved.parents:
        raise MacBuildError(f"拒绝清理非 build 子目录：{resolved}")
    if resolved.exists():
        shutil.rmtree(resolved)
    resolved.mkdir(parents=True, exist_ok=True)


def _verify_universal_python() -> None:
    result = _run(["lipo", "-archs", str(Path(sys.executable).resolve())], capture=True)
    architectures = set(result.stdout.split())
    missing = {"arm64", "x86_64"} - architectures
    if missing:
        raise MacBuildError(
            f"universal2 构建需要 universal2 Python；{sys.executable} 缺少 {sorted(missing)}"
        )


def _write_icns(iconset_dir: Path, output_path: Path) -> None:
    iconset_dir.mkdir(parents=True, exist_ok=True)
    icon_entries = (
        (16, "icon_16x16.png"),
        (32, "icon_16x16@2x.png"),
        (32, "icon_32x32.png"),
        (64, "icon_32x32@2x.png"),
        (128, "icon_128x128.png"),
        (256, "icon_128x128@2x.png"),
        (256, "icon_256x256.png"),
        (512, "icon_256x256@2x.png"),
        (512, "icon_512x512.png"),
        (1024, "icon_512x512@2x.png"),
    )
    rendered = {}
    for size, name in icon_entries:
        if size not in rendered:
            rendered[size] = encode_png(render_icon(size))
        payload = rendered[size]
        (iconset_dir / name).write_bytes(payload)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    _run(["iconutil", "--convert", "icns", "--output", str(output_path), str(iconset_dir)])


def _write_spec(
    spec_path: Path,
    *,
    architecture: str,
    version: str,
    icon_path: Path,
    codesign_identity: Optional[str],
    entitlements_file: Optional[Path],
) -> None:
    entrypoint = REPO_ROOT / "hr_toolkit_app.py"
    readme = REPO_ROOT / "README.md"
    templates = REPO_ROOT / "hr_toolkit" / "templates"
    template_files = sorted(path for path in templates.glob("*.xlsx") if path.is_file())
    if not template_files:
        raise MacBuildError(f"模板白名单为空：{templates}")
    datas = [(str(readme), ".")]
    datas.extend((str(path), "hr_toolkit/templates") for path in template_files)
    if QML_DIR.is_dir() and (QML_DIR / "Main.qml").is_file():
        datas.append((str(QML_DIR), "hr_toolkit/gui_qt/qml"))
    if QT_NOTICE.is_file() and QT_HOOKS_DIR.is_dir():
        datas.append((str(QT_NOTICE), "third_party/qt"))
    spec = f'''# Generated by scripts/build_macos.py; do not commit build output.
from PyInstaller.utils.hooks import collect_all, copy_metadata
_ocr_datas, _ocr_binaries, _ocr_hidden = collect_all("rapidocr_onnxruntime")
_sevenzip_datas, _sevenzip_binaries, _sevenzip_hidden = collect_all("py7zr")
_rar_datas, _rar_binaries, _rar_hidden = collect_all("unrar")
_pypdf_metadata = copy_metadata("pypdf")
a = Analysis(
    [{str(entrypoint)!r}],
    pathex=[{str(REPO_ROOT)!r}],
    binaries=_ocr_binaries + _sevenzip_binaries + _rar_binaries,
    datas={datas!r} + _ocr_datas + _sevenzip_datas + _rar_datas + _pypdf_metadata,
    hiddenimports=[
        "xlrd",
        "pypdf",
        "PySide6.QtQml",
        "PySide6.QtWidgets",
    ] + _ocr_hidden + _sevenzip_hidden + _rar_hidden,
    hookspath=[{str(QT_HOOKS_DIR)!r}],
    hooksconfig={{}},
    excludes=[
        "pytest",
        "unittest",
        "test",
        "tests",
        "tkinter.test",
        "sqlite3.test",
        "ctypes.test",
        "idlelib",
        "pydoc",
        "pdb",
        "turtle",
        "doctest",
        "lib2to3",
        # The release application is Qt Quick-only.  The legacy Tk renderer
        # remains available when running from source, but shipping a second GUI
        # runtime adds several megabytes and cannot be reached in normal use.
        "tkinter",
        "hr_toolkit.gui",
        "PIL.ImageGrab",
        "PIL.ImageTk",
        "PIL._imagingtk",
        "PIL.AvifImagePlugin",
        "PIL._avif",
    ],
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="HRToolkit",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    argv_emulation=False,
    target_arch={architecture!r},
    codesign_identity={codesign_identity!r},
    entitlements_file={(str(entitlements_file) if entitlements_file else None)!r},
    icon={str(icon_path)!r},
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="HRToolkit",
)
app = BUNDLE(
    coll,
    name="HRToolkit.app",
    icon={str(icon_path)!r},
    bundle_identifier="com.xhzwjc.hrtoolkit",
    info_plist={{
        "CFBundleDisplayName": "HR Toolkit",
        "CFBundleName": "HRToolkit",
        "CFBundlePackageType": "APPL",
        "CFBundleShortVersionString": {version!r},
        "CFBundleVersion": {version!r},
        "LSApplicationCategoryType": "public.app-category.business",
        "NSHighResolutionCapable": True,
    }},
)
'''
    spec_path.write_text(spec, encoding="utf-8")


def _create_dmg(app_path: Path, dmg_path: Path, staging_dir: Path, version: str) -> None:
    if staging_dir.exists():
        shutil.rmtree(staging_dir)
    staging_dir.mkdir(parents=True)
    staged_app = staging_dir / "HRToolkit.app"
    _run(["ditto", str(app_path), str(staged_app)])
    (staging_dir / "Applications").symlink_to("/Applications", target_is_directory=True)
    dmg_path.parent.mkdir(parents=True, exist_ok=True)
    if dmg_path.exists():
        dmg_path.unlink()
    uncompressed_dmg = staging_dir.parent / f"{dmg_path.stem}.uncompressed.dmg"
    uncompressed_dmg.unlink(missing_ok=True)
    try:
        _run_hdiutil_with_busy_retry(
            [
                "hdiutil",
                "create",
                "-ov",
                "-format",
                "UDRO",
                "-volname",
                f"HR Toolkit {version}",
                "-srcfolder",
                str(staging_dir),
                str(uncompressed_dmg),
            ],
            cleanup_paths=(uncompressed_dmg,),
        )
        _run_hdiutil_with_busy_retry(
            [
                "hdiutil",
                "convert",
                str(uncompressed_dmg),
                "-format",
                EXPECTED_DMG_FORMAT,
                "-o",
                str(dmg_path),
            ],
            cleanup_paths=(dmg_path,),
        )
    finally:
        uncompressed_dmg.unlink(missing_ok=True)


def remove_unused_qt_translations(app_path: Path) -> int:
    """Keep only English and Chinese Qt strings used by the Chinese desktop app."""

    translations = app_path / "Contents" / "Resources" / "PySide6" / "Qt" / "translations"
    removed_bytes = 0
    if not translations.is_dir():
        return removed_bytes
    for path in sorted(translations.glob("*.qm")):
        if path.name.endswith(QT_TRANSLATION_SUFFIXES):
            continue
        if path.is_symlink() or not path.is_file():
            raise MacBuildError(f"拒绝移除非普通 Qt 翻译文件：{path}")
        removed_bytes += path.stat().st_size
        path.unlink()
    return removed_bytes


def remove_qt_development_plugins(app_path: Path) -> int:
    """Remove QML debugger/profiler plugins that production never enables."""

    plugin_dir = (
        app_path
        / "Contents"
        / "Frameworks"
        / "PySide6"
        / "Qt"
        / "plugins"
        / "qmltooling"
    )
    if not plugin_dir.exists():
        return 0
    if plugin_dir.is_symlink() or not plugin_dir.is_dir():
        raise MacBuildError(f"拒绝移除异常 Qt 开发插件路径：{plugin_dir}")
    files = [path for path in plugin_dir.rglob("*") if path.is_file()]
    removed_bytes = sum(path.stat().st_size for path in files)
    shutil.rmtree(plugin_dir)
    return removed_bytes


def remove_unused_qt_image_format_plugins(app_path: Path) -> int:
    """Keep only codecs for image formats accepted by the business tools."""

    plugin_dir = (
        app_path
        / "Contents"
        / "Frameworks"
        / "PySide6"
        / "Qt"
        / "plugins"
        / "imageformats"
    )
    if not plugin_dir.is_dir():
        return 0
    removed_bytes = 0
    for path in sorted(plugin_dir.glob("*.dylib")):
        if path.name in QT_IMAGE_FORMAT_PLUGIN_NAMES:
            continue
        if path.is_symlink() or not path.is_file():
            raise MacBuildError(f"拒绝移除异常 Qt 图片格式插件：{path}")
        removed_bytes += path.stat().st_size
        path.unlink()
    return removed_bytes


def build_macos(
    *,
    version: str,
    architecture: str,
    output_dir: Path,
    work_dir: Path,
    codesign_identity: Optional[str] = None,
    entitlements_file: Optional[Path] = None,
) -> Path:
    if sys.platform != "darwin":
        raise MacBuildError("macOS .app/DMG 必须在 macOS 上构建")
    validate_version(version)
    project_version = read_project_version()
    if project_version != version:
        raise MacBuildError(
            f"构建版本 {version} 与 hr_toolkit.__version__ {project_version} 不一致"
        )
    if architecture not in ARCHITECTURES:
        raise MacBuildError(f"未知架构：{architecture}")
    ensure_build_dependencies()
    if not work_dir.is_absolute():
        work_dir = REPO_ROOT / work_dir
    if not output_dir.is_absolute():
        output_dir = REPO_ROOT / output_dir
    work_dir = work_dir.resolve()
    output_dir = output_dir.resolve()
    if shutil.which("hdiutil") is None or shutil.which("iconutil") is None or shutil.which("lipo") is None:
        raise MacBuildError("缺少 hdiutil/iconutil/lipo，无法完成原生 macOS 构建验证")
    if architecture == "universal2":
        _verify_universal_python()
    if codesign_identity and entitlements_file is None:
        entitlements_file = DEFAULT_ENTITLEMENTS
    if entitlements_file is not None and not entitlements_file.is_file():
        raise MacBuildError(f"签名 entitlements 文件不存在：{entitlements_file}")

    _safe_clean_directory(work_dir)
    pyinstaller_build = work_dir / "pyinstaller-build"
    pyinstaller_dist = work_dir / "pyinstaller-dist"
    iconset_dir = work_dir / "HRToolkit.iconset"
    icon_path = work_dir / "HRToolkit.icns"
    spec_path = work_dir / "HRToolkit.spec"
    staging_dir = work_dir / "dmg-staging"
    _write_icns(iconset_dir, icon_path)
    _write_spec(
        spec_path,
        architecture=architecture,
        version=version,
        icon_path=icon_path,
        codesign_identity=codesign_identity,
        entitlements_file=entitlements_file,
    )

    _run(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--workpath",
            str(pyinstaller_build),
            "--distpath",
            str(pyinstaller_dist),
            str(spec_path),
        ]
    )
    app_path = pyinstaller_dist / "HRToolkit.app"
    removed_translation_bytes = remove_unused_qt_translations(app_path)
    if removed_translation_bytes:
        print(f"已移除未使用的 Qt 翻译：{removed_translation_bytes} 字节")
    removed_development_bytes = remove_qt_development_plugins(app_path)
    if removed_development_bytes:
        print(f"已移除 Qt QML 开发插件：{removed_development_bytes} 字节")
    removed_image_plugin_bytes = remove_unused_qt_image_format_plugins(app_path)
    if removed_image_plugin_bytes:
        print(f"已移除未使用的 Qt 图片格式插件：{removed_image_plugin_bytes} 字节")
    mach_o_count = verify_app_bundle(
        app_path,
        version=version,
        architecture=architecture,
        smoke_test=True,
    )
    print(f".app 验证通过：{mach_o_count} 个 Mach-O")

    suffix = ARCH_SUFFIXES[architecture]
    dmg_path = output_dir / f"HRToolkit_{version}_{suffix}.dmg"
    _create_dmg(app_path, dmg_path, staging_dir, version)
    dmg_mach_o_count = verify_dmg(
        dmg_path,
        version=version,
        architecture=architecture,
        smoke_test=True,
    )
    print(f"DMG 验证通过：{dmg_mach_o_count} 个 Mach-O")
    dmg_size = require_release_asset_under_limit(dmg_path)
    print(f"DMG 体积门禁通过：{dmg_size} 字节")
    if codesign_identity:
        _run(["codesign", "--verify", "--deep", "--strict", str(app_path)])
    return dmg_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="构建并验证 HRToolkit 标准 macOS .app/DMG")
    parser.add_argument("--version", required=True)
    parser.add_argument("--architecture", choices=ARCHITECTURES, default="universal2")
    parser.add_argument("--output-dir", type=Path, default=REPO_ROOT / "dist" / "release-assets")
    parser.add_argument("--work-dir", type=Path, help="必须位于仓库 build/ 下")
    parser.add_argument(
        "--codesign-identity",
        default=os.environ.get("MACOS_CODESIGN_IDENTITY") or None,
        help="未来 Developer ID 签名入口；当前留空，仅使用 PyInstaller 必需的 ad-hoc 签名",
    )
    parser.add_argument(
        "--entitlements-file",
        type=Path,
        default=Path(os.environ["MACOS_ENTITLEMENTS_FILE"]) if os.environ.get("MACOS_ENTITLEMENTS_FILE") else None,
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    work_dir = args.work_dir or (REPO_ROOT / "build" / "macos" / args.architecture)
    dmg_path = build_macos(
        version=args.version,
        architecture=args.architecture,
        output_dir=args.output_dir,
        work_dir=work_dir,
        codesign_identity=args.codesign_identity,
        entitlements_file=args.entitlements_file,
    )
    print(f"已生成 macOS 安装包：{dmg_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
