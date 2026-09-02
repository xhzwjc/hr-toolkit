from __future__ import annotations

import argparse
import importlib.metadata as importlib_metadata
import importlib.util
import os
import platform
import shutil
import struct
import subprocess
import sys
import tempfile
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from versioning import read_project_version


APP_NAME = "HRToolkit"
UPDATER_NAME = "HRToolkitUpdater"
APP_ENTRYPOINT = REPO_ROOT / "hr_toolkit_app.py"
UPDATER_ENTRYPOINT = REPO_ROOT / "hr_toolkit_updater.py"
WINDOWS_MANIFEST = REPO_ROOT / "packaging" / "windows" / "HRToolkit.manifest"
WINDOWS_WIN7_MANIFEST = REPO_ROOT / "packaging" / "windows" / "HRToolkit.win7.manifest"
WINDOWS_ICON = REPO_ROOT / "packaging" / "windows" / "HRToolkit.ico"
WIN7_THIRD_PARTY_NOTICE = (
    REPO_ROOT / "packaging" / "windows" / "win7" / "THIRD-PARTY-NOTICES.txt"
)
WIN7_PYINSTALLER_HOOKS_DIR = REPO_ROOT / "packaging" / "windows" / "win7" / "hooks"
WIN7_PYINSTALLER_HOOK = WIN7_PYINSTALLER_HOOKS_DIR / "hook-hr_toolkit.py"
README_FILE = REPO_ROOT / "README.md"
TEMPLATES_DIR = REPO_ROOT / "hr_toolkit" / "templates"
QML_DIR = REPO_ROOT / "hr_toolkit" / "gui_qt" / "qml"
QT_NOTICE = REPO_ROOT / "packaging" / "qt" / "THIRD-PARTY-NOTICES.txt"
QT_PYINSTALLER_HOOKS_DIR = REPO_ROOT / "packaging" / "qt" / "hooks"
# PyInstaller deliberately freezes PySide beside its Windows wheel layout:
# ``PySide6/...`` and ``PySide2/...``.  The extra ``Qt`` directory is used by
# PySide wheels on macOS/Linux only; assuming that layout on Windows makes a
# valid Qt Quick payload look empty and skips production pruning.
QT6_WINDOWS_RUNTIME_ROOT = Path("PySide6")
QT5_WINDOWS_RUNTIME_ROOT = Path("PySide2")

QT6_REQUIRED_QML_FILES = (
    "QtCore/qmldir",
    "QtQml/qmldir",
    "QtQml/Models/qmldir",
    "QtQml/WorkerScript/qmldir",
    "QtQml/WorkerScript/workerscriptplugin.dll",
    "QtQuick/qmldir",
    "QtQuick/Layouts/qmldir",
    "QtQuick/Templates/qmldir",
    "QtQuick/Window/qmldir",
    "QtQuick/Controls/qmldir",
    "QtQuick/Controls/Basic/qmldir",
    "QtQuick/Controls/Basic/Button.qml",
)
QT5_REQUIRED_QML_FILES = (
    "QtQml/qmldir",
    "QtQml/Models.2/qmldir",
    "QtQml/WorkerScript.2/qmldir",
    "QtQml/WorkerScript.2/workerscriptplugin.dll",
    "QtQuick.2/qmldir",
    "QtQuick/Layouts/qmldir",
    "QtQuick/Templates.2/qmldir",
    "QtQuick/Window.2/qmldir",
    "QtQuick/Controls.2/qmldir",
    "QtQuick/Controls.2/Button.qml",
)
QT5_REQUIRED_RUNTIME_FILES = (
    "PySide2/Qt5Core.dll",
    "PySide2/Qt5Gui.dll",
    "PySide2/Qt5Qml.dll",
    "PySide2/Qt5QmlWorkerScript.dll",
    "PySide2/Qt5Quick.dll",
    "PySide2/Qt5QuickControls2.dll",
    "PySide2/QtCore.pyd",
    "PySide2/QtGui.pyd",
    "PySide2/QtQml.pyd",
    "PySide2/d3dcompiler_47.dll",
    "PySide2/libEGL.dll",
    "PySide2/libGLESv2.dll",
    "PySide2/plugins/platforms/qoffscreen.dll",
    "PySide2/plugins/platforms/qwindows.dll",
)

WINDOWS_BUILD_MODULES = {
    "PyInstaller": "pyinstaller",
    "PySide6.QtQml": "PySide6_Essentials",
    "PySide6.QtQuick": "PySide6_Essentials",
    "PySide6.QtQuickControls2": "PySide6_Essentials",
    "certifi": "certifi",
    "cv2": "opencv-python",
    "onnxruntime": "onnxruntime",
    "openpyxl": "openpyxl",
    "PIL.Image": "Pillow",
    "py7zr": "py7zr",
    "rapidocr_onnxruntime": "rapidocr-onnxruntime",
    "pypdf": "pypdf",
    "unrar.cffi.rarfile": "unrar2-cffi",
    "xlrd": "xlrd",
    "pythoncom": "pywin32",
    "pywintypes": "pywin32",
    "win32com.client": "pywin32",
    "win32timezone": "pywin32",
}
WINDOWS_WIN7_BUILD_MODULES = {
    "PyInstaller": "pyinstaller",
    "PySide2.QtQml": "PySide2",
    "PySide2.QtQuick": "PySide2",
    "PySide2.QtQuickControls2": "PySide2",
    "certifi": "certifi",
    "cv2": "opencv-python",
    "onnxruntime": "onnxruntime",
    "openpyxl": "openpyxl",
    "PIL.Image": "Pillow",
    "pefile": "pefile",
    "py7zr": "py7zr",
    "rapidocr_onnxruntime": "rapidocr-onnxruntime",
    "pypdfium2": "pypdfium2",
    "xlrd": "xlrd",
    "pythoncom": "pywin32",
    "pywintypes": "pywin32",
    "win32com.client": "pywin32",
    "win32timezone": "pywin32",
}
WIN7_PINNED_DISTRIBUTIONS = {
    "numpy": "1.24.4",
    "onnxruntime": "1.11.1",
    "opencv-python": "4.8.1.78",
    "Pillow": "10.4.0",
    "PyInstaller": "6.21.0",
    "PySide2": "5.15.2.1",
    "py7zr": "0.22.0",
    "pypdfium2": "4.27.0",
    "pywin32": "306",
    "rapidocr-onnxruntime": "1.4.4",
    "shiboken2": "5.15.2.1",
}
HIDDEN_IMPORTS = (
    "PySide6.QtQml",
    "PySide6.QtWidgets",
    "pythoncom",
    "pywintypes",
    "win32com.client",
    "win32timezone",
    "py7zr",
    "pypdf",
    "unrar.cffi.rarfile",
    "unrar.cffi.unrarlib",
    "xlrd",
)
WIN7_HIDDEN_IMPORTS = (
    "PySide2.QtQml",
    "PySide2.QtWidgets",
    "pythoncom",
    "pywintypes",
    "win32com.client",
    "win32timezone",
    "py7zr",
    "pypdfium2",
    "xlrd",
)
COLLECT_ALL_MODULES = (
    "rapidocr_onnxruntime",
    "py7zr",
    "unrar",
)
WIN7_COLLECT_ALL_MODULES = (
    "rapidocr_onnxruntime",
    "py7zr",
    "pypdfium2",
    "pypdfium2_raw",
)
EXCLUDED_MODULES = (
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
)
# Release applications are Qt Quick-only.  Keep the legacy Tk renderer
# available in source checkouts, while avoiding a second GUI runtime in every
# installed main application.  The standalone updater still uses Tk and must
# therefore use EXCLUDED_MODULES rather than this extended tuple.
MAIN_APP_EXCLUDED_MODULES = (
    *EXCLUDED_MODULES,
    "tkinter",
    "hr_toolkit.gui",
    "PIL.ImageGrab",
    "PIL.ImageTk",
    "PIL._imagingtk",
    "PIL.AvifImagePlugin",
    "PIL._avif",
)
RELEASE_TEMPLATE_NAMES = (
    "archive_company_template.xlsx",
    "archive_summary_template.xlsx",
    "data_statistics_template.xlsx",
    "insurance_ledger_template.xlsx",
    "personnel_change_summary_template.xlsx",
    "social_security_detail_template.xlsx",
    "social_security_summary_template.xlsx",
)
PE_MACHINE_AMD64 = 0x8664
WINDOWS_TARGET_MODERN = "modern"
WINDOWS_TARGET_WIN7 = "win7"
WINDOWS_TARGETS = (WINDOWS_TARGET_MODERN, WINDOWS_TARGET_WIN7)
WIN7_QT_SMOKE_ENV = {
    # Exercise the bundled Qt/QML runtime without coupling unattended build
    # verification to a hosted runner's interactive desktop or graphics stack.
    # These values are passed only to the smoke subprocess, never embedded in
    # the installed application.
    "QT_QPA_PLATFORM": "offscreen",
    "QT_QUICK_BACKEND": "software",
    "QSG_RENDER_LOOP": "basic",
}
WIN7_REQUIRED_UCRT_FILES = (
    "api-ms-win-core-console-l1-1-0.dll",
    "api-ms-win-core-datetime-l1-1-0.dll",
    "api-ms-win-core-debug-l1-1-0.dll",
    "api-ms-win-core-errorhandling-l1-1-0.dll",
    "api-ms-win-core-file-l1-1-0.dll",
    "api-ms-win-core-file-l1-2-0.dll",
    "api-ms-win-core-file-l2-1-0.dll",
    "api-ms-win-core-handle-l1-1-0.dll",
    "api-ms-win-core-heap-l1-1-0.dll",
    "api-ms-win-core-interlocked-l1-1-0.dll",
    "api-ms-win-core-libraryloader-l1-1-0.dll",
    "api-ms-win-core-localization-l1-2-0.dll",
    "api-ms-win-core-memory-l1-1-0.dll",
    "api-ms-win-core-namedpipe-l1-1-0.dll",
    "api-ms-win-core-processenvironment-l1-1-0.dll",
    "api-ms-win-core-processthreads-l1-1-0.dll",
    "api-ms-win-core-processthreads-l1-1-1.dll",
    "api-ms-win-core-profile-l1-1-0.dll",
    "api-ms-win-core-rtlsupport-l1-1-0.dll",
    "api-ms-win-core-string-l1-1-0.dll",
    "api-ms-win-core-synch-l1-1-0.dll",
    "api-ms-win-core-synch-l1-2-0.dll",
    "api-ms-win-core-sysinfo-l1-1-0.dll",
    "api-ms-win-core-timezone-l1-1-0.dll",
    "api-ms-win-core-util-l1-1-0.dll",
    "api-ms-win-crt-conio-l1-1-0.dll",
    "api-ms-win-crt-convert-l1-1-0.dll",
    "api-ms-win-crt-environment-l1-1-0.dll",
    "api-ms-win-crt-filesystem-l1-1-0.dll",
    "api-ms-win-crt-heap-l1-1-0.dll",
    "api-ms-win-crt-locale-l1-1-0.dll",
    "api-ms-win-crt-math-l1-1-0.dll",
    "api-ms-win-crt-multibyte-l1-1-0.dll",
    "api-ms-win-crt-private-l1-1-0.dll",
    "api-ms-win-crt-process-l1-1-0.dll",
    "api-ms-win-crt-runtime-l1-1-0.dll",
    "api-ms-win-crt-stdio-l1-1-0.dll",
    "api-ms-win-crt-string-l1-1-0.dll",
    "api-ms-win-crt-time-l1-1-0.dll",
    "api-ms-win-crt-utility-l1-1-0.dll",
    "ucrtbase.dll",
)
WIN7_REQUIRED_UCRT_FILE_KEYS = frozenset(
    name.casefold() for name in WIN7_REQUIRED_UCRT_FILES
)
WIN7_REQUIRED_7ZIP_FILES = ("7z.exe", "7z.dll", "License.txt")
WIN7_7ZIP_OVERRIDE_ENV = "HR_TOOLKIT_7ZIP_EXE"
WIN7_REQUIRED_VC_RUNTIME_FILES = (
    "msvcp140.dll",
    "vcruntime140.dll",
    "vcruntime140_1.dll",
)
WIN7_REQUIRED_PYTHON_RUNTIME_FILES = (
    "python38.dll",
    "python3.dll",
)
WIN7_FORBIDDEN_DLLS = frozenset(
    {
        "api-ms-win-core-path-l1-1-0.dll",
        "api-ms-win-core-path-l1-1-1.dll",
    }
)
WIN7_API_SET_PREFIXES = ("api-ms-win-", "ext-ms-win-")
WIN7_FORBIDDEN_IMPORTS = frozenset(
    {
        "copyfile2",
        "createfile2",
        "getcurrentpackagefullname",
        "getcurrentpackageid",
        "getfileinformationbyname",
        "getoverlappedresultex",
        "getpackagefullname",
        "getpackagepathbyfullname",
        "getsystemtimeadjustmentprecise",
        "getsystemtimepreciseasfiletime",
        "getthreaddescription",
        "iswow64process2",
        "setprocessmitigationpolicy",
        "setsystemtimeadjustmentprecise",
        "setthreaddescription",
        "waitonaddress",
        "wakebyaddressall",
        "wakebyaddresssingle",
    }
)
FORBIDDEN_PAYLOAD_PARTS = {
    "__pycache__",
    ".pytest_cache",
    "tests",
    "test",
    "outputs",
    "output",
    ".hrtoolkit",
    "上传资料",
    "处理结果",
    "补充资料",
    "共用资料",
    "附件",
    "二期新增的附件",
    "问题汇总",
    "二期问题汇总表",
    "问题1-3相关数据及模板",
    "模板",
}
FORBIDDEN_DATA_SUFFIXES = {".db", ".sqlite", ".sqlite3"}
FORBIDDEN_DATA_FILENAMES = {
    "history.db-wal",
    "history.db-shm",
    "history.db-journal",
    "history.db.bak",
    "history.db.backup",
    ".hrtoolkit-data-v1",
    ".archive.lock",
    ".manifest.lock",
    ".database-access.lock",
    ".database-recovery-pending.json",
    ".project.lock",
    "project-write.lock",
}
OPENCV_VIDEOIO_FFMPEG_PATTERN = "opencv_videoio_ffmpeg*.dll"
QT_TRANSLATION_SUFFIXES = ("_en.qm", "_zh_CN.qm", "_zh_TW.qm")
QT_IMAGE_FORMAT_PLUGIN_NAMES = frozenset({"qgif.dll", "qjpeg.dll", "qwebp.dll"})


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="在 Windows x64 上纯构建 HRToolkit onedir 程序和 onefile Updater。"
    )
    parser.add_argument("--version", default=read_project_version(), help="必须与 hr_toolkit.__version__ 一致")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=REPO_ROOT / "dist" / "windows",
        help="PyInstaller 输出目录",
    )
    parser.add_argument(
        "--work-dir",
        type=Path,
        default=REPO_ROOT / "build" / "windows",
        help="PyInstaller 临时工作目录",
    )
    parser.add_argument(
        "--skip-runtime-smoke",
        action="store_true",
        help="仅供诊断；跳过打包后可执行文件的无界面启动检查",
    )
    parser.add_argument(
        "--target",
        choices=WINDOWS_TARGETS,
        default=WINDOWS_TARGET_MODERN,
        help="modern 保持现有构建；win7 生成 Windows 7 SP1 x64 兼容构建",
    )
    parser.add_argument(
        "--seven-zip-dir",
        type=Path,
        help="win7 构建所需的官方 7-Zip x64 目录（7z.exe/7z.dll/License.txt）",
    )
    parser.add_argument(
        "--ucrt-dir",
        type=Path,
        help="win7 构建所需的 Windows SDK app-local UCRT x64 目录",
    )
    parser.add_argument(
        "--vc-runtime-dir",
        type=Path,
        help="win7 构建所需的 Visual C++ 2015-2019 app-local x64 目录",
    )
    args = parser.parse_args(argv)

    version = validate_build_version(args.version)
    target = validate_windows_target(args.target)
    seven_zip_dir, ucrt_dir, vc_runtime_dir = validate_win7_runtime_sources(
        target=target,
        seven_zip_dir=args.seven_zip_dir,
        ucrt_dir=args.ucrt_dir,
        vc_runtime_dir=args.vc_runtime_dir,
    )
    ensure_windows_x64_build_environment(target)
    ensure_build_dependencies(target)

    app_dir, updater = build_windows_binaries(
        version=version,
        output_dir=args.output_dir.resolve(),
        work_dir=args.work_dir.resolve(),
        target=target,
        seven_zip_dir=seven_zip_dir,
        ucrt_dir=ucrt_dir,
        vc_runtime_dir=vc_runtime_dir,
    )
    verify_windows_payload(app_dir, target=target)
    verify_pe_x64(app_dir / f"{APP_NAME}.exe")
    verify_pe_x64(updater)
    if target == WINDOWS_TARGET_WIN7:
        verify_win7_pe_compatibility((*_payload_pe_files(app_dir), updater))
    if not args.skip_runtime_smoke:
        run_runtime_smoke(
            app_dir / f"{APP_NAME}.exe",
            updater,
            target=target,
        )

    print(f"Windows 程序目录：{app_dir}")
    print(f"Windows 更新程序：{updater}")
    return 0


def validate_build_version(version: str) -> str:
    version = version.strip()
    validate_stable_semver(version)
    project_version = read_project_version()
    if version != project_version:
        raise ValueError(
            f"构建版本 {version} 与 hr_toolkit.__version__ {project_version} 不一致。"
        )
    return version


def validate_stable_semver(version: str) -> tuple[int, int, int]:
    parts = version.split(".")
    if len(parts) != 3 or any(not part.isdigit() for part in parts):
        raise ValueError(f"版本号必须是稳定 SemVer x.y.z：{version}")
    if any(len(part) > 1 and part.startswith("0") for part in parts):
        raise ValueError(f"版本号不能包含前导零：{version}")
    return int(parts[0]), int(parts[1]), int(parts[2])


def validate_windows_target(target: str) -> str:
    if target not in WINDOWS_TARGETS:
        raise ValueError(f"未知 Windows 构建目标：{target}")
    return target


def windows_asset_suffix(target: str = WINDOWS_TARGET_MODERN) -> str:
    target = validate_windows_target(target)
    return "win7_x64" if target == WINDOWS_TARGET_WIN7 else "x64"


def windows_setup_asset_name(
    version: str,
    target: str = WINDOWS_TARGET_MODERN,
) -> str:
    validate_stable_semver(version)
    return f"HRToolkit_{version}_{windows_asset_suffix(target)}-setup.exe"


def windows_msi_asset_name(
    version: str,
    target: str = WINDOWS_TARGET_MODERN,
) -> str:
    validate_stable_semver(version)
    return f"HRToolkit_{version}_{windows_asset_suffix(target)}.msi"


def validate_win7_runtime_sources(
    *,
    target: str,
    seven_zip_dir: Path | None,
    ucrt_dir: Path | None,
    vc_runtime_dir: Path | None = None,
) -> tuple[Path | None, Path | None, Path | None]:
    target = validate_windows_target(target)
    if target != WINDOWS_TARGET_WIN7:
        return None, None, None
    if seven_zip_dir is None or ucrt_dir is None or vc_runtime_dir is None:
        raise ValueError(
            "Windows 7 构建必须提供 --seven-zip-dir、--ucrt-dir 和 --vc-runtime-dir。"
        )

    resolved_7zip = seven_zip_dir.resolve()
    resolved_ucrt = ucrt_dir.resolve()
    resolved_vc_runtime = vc_runtime_dir.resolve()
    _require_files(resolved_7zip, WIN7_REQUIRED_7ZIP_FILES, label="7-Zip")
    _require_files(resolved_ucrt, WIN7_REQUIRED_UCRT_FILES, label="app-local UCRT")
    _require_files(
        resolved_vc_runtime,
        WIN7_REQUIRED_VC_RUNTIME_FILES,
        label="Visual C++ app-local runtime",
    )
    if not WIN7_THIRD_PARTY_NOTICE.is_file():
        raise RuntimeError(f"缺少 Win7 第三方许可说明：{WIN7_THIRD_PARTY_NOTICE}")
    if not WINDOWS_WIN7_MANIFEST.is_file():
        raise RuntimeError(f"缺少 Win7 应用清单：{WINDOWS_WIN7_MANIFEST}")
    if not WIN7_PYINSTALLER_HOOK.is_file():
        raise RuntimeError(f"缺少 Win7 PyInstaller 钩子：{WIN7_PYINSTALLER_HOOK}")
    return resolved_7zip, resolved_ucrt, resolved_vc_runtime


def _require_files(directory: Path, names: tuple[str, ...], *, label: str) -> None:
    if not directory.is_dir():
        raise RuntimeError(f"{label} 目录不存在：{directory}")
    missing = [name for name in names if not (directory / name).is_file()]
    if missing:
        raise RuntimeError(f"{label} 目录缺少文件：{missing}")


def ensure_windows_x64_build_environment(target: str = WINDOWS_TARGET_MODERN) -> None:
    target = validate_windows_target(target)
    if not sys.platform.startswith("win"):
        raise RuntimeError("Windows 产物必须由 Windows runner 构建。")
    if target == WINDOWS_TARGET_WIN7 and sys.version_info[:2] != (3, 8):
        raise RuntimeError("Windows 7 兼容构建必须使用 Python 3.8。")
    if target == WINDOWS_TARGET_MODERN and sys.version_info < (3, 9):
        raise RuntimeError("构建 Python 必须为 3.9 或更高版本。")
    machine = platform.machine().lower()
    if struct.calcsize("P") != 8 or machine not in {"amd64", "x86_64"}:
        raise RuntimeError(f"必须使用 Windows x64 Python，当前架构：{platform.machine()}")


def ensure_build_dependencies(target: str = WINDOWS_TARGET_MODERN) -> None:
    target = validate_windows_target(target)
    modules = (
        WINDOWS_WIN7_BUILD_MODULES
        if target == WINDOWS_TARGET_WIN7
        else WINDOWS_BUILD_MODULES
    )
    missing = [module for module in modules if not _module_exists(module)]
    if missing:
        packages = sorted({modules[module] for module in missing})
        raise RuntimeError(
            "Windows 打包环境缺少依赖模块："
            + ", ".join(missing)
            + "。请安装："
            + " ".join(packages)
        )
    if target == WINDOWS_TARGET_WIN7:
        _ensure_win7_pinned_distributions()


def _ensure_win7_pinned_distributions() -> None:
    mismatches: list[str] = []
    for distribution, expected in WIN7_PINNED_DISTRIBUTIONS.items():
        try:
            actual = importlib_metadata.version(distribution)
        except importlib_metadata.PackageNotFoundError:
            mismatches.append(f"{distribution}=未安装（要求 {expected}）")
            continue
        if actual != expected:
            mismatches.append(f"{distribution}={actual}（要求 {expected}）")
    if mismatches:
        raise RuntimeError(
            "Windows 7 构建依赖必须与冻结清单完全一致：" + "；".join(mismatches)
        )


def build_windows_binaries(
    version: str,
    output_dir: Path,
    work_dir: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
    seven_zip_dir: Path | None = None,
    ucrt_dir: Path | None = None,
    vc_runtime_dir: Path | None = None,
) -> tuple[Path, Path]:
    target = validate_windows_target(target)
    output_dir.mkdir(parents=True, exist_ok=True)
    work_dir.mkdir(parents=True, exist_ok=True)
    app_dir = output_dir / APP_NAME
    updater_path = output_dir / f"{UPDATER_NAME}.exe"
    _remove_previous_output(app_dir, updater_path)

    version_file = work_dir / "HRToolkit.version.txt"
    updater_version_file = work_dir / "HRToolkitUpdater.version.txt"
    version_file.write_text(windows_version_info(version), encoding="utf-8")
    updater_version_file.write_text(
        windows_version_info(
            version,
            description="HRToolkit Updater",
            original_filename="HRToolkitUpdater.exe",
        ),
        encoding="utf-8",
    )

    main_command, updater_command = pyinstaller_commands(
        version=version,
        output_dir=output_dir,
        work_dir=work_dir,
        version_file=version_file,
        updater_version_file=updater_version_file,
        target=target,
        seven_zip_dir=seven_zip_dir,
        ucrt_dir=ucrt_dir,
        vc_runtime_dir=vc_runtime_dir,
    )
    _run(main_command)
    if not app_dir.is_dir():
        raise RuntimeError("PyInstaller 未生成预期的 HRToolkit onedir。")
    if target == WINDOWS_TARGET_WIN7:
        assert ucrt_dir is not None
        assert vc_runtime_dir is not None
        stage_win7_app_local_runtimes(
            app_dir=app_dir,
            ucrt_dir=ucrt_dir,
            vc_runtime_dir=vc_runtime_dir,
        )
    removed_bytes = remove_unused_opencv_videoio_ffmpeg(app_dir)
    if removed_bytes:
        print(f"已移除未使用的 OpenCV 视频后端：{removed_bytes} 字节")
    removed_translation_bytes = remove_unused_qt_translations(app_dir, target=target)
    if removed_translation_bytes:
        print(f"已移除未使用的 Qt 翻译：{removed_translation_bytes} 字节")
    removed_development_bytes = remove_qt_development_plugins(app_dir, target=target)
    if removed_development_bytes:
        print(f"已移除 Qt QML 开发插件：{removed_development_bytes} 字节")
    removed_image_plugin_bytes = remove_unused_qt_image_format_plugins(
        app_dir,
        target=target,
    )
    if removed_image_plugin_bytes:
        print(f"已移除未使用的 Qt 图片格式插件：{removed_image_plugin_bytes} 字节")
    _run(updater_command)
    if not app_dir.is_dir() or not updater_path.is_file():
        raise RuntimeError("PyInstaller 未生成预期的 HRToolkit onedir 和 Updater。")
    if target == WINDOWS_TARGET_WIN7:
        assert seven_zip_dir is not None
        assert ucrt_dir is not None
        assert vc_runtime_dir is not None
        verify_win7_runtime_source_integrity(
            app_dir=app_dir,
            updater=updater_path,
            seven_zip_dir=seven_zip_dir,
            ucrt_dir=ucrt_dir,
            vc_runtime_dir=vc_runtime_dir,
        )
    return app_dir, updater_path


def remove_unused_opencv_videoio_ffmpeg(app_dir: Path) -> int:
    """Remove only OpenCV's optional FFmpeg video I/O runtime from a Windows payload."""
    cv2_dir = app_dir / "_internal" / "cv2"
    matches = sorted(cv2_dir.glob(OPENCV_VIDEOIO_FFMPEG_PATTERN))
    removed_bytes = 0
    for path in matches:
        if path.is_symlink() or not path.is_file():
            raise RuntimeError(f"拒绝移除非普通 OpenCV 视频后端文件：{path}")
        removed_bytes += path.stat().st_size
        path.unlink()
    leftovers = sorted(cv2_dir.glob(OPENCV_VIDEOIO_FFMPEG_PATTERN))
    if leftovers:
        raise RuntimeError(f"OpenCV 视频后端未完全移除：{leftovers}")
    return removed_bytes


def remove_unused_qt_translations(
    app_dir: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> int:
    """Remove Qt translations outside the English and Chinese release locales."""

    target = validate_windows_target(target)
    internal = app_dir / "_internal"
    translations = (
        internal / QT5_WINDOWS_RUNTIME_ROOT / "translations"
        if target == WINDOWS_TARGET_WIN7
        else internal / QT6_WINDOWS_RUNTIME_ROOT / "translations"
    )
    removed_bytes = 0
    if not translations.is_dir():
        return removed_bytes
    for path in sorted(translations.glob("*.qm")):
        if path.name.endswith(QT_TRANSLATION_SUFFIXES):
            continue
        if path.is_symlink() or not path.is_file():
            raise RuntimeError(f"拒绝移除非普通 Qt 翻译文件：{path}")
        removed_bytes += path.stat().st_size
        path.unlink()
    return removed_bytes


def remove_qt_development_plugins(
    app_dir: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> int:
    """Remove QML debugger/profiler plugins from production Windows payloads."""

    target = validate_windows_target(target)
    internal = app_dir / "_internal"
    plugin_dir = (
        internal / QT5_WINDOWS_RUNTIME_ROOT / "plugins" / "qmltooling"
        if target == WINDOWS_TARGET_WIN7
        else internal / QT6_WINDOWS_RUNTIME_ROOT / "plugins" / "qmltooling"
    )
    if not plugin_dir.exists():
        return 0
    if plugin_dir.is_symlink() or not plugin_dir.is_dir():
        raise RuntimeError(f"拒绝移除异常 Qt 开发插件路径：{plugin_dir}")
    files = [path for path in plugin_dir.rglob("*") if path.is_file()]
    removed_bytes = sum(path.stat().st_size for path in files)
    shutil.rmtree(plugin_dir)
    return removed_bytes


def remove_unused_qt_image_format_plugins(
    app_dir: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> int:
    """Keep only codecs for image formats accepted by the business tools."""

    target = validate_windows_target(target)
    internal = app_dir / "_internal"
    plugin_dir = (
        internal / QT5_WINDOWS_RUNTIME_ROOT / "plugins" / "imageformats"
        if target == WINDOWS_TARGET_WIN7
        else internal / QT6_WINDOWS_RUNTIME_ROOT / "plugins" / "imageformats"
    )
    if not plugin_dir.is_dir():
        return 0
    removed_bytes = 0
    for path in sorted(plugin_dir.glob("*.dll")):
        if path.name.casefold() in QT_IMAGE_FORMAT_PLUGIN_NAMES:
            continue
        if path.is_symlink() or not path.is_file():
            raise RuntimeError(f"拒绝移除异常 Qt 图片格式插件：{path}")
        removed_bytes += path.stat().st_size
        path.unlink()
    return removed_bytes


def stage_win7_app_local_runtimes(
    *,
    app_dir: Path,
    ucrt_dir: Path,
    vc_runtime_dir: Path,
) -> None:
    """Place downlevel runtimes beside the main EXE as required before Windows 8."""
    for source_dir, names, label in (
        (ucrt_dir, WIN7_REQUIRED_UCRT_FILES, "app-local UCRT"),
        (vc_runtime_dir, WIN7_REQUIRED_VC_RUNTIME_FILES, "Visual C++ app-local runtime"),
    ):
        _require_files(source_dir, names, label=label)
        for name in names:
            matches = sorted(
                path
                for path in app_dir.rglob("*")
                if path.name.casefold() == name.casefold()
            )
            for path in matches:
                if path.is_symlink() or not path.is_file():
                    raise RuntimeError(f"拒绝替换非普通 Win7 运行库文件：{path}")
                path.unlink()
            shutil.copy2(source_dir / name, app_dir / name)


def pyinstaller_commands(
    *,
    version: str,
    output_dir: Path,
    work_dir: Path,
    version_file: Path | None = None,
    updater_version_file: Path | None = None,
    target: str = WINDOWS_TARGET_MODERN,
    seven_zip_dir: Path | None = None,
    ucrt_dir: Path | None = None,
    vc_runtime_dir: Path | None = None,
) -> tuple[list[str], list[str]]:
    validate_stable_semver(version)
    target = validate_windows_target(target)
    release_qml_files()
    if not QT_NOTICE.is_file() or not QT_PYINSTALLER_HOOKS_DIR.is_dir():
        raise RuntimeError("缺少 Qt 打包许可说明或 PyInstaller 钩子。")
    seven_zip_dir, ucrt_dir, vc_runtime_dir = validate_win7_runtime_sources(
        target=target,
        seven_zip_dir=seven_zip_dir,
        ucrt_dir=ucrt_dir,
        vc_runtime_dir=vc_runtime_dir,
    )
    version_file = version_file or (work_dir / "HRToolkit.version.txt")
    updater_version_file = updater_version_file or (work_dir / "HRToolkitUpdater.version.txt")
    spec_dir = work_dir / "spec"
    spec_dir.mkdir(parents=True, exist_ok=True)

    common = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--distpath",
        str(output_dir),
        "--specpath",
        str(spec_dir),
        "--icon",
        str(WINDOWS_ICON),
        "--version-file",
        str(version_file),
    ]
    common.extend(["--additional-hooks-dir", str(QT_PYINSTALLER_HOOKS_DIR)])
    if target == WINDOWS_TARGET_WIN7:
        common.extend(["--additional-hooks-dir", str(WIN7_PYINSTALLER_HOOKS_DIR)])
    main = [
        *common,
        "--name",
        APP_NAME,
        "--onedir",
        "--windowed",
        "--workpath",
        str(work_dir / APP_NAME),
        "--manifest",
        str(WINDOWS_WIN7_MANIFEST if target == WINDOWS_TARGET_WIN7 else WINDOWS_MANIFEST),
        "--add-data",
        f"{README_FILE};.",
        "--add-data",
        f"{QML_DIR};hr_toolkit/gui_qt/qml",
        "--add-data",
        f"{QT_NOTICE};third_party/qt",
        "--copy-metadata",
        "pypdfium2" if target == WINDOWS_TARGET_WIN7 else "pypdf",
    ]
    if target == WINDOWS_TARGET_WIN7:
        assert seven_zip_dir is not None
        main.extend(
            [
                "--add-binary",
                f"{seven_zip_dir / '7z.exe'};third_party/7zip",
                "--add-binary",
                f"{seven_zip_dir / '7z.dll'};third_party/7zip",
                "--add-data",
                f"{seven_zip_dir / 'License.txt'};third_party/7zip",
                "--add-data",
                f"{WIN7_THIRD_PARTY_NOTICE};third_party/7zip",
            ]
        )
    for template in release_template_files():
        main.extend(
            [
                "--add-data",
                f"{template};hr_toolkit/templates",
            ]
        )
    hidden_imports = WIN7_HIDDEN_IMPORTS if target == WINDOWS_TARGET_WIN7 else HIDDEN_IMPORTS
    collect_all_modules = (
        WIN7_COLLECT_ALL_MODULES
        if target == WINDOWS_TARGET_WIN7
        else COLLECT_ALL_MODULES
    )
    for module in hidden_imports:
        main.extend(["--hidden-import", module])
    for module in MAIN_APP_EXCLUDED_MODULES:
        main.extend(["--exclude-module", module])
    for module in collect_all_modules:
        main.extend(["--collect-all", module])
    main.append(str(APP_ENTRYPOINT))

    updater = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--distpath",
        str(output_dir),
        "--specpath",
        str(spec_dir),
        "--icon",
        str(WINDOWS_ICON),
        "--version-file",
        str(updater_version_file),
        "--name",
        UPDATER_NAME,
        "--onefile",
        "--windowed",
        "--workpath",
        str(work_dir / UPDATER_NAME),
    ]
    if target == WINDOWS_TARGET_WIN7:
        updater.extend(["--additional-hooks-dir", str(WIN7_PYINSTALLER_HOOKS_DIR)])
        assert ucrt_dir is not None
        for runtime_name in WIN7_REQUIRED_UCRT_FILES:
            updater.extend(["--add-binary", f"{ucrt_dir / runtime_name};."])
        assert vc_runtime_dir is not None
        for runtime_name in WIN7_REQUIRED_VC_RUNTIME_FILES:
            updater.extend(["--add-binary", f"{vc_runtime_dir / runtime_name};."])
    for module in EXCLUDED_MODULES:
        updater.extend(["--exclude-module", module])
    updater.append(str(UPDATER_ENTRYPOINT))
    return main, updater


def release_template_files() -> tuple[Path, ...]:
    expected = set(RELEASE_TEMPLATE_NAMES)
    discovered = {path.name for path in TEMPLATES_DIR.glob("*.xlsx") if path.is_file()}
    if discovered != expected:
        missing = sorted(expected - discovered)
        extra = sorted(discovered - expected)
        raise RuntimeError(f"内置模板白名单不一致，缺少={missing}，多出={extra}")
    return tuple(TEMPLATES_DIR / name for name in RELEASE_TEMPLATE_NAMES)


def release_qml_files() -> tuple[Path, ...]:
    files = tuple(sorted(path for path in QML_DIR.rglob("*.qml") if path.is_file()))
    required = {
        QML_DIR / "Main.qml",
        QML_DIR / "components" / "AppButton.qml",
        QML_DIR / "components" / "Card.qml",
    }
    missing = sorted(str(path) for path in required if path not in files)
    if missing:
        raise RuntimeError(f"Qt Quick 资源不完整，缺少={missing}")
    return files


def windows_version_info(
    version: str,
    *,
    description: str = "HRToolkit",
    original_filename: str = "HRToolkit.exe",
) -> str:
    major, minor, patch = validate_stable_semver(version)
    numeric = f"{major}, {minor}, {patch}, 0"
    return f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({numeric}),
    prodvers=({numeric}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '080404B0',
        [StringStruct('CompanyName', 'xhzwjc'),
         StringStruct('FileDescription', '{description}'),
         StringStruct('FileVersion', '{version}'),
         StringStruct('InternalName', 'HRToolkit'),
         StringStruct('OriginalFilename', '{original_filename}'),
         StringStruct('ProductName', 'HRToolkit'),
         StringStruct('ProductVersion', '{version}')])
    ]),
    VarFileInfo([VarStruct('Translation', [2052, 1200])])
  ]
)
"""


def verify_windows_payload(
    app_dir: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> None:
    target = validate_windows_target(target)
    app_dir = app_dir.resolve()
    launcher = app_dir / f"{APP_NAME}.exe"
    internal = app_dir / "_internal"
    if not launcher.is_file():
        raise RuntimeError(f"程序目录缺少 {launcher.name}：{app_dir}")
    if not internal.is_dir():
        raise RuntimeError(f"程序目录缺少 _internal：{app_dir}")

    video_backends = sorted(
        (internal / "cv2").glob(OPENCV_VIDEOIO_FFMPEG_PATTERN)
    )
    if video_backends:
        raise RuntimeError(f"程序包仍包含未使用的 OpenCV 视频后端：{video_backends}")

    files = sorted((path for path in app_dir.rglob("*") if path.is_file()), key=lambda path: path.as_posix())
    for path in files:
        relative = path.relative_to(app_dir)
        lowered_parts = {part.lower() for part in relative.parts}
        if lowered_parts & {part.lower() for part in FORBIDDEN_PAYLOAD_PARTS}:
            raise RuntimeError(f"程序包包含禁止目录或缓存：{relative}")
        if path.suffix.lower() in {".log", ".xlsm", ".xls"} | FORBIDDEN_DATA_SUFFIXES:
            raise RuntimeError(f"程序包包含非白名单或用户项目数据：{relative}")
        if (
            path.name.lower() in FORBIDDEN_DATA_FILENAMES
            or path.name.lower().startswith("history.db")
            or path.name.lower().startswith(".trash-move-")
        ):
            raise RuntimeError(f"程序包包含非白名单或用户项目数据：{relative}")
        if path.suffix.lower() == ".xlsx" and not _is_template_payload_path(relative):
            raise RuntimeError(f"程序包包含模板目录之外的 Excel：{relative}")
        if "pil" in lowered_parts and "avif" in path.name.casefold():
            raise RuntimeError(f"程序包包含业务不支持的 Pillow AVIF 解码运行库：{relative}")

    root_files = {path.name for path in app_dir.iterdir() if path.is_file()}
    allowed_root_files = {
        f"{APP_NAME}.exe",
        f"{UPDATER_NAME}.exe",
        "update_url.txt",
    }
    if target == WINDOWS_TARGET_WIN7:
        allowed_root_files.update(WIN7_REQUIRED_UCRT_FILES)
        allowed_root_files.update(WIN7_REQUIRED_VC_RUNTIME_FILES)
    unexpected_root_files = root_files - allowed_root_files
    if unexpected_root_files:
        raise RuntimeError(f"程序包根目录包含非白名单文件：{sorted(unexpected_root_files)}")

    expected_templates = {path.name for path in release_template_files()}
    packaged_templates = {
        path.name
        for path in files
        if path.suffix.lower() == ".xlsx" and _is_template_payload_path(path.relative_to(app_dir))
    }
    if packaged_templates != expected_templates:
        missing = sorted(expected_templates - packaged_templates)
        extra = sorted(packaged_templates - expected_templates)
        raise RuntimeError(f"内置模板集合不一致，缺少={missing}，多出={extra}")

    readmes = [path for path in files if path.name == README_FILE.name]
    if len(readmes) != 1 or readmes[0].read_bytes() != README_FILE.read_bytes():
        raise RuntimeError("程序包必须且只能包含一份与仓库一致的 README.md。")
    packaged_qml_root = internal / "hr_toolkit" / "gui_qt" / "qml"
    for source in release_qml_files():
        relative = source.relative_to(QML_DIR)
        packaged = packaged_qml_root / relative
        if not packaged.is_file() or packaged.read_bytes() != source.read_bytes():
            raise RuntimeError(f"程序包 Qt Quick 资源缺失或内容不一致：{relative}")
    qt_notice = internal / "third_party" / "qt" / QT_NOTICE.name
    if not qt_notice.is_file() or qt_notice.read_bytes() != QT_NOTICE.read_bytes():
        raise RuntimeError("程序包缺少正确的 Qt 第三方许可说明。")
    verify_packaged_qt_qml(internal, target=target)
    verify_payload_pe_architecture(app_dir)
    if target == WINDOWS_TARGET_WIN7:
        verify_win7_runtime_payload(app_dir)


def verify_packaged_qt_qml(internal: Path, *, target: str) -> None:
    target = validate_windows_target(target)
    if target == WINDOWS_TARGET_WIN7:
        qt_root = internal / QT5_WINDOWS_RUNTIME_ROOT
        required = QT5_REQUIRED_QML_FILES
    else:
        qt_root = internal / QT6_WINDOWS_RUNTIME_ROOT
        required = QT6_REQUIRED_QML_FILES
    qml_root = qt_root / "qml"
    translations = qt_root / "translations"
    missing = [relative for relative in required if not (qml_root / relative).is_file()]
    if missing:
        raise RuntimeError(
            f"程序包缺少 {target} Qt Quick 运行时资源：{missing}"
        )
    unexpected_translations = sorted(
        path.name
        for path in translations.glob("*.qm")
        if not path.name.endswith(QT_TRANSLATION_SUFFIXES)
    )
    if unexpected_translations:
        raise RuntimeError(
            f"程序包包含未使用的 {target} Qt 翻译：{unexpected_translations}"
        )
    plugin_root = qt_root / "plugins"
    if (plugin_root / "qmltooling").exists():
        raise RuntimeError(f"程序包包含仅供调试/分析使用的 {target} Qt QML 开发插件")
    unexpected_image_plugins = sorted(
        path.name
        for path in (plugin_root / "imageformats").glob("*.dll")
        if path.name.casefold() not in QT_IMAGE_FORMAT_PLUGIN_NAMES
    )
    if unexpected_image_plugins:
        raise RuntimeError(
            f"程序包包含业务不使用的 {target} Qt 图片格式插件："
            f"{unexpected_image_plugins}"
        )
    if target == WINDOWS_TARGET_WIN7:
        missing_runtime = [
            relative
            for relative in QT5_REQUIRED_RUNTIME_FILES
            if not (internal / relative).is_file()
        ]
        if missing_runtime:
            raise RuntimeError(
                f"程序包缺少 Windows 7 Qt/ANGLE 运行时：{missing_runtime}"
            )


def verify_win7_runtime_payload(app_dir: Path) -> None:
    internal = app_dir / "_internal"
    forbidden_runtime_names = {
        path.name.casefold()
        for path in app_dir.rglob("*")
        if path.is_file() and path.name.casefold() in WIN7_FORBIDDEN_DLLS
    }
    if forbidden_runtime_names:
        raise RuntimeError(
            "Windows 7 程序包不得通过旁加载伪造缺失的系统 API："
            f"{sorted(forbidden_runtime_names)}"
        )
    unexpected_api_sets = sorted(
        str(path.relative_to(app_dir))
        for path in app_dir.rglob("*")
        if path.is_file()
        and path.name.casefold().startswith(WIN7_API_SET_PREFIXES)
        and path.name.casefold() not in WIN7_REQUIRED_UCRT_FILE_KEYS
    )
    if unexpected_api_sets:
        raise RuntimeError(
            "Windows 7 程序包混入未锁定的系统 API Set："
            f"{unexpected_api_sets}"
        )
    _require_files(
        internal,
        WIN7_REQUIRED_PYTHON_RUNTIME_FILES,
        label="Win7 payload Python 3.8 runtime",
    )
    allowed_python_runtime = {
        name.casefold() for name in WIN7_REQUIRED_PYTHON_RUNTIME_FILES
    }
    unexpected_python = sorted(
        path.name
        for path in internal.glob("python3*.dll")
        if path.name.casefold() not in allowed_python_runtime
    )
    if unexpected_python:
        raise RuntimeError(f"Windows 7 程序包混入其他 Python 运行时：{unexpected_python}")
    _require_files(app_dir, WIN7_REQUIRED_UCRT_FILES, label="Win7 payload UCRT")
    _require_files(
        app_dir,
        WIN7_REQUIRED_VC_RUNTIME_FILES,
        label="Win7 payload Visual C++ runtime",
    )
    seven_zip_dir = internal / "third_party" / "7zip"
    _require_files(seven_zip_dir, WIN7_REQUIRED_7ZIP_FILES, label="Win7 payload 7-Zip")
    notice = seven_zip_dir / WIN7_THIRD_PARTY_NOTICE.name
    if not notice.is_file() or notice.read_bytes() != WIN7_THIRD_PARTY_NOTICE.read_bytes():
        raise RuntimeError("Windows 7 程序包缺少正确的 7-Zip 第三方许可说明。")


def verify_win7_runtime_source_integrity(
    *,
    app_dir: Path,
    updater: Path,
    seven_zip_dir: Path,
    ucrt_dir: Path,
    vc_runtime_dir: Path,
    archive_reader_cls=None,
) -> None:
    """Ensure PyInstaller did not replace the pinned Win7 compatibility runtimes."""
    internal = app_dir / "_internal"
    _require_matching_files(
        source_dir=ucrt_dir,
        payload_dir=app_dir,
        names=WIN7_REQUIRED_UCRT_FILES,
        label="app-local UCRT",
    )
    _require_matching_files(
        source_dir=vc_runtime_dir,
        payload_dir=app_dir,
        names=WIN7_REQUIRED_VC_RUNTIME_FILES,
        label="Visual C++ app-local runtime",
    )
    _require_matching_files(
        source_dir=seven_zip_dir,
        payload_dir=internal / "third_party" / "7zip",
        names=WIN7_REQUIRED_7ZIP_FILES,
        label="7-Zip",
    )

    updater_runtime_sources = tuple(
        (name, ucrt_dir / name) for name in WIN7_REQUIRED_UCRT_FILES
    ) + tuple(
        (name, vc_runtime_dir / name) for name in WIN7_REQUIRED_VC_RUNTIME_FILES
    )
    _verify_onefile_embedded_files(
        updater,
        updater_runtime_sources,
        archive_reader_cls=archive_reader_cls,
    )


def _require_matching_files(
    *,
    source_dir: Path,
    payload_dir: Path,
    names: tuple[str, ...],
    label: str,
) -> None:
    for name in names:
        source = source_dir / name
        payload = payload_dir / name
        if not source.is_file() or not payload.is_file():
            raise RuntimeError(f"{label} 完整性检查缺少文件：{name}")
        if payload.read_bytes() != source.read_bytes():
            raise RuntimeError(f"{label} 未使用已锁定的源文件：{name}")


def _verify_onefile_embedded_files(
    executable: Path,
    expected_sources: tuple[tuple[str, Path], ...],
    *,
    archive_reader_cls=None,
) -> None:
    if archive_reader_cls is None:
        try:
            from PyInstaller.archive.readers import CArchiveReader
        except ImportError as exc:
            raise RuntimeError("无法读取 Win7 Updater 内嵌运行库：缺少 PyInstaller。") from exc
        archive_reader_cls = CArchiveReader

    try:
        archive = archive_reader_cls(str(executable))
    except Exception as exc:
        raise RuntimeError(f"无法读取 Win7 Updater 内嵌归档：{executable}") from exc

    toc_names: dict[str, list[str]] = {}
    for raw_name in archive.toc:
        name = str(raw_name).replace("\\", "/")
        while name.startswith("./"):
            name = name[2:]
        toc_names.setdefault(name.casefold(), []).append(str(raw_name))

    for expected_name, source in expected_sources:
        matches = toc_names.get(expected_name.casefold(), [])
        if len(matches) != 1:
            raise RuntimeError(
                "Win7 Updater 内嵌运行库缺失或重复："
                f"{expected_name}（匹配数 {len(matches)}）"
            )
        try:
            embedded = archive.extract(matches[0])
        except Exception as exc:
            raise RuntimeError(
                f"无法提取 Win7 Updater 内嵌运行库：{expected_name}"
            ) from exc
        if not source.is_file() or embedded != source.read_bytes():
            raise RuntimeError(
                f"Win7 Updater 未使用已锁定的运行库源文件：{expected_name}"
            )

    unexpected_api_sets = sorted(
        raw_name
        for normalized_name, raw_names in toc_names.items()
        if normalized_name.rsplit("/", 1)[-1].startswith(WIN7_API_SET_PREFIXES)
        and normalized_name.rsplit("/", 1)[-1] not in WIN7_REQUIRED_UCRT_FILE_KEYS
        for raw_name in raw_names
    )
    if unexpected_api_sets:
        raise RuntimeError(
            "Win7 Updater 内嵌了未锁定的系统 API Set："
            f"{unexpected_api_sets}"
        )


def verify_pe_x64(executable: Path) -> None:
    machine = read_pe_machine(executable)
    if machine != PE_MACHINE_AMD64:
        raise RuntimeError(
            f"{executable.name} 不是 x64 PE（machine=0x{machine:04x}，期望 0x{PE_MACHINE_AMD64:04x}）。"
        )


def verify_payload_pe_architecture(app_dir: Path) -> None:
    pe_files = _payload_pe_files(app_dir)
    if not pe_files:
        raise RuntimeError(f"程序目录未发现 Windows PE 文件：{app_dir}")
    for path in pe_files:
        verify_pe_x64(path)


def _payload_pe_files(app_dir: Path) -> tuple[Path, ...]:
    return tuple(
        sorted(
            (
                path
                for path in app_dir.rglob("*")
                if path.is_file() and path.suffix.lower() in {".exe", ".dll", ".pyd"}
            ),
            key=lambda path: path.as_posix(),
        )
    )


def verify_win7_pe_compatibility(pe_files: tuple[Path, ...]) -> None:
    try:
        import pefile
    except ImportError as exc:
        raise RuntimeError("Win7 PE 兼容检查缺少 pefile。") from exc

    pe_files = tuple(pe_files)
    payload_by_name: dict[str, Path] = {}
    for path in pe_files:
        payload_by_name.setdefault(path.name.casefold(), path)
    runtime_exports = {
        name.casefold(): _read_pe_exports(pefile, payload_by_name[name.casefold()])
        for name in WIN7_REQUIRED_VC_RUNTIME_FILES
        if name.casefold() in payload_by_name
    }

    violations: list[str] = []
    for path in pe_files:
        try:
            pe = pefile.PE(str(path), fast_load=True)
            pe.parse_data_directories(
                directories=[
                    pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_IMPORT"],
                    pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT"],
                ]
            )
        except pefile.PEFormatError as exc:
            raise RuntimeError(f"无法检查 Win7 PE 导入表：{path}：{exc}") from exc
        try:
            subsystem = (
                int(pe.OPTIONAL_HEADER.MajorSubsystemVersion),
                int(pe.OPTIONAL_HEADER.MinorSubsystemVersion),
            )
            # Microsoft's down-level app-local UCRT is explicitly supported on
            # Windows 7, although its own PE headers report subsystem 10.0.
            if (
                subsystem > (6, 1)
                and path.name.casefold() not in WIN7_REQUIRED_UCRT_FILE_KEYS
            ):
                violations.append(f"{path.name}: subsystem={subsystem[0]}.{subsystem[1]}")
            entries = tuple(getattr(pe, "DIRECTORY_ENTRY_IMPORT", ())) + tuple(
                getattr(pe, "DIRECTORY_ENTRY_DELAY_IMPORT", ())
            )
            for entry in entries:
                dll_name = _decode_pe_name(getattr(entry, "dll", b""))
                normalized_dll = dll_name.casefold()
                if normalized_dll in WIN7_FORBIDDEN_DLLS:
                    violations.append(f"{path.name}: {dll_name}")
                if (
                    normalized_dll.startswith(WIN7_API_SET_PREFIXES)
                    and normalized_dll not in payload_by_name
                ):
                    violations.append(f"{path.name}: 未随包提供 {dll_name}")
                if (
                    _is_win7_app_local_vc_runtime(normalized_dll)
                    and normalized_dll not in payload_by_name
                ):
                    violations.append(f"{path.name}: 缺少 app-local {dll_name}")
                exported_names, exported_ordinals = runtime_exports.get(
                    normalized_dll,
                    (frozenset(), frozenset()),
                )
                for imported in getattr(entry, "imports", ()):
                    raw_symbol = getattr(imported, "name", None)
                    symbol = _decode_pe_name(raw_symbol)
                    normalized = symbol.casefold()
                    if (
                        normalized in WIN7_FORBIDDEN_IMPORTS
                        or normalized.startswith("pss")
                        or normalized.startswith("pathcch")
                    ):
                        violations.append(f"{path.name}: {dll_name}!{symbol}")
                    if normalized_dll in runtime_exports:
                        ordinal = int(getattr(imported, "ordinal", 0) or 0)
                        if raw_symbol is not None and raw_symbol not in exported_names:
                            violations.append(
                                f"{path.name}: {dll_name} 不导出 {symbol}"
                            )
                        elif raw_symbol is None and ordinal not in exported_ordinals:
                            violations.append(
                                f"{path.name}: {dll_name} 不导出序号 {ordinal}"
                            )
        finally:
            pe.close()

    if violations:
        preview = "；".join(violations[:20])
        suffix = "" if len(violations) <= 20 else f"；另有 {len(violations) - 20} 项"
        raise RuntimeError(f"程序包包含 Windows 7 不支持的 PE 依赖：{preview}{suffix}")


def _is_win7_app_local_vc_runtime(dll_name: str) -> bool:
    return dll_name.startswith(
        ("concrt140", "mfc140", "msvcp140", "vcomp140", "vcruntime140")
    ) and dll_name.endswith(".dll")


def _read_pe_exports(pefile_module, path: Path) -> tuple[frozenset[bytes], frozenset[int]]:
    try:
        pe = pefile_module.PE(str(path), fast_load=True)
        pe.parse_data_directories(
            directories=[
                pefile_module.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_EXPORT"]
            ]
        )
    except pefile_module.PEFormatError as exc:
        raise RuntimeError(f"无法检查 VC 运行库导出表：{path}：{exc}") from exc
    try:
        export_entry = getattr(pe, "DIRECTORY_ENTRY_EXPORT", None)
        if export_entry is None:
            raise RuntimeError(f"VC 运行库没有导出表：{path}")
        symbols = tuple(getattr(export_entry, "symbols", ()))
        names = frozenset(
            symbol.name for symbol in symbols if getattr(symbol, "name", None) is not None
        )
        ordinals = frozenset(
            int(symbol.ordinal) for symbol in symbols if getattr(symbol, "ordinal", None) is not None
        )
        return names, ordinals
    finally:
        pe.close()


def _decode_pe_name(value: bytes | str | None) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("ascii", errors="replace")
    return str(value)


def read_pe_machine(executable: Path) -> int:
    with executable.open("rb") as handle:
        if handle.read(2) != b"MZ":
            raise RuntimeError(f"不是有效的 Windows PE：{executable}")
        handle.seek(0x3C)
        offset_data = handle.read(4)
        if len(offset_data) != 4:
            raise RuntimeError(f"PE 头不完整：{executable}")
        pe_offset = struct.unpack("<I", offset_data)[0]
        handle.seek(pe_offset)
        if handle.read(4) != b"PE\0\0":
            raise RuntimeError(f"PE 签名无效：{executable}")
        machine_data = handle.read(2)
        if len(machine_data) != 2:
            raise RuntimeError(f"PE machine 字段缺失：{executable}")
        return struct.unpack("<H", machine_data)[0]


def run_runtime_smoke(
    app_executable: Path,
    updater_executable: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> None:
    target = validate_windows_target(target)
    verify_pe_x64(updater_executable)
    expected_version = read_project_version()
    with tempfile.TemporaryDirectory(prefix="hr_toolkit_runtime_check_") as tmp:
        output_path = Path(tmp) / "result.txt"
        env = dict(os.environ)
        if target == WINDOWS_TARGET_WIN7:
            env.pop(WIN7_7ZIP_OVERRIDE_ENV, None)
        env["HR_TOOLKIT_CHECK_OUTPUT"] = str(output_path)
        _run_packaged_check(
            [str(app_executable), "--version"],
            output_path=output_path,
            label="打包程序版本检查",
            timeout=60,
            env=env,
        )
        actual_version = output_path.read_text(encoding="utf-8").strip()
        if actual_version != expected_version:
            raise RuntimeError(
                f"打包程序版本不一致：期望 {expected_version}，实际 {actual_version or '空'}"
            )
        _run_packaged_check(
            [str(app_executable), "--smoke-test"],
            output_path=output_path,
            label="打包程序 smoke-test",
            timeout=180,
            env=env,
        )
        smoke_result = output_path.read_text(encoding="utf-8").strip()
        if f"HRToolkit {expected_version} smoke-test OK" not in smoke_result:
            raise RuntimeError(f"打包程序 smoke-test 输出不正确：{smoke_result or '空'}")
        _run_packaged_check(
            [str(app_executable), "--update-smoke-test"],
            output_path=output_path,
            label="打包程序 update-smoke-test",
            timeout=90,
            env=env,
        )
        update_smoke_result = output_path.read_text(encoding="utf-8").strip()
        expected_prefix = f"HRToolkit {expected_version} update-smoke-test OK; latest="
        if not update_smoke_result.startswith(expected_prefix):
            raise RuntimeError(
                f"打包程序 update-smoke-test 输出不正确：{update_smoke_result or '空'}"
            )
        qt_env = dict(env)
        qt_env["HR_TOOLKIT_SKIP_UPDATE"] = "1"
        if target == WINDOWS_TARGET_WIN7:
            qt_env.update(WIN7_QT_SMOKE_ENV)
        _run_packaged_check(
            [str(app_executable), "--qt-smoke-test"],
            output_path=output_path,
            label="打包程序 Qt Quick smoke-test",
            timeout=90,
            env=qt_env,
        )
        qt_smoke_result = output_path.read_text(encoding="utf-8").strip()
        if qt_smoke_result != "HRToolkit Qt smoke-test OK":
            raise RuntimeError(
                f"打包程序 Qt Quick smoke-test 输出不正确：{qt_smoke_result or '空'}"
            )
        if target == WINDOWS_TARGET_WIN7:
            updater_smoke_dir = Path(tmp) / "updater"
            updater_smoke_dir.mkdir()
            updater_smoke = updater_smoke_dir / updater_executable.name
            shutil.copy2(updater_executable, updater_smoke)
            for runtime_name in (
                *WIN7_REQUIRED_UCRT_FILES,
                *WIN7_REQUIRED_VC_RUNTIME_FILES,
            ):
                source = app_executable.parent / runtime_name
                if source.is_symlink() or not source.is_file():
                    raise RuntimeError(
                        f"Win7 Updater 启动检查缺少 app-local 运行库：{runtime_name}"
                    )
                shutil.copy2(source, updater_smoke_dir / runtime_name)
            _run_packaged_check(
                [str(updater_smoke), "--smoke-test"],
                output_path=output_path,
                label="打包更新程序 smoke-test",
                timeout=90,
                env=env,
            )
            updater_smoke_result = output_path.read_text(encoding="utf-8").strip()
            expected_updater = f"HRToolkitUpdater {expected_version} smoke-test OK"
            if updater_smoke_result != expected_updater:
                raise RuntimeError(
                    "打包更新程序 smoke-test 输出不正确："
                    f"{updater_smoke_result or '空'}"
                )


def _run_packaged_check(
    command: list[str],
    *,
    output_path: Path,
    label: str,
    timeout: int,
    env: dict[str, str],
) -> None:
    output_path.unlink(missing_ok=True)
    try:
        _run(command, timeout=timeout, env=env)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
        try:
            detail = output_path.read_text(encoding="utf-8").strip()
        except OSError:
            detail = ""
        suffix = f"；程序记录：{detail}" if detail else ""
        raise RuntimeError(f"{label}失败：{exc}{suffix}") from exc


def _is_template_payload_path(relative: Path) -> bool:
    parts = tuple(part.lower() for part in relative.parts)
    return len(parts) >= 3 and parts[-3:-1] == ("hr_toolkit", "templates")


def _remove_previous_output(app_dir: Path, updater_path: Path) -> None:
    if app_dir.exists():
        if not app_dir.is_dir() or app_dir.name != APP_NAME:
            raise RuntimeError(f"拒绝清理非预期构建目录：{app_dir}")
        shutil.rmtree(app_dir)
    if updater_path.exists():
        if not updater_path.is_file() or updater_path.name != f"{UPDATER_NAME}.exe":
            raise RuntimeError(f"拒绝清理非预期构建文件：{updater_path}")
        updater_path.unlink()


def _run(
    command: list[str],
    *,
    timeout: int | None = None,
    env: dict[str, str] | None = None,
) -> None:
    print("执行：" + subprocess.list2cmdline(command))
    subprocess.run(command, cwd=REPO_ROOT, check=True, timeout=timeout, env=env)


def _module_exists(module: str) -> bool:
    try:
        return importlib.util.find_spec(module) is not None
    except ModuleNotFoundError:
        return False


if __name__ == "__main__":
    raise SystemExit(main())
