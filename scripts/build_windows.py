from __future__ import annotations

import argparse
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
WINDOWS_ICON = REPO_ROOT / "packaging" / "windows" / "HRToolkit.ico"
README_FILE = REPO_ROOT / "README.md"
TEMPLATES_DIR = REPO_ROOT / "hr_toolkit" / "templates"

WINDOWS_BUILD_MODULES = {
    "PyInstaller": "pyinstaller",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd",
    "pythoncom": "pywin32",
    "pywintypes": "pywin32",
    "win32com.client": "pywin32",
    "win32timezone": "pywin32",
}
HIDDEN_IMPORTS = (
    "pythoncom",
    "pywintypes",
    "win32com.client",
    "win32timezone",
    "xlrd",
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
FORBIDDEN_PAYLOAD_PARTS = {
    "__pycache__",
    ".pytest_cache",
    "tests",
    "test",
    "outputs",
    "output",
    "附件",
    "二期新增的附件",
}


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
    args = parser.parse_args(argv)

    version = validate_build_version(args.version)
    ensure_windows_x64_build_environment()
    ensure_build_dependencies()

    app_dir, updater = build_windows_binaries(
        version=version,
        output_dir=args.output_dir.resolve(),
        work_dir=args.work_dir.resolve(),
    )
    verify_windows_payload(app_dir)
    verify_pe_x64(app_dir / f"{APP_NAME}.exe")
    verify_pe_x64(updater)
    if not args.skip_runtime_smoke:
        run_runtime_smoke(app_dir / f"{APP_NAME}.exe", updater)

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


def ensure_windows_x64_build_environment() -> None:
    if not sys.platform.startswith("win"):
        raise RuntimeError("Windows 产物必须由 Windows runner 构建。")
    if sys.version_info < (3, 9):
        raise RuntimeError("构建 Python 必须为 3.9 或更高版本。")
    machine = platform.machine().lower()
    if struct.calcsize("P") != 8 or machine not in {"amd64", "x86_64"}:
        raise RuntimeError(f"必须使用 Windows x64 Python，当前架构：{platform.machine()}")


def ensure_build_dependencies() -> None:
    missing = [module for module in WINDOWS_BUILD_MODULES if not _module_exists(module)]
    if not missing:
        return
    packages = sorted({WINDOWS_BUILD_MODULES[module] for module in missing})
    raise RuntimeError(
        "Windows 打包环境缺少依赖模块："
        + ", ".join(missing)
        + "。请安装："
        + " ".join(packages)
    )


def build_windows_binaries(version: str, output_dir: Path, work_dir: Path) -> tuple[Path, Path]:
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
    )
    _run(main_command)
    _run(updater_command)
    if not app_dir.is_dir() or not updater_path.is_file():
        raise RuntimeError("PyInstaller 未生成预期的 HRToolkit onedir 和 Updater。")
    return app_dir, updater_path


def pyinstaller_commands(
    *,
    version: str,
    output_dir: Path,
    work_dir: Path,
    version_file: Path | None = None,
    updater_version_file: Path | None = None,
) -> tuple[list[str], list[str]]:
    validate_stable_semver(version)
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
    main = [
        *common,
        "--name",
        APP_NAME,
        "--onedir",
        "--windowed",
        "--workpath",
        str(work_dir / APP_NAME),
        "--manifest",
        str(WINDOWS_MANIFEST),
        "--add-data",
        f"{README_FILE};.",
    ]
    for template in release_template_files():
        main.extend(
            [
                "--add-data",
                f"{template};hr_toolkit/templates",
            ]
        )
    for module in HIDDEN_IMPORTS:
        main.extend(["--hidden-import", module])
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
        str(UPDATER_ENTRYPOINT),
    ]
    return main, updater


def release_template_files() -> tuple[Path, ...]:
    expected = set(RELEASE_TEMPLATE_NAMES)
    discovered = {path.name for path in TEMPLATES_DIR.glob("*.xlsx") if path.is_file()}
    if discovered != expected:
        missing = sorted(expected - discovered)
        extra = sorted(discovered - expected)
        raise RuntimeError(f"内置模板白名单不一致，缺少={missing}，多出={extra}")
    return tuple(TEMPLATES_DIR / name for name in RELEASE_TEMPLATE_NAMES)


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


def verify_windows_payload(app_dir: Path) -> None:
    app_dir = app_dir.resolve()
    launcher = app_dir / f"{APP_NAME}.exe"
    internal = app_dir / "_internal"
    if not launcher.is_file():
        raise RuntimeError(f"程序目录缺少 {launcher.name}：{app_dir}")
    if not internal.is_dir():
        raise RuntimeError(f"程序目录缺少 _internal：{app_dir}")

    files = sorted((path for path in app_dir.rglob("*") if path.is_file()), key=lambda path: path.as_posix())
    for path in files:
        relative = path.relative_to(app_dir)
        lowered_parts = {part.lower() for part in relative.parts}
        if lowered_parts & {part.lower() for part in FORBIDDEN_PAYLOAD_PARTS}:
            raise RuntimeError(f"程序包包含禁止目录或缓存：{relative}")
        if path.suffix.lower() in {".log", ".xlsm", ".xls"}:
            raise RuntimeError(f"程序包包含非白名单数据文件：{relative}")
        if path.suffix.lower() == ".xlsx" and not _is_template_payload_path(relative):
            raise RuntimeError(f"程序包包含模板目录之外的 Excel：{relative}")

    root_files = {path.name for path in app_dir.iterdir() if path.is_file()}
    allowed_root_files = {
        f"{APP_NAME}.exe",
        f"{UPDATER_NAME}.exe",
        "update_url.txt",
    }
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
    verify_payload_pe_architecture(app_dir)


def verify_pe_x64(executable: Path) -> None:
    machine = read_pe_machine(executable)
    if machine != PE_MACHINE_AMD64:
        raise RuntimeError(
            f"{executable.name} 不是 x64 PE（machine=0x{machine:04x}，期望 0x{PE_MACHINE_AMD64:04x}）。"
        )


def verify_payload_pe_architecture(app_dir: Path) -> None:
    pe_files = sorted(
        (
            path
            for path in app_dir.rglob("*")
            if path.is_file() and path.suffix.lower() in {".exe", ".dll", ".pyd"}
        ),
        key=lambda path: path.as_posix(),
    )
    if not pe_files:
        raise RuntimeError(f"程序目录未发现 Windows PE 文件：{app_dir}")
    for path in pe_files:
        verify_pe_x64(path)


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


def run_runtime_smoke(app_executable: Path, updater_executable: Path) -> None:
    # Updater 使用 PyInstaller --windowed；argparse --help 在无控制台环境没有
    # 稳定 stdout，因此这里只做 PE 架构检查。其替换/回滚逻辑由 unittest 覆盖。
    verify_pe_x64(updater_executable)
    expected_version = read_project_version()
    with tempfile.TemporaryDirectory(prefix="hr_toolkit_runtime_check_") as tmp:
        output_path = Path(tmp) / "result.txt"
        env = dict(os.environ)
        env["HR_TOOLKIT_CHECK_OUTPUT"] = str(output_path)
        _run([str(app_executable), "--version"], timeout=60, env=env)
        actual_version = output_path.read_text(encoding="utf-8").strip()
        if actual_version != expected_version:
            raise RuntimeError(
                f"打包程序版本不一致：期望 {expected_version}，实际 {actual_version or '空'}"
            )
        _run([str(app_executable), "--smoke-test"], timeout=60, env=env)
        smoke_result = output_path.read_text(encoding="utf-8").strip()
        if f"HRToolkit {expected_version} smoke-test OK" not in smoke_result:
            raise RuntimeError(f"打包程序 smoke-test 输出不正确：{smoke_result or '空'}")


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
