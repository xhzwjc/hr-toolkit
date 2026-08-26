from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from build_update_assets import stage_windows_payload, verify_staged_payload
from build_windows import (
    APP_NAME,
    REPO_ROOT,
    UPDATER_NAME,
    WINDOWS_ICON,
    WINDOWS_TARGET_MODERN,
    WINDOWS_TARGET_WIN7,
    WINDOWS_TARGETS,
    run_runtime_smoke,
    validate_build_version,
    validate_windows_target,
    windows_asset_suffix,
    windows_msi_asset_name,
    windows_setup_asset_name,
)
from generate_release_metadata import require_release_assets_under_limit


INNO_SCRIPT = REPO_ROOT / "packaging" / "windows" / "HRToolkit.iss"
# Pinned from jrsoftware/issrc tag is-6_7_1 with trailing whitespace normalized.
INNO_LANGUAGE_FILE = REPO_ROOT / "packaging" / "windows" / "ChineseSimplified.isl"
INNO_LANGUAGE_SHA256 = "75ec648a9c1b547b1c35113b06bc85cede51c1c1d7d089af8fd974331f930570"
WIX_SOURCE = REPO_ROOT / "packaging" / "windows" / "HRToolkit.wxs"
INNO_ENV = "INNO_SETUP_COMPILER"
WIX_ENV = "WIX_EXECUTABLE"
WIX_NAMESPACE = "http://wixtoolset.org/schemas/v4/wxs"
MSI_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
COMPONENT_NAMESPACE = uuid.UUID("e51355f4-644d-4d85-8525-ef9c08b088f2")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="从已构建的 Windows payload 纯生成普通用户 EXE 与 per-user MSI 安装器。"
    )
    parser.add_argument("--version", required=True, help="必须与 hr_toolkit.__version__ 一致")
    parser.add_argument("--app-dir", required=True, type=Path, help="PyInstaller HRToolkit onedir")
    parser.add_argument("--updater", required=True, type=Path, help="HRToolkitUpdater.exe")
    parser.add_argument("--output-dir", required=True, type=Path, help="安装器输出目录")
    parser.add_argument("--inno-compiler", help=f"ISCC.exe；也可设置 {INNO_ENV}")
    parser.add_argument("--wix-executable", help=f"WiX v4+ wix.exe；也可设置 {WIX_ENV}")
    parser.add_argument(
        "--target",
        choices=WINDOWS_TARGETS,
        default=WINDOWS_TARGET_MODERN,
        help="modern 保持现有安装器；win7 生成 Windows 7 SP1 x64 安装器",
    )
    parser.add_argument(
        "--skip-install-smoke",
        action="store_true",
        help="仅供诊断；跳过 EXE/MSI 静默安装、运行和卸载验证",
    )
    parser.add_argument(
        "--inno-sign-tool-name",
        help="未来签名入口：传入已通过 ISCC /S 注册的 SignTool 名称",
    )
    args = parser.parse_args(argv)

    version = validate_build_version(args.version)
    ensure_windows_runtime()
    exe_path, msi_path = build_windows_installers(
        version=version,
        app_dir=args.app_dir.resolve(),
        updater=args.updater.resolve(),
        output_dir=args.output_dir.resolve(),
        inno_compiler=resolve_inno_compiler(args.inno_compiler),
        wix_executable=resolve_wix_executable(args.wix_executable),
        install_smoke=not args.skip_install_smoke,
        inno_sign_tool_name=args.inno_sign_tool_name,
        target=args.target,
    )
    print(f"Windows EXE 安装器：{exe_path}")
    print(f"Windows MSI 安装器：{msi_path}")
    return 0


def ensure_windows_runtime() -> None:
    if not sys.platform.startswith("win"):
        raise RuntimeError("Windows 安装器必须由 Windows runner 构建和验证。")


def installer_asset_names(
    version: str,
    target: str = WINDOWS_TARGET_MODERN,
) -> tuple[str, str]:
    validate_build_version(version)
    return (
        windows_setup_asset_name(version, target),
        windows_msi_asset_name(version, target),
    )


def _inno_min_windows_version(target: str) -> str:
    target = validate_windows_target(target)
    return "6.1sp1" if target == WINDOWS_TARGET_WIN7 else "6.3"


def build_windows_installers(
    *,
    version: str,
    app_dir: Path,
    updater: Path,
    output_dir: Path,
    inno_compiler: str,
    wix_executable: str,
    install_smoke: bool = True,
    inno_sign_tool_name: str | None = None,
    target: str = WINDOWS_TARGET_MODERN,
) -> tuple[Path, Path]:
    validate_build_version(version)
    target = validate_windows_target(target)
    validate_installer_definitions()
    # Python 3.8 on Windows may leave a non-existent path relative after
    # Path.resolve(). Inno interprets a relative OutputDir from the .iss file's
    # directory, so normalize it without requiring the directory to exist.
    output_dir = Path(os.path.abspath(output_dir))
    output_dir.mkdir(parents=True, exist_ok=True)
    exe_name, msi_name = installer_asset_names(version, target)
    exe_path = output_dir / exe_name
    msi_path = output_dir / msi_name
    exe_path.unlink(missing_ok=True)
    msi_path.unlink(missing_ok=True)

    with tempfile.TemporaryDirectory(prefix="hr_toolkit_windows_installers_") as tmp:
        temp_dir = Path(tmp)
        payload_dir = temp_dir / "payload"
        stage_windows_payload(
            app_dir=app_dir,
            updater=updater,
            target_dir=payload_dir,
            target=target,
        )
        wix_fragment = temp_dir / "HRToolkitPayload.wxs"
        generate_wix_payload_fragment(payload_dir, wix_fragment, target=target)

        _run(
            inno_compile_command(
                compiler=inno_compiler,
                version=version,
                payload_dir=payload_dir,
                output_dir=output_dir,
                sign_tool_name=inno_sign_tool_name,
                target=target,
            )
        )
        _run(
            wix_build_command(
                wix_executable=wix_executable,
                version=version,
                payload_fragment=wix_fragment,
                output_path=msi_path,
                target=target,
            )
        )

    verify_installer_outputs(exe_path, msi_path)
    require_release_assets_under_limit((exe_path, msi_path))
    if install_smoke:
        smoke_test_installers(exe_path, msi_path, target=target)
    return exe_path, msi_path


def inno_compile_command(
    *,
    compiler: str,
    version: str,
    payload_dir: Path,
    output_dir: Path,
    sign_tool_name: str | None = None,
    target: str = WINDOWS_TARGET_MODERN,
) -> list[str]:
    validate_build_version(version)
    target = validate_windows_target(target)
    absolute_output_dir = Path(os.path.abspath(output_dir))
    command = [
        compiler,
        "/Qp",
        f"/DMyAppVersion={version}",
        f"/DSourceDir={payload_dir}",
        f"/DOutputDir={absolute_output_dir}",
        f"/DSetupIconFile={WINDOWS_ICON}",
        f"/DInstallerSuffix={windows_asset_suffix(target)}",
        f"/DMinWindowsVersion={_inno_min_windows_version(target)}",
    ]
    if target == WINDOWS_TARGET_WIN7:
        command.append("/DWin7Compatibility=1")
        command.append("/DCleanExistingPayload=1")
    if sign_tool_name:
        command.append(f"/DSignToolName={sign_tool_name}")
    command.append(str(INNO_SCRIPT))
    return command


def wix_build_command(
    *,
    wix_executable: str,
    version: str,
    payload_fragment: Path,
    output_path: Path,
    target: str = WINDOWS_TARGET_MODERN,
) -> list[str]:
    validate_build_version(version)
    target = validate_windows_target(target)
    return [
        wix_executable,
        "build",
        str(WIX_SOURCE),
        str(payload_fragment),
        "-arch",
        "x64",
        "-d",
        f"AppVersion={version}",
        "-d",
        f"WindowsTarget={target}",
        "-pdbtype",
        "none",
        "-o",
        str(output_path),
    ]


def generate_wix_payload_fragment(
    payload_dir: Path,
    output_path: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> Path:
    payload_dir = payload_dir.resolve()
    verify_staged_payload(payload_dir, target=target)
    ET.register_namespace("", WIX_NAMESPACE)
    wix = ET.Element(_wix("Wix"))
    directory_fragment = ET.SubElement(wix, _wix("Fragment"))
    app_ref = ET.SubElement(directory_fragment, _wix("DirectoryRef"), {"Id": "APPDIR"})
    group_fragment = ET.SubElement(wix, _wix("Fragment"))
    group = ET.SubElement(group_fragment, _wix("ComponentGroup"), {"Id": "PayloadComponents"})

    directory_nodes: dict[Path, ET.Element] = {Path(): app_ref}
    files = sorted(
        (path for path in payload_dir.rglob("*") if path.is_file()),
        key=lambda path: path.relative_to(payload_dir).as_posix(),
    )
    for file_path in files:
        relative = file_path.relative_to(payload_dir)
        parent_node = _ensure_wix_directory(relative.parent, directory_nodes)
        token = hashlib.sha256(relative.as_posix().lower().encode("utf-8")).hexdigest()[:24]
        component_id = f"Cmp_{token}"
        component_guid = "{" + str(uuid.uuid5(COMPONENT_NAMESPACE, relative.as_posix().lower())).upper() + "}"
        component = ET.SubElement(
            parent_node,
            _wix("Component"),
            {"Id": component_id, "Guid": component_guid},
        )
        ET.SubElement(
            component,
            _wix("File"),
            {
                "Id": f"Fil_{token}",
                "Source": str(file_path),
                "KeyPath": "yes",
            },
        )
        ET.SubElement(group, _wix("ComponentRef"), {"Id": component_id})

    tree = ET.ElementTree(wix)
    _indent_xml(tree)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    return output_path


def _indent_xml(tree: ET.ElementTree) -> None:
    indent = getattr(ET, "indent", None)
    if indent is not None:
        indent(tree, space="  ")
        return

    # xml.etree.ElementTree.indent was added in Python 3.9. Keep the exact
    # modern path above and provide only the Win7/Python 3.8 build fallback.
    def apply(element: ET.Element, level: int = 0) -> None:
        padding = "\n" + ("  " * level)
        child_padding = "\n" + ("  " * (level + 1))
        children = list(element)
        if children:
            if not element.text or not element.text.strip():
                element.text = child_padding
            for child in children:
                apply(child, level + 1)
            if not children[-1].tail or not children[-1].tail.strip():
                children[-1].tail = padding
        if level and (not element.tail or not element.tail.strip()):
            element.tail = padding

    apply(tree.getroot())


def validate_installer_definitions() -> None:
    if not INNO_LANGUAGE_FILE.is_file():
        raise RuntimeError(f"缺少固定的 Inno 简体中文语言文件：{INNO_LANGUAGE_FILE}")
    language_sha256 = hashlib.sha256(INNO_LANGUAGE_FILE.read_bytes()).hexdigest()
    if language_sha256 != INNO_LANGUAGE_SHA256:
        raise RuntimeError(
            "Inno 简体中文语言文件校验失败："
            f"期望 {INNO_LANGUAGE_SHA256}，实际 {language_sha256}"
        )

    iss = INNO_SCRIPT.read_text(encoding="utf-8")
    required_iss = (
        "DefaultDirName={localappdata}\\Programs\\HRToolkit",
        "PrivilegesRequired=lowest",
        "ArchitecturesAllowed=x64compatible",
        "MinVersion={#MinWindowsVersion}",
        "#ifdef Win7Compatibility",
        "OnlyBelowVersion=6.3",
        'DestDir: "{app}\\app"',
        'UninstallDisplayIcon={app}\\app\\{#MyAppExeName}',
        'MessagesFile: "compiler:Default.isl,ChineseSimplified.isl"',
        "#ifdef CleanExistingPayload",
        "Check: ExistingPayloadIsWin7",
        "function ExistingPayloadIsWin7: Boolean;",
    )
    missing_iss = [value for value in required_iss if value not in iss]
    if missing_iss:
        raise RuntimeError(f"Inno Setup 配置缺少权限/布局约束：{missing_iss}")

    wix = WIX_SOURCE.read_text(encoding="utf-8")
    required_wix = (
        'Scope="perUser"',
        'Id="LocalAppDataFolder"',
        'Id="APPDIR" Name="app"',
        'Target="[APPDIR]HRToolkit.exe"',
        'Id="PayloadComponents"',
        "$(var.WindowsTarget)",
        "VersionNT64",
        'AllowSameVersionUpgrades="yes"',
    )
    missing_wix = [value for value in required_wix if value not in wix]
    if missing_wix:
        raise RuntimeError(f"WiX 配置缺少权限/布局约束：{missing_wix}")


def verify_installer_outputs(exe_path: Path, msi_path: Path) -> None:
    if not exe_path.is_file() or _read_prefix(exe_path, 2) != b"MZ":
        raise RuntimeError(f"EXE 安装器无效：{exe_path}")
    if not msi_path.is_file() or _read_prefix(msi_path, 8) != MSI_MAGIC:
        raise RuntimeError(f"MSI 安装器无效：{msi_path}")


def smoke_test_installers(
    exe_path: Path,
    msi_path: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> None:
    target = validate_windows_target(target)
    ensure_windows_runtime()
    if target == WINDOWS_TARGET_WIN7 and not _supports_win7_installer_smoke():
        print(
            "跳过 Win7 EXE/MSI 安装冒烟：当前 Windows 不在安装器支持的 "
            "Windows 7 SP1/Windows 8 范围内；真实验收由干净 Win7 SP1 环境完成。"
        )
        return
    with tempfile.TemporaryDirectory(prefix="hr_toolkit_installer_smoke_") as tmp:
        root = Path(tmp)
        _smoke_test_inno(exe_path, root / "inno", target=target)
        _smoke_test_msi(msi_path, root / "msi", target=target)


def _supports_win7_installer_smoke() -> bool:
    try:
        version = sys.getwindowsversion()
        platform_version = getattr(version, "platform_version", ())
        if len(platform_version) >= 2:
            major_minor = (int(platform_version[0]), int(platform_version[1]))
        else:
            major_minor = (int(version.major), int(version.minor))
        service_pack_major = int(getattr(version, "service_pack_major", 0))
    except (AttributeError, IndexError, TypeError, ValueError):
        return False
    return major_minor == (6, 2) or (
        major_minor == (6, 1) and service_pack_major >= 1
    )


def resolve_inno_compiler(explicit: str | None = None) -> str:
    return _resolve_executable(
        explicit or os.environ.get(INNO_ENV),
        names=("ISCC.exe", "iscc"),
        candidates=(
            Path(os.environ.get("ProgramFiles(x86)", "")) / "Inno Setup 6" / "ISCC.exe",
            Path(os.environ.get("ProgramFiles", "")) / "Inno Setup 6" / "ISCC.exe",
        ),
        label="Inno Setup 6 ISCC.exe",
    )


def resolve_wix_executable(explicit: str | None = None) -> str:
    return _resolve_executable(
        explicit or os.environ.get(WIX_ENV),
        names=("wix.exe", "wix"),
        candidates=(
            Path(os.environ.get("USERPROFILE", "")) / ".dotnet" / "tools" / "wix.exe",
        ),
        label="WiX Toolset v4+ wix.exe",
    )


def _ensure_wix_directory(
    relative_dir: Path,
    directory_nodes: dict[Path, ET.Element],
) -> ET.Element:
    if relative_dir in directory_nodes:
        return directory_nodes[relative_dir]
    parent = _ensure_wix_directory(relative_dir.parent, directory_nodes)
    token = hashlib.sha256(relative_dir.as_posix().lower().encode("utf-8")).hexdigest()[:24]
    node = ET.SubElement(
        parent,
        _wix("Directory"),
        {"Id": f"Dir_{token}", "Name": relative_dir.name},
    )
    directory_nodes[relative_dir] = node
    return node


def _smoke_test_inno(
    installer: Path,
    install_root: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> None:
    install_root.parent.mkdir(parents=True, exist_ok=True)
    install_command = [
        str(installer),
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
        "/SP-",
        f"/DIR={install_root}",
    ]
    _run_windows_installer(install_command)
    try:
        payload = install_root / "app"
        verify_staged_payload(payload, target=target)
        uninstallers = sorted(install_root.glob("unins*.exe"))
        if not uninstallers:
            raise RuntimeError("Inno 安装器没有把卸载器保留在 payload 外层。")
        if any(path.parent == payload for path in uninstallers):
            raise RuntimeError("Inno 卸载器错误地位于会被自更新替换的 app 目录。")
        run_runtime_smoke(
            payload / f"{APP_NAME}.exe",
            payload / f"{UPDATER_NAME}.exe",
            target=target,
        )
    finally:
        uninstallers = sorted(install_root.glob("unins*.exe"))
        if uninstallers:
            _run_windows_installer(
                [str(uninstallers[0]), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"]
            )


def _smoke_test_msi(
    installer: Path,
    install_root: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> None:
    install_root.parent.mkdir(parents=True, exist_ok=True)
    install_command = [
        "msiexec.exe",
        "/i",
        str(installer),
        "/qn",
        "/norestart",
        f"INSTALLROOT={install_root}",
    ]
    _run_windows_installer(install_command)
    try:
        payload = install_root / "app"
        verify_staged_payload(payload, target=target)
        run_runtime_smoke(
            payload / f"{APP_NAME}.exe",
            payload / f"{UPDATER_NAME}.exe",
            target=target,
        )
    finally:
        _run_windows_installer(
            ["msiexec.exe", "/x", str(installer), "/qn", "/norestart"],
            check=False,
        )


def _resolve_executable(
    explicit: str | None,
    *,
    names: tuple[str, ...],
    candidates: tuple[Path, ...],
    label: str,
) -> str:
    if explicit:
        explicit_path = Path(explicit)
        if explicit_path.is_file():
            return str(explicit_path)
        resolved = shutil.which(explicit)
        if resolved:
            return resolved
        raise RuntimeError(f"未找到 {label}：{explicit}")
    for name in names:
        resolved = shutil.which(name)
        if resolved:
            return resolved
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    raise RuntimeError(f"未找到 {label}，请安装后加入 PATH，或使用命令行参数指定。")


def _run(command: list[str], *, timeout: int | None = None) -> None:
    print("执行：" + subprocess.list2cmdline(command))
    subprocess.run(command, cwd=REPO_ROOT, check=True, timeout=timeout)


def _run_windows_installer(command: list[str], *, check: bool = True) -> None:
    print("验证：" + subprocess.list2cmdline(command))
    result = subprocess.run(command, cwd=REPO_ROOT, timeout=180, check=False)
    if check and result.returncode not in {0, 3010}:
        raise RuntimeError(
            f"安装器验证失败，退出码 {result.returncode}：{subprocess.list2cmdline(command)}"
        )


def _wix(tag: str) -> str:
    return f"{{{WIX_NAMESPACE}}}{tag}"


def _read_prefix(path: Path, size: int) -> bytes:
    with path.open("rb") as handle:
        return handle.read(size)


if __name__ == "__main__":
    raise SystemExit(main())
