from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import stat
import sys
import tempfile
import zipfile
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from build_windows import (
    APP_NAME,
    MODERN_VC_RUNTIME_FILES,
    UPDATER_NAME,
    WIN7_REQUIRED_UCRT_FILES,
    WIN7_REQUIRED_VC_RUNTIME_FILES,
    WINDOWS_TARGET_MODERN,
    WINDOWS_TARGET_WIN7,
    WINDOWS_TARGETS,
    _payload_pe_files,
    run_runtime_smoke,
    validate_build_version,
    validate_windows_target,
    verify_win7_pe_compatibility,
    verify_pe_x64,
    verify_windows_payload,
    windows_setup_asset_name,
)


GITEE_REPOSITORY = "optimistic-little-sunspot/hr-toolkit"
GITEE_LATEST_RELEASE_API_URL = (
    f"https://gitee.com/api/v5/repos/{GITEE_REPOSITORY}/releases/latest"
)
UPDATE_MANIFEST_URLS = (
    GITEE_LATEST_RELEASE_API_URL,
)
LEGACY_MANIFEST_NAME = "legacy-server-latest.json"
WIN7_LEGACY_MANIFEST_NAME = "legacy-server-latest-win7.json"
ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="校验 Windows onedir 并生成旧服务器桥接更新清单。"
    )
    parser.add_argument("--version", required=True, help="必须与 hr_toolkit.__version__ 一致")
    parser.add_argument("--app-dir", required=True, type=Path, help="PyInstaller HRToolkit onedir")
    parser.add_argument("--updater", required=True, type=Path, help="HRToolkitUpdater.exe")
    parser.add_argument("--output-dir", required=True, type=Path, help="更新资产输出目录")
    parser.add_argument("--notes", nargs="*", default=None, help="旧服务器桥接更新说明")
    parser.add_argument("--optional", action="store_true", help="旧服务器桥接清单标记为可选更新")
    parser.add_argument(
        "--target",
        choices=WINDOWS_TARGETS,
        default=WINDOWS_TARGET_MODERN,
        help="modern 保持现有更新资产；win7 生成独立的 Windows 7 更新资产",
    )
    parser.add_argument(
        "--skip-runtime-smoke",
        action="store_true",
        help="已完成等价构建/安装验收时，跳过重复 staging 和无界面启动检查",
    )
    args = parser.parse_args(argv)

    version = validate_build_version(args.version)
    setup_path, manifest_path = build_update_assets(
        version=version,
        app_dir=args.app_dir.resolve(),
        updater=args.updater.resolve(),
        output_dir=args.output_dir.resolve(),
        notes=args.notes,
        mandatory=not args.optional,
        runtime_smoke=not args.skip_runtime_smoke,
        target=args.target,
    )
    print(f"Windows 安装更新包：{setup_path}")
    print(f"旧服务器桥接清单：{manifest_path}")
    return 0


def build_update_assets(
    *,
    version: str,
    app_dir: Path,
    updater: Path,
    output_dir: Path,
    notes: list[str] | None = None,
    mandatory: bool = True,
    runtime_smoke: bool = True,
    target: str = WINDOWS_TARGET_MODERN,
) -> tuple[Path, Path]:
    validate_build_version(version)
    target = validate_windows_target(target)
    output_dir.mkdir(parents=True, exist_ok=True)
    setup_name = windows_setup_asset_name(version, target)
    setup_path = output_dir / setup_name
    manifest_path = output_dir / legacy_manifest_name(target)

    if runtime_smoke:
        with tempfile.TemporaryDirectory(prefix="hr_toolkit_windows_payload_") as tmp:
            payload_dir = Path(tmp) / APP_NAME
            stage_windows_payload(
                app_dir=app_dir,
                updater=updater,
                target_dir=payload_dir,
                target=target,
            )
            run_runtime_smoke(
                payload_dir / f"{APP_NAME}.exe",
                payload_dir / f"{UPDATER_NAME}.exe",
                target=target,
            )
    else:
        verify_windows_update_sources(app_dir=app_dir, updater=updater, target=target)

    digest = sha256_file(setup_path) if setup_path.exists() else ("0" * 64)
    manifest = legacy_server_manifest(
        version=version,
        filename=setup_name,
        sha256=digest,
        notes=notes,
        mandatory=mandatory,
        target=target,
    )
    _write_json_atomically(manifest_path, manifest)
    return setup_path, manifest_path


def stage_windows_payload(
    *,
    app_dir: Path,
    updater: Path,
    target_dir: Path,
    target: str = WINDOWS_TARGET_MODERN,
) -> Path:
    target = validate_windows_target(target)
    verify_windows_update_sources(app_dir=app_dir, updater=updater, target=target)
    if target_dir.exists():
        raise RuntimeError(f"staging 目录必须不存在：{target_dir}")

    shutil.copytree(app_dir, target_dir)
    shutil.copy2(updater, target_dir / f"{UPDATER_NAME}.exe")
    (target_dir / "update_url.txt").write_text(
        "\n".join(UPDATE_MANIFEST_URLS) + "\n",
        encoding="utf-8",
    )
    verify_staged_payload(target_dir, target=target)
    return target_dir


def verify_windows_update_sources(
    *,
    app_dir: Path,
    updater: Path,
    target: str = WINDOWS_TARGET_MODERN,
) -> None:
    target = validate_windows_target(target)
    verify_windows_payload(app_dir, target=target)
    verify_pe_x64(app_dir / f"{APP_NAME}.exe")
    verify_pe_x64(updater)
    if target == WINDOWS_TARGET_WIN7:
        verify_win7_pe_compatibility((*_payload_pe_files(app_dir), updater))
    if updater.name.lower() != f"{UPDATER_NAME}.exe".lower():
        raise RuntimeError(f"更新程序名称必须为 {UPDATER_NAME}.exe：{updater}")


def verify_staged_payload(
    payload_dir: Path,
    *,
    target: str = WINDOWS_TARGET_MODERN,
) -> None:
    target = validate_windows_target(target)
    verify_windows_payload(payload_dir, target=target)
    verify_pe_x64(payload_dir / f"{APP_NAME}.exe")
    verify_pe_x64(payload_dir / f"{UPDATER_NAME}.exe")
    root_files = {path.name for path in payload_dir.iterdir() if path.is_file()}
    expected_root_files = {
        f"{APP_NAME}.exe",
        f"{UPDATER_NAME}.exe",
        "update_url.txt",
    }
    if target == WINDOWS_TARGET_WIN7:
        expected_root_files.update(WIN7_REQUIRED_UCRT_FILES)
        expected_root_files.update(WIN7_REQUIRED_VC_RUNTIME_FILES)
    else:
        expected_root_files.update(
            name for name in MODERN_VC_RUNTIME_FILES if (payload_dir / name).is_file()
        )
    if root_files != expected_root_files:
        raise RuntimeError(
            f"Windows 更新 payload 根文件不符合白名单，实际={sorted(root_files)}"
        )
    update_urls = tuple(
        line.strip()
        for line in (payload_dir / "update_url.txt").read_text(encoding="utf-8").splitlines()
        if line.strip()
    )
    if update_urls != UPDATE_MANIFEST_URLS:
        raise RuntimeError(f"update_url.txt 地址不正确：{update_urls}")


def legacy_manifest_name(target: str = WINDOWS_TARGET_MODERN) -> str:
    target = validate_windows_target(target)
    return WIN7_LEGACY_MANIFEST_NAME if target == WINDOWS_TARGET_WIN7 else LEGACY_MANIFEST_NAME


def update_zip_name(version: str) -> str:
    validate_build_version(version)
    return f"HRToolkit-{version}-win-update.zip"


def legacy_server_manifest(
    *,
    version: str,
    filename: str,
    sha256: str,
    notes: list[str] | None,
    mandatory: bool,
    target: str = WINDOWS_TARGET_MODERN,
) -> dict:
    validate_build_version(version)
    target = validate_windows_target(target)
    if len(sha256) != 64 or any(character not in "0123456789abcdef" for character in sha256.lower()):
        raise ValueError("sha256 必须是 64 位十六进制字符串。")
    gitee_release_url = (
        f"https://gitee.com/{GITEE_REPOSITORY}/releases/download/v{version}/{filename}"
    )
    normalized_notes = [str(note).strip() for note in (notes or []) if str(note).strip()]
    if not normalized_notes:
        normalized_notes = [f"升级 HRToolkit 至 {version}"]
    platform_name = "windows-x64-win7" if target == WINDOWS_TARGET_WIN7 else "windows"
    return {
        "version": version,
        "mandatory": bool(mandatory),
        "notes": normalized_notes,
        "platforms": {
            platform_name: {
                "version": version,
                "file_url": gitee_release_url,
                "sha256": sha256.lower(),
            }
        },
    }


def write_deterministic_zip(payload_dir: Path, zip_path: Path) -> None:
    payload_dir = payload_dir.resolve()
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = zip_path.with_name(zip_path.name + ".tmp")
    temporary.unlink(missing_ok=True)
    try:
        with zipfile.ZipFile(
            temporary,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=9,
        ) as archive:
            files = sorted(
                (path for path in payload_dir.rglob("*") if path.is_file()),
                key=lambda path: path.relative_to(payload_dir).as_posix(),
            )
            for path in files:
                relative = path.relative_to(payload_dir).as_posix()
                info = zipfile.ZipInfo(relative, date_time=ZIP_EPOCH)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 3
                info.external_attr = (stat.S_IFREG | 0o644) << 16
                with path.open("rb") as source, archive.open(info, "w") as target:
                    shutil.copyfileobj(source, target, length=1024 * 1024)
        os.replace(temporary, zip_path)
    finally:
        temporary.unlink(missing_ok=True)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _write_json_atomically(path: Path, payload: dict) -> None:
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


if __name__ == "__main__":
    raise SystemExit(main())
