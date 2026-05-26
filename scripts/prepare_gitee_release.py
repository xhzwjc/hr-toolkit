from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from versioning import read_project_version


REPO_ROOT = Path(__file__).resolve().parents[1]
GITEE_RAW_BASE = "https://gitee.com/optimistic-little-sunspot/hr-toolkit/raw/main"
SCRIPT_HUB_MANIFEST_URL = "http://hr.seedlingintl.com/api/static/hr-toolkit/latest.json"
SCRIPT_HUB_RELEASE_BASE_URL = "http://hr.seedlingintl.com/api/static/hr-toolkit/releases"
DEFAULT_BUNDLE_DIR = REPO_ROOT / "release" / "scripthub_static" / "hr-toolkit"
UPDATE_URL = SCRIPT_HUB_MANIFEST_URL


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成 Gitee 自动更新发布文件")
    parser.add_argument("--platform", choices=["windows", "macos"], default=_default_platform())
    parser.add_argument("--version", default=read_project_version(), help="发布版本号，默认读取 hr_toolkit.__version__")
    parser.add_argument("--notes", nargs="*", default=["更新 HR工具箱"], help="更新说明，可写多条")
    parser.add_argument("--app-dir", type=Path, default=REPO_ROOT / "dist" / "HRToolkit", help="PyInstaller 输出目录")
    parser.add_argument("--updater", type=Path, help="HRToolkitUpdater 文件路径；默认从 dist 中查找")
    parser.add_argument("--output-dir", type=Path, default=REPO_ROOT / "release" / "downloads", help="zip 输出目录")
    parser.add_argument("--manifest-url", default=SCRIPT_HUB_MANIFEST_URL, help="客户端检查更新的 latest.json 地址")
    parser.add_argument("--release-base-url", default=SCRIPT_HUB_RELEASE_BASE_URL, help="更新包下载目录 URL")
    parser.add_argument("--bundle-dir", type=Path, default=DEFAULT_BUNDLE_DIR, help="生成给 ScriptHub 静态目录的一次性复制文件夹")
    parser.add_argument("--publish-dir", type=Path, help="可选：直接复制到 ScriptHub 的 fastApiProject/static/hr-toolkit 目录")
    args = parser.parse_args(argv)

    app_dir = args.app_dir
    if not app_dir.exists() or not app_dir.is_dir():
        raise SystemExit(f"未找到程序目录：{app_dir}\n请先完成 PyInstaller 打包。")

    updater = args.updater or _default_updater_path(args.platform)
    if not updater.exists():
        raise SystemExit(f"未找到更新程序：{updater}\n请先打包 HRToolkitUpdater。")

    _copy_updater(app_dir, updater, args.platform)
    _write_update_url(app_dir, args.manifest_url)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    zip_path = args.output_dir / f"HRToolkit-{args.version}-{_platform_suffix(args.platform)}.zip"
    _zip_app_dir(app_dir, zip_path)
    digest = _sha256_file(zip_path)
    manifest_path = REPO_ROOT / "release" / "latest.json"
    manifest = _load_manifest(manifest_path)
    manifest["version"] = args.version
    manifest["mandatory"] = True
    manifest["notes"] = args.notes
    platforms = manifest.setdefault("platforms", {})
    platforms[args.platform] = {
        "file_url": f"{args.release_base_url.rstrip('/')}/{zip_path.name}",
        "sha256": digest,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    bundle_dir = _write_static_bundle(args.bundle_dir, manifest_path, zip_path)
    if args.publish_dir is not None:
        _copy_static_bundle(bundle_dir, args.publish_dir)

    print(f"已生成更新包：{zip_path}")
    print(f"SHA256：{digest}")
    print(f"已更新配置：{manifest_path}")
    print(f"已生成 ScriptHub 静态发布目录：{bundle_dir}")
    if args.publish_dir is not None:
        print(f"已复制到 ScriptHub 目录：{args.publish_dir}")
    else:
        print("下一步把 release/scripthub_static/hr-toolkit 整个复制到 ScriptHub 项目的 fastApiProject/static/ 下。")
    return 0


def _default_platform() -> str:
    if sys.platform == "darwin":
        return "macos"
    return "windows"


def _platform_suffix(platform: str) -> str:
    return "win" if platform == "windows" else "mac"


def _default_updater_path(platform: str) -> Path:
    if platform == "windows":
        return REPO_ROOT / "dist" / "HRToolkitUpdater.exe"
    return REPO_ROOT / "dist" / "HRToolkitUpdater"


def _copy_updater(app_dir: Path, updater: Path, platform: str) -> None:
    target_name = "HRToolkitUpdater.exe" if platform == "windows" else "HRToolkitUpdater"
    shutil.copy2(updater, app_dir / target_name)


def _write_update_url(app_dir: Path, manifest_url: str) -> None:
    (app_dir / "update_url.txt").write_text(manifest_url.strip() + "\n", encoding="utf-8")


def _zip_app_dir(app_dir: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for file_path in sorted(app_dir.rglob("*")):
            if file_path.is_file():
                archive.write(file_path, file_path.relative_to(app_dir))


def _sha256_file(file_path: Path) -> str:
    digest = hashlib.sha256()
    with file_path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_manifest(path: Path) -> dict:
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {"version": "0.0.0", "mandatory": True, "notes": [], "platforms": {}}


def _write_static_bundle(bundle_dir: Path, manifest_path: Path, zip_path: Path) -> Path:
    releases_dir = bundle_dir / "releases"
    releases_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(manifest_path, bundle_dir / "latest.json")
    shutil.copy2(zip_path, releases_dir / zip_path.name)
    return bundle_dir


def _copy_static_bundle(bundle_dir: Path, publish_dir: Path) -> None:
    releases_dir = publish_dir / "releases"
    releases_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(bundle_dir / "latest.json", publish_dir / "latest.json")
    for file_path in (bundle_dir / "releases").glob("*.zip"):
        shutil.copy2(file_path, releases_dir / file_path.name)


if __name__ == "__main__":
    raise SystemExit(main())
