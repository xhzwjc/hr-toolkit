from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import urllib.parse
from pathlib import Path
from typing import Iterable, Optional, Sequence


REPO_ROOT = Path(__file__).resolve().parents[1]
VERSION_FILE = REPO_ROOT / "hr_toolkit" / "__init__.py"
SEMVER_PATTERN = re.compile(r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")
REPOSITORY_PATTERN = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$")
VERSION_ASSIGNMENT_PATTERN = re.compile(r'^__version__\s*=\s*"([^"]+)"\s*$', re.MULTILINE)
# GitHub Release accepts individual assets below 2 GiB.  Installer generation
# and unified metadata target GitHub, so they must not inherit Gitee's much
# smaller attachment ceiling.
MAX_RELEASE_ASSET_BYTES = 2 * 1024 * 1024 * 1024
# Gitee publishing remains independently constrained and currently mirrors
# only source metadata/checksums; retain its conservative decimal 100 MB cap.
GITEE_ATTACHMENT_SAFE_MAX_BYTES = 100_000_000
UPDATE_REPOSITORY = "optimistic-little-sunspot/hr-toolkit"
UPDATE_DOWNLOAD_BASE_URL = f"https://gitee.com/{UPDATE_REPOSITORY}/releases/download"


class ReleaseMetadataError(RuntimeError):
    """Raised when release assets or version metadata are inconsistent."""


def require_release_asset_under_limit(
    path: Path,
    *,
    max_bytes: int = MAX_RELEASE_ASSET_BYTES,
) -> int:
    """Require one direct release asset to stay strictly below its byte limit."""
    if max_bytes < 1:
        raise ReleaseMetadataError("发布资产单文件上限必须是正整数。")
    if not path.is_file():
        raise ReleaseMetadataError(f"发布资产不存在或不是普通文件：{path}")
    size = path.stat().st_size
    if size >= max_bytes:
        raise ReleaseMetadataError(
            f"发布资产 {path.name} 体积超限：{size} 字节，必须严格小于 {max_bytes} 字节。"
        )
    return size


def require_release_assets_under_limit(
    paths: Iterable[Path],
    *,
    max_bytes: int = MAX_RELEASE_ASSET_BYTES,
) -> tuple[tuple[Path, int], ...]:
    """Apply the same strict size gate to every supplied direct release asset."""
    return tuple(
        (path, require_release_asset_under_limit(path, max_bytes=max_bytes))
        for path in paths
    )


def validate_version(version: str) -> str:
    if not SEMVER_PATTERN.fullmatch(version):
        raise ReleaseMetadataError(f"版本号必须是严格的 MAJOR.MINOR.PATCH SemVer：{version!r}")
    return version


def read_project_version(version_file: Path = VERSION_FILE) -> str:
    match = VERSION_ASSIGNMENT_PATTERN.search(version_file.read_text(encoding="utf-8"))
    if match is None:
        raise ReleaseMetadataError(f"无法从 {version_file} 读取 __version__")
    return validate_version(match.group(1))


def validate_release_identity(version: str, tag: str, project_version: str) -> None:
    validate_version(version)
    validate_version(project_version)
    expected_tag = f"v{version}"
    if tag != expected_tag:
        raise ReleaseMetadataError(f"Tag 与版本不一致：期望 {expected_tag}，实际 {tag}")
    if project_version != version:
        raise ReleaseMetadataError(
            f"Tag 版本 {version} 与 hr_toolkit.__version__ {project_version} 不一致"
        )


def release_asset_names(version: str, *, mac_variant: str) -> tuple[str, ...]:
    validate_version(version)
    windows = (
        f"HRToolkit_{version}_x64-setup.exe",
        f"HRToolkit_{version}_x64.msi",
        f"HRToolkit_{version}_win7_x64-setup.exe",
        f"HRToolkit_{version}_win7_x64.msi",
    )
    if mac_variant == "universal2":
        mac = (f"HRToolkit_{version}_universal.dmg",)
    elif mac_variant == "split":
        mac = (
            f"HRToolkit_{version}_arm64.dmg",
            f"HRToolkit_{version}_x64.dmg",
        )
    else:
        raise ReleaseMetadataError(f"未知 macOS 资产模式：{mac_variant}")
    return windows + mac


def detect_mac_variant(assets_dir: Path, version: str) -> str:
    universal = assets_dir / f"HRToolkit_{version}_universal.dmg"
    arm64 = assets_dir / f"HRToolkit_{version}_arm64.dmg"
    x64 = assets_dir / f"HRToolkit_{version}_x64.dmg"
    if universal.exists():
        if arm64.exists() or x64.exists():
            raise ReleaseMetadataError("不能同时发布 universal 与分架构 macOS DMG")
        return "universal2"
    if arm64.exists() and x64.exists():
        return "split"
    if arm64.exists() or x64.exists():
        raise ReleaseMetadataError("universal2 失败后必须同时提供 arm64 与 x64 两个真实架构 DMG")
    raise ReleaseMetadataError("缺少 macOS DMG：需要一个 universal DMG 或 arm64+x64 两个 DMG")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validated_url(value: str, *, label: str, strip_trailing_slash: bool = False) -> str:
    normalized = value.strip()
    parsed = urllib.parse.urlparse(normalized)
    if parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.password:
        raise ReleaseMetadataError(f"{label} 必须是无凭据的 HTTPS URL：{value!r}")
    if strip_trailing_slash:
        normalized = normalized.rstrip("/")
    return normalized


def _asset_url(download_base_url: str, tag: str, name: str) -> str:
    return f"{download_base_url}/{tag}/{name}"


def _asset_payload(
    assets_dir: Path,
    tag: str,
    version: str,
    name: str,
    *,
    update_mode: str,
    download_base_url: str,
    fallback_download_base_url: str | None,
    primary_download_max_bytes: int | None,
    primary_download_asset_names: frozenset[str] | None,
) -> dict:
    asset_path = assets_dir / name
    primary_base_url = download_base_url
    secondary_base_url = fallback_download_base_url
    exceeds_primary_limit = (
        primary_download_max_bytes is not None
        and asset_path.stat().st_size >= primary_download_max_bytes
    )
    excluded_from_primary = (
        primary_download_asset_names is not None
        and name not in primary_download_asset_names
    )
    if exceeds_primary_limit or excluded_from_primary:
        if not fallback_download_base_url:
            raise ReleaseMetadataError(
                f"资产 {name} 不使用主下载源，但未配置备用下载基础地址。"
            )
        primary_base_url = fallback_download_base_url
        secondary_base_url = None
    payload = {
        "version": version,
        "file_url": _asset_url(primary_base_url, tag, name),
        "sha256": sha256_file(asset_path),
        "update_mode": update_mode,
    }
    if secondary_base_url:
        payload["fallback_urls"] = [
            _asset_url(secondary_base_url, tag, name)
        ]
    return payload


def build_latest_manifest(
    assets_dir: Path,
    *,
    version: str,
    tag: str,
    repository: str,
    notes: Sequence[str],
    mandatory: bool,
    mac_variant: str,
    download_base_url: str | None = None,
    release_url: str | None = None,
    fallback_download_base_url: str | None = None,
    primary_download_max_bytes: int | None = None,
    primary_download_asset_names: Sequence[str] | None = None,
) -> dict:
    if primary_download_max_bytes is not None and primary_download_max_bytes < 1:
        raise ReleaseMetadataError("主下载源单文件上限必须是正整数。")
    download_base_url = _validated_url(
        download_base_url or UPDATE_DOWNLOAD_BASE_URL,
        label="下载基础地址",
        strip_trailing_slash=True,
    )
    release_url = _validated_url(
        release_url or f"https://gitee.com/{UPDATE_REPOSITORY}/releases/tag/{tag}",
        label="Release 页面地址",
    )
    if fallback_download_base_url:
        fallback_download_base_url = _validated_url(
            fallback_download_base_url,
            label="备用下载基础地址",
            strip_trailing_slash=True,
        )
    primary_asset_names = (
        frozenset(primary_download_asset_names)
        if primary_download_asset_names is not None
        else None
    )
    windows_installer = f"HRToolkit_{version}_x64-setup.exe"
    win7_installer = f"HRToolkit_{version}_win7_x64-setup.exe"
    modern_payload = _asset_payload(
        assets_dir,
        tag,
        version,
        windows_installer,
        update_mode="auto",
        download_base_url=download_base_url,
        fallback_download_base_url=fallback_download_base_url,
        primary_download_max_bytes=primary_download_max_bytes,
        primary_download_asset_names=primary_asset_names,
    )
    platforms = {
        # windows 保留给旧版现代客户端；新客户端使用显式架构通道。
        "windows": dict(modern_payload),
        "windows-x64-modern": dict(modern_payload),
        "windows-x64-win7": _asset_payload(
            assets_dir,
            tag,
            version,
            win7_installer,
            update_mode="auto",
            download_base_url=download_base_url,
            fallback_download_base_url=fallback_download_base_url,
            primary_download_max_bytes=primary_download_max_bytes,
            primary_download_asset_names=primary_asset_names,
        )
    }
    if mac_variant == "universal2":
        mac_name = f"HRToolkit_{version}_universal.dmg"
        platforms["macos"] = _asset_payload(
            assets_dir,
            tag,
            version,
            mac_name,
            update_mode="manual",
            download_base_url=download_base_url,
            fallback_download_base_url=fallback_download_base_url,
            primary_download_max_bytes=primary_download_max_bytes,
            primary_download_asset_names=primary_asset_names,
        )
    else:
        for platform_key, suffix in (("macos-arm64", "arm64"), ("macos-x64", "x64")):
            mac_name = f"HRToolkit_{version}_{suffix}.dmg"
            platforms[platform_key] = _asset_payload(
                assets_dir,
                tag,
                version,
                mac_name,
                update_mode="manual",
                download_base_url=download_base_url,
                fallback_download_base_url=fallback_download_base_url,
                primary_download_max_bytes=primary_download_max_bytes,
                primary_download_asset_names=primary_asset_names,
            )

    return {
        "version": version,
        "mandatory": mandatory,
        "notes": list(notes),
        "release_url": release_url,
        "platforms": platforms,
    }


def _validate_asset_directory(assets_dir: Path, expected_names: Iterable[str]) -> None:
    if not assets_dir.is_dir():
        raise ReleaseMetadataError(f"发布资产目录不存在：{assets_dir}")
    expected = set(expected_names)
    metadata_names = {"latest.json", "SHA256SUMS.txt"}
    actual_files = {path.name for path in assets_dir.iterdir() if path.is_file()}
    unexpected = actual_files - expected - metadata_names
    missing = expected - actual_files
    if missing:
        raise ReleaseMetadataError("缺少发布资产：" + ", ".join(sorted(missing)))
    if unexpected:
        raise ReleaseMetadataError("发布目录包含非白名单文件：" + ", ".join(sorted(unexpected)))
    for name in sorted(expected):
        path = assets_dir / name
        if not path.is_file() or path.stat().st_size <= 0:
            raise ReleaseMetadataError(f"发布资产为空或不是普通文件：{path}")


def _atomic_write_text(path: Path, content: str) -> None:
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(content, encoding="utf-8")
    temporary.replace(path)


def generate_release_metadata(
    assets_dir: Path,
    *,
    version: str,
    tag: str,
    repository: str,
    project_version: str,
    notes: Optional[Sequence[str]] = None,
    mandatory: bool = True,
    download_base_url: str | None = None,
    release_url: str | None = None,
    fallback_download_base_url: str | None = None,
    primary_download_max_bytes: int | None = None,
    primary_download_asset_names: Sequence[str] | None = None,
) -> tuple[Path, Path, tuple[str, ...]]:
    validate_release_identity(version, tag, project_version)
    if not REPOSITORY_PATTERN.fullmatch(repository):
        raise ReleaseMetadataError(f"仓库名必须是 owner/repo：{repository!r}")
    mac_variant = detect_mac_variant(assets_dir, version)
    asset_names = release_asset_names(version, mac_variant=mac_variant)
    _validate_asset_directory(assets_dir, asset_names)
    require_release_assets_under_limit(assets_dir / name for name in asset_names)

    normalized_notes = tuple(note.strip() for note in (notes or (f"HR Toolkit v{version}",)) if note.strip())
    if not normalized_notes:
        normalized_notes = (f"HR Toolkit v{version}",)

    manifest = build_latest_manifest(
        assets_dir,
        version=version,
        tag=tag,
        repository=repository,
        notes=normalized_notes,
        mandatory=mandatory,
        mac_variant=mac_variant,
        download_base_url=download_base_url,
        release_url=release_url,
        fallback_download_base_url=fallback_download_base_url,
        primary_download_max_bytes=primary_download_max_bytes,
        primary_download_asset_names=primary_download_asset_names,
    )
    latest_path = assets_dir / "latest.json"
    _atomic_write_text(latest_path, json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")

    checksum_names = tuple(sorted(asset_names + (latest_path.name,)))
    checksum_lines = [f"{sha256_file(assets_dir / name)}  {name}" for name in checksum_names]
    checksums_path = assets_dir / "SHA256SUMS.txt"
    _atomic_write_text(checksums_path, "\n".join(checksum_lines) + "\n")
    return latest_path, checksums_path, asset_names + (latest_path.name, checksums_path.name)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="验证跨平台资产并生成 Release 元数据")
    parser.add_argument("--version", required=True, help="严格 MAJOR.MINOR.PATCH 版本")
    parser.add_argument("--tag", help="Git Tag，默认 v<version>")
    parser.add_argument("--assets-dir", type=Path, default=Path("release-assets"))
    parser.add_argument(
        "--repository",
        default=os.environ.get("GITHUB_REPOSITORY", "xhzwjc/hr-toolkit"),
        help="owner/repo（默认 GitHub 主仓库）",
    )
    parser.add_argument("--project-version-file", type=Path, default=VERSION_FILE)
    parser.add_argument("--notes", nargs="*", help="latest.json 更新说明")
    parser.add_argument(
        "--download-base-url",
        help="客户端资产下载基础地址，默认仅使用 Gitee releases/download",
    )
    parser.add_argument("--release-url", help="latest.json 中的 Release 页面地址")
    parser.add_argument(
        "--fallback-download-base-url",
        help="显式配置备用地址；正式客户端只使用 Gitee，默认不设置备用源",
    )
    parser.add_argument(
        "--primary-download-max-bytes",
        type=_positive_int,
        help="主下载源允许的单文件字节上限；超限资产改用备用源作为主地址",
    )
    parser.add_argument(
        "--primary-download-asset",
        action="append",
        dest="primary_download_assets",
        help="允许使用主下载源的平台资产名；可重复指定",
    )
    parser.add_argument("--optional", action="store_true", help="将更新标记为非强制")
    parser.add_argument("--check-only", action="store_true", help="只检查 Tag 与项目版本，不读取资产")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    version = validate_version(args.version)
    tag = args.tag or f"v{version}"
    project_version = read_project_version(args.project_version_file)
    validate_release_identity(version, tag, project_version)
    if args.check_only:
        print(f"版本检查通过：{tag} == hr_toolkit.__version__ {project_version}")
        return 0

    latest_path, checksums_path, asset_names = generate_release_metadata(
        args.assets_dir,
        version=version,
        tag=tag,
        repository=args.repository,
        project_version=project_version,
        notes=args.notes,
        mandatory=not args.optional,
        download_base_url=args.download_base_url,
        release_url=args.release_url,
        fallback_download_base_url=args.fallback_download_base_url,
        primary_download_max_bytes=args.primary_download_max_bytes,
        primary_download_asset_names=args.primary_download_assets,
    )
    print(f"已生成：{latest_path}")
    print(f"已生成：{checksums_path}")
    print("Release 直接资产：")
    for name in sorted(asset_names):
        print(f"- {name}")
    return 0


def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed < 1:
        raise argparse.ArgumentTypeError("必须是正整数")
    return parsed


if __name__ == "__main__":
    raise SystemExit(main())
