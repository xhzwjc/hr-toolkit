from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable


DEFAULT_UPDATE_MANIFEST_URL = "https://gitee.com/optimistic-little-sunspot/hr-toolkit/raw/main/release/latest.json"
UPDATE_URL_ENV = "HR_TOOLKIT_UPDATE_URL"
SKIP_UPDATE_ENV = "HR_TOOLKIT_SKIP_UPDATE"
FORCE_UPDATE_ENV = "HR_TOOLKIT_FORCE_UPDATE_CHECK"
UPDATER_PATH_ENV = "HR_TOOLKIT_UPDATER_PATH"
UPDATE_URL_FILE = "update_url.txt"
USER_AGENT = "HRToolkit-Updater/1.0"


class UpdateError(RuntimeError):
    """Raised when update metadata, download, or launch fails."""


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    file_url: str
    sha256: str
    notes: tuple[str, ...]
    mandatory: bool
    manifest_url: str


def update_check_enabled() -> bool:
    if _truthy(os.environ.get(SKIP_UPDATE_ENV)):
        return False
    if _truthy(os.environ.get(FORCE_UPDATE_ENV)):
        return True
    return bool(getattr(sys, "frozen", False))


def update_manifest_url() -> str:
    env_url = os.environ.get(UPDATE_URL_ENV, "").strip()
    if env_url:
        return env_url
    config_url = _read_update_url_file()
    if config_url:
        return config_url
    return DEFAULT_UPDATE_MANIFEST_URL


def check_for_update(current_version: str, manifest_url: str | None = None, platform: str | None = None) -> UpdateInfo | None:
    manifest_url = manifest_url or update_manifest_url()
    manifest = fetch_update_manifest(manifest_url)
    remote_version = manifest_version(manifest, platform=platform or platform_key())
    if not is_newer_version(remote_version, current_version):
        return None
    return parse_update_manifest(manifest, manifest_url=manifest_url, platform=platform or platform_key())


def fetch_update_manifest(manifest_url: str, timeout: int = 10) -> dict[str, Any]:
    request = urllib.request.Request(manifest_url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = response.read().decode("utf-8-sig")
    except Exception as exc:
        raise UpdateError(f"无法读取更新配置：{exc}") from exc
    try:
        data = json.loads(payload)
    except json.JSONDecodeError as exc:
        raise UpdateError("更新配置不是有效的 JSON。") from exc
    if not isinstance(data, dict):
        raise UpdateError("更新配置格式不正确。")
    return data


def parse_update_manifest(manifest: dict[str, Any], manifest_url: str, platform: str) -> UpdateInfo:
    platform_payload = _platform_payload(manifest, platform)
    version = _manifest_version_from_payload(platform_payload, manifest)
    file_url = str(
        platform_payload.get("file_url")
        or platform_payload.get("url")
        or manifest.get("file_url")
        or manifest.get("url")
        or ""
    ).strip()
    sha256 = str(platform_payload.get("sha256") or manifest.get("sha256") or "").strip().lower()
    mandatory = bool(platform_payload.get("mandatory", manifest.get("mandatory", True)))
    notes_value = platform_payload.get("notes", manifest.get("notes", ()))

    if not version:
        raise UpdateError("更新配置缺少 version。")
    if not file_url:
        raise UpdateError("更新配置缺少 file_url。")
    if not sha256:
        raise UpdateError("更新配置缺少 sha256。")

    file_url = urllib.parse.urljoin(manifest_url, file_url)
    notes = _normalize_notes(notes_value)
    return UpdateInfo(
        version=version,
        file_url=file_url,
        sha256=sha256,
        notes=notes,
        mandatory=mandatory,
        manifest_url=manifest_url,
    )


def manifest_version(manifest: dict[str, Any], platform: str) -> str:
    platform_payload = _platform_payload(manifest, platform)
    return _manifest_version_from_payload(platform_payload, manifest)


def download_update_package(
    update: UpdateInfo,
    dest_dir: Path | None = None,
    progress_callback: Callable[[int, int], None] | None = None,
) -> Path:
    if dest_dir is None:
        dest_dir = Path(tempfile.mkdtemp(prefix="hr_toolkit_update_"))
    else:
        dest_dir.mkdir(parents=True, exist_ok=True)

    filename = Path(urllib.parse.urlparse(update.file_url).path).name or f"HRToolkit-{update.version}.zip"
    final_path = dest_dir / filename
    temp_path = dest_dir / f"{filename}.download"
    request = urllib.request.Request(update.file_url, headers={"User-Agent": USER_AGENT})

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            total = int(response.headers.get("Content-Length") or 0)
            downloaded = 0
            with temp_path.open("wb") as output:
                while True:
                    chunk = response.read(1024 * 256)
                    if not chunk:
                        break
                    output.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback is not None:
                        progress_callback(downloaded, total)
        os.replace(temp_path, final_path)
    except Exception as exc:
        if temp_path.exists():
            temp_path.unlink()
        raise UpdateError(f"更新包下载失败：{exc}") from exc

    actual_sha256 = sha256_file(final_path)
    if actual_sha256.lower() != update.sha256.lower():
        final_path.unlink(missing_ok=True)
        raise UpdateError("更新包校验失败，请检查 latest.json 中的 sha256。")
    return final_path


def launch_update_replacement(
    package_path: Path,
    app_dir: Path | None = None,
    launcher_path: Path | None = None,
    wait_pid: int | None = None,
) -> None:
    app_dir = app_dir or current_app_dir()
    launcher_path = launcher_path or current_launcher_path()
    updater_path = find_updater_executable(app_dir)
    temp_dir = Path(tempfile.mkdtemp(prefix="hr_toolkit_updater_"))
    temp_updater = temp_dir / updater_path.name
    shutil.copy2(updater_path, temp_updater)
    if not sys.platform.startswith("win"):
        temp_updater.chmod(temp_updater.stat().st_mode | 0o111)

    args = [
        str(temp_updater),
        "--zip",
        str(package_path),
        "--app-dir",
        str(app_dir),
        "--launcher",
        launcher_path.name,
        "--wait-pid",
        str(wait_pid or os.getpid()),
        "--relaunch",
    ]
    subprocess.Popen(args, close_fds=True)


def find_updater_executable(app_dir: Path) -> Path:
    env_path = os.environ.get(UPDATER_PATH_ENV)
    if env_path:
        candidate = Path(env_path)
        if candidate.exists():
            return candidate
    names = ["HRToolkitUpdater.exe"] if sys.platform.startswith("win") else ["HRToolkitUpdater"]
    for name in names:
        candidate = app_dir / name
        if candidate.exists():
            return candidate
        candidate = app_dir / "_internal" / name
        if candidate.exists():
            return candidate
    raise UpdateError("未找到更新程序 HRToolkitUpdater，请重新打包发布。")


def current_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd().resolve()


def current_launcher_path() -> Path:
    return Path(sys.executable).resolve()


def _read_update_url_file() -> str | None:
    for parent in _update_url_search_dirs():
        path = parent / UPDATE_URL_FILE
        if not path.exists():
            continue
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            value = line.strip()
            if value and not value.startswith("#"):
                return value
    return None


def _update_url_search_dirs() -> tuple[Path, ...]:
    dirs = [current_app_dir()]
    if getattr(sys, "frozen", False):
        dirs.append(Path(sys.executable).resolve().parent)
    else:
        dirs.append(Path.cwd().resolve())
    unique: list[Path] = []
    for path in dirs:
        if path not in unique:
            unique.append(path)
    return tuple(unique)


def sha256_file(file_path: Path) -> str:
    digest = hashlib.sha256()
    with file_path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def is_newer_version(remote_version: str, current_version: str) -> bool:
    return _version_parts(remote_version) > _version_parts(current_version)


def platform_key() -> str:
    if sys.platform.startswith("win"):
        return "windows"
    if sys.platform == "darwin":
        return "macos"
    return "linux"


def _platform_payload(manifest: dict[str, Any], platform: str) -> dict[str, Any]:
    platforms = manifest.get("platforms")
    if not isinstance(platforms, dict):
        return manifest
    aliases = {
        "windows": ("windows", "win", "win32"),
        "macos": ("macos", "darwin", "mac"),
        "linux": ("linux",),
    }.get(platform, (platform,))
    for key in aliases:
        payload = platforms.get(key)
        if isinstance(payload, dict):
            merged = dict(manifest)
            merged.pop("platforms", None)
            merged.update(payload)
            return merged
    raise UpdateError(f"更新配置中没有 {platform} 平台的安装包。")


def _normalize_notes(value: Any) -> tuple[str, ...]:
    if isinstance(value, str):
        return (value,) if value.strip() else ()
    if isinstance(value, list):
        return tuple(str(item).strip() for item in value if str(item).strip())
    return ()


def _manifest_version_from_payload(platform_payload: dict[str, Any], manifest: dict[str, Any]) -> str:
    version = str(platform_payload.get("version") or manifest.get("version") or "").strip()
    if not version:
        raise UpdateError("更新配置缺少 version。")
    return version


def _version_parts(value: str) -> tuple[int, ...]:
    parts = [int(part) for part in re.findall(r"\d+", value)]
    while len(parts) < 4:
        parts.append(0)
    return tuple(parts[:4])


def _truthy(value: str | None) -> bool:
    return value is not None and value.strip().lower() in {"1", "true", "yes", "on"}
