from __future__ import annotations

import hashlib
import json
import os
import platform as platform_module
import re
import shutil
import ssl
import stat
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable


GITEE_REPOSITORY = "optimistic-little-sunspot/hr-toolkit"
GITEE_LATEST_RELEASE_API_URL = f"https://gitee.com/api/v5/repos/{GITEE_REPOSITORY}/releases/latest"
DEFAULT_UPDATE_MANIFEST_URLS = (
    GITEE_LATEST_RELEASE_API_URL,
)
# 保留单地址常量，兼容已有调用；默认值现在是国内 Gitee 源。
DEFAULT_UPDATE_MANIFEST_URL = DEFAULT_UPDATE_MANIFEST_URLS[0]
UPDATE_URL_ENV = "HR_TOOLKIT_UPDATE_URL"
SKIP_UPDATE_ENV = "HR_TOOLKIT_SKIP_UPDATE"
FORCE_UPDATE_ENV = "HR_TOOLKIT_FORCE_UPDATE_CHECK"
UPDATER_PATH_ENV = "HR_TOOLKIT_UPDATER_PATH"
UPDATE_URL_FILE = "update_url.txt"
UPDATE_LOG_FILE = "HRToolkit_update.log"
USER_AGENT = "HRToolkit-Updater/1.0"
UPDATE_TEMP_PREFIXES = ("hr_toolkit_update_", "hr_toolkit_updater_", "hr_toolkit_extract_")
# 解压 + 新旧两份目录，预留下载包体积的数倍空间
DISK_SPACE_FACTOR = 4
UPDATE_LOG_MAX_BYTES = 1024 * 1024
UPDATE_LOG_KEEP_BYTES = 256 * 1024
UPDATE_MANIFEST_MAX_BYTES = 2 * 1024 * 1024
UPDATE_PACKAGE_MAX_BYTES = 4 * 1024 * 1024 * 1024
UPDATE_UPDATER_MAX_BYTES = 256 * 1024 * 1024
UPDATE_ZIP_MAX_COMPRESSION_RATIO = 1000
WIN7_UPDATER_APP_LOCAL_RUNTIME_FILES = (
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
    "msvcp140.dll",
    "vcruntime140.dll",
    "vcruntime140_1.dll",
)


class UpdateError(RuntimeError):
    """Raised when update metadata, download, or launch fails."""


class UpdateCancelledError(UpdateError):
    """Raised when update metadata, download, or launch is cancelled by the user."""


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    file_url: str
    sha256: str
    notes: tuple[str, ...]
    mandatory: bool
    manifest_url: str
    update_mode: str = "auto"
    fallback_urls: tuple[str, ...] = ()

    @property
    def download_urls(self) -> tuple[str, ...]:
        # Public desktop builds serve domestic users from Gitee only.  Release
        # metadata may retain GitHub mirrors for operators, but the application
        # itself must not probe or download from GitHub.
        return tuple(
            url
            for url in _dedupe_urls((self.file_url, *self.fallback_urls))
            if not _is_github_url(url)
        )


def create_https_context() -> ssl.SSLContext:
    """Create a validating TLS context backed by the bundled Mozilla CA store."""
    try:
        import certifi

        context = ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        try:
            context = ssl.create_default_context()
        except (OSError, ssl.SSLError) as exc:
            raise UpdateError(f"无法加载 HTTPS 根证书：{exc}") from exc
    except (OSError, ssl.SSLError) as exc:
        raise UpdateError(f"无法加载 HTTPS 根证书：{exc}") from exc
    if context.verify_mode != ssl.CERT_REQUIRED or not context.check_hostname:
        raise UpdateError("HTTPS 证书校验未正确启用。")
    return context


def _open_url(request: urllib.request.Request, *, timeout: int):
    if _is_github_url(request.full_url):
        raise UpdateError("客户端更新仅使用 Gitee，不会访问 GitHub。")
    scheme = urllib.parse.urlparse(request.full_url).scheme.lower()
    if scheme not in {"https", "http", "file"}:
        raise UpdateError(f"不支持的更新地址协议：{scheme or '空'}。")
    kwargs: dict[str, Any] = {"timeout": timeout}
    if scheme == "https":
        kwargs["context"] = create_https_context()
    return urllib.request.urlopen(request, **kwargs)


def update_check_enabled() -> bool:
    if _truthy(os.environ.get(SKIP_UPDATE_ENV)):
        return False
    if _truthy(os.environ.get(FORCE_UPDATE_ENV)):
        return True
    return bool(getattr(sys, "frozen", False))


def update_manifest_url() -> str:
    return update_manifest_urls()[0]


def update_manifest_urls() -> tuple[str, ...]:
    env_urls = tuple(
        url
        for url in _normalize_url_lines(os.environ.get(UPDATE_URL_ENV, ""))
        if not _is_github_url(url)
    )
    if env_urls:
        return env_urls
    config_urls = tuple(
        url for url in _read_update_url_files() if not _is_github_url(url)
    )
    if config_urls:
        return config_urls
    return DEFAULT_UPDATE_MANIFEST_URLS


def check_for_update(current_version: str, manifest_url: str | None = None, platform: str | None = None) -> UpdateInfo | None:
    manifest_urls = (manifest_url,) if manifest_url else update_manifest_urls()
    selected_platform = platform or platform_key()
    errors: list[str] = []
    for candidate_url in manifest_urls:
        try:
            manifest, resolved_manifest_url = load_update_manifest(candidate_url)
            remote_version = manifest_version(manifest, platform=selected_platform)
            if not is_newer_version(remote_version, current_version):
                return None
            update = parse_update_manifest(
                manifest,
                manifest_url=resolved_manifest_url,
                platform=selected_platform,
            )
            if not update.download_urls:
                raise UpdateError("此平台尚无可用的 Gitee 安装包，请联系发布者补充后重试。")
            return update
        except Exception as exc:
            errors.append(f"{_update_source_name(candidate_url)}：{exc}")
    raise UpdateError("所有更新源均不可用，已按顺序尝试：" + "；".join(errors))


def fetch_update_manifest(manifest_url: str, timeout: int = 10) -> dict[str, Any]:
    manifest, _resolved_url = load_update_manifest(manifest_url, timeout=timeout)
    return manifest


def load_update_manifest(manifest_url: str, timeout: int = 10) -> tuple[dict[str, Any], str]:
    data = _fetch_json_object(manifest_url, timeout=timeout)
    if _is_release_discovery_url(manifest_url):
        asset_url = _find_manifest_asset_url(data)
        if not asset_url:
            raise UpdateError("最新 Release 缺少 latest.json 附件。")
        is_gitee = urllib.parse.urlparse(manifest_url).hostname == "gitee.com"
        if is_gitee and _is_github_url(asset_url):
            raise UpdateError("Gitee 更新配置不能指向 GitHub。")
        manifest = _fetch_json_object(asset_url, timeout=timeout)
        if is_gitee:
            manifest = _bind_gitee_release_assets(manifest, data, manifest_url)
        return manifest, asset_url
    return data, manifest_url


def _bind_gitee_release_assets(
    manifest: dict[str, Any], release: dict[str, Any], discovery_url: str,
) -> dict[str, Any]:
    """Use only attachments actually uploaded to this Gitee release.

    Older releases can contain a latest.json copied from GitHub. Resolve its
    exact asset name against Gitee's own list; never guess a mirror URL or
    query GitHub. Keep versions, platform lanes and SHA256 values unchanged.
    """
    path = urllib.parse.urlparse(discovery_url).path
    match = re.fullmatch(r"/api/v5/repos/([^/]+/[^/]+)/releases/latest/?", path)
    if match is None:
        return manifest
    tag = str(release.get("tag_name") or "")
    if tag != "v" + str(manifest.get("version") or "").lstrip("v"):
        raise UpdateError("Gitee 发行版与更新配置的版本不一致。")
    prefix = f"https://gitee.com/{match.group(1)}/releases/download/{tag}/"
    assets = release.get("assets") or release.get("attach_files") or []
    available: dict[str, str] = {}
    for asset in assets:
        if not isinstance(asset, dict):
            continue
        name = str(asset.get("name") or "")
        url = str(asset.get("browser_download_url") or asset.get("url") or "")
        if (
            name and "/" not in name and "\\" not in name
            and url == prefix + urllib.parse.quote(name)
        ):
            available[name] = url

    def resolve(payload: dict[str, Any]) -> dict[str, Any] | None:
        original_url = str(
            payload.get("file_url") or payload.get("url")
            or manifest.get("file_url") or manifest.get("url") or ""
        )
        filename = urllib.parse.unquote(urllib.parse.urlparse(original_url).path.rsplit("/", 1)[-1])
        url = available.get(filename)
        if not url:
            return None
        normalized = dict(payload)
        normalized["file_url"] = url
        for field in ("url", "fallback_urls", "fallback_url"):
            normalized.pop(field, None)
        return normalized

    result = dict(manifest)
    result["release_url"] = f"https://gitee.com/{match.group(1)}/releases/tag/{tag}"
    if isinstance(manifest.get("platforms"), dict):
        for field in ("file_url", "url", "fallback_urls", "fallback_url"):
            result.pop(field, None)
        platforms = {}
        for key, payload in manifest["platforms"].items():
            if isinstance(payload, dict):
                resolved = resolve(payload)
                if resolved is not None:
                    platforms[key] = resolved
        result["platforms"] = platforms
    elif manifest.get("file_url") or manifest.get("url"):
        result = resolve(result)
        if result is None:
            raise UpdateError("Gitee 尚未上传此平台安装包，请稍后重试。")
    else:
        return manifest
    return result


def _fetch_json_object(url: str, *, timeout: int) -> dict[str, Any]:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with _open_url(request, timeout=timeout) as response:
            content_length = _response_content_length(response)
            if content_length > UPDATE_MANIFEST_MAX_BYTES:
                raise UpdateError("更新配置文件过大，已拒绝读取。")
            raw_payload = response.read(UPDATE_MANIFEST_MAX_BYTES + 1)
            if len(raw_payload) > UPDATE_MANIFEST_MAX_BYTES:
                raise UpdateError("更新配置文件过大，已拒绝读取。")
            payload = raw_payload.decode("utf-8-sig")
    except UpdateError:
        raise
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
    default_update_mode = "manual" if platform == "macos" else "auto"
    update_mode = str(platform_payload.get("update_mode") or default_update_mode).strip().lower()
    fallback_urls = _normalize_manifest_urls(
        platform_payload.get("fallback_urls") or platform_payload.get("fallback_url"),
        manifest_url,
    )

    if not version:
        raise UpdateError("更新配置缺少 version。")
    if not file_url:
        raise UpdateError("更新配置缺少 file_url。")
    if not sha256:
        raise UpdateError("更新配置缺少 sha256。")
    if update_mode not in {"auto", "manual"}:
        raise UpdateError("更新配置中的 update_mode 只能是 auto 或 manual。")

    file_url = urllib.parse.urljoin(manifest_url, file_url)
    manifest_scheme = urllib.parse.urlparse(manifest_url).scheme.lower()
    for candidate_url in (file_url, *fallback_urls):
        candidate_scheme = urllib.parse.urlparse(candidate_url).scheme.lower()
        if candidate_scheme not in {"https", "http", "file"}:
            raise UpdateError(f"更新配置包含不支持的地址协议：{candidate_scheme or '空'}。")
        if candidate_scheme == "file" and manifest_scheme != "file":
            raise UpdateError("远程更新配置不能引用本机文件。")
    notes = _normalize_notes(notes_value)
    return UpdateInfo(
        version=version,
        file_url=file_url,
        sha256=sha256,
        notes=notes,
        mandatory=mandatory,
        manifest_url=manifest_url,
        update_mode=update_mode,
        fallback_urls=tuple(url for url in fallback_urls if url != file_url),
    )


def manifest_version(manifest: dict[str, Any], platform: str) -> str:
    platform_payload = _platform_payload(manifest, platform)
    return _manifest_version_from_payload(platform_payload, manifest)


def download_update_package(
    update: UpdateInfo,
    dest_dir: Path | None = None,
    progress_callback: Callable[[int, int], None] | None = None,
    cancel_event: Any | None = None,
) -> Path:
    if update.update_mode != "auto":
        raise UpdateError("当前平台使用手动安装包，不能交给自动更新器。")
    if dest_dir is None:
        dest_dir = Path(tempfile.mkdtemp(prefix="hr_toolkit_update_"))
    else:
        dest_dir.mkdir(parents=True, exist_ok=True)

    filename = Path(urllib.parse.urlparse(update.file_url).path).name or f"HRToolkit-{update.version}.zip"
    final_path = dest_dir / filename
    temp_path = dest_dir / f"{filename}.download"
    failures: list[str] = []
    download_urls = update.download_urls
    if not download_urls:
        raise UpdateError("更新配置没有可用的 Gitee 下载地址。")
    for download_url in download_urls:
        if cancel_event is not None and cancel_event.is_set():
            temp_path.unlink(missing_ok=True)
            final_path.unlink(missing_ok=True)
            raise UpdateCancelledError("用户已取消更新包下载。")
        request = urllib.request.Request(download_url, headers={"User-Agent": USER_AGENT})
        temp_path.unlink(missing_ok=True)
        final_path.unlink(missing_ok=True)
        try:
            with _open_url(request, timeout=60) as response:
                total = _response_content_length(response)
                if total > UPDATE_PACKAGE_MAX_BYTES:
                    raise UpdateError("更新包超过允许的最大体积，已停止下载。")
                _ensure_disk_space(dest_dir, total)
                downloaded = 0
                with temp_path.open("wb") as output:
                    while True:
                        if cancel_event is not None and cancel_event.is_set():
                            raise UpdateCancelledError("用户已取消更新包下载。")
                        chunk = response.read(1024 * 256)
                        if not chunk:
                            break
                        output.write(chunk)
                        downloaded += len(chunk)
                        if downloaded > UPDATE_PACKAGE_MAX_BYTES:
                            raise UpdateError("更新包超过允许的最大体积，已停止下载。")
                        if cancel_event is not None and cancel_event.is_set():
                            raise UpdateCancelledError("用户已取消更新包下载。")
                        if progress_callback is not None:
                            progress_callback(downloaded, total)
            if cancel_event is not None and cancel_event.is_set():
                temp_path.unlink(missing_ok=True)
                final_path.unlink(missing_ok=True)
                raise UpdateCancelledError("用户已取消更新包下载。")
            os.replace(temp_path, final_path)
            actual_sha256 = sha256_file(final_path)
            if actual_sha256.lower() != update.sha256.lower():
                raise UpdateError("更新包 SHA256 校验失败。")
            return final_path
        except UpdateCancelledError:
            temp_path.unlink(missing_ok=True)
            final_path.unlink(missing_ok=True)
            raise
        except Exception as exc:
            temp_path.unlink(missing_ok=True)
            final_path.unlink(missing_ok=True)
            failures.append(f"{_update_source_name(download_url)}：{exc}")
    raise UpdateError("更新包下载失败，已按顺序尝试：" + "；".join(failures))


def resolve_download_url(update: UpdateInfo, timeout: int = 10) -> str:
    """Return the first reachable package URL without downloading its body."""
    failures: list[str] = []
    download_urls = update.download_urls
    if not download_urls:
        raise UpdateError("更新配置没有可用的 Gitee 下载地址。")
    for download_url in download_urls:
        request = urllib.request.Request(
            download_url,
            headers={"User-Agent": USER_AGENT, "Range": "bytes=0-0"},
            method="GET",
        )
        try:
            with _open_url(request, timeout=timeout):
                return download_url
        except Exception as exc:
            failures.append(f"{_update_source_name(download_url)}：{exc}")
    raise UpdateError("更新包地址均不可用，已按顺序尝试：" + "；".join(failures))


def launch_update_replacement(
    package_path: Path,
    app_dir: Path | None = None,
    launcher_path: Path | None = None,
    wait_pid: int | None = None,
) -> None:
    app_dir = app_dir or current_app_dir()
    launcher_path = launcher_path or current_launcher_path()
    temp_dir = Path(tempfile.mkdtemp(prefix="hr_toolkit_updater_"))
    log_file = app_dir.parent / UPDATE_LOG_FILE
    _append_update_log(log_file, "准备启动更新程序。")
    _append_update_log(log_file, f"程序目录：{app_dir}")
    _append_update_log(log_file, f"更新包：{package_path}")

    temp_updater = _extract_package_updater(package_path, temp_dir)
    if temp_updater is None:
        updater_path = find_updater_executable(app_dir)
        temp_updater = temp_dir / updater_path.name
        shutil.copy2(updater_path, temp_updater)
        _append_update_log(log_file, f"未在更新包中找到更新程序，使用当前目录更新程序：{updater_path}")
    else:
        _append_update_log(log_file, f"使用更新包中的更新程序：{temp_updater}")

    if _stage_win7_updater_app_local_runtimes(app_dir, temp_updater.parent):
        _append_update_log(log_file, "已为 Win7 临时更新程序复制 app-local UCRT/VC 运行库。")

    if not sys.platform.startswith("win"):
        temp_updater.chmod(temp_updater.stat().st_mode | 0o111)

    pkg_flag = "--installer" if package_path.suffix.lower() == ".exe" else "--zip"
    args = [
        str(temp_updater),
        pkg_flag,
        str(package_path),
        "--app-dir",
        str(app_dir),
        "--launcher",
        launcher_path.name,
        "--wait-pid",
        str(wait_pid or os.getpid()),
        "--log-file",
        str(log_file),
        "--relaunch",
        "--ui",
    ]
    _append_update_log(log_file, "更新程序参数：" + " ".join(args[1:]))
    subprocess.Popen(args, cwd=str(app_dir.parent), close_fds=True)


def _stage_win7_updater_app_local_runtimes(app_dir: Path, target_dir: Path) -> bool:
    if not sys.platform.startswith("win") or _windows_platform_key() != "windows-x64-win7":
        return False
    for name in WIN7_UPDATER_APP_LOCAL_RUNTIME_FILES:
        source = app_dir / name
        if source.is_symlink() or not source.is_file():
            raise UpdateError(f"Win7 更新运行库缺失或不是普通文件：{name}")
        shutil.copy2(source, target_dir / name)
    return True


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


def _extract_package_updater(package_path: Path, temp_dir: Path) -> Path | None:
    if package_path.suffix.lower() == ".exe":
        return None
    names = ["HRToolkitUpdater.exe"] if sys.platform.startswith("win") else ["HRToolkitUpdater"]
    try:
        with zipfile.ZipFile(package_path) as archive:
            for member in archive.infolist():
                if member.is_dir():
                    continue
                member_name = Path(member.filename).name
                if member_name not in names:
                    continue
                unix_mode = (member.external_attr >> 16) & 0xFFFF
                if stat.S_IFMT(unix_mode) not in (0, stat.S_IFREG):
                    raise UpdateError("更新包中的独立更新程序不是普通文件，已拒绝执行。")
                if member.file_size > UPDATE_UPDATER_MAX_BYTES:
                    raise UpdateError("更新包中的独立更新程序体积异常，已拒绝执行。")
                if _zip_compression_ratio(member) > UPDATE_ZIP_MAX_COMPRESSION_RATIO:
                    raise UpdateError("更新包中的独立更新程序压缩比异常，已拒绝执行。")
                target = temp_dir / member_name
                with archive.open(member) as source, target.open("wb") as output:
                    copied = 0
                    while True:
                        chunk = source.read(1024 * 1024)
                        if not chunk:
                            break
                        copied += len(chunk)
                        if copied > UPDATE_UPDATER_MAX_BYTES:
                            raise UpdateError("更新包中的独立更新程序体积异常，已拒绝执行。")
                        output.write(chunk)
                if copied != member.file_size:
                    target.unlink(missing_ok=True)
                    raise UpdateError("更新包中的独立更新程序数据不完整。")
                return target
    except zipfile.BadZipFile as exc:
        raise UpdateError("更新包不是有效的 zip 文件。") from exc
    return None


def _append_update_log(log_file: Path, text: str) -> None:
    try:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        trim_log_file(log_file)
        with log_file.open("a", encoding="utf-8") as handle:
            handle.write(text + "\n")
    except OSError:
        pass


def trim_log_file(log_file: Path, max_bytes: int = UPDATE_LOG_MAX_BYTES, keep_bytes: int = UPDATE_LOG_KEEP_BYTES) -> None:
    """日志超限时只保留末尾内容，避免更新日志无限增长。"""
    try:
        if not log_file.exists():
            return
        size = log_file.stat().st_size
        if size <= max_bytes:
            return
        with log_file.open("rb") as source:
            source.seek(max(size - keep_bytes, 0))
            data = source.read()
        newline = data.find(b"\n")
        if newline >= 0:
            data = data[newline + 1 :]
        log_file.write_bytes(b"(...earlier log trimmed...)\n" + data)
    except OSError:
        pass


def _response_content_length(response: Any) -> int:
    headers = getattr(response, "headers", None)
    if headers is None:
        return 0
    try:
        value = int(headers.get("Content-Length") or 0)
    except (TypeError, ValueError):
        return 0
    return max(value, 0)


def _zip_compression_ratio(member: zipfile.ZipInfo) -> float:
    if member.file_size <= 0:
        return 0.0
    if member.compress_size <= 0:
        return float("inf")
    return member.file_size / member.compress_size


def cleanup_stale_update_files(max_age_days: float = 3, temp_dir: Path | None = None) -> int:
    """清理历史更新遗留的临时目录（下载包、解压目录、更新程序副本）。

    只清理超过 max_age_days 的条目，避免碰到正在进行的更新。返回清理数量。
    """
    root = Path(temp_dir) if temp_dir is not None else Path(tempfile.gettempdir())
    cutoff = time.time() - max_age_days * 86400
    removed = 0
    try:
        entries = list(root.iterdir())
    except OSError:
        return 0
    for entry in entries:
        if not entry.name.startswith(UPDATE_TEMP_PREFIXES):
            continue
        try:
            if entry.stat().st_mtime > cutoff:
                continue
        except OSError:
            continue
        if entry.is_dir():
            shutil.rmtree(entry, ignore_errors=True)
        else:
            entry.unlink(missing_ok=True)
        if not entry.exists():
            removed += 1
    return removed


def _ensure_disk_space(dest_dir: Path, download_size: int) -> None:
    if download_size <= 0:
        return
    try:
        free = shutil.disk_usage(dest_dir).free
    except OSError:
        return
    required = download_size * DISK_SPACE_FACTOR
    if free < required:
        raise UpdateError(
            f"磁盘空间不足：安装更新约需 {required / 1024 / 1024:.0f} MB 可用空间，"
            f"当前仅剩 {free / 1024 / 1024:.0f} MB。请清理磁盘后重试。"
        )


def current_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd().resolve()


def current_launcher_path() -> Path:
    return Path(sys.executable).resolve()


def _read_update_url_file() -> str | None:
    urls = _read_update_url_files()
    return urls[0] if urls else None


def _read_update_url_files() -> tuple[str, ...]:
    for parent in _update_url_search_dirs():
        path = parent / UPDATE_URL_FILE
        if not path.exists():
            continue
        urls = _normalize_url_lines(path.read_text(encoding="utf-8-sig"))
        if urls:
            return urls
    return ()


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
        return _windows_platform_key()
    if sys.platform == "darwin":
        return "macos"
    return "linux"


def _windows_platform_key() -> str:
    try:
        version = sys.getwindowsversion()
        platform_version = getattr(version, "platform_version", ())
        if len(platform_version) >= 2:
            major_minor = (int(platform_version[0]), int(platform_version[1]))
        else:
            major_minor = (int(version.major), int(version.minor))
    except (AttributeError, IndexError, TypeError, ValueError):
        # 非 Windows 单元测试或异常运行时保守选择现代通道，不把现代机器
        # 错误降级到已经冻结依赖的 Win7 构建。
        return "windows-x64-modern"
    # Python 3.12 requires Windows 8.1 or newer. Keep Windows 7 SP1 and
    # Windows 8/Server 2012 on the frozen Python 3.8 compatibility lane.
    if (6, 1) <= major_minor < (6, 3):
        return "windows-x64-win7"
    return "windows-x64-modern"


def _platform_payload(manifest: dict[str, Any], platform: str) -> dict[str, Any]:
    platforms = manifest.get("platforms")
    if not isinstance(platforms, dict):
        return manifest
    if platform == "macos":
        machine = platform_module.machine().strip().lower()
        if machine in {"arm64", "aarch64"}:
            architecture_aliases = ("macos-arm64",)
        elif machine in {"x86_64", "amd64"}:
            architecture_aliases = ("macos-x64", "macos-x86_64")
        else:
            architecture_aliases = ()
        aliases = architecture_aliases + ("macos", "darwin", "mac")
    elif platform == "windows-x64-win7":
        # Win7 兼容包绝不能退回通用 windows 项，否则会重新下载 Python 3.12
        # 现代包并在下次启动时触发 KERNEL32.dll 入口点错误。
        aliases = ("windows-x64-win7", "windows-win7")
    elif platform == "windows-x64-modern":
        aliases = ("windows-x64-modern", "windows", "win", "win32")
    else:
        aliases = {
            "windows": ("windows", "windows-x64-modern", "win", "win32"),
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


def _normalize_manifest_urls(value: Any, manifest_url: str) -> tuple[str, ...]:
    if isinstance(value, str):
        values: Iterable[Any] = (value,)
    elif isinstance(value, (list, tuple)):
        values = value
    else:
        values = ()
    urls = (
        urllib.parse.urljoin(manifest_url, str(item).strip())
        for item in values
        if str(item).strip()
    )
    return _dedupe_urls(urls)


def _normalize_url_lines(value: str) -> tuple[str, ...]:
    return _dedupe_urls(
        line.strip()
        for line in value.splitlines()
        if line.strip() and not line.strip().startswith("#")
    )


def _dedupe_urls(urls: Iterable[str]) -> tuple[str, ...]:
    unique: list[str] = []
    for url in urls:
        normalized = str(url).strip()
        if normalized and normalized not in unique:
            unique.append(normalized)
    return tuple(unique)


def _is_github_url(url: str) -> bool:
    host = (urllib.parse.urlparse(str(url)).hostname or "").casefold()
    return any(
        host == domain or host.endswith("." + domain)
        for domain in ("github.com", "githubusercontent.com", "githubassets.com")
    )


def _is_release_discovery_url(url: str) -> bool:
    parsed = urllib.parse.urlparse(url)
    return parsed.path.rstrip("/").endswith("/releases/latest")


def _find_manifest_asset_url(release: dict[str, Any]) -> str | None:
    assets = release.get("assets") or release.get("attach_files")
    if not isinstance(assets, list):
        return None
    for asset in assets:
        if not isinstance(asset, dict) or str(asset.get("name", "")).strip() != "latest.json":
            continue
        url = str(asset.get("browser_download_url") or asset.get("url") or "").strip()
        if url:
            return url
    return None


def _update_source_name(url: str) -> str:
    host = (urllib.parse.urlparse(url).hostname or "").lower()
    if host == "gitee.com" or host.endswith(".gitee.com"):
        return "Gitee 国内源"
    if host == "github.com" or host.endswith(".github.com") or host.endswith("githubusercontent.com"):
        return "GitHub 备用源"
    return url


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
