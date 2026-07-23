from __future__ import annotations

import argparse
import json
import mimetypes
import os
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path
from typing import Any, Mapping, Sequence

from generate_release_metadata import (
    REPOSITORY_PATTERN,
    detect_mac_variant,
    release_asset_names,
    sha256_file,
    validate_release_identity,
    validate_version,
)


DEFAULT_REPOSITORY = "optimistic-little-sunspot/hr-toolkit"
DEFAULT_API_BASE = "https://gitee.com/api/v5"
DEFAULT_GITHUB_REPOSITORY = "xhzwjc/hr-toolkit"
METADATA_NAMES = ("latest.json", "SHA256SUMS.txt")
USER_AGENT = "HRToolkit-Gitee-Publisher/1.0"


class GiteeReleaseError(RuntimeError):
    """Raised when a mirrored release cannot be created or verified."""


class GiteeClient:
    def __init__(self, token: str, *, api_base: str = DEFAULT_API_BASE, timeout: int = 120) -> None:
        if not token.strip():
            raise GiteeReleaseError("GITEE_TOKEN 不能为空。")
        self._token = token.strip()
        self.api_base = api_base.rstrip("/")
        self.timeout = timeout
        self._ssl_context = ssl.create_default_context()

    def get_release_by_tag(self, repository: str, tag: str) -> dict[str, Any] | None:
        path = f"/repos/{_quoted_repository(repository)}/releases/tags/{urllib.parse.quote(tag, safe='')}"
        result = self._request_json("GET", path, allow_not_found=True)
        if result is None:
            return None
        return _require_object(result, "按 Tag 查询 Release")

    def create_release(
        self,
        repository: str,
        *,
        tag: str,
        target_commitish: str,
        name: str,
        body: str,
    ) -> dict[str, Any]:
        result = self._request_json(
            "POST",
            f"/repos/{_quoted_repository(repository)}/releases",
            fields={
                "tag_name": tag,
                "target_commitish": target_commitish,
                "name": name,
                "body": body,
                "prerelease": "false",
            },
        )
        return _require_object(result, "创建 Release")

    def update_release(
        self,
        repository: str,
        release_id: str,
        *,
        tag: str,
        name: str,
        body: str,
    ) -> dict[str, Any]:
        result = self._request_json(
            "PATCH",
            f"/repos/{_quoted_repository(repository)}/releases/{urllib.parse.quote(release_id, safe='')}",
            fields={
                "tag_name": tag,
                "name": name,
                "body": body,
                "prerelease": "false",
            },
        )
        return _require_object(result, "更新 Release")

    def list_attachments(self, repository: str, release_id: str) -> list[dict[str, Any]]:
        result = self._request_json(
            "GET",
            f"/repos/{_quoted_repository(repository)}/releases/{urllib.parse.quote(release_id, safe='')}/attach_files",
            query={"page": "1", "per_page": "100"},
        )
        if isinstance(result, dict) and isinstance(result.get("data"), list):
            result = result["data"]
        if not isinstance(result, list) or not all(isinstance(item, dict) for item in result):
            raise GiteeReleaseError("Gitee 附件列表格式不正确。")
        return result

    def delete_attachment(self, repository: str, release_id: str, attachment_id: str) -> None:
        self._request_json(
            "DELETE",
            f"/repos/{_quoted_repository(repository)}/releases/{urllib.parse.quote(release_id, safe='')}"
            f"/attach_files/{urllib.parse.quote(attachment_id, safe='')}",
            expect_json=False,
        )

    def upload_attachment(
        self,
        repository: str,
        release_id: str,
        file_path: Path,
    ) -> dict[str, Any]:
        result = self._request_json(
            "POST",
            f"/repos/{_quoted_repository(repository)}/releases/{urllib.parse.quote(release_id, safe='')}/attach_files",
            file_path=file_path,
        )
        return _require_object(result, f"上传 {file_path.name}")

    def get_public_latest_release(self, repository: str) -> dict[str, Any]:
        result = self._request_json(
            "GET",
            f"/repos/{_quoted_repository(repository)}/releases/latest",
            authenticated=False,
        )
        return _require_object(result, "读取公开最新 Release")

    def get_public_json(self, url: str) -> dict[str, Any]:
        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        try:
            with urllib.request.urlopen(
                request,
                timeout=self.timeout,
                context=self._ssl_context,
            ) as response:
                payload = response.read().decode("utf-8-sig")
        except Exception as exc:
            raise GiteeReleaseError(f"无法读取公开附件：{_safe_exception(exc, self._token)}") from None
        try:
            data = json.loads(payload)
        except json.JSONDecodeError as exc:
            raise GiteeReleaseError(f"公开附件不是有效 JSON：{exc}") from None
        return _require_object(data, "读取公开 JSON 附件")

    def _request_json(
        self,
        method: str,
        path: str,
        *,
        fields: Mapping[str, str] | None = None,
        query: Mapping[str, str] | None = None,
        file_path: Path | None = None,
        authenticated: bool = True,
        allow_not_found: bool = False,
        expect_json: bool = True,
    ) -> Any:
        query_values = dict(query or {})
        headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
        data: bytes | None = None
        if authenticated and method in {"GET", "DELETE"}:
            query_values["access_token"] = self._token
        if file_path is not None:
            if not file_path.is_file():
                raise GiteeReleaseError(f"上传文件不存在：{file_path}")
            multipart_fields = dict(fields or {})
            if authenticated:
                multipart_fields["access_token"] = self._token
            data, content_type = _multipart_body(multipart_fields, file_path)
            headers["Content-Type"] = content_type
        elif fields is not None or (authenticated and method in {"POST", "PATCH"}):
            form_fields = dict(fields or {})
            if authenticated:
                form_fields["access_token"] = self._token
            data = urllib.parse.urlencode(form_fields).encode("utf-8")
            headers["Content-Type"] = "application/x-www-form-urlencoded"

        url = self.api_base + path
        if query_values:
            url += "?" + urllib.parse.urlencode(query_values)
        request = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(
                request,
                timeout=self.timeout,
                context=self._ssl_context,
            ) as response:
                payload = response.read()
        except urllib.error.HTTPError as exc:
            if allow_not_found and exc.code == 404:
                return None
            body = exc.read().decode("utf-8", errors="replace")[:1000]
            detail = _redact(body, self._token)
            raise GiteeReleaseError(
                f"Gitee API {method} {path} 返回 HTTP {exc.code}：{detail}"
            ) from None
        except Exception as exc:
            raise GiteeReleaseError(
                f"Gitee API {method} {path} 请求失败：{_safe_exception(exc, self._token)}"
            ) from None

        if not expect_json or not payload.strip():
            return None
        try:
            return json.loads(payload.decode("utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise GiteeReleaseError(f"Gitee API {method} {path} 返回无效 JSON：{exc}") from None


def expected_asset_names(assets_dir: Path, version: str) -> tuple[str, ...]:
    mac_variant = detect_mac_variant(assets_dir, version)
    return release_asset_names(version, mac_variant=mac_variant) + METADATA_NAMES


def validate_mirror_assets(
    assets_dir: Path,
    *,
    version: str,
    tag: str,
    repository: str,
    github_repository: str = DEFAULT_GITHUB_REPOSITORY,
) -> tuple[str, ...]:
    names = expected_asset_names(assets_dir, version)
    actual = {path.name for path in assets_dir.iterdir() if path.is_file()} if assets_dir.is_dir() else set()
    if actual != set(names):
        raise GiteeReleaseError(
            f"Gitee 镜像资产必须严格匹配白名单：期望={sorted(names)}，实际={sorted(actual)}"
        )
    for name in names:
        path = assets_dir / name
        if not path.is_file() or path.stat().st_size <= 0:
            raise GiteeReleaseError(f"Gitee 镜像资产为空或不存在：{path}")

    latest = _read_json_file(assets_dir / "latest.json")
    if latest.get("version") != version:
        raise GiteeReleaseError("latest.json 版本与发布版本不一致。")
    expected_release_url = f"https://gitee.com/{repository}/releases/tag/{tag}"
    if latest.get("release_url") != expected_release_url:
        raise GiteeReleaseError("latest.json 的 Release 页面不是 Gitee。")

    platforms = latest.get("platforms")
    if not isinstance(platforms, dict) or not platforms:
        raise GiteeReleaseError("latest.json 缺少 platforms。")
    gitee_prefix = f"https://gitee.com/{repository}/releases/download/{tag}/"
    github_prefix = f"https://github.com/{github_repository}/releases/download/{tag}/"
    binary_names = set(names) - set(METADATA_NAMES)
    for key, payload in platforms.items():
        if not isinstance(payload, dict):
            raise GiteeReleaseError(f"latest.json 平台 {key} 格式不正确。")
        file_url = str(payload.get("file_url") or "")
        fallback_urls = payload.get("fallback_urls")
        if not file_url.startswith(gitee_prefix):
            raise GiteeReleaseError(f"latest.json 平台 {key} 未优先使用 Gitee。")
        filename = urllib.parse.unquote(urllib.parse.urlparse(file_url).path.rsplit("/", 1)[-1])
        if filename not in binary_names:
            raise GiteeReleaseError(f"latest.json 平台 {key} 指向非白名单资产：{filename}")
        if fallback_urls != [github_prefix + filename]:
            raise GiteeReleaseError(f"latest.json 平台 {key} 未配置 GitHub 备用地址。")
        if str(payload.get("sha256") or "").lower() != sha256_file(assets_dir / filename):
            raise GiteeReleaseError(f"latest.json 平台 {key} 的 SHA256 不正确。")

    _verify_checksum_file(assets_dir, names)
    return names


def publish_gitee_release(
    client: GiteeClient,
    *,
    assets_dir: Path,
    version: str,
    tag: str,
    repository: str,
    target_commitish: str,
    name: str,
    body: str,
) -> tuple[dict[str, Any], tuple[str, ...]]:
    names = validate_mirror_assets(
        assets_dir,
        version=version,
        tag=tag,
        repository=repository,
    )
    release = client.get_release_by_tag(repository, tag)
    if release is None:
        release = client.create_release(
            repository,
            tag=tag,
            target_commitish=target_commitish,
            name=name,
            body=body,
        )
    else:
        release_id = _release_id(release)
        release = client.update_release(
            repository,
            release_id,
            tag=tag,
            name=name,
            body=body,
        )

    release_id = _release_id(release)
    for attachment in client.list_attachments(repository, release_id):
        client.delete_attachment(repository, release_id, _attachment_id(attachment))

    upload_order = sorted(name for name in names if name != "latest.json") + ["latest.json"]
    for asset_name in upload_order:
        client.upload_attachment(repository, release_id, assets_dir / asset_name)

    attachments = client.list_attachments(repository, release_id)
    _verify_attachment_list(attachments, assets_dir, names)
    return release, names


def verify_public_release(
    client: GiteeClient,
    *,
    assets_dir: Path,
    repository: str,
    tag: str,
    names: Sequence[str],
    attempts: int = 5,
    retry_delay: float = 2.0,
) -> None:
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            release = client.get_public_latest_release(repository)
            if str(release.get("tag_name") or "") != tag:
                raise GiteeReleaseError("Gitee 公开 latest Release 尚未指向当前 Tag。")
            assets = release.get("assets") or release.get("attach_files")
            if not isinstance(assets, list):
                raise GiteeReleaseError("Gitee 公开 Release 缺少附件列表。")
            public_assets = {
                str(item.get("name") or "").strip(): item
                for item in assets
                if isinstance(item, dict) and str(item.get("name") or "").strip()
            }
            missing = set(names) - set(public_assets)
            if missing:
                raise GiteeReleaseError(
                    "Gitee 公开 Release 缺少附件：" + ", ".join(sorted(missing))
                )
            for asset_name in names:
                if not str(public_assets[asset_name].get("browser_download_url") or "").strip():
                    raise GiteeReleaseError(f"Gitee 公开附件缺少下载地址：{asset_name}")
            manifest_asset = next(
                (item for item in assets if isinstance(item, dict) and item.get("name") == "latest.json"),
                None,
            )
            if manifest_asset is None:
                raise GiteeReleaseError("Gitee 公开 Release 缺少 latest.json。")
            public_url = str(manifest_asset.get("browser_download_url") or "").strip()
            if not public_url:
                raise GiteeReleaseError("Gitee latest.json 缺少公开下载地址。")
            expected = _read_json_file(assets_dir / "latest.json")
            if client.get_public_json(public_url) != expected:
                raise GiteeReleaseError("Gitee 公开 latest.json 与本地生成内容不一致。")
            return
        except Exception as exc:
            last_error = exc
            if attempt + 1 < attempts:
                time.sleep(retry_delay)
    raise GiteeReleaseError(f"Gitee 公开 Release 验证失败：{last_error}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="幂等发布并验证 Gitee Release 镜像")
    parser.add_argument("--version", required=True)
    parser.add_argument("--tag", help="默认 v<version>")
    parser.add_argument("--repository", default=DEFAULT_REPOSITORY)
    parser.add_argument("--target-commitish", required=True)
    parser.add_argument("--assets-dir", type=Path, required=True)
    parser.add_argument("--api-base", default=DEFAULT_API_BASE)
    parser.add_argument("--token-env", default="GITEE_TOKEN")
    parser.add_argument("--name")
    parser.add_argument("--body")
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    version = validate_version(args.version)
    tag = args.tag or f"v{version}"
    validate_release_identity(version, tag, version)
    if not REPOSITORY_PATTERN.fullmatch(args.repository):
        raise GiteeReleaseError(f"Gitee 仓库名必须是 owner/repo：{args.repository!r}")
    names = validate_mirror_assets(
        args.assets_dir,
        version=version,
        tag=tag,
        repository=args.repository,
    )
    if args.dry_run:
        print(f"Gitee 镜像 dry-run 通过：{tag}")
        for asset_name in sorted(names):
            print(f"- {asset_name}")
        return 0

    token = os.environ.get(args.token_env, "")
    if not token.strip():
        raise GiteeReleaseError(f"缺少 GitHub Actions Secret：{args.token_env}")
    client = GiteeClient(token, api_base=args.api_base)
    release_name = args.name or f"HR Toolkit {tag}"
    body = args.body or "GitHub Release 构建完成后自动同步的国内下载镜像。"
    _release, names = publish_gitee_release(
        client,
        assets_dir=args.assets_dir,
        version=version,
        tag=tag,
        repository=args.repository,
        target_commitish=args.target_commitish,
        name=release_name,
        body=body,
    )
    verify_public_release(
        client,
        assets_dir=args.assets_dir,
        repository=args.repository,
        tag=tag,
        names=names,
    )
    print(f"Gitee Release 镜像发布并验证成功：https://gitee.com/{args.repository}/releases/tag/{tag}")
    return 0


def _quoted_repository(repository: str) -> str:
    owner, name = repository.split("/", 1)
    return f"{urllib.parse.quote(owner, safe='')}/{urllib.parse.quote(name, safe='')}"


def _multipart_body(fields: Mapping[str, str], file_path: Path) -> tuple[bytes, str]:
    boundary = "----HRToolkit" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for key, value in fields.items():
        chunks.extend(
            (
                f"--{boundary}\r\n".encode("ascii"),
                f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode("utf-8"),
                str(value).encode("utf-8"),
                b"\r\n",
            )
        )
    content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    chunks.extend(
        (
            f"--{boundary}\r\n".encode("ascii"),
            f'Content-Disposition: form-data; name="file"; filename="{file_path.name}"\r\n'.encode("utf-8"),
            f"Content-Type: {content_type}\r\n\r\n".encode("ascii"),
            file_path.read_bytes(),
            b"\r\n",
            f"--{boundary}--\r\n".encode("ascii"),
        )
    )
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def _require_object(value: Any, operation: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise GiteeReleaseError(f"{operation}返回格式不正确。")
    return value


def _release_id(release: Mapping[str, Any]) -> str:
    value = str(release.get("id") or "").strip()
    if not value:
        raise GiteeReleaseError("Gitee Release 缺少 id。")
    return value


def _attachment_id(attachment: Mapping[str, Any]) -> str:
    value = str(attachment.get("id") or "").strip()
    if not value:
        raise GiteeReleaseError("Gitee Release 附件缺少 id。")
    return value


def _attachment_size(attachment: Mapping[str, Any]) -> int:
    try:
        return int(attachment.get("size"))
    except (TypeError, ValueError):
        raise GiteeReleaseError(f"Gitee Release 附件大小无效：{attachment.get('name')}") from None


def _verify_attachment_list(
    attachments: Sequence[Any],
    assets_dir: Path,
    names: Sequence[str],
) -> None:
    actual: dict[str, int] = {}
    for attachment in attachments:
        if not isinstance(attachment, dict):
            raise GiteeReleaseError("Gitee Release 附件格式不正确。")
        name = str(attachment.get("name") or "").strip()
        if not name or name in actual:
            raise GiteeReleaseError("Gitee Release 附件名称为空或重复。")
        actual[name] = _attachment_size(attachment)
    expected = {name: (assets_dir / name).stat().st_size for name in names}
    if actual != expected:
        raise GiteeReleaseError(f"Gitee Release 附件不匹配：期望={expected}，实际={actual}")


def _read_json_file(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise GiteeReleaseError(f"无法读取 {path.name}：{exc}") from None
    return _require_object(value, f"读取 {path.name}")


def _verify_checksum_file(assets_dir: Path, names: Sequence[str]) -> None:
    checksum_path = assets_dir / "SHA256SUMS.txt"
    expected_names = sorted(set(names) - {"SHA256SUMS.txt"})
    expected = {
        name: sha256_file(assets_dir / name)
        for name in expected_names
    }
    actual: dict[str, str] = {}
    for line in checksum_path.read_text(encoding="utf-8").splitlines():
        try:
            digest, name = line.split("  ", 1)
        except ValueError:
            raise GiteeReleaseError("SHA256SUMS.txt 格式不正确。") from None
        if name in actual:
            raise GiteeReleaseError(f"SHA256SUMS.txt 包含重复资产：{name}")
        actual[name] = digest.lower()
    if actual != expected:
        raise GiteeReleaseError("SHA256SUMS.txt 内容或资产 SHA256 不正确。")


def _redact(value: str, secret: str) -> str:
    return value.replace(secret, "***") if secret else value


def _safe_exception(exc: Exception, secret: str) -> str:
    return _redact(str(exc), secret)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except GiteeReleaseError as exc:
        raise SystemExit(f"Gitee 发布失败：{exc}") from None
