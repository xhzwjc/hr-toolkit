from __future__ import annotations

import stat
import shutil
import unicodedata
import zipfile
from pathlib import Path

from .excel_compat import is_supported_excel_file


_ZIP_UTF8_FLAG = 0x0800
ZIP_MAX_MEMBERS = 10_000
ZIP_MAX_MEMBER_BYTES = 512 * 1024 * 1024
ZIP_MAX_TOTAL_BYTES = 2 * 1024 * 1024 * 1024
ZIP_MAX_COMPRESSION_RATIO = 200
ZIP_RATIO_CHECK_MIN_BYTES = 1024 * 1024
ZIP_COPY_BUFFER_BYTES = 1024 * 1024
ZIP_MIN_FREE_SPACE_BYTES = 128 * 1024 * 1024


def normalize_input_paths(input_path: str | Path | list[str | Path], empty_message: str) -> list[Path]:
    raw_paths = input_path if isinstance(input_path, list) else [input_path]
    paths = [Path(path).expanduser().resolve() for path in raw_paths]
    if not paths:
        raise ValueError(empty_message)
    return paths


def zip_member_name(member: zipfile.ZipInfo) -> str:
    """Windows 压缩工具常以 GBK 存储中文文件名且不设 UTF-8 标志，
    zipfile 会按 cp437 解码成乱码；此处按 GBK 还原。"""
    name = member.filename
    if member.flag_bits & _ZIP_UTF8_FLAG:
        return name
    try:
        return name.encode("cp437").decode("gbk")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return name


def extract_zip_excel_files(
    zip_path: Path,
    temp_dir: Path,
    warnings: list[str],
    *,
    subdir: str | None = None,
) -> list[Path]:
    extract_dir = temp_dir / f"zip_{len(list(temp_dir.glob('zip_*'))) + 1}"
    if subdir:
        extract_dir = extract_dir / subdir
    try:
        with zipfile.ZipFile(zip_path) as archive:
            members = archive.infolist()
            declared_total_bytes = _validate_zip_limits(zip_path, members)
            _validate_zip_free_space(temp_dir, declared_total_bytes)
            extract_dir.mkdir(parents=True, exist_ok=True)
            extract_root = extract_dir.resolve()
            extracted_bytes = 0
            for member in members:
                member_parts = _safe_zip_member_parts(member)
                if member_parts is None:
                    warnings.append(f"{zip_path.name} 中存在不安全路径，已跳过：{member.filename}")
                    continue
                target = extract_dir.joinpath(*member_parts).resolve()
                if not target.is_relative_to(extract_root):
                    warnings.append(f"{zip_path.name} 中存在不安全路径，已跳过：{member.filename}")
                    continue
                if _zip_member_is_symlink(member):
                    warnings.append(f"{zip_path.name} 中存在链接文件，已跳过：{member.filename}")
                    continue
                if member.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                partial_target = target.with_name(target.name + ".partial")
                member_bytes = 0
                with archive.open(member) as source, partial_target.open("wb") as output:
                    while True:
                        chunk = source.read(ZIP_COPY_BUFFER_BYTES)
                        if not chunk:
                            break
                        member_bytes += len(chunk)
                        extracted_bytes += len(chunk)
                        if member_bytes > ZIP_MAX_MEMBER_BYTES or extracted_bytes > ZIP_MAX_TOTAL_BYTES:
                            raise ValueError("解压后的文件大小超过安全上限")
                        output.write(chunk)
                partial_target.replace(target)
    except Exception as exc:
        shutil.rmtree(extract_dir, ignore_errors=True)
        warnings.append(f"{zip_path.name} 解压失败，已跳过：{exc}")
        return []
    return sorted(path for path in extract_dir.rglob("*") if path.is_file() and is_supported_excel_file(path))


def _validate_zip_limits(zip_path: Path, members: list[zipfile.ZipInfo]) -> int:
    if len(members) > ZIP_MAX_MEMBERS:
        raise ValueError(f"压缩包文件数量超过 {ZIP_MAX_MEMBERS} 个安全上限")

    total_bytes = 0
    seen_names: set[str] = set()
    for member in members:
        member_parts = _safe_zip_member_parts(member)
        if member_parts is not None:
            normalized_name = "/".join(member_parts)
            canonical_name = normalized_name.casefold()
            if canonical_name in seen_names:
                raise ValueError(f"压缩包包含重复路径：{normalized_name}")
            seen_names.add(canonical_name)
        if member.is_dir():
            continue
        if member.file_size < 0 or member.compress_size < 0:
            raise ValueError("压缩包包含无效的文件大小")
        if member.file_size > ZIP_MAX_MEMBER_BYTES:
            raise ValueError(f"{zip_member_name(member)} 解压后超过单文件安全上限")
        total_bytes += member.file_size
        if total_bytes > ZIP_MAX_TOTAL_BYTES:
            raise ValueError("压缩包解压后的总大小超过安全上限")
        if member.file_size >= ZIP_RATIO_CHECK_MIN_BYTES:
            ratio = member.file_size / max(member.compress_size, 1)
            if ratio > ZIP_MAX_COMPRESSION_RATIO:
                raise ValueError(f"{zip_member_name(member)} 的压缩比例异常")
    return total_bytes


def _validate_zip_free_space(temp_dir: Path, required_bytes: int) -> None:
    available = shutil.disk_usage(temp_dir).free
    if available < required_bytes + ZIP_MIN_FREE_SPACE_BYTES:
        raise ValueError("临时磁盘空间不足，无法安全解压压缩包")


def _zip_member_is_symlink(member: zipfile.ZipInfo) -> bool:
    unix_mode = (member.external_attr >> 16) & 0xFFFF
    return stat.S_ISLNK(unix_mode)


def _safe_zip_member_parts(member: zipfile.ZipInfo) -> tuple[str, ...] | None:
    normalized = unicodedata.normalize("NFC", zip_member_name(member).replace("\\", "/"))
    if normalized.startswith("/"):
        return None
    raw_parts = normalized.split("/")
    if any(part == ".." for part in raw_parts):
        return None
    parts = tuple(part for part in raw_parts if part not in {"", "."})
    if parts and len(parts[0]) >= 2 and parts[0][1] == ":":
        return None
    return parts
