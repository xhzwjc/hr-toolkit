from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from .excel_compat import is_supported_excel_file


_ZIP_UTF8_FLAG = 0x0800


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
    extract_dir.mkdir(parents=True, exist_ok=True)
    extract_root = extract_dir.resolve()
    try:
        with zipfile.ZipFile(zip_path) as archive:
            for member in archive.infolist():
                target = (extract_dir / zip_member_name(member)).resolve()
                if not target.is_relative_to(extract_root):
                    warnings.append(f"{zip_path.name} 中存在不安全路径，已跳过：{member.filename}")
                    continue
                if member.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, target.open("wb") as output:
                    shutil.copyfileobj(source, output)
    except Exception as exc:
        warnings.append(f"{zip_path.name} 解压失败，已跳过：{exc}")
        return []
    return sorted(path for path in extract_dir.rglob("*") if path.is_file() and is_supported_excel_file(path))
