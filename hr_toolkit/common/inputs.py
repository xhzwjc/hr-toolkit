from __future__ import annotations

import shutil
import stat
import tarfile
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO

from .excel_compat import is_supported_excel_file


_ZIP_UTF8_FLAG = 0x0800

# 这里列出的格式由桌面端和 CLI 共同支持。多段后缀必须放在短后缀前面，
# archive_suffix() 会按长度匹配，避免把 .tar.gz 误识别成 .gz。
SUPPORTED_ARCHIVE_SUFFIXES = (
    ".tar.bz2",
    ".tar.gz",
    ".tar.xz",
    ".tbz2",
    ".tgz",
    ".txz",
    ".7z",
    ".rar",
    ".tar",
    ".zip",
)
SUPPORTED_ARCHIVE_SUFFIX_SET = frozenset(SUPPORTED_ARCHIVE_SUFFIXES)
ARCHIVE_FILE_DIALOG_PATTERN = " ".join(f"*{suffix}" for suffix in SUPPORTED_ARCHIVE_SUFFIXES)
ARCHIVE_FORMAT_DESCRIPTION = "ZIP、RAR、7Z、TAR"

ARCHIVE_MAX_MEMBERS = 10_000
ARCHIVE_MAX_MEMBER_BYTES = 512 * 1024 * 1024
ARCHIVE_MAX_TOTAL_BYTES = 2 * 1024 * 1024 * 1024
ARCHIVE_MAX_COMPRESSION_RATIO = 200
ARCHIVE_RATIO_CHECK_MIN_BYTES = 1024 * 1024
ARCHIVE_COPY_BUFFER_BYTES = 1024 * 1024
ARCHIVE_MIN_FREE_SPACE_BYTES = 128 * 1024 * 1024


@dataclass(frozen=True)
class _ArchiveMember:
    token: object
    name: str
    file_size: int
    compress_size: int | None
    is_dir: bool = False
    is_link: bool = False
    is_regular: bool = True


@dataclass(frozen=True)
class _ArchiveMemberPlan:
    member: _ArchiveMember
    parts: tuple[str, ...]


def normalize_input_paths(input_path: str | Path | list[str | Path], empty_message: str) -> list[Path]:
    raw_paths = input_path if isinstance(input_path, list) else [input_path]
    paths = [Path(path).expanduser().resolve() for path in raw_paths]
    if not paths:
        raise ValueError(empty_message)
    return paths


def archive_suffix(path: str | Path) -> str | None:
    """返回完整压缩包后缀，例如 ``.tar.gz``，不是 Path.suffix 的 ``.gz``。"""
    name = Path(path).name.casefold()
    for suffix in SUPPORTED_ARCHIVE_SUFFIXES:
        if name.endswith(suffix):
            return suffix
    return None


def archive_stem(path: str | Path) -> str:
    name = Path(path).name
    suffix = archive_suffix(name)
    return name[: -len(suffix)] if suffix else Path(name).stem


def is_supported_archive_file(path: str | Path) -> bool:
    return archive_suffix(path) is not None


def zip_member_name(member: zipfile.ZipInfo) -> str:
    """还原 Windows 压缩工具以 GBK 写入、但未设置 UTF-8 标志的中文文件名。"""
    name = member.filename
    if member.flag_bits & _ZIP_UTF8_FLAG:
        return name
    try:
        return name.encode("cp437").decode("gbk")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return name


def extract_archive_excel_files(
    archive_path: Path,
    temp_dir: Path,
    warnings: list[str],
    *,
    subdir: str | None = None,
) -> list[Path]:
    """安全解压受支持的压缩包，并以确定顺序返回其中的 Excel 文件。

    所有格式共用路径、重复项、大小、压缩比例和磁盘空间限制，确保同一批
    文件使用 ZIP、RAR、7Z 或 TAR 打包时得到相同的文件集合与顺序。
    """
    archive_path = Path(archive_path)
    if not is_supported_archive_file(archive_path):
        warnings.append(f"{archive_path.name} 不是受支持的压缩包，已跳过")
        return []

    extract_root = _allocate_extract_root(temp_dir)
    extract_dir = extract_root
    try:
        if subdir:
            subdir_parts = _safe_member_parts(subdir)
            if not subdir_parts:
                raise ValueError("压缩包上下文目录名称无效")
            extract_dir = extract_root.joinpath(*subdir_parts)
            extract_dir.mkdir(parents=True, exist_ok=False)

        archive_format = _detect_archive_format(archive_path)
        if archive_format == "zip":
            _extract_zip(archive_path, extract_dir, temp_dir, warnings)
        elif archive_format == "rar":
            _extract_rar(archive_path, extract_dir, temp_dir, warnings)
        elif archive_format == "7z":
            _extract_7z(archive_path, extract_dir, temp_dir, warnings)
        else:
            _extract_tar(archive_path, extract_dir, temp_dir, warnings)
    except Exception as exc:
        shutil.rmtree(extract_root, ignore_errors=True)
        warnings.append(f"{archive_path.name} 解压失败，已跳过：{_friendly_archive_error(exc)}")
        return []

    files = [
        path
        for path in extract_dir.rglob("*")
        if path.is_file() and not path.is_symlink() and is_supported_excel_file(path)
    ]
    return sorted(files, key=lambda path: _relative_sort_key(path, extract_dir))


def extract_zip_excel_files(
    zip_path: Path,
    temp_dir: Path,
    warnings: list[str],
    *,
    subdir: str | None = None,
) -> list[Path]:
    """兼容旧调用方；新代码应使用 extract_archive_excel_files。"""
    return extract_archive_excel_files(zip_path, temp_dir, warnings, subdir=subdir)


def _allocate_extract_root(temp_dir: Path) -> Path:
    temp_dir.mkdir(parents=True, exist_ok=True)
    index = 1
    while True:
        candidate = temp_dir / f"archive_{index}"
        try:
            candidate.mkdir(parents=False, exist_ok=False)
            return candidate
        except FileExistsError:
            index += 1


def _detect_archive_format(archive_path: Path) -> str:
    # 以实际内容识别格式，后缀只负责决定该文件是否属于用户允许选择的压缩包。
    # 这样即使第三方软件写错后缀，也不会把 RAR 当 ZIP 解析。
    with archive_path.open("rb") as source:
        signature = source.read(8)
    if signature.startswith((b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")):
        return "zip"
    if signature.startswith((b"Rar!\x1a\x07\x00", b"Rar!\x1a\x07\x01\x00")):
        return "rar"
    if signature.startswith(b"7z\xbc\xaf\x27\x1c"):
        return "7z"

    if tarfile.is_tarfile(archive_path):
        return "tar"
    # 保留 Python zipfile 原有的自解压 ZIP/前置数据兼容；必须放在 TAR
    # 判断之后，否则 TAR 内嵌的 .xlsx 尾部会被误判为整个文件是 ZIP。
    if zipfile.is_zipfile(archive_path):
        return "zip"
    raise ValueError("文件内容不是有效的 ZIP、RAR、7Z 或 TAR 压缩包")


def _extract_zip(archive_path: Path, extract_dir: Path, temp_dir: Path, warnings: list[str]) -> None:
    with zipfile.ZipFile(archive_path) as archive:
        infos = archive.infolist()
        if any(info.flag_bits & 0x1 for info in infos):
            raise ValueError("暂不支持带密码的压缩包")
        members = [
            _ArchiveMember(
                token=info,
                name=zip_member_name(info),
                file_size=info.file_size,
                compress_size=info.compress_size,
                is_dir=info.is_dir(),
                is_link=_zip_member_is_symlink(info),
            )
            for info in infos
        ]
        plans = _validate_archive_members(
            archive_path,
            members,
            temp_dir,
            warnings,
            archive_compressed_bytes=archive_path.stat().st_size,
        )
        actual_total = [0]
        for plan in plans:
            target = _target_for_plan(extract_dir, plan)
            if plan.member.is_dir:
                target.mkdir(parents=True, exist_ok=True)
                continue
            with archive.open(plan.member.token) as source:
                _copy_member_stream(source, target, plan.member.file_size, actual_total)


def _extract_rar(archive_path: Path, extract_dir: Path, temp_dir: Path, warnings: list[str]) -> None:
    try:
        from unrar.cffi.unrarlib import FLAGS_RHDF_DIRECTORY, RarArchive
    except ImportError as exc:  # pragma: no cover - 安装/打包检查覆盖
        raise RuntimeError("RAR 解压组件未安装完整") from exc

    members: list[_ArchiveMember] = []
    with RarArchive.open_for_processing(archive_path) as archive:
        for header in archive.iterate_headers():
            file_size = header.UnpSize + (header.UnpSizeHigh << 32)
            compress_size = header.PackSize + (header.PackSizeHigh << 32)
            header_data = header.headerDataEx
            is_dir = bool(header.Flags & FLAGS_RHDF_DIRECTORY)
            is_link = bool(getattr(header_data, "RedirType", 0)) or (
                header.HostOS == 3 and stat.S_ISLNK(header_data.FileAttr)
            )
            is_regular = (
                is_dir
                or is_link
                or header.HostOS != 3
                or stat.S_ISREG(header_data.FileAttr)
            )
            members.append(
                _ArchiveMember(
                    token=header.FileNameW,
                    name=header.FileNameW,
                    file_size=file_size,
                    compress_size=compress_size,
                    is_dir=is_dir,
                    is_link=is_link,
                    is_regular=is_regular,
                )
            )
            header.skip()
    plans = _validate_archive_members(
        archive_path,
        members,
        temp_dir,
        warnings,
        archive_compressed_bytes=archive_path.stat().st_size,
    )
    plan_by_name = {
        "/".join(plan.parts).casefold(): plan
        for plan in plans
    }
    actual_total = [0]
    created_names: set[str] = set()
    for plan in plans:
        if plan.member.is_dir:
            _target_for_plan(extract_dir, plan).mkdir(parents=True, exist_ok=True)

    # 直接使用 UnRAR 的流式回调，避免高层 RarFile.open() 把每个成员完整载入内存。
    with RarArchive.open_for_processing(archive_path) as archive:
        for header in archive.iterate_headers():
            parts = _safe_member_parts(header.FileNameW)
            key = None if not parts else "/".join(parts).casefold()
            plan = None if key is None else plan_by_name.get(key)
            if plan is None or plan.member.is_dir:
                header.skip()
                continue
            _copy_rar_header(
                header,
                _target_for_plan(extract_dir, plan),
                plan.member.file_size,
                actual_total,
            )
            created_names.add(key)

    missing_names = [
        plan.member.name
        for key, plan in plan_by_name.items()
        if not plan.member.is_dir and key not in created_names
    ]
    if missing_names:
        raise ValueError(f"RAR 未解压预期成员：{missing_names[0]}")


def _extract_7z(archive_path: Path, extract_dir: Path, temp_dir: Path, warnings: list[str]) -> None:
    try:
        import py7zr
    except ImportError as exc:  # pragma: no cover - 安装/打包检查覆盖
        raise RuntimeError("7Z 解压组件未安装完整") from exc

    with py7zr.SevenZipFile(archive_path, mode="r") as archive:
        if archive.password_protected:
            raise ValueError("暂不支持带密码的压缩包")
        infos = archive.list()
        raw_infos = list(archive.files)
        if len(infos) != len(raw_infos):
            raise ValueError("7Z 文件列表不完整")
        members: list[_ArchiveMember] = []
        for info, raw_info in zip(infos, raw_infos):
            if info.filename != raw_info.filename:
                raise ValueError("7Z 文件列表顺序不一致")
            members.append(
                _ArchiveMember(
                    token=info,
                    name=info.filename,
                    file_size=info.uncompressed,
                    compress_size=info.compressed,
                    is_dir=raw_info.is_directory,
                    # py7zr 1.0 的公开 FileInfo 尚未包含这些字段，但底层
                    # ArchiveFile 已包含；统一从这里读取可避免依赖版本分叉。
                    is_link=raw_info.is_symlink or raw_info.is_junction,
                    is_regular=not raw_info.is_socket,
                )
            )
        plans = _validate_archive_members(
            archive_path,
            members,
            temp_dir,
            warnings,
            archive_compressed_bytes=archive_path.stat().st_size,
        )
        for plan in plans:
            if plan.member.is_dir:
                _target_for_plan(extract_dir, plan).mkdir(parents=True, exist_ok=True)
        factory = _build_7z_writer_factory(py7zr, extract_dir, plans)
        targets = [plan.member.name for plan in plans if not plan.member.is_dir]
        try:
            if targets:
                archive.extract(targets=targets, factory=factory)
            factory.finish()
        except Exception:
            factory.abort()
            raise


def _extract_tar(archive_path: Path, extract_dir: Path, temp_dir: Path, warnings: list[str]) -> None:
    with tarfile.open(archive_path, mode="r:*") as archive:
        infos = archive.getmembers()
        members = [
            _ArchiveMember(
                token=info,
                name=info.name,
                file_size=info.size,
                compress_size=None,
                is_dir=info.isdir(),
                is_link=info.issym() or info.islnk(),
                is_regular=info.isdir() or info.isreg(),
            )
            for info in infos
        ]
        plans = _validate_archive_members(
            archive_path,
            members,
            temp_dir,
            warnings,
            archive_compressed_bytes=archive_path.stat().st_size,
        )
        actual_total = [0]
        for plan in plans:
            target = _target_for_plan(extract_dir, plan)
            if plan.member.is_dir:
                target.mkdir(parents=True, exist_ok=True)
                continue
            source = archive.extractfile(plan.member.token)
            if source is None:
                raise ValueError(f"无法读取压缩包成员：{plan.member.name}")
            with source:
                _copy_member_stream(source, target, plan.member.file_size, actual_total)


def _validate_archive_members(
    archive_path: Path,
    members: list[_ArchiveMember],
    temp_dir: Path,
    warnings: list[str],
    *,
    archive_compressed_bytes: int,
) -> list[_ArchiveMemberPlan]:
    if len(members) > ARCHIVE_MAX_MEMBERS:
        raise ValueError(f"压缩包文件数量超过 {ARCHIVE_MAX_MEMBERS} 个安全上限")

    total_bytes = 0
    seen_names: set[str] = set()
    regular_names: set[str] = set()
    parent_names: set[str] = set()
    plans: list[_ArchiveMemberPlan] = []
    for member in members:
        if member.file_size < 0 or (member.compress_size is not None and member.compress_size < 0):
            raise ValueError("压缩包包含无效的文件大小")
        if not member.is_dir:
            if member.file_size > ARCHIVE_MAX_MEMBER_BYTES:
                raise ValueError(f"{member.name} 解压后超过单文件安全上限")
            total_bytes += member.file_size
            if total_bytes > ARCHIVE_MAX_TOTAL_BYTES:
                raise ValueError("压缩包解压后的总大小超过安全上限")
            if member.file_size >= ARCHIVE_RATIO_CHECK_MIN_BYTES and member.compress_size is not None:
                ratio = member.file_size / max(member.compress_size, 1)
                if ratio > ARCHIVE_MAX_COMPRESSION_RATIO:
                    raise ValueError(f"{member.name} 的压缩比例异常")

        parts = _safe_member_parts(member.name)
        if parts is None:
            warnings.append(f"{archive_path.name} 中存在不安全路径，已跳过：{member.name}")
            continue
        if not parts:
            continue
        normalized_name = "/".join(parts)
        canonical_name = normalized_name.casefold()
        if canonical_name in seen_names:
            raise ValueError(f"压缩包包含重复路径：{normalized_name}")
        canonical_parents = {
            "/".join(part.casefold() for part in parts[:index])
            for index in range(1, len(parts))
        }
        if canonical_parents & regular_names or (not member.is_dir and canonical_name in parent_names):
            raise ValueError(f"压缩包包含文件与目录路径冲突：{normalized_name}")
        seen_names.add(canonical_name)
        parent_names.update(canonical_parents)
        if not member.is_dir:
            regular_names.add(canonical_name)
        if member.is_link:
            warnings.append(f"{archive_path.name} 中存在链接文件，已跳过：{member.name}")
            continue
        if not member.is_regular:
            warnings.append(f"{archive_path.name} 中存在特殊文件，已跳过：{member.name}")
            continue
        plans.append(_ArchiveMemberPlan(member=member, parts=parts))

    if total_bytes >= ARCHIVE_RATIO_CHECK_MIN_BYTES:
        overall_ratio = total_bytes / max(archive_compressed_bytes, 1)
        if overall_ratio > ARCHIVE_MAX_COMPRESSION_RATIO:
            raise ValueError("压缩包的整体压缩比例异常")
    _validate_archive_free_space(temp_dir, total_bytes)
    return plans


def _validate_archive_free_space(temp_dir: Path, required_bytes: int) -> None:
    available = shutil.disk_usage(temp_dir).free
    if available < required_bytes + ARCHIVE_MIN_FREE_SPACE_BYTES:
        raise ValueError("临时磁盘空间不足，无法安全解压压缩包")


def _target_for_plan(extract_dir: Path, plan: _ArchiveMemberPlan) -> Path:
    extract_root = extract_dir.resolve()
    target = extract_dir.joinpath(*plan.parts).resolve()
    if not target.is_relative_to(extract_root):
        raise ValueError(f"压缩包成员路径不安全：{plan.member.name}")
    return target


def _copy_member_stream(
    source: BinaryIO,
    target: Path,
    expected_bytes: int,
    actual_total: list[int],
) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    partial_target = target.with_name(target.name + ".partial")
    member_bytes = 0
    try:
        with partial_target.open("wb") as output:
            while True:
                chunk = source.read(ARCHIVE_COPY_BUFFER_BYTES)
                if not chunk:
                    break
                next_member_bytes = member_bytes + len(chunk)
                next_total_bytes = actual_total[0] + len(chunk)
                if next_member_bytes > expected_bytes:
                    raise ValueError(f"压缩包成员实际大小不一致：{target.name}")
                if (
                    next_member_bytes > ARCHIVE_MAX_MEMBER_BYTES
                    or next_total_bytes > ARCHIVE_MAX_TOTAL_BYTES
                ):
                    raise ValueError("解压后的文件大小超过安全上限")
                written = output.write(chunk)
                if written != len(chunk):
                    raise OSError(f"无法完整写入解压文件：{target.name}")
                member_bytes = next_member_bytes
                actual_total[0] = next_total_bytes
        if member_bytes != expected_bytes:
            raise ValueError(f"压缩包成员实际大小不一致：{target.name}")
        partial_target.replace(target)
    except Exception:
        partial_target.unlink(missing_ok=True)
        raise


def _copy_rar_header(header, target: Path, expected_bytes: int, actual_total: list[int]) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    partial_target = target.with_name(target.name + ".partial")
    member_bytes = 0
    callback_error: Exception | None = None
    try:
        with partial_target.open("wb") as output:

            def write_chunk(chunk: bytes) -> None:
                nonlocal callback_error, member_bytes
                if callback_error is not None:
                    return
                try:
                    next_member_bytes = member_bytes + len(chunk)
                    next_total_bytes = actual_total[0] + len(chunk)
                    if next_member_bytes > expected_bytes:
                        raise ValueError(f"压缩包成员实际大小不一致：{target.name}")
                    if (
                        next_member_bytes > ARCHIVE_MAX_MEMBER_BYTES
                        or next_total_bytes > ARCHIVE_MAX_TOTAL_BYTES
                    ):
                        raise ValueError("解压后的文件大小超过安全上限")
                    written = output.write(chunk)
                    if written != len(chunk):
                        raise OSError(f"无法完整写入解压文件：{target.name}")
                    member_bytes = next_member_bytes
                    actual_total[0] = next_total_bytes
                except Exception as exc:  # CFFI 回调边界不能向外抛出异常
                    callback_error = exc

            header.test(write_chunk)
        if callback_error is not None:
            raise callback_error
        if member_bytes != expected_bytes:
            raise ValueError(f"压缩包成员实际大小不一致：{target.name}")
        partial_target.replace(target)
    except Exception:
        partial_target.unlink(missing_ok=True)
        raise


def _build_7z_writer_factory(py7zr, extract_dir: Path, plans: list[_ArchiveMemberPlan]):
    targets = {
        "/".join(plan.parts).casefold(): (_target_for_plan(extract_dir, plan), plan.member.file_size)
        for plan in plans
        if not plan.member.is_dir
    }
    actual_total = [0]

    class _Writer(py7zr.Py7zIO):
        def __init__(self, target: Path, expected_bytes: int) -> None:
            self.target = target
            self.expected_bytes = expected_bytes
            self.partial = target.with_name(target.name + ".partial")
            target.parent.mkdir(parents=True, exist_ok=True)
            self.handle = self.partial.open("w+b")
            self.max_written = 0
            self.closed = False

        def write(self, data: bytes | bytearray) -> int:
            end = self.handle.tell() + len(data)
            next_max_written = max(self.max_written, end)
            next_total = actual_total[0] + (next_max_written - self.max_written)
            if next_max_written > self.expected_bytes:
                raise ValueError(f"压缩包成员实际大小不一致：{self.target.name}")
            if next_max_written > ARCHIVE_MAX_MEMBER_BYTES or next_total > ARCHIVE_MAX_TOTAL_BYTES:
                raise ValueError("解压后的文件大小超过安全上限")
            written = self.handle.write(data)
            if written != len(data):
                raise OSError(f"无法完整写入解压文件：{self.target.name}")
            self.max_written = next_max_written
            actual_total[0] = next_total
            return written

        def read(self, size: int | None = None) -> bytes:
            return self.handle.read(-1 if size is None else size)

        def seek(self, offset: int, whence: int = 0) -> int:
            return self.handle.seek(offset, whence)

        def flush(self) -> None:
            self.handle.flush()

        def size(self) -> int:
            return self.max_written

        def finish(self) -> None:
            if self.closed:
                return
            self.handle.flush()
            self.handle.close()
            self.closed = True
            if self.max_written != self.expected_bytes:
                raise ValueError(f"压缩包成员实际大小不一致：{self.target.name}")
            self.partial.replace(self.target)

        def abort(self) -> None:
            if not self.closed:
                self.handle.close()
                self.closed = True
            self.partial.unlink(missing_ok=True)

    class _Factory(py7zr.WriterFactory):
        def __init__(self) -> None:
            self.writers: list[_Writer] = []
            self.created_targets: set[Path] = set()

        def create(self, filename: str):
            parts = _safe_member_parts(filename)
            entry = None if not parts else targets.get("/".join(parts).casefold())
            if entry is None:
                raise ValueError(f"7Z 返回未通过安全校验的成员：{filename}")
            writer = _Writer(entry[0], entry[1])
            self.writers.append(writer)
            self.created_targets.add(entry[0])
            return writer

        def finish(self) -> None:
            for writer in self.writers:
                writer.finish()
            missing = [(target, size) for target, size in targets.values() if target not in self.created_targets]
            for target, size in missing:
                if size != 0:
                    raise ValueError(f"7Z 未解压预期成员：{target.name}")
                target.parent.mkdir(parents=True, exist_ok=True)
                target.touch(exist_ok=False)

        def abort(self) -> None:
            for writer in self.writers:
                writer.abort()

    return _Factory()


def _zip_member_is_symlink(member: zipfile.ZipInfo) -> bool:
    unix_mode = (member.external_attr >> 16) & 0xFFFF
    return stat.S_ISLNK(unix_mode)


def _safe_member_parts(name: str) -> tuple[str, ...] | None:
    normalized = unicodedata.normalize("NFC", str(name).replace("\\", "/"))
    if "\x00" in normalized or normalized.startswith("/"):
        return None
    raw_parts = normalized.split("/")
    if any(part == ".." for part in raw_parts):
        return None
    parts = tuple(part for part in raw_parts if part not in {"", "."})
    if parts and len(parts[0]) >= 2 and parts[0][1] == ":":
        return None
    return parts


def _relative_sort_key(path: Path, root: Path) -> tuple[str, str]:
    relative = unicodedata.normalize("NFC", path.relative_to(root).as_posix())
    return relative.casefold(), relative


def _friendly_archive_error(exc: Exception) -> str:
    text = str(exc).strip()
    lowered = text.casefold()
    if "password" in lowered or "encrypted" in lowered or "密码" in text:
        return "压缩包带密码，当前未提供密码输入，无法读取"
    return text or exc.__class__.__name__
