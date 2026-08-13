"""Self-contained, user-visible project workspaces for HR Toolkit.

The project directory is the portable unit.  User files live in ordinary,
visible folders while control data (project identity, batch manifests, locks,
staging files and the recycle bin) lives below ``.hrtoolkit``.  Batch manifests
are the source of truth; every stored path is relative to the project root.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import stat as stat_module
import sys
import threading
import uuid
from contextlib import AbstractContextManager
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator, Literal, Sequence

try:
    from zoneinfo import ZoneInfo

    _SHANGHAI_TZ = ZoneInfo("Asia/Shanghai")
except Exception:  # pragma: no cover - 极端环境兜底
    from datetime import timedelta

    _SHANGHAI_TZ = timezone(timedelta(hours=8))


def _shanghai_now() -> datetime:
    return datetime.now(tz=_SHANGHAI_TZ)

from hr_toolkit.history_store import (
    COPY_BUFFER_BYTES,
    MIN_FREE_SPACE_BYTES,
    _atomic_write_text,
    _exclusive_file_lock,
    _is_link_like,
    _make_private,
    _mkdir_private,
    _safe_component,
    _sha256_file,
)


PROJECT_FORMAT_VERSION = 1
PROJECT_METADATA_DIR = ".hrtoolkit"
PROJECT_FILE_NAME = "project.json"
BATCH_MANIFEST_DIR = "manifests"
PROJECT_STAGING_DIR = "staging"
PROJECT_TRASH_DIR = "trash"
PROJECT_QUARANTINE_DIR = "quarantine"
PROJECT_WRITE_LOCK = "project-write.lock"
VISIBLE_IMPORT_JOURNAL = "operation.json"
VISIBLE_IMPORT_KIND = "visible_directory_import"
VISIBLE_IMPORT_VERSION = 1
COMMON_VISIBLE_DIR = "共用资料"

CATEGORY_UPLOADS = "uploads"
CATEGORY_RESULTS = "results"
CATEGORY_SUPPLEMENTS = "supplements"
CATEGORIES = (CATEGORY_UPLOADS, CATEGORY_RESULTS, CATEGORY_SUPPLEMENTS)
CATEGORY_VISIBLE_NAMES = {
    CATEGORY_UPLOADS: "上传资料",
    CATEGORY_RESULTS: "处理结果",
    CATEGORY_SUPPLEMENTS: "补充资料",
}

BATCH_STATUSES = {"draft", "running", "success", "failed", "stopped"}
TERMINAL_BATCH_STATUSES = {"success", "failed", "stopped"}
PROJECT_COPY_RESERVE_RATIO = 0.05
IGNORED_IMPORT_NAMES = {".ds_store", "thumbs.db", "desktop.ini"}
IGNORED_IMPORT_SUFFIXES = {".tmp", ".temp"}
FORBIDDEN_IMPORT_SUFFIXES = {
    ".app",
    ".bat",
    ".cmd",
    ".com",
    ".dll",
    ".exe",
    ".hta",
    ".jar",
    ".js",
    ".jse",
    ".lnk",
    ".msi",
    ".msp",
    ".ps1",
    ".py",
    ".pyw",
    ".reg",
    ".scr",
    ".sh",
    ".url",
    ".vbe",
    ".vbs",
    ".wsf",
    ".wsh",
}
PROJECT_NAME_MAX_LENGTH = 120
PROJECT_NAME_FORBIDDEN_CHARS = frozenset('<>:"/\\|?*')
WINDOWS_RESERVED_NAMES = {
    "con",
    "prn",
    "aux",
    "nul",
    *(f"com{index}" for index in range(1, 10)),
    *(f"lpt{index}" for index in range(1, 10)),
}


class ProjectStoreError(RuntimeError):
    """Raised when a project operation cannot be completed without data risk."""


class ImportCancelled(ProjectStoreError):
    """Raised when an import is safely stopped before final publication."""


@dataclass(frozen=True)
class ImportProgress:
    """A business-neutral snapshot for rendering import progress in the GUI."""

    phase: Literal["checking", "copying", "finalizing"]
    current_name: str | None = None
    files_scanned: int = 0
    files_completed: int = 0
    files_total: int | None = None
    bytes_copied: int = 0
    bytes_total: int | None = None


def validate_project_name(value: str) -> str:
    """Return a portable project name or raise a human-readable error."""

    project_name = str(value).strip()
    if not project_name:
        raise ProjectStoreError("项目名称不能为空。")
    if len(project_name) > PROJECT_NAME_MAX_LENGTH:
        raise ProjectStoreError(f"项目名称不能超过 {PROJECT_NAME_MAX_LENGTH} 个字。")
    if project_name in {".", ".."}:
        raise ProjectStoreError("项目名称不能使用英文句点。")
    if project_name.endswith("."):
        raise ProjectStoreError("项目名称末尾不能使用句点。")
    if any(ord(character) < 32 for character in project_name):
        raise ProjectStoreError("项目名称不能包含换行或控制字符。")
    if any(character in PROJECT_NAME_FORBIDDEN_CHARS for character in project_name):
        raise ProjectStoreError('项目名称不能包含 \\ / : * ? " < > |。')
    if project_name.casefold() == PROJECT_METADATA_DIR.casefold():
        raise ProjectStoreError("该名称是项目保留名称，请换一个名称。")
    portable_base = project_name.split(".", 1)[0].casefold()
    if portable_base in WINDOWS_RESERVED_NAMES:
        raise ProjectStoreError("该名称是 Windows 系统保留名称，请换一个名称。")
    return project_name


@dataclass(frozen=True)
class ProjectWorkspace:
    project_id: str
    name: str
    root: Path
    created_at: str
    format_version: int
    writable: bool
    read_only_reason: str | None = None

    @property
    def metadata_dir(self) -> Path:
        return self.root / PROJECT_METADATA_DIR

    @property
    def common_root(self) -> Path:
        return self.root / COMMON_VISIBLE_DIR


@dataclass(frozen=True)
class ProjectFile:
    id: str
    batch_id: str
    category: str
    role: str
    display_name: str
    relative_path: str
    size_bytes: int
    sha256: str
    modified_ns: int

    def path(self, workspace: ProjectWorkspace) -> Path:
        return _project_join(workspace.root, self.relative_path)


@dataclass(frozen=True)
class BatchSummary:
    id: str
    group_name: str
    tool_id: str
    tool_name: str
    status: str
    directory_name: str
    business_description: str
    business_period: str
    created_at: str
    started_at: str | None
    finished_at: str | None
    error_message: str | None
    deleted_at: str | None = None


@dataclass(frozen=True)
class BatchDetail:
    summary: BatchSummary
    directories: dict[str, Path]
    files: tuple[ProjectFile, ...]

    def files_for(self, category: str) -> tuple[ProjectFile, ...]:
        return tuple(item for item in self.files if item.category == category)


@dataclass(frozen=True)
class TrashBatchDetail:
    """Safe, aggregate-only information for the project recycle-bin UI."""

    summary: BatchSummary
    original_relative_path: str
    upload_count: int
    result_count: int
    supplement_count: int
    total_size_bytes: int


@dataclass(frozen=True)
class _ExternalSourceItem:
    path: Path
    source_index: int
    top_name: str
    top_is_directory: bool
    relative_parts: tuple[str, ...]
    expected_size_bytes: int | None = None
    expected_sha256: str | None = None


class ProjectStore:
    """One opened project workspace.

    Use :meth:`create` or :meth:`open`.  A writable store owns the project-wide
    writer lock until :meth:`close`; a second instance opens read-only by
    default.  Unknown future project versions are always read-only.
    """

    def __init__(
        self,
        workspace: ProjectWorkspace,
        *,
        writer_lock: AbstractContextManager[None] | None = None,
    ) -> None:
        self.workspace = workspace
        self.root = workspace.root
        self.metadata_dir = workspace.metadata_dir
        self.manifest_dir = self.metadata_dir / BATCH_MANIFEST_DIR
        self.staging_dir = self.metadata_dir / PROJECT_STAGING_DIR
        self.trash_dir = self.metadata_dir / PROJECT_TRASH_DIR
        self.quarantine_dir = self.metadata_dir / PROJECT_QUARANTINE_DIR
        self._writer_lock = writer_lock
        self._closed = False
        self._session_id = uuid.uuid4().hex
        self._mutex = threading.RLock()

    @classmethod
    def create(cls, root: str | Path, name: str) -> "ProjectStore":
        project_root = _validate_project_location(Path(root).expanduser(), for_write=True)
        project_name = validate_project_name(name)

        created_root = False
        created_metadata = False
        try:
            if project_root.exists():
                if not project_root.is_dir() or _is_link_like(project_root):
                    raise ProjectStoreError("项目位置不是安全的普通文件夹。")
                if next(project_root.iterdir(), None) is not None:
                    raise ProjectStoreError("新项目必须使用空文件夹，避免覆盖已有资料。")
            else:
                project_root.mkdir(parents=True, exist_ok=False)
                created_root = True
            _make_private(project_root, directory=True)

            metadata_dir = project_root / PROJECT_METADATA_DIR
            metadata_dir.mkdir(mode=0o700, exist_ok=False)
            created_metadata = True
            _make_private(metadata_dir, directory=True)
            _hide_on_windows(metadata_dir)
            for name_part in (
                BATCH_MANIFEST_DIR,
                PROJECT_STAGING_DIR,
                PROJECT_TRASH_DIR,
                PROJECT_QUARANTINE_DIR,
            ):
                _mkdir_private(metadata_dir / name_part)
            (project_root / COMMON_VISIBLE_DIR).mkdir(mode=0o700, exist_ok=False)
            _make_private(project_root / COMMON_VISIBLE_DIR, directory=True)

            created_at = _utc_now()
            payload = {
                "format_version": PROJECT_FORMAT_VERSION,
                "minimum_reader_version": 1,
                "project_id": uuid.uuid4().hex,
                "name": project_name,
                "created_at": created_at,
            }
            _write_json(metadata_dir / PROJECT_FILE_NAME, payload)
            return cls.open(project_root, writable=True, read_only_fallback=False)
        except Exception:
            if created_root:
                shutil.rmtree(project_root, ignore_errors=True)
            elif created_metadata:
                # The caller supplied an empty directory.  Remove only entries
                # this method created and leave the caller's directory itself.
                try:
                    (project_root / COMMON_VISIBLE_DIR).rmdir()
                except OSError:
                    pass
                shutil.rmtree(project_root / PROJECT_METADATA_DIR, ignore_errors=True)
            raise

    @classmethod
    def open(
        cls,
        root: str | Path,
        *,
        writable: bool = True,
        read_only_fallback: bool = True,
    ) -> "ProjectStore":
        requested = Path(root).expanduser()
        try:
            project_root = _validate_project_location(requested, for_write=writable)
            unsafe_reason = None
        except ProjectStoreError as exc:
            if not writable or not read_only_fallback:
                raise
            project_root = _validate_project_location(requested, for_write=False)
            unsafe_reason = str(exc)

        metadata_dir = project_root / PROJECT_METADATA_DIR
        project_path = metadata_dir / PROJECT_FILE_NAME
        _require_regular_directory(metadata_dir, "项目管理目录")
        _require_regular_file(project_path, "项目标记")
        payload = _read_json(project_path)
        format_version = _required_int(payload, "format_version")
        if format_version < 1:
            raise ProjectStoreError("项目版本无效，不能安全打开。")
        project_id = _required_uuid(payload, "project_id")
        project_name = str(payload.get("name") or "").strip()
        created_at = str(payload.get("created_at") or "")
        if not project_name or not created_at:
            raise ProjectStoreError("项目标记缺少必要信息。")

        writer_lock: AbstractContextManager[None] | None = None
        read_only_reason = unsafe_reason
        wants_write = writable and unsafe_reason is None
        if format_version > PROJECT_FORMAT_VERSION:
            wants_write = False
            read_only_reason = "该项目由更高版本 HRToolkit 创建，当前以只读方式打开。"
        elif wants_write:
            lock_path = metadata_dir / PROJECT_WRITE_LOCK
            writer_lock = _exclusive_file_lock(lock_path, blocking=False)
            try:
                writer_lock.__enter__()
            except BlockingIOError as exc:
                writer_lock = None
                if not read_only_fallback:
                    raise ProjectStoreError("项目正在另一个 HRToolkit 窗口中使用。") from exc
                wants_write = False
                read_only_reason = "项目正在另一个窗口中使用，当前以只读方式打开。"
            except Exception as exc:
                writer_lock = None
                if not read_only_fallback:
                    raise ProjectStoreError(f"无法锁定项目：{exc}") from exc
                wants_write = False
                read_only_reason = "无法取得项目写入权限，当前以只读方式打开。"

        workspace = ProjectWorkspace(
            project_id=project_id,
            name=project_name,
            root=project_root,
            created_at=created_at,
            format_version=format_version,
            writable=wants_write,
            read_only_reason=read_only_reason,
        )
        store = cls(workspace, writer_lock=writer_lock)
        if wants_write:
            try:
                store._prepare_known_layout()
                store._recover_workspace()
            except Exception:
                store.close()
                raise
        return store

    @property
    def writable(self) -> bool:
        return self.workspace.writable and not self._closed

    def __enter__(self) -> "ProjectStore":
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        self.close()

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        if self._writer_lock is not None:
            self._writer_lock.__exit__(None, None, None)
            self._writer_lock = None

    def _require_writable(self) -> None:
        if self._closed:
            raise ProjectStoreError("项目已经关闭。")
        if not self.workspace.writable:
            reason = self.workspace.read_only_reason or "项目当前为只读状态。"
            raise ProjectStoreError(reason)

    def _prepare_known_layout(self) -> None:
        _require_regular_directory(self.metadata_dir, "项目管理目录")
        _hide_on_windows(self.metadata_dir)
        for path in (self.manifest_dir, self.staging_dir, self.trash_dir, self.quarantine_dir):
            if path.exists():
                _require_regular_directory(path, "项目内部目录")
            else:
                _mkdir_private(path)
        common_root = self.root / COMMON_VISIBLE_DIR
        if common_root.exists():
            _require_regular_directory(common_root, COMMON_VISIBLE_DIR)
        else:
            common_root.mkdir(mode=0o700)
            _make_private(common_root, directory=True)

    def create_draft_batch(
        self,
        *,
        group_name: str,
        tool_id: str,
        tool_name: str,
        business_description: str = "",
        business_period: str = "",
    ) -> BatchDetail:
        self._require_writable()
        group = str(group_name).strip() or "默认分组"
        clean_tool_id = str(tool_id).strip()
        clean_tool_name = str(tool_name).strip()
        if not clean_tool_id or not clean_tool_name:
            raise ProjectStoreError("工具信息不能为空。")
        batch_id = uuid.uuid4().hex
        draft_name = f"草稿_{batch_id[:8]}"
        directories = self._batch_relative_directories(group, clean_tool_name, draft_name)
        payload = {
            "format_version": PROJECT_FORMAT_VERSION,
            "batch": {
                "id": batch_id,
                "group_name": group,
                "tool_id": clean_tool_id,
                "tool_name": clean_tool_name,
                "status": "draft",
                "directory_name": draft_name,
                "business_description": str(business_description).strip(),
                "business_period": str(business_period).strip(),
                "created_at": _utc_now(),
                "started_at": None,
                "finished_at": None,
                "error_message": None,
                "writer_session": None,
                "deleted_at": None,
            },
            "directories": directories,
            "files": [],
            "pending_import": None,
            "pending_rename": None,
            "pending_trash": None,
            "recovery_notes": [],
        }
        manifest_path = self._manifest_path(batch_id)
        with self._mutex:
            batch_relative_root = Path(directories[CATEGORY_UPLOADS]).parent
            _assert_no_link_components(self.root, batch_relative_root)
            _write_json(manifest_path, payload)
            batch_root = _project_join(self.root, directories[CATEGORY_UPLOADS]).parent
            try:
                batch_root.mkdir(parents=True, exist_ok=False)
                _make_private(batch_root, directory=True)
                for category in (CATEGORY_UPLOADS, CATEGORY_RESULTS):
                    path = _project_join(self.root, directories[category])
                    path.mkdir(mode=0o700, exist_ok=False)
                    _make_private(path, directory=True)
            except Exception:
                shutil.rmtree(batch_root, ignore_errors=True)
                manifest_path.unlink(missing_ok=True)
                raise
        detail = self.get_batch(batch_id)
        assert detail is not None
        return detail

    def create_draft(
        self,
        *,
        group_name: str,
        tool_id: str,
        tool_name: str,
        business_description: str = "",
        business_period: str = "",
    ) -> BatchDetail:
        return self.create_draft_batch(
            group_name=group_name,
            tool_id=tool_id,
            tool_name=tool_name,
            business_description=business_description,
            business_period=business_period,
        )

    def start_processing(
        self,
        batch_id: str,
        *,
        business_description: str | None = None,
        business_period: str | None = None,
        now: datetime | None = None,
    ) -> BatchDetail:
        self._require_writable()
        with self._mutex:
            manifest = self._load_active_manifest(batch_id)
            batch = _batch_object(manifest)
            if batch["status"] != "draft":
                raise ProjectStoreError("只有草稿批次可以开始处理。")
            self._verify_manifest_files(manifest)
            self._assert_no_unregistered_results(manifest)
            description = str(
                batch.get("business_description") if business_description is None else business_description
            ).strip() or "未命名事项"
            period = str(batch.get("business_period") if business_period is None else business_period).strip() or "未填写期间"
            # 项目批次目录命名固定使用北京时间，与运行环境时区无关
            local_now = (now or _shanghai_now()).astimezone(_SHANGHAI_TZ)
            base_name = "_".join(
                (
                    local_now.strftime("%Y%m%d_%H%M%S"),
                    _visible_component(description),
                    _visible_component(period),
                )
            )
            old_directories = _directory_map(manifest)
            new_name, new_directories = self._unique_batch_directories(
                str(batch["group_name"]),
                str(batch["tool_name"]),
                base_name,
            )
            manifest["pending_rename"] = {
                "old_directories": old_directories,
                "new_directories": new_directories,
                "directory_name": new_name,
                "business_description": description,
                "business_period": period,
                "started_at": _utc_from_datetime(local_now),
                "writer_session": self._session_id,
            }
            self._write_active_manifest(batch_id, manifest)
            self._recover_pending_rename(batch_id, manifest)
        detail = self.get_batch(batch_id)
        assert detail is not None
        return detail

    def start_batch(
        self,
        batch_id: str,
        *,
        business_description: str | None = None,
        business_period: str | None = None,
        now: datetime | None = None,
    ) -> BatchDetail:
        return self.start_processing(
            batch_id,
            business_description=business_description,
            business_period=business_period,
            now=now,
        )

    def mark_success(self, batch_id: str) -> BatchDetail:
        return self._finish_batch(batch_id, "success", None)

    def mark_failed(self, batch_id: str, error: str) -> BatchDetail:
        return self._finish_batch(batch_id, "failed", error)

    def mark_stopped(self, batch_id: str, error: str = "用户停止了本次处理。") -> BatchDetail:
        return self._finish_batch(batch_id, "stopped", error)

    def _finish_batch(self, batch_id: str, status: str, error: str | None) -> BatchDetail:
        self._require_writable()
        if status not in TERMINAL_BATCH_STATUSES:
            raise ValueError(f"不支持的批次状态：{status}")
        with self._mutex:
            manifest = self._load_active_manifest(batch_id)
            batch = _batch_object(manifest)
            if batch["status"] != "running":
                raise ProjectStoreError("只有正在处理的批次可以结束。")
            if status == "success":
                self._verify_manifest_files(manifest)
                self._assert_no_unregistered_results(manifest)
            else:
                self._quarantine_unregistered_results(manifest, reason=status)
            batch["status"] = status
            batch["finished_at"] = _utc_now()
            batch["error_message"] = _clean_error(error)
            batch["writer_session"] = None
            self._write_active_manifest(batch_id, manifest)
        detail = self.get_batch(batch_id)
        assert detail is not None
        return detail

    def import_sources(
        self,
        batch_id: str,
        sources: Iterable[str | Path],
        *,
        category: str = CATEGORY_UPLOADS,
        role: str = "main",
        cancelled: Callable[[], bool] | None = None,
        progress: Callable[[int, int, str], None] | None = None,
        on_progress: Callable[[ImportProgress], None] | None = None,
    ) -> tuple[ProjectFile, ...]:
        """Copy external regular files into a draft/running batch.

        The sources are never moved or written.  The copy is staged, hashed,
        declared in the batch manifest, and only then published without
        replacing an existing project file.
        """

        return self._copy_sources_to_batch(
            batch_id,
            sources,
            category=category,
            role=role,
            project_sources=False,
            cancelled=cancelled,
            progress=progress,
            on_progress=on_progress,
        )

    def copy_project_sources(
        self,
        batch_id: str,
        sources: Iterable[str | Path],
        *,
        category: str = CATEGORY_UPLOADS,
        role: str = "main",
        cancelled: Callable[[], bool] | None = None,
        progress: Callable[[int, int, str], None] | None = None,
        on_progress: Callable[[ImportProgress], None] | None = None,
    ) -> tuple[ProjectFile, ...]:
        """Snapshot existing visible project material into another batch.

        Sources may come from common material, another batch's uploads or
        supplements, or the registered results of a successfully completed
        batch.  The destination receives an independent, hashed copy.
        """

        return self._copy_sources_to_batch(
            batch_id,
            sources,
            category=category,
            role=role,
            project_sources=True,
            cancelled=cancelled,
            progress=progress,
            on_progress=on_progress,
        )

    def _copy_sources_to_batch(
        self,
        batch_id: str,
        sources: Iterable[str | Path],
        *,
        category: str,
        role: str,
        project_sources: bool,
        cancelled: Callable[[], bool] | None,
        progress: Callable[[int, int, str], None] | None,
        on_progress: Callable[[ImportProgress], None] | None,
    ) -> tuple[ProjectFile, ...]:
        """Stage, hash, declare and publish files into a batch."""

        self._require_writable()
        if category not in {CATEGORY_UPLOADS, CATEGORY_SUPPLEMENTS}:
            raise ProjectStoreError("资料只能导入到上传资料或补充资料。")
        with self._mutex:
            _report_import_progress(on_progress, ImportProgress(phase="checking"))
            _raise_if_cancelled(cancelled)
            manifest = self._load_active_manifest(batch_id)
            batch = _batch_object(manifest)
            allowed = {"draft"} if category == CATEGORY_UPLOADS else BATCH_STATUSES
            if batch["status"] not in allowed:
                raise ProjectStoreError("当前批次状态不允许继续导入资料。")
            if manifest.get("pending_import"):
                self._recover_pending_import(batch_id, manifest)
                manifest = self._load_active_manifest(batch_id)

            destination_relative = Path(_directory_map(manifest)[category])
            _assert_no_link_components(self.root, destination_relative)
            destination_root = _project_join(self.root, destination_relative)
            if project_sources:
                source_items = self._collect_project_sources(
                    sources,
                    target_batch_root=destination_root.parent,
                    cancelled=cancelled,
                    on_progress=on_progress,
                )
            else:
                source_items = self._collect_external_sources(
                    sources,
                    cancelled=cancelled,
                    on_progress=on_progress,
                )
            _raise_if_cancelled(cancelled)
            required_bytes = _source_total_bytes(source_items, cancelled=cancelled)
            _report_import_progress(
                on_progress,
                ImportProgress(
                    phase="checking",
                    files_scanned=len(source_items),
                    files_total=len(source_items),
                    bytes_total=required_bytes,
                ),
            )
            _raise_if_cancelled(cancelled)
            self._ensure_free_space(required_bytes)
            _raise_if_cancelled(cancelled)
            reserved_keys = _existing_name_keys(destination_root, cancelled=cancelled)
            planned: list[tuple[_ExternalSourceItem, Path, str]] = []
            assigned_top_paths: dict[int, Path] = {}
            for item in source_items:
                _raise_if_cancelled(cancelled)
                top_path = assigned_top_paths.get(item.source_index)
                if top_path is None:
                    desired_top = _project_join(destination_root, _visible_component(item.top_name))
                    top_path = _unique_destination(
                        desired_top,
                        reserved_keys,
                        destination_root,
                        is_directory=item.top_is_directory,
                        cancelled=cancelled,
                    )
                    assigned_top_paths[item.source_index] = top_path
                    reserved_keys.add(_portable_relative_key(top_path.relative_to(destination_root)))
                if item.top_is_directory:
                    safe_parts = tuple(_visible_component(part) for part in item.relative_parts)
                    desired = _project_join(top_path, Path(*safe_parts))
                    destination = _unique_destination(
                        desired,
                        reserved_keys,
                        destination_root,
                        cancelled=cancelled,
                    )
                else:
                    destination = top_path
                reserved_keys.add(_portable_relative_key(destination.relative_to(destination_root)))
                planned.append((item, destination, item.path.name))

            operation_id = uuid.uuid4().hex
            staging_root = self.staging_dir / operation_id
            staging_root.mkdir(mode=0o700, exist_ok=False)
            _make_private(staging_root, directory=True)
            pending_items: list[dict[str, Any]] = []
            copied_bytes = 0
            try:
                for index, (source_item, destination, display_name) in enumerate(planned):
                    _raise_if_cancelled(cancelled)
                    staging_path = staging_root / f"{index:08d}.data"
                    metadata = _copy_external_file(
                        source_item.path,
                        staging_path,
                        cancelled=cancelled,
                        on_chunk=lambda size, name=display_name, completed=index: _report_copy_callbacks(
                            progress,
                            on_progress,
                            copied_bytes + size,
                            required_bytes,
                            name,
                            files_scanned=len(source_items),
                            files_completed=completed,
                            files_total=len(planned),
                        ),
                    )
                    if (
                        source_item.expected_size_bytes is not None
                        and source_item.expected_sha256 is not None
                        and (
                            int(metadata["size_bytes"]) != source_item.expected_size_bytes
                            or str(metadata["sha256"]) != source_item.expected_sha256
                        )
                    ):
                        staging_path.unlink(missing_ok=True)
                        raise ProjectStoreError(
                            f"项目资料与原清单不一致，不能复用：{display_name}"
                        )
                    copied_bytes += int(metadata["size_bytes"])
                    _report_copy_callbacks(
                        progress=None,
                        on_progress=on_progress,
                        copied_bytes=copied_bytes,
                        total_bytes=required_bytes,
                        name=display_name,
                        files_scanned=len(source_items),
                        files_completed=index + 1,
                        files_total=len(planned),
                    )
                    pending_items.append(
                        {
                            "id": uuid.uuid4().hex,
                            "batch_id": batch_id,
                            "category": category,
                            "role": str(role)[:100],
                            "display_name": display_name,
                            "relative_path": destination.relative_to(self.root).as_posix(),
                            "staging_path": staging_path.relative_to(self.root).as_posix(),
                            "size_bytes": int(metadata["size_bytes"]),
                            "sha256": str(metadata["sha256"]),
                            "modified_ns": int(metadata["modified_ns"]),
                        }
                    )
                _raise_if_cancelled(cancelled)
                _report_import_progress(
                    on_progress,
                    ImportProgress(
                        phase="finalizing",
                        files_scanned=len(source_items),
                        files_completed=len(planned),
                        files_total=len(planned),
                        bytes_copied=copied_bytes,
                        bytes_total=required_bytes,
                    ),
                )
                manifest["pending_import"] = {
                    "operation_id": operation_id,
                    "created_at": _utc_now(),
                    "items": pending_items,
                }
                self._write_active_manifest(batch_id, manifest)
                self._recover_pending_import(batch_id, manifest)
            except Exception:
                current = self._load_active_manifest(batch_id)
                if not current.get("pending_import"):
                    shutil.rmtree(staging_root, ignore_errors=True)
                raise
        detail = self.get_batch(batch_id)
        assert detail is not None
        ids = {str(item["id"]) for item in pending_items}
        return tuple(item for item in detail.files if item.id in ids)

    def result_directory(self, batch_id: str) -> Path:
        self._require_writable()
        detail = self.get_batch(batch_id)
        if detail is None:
            raise ProjectStoreError("批次不存在。")
        if detail.summary.status != "running":
            raise ProjectStoreError("只有正在处理的批次可以写入结果。")
        return detail.directories[CATEGORY_RESULTS]

    def create_result_working_copy(self, batch_id: str, source: str | Path) -> Path:
        """Create a verified directory copy for tools that modify a folder tree.

        Only files declared in this batch's upload manifest are copied.  The
        copy is made file-by-file with no-follow opens and every copied digest
        must still match the upload snapshot, so later links or unregistered
        files in the visible upload directory cannot enter formal results.
        """

        self._require_writable()
        with self._mutex:
            manifest = self._load_active_manifest(batch_id)
            batch = _batch_object(manifest)
            if batch["status"] != "running":
                raise ProjectStoreError("只有正在处理的批次可以建立结果副本。")
            directories = _directory_map(manifest)
            upload_root = _project_join(self.root, directories[CATEGORY_UPLOADS])
            result_root = _project_join(self.root, directories[CATEGORY_RESULTS])
            raw_source = Path(source).expanduser()
            source_path = raw_source if raw_source.is_absolute() else self.root / raw_source
            _assert_existing_ancestors_are_real(source_path.absolute())
            if _is_link_like(source_path) or not source_path.exists():
                raise ProjectStoreError("人员资料文件夹不存在或是链接。")
            source_path = source_path.resolve()
            if not _is_inside(source_path, upload_root):
                raise ProjectStoreError("只能从本次批次的上传资料建立结果副本。")
            _assert_no_link_components(self.root, source_path.relative_to(self.root))
            _require_regular_directory(source_path, "人员资料文件夹")

            registered = {
                _project_join(self.root, str(item["relative_path"])): item
                for item in _file_objects(manifest)
                if str(item.get("category")) == CATEGORY_UPLOADS
                and _is_inside(
                    _project_join(self.root, str(item["relative_path"])),
                    source_path,
                )
            }
            walked = list(_walk_directory_strict(source_path, ignore_temporary=False))
            actual_paths = {path for path, _parts in walked}
            if actual_paths != set(registered):
                raise ProjectStoreError("上传资料与本次清单不一致，已停止建立处理副本。")

            destination = _project_join(result_root, _visible_component(source_path.name))
            if destination.exists() or _is_link_like(destination):
                raise ProjectStoreError(f"处理结果目录中已存在同名文件夹：{destination.name}")

            directory_parts: set[tuple[str, ...]] = {()}
            for current, child_directories, child_files in os.walk(source_path, followlinks=False):
                current_path = Path(current)
                for name in (*child_directories, *child_files):
                    if _is_link_like(current_path / name):
                        raise ProjectStoreError(f"上传资料包含链接，已停止处理：{name}")
                current_parts = current_path.relative_to(source_path).parts
                directory_parts.add(current_parts)
                for name in child_directories:
                    directory_parts.add((*current_parts, name))

            destination.mkdir(mode=0o700, exist_ok=False)
            _make_private(destination, directory=True)
            try:
                for relative_parts in sorted(directory_parts, key=lambda parts: (len(parts), parts)):
                    directory = (
                        destination
                        if not relative_parts
                        else _project_join(destination, Path(*relative_parts))
                    )
                    directory.mkdir(parents=True, exist_ok=True)
                    _make_private(directory, directory=True)
                for source_file, relative_parts in walked:
                    _assert_existing_ancestors_are_real(source_file.absolute())
                    destination_file = _project_join(destination, Path(*relative_parts))
                    metadata = _copy_external_file(
                        source_file,
                        destination_file,
                        cancelled=None,
                        on_chunk=None,
                    )
                    expected = registered[source_file]
                    if (
                        int(metadata["size_bytes"]) != int(expected["size_bytes"])
                        or str(metadata["sha256"]) != str(expected["sha256"])
                    ):
                        raise ProjectStoreError(
                            f"上传资料已发生变化，不能继续处理：{source_file.name}"
                        )
            except Exception:
                shutil.rmtree(destination, ignore_errors=True)
                raise
            return destination

    def register_results(
        self,
        batch_id: str,
        paths: Iterable[str | Path] | str | Path,
        *,
        role: str = "result",
    ) -> tuple[ProjectFile, ...]:
        """Register results already written inside the batch result folder.

        No file is copied.  Registration records a stable size and SHA-256 and
        rejects paths outside the batch's result directory.
        """

        self._require_writable()
        raw_paths: Sequence[str | Path]
        if isinstance(paths, (str, Path)):
            raw_paths = (paths,)
        else:
            raw_paths = tuple(paths)
        with self._mutex:
            manifest = self._load_active_manifest(batch_id)
            batch = _batch_object(manifest)
            if batch["status"] != "running":
                raise ProjectStoreError("只有正在处理的批次可以登记结果。")
            result_root = _project_join(self.root, _directory_map(manifest)[CATEGORY_RESULTS])
            candidates: list[Path] = []
            for raw in raw_paths:
                path = Path(raw).expanduser()
                if _is_link_like(path) or not path.exists():
                    raise ProjectStoreError(f"结果文件不存在或是链接：{path.name or path}")
                resolved = path.resolve()
                if not _is_inside(resolved, result_root):
                    raise ProjectStoreError("只能登记当前批次结果目录中的文件。")
                if resolved.is_dir():
                    candidates.extend(path_item for path_item, _parts in _walk_external_directory(resolved))
                elif resolved.is_file():
                    if _is_ignored_import_file(resolved):
                        continue
                    _reject_forbidden_import_file(resolved)
                    candidates.append(resolved)
                else:
                    raise ProjectStoreError(f"结果不是普通文件：{resolved.name}")
            candidates = sorted(set(candidates), key=lambda item: item.as_posix().casefold())
            existing = {str(item["relative_path"]): item for item in _file_objects(manifest)}
            added: list[dict[str, Any]] = []
            for path in candidates:
                relative = path.relative_to(self.root).as_posix()
                metadata = _hash_stable_file(path)
                prior = existing.get(relative)
                if prior is not None:
                    if (
                        int(prior["size_bytes"]) == int(metadata["size_bytes"])
                        and str(prior["sha256"]) == str(metadata["sha256"])
                    ):
                        continue
                    raise ProjectStoreError(f"已登记结果发生变化，请保存为新文件：{path.name}")
                added.append(
                    {
                        "id": uuid.uuid4().hex,
                        "batch_id": batch_id,
                        "category": CATEGORY_RESULTS,
                        "role": str(role)[:100],
                        "display_name": path.name,
                        "relative_path": relative,
                        "size_bytes": int(metadata["size_bytes"]),
                        "sha256": str(metadata["sha256"]),
                        "modified_ns": int(metadata["modified_ns"]),
                    }
                )
            manifest["files"] = [*_file_objects(manifest), *added]
            self._write_active_manifest(batch_id, manifest)
        ids = {str(item["id"]) for item in added}
        detail = self.get_batch(batch_id)
        assert detail is not None
        return tuple(item for item in detail.files if item.id in ids)

    def new_folder(self, parent: str | Path, name: str) -> Path:
        """Create a visible user folder without entering results or metadata."""

        self._require_writable()
        with self._mutex:
            clean_name = _visible_component(name)
            if clean_name.casefold() in {
                PROJECT_METADATA_DIR.casefold(),
                CATEGORY_VISIBLE_NAMES[CATEGORY_RESULTS].casefold(),
            }:
                raise ProjectStoreError("该文件夹名称属于项目保留名称，请换一个名称。")
            parent_path = self._visible_writable_directory(
                parent,
                create_declared_supplement=True,
            )
            destination = _project_join(parent_path, clean_name)
            if destination.exists() or _is_link_like(destination):
                raise ProjectStoreError(f"文件夹已存在：{clean_name}")
            destination.mkdir(mode=0o700, exist_ok=False)
            _make_private(destination, directory=True)
            return destination

    def import_common_sources(
        self,
        sources: Iterable[str | Path],
        *,
        subdirectory: str | Path | None = None,
        cancelled: Callable[[], bool] | None = None,
        progress: Callable[[int, int, str], None] | None = None,
        on_progress: Callable[[ImportProgress], None] | None = None,
    ) -> tuple[Path, ...]:
        self._require_writable()
        destination = self.workspace.common_root
        if subdirectory is not None and str(subdirectory).strip():
            relative = _validated_relative_path(str(subdirectory))
            _assert_no_link_components(destination, Path(relative))
            destination = _project_join(destination, relative)
            destination.mkdir(parents=True, exist_ok=True)
        return self.import_to_directory(
            destination,
            sources,
            cancelled=cancelled,
            progress=progress,
            on_progress=on_progress,
        )

    def import_to_directory(
        self,
        destination: str | Path,
        sources: Iterable[str | Path],
        *,
        cancelled: Callable[[], bool] | None = None,
        progress: Callable[[int, int, str], None] | None = None,
        on_progress: Callable[[ImportProgress], None] | None = None,
    ) -> tuple[Path, ...]:
        """Safely copy external material into an ordinary visible folder."""

        self._require_writable()
        with self._mutex:
            _report_import_progress(on_progress, ImportProgress(phase="checking"))
            _raise_if_cancelled(cancelled)
            destination_root = self._visible_writable_directory(destination)
            source_items = self._collect_external_sources(
                sources,
                cancelled=cancelled,
                on_progress=on_progress,
            )
            _raise_if_cancelled(cancelled)
            required_bytes = _source_total_bytes(source_items, cancelled=cancelled)
            _report_import_progress(
                on_progress,
                ImportProgress(
                    phase="checking",
                    files_scanned=len(source_items),
                    files_total=len(source_items),
                    bytes_total=required_bytes,
                ),
            )
            _raise_if_cancelled(cancelled)
            self._ensure_free_space(required_bytes)
            _raise_if_cancelled(cancelled)
            reserved_keys = _existing_name_keys(destination_root, cancelled=cancelled)
            assigned_top_paths: dict[int, Path] = {}
            planned: list[tuple[_ExternalSourceItem, Path]] = []
            for item in source_items:
                _raise_if_cancelled(cancelled)
                top_path = assigned_top_paths.get(item.source_index)
                if top_path is None:
                    top_path = _unique_destination(
                        _project_join(destination_root, _visible_component(item.top_name)),
                        reserved_keys,
                        destination_root,
                        is_directory=item.top_is_directory,
                        cancelled=cancelled,
                    )
                    assigned_top_paths[item.source_index] = top_path
                    reserved_keys.add(_portable_relative_key(top_path.relative_to(destination_root)))
                if item.top_is_directory:
                    desired = _project_join(
                        top_path,
                        Path(*(_visible_component(part) for part in item.relative_parts)),
                    )
                    final_path = _unique_destination(
                        desired,
                        reserved_keys,
                        destination_root,
                        cancelled=cancelled,
                    )
                else:
                    final_path = top_path
                reserved_keys.add(_portable_relative_key(final_path.relative_to(destination_root)))
                planned.append((item, final_path))

            operation_id = uuid.uuid4().hex
            staging_root = self.staging_dir / operation_id
            staging_files = staging_root / "files"
            staged: list[tuple[Path, Path, int, str]] = []
            copied_bytes = 0
            journal_written = False
            try:
                staging_root.mkdir(mode=0o700, exist_ok=False)
                _make_private(staging_root, directory=True)
                staging_files.mkdir(mode=0o700, exist_ok=False)
                _make_private(staging_files, directory=True)
                for index, (item, final_path) in enumerate(planned):
                    _raise_if_cancelled(cancelled)
                    staging_path = staging_files / f"{index:08d}.data"
                    metadata = _copy_external_file(
                        item.path,
                        staging_path,
                        cancelled=cancelled,
                        on_chunk=lambda size, name=item.path.name, completed=index: _report_copy_callbacks(
                            progress,
                            on_progress,
                            copied_bytes + size,
                            required_bytes,
                            name,
                            files_scanned=len(source_items),
                            files_completed=completed,
                            files_total=len(planned),
                        ),
                    )
                    copied_bytes += int(metadata["size_bytes"])
                    _report_copy_callbacks(
                        progress=None,
                        on_progress=on_progress,
                        copied_bytes=copied_bytes,
                        total_bytes=required_bytes,
                        name=item.path.name,
                        files_scanned=len(source_items),
                        files_completed=index + 1,
                        files_total=len(planned),
                    )
                    staged.append(
                        (
                            staging_path,
                            final_path,
                            int(metadata["size_bytes"]),
                            str(metadata["sha256"]),
                        )
                    )
                _raise_if_cancelled(cancelled)
                _report_import_progress(
                    on_progress,
                    ImportProgress(
                        phase="finalizing",
                        files_scanned=len(source_items),
                        files_completed=len(planned),
                        files_total=len(planned),
                        bytes_copied=copied_bytes,
                        bytes_total=required_bytes,
                    ),
                )
                journal = {
                    "version": VISIBLE_IMPORT_VERSION,
                    "kind": VISIBLE_IMPORT_KIND,
                    "project_id": self.workspace.project_id,
                    "operation_id": operation_id,
                    "state": "finalizing",
                    "created_at": _utc_now(),
                    "destination_relative": destination_root.relative_to(self.root).as_posix(),
                    "items": [
                        {
                            "staging_relative": staging_path.relative_to(staging_root).as_posix(),
                            "destination_relative": final_path.relative_to(self.root).as_posix(),
                            "size_bytes": size,
                            "sha256": digest,
                        }
                        for staging_path, final_path, size, digest in staged
                    ],
                }
                _write_json(staging_root / VISIBLE_IMPORT_JOURNAL, journal)
                journal_written = True
                self._recover_visible_directory_import(staging_root)
            except Exception:
                if not journal_written:
                    shutil.rmtree(staging_root, ignore_errors=True)
                raise
            return tuple(final_path for _staging_path, final_path, _size, _digest in staged)

    def _visible_writable_directory(
        self,
        value: str | Path,
        *,
        create_declared_supplement: bool = False,
    ) -> Path:
        raw = Path(value).expanduser()
        if raw.is_absolute():
            _assert_existing_ancestors_are_real(raw.absolute())
            path = raw.resolve()
        else:
            relative_input = Path(_validated_relative_path(str(raw)))
            _assert_no_link_components(self.root, relative_input)
            path = _project_join(self.root, relative_input)
        if not _is_inside(path, self.root):
            raise ProjectStoreError("只能写入当前项目中的普通目录。")
        relative = path.relative_to(self.root)
        if not relative.parts:
            raise ProjectStoreError("不能直接向项目根目录导入文件。")
        lowered_parts = {part.casefold() for part in relative.parts}
        if PROJECT_METADATA_DIR.casefold() in lowered_parts:
            raise ProjectStoreError("不能写入项目隐藏管理目录。")
        if CATEGORY_VISIBLE_NAMES[CATEGORY_RESULTS].casefold() in lowered_parts:
            raise ProjectStoreError("处理结果只能由工具生成并登记。")
        if not path.exists() and create_declared_supplement:
            if not self._is_declared_supplement_root(path):
                raise ProjectStoreError("只能创建清单中声明的补充资料目录。")
            _assert_no_link_components(self.root, relative.parent)
            path.mkdir(mode=0o700, exist_ok=False)
            _make_private(path, directory=True)
        _require_regular_directory(path, "目标文件夹")
        _assert_no_link_components(self.root, relative)
        return path

    def _is_declared_supplement_root(self, path: Path) -> bool:
        for manifest_path in sorted(self.manifest_dir.glob("*.json")):
            if _is_link_like(manifest_path) or not manifest_path.is_file():
                raise ProjectStoreError("批次清单目录包含不安全项目。")
            manifest = _read_json(manifest_path)
            summary = _summary_from_manifest(manifest)
            _validate_manifest_identity(manifest, summary.id)
            supplement_root = _project_join(
                self.root,
                _directory_map(manifest)[CATEGORY_SUPPLEMENTS],
            )
            if path == supplement_root:
                return True
        return False

    def list_batches(
        self,
        *,
        tool_id: str | None = None,
        group_name: str | None = None,
    ) -> tuple[BatchSummary, ...]:
        summaries: list[BatchSummary] = []
        if self.workspace.format_version > PROJECT_FORMAT_VERSION:
            return ()
        for path in sorted(self.manifest_dir.glob("*.json")):
            try:
                manifest = _read_json(path)
                summary = _summary_from_manifest(manifest)
            except (OSError, ValueError, TypeError, ProjectStoreError, json.JSONDecodeError):
                continue
            if tool_id is not None and summary.tool_id != tool_id:
                continue
            if group_name is not None and summary.group_name != group_name:
                continue
            summaries.append(summary)
        summaries.sort(key=lambda item: (item.started_at or item.created_at, item.id), reverse=True)
        return tuple(summaries)

    def get_batch(self, batch_id: str) -> BatchDetail | None:
        if self.workspace.format_version > PROJECT_FORMAT_VERSION:
            return None
        if not _is_uuid_hex(batch_id):
            raise ProjectStoreError("批次编号无效。")
        path = self._manifest_path(batch_id)
        if not path.is_file() or _is_link_like(path):
            return None
        manifest = _read_json(path)
        _validate_manifest_identity(manifest, batch_id)
        return self._detail_from_manifest(manifest)

    def _load_trash_manifest(self, batch_id: str) -> dict[str, Any]:
        clean_batch_id = str(batch_id).lower()
        if not _is_uuid_hex(clean_batch_id):
            raise ProjectStoreError("批次编号无效。")
        trash_entry = self.trash_dir / clean_batch_id
        _require_regular_directory(trash_entry, "回收站批次目录")
        manifest_path = trash_entry / "manifest.json"
        _require_regular_file(manifest_path, "回收站批次清单")
        manifest = _read_json(manifest_path)
        _validate_manifest_identity(manifest, clean_batch_id)
        summary = _summary_from_manifest(manifest)
        if summary.deleted_at is None:
            raise ProjectStoreError("回收站清单缺少移入时间。")
        _trash_original_directory_map(manifest, summary)
        return manifest

    def _trash_detail_from_manifest(self, manifest: dict[str, Any]) -> TrashBatchDetail:
        summary = _summary_from_manifest(manifest)
        original_directories = _trash_original_directory_map(manifest, summary)
        files = tuple(_project_file_from_dict(item) for item in _file_objects(manifest))
        counts = {
            category: sum(1 for item in files if item.category == category)
            for category in CATEGORIES
        }
        return TrashBatchDetail(
            summary=summary,
            original_relative_path=Path(original_directories[CATEGORY_UPLOADS]).parent.as_posix(),
            upload_count=counts[CATEGORY_UPLOADS],
            result_count=counts[CATEGORY_RESULTS],
            supplement_count=counts[CATEGORY_SUPPLEMENTS],
            total_size_bytes=sum(item.size_bytes for item in files),
        )

    def list_trash_details(self) -> tuple[TrashBatchDetail, ...]:
        if self.workspace.format_version > PROJECT_FORMAT_VERSION:
            return ()
        details: list[TrashBatchDetail] = []
        try:
            entries = sorted(self.trash_dir.iterdir()) if self.trash_dir.is_dir() else ()
        except OSError:
            return ()
        for entry in entries:
            try:
                manifest = self._load_trash_manifest(entry.name)
                details.append(self._trash_detail_from_manifest(manifest))
            except (OSError, ValueError, TypeError, ProjectStoreError, json.JSONDecodeError):
                continue
        details.sort(key=lambda item: (item.summary.deleted_at or "", item.summary.id), reverse=True)
        return tuple(details)

    def list_trash(self) -> tuple[BatchSummary, ...]:
        return tuple(item.summary for item in self.list_trash_details())

    def move_to_trash(self, batch_id: str) -> Path:
        self._require_writable()
        with self._mutex:
            manifest = self._load_active_manifest(batch_id)
            batch = _batch_object(manifest)
            if batch["status"] == "running":
                raise ProjectStoreError("正在处理的批次不能移到回收站。")
            trash_entry = self.trash_dir / batch_id
            if trash_entry.exists():
                raise ProjectStoreError("回收站中已存在同一批次。")
            manifest["pending_trash"] = {
                "trash_relative": trash_entry.relative_to(self.root).as_posix(),
                "original_directories": _directory_map(manifest),
                "created_at": _utc_now(),
            }
            self._write_active_manifest(batch_id, manifest)
            self._recover_pending_trash(batch_id, manifest)
            return trash_entry

    def restore_from_trash(self, batch_id: str) -> BatchDetail:
        self._require_writable()
        clean_batch_id = str(batch_id).lower()
        if not _is_uuid_hex(clean_batch_id):
            raise ProjectStoreError("批次编号无效。")
        with self._mutex:
            manifest = self._load_trash_manifest(clean_batch_id)
            if isinstance(manifest.get("pending_restore"), dict):
                self._recover_pending_restore(clean_batch_id, manifest)
                detail = self.get_batch(clean_batch_id)
                assert detail is not None
                return detail
            self._verify_manifest_files_at_directories(manifest, _directory_map(manifest))
            batch = _batch_object(manifest)
            target_name, target_directories = self._unique_restore_directories(
                str(batch["group_name"]),
                str(batch["tool_name"]),
                str(batch["directory_name"]),
            )
            manifest["pending_restore"] = {
                "target_directories": target_directories,
                "directory_name": target_name,
                "created_at": _utc_now(),
            }
            _write_json(self.trash_dir / clean_batch_id / "manifest.json", manifest)
            self._recover_pending_restore(clean_batch_id, manifest, files_verified=True)
        detail = self.get_batch(clean_batch_id)
        assert detail is not None
        return detail

    def verify_batch_files(self, batch_id: str) -> bool:
        manifest = self._load_active_manifest(batch_id)
        self._verify_manifest_files(manifest)
        return True

    def integrity_check(self) -> bool:
        if self.workspace.format_version > PROJECT_FORMAT_VERSION:
            return False
        try:
            marker = _read_json(self.metadata_dir / PROJECT_FILE_NAME)
            if _required_uuid(marker, "project_id") != self.workspace.project_id:
                return False
            for summary in self.list_batches():
                manifest = self._load_active_manifest(summary.id)
                self._verify_manifest_files(manifest)
                if summary.status in TERMINAL_BATCH_STATUSES:
                    self._assert_no_unregistered_results(manifest)
            for summary in self.list_trash():
                manifest_path = self.trash_dir / summary.id / "manifest.json"
                manifest = _read_json(manifest_path)
                _summary_from_manifest(manifest)
                self._verify_manifest_files(manifest)
            return True
        except (OSError, ValueError, TypeError, ProjectStoreError, json.JSONDecodeError):
            return False

    def refresh(self) -> tuple[BatchSummary, ...]:
        if self.writable:
            with self._mutex:
                self._recover_workspace()
        return self.list_batches()

    def _batch_relative_directories(self, group: str, tool_name: str, batch_name: str) -> dict[str, str]:
        group_part = _business_component(group, "业务分组")
        tool_part = _business_component(tool_name, "工具")
        batch_part = _visible_component(batch_name)
        batch_root = Path(group_part) / tool_part / batch_part
        return {
            category: (batch_root / visible).as_posix()
            for category, visible in CATEGORY_VISIBLE_NAMES.items()
        }

    def _unique_batch_directories(
        self,
        group: str,
        tool_name: str,
        desired_name: str,
    ) -> tuple[str, dict[str, str]]:
        for index in range(1, 10_000):
            name = desired_name if index == 1 else f"{desired_name}_{index}"
            directories = self._batch_relative_directories(group, tool_name, name)
            _assert_no_link_components(
                self.root,
                Path(directories[CATEGORY_UPLOADS]).parent,
            )
            batch_root = _project_join(self.root, directories[CATEGORY_UPLOADS]).parent
            if not batch_root.exists() and not _is_link_like(batch_root):
                return name, directories
        raise ProjectStoreError("同名批次过多，请调整业务说明或期间。")

    def _unique_restore_directories(
        self,
        group: str,
        tool_name: str,
        desired_name: str,
    ) -> tuple[str, dict[str, str]]:
        """Choose a human-readable restore name without changing run naming."""

        for index in range(1, 10_000):
            name = desired_name if index == 1 else f"{desired_name} ({index})"
            directories = self._batch_relative_directories(group, tool_name, name)
            _assert_no_link_components(
                self.root,
                Path(directories[CATEGORY_UPLOADS]).parent,
            )
            batch_root = _project_join(self.root, directories[CATEGORY_UPLOADS]).parent
            if not batch_root.exists() and not _is_link_like(batch_root):
                return name, directories
        raise ProjectStoreError("同名恢复记录过多，请先整理该功能下的项目文件。")

    def _collect_external_sources(
        self,
        sources: Iterable[str | Path],
        *,
        cancelled: Callable[[], bool] | None = None,
        on_progress: Callable[[ImportProgress], None] | None = None,
    ) -> list[_ExternalSourceItem]:
        items: list[_ExternalSourceItem] = []
        raw_sources: list[Path] = []
        for item in sources:
            _raise_if_cancelled(cancelled)
            raw_sources.append(Path(item).expanduser())
        if not raw_sources:
            raise ProjectStoreError("请选择要导入的文件或文件夹。")
        for source_index, source in enumerate(raw_sources):
            _raise_if_cancelled(cancelled)
            _assert_existing_ancestors_are_real(
                source.absolute(),
                allow_macos_root_aliases=True,
            )
            if _is_link_like(source) or not source.exists():
                raise ProjectStoreError(f"来源不存在或是链接：{source.name or source}")
            resolved = source.resolve()
            if _paths_overlap(resolved, self.root):
                raise ProjectStoreError("不能把项目自身或包含项目的文件夹再次导入。")
            if resolved.is_file():
                _require_regular_file(resolved, "导入文件")
                if _is_ignored_import_file(resolved):
                    continue
                _reject_forbidden_import_file(resolved)
                items.append(
                    _ExternalSourceItem(
                        path=resolved,
                        source_index=source_index,
                        top_name=resolved.name,
                        top_is_directory=False,
                        relative_parts=(),
                    )
                )
                _report_checking_progress(on_progress, items, resolved.name)
                _raise_if_cancelled(cancelled)
            elif resolved.is_dir():
                for child, relative_parts in _walk_external_directory(
                    resolved,
                    cancelled=cancelled,
                ):
                    _raise_if_cancelled(cancelled)
                    items.append(
                        _ExternalSourceItem(
                            path=child,
                            source_index=source_index,
                            top_name=resolved.name,
                            top_is_directory=True,
                            relative_parts=relative_parts,
                        )
                    )
                    _report_checking_progress(on_progress, items, child.name)
                    _raise_if_cancelled(cancelled)
            else:
                raise ProjectStoreError(f"来源不是普通文件或文件夹：{resolved.name}")
        if not items:
            raise ProjectStoreError("所选文件夹内没有可导入的资料。")
        return items

    def _collect_project_sources(
        self,
        sources: Iterable[str | Path],
        *,
        target_batch_root: Path,
        cancelled: Callable[[], bool] | None = None,
        on_progress: Callable[[ImportProgress], None] | None = None,
    ) -> list[_ExternalSourceItem]:
        """Collect safe, visible project material for an independent snapshot."""

        raw_sources: list[Path] = []
        for item in sources:
            _raise_if_cancelled(cancelled)
            source = Path(item).expanduser()
            raw_sources.append(source if source.is_absolute() else self.root / source)
        if not raw_sources:
            raise ProjectStoreError("请选择要复用的项目文件或文件夹。")

        source_roots: list[tuple[Path, str, dict[str, Any] | None]] = [
            (self.workspace.common_root, "common", None)
        ]
        for manifest_path in sorted(self.manifest_dir.glob("*.json")):
            _raise_if_cancelled(cancelled)
            if _is_link_like(manifest_path) or not manifest_path.is_file():
                raise ProjectStoreError("批次清单目录包含不安全项目。")
            manifest = _read_json(manifest_path)
            summary = _summary_from_manifest(manifest)
            _validate_manifest_identity(manifest, summary.id)
            directories = _directory_map(manifest)
            source_roots.extend(
                (
                    (_project_join(self.root, directories[CATEGORY_UPLOADS]), CATEGORY_UPLOADS, manifest),
                    (
                        _project_join(self.root, directories[CATEGORY_SUPPLEMENTS]),
                        CATEGORY_SUPPLEMENTS,
                        manifest,
                    ),
                    (_project_join(self.root, directories[CATEGORY_RESULTS]), CATEGORY_RESULTS, manifest),
                )
            )

        items: list[_ExternalSourceItem] = []
        for source_index, source in enumerate(raw_sources):
            _raise_if_cancelled(cancelled)
            _assert_existing_ancestors_are_real(source.absolute())
            if _is_link_like(source) or not source.exists():
                raise ProjectStoreError(f"项目来源不存在或是链接：{source.name or source}")
            resolved = source.resolve()
            if not _is_inside(resolved, self.root):
                raise ProjectStoreError("此入口只能复用当前项目中的资料。")
            relative = resolved.relative_to(self.root)
            if any(part.casefold() == PROJECT_METADATA_DIR.casefold() for part in relative.parts):
                raise ProjectStoreError("项目隐藏管理目录不能作为处理资料。")
            _assert_no_link_components(self.root, relative)
            if _paths_overlap(resolved, target_batch_root):
                raise ProjectStoreError("不能把当前目标批次自身作为输入资料。")

            matches = [rule for rule in source_roots if _is_inside(resolved, rule[0])]
            if not matches:
                raise ProjectStoreError("只能复用共用资料或批次中的上传、补充、已完成结果。")
            source_root, source_kind, source_manifest = max(matches, key=lambda rule: len(rule[0].parts))
            registered_by_path: dict[Path, dict[str, Any]] = {}
            if source_manifest is not None:
                for registered_item in _file_objects(source_manifest):
                    _raise_if_cancelled(cancelled)
                    if str(registered_item.get("category")) != source_kind:
                        continue
                    registered_path = _project_join(
                        self.root,
                        str(registered_item["relative_path"]),
                    )
                    registered_by_path[registered_path] = registered_item
            if source_kind == CATEGORY_RESULTS:
                assert source_manifest is not None
                source_summary = _summary_from_manifest(source_manifest)
                if source_summary.status != "success":
                    raise ProjectStoreError("只有已成功完成批次的处理结果可以复用。")

            if resolved.is_file():
                _require_regular_file(resolved, "项目来源文件")
                if _is_ignored_import_file(resolved):
                    continue
                _reject_forbidden_import_file(resolved)
                registered_item = registered_by_path.get(resolved)
                if source_manifest is not None and registered_item is None:
                    raise ProjectStoreError("只能复用清单中已登记且未改动的项目文件。")
                items.append(
                    _ExternalSourceItem(
                        path=resolved,
                        source_index=source_index,
                        top_name=resolved.name,
                        top_is_directory=False,
                        relative_parts=(),
                        expected_size_bytes=(
                            int(registered_item["size_bytes"])
                            if registered_item is not None
                            else None
                        ),
                        expected_sha256=(
                            str(registered_item["sha256"])
                            if registered_item is not None
                            else None
                        ),
                    )
                )
                _report_checking_progress(on_progress, items, resolved.name)
                _raise_if_cancelled(cancelled)
            elif resolved.is_dir():
                _require_regular_directory(source_root, "项目资料目录")
                walked = list(
                    _walk_external_directory(
                        resolved,
                        cancelled=cancelled,
                    )
                )
                if source_manifest is not None:
                    actual_paths: set[Path] = set()
                    for child, _relative_parts in walked:
                        _raise_if_cancelled(cancelled)
                        actual_paths.add(child)
                    registered_paths: set[Path] = set()
                    for path in registered_by_path:
                        _raise_if_cancelled(cancelled)
                        if _is_inside(path, resolved):
                            registered_paths.add(path)
                    if actual_paths != registered_paths:
                        raise ProjectStoreError("项目资料与原清单不一致，不能作为新批次输入。")
                for child, relative_parts in walked:
                    _raise_if_cancelled(cancelled)
                    if any(
                        part.casefold() == PROJECT_METADATA_DIR.casefold()
                        for part in relative_parts
                    ):
                        raise ProjectStoreError("项目隐藏管理目录不能作为处理资料。")
                    registered_item = registered_by_path.get(child)
                    items.append(
                        _ExternalSourceItem(
                            path=child,
                            source_index=source_index,
                            top_name=resolved.name,
                            top_is_directory=True,
                            relative_parts=relative_parts,
                            expected_size_bytes=(
                                int(registered_item["size_bytes"])
                                if registered_item is not None
                                else None
                            ),
                            expected_sha256=(
                                str(registered_item["sha256"])
                                if registered_item is not None
                                else None
                            ),
                        )
                    )
                    _report_checking_progress(on_progress, items, child.name)
                    _raise_if_cancelled(cancelled)
            else:
                raise ProjectStoreError(f"项目来源不是普通文件或文件夹：{resolved.name}")
        if not items:
            raise ProjectStoreError("所选项目资料中没有可复用的文件。")
        return items

    def _ensure_free_space(self, required_bytes: int) -> None:
        reserve = max(MIN_FREE_SPACE_BYTES, int(required_bytes * PROJECT_COPY_RESERVE_RATIO))
        try:
            free = shutil.disk_usage(self.root).free
        except OSError as exc:
            raise ProjectStoreError(f"无法检查项目剩余空间：{exc}") from exc
        if free < required_bytes + reserve:
            need_gb = (required_bytes + reserve) / 1024 / 1024 / 1024
            free_gb = free / 1024 / 1024 / 1024
            raise ProjectStoreError(f"项目空间不足，需要约 {need_gb:.1f} GB，当前可用 {free_gb:.1f} GB。")

    def _manifest_path(self, batch_id: str) -> Path:
        if not _is_uuid_hex(batch_id):
            raise ProjectStoreError("批次编号无效。")
        return self.manifest_dir / f"{batch_id.lower()}.json"

    def _load_active_manifest(self, batch_id: str) -> dict[str, Any]:
        path = self._manifest_path(batch_id)
        _require_regular_file(path, "批次清单")
        manifest = _read_json(path)
        if _summary_from_manifest(manifest).id != batch_id.lower():
            raise ProjectStoreError("批次清单与编号不一致。")
        _validate_manifest_identity(manifest, batch_id)
        return manifest

    def _write_active_manifest(self, batch_id: str, manifest: dict[str, Any]) -> None:
        _validate_manifest_identity(manifest, batch_id)
        _write_json(self._manifest_path(batch_id), manifest)

    def _detail_from_manifest(self, manifest: dict[str, Any]) -> BatchDetail:
        summary = _summary_from_manifest(manifest)
        directories = {
            category: _project_join(self.root, relative)
            for category, relative in _directory_map(manifest).items()
        }
        files = tuple(_project_file_from_dict(item) for item in _file_objects(manifest))
        return BatchDetail(summary=summary, directories=directories, files=files)

    def _verify_manifest_files(self, manifest: dict[str, Any]) -> None:
        directories = _directory_map(manifest)
        for item in _file_objects(manifest):
            category = str(item["category"])
            if category not in CATEGORIES:
                raise ProjectStoreError("批次文件分类无效。")
            path = _project_join(self.root, str(item["relative_path"]))
            expected_root = _project_join(self.root, directories[category])
            if not _is_inside(path, expected_root) or _is_link_like(path) or not path.is_file():
                raise ProjectStoreError(f"批次文件已移动或不安全：{item.get('display_name', path.name)}")
            if path.stat().st_size != int(item["size_bytes"]) or _sha256_file(path) != str(item["sha256"]):
                raise ProjectStoreError(f"批次文件已发生变化：{item.get('display_name', path.name)}")

    def _verify_manifest_files_at_directories(
        self,
        manifest: dict[str, Any],
        actual_directories: dict[str, str],
    ) -> None:
        """Verify registered recycle-bin files and reject undeclared additions."""

        declared_directories = _directory_map(manifest)
        checked_directories = _validated_directory_map(
            actual_directories,
            allow_hidden=any(
                Path(relative).parts[:1] == (PROJECT_METADATA_DIR,)
                for relative in actual_directories.values()
            ),
        )
        expected_paths: set[Path] = set()
        for raw_item in _file_objects(manifest):
            item = _project_file_from_dict(raw_item)
            declared_root = Path(declared_directories[item.category])
            try:
                inside = Path(item.relative_path).relative_to(declared_root)
            except ValueError as exc:
                raise ProjectStoreError("项目文件路径不属于批次目录。") from exc
            destination_root = _project_join(self.root, checked_directories[item.category])
            path = _project_join(destination_root, inside)
            if path in expected_paths:
                raise ProjectStoreError("回收站清单包含重复的登记文件。")
            expected_paths.add(path)
            if _is_link_like(path) or not path.is_file():
                raise ProjectStoreError(f"批次文件已移动或不安全：{item.display_name}")
            if path.stat().st_size != item.size_bytes or _sha256_file(path) != item.sha256:
                raise ProjectStoreError(f"批次文件已发生变化：{item.display_name}")

        batch_root = _project_join(self.root, checked_directories[CATEGORY_UPLOADS]).parent
        _require_regular_directory(batch_root, "回收站批次资料")
        actual_paths = {
            path
            for path, _relative_parts in _walk_directory_strict(
                batch_root,
                ignore_temporary=False,
            )
        }
        undeclared = sorted(actual_paths - expected_paths, key=lambda path: path.as_posix().casefold())
        if undeclared:
            names = "、".join(path.name for path in undeclared[:3])
            suffix = "等" if len(undeclared) > 3 else ""
            raise ProjectStoreError(f"回收站资料中存在未登记文件：{names}{suffix}")

    def _unregistered_result_files(self, manifest: dict[str, Any]) -> list[Path]:
        result_root = _project_join(self.root, _directory_map(manifest)[CATEGORY_RESULTS])
        if not result_root.exists():
            return []
        _require_regular_directory(result_root, "处理结果目录")
        registered = {
            _project_join(self.root, str(item["relative_path"]))
            for item in _file_objects(manifest)
            if str(item.get("category")) == CATEGORY_RESULTS
        }
        return [path for path, _parts in _walk_directory_strict(result_root, ignore_temporary=False) if path not in registered]

    def _assert_no_unregistered_results(self, manifest: dict[str, Any]) -> None:
        unregistered = self._unregistered_result_files(manifest)
        if unregistered:
            names = "、".join(path.name for path in unregistered[:3])
            suffix = "等" if len(unregistered) > 3 else ""
            raise ProjectStoreError(f"处理结果中仍有未登记文件：{names}{suffix}")

    def _quarantine_unregistered_results(self, manifest: dict[str, Any], *, reason: str) -> None:
        unregistered = self._unregistered_result_files(manifest)
        if not unregistered:
            return
        batch_id = _summary_from_manifest(manifest).id
        result_root = _project_join(self.root, _directory_map(manifest)[CATEGORY_RESULTS])
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        quarantine_root = self.quarantine_dir / batch_id / stamp
        moved: list[dict[str, str]] = []
        for source in unregistered:
            relative = source.relative_to(result_root)
            destination = quarantine_root / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            if destination.exists() or _is_link_like(destination):
                raise ProjectStoreError(f"未完成结果隔离位置冲突：{source.name}")
            try:
                source.rename(destination)
            except OSError as exc:
                raise ProjectStoreError(f"无法安全隔离未完成结果，请关闭文件后重试：{source.name}") from exc
            moved.append(
                {
                    "from": source.relative_to(self.root).as_posix(),
                    "to": destination.relative_to(self.root).as_posix(),
                }
            )
        _remove_empty_directories(result_root)
        notes = manifest.setdefault("recovery_notes", [])
        if isinstance(notes, list):
            notes.append(
                {
                    "at": _utc_now(),
                    "action": "quarantine_unregistered_results",
                    "reason": reason,
                    "files": moved,
                }
            )

    def _recover_workspace(self) -> None:
        referenced_staging: set[str] = set()
        for path in sorted(self.manifest_dir.glob("*.json")):
            if _is_link_like(path) or not path.is_file():
                raise ProjectStoreError("批次清单目录包含不安全项目。")
            manifest = _read_json(path)
            summary = _summary_from_manifest(manifest)
            batch_id = summary.id
            if manifest.get("pending_rename"):
                self._recover_pending_rename(batch_id, manifest)
                manifest = self._load_active_manifest(batch_id)
            if manifest.get("pending_import"):
                operation = manifest["pending_import"]
                if isinstance(operation, dict):
                    referenced_staging.add(str(operation.get("operation_id") or ""))
                self._recover_pending_import(batch_id, manifest)
                manifest = self._load_active_manifest(batch_id)
            if manifest.get("pending_trash"):
                self._recover_pending_trash(batch_id, manifest)
                continue
            batch = _batch_object(manifest)
            if batch["status"] == "running":
                self._quarantine_unregistered_results(manifest, reason="startup_recovery")
                batch["status"] = "stopped"
                batch["finished_at"] = _utc_now()
                batch["error_message"] = "程序上次退出，本次处理未正常完成。"
                batch["writer_session"] = None
                self._write_active_manifest(batch_id, manifest)

        for entry in sorted(self.trash_dir.iterdir()):
            if _is_link_like(entry) or not entry.is_dir():
                raise ProjectStoreError("项目回收站包含不安全项目。")
            manifest_path = entry / "manifest.json"
            if not manifest_path.exists():
                continue
            manifest = _read_json(manifest_path)
            if manifest.get("pending_restore"):
                self._recover_pending_restore(entry.name, manifest)

        for path in sorted(self.staging_dir.iterdir()):
            if path.name not in referenced_staging:
                if _is_link_like(path) or not path.is_dir():
                    raise ProjectStoreError("项目临时目录包含不安全项目。")
                journal_path = path / VISIBLE_IMPORT_JOURNAL
                if journal_path.exists() or _is_link_like(journal_path):
                    self._recover_visible_directory_import(path)
                else:
                    shutil.rmtree(path)

    def _recover_pending_import(self, batch_id: str, manifest: dict[str, Any]) -> None:
        pending = manifest.get("pending_import")
        if not isinstance(pending, dict):
            return
        operation_id = str(pending.get("operation_id") or "")
        if not _is_uuid_hex(operation_id):
            raise ProjectStoreError("导入恢复记录无效。")
        raw_items = pending.get("items")
        if not isinstance(raw_items, list):
            raise ProjectStoreError("导入恢复文件列表无效。")
        published: list[Path] = []
        try:
            for item in raw_items:
                if not isinstance(item, dict):
                    raise ProjectStoreError("导入恢复文件信息无效。")
                if _required_uuid(item, "batch_id") != batch_id or not _is_uuid_hex(str(item.get("id") or "")):
                    raise ProjectStoreError("导入恢复文件与批次不一致。")
                category = str(item.get("category") or "")
                if category not in {CATEGORY_UPLOADS, CATEGORY_SUPPLEMENTS}:
                    raise ProjectStoreError("导入恢复文件分类无效。")
                destination = _project_join(self.root, str(item["relative_path"]))
                expected_root = _project_join(self.root, _directory_map(manifest)[category])
                if not _is_inside(destination, expected_root) or destination == expected_root:
                    raise ProjectStoreError("导入恢复目标不属于当前批次。")
                staging = _project_join(self.root, str(item["staging_path"]))
                if not _is_inside(staging, self.staging_dir / operation_id):
                    raise ProjectStoreError("导入临时路径越界。")
                expected_size = int(item["size_bytes"])
                expected_hash = str(item["sha256"])
                if destination.exists():
                    if (
                        _is_link_like(destination)
                        or not destination.is_file()
                        or destination.stat().st_size != expected_size
                        or _sha256_file(destination) != expected_hash
                    ):
                        raise ProjectStoreError(f"导入目标发生冲突：{destination.name}")
                else:
                    if (
                        _is_link_like(staging)
                        or not staging.is_file()
                        or staging.stat().st_size != expected_size
                        or _sha256_file(staging) != expected_hash
                    ):
                        raise ProjectStoreError(f"导入临时文件不完整：{destination.name}")
                    _assert_no_link_components(self.root, destination.parent.relative_to(self.root))
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    _publish_no_replace(staging, destination)
                    _make_private(destination)
                    published.append(destination)
            known = {str(item["id"]) for item in _file_objects(manifest)}
            manifest["files"] = [
                *_file_objects(manifest),
                *(
                    {key: value for key, value in item.items() if key != "staging_path"}
                    for item in raw_items
                    if str(item["id"]) not in known
                ),
            ]
            manifest["pending_import"] = None
            self._write_active_manifest(batch_id, manifest)
            shutil.rmtree(self.staging_dir / operation_id, ignore_errors=True)
        except Exception:
            # Keep a complete, declared operation for the next open.  Files are
            # never reported in ``files`` until the manifest finalization above.
            raise

    def _recover_visible_directory_import(self, staging_root: Path) -> None:
        """Idempotently finish a journalled import into an ordinary visible folder."""

        if (
            staging_root.parent != self.staging_dir
            or not _is_uuid_hex(staging_root.name)
            or not _is_inside(staging_root, self.staging_dir)
        ):
            raise ProjectStoreError("普通目录导入临时位置无效。")
        _require_regular_directory(staging_root, "普通目录导入临时目录")
        journal_path = staging_root / VISIBLE_IMPORT_JOURNAL
        _require_regular_file(journal_path, "普通目录导入恢复记录")
        journal = _read_json(journal_path)
        try:
            version = int(journal.get("version"))
        except (TypeError, ValueError) as exc:
            raise ProjectStoreError("普通目录导入恢复版本无效。") from exc
        if version != VISIBLE_IMPORT_VERSION or str(journal.get("kind")) != VISIBLE_IMPORT_KIND:
            raise ProjectStoreError("普通目录导入恢复记录格式不受支持。")
        if _required_uuid(journal, "project_id") != self.workspace.project_id:
            raise ProjectStoreError("普通目录导入恢复记录不属于当前项目。")
        operation_id = _required_uuid(journal, "operation_id")
        if operation_id != staging_root.name:
            raise ProjectStoreError("普通目录导入恢复编号不一致。")
        if str(journal.get("state")) != "finalizing":
            raise ProjectStoreError("普通目录导入恢复状态无效。")

        destination_root = _project_join(self.root, str(journal.get("destination_relative") or ""))
        destination_root = self._visible_writable_directory(destination_root)
        raw_items = journal.get("items")
        if not isinstance(raw_items, list) or not raw_items:
            raise ProjectStoreError("普通目录导入恢复文件列表无效。")

        seen_staging: set[str] = set()
        seen_destinations: set[str] = set()
        validated: list[tuple[Path, Path, int, str]] = []
        staging_files_root = staging_root / "files"
        _require_regular_directory(staging_files_root, "普通目录导入临时文件目录")
        for raw_item in raw_items:
            if not isinstance(raw_item, dict):
                raise ProjectStoreError("普通目录导入恢复文件信息无效。")
            staging_relative = _validated_relative_path(str(raw_item.get("staging_relative") or ""))
            destination_relative = _validated_relative_path(
                str(raw_item.get("destination_relative") or "")
            )
            if staging_relative in seen_staging or destination_relative in seen_destinations:
                raise ProjectStoreError("普通目录导入恢复记录包含重复文件。")
            seen_staging.add(staging_relative)
            seen_destinations.add(destination_relative)
            staging = _project_join(staging_root, staging_relative)
            destination = _project_join(self.root, destination_relative)
            if not _is_inside(staging, staging_files_root):
                raise ProjectStoreError("普通目录导入临时文件路径越界。")
            if destination == destination_root or not _is_inside(destination, destination_root):
                raise ProjectStoreError("普通目录导入恢复目标越界。")
            try:
                expected_size = int(raw_item["size_bytes"])
            except (KeyError, TypeError, ValueError) as exc:
                raise ProjectStoreError("普通目录导入恢复文件大小无效。") from exc
            expected_hash = str(raw_item.get("sha256") or "").lower()
            if (
                expected_size < 0
                or len(expected_hash) != 64
                or any(character not in "0123456789abcdef" for character in expected_hash)
            ):
                raise ProjectStoreError("普通目录导入恢复校验信息无效。")
            validated.append((staging, destination, expected_size, expected_hash))

        for staging, destination, expected_size, expected_hash in validated:
            if destination.exists() or _is_link_like(destination):
                if (
                    _is_link_like(destination)
                    or not destination.is_file()
                    or destination.stat().st_size != expected_size
                    or _sha256_file(destination) != expected_hash
                ):
                    raise ProjectStoreError(f"导入目标发生冲突，未覆盖：{destination.name}")
                continue
            if (
                _is_link_like(staging)
                or not staging.is_file()
                or staging.stat().st_size != expected_size
                or _sha256_file(staging) != expected_hash
            ):
                raise ProjectStoreError(f"普通目录导入临时文件不完整：{destination.name}")
            destination_parent_relative = destination.parent.relative_to(self.root)
            _assert_no_link_components(self.root, destination_parent_relative)
            destination.parent.mkdir(parents=True, exist_ok=True)
            _assert_no_link_components(self.root, destination_parent_relative)
            _publish_no_replace(staging, destination)
            _make_private(destination)

        shutil.rmtree(staging_root)

    def _recover_pending_rename(self, batch_id: str, manifest: dict[str, Any]) -> None:
        pending = manifest.get("pending_rename")
        if not isinstance(pending, dict):
            return
        old_directories = _validated_directory_map(pending.get("old_directories"))
        new_directories = _validated_directory_map(pending.get("new_directories"))
        old_batch_root = _project_join(self.root, old_directories[CATEGORY_UPLOADS]).parent
        new_batch_root = _project_join(self.root, new_directories[CATEGORY_UPLOADS]).parent
        old_exists = old_batch_root.is_dir() and not _is_link_like(old_batch_root)
        new_exists = new_batch_root.is_dir() and not _is_link_like(new_batch_root)
        if old_exists and new_exists:
            raise ProjectStoreError("批次改名恢复发现两个目录，已停止写入。")
        if old_exists:
            _assert_no_link_components(self.root, new_batch_root.parent.relative_to(self.root))
            new_batch_root.parent.mkdir(parents=True, exist_ok=True)
            old_batch_root.rename(new_batch_root)
        elif not new_exists:
            raise ProjectStoreError("批次改名恢复找不到原资料目录。")
        for item in _file_objects(manifest):
            category = str(item["category"])
            old_prefix = old_directories[category]
            new_prefix = new_directories[category]
            relative = str(item["relative_path"])
            if relative == old_prefix or relative.startswith(old_prefix + "/"):
                item["relative_path"] = new_prefix + relative[len(old_prefix) :]
            elif not (relative == new_prefix or relative.startswith(new_prefix + "/")):
                raise ProjectStoreError("批次文件路径与待恢复改名不一致。")
        batch = _batch_object(manifest)
        batch["status"] = "running"
        batch["directory_name"] = str(pending["directory_name"])
        batch["business_description"] = str(pending["business_description"])
        batch["business_period"] = str(pending["business_period"])
        batch["started_at"] = str(pending["started_at"])
        batch["writer_session"] = str(pending.get("writer_session") or self._session_id)
        manifest["directories"] = new_directories
        manifest["pending_rename"] = None
        self._write_active_manifest(batch_id, manifest)

    def _recover_pending_trash(self, batch_id: str, manifest: dict[str, Any]) -> None:
        pending = manifest.get("pending_trash")
        if not isinstance(pending, dict):
            return
        trash_entry = _project_join(self.root, str(pending["trash_relative"]))
        if trash_entry.parent != self.trash_dir or trash_entry.name != batch_id:
            raise ProjectStoreError("回收站恢复路径越界。")
        trash_entry.mkdir(mode=0o700, exist_ok=True)
        _make_private(trash_entry, directory=True)
        directories = _validated_directory_map(pending.get("original_directories"))
        source_batch_root = _project_join(self.root, directories[CATEGORY_UPLOADS]).parent
        trash_batch_root = trash_entry / "batch"
        source_exists = source_batch_root.is_dir() and not _is_link_like(source_batch_root)
        destination_exists = trash_batch_root.is_dir() and not _is_link_like(trash_batch_root)
        if source_exists and destination_exists:
            raise ProjectStoreError("移入回收站时发现两个资料目录。")
        if source_exists:
            source_batch_root.rename(trash_batch_root)
        elif not destination_exists:
            raise ProjectStoreError("待移入回收站的批次目录不存在。")
        trash_directories = {
            category: (trash_batch_root / CATEGORY_VISIBLE_NAMES[category]).relative_to(self.root).as_posix()
            for category in CATEGORIES
        }
        for item in _file_objects(manifest):
            category = str(item["category"])
            old_prefix = directories[category]
            new_prefix = trash_directories[category]
            relative = str(item["relative_path"])
            if relative == old_prefix or relative.startswith(old_prefix + "/"):
                item["relative_path"] = new_prefix + relative[len(old_prefix) :]
        batch = _batch_object(manifest)
        batch["deleted_at"] = _utc_now()
        batch["writer_session"] = None
        manifest["original_directories"] = directories
        manifest["directories"] = trash_directories
        manifest["pending_trash"] = None
        _write_json(trash_entry / "manifest.json", manifest)
        self._manifest_path(batch_id).unlink(missing_ok=True)

    def _recover_pending_restore(
        self,
        batch_id: str,
        manifest: dict[str, Any],
        *,
        files_verified: bool = False,
    ) -> None:
        pending = manifest.get("pending_restore")
        if not isinstance(pending, dict):
            return
        clean_batch_id = str(batch_id).lower()
        if not _is_uuid_hex(clean_batch_id):
            raise ProjectStoreError("批次编号无效。")
        _validate_manifest_identity(manifest, clean_batch_id)
        source_directories = _directory_map(manifest)
        target_directories = _validated_directory_map(pending.get("target_directories"))
        target_name = _visible_component(str(pending.get("directory_name") or ""))
        if Path(target_directories[CATEGORY_UPLOADS]).parts[2] != target_name:
            raise ProjectStoreError("回收站恢复名称与目标位置不一致。")
        source_batch_root = _project_join(self.root, source_directories[CATEGORY_UPLOADS]).parent
        target_batch_root = _project_join(self.root, target_directories[CATEGORY_UPLOADS]).parent
        source_exists = source_batch_root.is_dir() and not _is_link_like(source_batch_root)
        destination_exists = target_batch_root.is_dir() and not _is_link_like(target_batch_root)
        if source_exists and destination_exists:
            raise ProjectStoreError("恢复批次时发现两个资料目录。")
        if not source_exists and not destination_exists:
            raise ProjectStoreError("回收站中的批次资料不存在。")
        if not files_verified:
            self._verify_manifest_files_at_directories(
                manifest,
                source_directories if source_exists else target_directories,
            )
        if source_exists:
            _assert_no_link_components(self.root, target_batch_root.parent.relative_to(self.root))
            target_batch_root.parent.mkdir(parents=True, exist_ok=True)
            source_batch_root.rename(target_batch_root)
        for item in _file_objects(manifest):
            category = str(item["category"])
            old_prefix = source_directories[category]
            new_prefix = target_directories[category]
            relative = str(item["relative_path"])
            if relative == old_prefix or relative.startswith(old_prefix + "/"):
                item["relative_path"] = new_prefix + relative[len(old_prefix) :]
        batch = _batch_object(manifest)
        batch["directory_name"] = target_name
        batch["deleted_at"] = None
        manifest["directories"] = target_directories
        manifest["pending_restore"] = None
        manifest.pop("original_directories", None)
        active_path = self._manifest_path(clean_batch_id)
        if active_path.exists():
            existing = _read_json(active_path)
            _validate_manifest_identity(existing, clean_batch_id)
            if _directory_map(existing) != target_directories:
                raise ProjectStoreError("活动项目中已存在位置不同的同一批次。")
        else:
            _write_json(active_path, manifest)
        trash_entry = self.trash_dir / clean_batch_id
        (trash_entry / "manifest.json").unlink(missing_ok=True)
        try:
            trash_entry.rmdir()
        except OSError:
            pass


def _summary_from_manifest(manifest: dict[str, Any]) -> BatchSummary:
    if not isinstance(manifest, dict) or int(manifest.get("format_version", 0)) != PROJECT_FORMAT_VERSION:
        raise ProjectStoreError("批次清单版本无效。")
    batch = _batch_object(manifest)
    batch_id = _required_uuid(batch, "id")
    status = str(batch.get("status") or "")
    if status not in BATCH_STATUSES:
        raise ProjectStoreError("批次状态无效。")
    directories = _validated_directory_map(
        manifest.get("directories"),
        allow_hidden=bool(batch.get("deleted_at")) or bool(manifest.get("pending_restore")),
    )
    directory_parts = Path(directories[CATEGORY_UPLOADS]).parts
    if directory_parts[0] == PROJECT_METADATA_DIR:
        if directory_parts[2] != batch_id or directory_parts[3] != "batch":
            raise ProjectStoreError("回收站批次目录与编号不一致。")
    else:
        expected_prefix = (
            _business_component(str(batch.get("group_name") or ""), "业务分组"),
            _business_component(str(batch.get("tool_name") or ""), "工具"),
            _visible_component(str(batch.get("directory_name") or "")),
        )
        if directory_parts[:3] != expected_prefix:
            raise ProjectStoreError("批次目录与业务信息不一致。")
    return BatchSummary(
        id=batch_id,
        group_name=str(batch.get("group_name") or ""),
        tool_id=str(batch.get("tool_id") or ""),
        tool_name=str(batch.get("tool_name") or ""),
        status=status,
        directory_name=str(batch.get("directory_name") or ""),
        business_description=str(batch.get("business_description") or ""),
        business_period=str(batch.get("business_period") or ""),
        created_at=str(batch.get("created_at") or ""),
        started_at=_optional_text(batch.get("started_at")),
        finished_at=_optional_text(batch.get("finished_at")),
        error_message=_optional_text(batch.get("error_message")),
        deleted_at=_optional_text(batch.get("deleted_at")),
    )


def _project_file_from_dict(item: dict[str, Any]) -> ProjectFile:
    category = str(item.get("category") or "")
    if category not in CATEGORIES:
        raise ProjectStoreError("项目文件分类无效。")
    relative = _validated_relative_path(str(item.get("relative_path") or ""))
    digest = str(item.get("sha256") or "").lower()
    if len(digest) != 64 or any(character not in "0123456789abcdef" for character in digest):
        raise ProjectStoreError("项目文件校验值无效。")
    return ProjectFile(
        id=_required_uuid(item, "id"),
        batch_id=_required_uuid(item, "batch_id"),
        category=category,
        role=str(item.get("role") or "main"),
        display_name=Path(str(item.get("display_name") or Path(relative).name)).name,
        relative_path=relative,
        size_bytes=max(0, int(item.get("size_bytes", 0))),
        sha256=digest,
        modified_ns=max(0, int(item.get("modified_ns", 0))),
    )


def _batch_object(manifest: dict[str, Any]) -> dict[str, Any]:
    batch = manifest.get("batch")
    if not isinstance(batch, dict):
        raise ProjectStoreError("批次清单缺少任务信息。")
    return batch


def _file_objects(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    files = manifest.get("files")
    if not isinstance(files, list) or any(not isinstance(item, dict) for item in files):
        raise ProjectStoreError("批次文件清单无效。")
    return files


def _directory_map(manifest: dict[str, Any]) -> dict[str, str]:
    batch = _batch_object(manifest)
    return _validated_directory_map(
        manifest.get("directories"),
        allow_hidden=bool(batch.get("deleted_at")) or bool(manifest.get("pending_restore")),
    )


def _trash_original_directory_map(
    manifest: dict[str, Any],
    summary: BatchSummary | None = None,
) -> dict[str, str]:
    original = manifest.get("original_directories")
    if not isinstance(original, dict):
        raise ProjectStoreError("回收站清单缺少原位置。")
    directories = _validated_directory_map(original)
    batch_summary = summary or _summary_from_manifest(manifest)
    expected_prefix = (
        _business_component(batch_summary.group_name, "业务分组"),
        _business_component(batch_summary.tool_name, "工具"),
        _visible_component(batch_summary.directory_name),
    )
    if Path(directories[CATEGORY_UPLOADS]).parts[:3] != expected_prefix:
        raise ProjectStoreError("回收站原位置与批次信息不一致。")
    return directories


def _validated_directory_map(value: Any, *, allow_hidden: bool = False) -> dict[str, str]:
    if not isinstance(value, dict) or set(value) != set(CATEGORIES):
        raise ProjectStoreError("批次目录信息无效。")
    result: dict[str, str] = {}
    for category in CATEGORIES:
        relative = _validated_relative_path(str(value[category]))
        parts = Path(relative).parts
        visible_shape = len(parts) == 4 and parts[-1] == CATEGORY_VISIBLE_NAMES[category]
        hidden_shape = (
            allow_hidden
            and len(parts) == 5
            and parts[0] == PROJECT_METADATA_DIR
            and parts[1] == PROJECT_TRASH_DIR
            and parts[3] == "batch"
            and parts[-1] == CATEGORY_VISIBLE_NAMES[category]
        )
        if not visible_shape and not hidden_shape:
            raise ProjectStoreError("批次目录不属于正确分类。")
        if visible_shape:
            if parts[0] != _business_component(parts[0], "业务分组"):
                raise ProjectStoreError("批次业务分组目录无效。")
            if parts[1] != _business_component(parts[1], "工具"):
                raise ProjectStoreError("批次工具目录无效。")
            if parts[2] != _visible_component(parts[2]):
                raise ProjectStoreError("批次目录名称无效。")
        result[category] = relative
    common_roots = {Path(relative).parts[:3] for relative in result.values()}
    if len(common_roots) != 1:
        raise ProjectStoreError("同一批次的资料目录不在一起。")
    return result


def _validate_manifest_identity(manifest: dict[str, Any], batch_id: str) -> None:
    if _summary_from_manifest(manifest).id != batch_id.lower():
        raise ProjectStoreError("批次清单与编号不一致。")
    directories = _directory_map(manifest)
    for raw_item in _file_objects(manifest):
        item = _project_file_from_dict(raw_item)
        if item.batch_id != batch_id.lower():
            raise ProjectStoreError("项目文件与批次不一致。")
        relative = Path(item.relative_path)
        try:
            inside = relative.relative_to(Path(directories[item.category]))
        except ValueError as exc:
            raise ProjectStoreError("项目文件路径不属于批次目录。") from exc
        if not inside.parts:
            raise ProjectStoreError("项目文件路径不能指向批次目录本身。")


def _read_json(path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProjectStoreError(f"无法读取项目清单 {path.name}：{exc}") from exc
    if not isinstance(payload, dict):
        raise ProjectStoreError(f"项目清单格式无效：{path.name}")
    return payload


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    try:
        _atomic_write_text(
            path,
            json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True, default=str) + "\n",
        )
    except Exception as exc:
        if isinstance(exc, ProjectStoreError):
            raise
        raise ProjectStoreError(f"无法保存项目清单 {path.name}：{exc}") from exc


def _validate_project_location(path: Path, *, for_write: bool) -> Path:
    absolute = path.absolute()
    # Preserve the path the user actually chose long enough to inspect every
    # existing parent.  Resolving first would erase a linked parent and could
    # make an external location look like an ordinary local folder.
    _assert_existing_ancestors_are_real(
        absolute,
        allow_macos_root_aliases=True,
    )
    if absolute.exists() and _is_link_like(absolute):
        raise ProjectStoreError("项目位置不能是链接或系统重定向目录。")
    resolved = absolute.resolve()
    _assert_existing_ancestors_are_real(resolved)
    anchor = Path(resolved.anchor)
    home = Path.home().resolve()
    if resolved == anchor or resolved == home:
        raise ProjectStoreError("项目位置过于宽泛，请选择专用子文件夹。")
    for ancestor in (resolved, *resolved.parents):
        marker = ancestor / PROJECT_METADATA_DIR / PROJECT_FILE_NAME
        if ancestor != resolved and marker.is_file() and not _is_link_like(marker):
            raise ProjectStoreError("不能在另一个 HRToolkit 项目中创建子项目。")
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).resolve().parent
        if resolved == app_dir or _is_inside(resolved, app_dir):
            raise ProjectStoreError("项目不能放在程序安装目录。")
    if for_write:
        unsafe_reason = _unsafe_active_location_reason(resolved)
        if unsafe_reason:
            raise ProjectStoreError(unsafe_reason)
    return resolved


def _unsafe_active_location_reason(path: Path) -> str | None:
    text = str(path)
    if os.name == "nt":
        if text.startswith("\\\\"):
            return "共享盘不能作为活动项目位置，请先在本机创建项目。"
        try:
            import ctypes

            drive_type = int(ctypes.windll.kernel32.GetDriveTypeW(str(Path(path.anchor))))
            if drive_type == 2:
                return "U 盘或移动磁盘不能作为活动项目位置，可用于项目备份。"
            if drive_type == 4:
                return "网络共享盘不能作为活动项目位置，请先在本机创建项目。"
        except Exception:
            pass
    elif sys.platform == "darwin":
        try:
            path.relative_to(Path("/Volumes"))
            return "外接磁盘不能作为活动项目位置，可用于项目备份。"
        except ValueError:
            pass

    known_sync_roots: list[Path] = []
    for variable in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer", "Dropbox"):
        value = os.environ.get(variable, "").strip()
        if value:
            known_sync_roots.append(Path(value).expanduser().resolve())
    known_sync_roots.extend(
        (
            Path.home() / "Library" / "CloudStorage",
            Path.home() / "Library" / "Mobile Documents",
        )
    )
    if any(root.exists() and _is_inside(path, root) for root in known_sync_roots):
        return "同步盘不能作为活动项目位置，请使用本机文件夹后再备份。"
    return None


def _assert_existing_ancestors_are_real(
    path: Path,
    *,
    allow_macos_root_aliases: bool = False,
) -> None:
    chain: list[Path] = []
    current = path
    while current != current.parent:
        chain.append(current)
        current = current.parent
    chain.append(current)
    for item in reversed(chain):
        if item.exists() and _is_link_like(item):
            if allow_macos_root_aliases and _is_known_macos_root_alias(item):
                continue
            raise ProjectStoreError("项目路径不能经过链接或系统重定向目录。")


def _is_known_macos_root_alias(path: Path) -> bool:
    """Allow only Apple's fixed root aliases, never user-created link parents."""

    if sys.platform != "darwin" or path.parent != Path(path.anchor):
        return False
    expected = {
        "etc": Path("/private/etc"),
        "tmp": Path("/private/tmp"),
        "var": Path("/private/var"),
    }.get(path.name)
    if expected is None:
        return False
    try:
        return path.resolve(strict=True) == expected
    except OSError:
        return False


def _require_regular_directory(path: Path, label: str) -> None:
    try:
        item_stat = path.lstat()
    except OSError as exc:
        raise ProjectStoreError(f"{label}不存在或无法读取。") from exc
    if (
        stat_module.S_ISLNK(item_stat.st_mode)
        or bool(getattr(item_stat, "st_file_attributes", 0) & 0x400)
        or not stat_module.S_ISDIR(item_stat.st_mode)
    ):
        raise ProjectStoreError(f"{label}不是安全的普通文件夹。")


def _require_regular_file(path: Path, label: str) -> None:
    try:
        item_stat = path.lstat()
    except OSError as exc:
        raise ProjectStoreError(f"{label}不存在或无法读取。") from exc
    if (
        stat_module.S_ISLNK(item_stat.st_mode)
        or bool(getattr(item_stat, "st_file_attributes", 0) & 0x400)
        or not stat_module.S_ISREG(item_stat.st_mode)
    ):
        raise ProjectStoreError(f"{label}不是安全的普通文件。")


def _walk_external_directory(
    root: Path,
    *,
    cancelled: Callable[[], bool] | None = None,
) -> Iterator[tuple[Path, tuple[str, ...]]]:
    yield from _walk_directory_strict(
        root,
        ignore_temporary=True,
        cancelled=cancelled,
    )


def _walk_directory_strict(
    root: Path,
    *,
    ignore_temporary: bool,
    cancelled: Callable[[], bool] | None = None,
) -> Iterator[tuple[Path, tuple[str, ...]]]:
    _raise_if_cancelled(cancelled)
    _require_regular_directory(root, "导入文件夹")
    for current, directories, files in os.walk(root, followlinks=False):
        _raise_if_cancelled(cancelled)
        current_path = Path(current)
        for name in sorted((*directories, *files), key=str.casefold):
            _raise_if_cancelled(cancelled)
            candidate = current_path / name
            if _is_link_like(candidate):
                raise ProjectStoreError(f"所选文件夹包含链接，未导入：{candidate.name}")
        directories[:] = sorted(directories, key=str.casefold)
        for name in sorted(files, key=str.casefold):
            _raise_if_cancelled(cancelled)
            path = current_path / name
            _require_regular_file(path, "导入文件")
            if ignore_temporary:
                if _is_ignored_import_file(path):
                    continue
                _reject_forbidden_import_file(path)
            resolved = path.resolve()
            if not _is_inside(resolved, root):
                raise ProjectStoreError("导入文件路径越界。")
            yield resolved, resolved.relative_to(root).parts


def _is_ignored_import_file(path: Path) -> bool:
    lowered = path.name.casefold()
    return (
        lowered in IGNORED_IMPORT_NAMES
        or lowered.startswith(("~$", ".~lock."))
        or path.suffix.casefold() in IGNORED_IMPORT_SUFFIXES
    )


def _reject_forbidden_import_file(path: Path) -> None:
    if path.suffix.casefold() in FORBIDDEN_IMPORT_SUFFIXES:
        raise ProjectStoreError(f"为保护电脑安全，不能导入此类文件：{path.name}")


def _remove_empty_directories(root: Path) -> None:
    if not root.is_dir() or _is_link_like(root):
        return
    directories = sorted(
        (path for path in root.rglob("*") if path.is_dir() and not _is_link_like(path)),
        key=lambda item: len(item.parts),
        reverse=True,
    )
    for directory in directories:
        try:
            directory.rmdir()
        except OSError:
            pass


def _copy_external_file(
    source: Path,
    destination: Path,
    *,
    cancelled: Callable[[], bool] | None,
    on_chunk: Callable[[int], None] | None,
) -> dict[str, Any]:
    _require_regular_file(source, "导入文件")
    destination.parent.mkdir(parents=True, exist_ok=True)
    flags = os.O_RDONLY
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    descriptor = os.open(source, flags)
    digest = hashlib.sha256()
    written = 0
    try:
        before = os.fstat(descriptor)
        with os.fdopen(descriptor, "rb", closefd=False) as input_handle, destination.open("xb") as output_handle:
            while True:
                _raise_if_cancelled(cancelled)
                chunk = input_handle.read(COPY_BUFFER_BYTES)
                if not chunk:
                    break
                output_handle.write(chunk)
                digest.update(chunk)
                written += len(chunk)
                if on_chunk is not None:
                    on_chunk(written)
            output_handle.flush()
            os.fsync(output_handle.fileno())
        after_handle = os.fstat(descriptor)
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    finally:
        os.close(descriptor)
    try:
        after_path = source.stat()
    except OSError as exc:
        destination.unlink(missing_ok=True)
        raise ProjectStoreError(f"复制期间来源文件被移动：{source.name}") from exc
    identity_before = (before.st_dev, before.st_ino, before.st_size, before.st_mtime_ns)
    identity_after_handle = (
        after_handle.st_dev,
        after_handle.st_ino,
        after_handle.st_size,
        after_handle.st_mtime_ns,
    )
    identity_after_path = (
        after_path.st_dev,
        after_path.st_ino,
        after_path.st_size,
        after_path.st_mtime_ns,
    )
    if written != before.st_size or identity_before != identity_after_handle or identity_before != identity_after_path:
        destination.unlink(missing_ok=True)
        raise ProjectStoreError(f"复制期间文件发生变化，请关闭文件后重试：{source.name}")
    _make_private(destination)
    return {
        "size_bytes": written,
        "sha256": digest.hexdigest(),
        "modified_ns": before.st_mtime_ns,
    }


def _hash_stable_file(path: Path) -> dict[str, Any]:
    _require_regular_file(path, "项目文件")
    before = path.stat()
    digest = _sha256_file(path)
    after = path.stat()
    if (
        (before.st_dev, before.st_ino, before.st_size, before.st_mtime_ns)
        != (after.st_dev, after.st_ino, after.st_size, after.st_mtime_ns)
    ):
        raise ProjectStoreError(f"校验期间文件发生变化：{path.name}")
    return {"size_bytes": after.st_size, "sha256": digest, "modified_ns": after.st_mtime_ns}


def _publish_no_replace(staging: Path, destination: Path) -> None:
    if destination.exists() or _is_link_like(destination):
        raise ProjectStoreError(f"目标文件已存在，不能覆盖：{destination.name}")
    try:
        os.link(staging, destination)
        staging.unlink()
    except FileExistsError as exc:
        raise ProjectStoreError(f"目标文件已存在，不能覆盖：{destination.name}") from exc
    except OSError as exc:
        # Windows os.rename refuses to replace an existing destination.  On
        # POSIX the hard-link path above is the no-clobber primitive.
        if os.name != "nt":
            raise ProjectStoreError(f"无法安全发布文件：{destination.name}") from exc
        try:
            staging.rename(destination)
        except FileExistsError as rename_exc:
            raise ProjectStoreError(f"目标文件已存在，不能覆盖：{destination.name}") from rename_exc


def _existing_name_keys(
    root: Path,
    *,
    cancelled: Callable[[], bool] | None = None,
) -> set[str]:
    keys: set[str] = set()
    _raise_if_cancelled(cancelled)
    if not root.exists():
        return keys
    _require_regular_directory(root, "批次资料目录")
    for path in root.rglob("*"):
        _raise_if_cancelled(cancelled)
        if _is_link_like(path):
            raise ProjectStoreError("批次资料目录包含链接，已停止导入。")
        keys.add(_portable_relative_key(path.relative_to(root)))
    return keys


def _source_total_bytes(
    items: Sequence[_ExternalSourceItem],
    *,
    cancelled: Callable[[], bool] | None = None,
) -> int:
    total = 0
    for item in items:
        _raise_if_cancelled(cancelled)
        total += item.path.stat().st_size
    return total


def _unique_destination(
    path: Path,
    reserved_keys: set[str],
    destination_root: Path,
    *,
    is_directory: bool = False,
    cancelled: Callable[[], bool] | None = None,
) -> Path:
    suffixes = "" if is_directory else "".join(path.suffixes)
    base = path.name[: -len(suffixes)] if suffixes else path.name
    for index in range(1, 10_000):
        _raise_if_cancelled(cancelled)
        candidate = path if index == 1 else path.with_name(f"{base} ({index}){suffixes}")
        try:
            key = _portable_relative_key(candidate.relative_to(destination_root))
        except ValueError as exc:
            raise ProjectStoreError("导入目标路径越界。") from exc
        collision = key in reserved_keys
        if not collision and not candidate.exists() and not _is_link_like(candidate):
            return candidate
    raise ProjectStoreError(f"同名文件过多，无法导入：{path.name}")


def _portable_relative_key(path: Path) -> str:
    return "/".join(_portable_text_key(part) for part in path.parts)


def _portable_text_key(value: str) -> str:
    import unicodedata

    return unicodedata.normalize("NFC", value).casefold().rstrip(" .")


def _visible_component(value: str) -> str:
    component = _safe_component(str(value))
    if component.casefold() == PROJECT_METADATA_DIR.casefold():
        raise ProjectStoreError("文件或文件夹名称不能使用项目隐藏管理目录名称。")
    return component


def _business_component(value: str, label: str) -> str:
    component = _visible_component(value)
    reserved = {
        COMMON_VISIBLE_DIR.casefold(),
        *(name.casefold() for name in CATEGORY_VISIBLE_NAMES.values()),
    }
    if component.casefold() in reserved:
        raise ProjectStoreError(f"{label}名称属于项目保留名称，请换一个名称。")
    return component


def _validated_relative_path(value: str) -> str:
    path = Path(value)
    if path.is_absolute() or not path.parts or any(part in {"", ".", ".."} for part in path.parts):
        raise ProjectStoreError("项目相对路径无效。")
    if path.drive or str(value).startswith(("/", "\\")):
        raise ProjectStoreError("项目相对路径无效。")
    return path.as_posix()


def _project_join(root: Path, relative: str | Path) -> Path:
    normalized = _validated_relative_path(str(relative))
    target = (root / Path(normalized)).resolve()
    if not _is_inside(target, root.resolve()):
        raise ProjectStoreError("项目路径越界。")
    return target


def _assert_no_link_components(root: Path, relative: Path) -> None:
    current = root
    if _is_link_like(current):
        raise ProjectStoreError("项目根目录不安全。")
    for part in relative.parts:
        current = current / part
        if _is_link_like(current):
            raise ProjectStoreError("项目路径包含链接或系统重定向目录。")


def _is_inside(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _paths_overlap(first: Path, second: Path) -> bool:
    return _is_inside(first, second) or _is_inside(second, first)


def _required_uuid(payload: dict[str, Any], key: str) -> str:
    value = str(payload.get(key) or "").lower()
    if not _is_uuid_hex(value):
        raise ProjectStoreError(f"{key} 无效。")
    return value


def _is_uuid_hex(value: str) -> bool:
    try:
        return len(value) == 32 and uuid.UUID(hex=value).hex == value.lower()
    except (ValueError, AttributeError):
        return False


def _required_int(payload: dict[str, Any], key: str) -> int:
    try:
        return int(payload[key])
    except (KeyError, TypeError, ValueError) as exc:
        raise ProjectStoreError(f"项目标记缺少 {key}。") from exc


def _optional_text(value: Any) -> str | None:
    return None if value is None else str(value)


def _raise_if_cancelled(cancelled: Callable[[], bool] | None) -> None:
    if cancelled is not None and cancelled():
        raise ImportCancelled("本次导入已停止。")


def _report_import_progress(
    callback: Callable[[ImportProgress], None] | None,
    event: ImportProgress,
) -> None:
    if callback is not None:
        callback(event)


def _report_checking_progress(
    callback: Callable[[ImportProgress], None] | None,
    items: Sequence[_ExternalSourceItem],
    name: str,
) -> None:
    _report_import_progress(
        callback,
        ImportProgress(
            phase="checking",
            current_name=name,
            files_scanned=len(items),
        ),
    )


def _report_copy_callbacks(
    progress: Callable[[int, int, str], None] | None,
    on_progress: Callable[[ImportProgress], None] | None,
    copied_bytes: int,
    total_bytes: int,
    name: str,
    *,
    files_scanned: int,
    files_completed: int,
    files_total: int,
) -> None:
    _report_copy_progress(progress, copied_bytes, total_bytes, name)
    _report_import_progress(
        on_progress,
        ImportProgress(
            phase="copying",
            current_name=name,
            files_scanned=files_scanned,
            files_completed=files_completed,
            files_total=files_total,
            bytes_copied=copied_bytes,
            bytes_total=total_bytes,
        ),
    )


def _report_copy_progress(
    callback: Callable[[int, int, str], None] | None,
    copied_bytes: int,
    total_bytes: int,
    name: str,
) -> None:
    if callback is not None:
        callback(copied_bytes, total_bytes, name)


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _utc_from_datetime(value: datetime) -> str:
    if value.tzinfo is None:
        value = value.astimezone()
    return value.astimezone(timezone.utc).isoformat(timespec="seconds")


def _clean_error(value: str | None) -> str | None:
    if value is None:
        return None
    return " ".join(str(value).split())[:2000]


def _hide_on_windows(path: Path) -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        hidden_attribute = 0x2
        attributes = ctypes.windll.kernel32.GetFileAttributesW(str(path))
        if attributes != -1:
            ctypes.windll.kernel32.SetFileAttributesW(str(path), attributes | hidden_attribute)
    except Exception:
        # The leading dot still keeps the metadata visually separate in the
        # in-app tree; failure to add the Explorer hidden flag is non-fatal.
        pass
