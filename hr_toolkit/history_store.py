"""Persistent local history for HR Toolkit runs.

The SQLite database stores only searchable metadata.  The original inputs and
generated outputs remain ordinary files under a per-user data directory so a
record can still be inspected or recovered without the application.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sqlite3
import stat as stat_module
import struct
import sys
import threading
import unicodedata
import uuid
import weakref
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator


DATA_DIR_ENV = "HR_TOOLKIT_DATA_DIR"
SCHEMA_VERSION = 1
SQLITE_APPLICATION_ID = 0x4852544B  # "HRTK"
COPY_BUFFER_BYTES = 1024 * 1024
MIN_FREE_SPACE_BYTES = 128 * 1024 * 1024
FREE_SPACE_RATIO = 0.05
HISTORY_PAGE_SIZE = 50
ALLOWED_STATUSES = {"running", "success", "failed", "stopped"}
TERMINAL_STATUSES = ALLOWED_STATUSES - {"running"}
MARKER_NAME = ".hrtoolkit-data-v1"
MARKER_CONTENT = "HRToolkit local history data\n"
DATABASE_NAME = "history.db"
MANIFEST_NAME = "manifest.json"
DATABASE_BACKUPS_DIR_NAME = "database-backups"
DATABASE_RECOVERY_PENDING_NAME = ".database-recovery-pending.json"
DATABASE_BACKUP_MANIFEST_NAME = "complete.json"
DATABASE_ACCESS_LOCK_NAME = ".database-access.lock"
TASK_ARCHIVE_LOCKS_DIR_NAME = ".task-archive-locks"
TRASH_MOVE_PENDING_PREFIX = ".trash-move-"
DEFAULT_ARCHIVE_SUFFIXES = frozenset({".xlsx", ".xls", ".zip"})
SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS tasks (
    id TEXT PRIMARY KEY,
    tool_id TEXT NOT NULL,
    tool_name TEXT NOT NULL,
    mode TEXT,
    app_version TEXT NOT NULL,
    status TEXT NOT NULL CHECK(status IN ('running', 'success', 'failed', 'stopped')),
    started_at TEXT NOT NULL,
    finished_at TEXT,
    parameters_json TEXT NOT NULL DEFAULT '{}',
    result_json TEXT NOT NULL DEFAULT '{}',
    error_message TEXT,
    owner_pid INTEGER,
    task_relpath TEXT NOT NULL UNIQUE,
    deleted_at TEXT
);
CREATE TABLE IF NOT EXISTS files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id TEXT NOT NULL REFERENCES tasks(id) ON DELETE CASCADE,
    kind TEXT NOT NULL CHECK(kind IN ('input', 'output')),
    role TEXT NOT NULL,
    display_name TEXT NOT NULL,
    original_path TEXT NOT NULL,
    relative_path TEXT NOT NULL,
    size_bytes INTEGER NOT NULL,
    sha256 TEXT NOT NULL,
    modified_ns INTEGER NOT NULL,
    UNIQUE(task_id, kind, relative_path)
);
CREATE INDEX IF NOT EXISTS idx_tasks_started_at ON tasks(started_at DESC);
CREATE INDEX IF NOT EXISTS idx_tasks_tool_started ON tasks(tool_id, started_at DESC);
CREATE INDEX IF NOT EXISTS idx_tasks_active ON tasks(deleted_at, started_at DESC);
CREATE INDEX IF NOT EXISTS idx_files_task_kind ON files(task_id, kind);
CREATE INDEX IF NOT EXISTS idx_files_display_name ON files(display_name);
"""


# ``msvcrt.locking`` treats a second handle in the same process differently
# from a second process and can raise ``OSError(36, "Resource deadlock
# avoided")`` when several worker threads open the same lock file at once.
# Keep a lightweight process-local lock in front of the OS lock.  The weak
# registry avoids retaining one Python lock forever for every historical task.
_PROCESS_FILE_LOCKS_GUARD = threading.Lock()
_PROCESS_FILE_LOCKS: weakref.WeakValueDictionary[str, threading.RLock] = weakref.WeakValueDictionary()


class HistoryStoreError(RuntimeError):
    """Raised when a run cannot be archived safely."""


class _CorruptHistoryDatabase(sqlite3.DatabaseError):
    """Internal signal used only for verified SQLite corruption."""


@dataclass(frozen=True)
class SourceSpec:
    path: Path
    role: str = "main"
    suffixes: frozenset[str] | None = DEFAULT_ARCHIVE_SUFFIXES


@dataclass(frozen=True)
class FileRecord:
    id: int
    task_id: str
    kind: str
    role: str
    display_name: str
    original_path: str
    archived_path: Path
    relative_path: str
    size_bytes: int
    sha256: str
    modified_ns: int


@dataclass(frozen=True)
class TaskSummary:
    id: str
    tool_id: str
    tool_name: str
    mode: str | None
    app_version: str
    status: str
    started_at: str
    finished_at: str | None
    input_names: tuple[str, ...]
    output_names: tuple[str, ...]
    error_message: str | None
    deleted_at: str | None = None


@dataclass(frozen=True)
class TaskDetail:
    summary: TaskSummary
    parameters: dict[str, Any]
    result: dict[str, Any]
    task_dir: Path
    input_dir: Path
    output_dir: Path
    inputs: tuple[FileRecord, ...]
    outputs: tuple[FileRecord, ...]


def default_history_root() -> Path:
    override = os.environ.get(DATA_DIR_ENV, "").strip()
    if override:
        return Path(override).expanduser()
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("LOCALAPPDATA", "").strip() or (Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME", "").strip() or (Path.home() / ".local" / "share"))
    return base / "HRToolkit" / "Data"


class HistoryStore:
    def __init__(self, root: str | Path | None = None) -> None:
        requested = Path(root).expanduser() if root is not None else default_history_root()
        self.root = _validate_data_root(requested)
        self.records_dir = self.root / "records"
        self.trash_dir = self.root / "trash"
        self.archive_locks_dir = self.root / TASK_ARCHIVE_LOCKS_DIR_NAME
        self.database_backups_dir = self.root / DATABASE_BACKUPS_DIR_NAME
        self.database_path = self.root / DATABASE_NAME
        self.recovered_database_backup: Path | None = None
        self.last_rebuild_report: dict[str, Any] = {
            "discovered": 0,
            "scanned": 0,
            "validated": 0,
            "restored": 0,
            "errors": [],
            "missing_manifests": [],
            "unexpected_entries": [],
        }
        self._task_locks: dict[str, threading.RLock] = {}
        self._task_io_locks: dict[str, threading.RLock] = {}
        self._locks_guard = threading.Lock()
        self._database_gate_lock = threading.RLock()
        self._database_gate_owner: int | None = None
        startup_lock = _startup_lock_path(self.root)
        with _exclusive_file_lock(startup_lock, private_parent=False):
            self._prepare_root()
            with self._database_gate():
                self._initialize_database_with_recovery()
                self._recover_pending_trash_moves()
        self.recover_incomplete_tasks()
        self._cleanup_recent_terminal_staging()

    def _prepare_root(self) -> None:
        marker = self.root / MARKER_NAME
        if self.root.exists() and not marker.exists():
            try:
                has_existing_content = next(self.root.iterdir(), None) is not None
            except OSError as exc:
                raise HistoryStoreError(f"无法读取资料库位置：{exc}") from exc
            if has_existing_content:
                raise HistoryStoreError("资料库位置已被其他文件占用，请使用 HRToolkit 专用文件夹。")
        _mkdir_private(self.root)
        _assert_storage_entry(marker, "资料库标记", kind="file", allow_missing=True)
        if not marker.exists():
            _atomic_write_text(marker, MARKER_CONTENT)
        elif marker.read_text(encoding="utf-8") != MARKER_CONTENT:
            raise HistoryStoreError("资料库标记不匹配，为保护现有文件，已停止使用该目录。")
        _make_private(marker)
        _assert_storage_entry(self.records_dir, "历史记录目录", kind="directory", allow_missing=True)
        _assert_storage_entry(self.trash_dir, "历史回收站目录", kind="directory", allow_missing=True)
        _assert_storage_entry(self.archive_locks_dir, "历史资料锁目录", kind="directory", allow_missing=True)
        _assert_storage_entry(self.database_backups_dir, "历史索引备份目录", kind="directory", allow_missing=True)
        _mkdir_private(self.records_dir)
        _mkdir_private(self.trash_dir)
        _mkdir_private(self.archive_locks_dir)
        recovery_marker = self.root / DATABASE_RECOVERY_PENDING_NAME
        _assert_storage_entry(recovery_marker, "历史索引恢复标记", kind="file", allow_missing=True)
        database_access_lock = self.root / DATABASE_ACCESS_LOCK_NAME
        _assert_storage_entry(database_access_lock, "历史索引访问锁", kind="file", allow_missing=True)
        for path in _database_artifact_paths(self.database_path):
            _assert_storage_entry(path, "历史数据库", kind="file", allow_missing=True)
        for path in self.root.glob(f"{TRASH_MOVE_PENDING_PREFIX}*"):
            if not _is_trash_move_marker_name(path.name):
                raise HistoryStoreError("回收站移动恢复标记名称无效。")
            _assert_storage_entry(path, "回收站移动恢复标记", kind="file", allow_missing=False)
        self._cleanup_stale_probe_directories()

    def _cleanup_stale_probe_directories(self) -> None:
        for path in self.root.glob(".database-probe-*"):
            if not _is_probe_directory_name(path.name):
                continue
            if _is_link_like(path) or not path.is_dir():
                raise HistoryStoreError("历史索引检查临时目录不安全。")
            try:
                shutil.rmtree(path)
            except OSError as exc:
                raise HistoryStoreError(f"无法清理历史索引检查临时目录：{exc}") from exc

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        with self._database_gate():
            connection = sqlite3.connect(self.database_path, timeout=10.0)
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys = ON")
            connection.execute("PRAGMA busy_timeout = 10000")
            try:
                yield connection
                connection.commit()
            except Exception:
                connection.rollback()
                raise
            finally:
                connection.close()

    @contextmanager
    def _database_gate(self) -> Iterator[None]:
        current_thread = threading.get_ident()
        with self._database_gate_lock:
            if self._database_gate_owner == current_thread:
                yield
                return
            lock_path = self.root / DATABASE_ACCESS_LOCK_NAME
            with _exclusive_file_lock(lock_path):
                self._database_gate_owner = current_thread
                try:
                    yield
                finally:
                    self._database_gate_owner = None

    def _initialize_database(self) -> None:
        for path in _database_artifact_paths(self.database_path):
            _assert_storage_entry(path, "历史数据库", kind="file", allow_missing=True)
        with self._connect() as connection:
            user_version = int(connection.execute("PRAGMA user_version").fetchone()[0])
            if user_version not in {0, SCHEMA_VERSION}:
                raise HistoryStoreError(f"历史资料库版本不兼容：{user_version}")
            existing_tables = {
                str(row[0])
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'"
                ).fetchall()
            }
            unexpected_tables = existing_tables - {"tasks", "files"}
            if unexpected_tables:
                raise HistoryStoreError("历史资料库结构无法识别，为保护现有数据，已停止使用。")
            connection.execute("PRAGMA journal_mode = WAL")
            connection.execute("PRAGMA synchronous = NORMAL")
            connection.executescript(SCHEMA_SQL)
            if _database_schema_fingerprint(connection) != _expected_database_schema_fingerprint():
                raise HistoryStoreError("历史资料库结构不匹配，为保护现有数据，已停止使用。")
            connection.execute(f"PRAGMA application_id = {SQLITE_APPLICATION_ID}")
            connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        _make_private(self.database_path)

    def _initialize_database_with_recovery(self) -> None:
        recovery_marker = self.root / DATABASE_RECOVERY_PENDING_NAME
        if recovery_marker.exists():
            self._normalize_pending_trash_moves_without_database()
            self._resume_pending_database_recovery(recovery_marker)
            return

        existing_artifacts = [path for path in _database_artifact_paths(self.database_path) if path.exists()]
        if not self.database_path.exists():
            if existing_artifacts:
                self._normalize_pending_trash_moves_without_database()
                self._recover_corrupt_database()
            elif self._has_archive_evidence():
                self._normalize_pending_trash_moves_without_database()
                self._recover_missing_database()
            else:
                self._initialize_database()
                self._quick_check_database()
            return

        if self.database_path.stat().st_size == 0 and self._has_archive_evidence():
            self._normalize_pending_trash_moves_without_database()
            self._recover_corrupt_database()
            return

        try:
            self._probe_existing_database()
        except HistoryStoreError:
            # A future schema or an unrecognized healthy database must never
            # be treated as damage and rewritten.
            raise
        except sqlite3.DatabaseError as exc:
            if not _is_database_corruption(exc):
                raise
            self._normalize_pending_trash_moves_without_database()
            self._recover_corrupt_database()
            return

        try:
            self._initialize_database()
            self._quick_check_database()
        except HistoryStoreError:
            raise
        except sqlite3.DatabaseError as exc:
            if not _is_database_corruption(exc):
                raise
            self._normalize_pending_trash_moves_without_database()
            self._recover_corrupt_database()

    def _probe_existing_database(self) -> None:
        artifacts = [path for path in _database_artifact_paths(self.database_path) if path.exists()]
        probe_dir: Path | None = None
        copied_wal_size = 0
        if any(path != self.database_path for path in artifacts):
            probe_dir = self.root / f".database-probe-{uuid.uuid4().hex}"
            probe_dir.mkdir(mode=0o700, exist_ok=False)
            _make_private(probe_dir, directory=True)
            try:
                before_stats = {
                    path.name: (path.stat().st_size, path.stat().st_mtime_ns)
                    for path in artifacts
                }
                for source in artifacts:
                    _assert_storage_entry(source, "历史数据库", kind="file", allow_missing=False)
                    # SHM is a disposable WAL index. Let SQLite recreate it only
                    # inside the isolated probe directory.
                    if source.name.endswith("-shm"):
                        continue
                    _copy_file_durable(source, probe_dir / source.name)
                    if source.name.endswith("-wal"):
                        copied_wal_size = source.stat().st_size
                after_artifacts = [path for path in _database_artifact_paths(self.database_path) if path.exists()]
                after_stats = {
                    path.name: (path.stat().st_size, path.stat().st_mtime_ns)
                    for path in after_artifacts
                }
                if before_stats != after_stats:
                    raise HistoryStoreError("历史索引仍在变化，请关闭其他 HRToolkit 窗口后重试。")
                probe_database = probe_dir / DATABASE_NAME
                base_uri = probe_database.resolve().as_uri() + "?mode=ro&immutable=1"
                base_connection = sqlite3.connect(base_uri, uri=True, timeout=10.0)
                base_connection.row_factory = sqlite3.Row
                try:
                    self._validate_probe_database(base_connection)
                finally:
                    base_connection.close()
                connection = sqlite3.connect(probe_database, timeout=10.0)
            except Exception:
                shutil.rmtree(probe_dir, ignore_errors=True)
                raise
        else:
            uri = self.database_path.resolve().as_uri() + "?mode=ro&immutable=1"
            connection = sqlite3.connect(uri, uri=True, timeout=10.0)
        connection.row_factory = sqlite3.Row
        try:
            if probe_dir is not None and copied_wal_size:
                probe_wal = probe_dir / f"{DATABASE_NAME}-wal"
                wal_frame_count = _validated_wal_frame_count(probe_wal)
                if wal_frame_count is None:
                    raise _CorruptHistoryDatabase("WAL 日志校验失败")
                checkpoint = connection.execute("PRAGMA wal_checkpoint(PASSIVE)").fetchone()
                if checkpoint is None:
                    raise _CorruptHistoryDatabase("WAL 日志无法校验")
                busy, log_frames, checkpointed_frames = (int(value) for value in checkpoint[:3])
                if (
                    busy != 0
                    or log_frames != wal_frame_count
                    or checkpointed_frames != log_frames
                ):
                    raise _CorruptHistoryDatabase("WAL 日志无法完整校验")
            connection.execute("PRAGMA query_only = ON")
            self._validate_probe_database(connection)
        finally:
            connection.close()
            if probe_dir is not None:
                shutil.rmtree(probe_dir, ignore_errors=True)

    def _validate_probe_database(self, connection: sqlite3.Connection) -> None:
        application_id = int(connection.execute("PRAGMA application_id").fetchone()[0])
        if application_id not in {0, SQLITE_APPLICATION_ID}:
            raise HistoryStoreError("历史资料库不是 HRToolkit 创建的，为保护现有数据，已停止使用。")
        user_version = int(connection.execute("PRAGMA user_version").fetchone()[0])
        if user_version not in {0, SCHEMA_VERSION}:
            raise HistoryStoreError(f"历史资料库版本不兼容：{user_version}")
        fingerprint = _database_schema_fingerprint(connection)
        if fingerprint:
            if fingerprint != _expected_database_schema_fingerprint():
                raise HistoryStoreError("历史资料库结构无法识别，为保护现有数据，已停止使用。")
        elif user_version != 0:
            raise HistoryStoreError("历史资料库缺少必要数据表，为保护现有数据，已停止使用。")
        row = connection.execute("PRAGMA quick_check(1)").fetchone()
        if row is None or str(row[0]).lower() != "ok":
            detail = "未知错误" if row is None else str(row[0])
            raise _CorruptHistoryDatabase(detail)

    def _quick_check_database(self) -> None:
        with self._connect() as connection:
            row = connection.execute("PRAGMA quick_check(1)").fetchone()
        if row is None or str(row[0]).lower() != "ok":
            detail = "未知错误" if row is None else str(row[0])
            raise _CorruptHistoryDatabase(detail)

    def _recover_corrupt_database(self) -> None:
        backup_dir = self._backup_database_artifacts()
        self._begin_database_recovery(backup_dir)

    def _recover_missing_database(self) -> None:
        backup_dir = self._backup_missing_database()
        self._begin_database_recovery(backup_dir)

    def _begin_database_recovery(self, backup_dir: Path) -> None:
        recovery_marker = self.root / DATABASE_RECOVERY_PENDING_NAME
        rebuild_name = f".history-rebuild-{uuid.uuid4().hex}.db"
        payload = {
            "schema_version": SCHEMA_VERSION,
            "backup_relpath": backup_dir.relative_to(self.root).as_posix(),
            "rebuild_name": rebuild_name,
            "created_at": _utc_now(),
        }
        _atomic_write_text(recovery_marker, _json_dumps(payload, pretty=True) + "\n")
        _make_private(recovery_marker)
        self._rebuild_and_publish_database(backup_dir, recovery_marker, self.root / rebuild_name)

    def _has_archive_evidence(self) -> bool:
        for archive_dir in (self.records_dir, self.trash_dir):
            try:
                if next(archive_dir.iterdir(), None) is not None:
                    return True
            except OSError as exc:
                raise HistoryStoreError(f"无法核对历史资料目录：{exc}") from exc
        return False

    def _resume_pending_database_recovery(self, recovery_marker: Path) -> None:
        try:
            payload = json.loads(recovery_marker.read_text(encoding="utf-8"))
            if not isinstance(payload, dict) or payload.get("schema_version") != SCHEMA_VERSION:
                raise ValueError("恢复标记版本无效")
            backup_relpath = Path(str(payload["backup_relpath"]))
            if backup_relpath.is_absolute() or ".." in backup_relpath.parts:
                raise ValueError("恢复备份路径无效")
            backup_dir = _safe_join(self.root, backup_relpath)
            if not _is_within(backup_dir, self.database_backups_dir):
                raise ValueError("恢复备份不属于历史资料库")
            rebuild_name = str(payload["rebuild_name"])
            if not _is_rebuild_database_name(rebuild_name):
                raise ValueError("恢复临时索引名无效")
            rebuild_path = self.root / rebuild_name
            self._verify_database_backup(backup_dir)
        except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
            raise HistoryStoreError(f"历史索引恢复标记无法核对，为保护资料已停止恢复：{exc}") from exc
        self._rebuild_and_publish_database(backup_dir, recovery_marker, rebuild_path)

    def _rebuild_and_publish_database(
        self,
        backup_dir: Path,
        recovery_marker: Path,
        rebuild_path: Path,
    ) -> None:
        official_database = self.database_path
        try:
            self._verify_database_backup(backup_dir)
            for path in _database_artifact_paths(rebuild_path):
                if not path.exists():
                    continue
                _assert_storage_entry(path, "临时历史索引", kind="file", allow_missing=False)
                path.unlink()

            self.database_path = rebuild_path
            try:
                self._initialize_database()
                restored = self.rebuild_index_from_manifests()
                discovered = int(self.last_rebuild_report.get("discovered", 0))
                validated = int(self.last_rebuild_report.get("validated", 0))
                scanned = int(self.last_rebuild_report.get("scanned", 0))
                errors = list(self.last_rebuild_report.get("errors", []))
                missing_manifests = list(self.last_rebuild_report.get("missing_manifests", []))
                unexpected_entries = list(self.last_rebuild_report.get("unexpected_entries", []))
                if missing_manifests or unexpected_entries:
                    raise HistoryStoreError("部分历史任务缺少清单或目录结构异常，已停止发布新索引。")
                if errors or discovered != scanned or validated != discovered:
                    raise HistoryStoreError("部分历史清单未通过完整校验，已停止发布新索引。")
                if discovered > 0 and restored == 0:
                    raise HistoryStoreError("没有任何历史清单通过安全校验，已停止发布新索引。")
                self._quick_check_database()
                with self._connect() as connection:
                    foreign_keys = connection.execute("PRAGMA foreign_key_check").fetchall()
                    if foreign_keys:
                        raise HistoryStoreError("重建后的历史索引关联校验失败。")
                    connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")
                    connection.execute("PRAGMA journal_mode = DELETE")
            finally:
                self.database_path = official_database

            if not rebuild_path.is_file() or _is_link_like(rebuild_path):
                raise HistoryStoreError("临时历史索引未完整生成。")
            for sidecar in _database_artifact_paths(rebuild_path)[1:]:
                if sidecar.exists():
                    _assert_storage_entry(sidecar, "临时历史索引", kind="file", allow_missing=False)
                    sidecar.unlink()
            self._publish_rebuilt_database(rebuild_path, official_database)
            _make_private(official_database)
            self._initialize_database()
            self._quick_check_database()
            report_payload = {
                **self.last_rebuild_report,
                "backup_relpath": backup_dir.relative_to(self.root).as_posix(),
                "completed_at": _utc_now(),
            }
            _atomic_write_text(
                backup_dir / "recovery-report.json",
                _json_dumps(report_payload, pretty=True) + "\n",
            )
            recovery_marker.unlink()
        except Exception as exc:
            self.database_path = official_database
            raise HistoryStoreError(
                f"历史索引已损坏，原索引已安全备份在 {backup_dir.name}，但自动整理失败：{exc}"
            ) from exc
        self.recovered_database_backup = backup_dir

    def _publish_rebuilt_database(self, rebuild_path: Path, official_database: Path) -> None:
        if (
            rebuild_path.parent != self.root
            or not _is_rebuild_database_name(rebuild_path.name)
            or official_database != self.root / DATABASE_NAME
        ):
            raise HistoryStoreError("历史索引发布路径无效。")
        for sidecar in _database_artifact_paths(official_database)[1:]:
            if not sidecar.exists():
                continue
            _assert_storage_entry(sidecar, "历史数据库", kind="file", allow_missing=False)
            sidecar.unlink()
        # os.replace semantics keep the old main database present until this
        # final same-filesystem operation succeeds.
        rebuild_path.replace(official_database)

    def _backup_database_artifacts(self) -> Path:
        _mkdir_private(self.database_backups_dir)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        token = uuid.uuid4().hex
        capture_dir = self.database_backups_dir / f".capturing-{token}"
        backup_dir = self.database_backups_dir / f"损坏索引_{stamp}_{token[:8]}"
        capture_dir.mkdir(mode=0o700, exist_ok=False)
        _make_private(capture_dir, directory=True)
        entries: list[dict[str, Any]] = []
        try:
            source_paths = [path for path in _database_artifact_paths(self.database_path) if path.exists()]
            before_stats = {
                path.name: (path.stat().st_size, path.stat().st_mtime_ns)
                for path in source_paths
            }
            for source in _database_artifact_paths(self.database_path):
                if not source.exists():
                    continue
                _assert_storage_entry(source, "历史数据库", kind="file", allow_missing=False)
                destination = capture_dir / source.name
                size_bytes, digest = _copy_file_durable(source, destination)
                entries.append({"name": source.name, "size_bytes": size_bytes, "sha256": digest})
            if not entries:
                raise HistoryStoreError("没有找到需要恢复的历史索引文件。")
            after_paths = [path for path in _database_artifact_paths(self.database_path) if path.exists()]
            after_stats = {
                path.name: (path.stat().st_size, path.stat().st_mtime_ns)
                for path in after_paths
            }
            if before_stats != after_stats:
                raise HistoryStoreError("备份期间历史索引仍在变化，请关闭其他 HRToolkit 窗口后重试。")
            manifest = {
                "schema_version": SCHEMA_VERSION,
                "created_at": _utc_now(),
                "reason": "database_corrupt",
                "files": entries,
            }
            _atomic_write_text(
                capture_dir / DATABASE_BACKUP_MANIFEST_NAME,
                _json_dumps(manifest, pretty=True) + "\n",
            )
            self._verify_database_backup(capture_dir)
            capture_dir.replace(backup_dir)
            self._verify_database_backup(backup_dir)
        except Exception as exc:
            raise HistoryStoreError(f"无法安全备份损坏的历史索引：{exc}") from exc
        return backup_dir

    def _backup_missing_database(self) -> Path:
        _mkdir_private(self.database_backups_dir)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        token = uuid.uuid4().hex
        capture_dir = self.database_backups_dir / f".capturing-{token}"
        backup_dir = self.database_backups_dir / f"缺失索引_{stamp}_{token[:8]}"
        capture_dir.mkdir(mode=0o700, exist_ok=False)
        _make_private(capture_dir, directory=True)
        try:
            manifest = {
                "schema_version": SCHEMA_VERSION,
                "created_at": _utc_now(),
                "reason": "database_missing",
                "files": [],
            }
            _atomic_write_text(
                capture_dir / DATABASE_BACKUP_MANIFEST_NAME,
                _json_dumps(manifest, pretty=True) + "\n",
            )
            self._verify_database_backup(capture_dir)
            capture_dir.replace(backup_dir)
            self._verify_database_backup(backup_dir)
        except Exception as exc:
            raise HistoryStoreError(f"无法记录历史索引缺失状态：{exc}") from exc
        return backup_dir

    def _verify_database_backup(self, backup_dir: Path) -> None:
        if _is_link_like(backup_dir) or not backup_dir.is_dir():
            raise HistoryStoreError("损坏索引备份目录不存在或不安全。")
        manifest_path = backup_dir / DATABASE_BACKUP_MANIFEST_NAME
        if _is_link_like(manifest_path) or not manifest_path.is_file():
            raise HistoryStoreError("损坏索引备份清单不存在或不安全。")
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict) or payload.get("schema_version") != SCHEMA_VERSION:
            raise HistoryStoreError("损坏索引备份清单版本无效。")
        reason = str(payload.get("reason") or "database_corrupt")
        if reason not in {"database_corrupt", "database_missing"}:
            raise HistoryStoreError("损坏索引备份原因无效。")
        raw_files = payload.get("files")
        if not isinstance(raw_files, list):
            raise HistoryStoreError("损坏索引备份清单内容无效。")
        if not raw_files and reason != "database_missing":
            raise HistoryStoreError("损坏索引备份清单为空。")
        if raw_files and reason == "database_missing":
            raise HistoryStoreError("缺失索引备份不应包含数据库文件。")
        allowed_names = {path.name for path in _database_artifact_paths(self.database_path)}
        seen: set[str] = set()
        for item in raw_files:
            if not isinstance(item, dict):
                raise HistoryStoreError("损坏索引备份清单内容无效。")
            name = str(item.get("name") or "")
            if name not in allowed_names or name in seen:
                raise HistoryStoreError("损坏索引备份文件名无效。")
            seen.add(name)
            path = backup_dir / name
            if _is_link_like(path) or not path.is_file():
                raise HistoryStoreError("损坏索引备份文件不存在或不安全。")
            expected_size = int(item["size_bytes"])
            expected_digest = str(item["sha256"])
            if path.stat().st_size != expected_size or _sha256_file(path) != expected_digest:
                raise HistoryStoreError("损坏索引备份校验失败。")

    def start_task(
        self,
        *,
        tool_id: str,
        tool_name: str,
        app_version: str,
        mode: str | None = None,
        parameters: dict[str, Any] | None = None,
    ) -> str:
        task_id = uuid.uuid4().hex
        now = _utc_now()
        local_now = datetime.now().astimezone()
        relpath = Path("records") / local_now.strftime("%Y") / local_now.strftime("%m") / (
            local_now.strftime("%Y%m%d_%H%M%S_") + task_id
        )
        task_dir = _safe_join(self.root, relpath)
        created_task_dir = False
        try:
            _mkdir_private(task_dir.parent)
            task_dir.mkdir(mode=0o700, exist_ok=False)
            created_task_dir = True
            _make_private(task_dir, directory=True)
            _mkdir_private(task_dir / "inputs")
            _mkdir_private(task_dir / "outputs")
            with self._connect() as connection:
                connection.execute(
                    """
                    INSERT INTO tasks (
                        id, tool_id, tool_name, mode, app_version, status, started_at,
                        parameters_json, owner_pid, task_relpath
                    ) VALUES (?, ?, ?, ?, ?, 'running', ?, ?, ?, ?)
                    """,
                    (
                        task_id,
                        str(tool_id),
                        str(tool_name),
                        None if mode is None else str(mode),
                        str(app_version),
                        now,
                        _json_dumps(parameters or {}),
                        os.getpid(),
                        relpath.as_posix(),
                    ),
                )
        except Exception:
            if created_task_dir:
                shutil.rmtree(task_dir, ignore_errors=True)
            raise
        try:
            self._write_manifest(task_id)
        except Exception:
            removed_from_index = False
            try:
                with self._connect() as connection:
                    cursor = connection.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
                    removed_from_index = cursor.rowcount > 0
            except Exception:
                # If the rollback itself cannot be written (for example the
                # disk is completely full), make a best-effort transition so
                # this process does not own an invisible running record.
                try:
                    with self._connect() as connection:
                        connection.execute(
                            """
                            UPDATE tasks
                            SET status = 'failed', finished_at = ?, owner_pid = NULL,
                                error_message = '无法建立历史清单，本次处理没有开始。'
                            WHERE id = ? AND status = 'running'
                            """,
                            (_utc_now(), task_id),
                        )
                except Exception:
                    pass
            if removed_from_index and created_task_dir:
                shutil.rmtree(task_dir, ignore_errors=True)
            raise
        return task_id

    def archive_sources(
        self,
        task_id: str,
        sources: Iterable[SourceSpec],
        *,
        cancelled: Callable[[], bool] | None = None,
    ) -> tuple[FileRecord, ...]:
        with self._task_io_lock(task_id):
            lock_path = self._archive_lock_path(task_id)
            with _exclusive_file_lock(lock_path):
                return self._archive_sources_locked(task_id, sources, cancelled=cancelled)

    def _archive_sources_locked(
        self,
        task_id: str,
        sources: Iterable[SourceSpec],
        *,
        cancelled: Callable[[], bool] | None = None,
    ) -> tuple[FileRecord, ...]:
        task = self.get_task(task_id)
        if task is None:
            raise HistoryStoreError("历史记录不存在。")
        if task.summary.status != "running":
            raise HistoryStoreError("本次历史记录已经结束，不能继续写入资料。")
        if not _is_safe_active_task_storage(task, self.records_dir):
            raise HistoryStoreError("历史记录目录不安全，已停止写入。")
        source_items: list[tuple[Path, str, tuple[str, ...]]] = []
        for item in self._iter_source_items(sources):
            _raise_if_cancelled(cancelled)
            source_items.append(item)
        required_bytes = sum(item[0].stat().st_size for item in source_items)
        self._ensure_free_space(required_bytes)
        destination_root = task.input_dir
        copied: list[dict[str, Any]] = []
        try:
            for source_path, role, relative_parts in source_items:
                _raise_if_cancelled(cancelled)
                role_dir = _safe_join(destination_root, Path(_safe_component(role)))
                _mkdir_private(role_dir)
                destination = _safe_join(role_dir, Path(*(_safe_component(part) for part in relative_parts)))
                destination = _unique_destination(destination)
                copied.append(
                    self._copy_and_hash(
                        source_path,
                        destination,
                        task_id,
                        "input",
                        role,
                        cancelled=cancelled,
                    )
                )
            self._write_manifest(task_id, pending_items=copied)
            records = self._insert_files(task_id, copied)
        except Exception:
            for item in copied:
                Path(item["archived_path"]).unlink(missing_ok=True)
            try:
                self._write_manifest(task_id)
            except Exception:
                pass
            raise
        self._write_manifest(task_id)
        return records

    def archive_output_directory(
        self,
        task_id: str,
        source_dir: str | Path,
        *,
        cancelled: Callable[[], bool] | None = None,
    ) -> tuple[FileRecord, ...]:
        with self._task_io_lock(task_id):
            lock_path = self._archive_lock_path(task_id)
            with _exclusive_file_lock(lock_path):
                return self._archive_output_directory_locked(
                    task_id,
                    source_dir,
                    cancelled=cancelled,
                )

    def _archive_output_directory_locked(
        self,
        task_id: str,
        source_dir: str | Path,
        *,
        cancelled: Callable[[], bool] | None = None,
    ) -> tuple[FileRecord, ...]:
        task = self.get_task(task_id)
        if task is None:
            raise HistoryStoreError("历史记录不存在。")
        if task.summary.status != "running":
            raise HistoryStoreError("本次历史记录已经结束，不能继续写入资料。")
        if not _is_safe_active_task_storage(task, self.records_dir):
            raise HistoryStoreError("历史记录目录不安全，已停止写入。")
        source_root = Path(source_dir).expanduser()
        if _is_link_like(source_root) or not source_root.is_dir():
            raise HistoryStoreError("结果文件夹不存在或不安全。")
        source_root = source_root.resolve()
        if _paths_overlap(source_root, self.root):
            raise HistoryStoreError("不能把资料库自身作为结果来源。")
        source_files: list[Path] = []
        for path in _walk_regular_files(source_root):
            _raise_if_cancelled(cancelled)
            if path.name != MANIFEST_NAME and not path.name.endswith(".partial"):
                source_files.append(path)
        required_bytes = sum(path.stat().st_size for path in source_files)
        self._ensure_free_space(required_bytes)
        copied: list[dict[str, Any]] = []
        try:
            for source_path in source_files:
                _raise_if_cancelled(cancelled)
                relative = source_path.relative_to(source_root)
                destination = _safe_join(
                    task.output_dir,
                    Path(*(_safe_component(part) for part in relative.parts)),
                )
                destination = _unique_destination(destination)
                copied.append(
                    self._copy_and_hash(
                        source_path,
                        destination,
                        task_id,
                        "output",
                        "result",
                        cancelled=cancelled,
                    )
                )
            self._write_manifest(task_id, pending_items=copied)
            records = self._insert_files(task_id, copied)
        except Exception:
            for item in copied:
                Path(item["archived_path"]).unlink(missing_ok=True)
            try:
                self._write_manifest(task_id)
            except Exception:
                pass
            raise
        self._write_manifest(task_id)
        return records

    def mark_success(self, task_id: str, result: dict[str, Any] | None = None) -> bool:
        return self._finish_task(task_id, "success", result=result)

    def mark_failed(self, task_id: str, error: str, result: dict[str, Any] | None = None) -> bool:
        return self._finish_task(task_id, "failed", result=result, error=error)

    def mark_stopped(self, task_id: str) -> bool:
        return self._finish_task(task_id, "stopped", error="用户停止了本次处理。")

    def _finish_task(
        self,
        task_id: str,
        status: str,
        *,
        result: dict[str, Any] | None = None,
        error: str | None = None,
    ) -> bool:
        if status not in TERMINAL_STATUSES:
            raise ValueError(f"不支持的历史状态：{status}")
        with self._task_lock(task_id):
            finished_at = _utc_now()
            with self._connect() as connection:
                previous = connection.execute(
                    """
                    SELECT status, finished_at, result_json, error_message, owner_pid
                    FROM tasks WHERE id = ?
                    """,
                    (task_id,),
                ).fetchone()
                if previous is None:
                    return False
                cursor = connection.execute(
                    """
                    UPDATE tasks
                    SET status = ?, finished_at = ?, result_json = ?, error_message = ?, owner_pid = NULL
                    WHERE id = ? AND status = 'running'
                    """,
                    (
                        status,
                        finished_at,
                        _json_dumps(result or {}),
                        _clean_error(error),
                        task_id,
                    ),
                )
                changed = cursor.rowcount > 0
            try:
                self._write_manifest(task_id)
            except Exception:
                if changed:
                    try:
                        with self._connect() as connection:
                            connection.execute(
                                """
                                UPDATE tasks
                                SET status = ?, finished_at = ?, result_json = ?,
                                    error_message = ?, owner_pid = ?
                                WHERE id = ? AND status = ? AND finished_at = ?
                                """,
                                (
                                    str(previous["status"]),
                                    previous["finished_at"],
                                    str(previous["result_json"]),
                                    previous["error_message"],
                                    previous["owner_pid"],
                                    task_id,
                                    status,
                                    finished_at,
                                ),
                            )
                    except Exception:
                        pass
                raise
            return changed

    def recover_incomplete_tasks(self) -> int:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT id, owner_pid FROM tasks WHERE status = 'running'"
            ).fetchall()
            stale_ids = [
                str(row["id"])
                for row in rows
                if row["owner_pid"] is None or not _pid_is_alive(int(row["owner_pid"]))
            ]
            if not stale_ids:
                return 0
        for task_id in stale_ids:
            try:
                self._recover_stale_task_files(task_id)
            except Exception:
                # Recovery is best effort; the task must still leave "running"
                # so HR never sees an abandoned run as active forever.
                pass
        with self._connect() as connection:
            placeholders = ",".join("?" for _ in stale_ids)
            connection.execute(
                f"""
                UPDATE tasks
                SET status = 'failed', finished_at = ?,
                    error_message = '程序上次意外关闭，本次处理未正常完成。'
                WHERE status = 'running' AND id IN ({placeholders})
                """,
                (_utc_now(), *stale_ids),
            )
        for task_id in stale_ids:
            self._cleanup_task_staging(task_id)
            self._write_manifest(task_id)
        return len(stale_ids)

    def _recover_stale_task_files(self, task_id: str) -> int:
        task = self.get_task(task_id)
        if task is None or not _is_valid_record_task_dir(task.task_dir, self.records_dir, task_id):
            return 0
        manifest_items: dict[str, dict[str, Any]] = {}
        manifest_path = task.task_dir / MANIFEST_NAME
        try:
            payload = json.loads(manifest_path.read_text(encoding="utf-8"))
            if payload.get("schema_version") == SCHEMA_VERSION and str(payload["task"]["id"]) == task_id:
                for item in payload.get("files", []):
                    if not isinstance(item, dict):
                        continue
                    relative_path = Path(str(item.get("relative_path") or ""))
                    kind = str(item.get("kind") or "")
                    if relative_path.is_absolute() or ".." in relative_path.parts or kind not in {"input", "output"}:
                        continue
                    archived_path = _safe_join(self.root, relative_path)
                    expected_root = task.input_dir if kind == "input" else task.output_dir
                    if not _is_within(archived_path, expected_root):
                        continue
                    manifest_items[relative_path.as_posix()] = item
        except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError):
            manifest_items = {}

        with self._connect() as connection:
            existing_paths = {
                str(row["relative_path"])
                for row in connection.execute(
                    "SELECT relative_path FROM files WHERE task_id = ?",
                    (task_id,),
                ).fetchall()
            }

        recovered: list[dict[str, Any]] = []
        for kind, archive_root in (("input", task.input_dir), ("output", task.output_dir)):
            if not archive_root.is_dir() or _is_link_like(archive_root):
                continue
            for archived_path in _walk_regular_files(archive_root):
                if _is_owned_staging_name(archived_path.name):
                    continue
                relative_path = archived_path.relative_to(self.root).as_posix()
                if relative_path in existing_paths:
                    continue
                relative_to_archive = archived_path.relative_to(archive_root)
                manifest_item = manifest_items.get(relative_path, {})
                role = str(manifest_item.get("role") or "")
                if not role:
                    role = relative_to_archive.parts[0] if kind == "input" and len(relative_to_archive.parts) > 1 else "result"
                file_stat = archived_path.stat()
                recovered.append(
                    {
                        "kind": kind,
                        "role": role[:100],
                        "display_name": Path(str(manifest_item.get("display_name") or archived_path.name)).name,
                        "original_path": Path(str(manifest_item.get("original_path") or archived_path.name)).name,
                        "relative_path": relative_path,
                        "size_bytes": file_stat.st_size,
                        "sha256": _sha256_file(archived_path),
                        "modified_ns": file_stat.st_mtime_ns,
                    }
                )
        if not recovered:
            return 0
        with self._connect() as connection:
            for item in recovered:
                connection.execute(
                    """
                    INSERT OR IGNORE INTO files (
                        task_id, kind, role, display_name, original_path,
                        relative_path, size_bytes, sha256, modified_ns
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        task_id,
                        item["kind"],
                        item["role"],
                        item["display_name"],
                        item["original_path"],
                        item["relative_path"],
                        item["size_bytes"],
                        item["sha256"],
                        item["modified_ns"],
                    ),
                )
        return len(recovered)

    def _cleanup_task_staging(self, task_id: str) -> None:
        task = self.get_task(task_id)
        if task is None or not _is_valid_record_task_dir(task.task_dir, self.records_dir, task_id):
            return
        registered_paths = {
            item.archived_path.resolve()
            for item in (*task.inputs, *task.outputs)
        }
        for path in task.task_dir.rglob("*"):
            if (
                _is_owned_staging_name(path.name)
                and path.resolve() not in registered_paths
                and (path.is_file() or path.is_symlink())
            ):
                try:
                    path.unlink()
                except OSError:
                    pass

    def _cleanup_recent_terminal_staging(self, limit: int = 100) -> None:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT id FROM tasks
                WHERE deleted_at IS NULL AND status IN ('success', 'failed', 'stopped')
                ORDER BY finished_at DESC
                LIMIT ?
                """,
                (max(1, int(limit)),),
            ).fetchall()
        for row in rows:
            self._cleanup_task_staging(str(row["id"]))

    def cleanup_task_staging(self, task_id: str) -> None:
        """Remove only HR Toolkit-owned temporary copies after a worker exits."""
        with self._task_io_lock(task_id):
            lock_path = self._archive_lock_path(task_id, require_running=False)
            with _exclusive_file_lock(lock_path):
                self._cleanup_task_staging(task_id)

    def list_tasks(
        self,
        *,
        search: str = "",
        tool_id: str | None = None,
        started_after: str | None = None,
        limit: int = HISTORY_PAGE_SIZE,
        offset: int = 0,
    ) -> tuple[tuple[TaskSummary, ...], int]:
        limit = max(1, min(int(limit), 200))
        offset = max(0, int(offset))
        conditions = ["tasks.deleted_at IS NULL"]
        parameters: list[Any] = []
        if tool_id:
            conditions.append("tasks.tool_id = ?")
            parameters.append(tool_id)
        if started_after:
            conditions.append("tasks.started_at >= ?")
            parameters.append(started_after)
        query = search.strip()
        if query:
            pattern = f"%{query}%"
            conditions.append(
                "(tasks.tool_name LIKE ? OR EXISTS ("
                "SELECT 1 FROM files WHERE files.task_id = tasks.id AND files.display_name LIKE ?))"
            )
            parameters.extend((pattern, pattern))
        where = " AND ".join(conditions)
        with self._connect() as connection:
            total = int(connection.execute(f"SELECT COUNT(*) FROM tasks WHERE {where}", parameters).fetchone()[0])
            rows = connection.execute(
                f"""
                SELECT * FROM tasks
                WHERE {where}
                ORDER BY started_at DESC, id DESC
                LIMIT ? OFFSET ?
                """,
                (*parameters, limit, offset),
            ).fetchall()
            names = self._file_names_for_tasks(connection, [str(row["id"]) for row in rows])
        summaries = tuple(self._summary_from_row(row, names.get(str(row["id"]), {})) for row in rows)
        return summaries, total

    def get_task(self, task_id: str) -> TaskDetail | None:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM tasks WHERE id = ?", (task_id,)).fetchone()
            if row is None:
                return None
            file_rows = connection.execute(
                "SELECT * FROM files WHERE task_id = ? ORDER BY kind, id",
                (task_id,),
            ).fetchall()
        files = tuple(self._file_from_row(file_row) for file_row in file_rows)
        task_relpath = Path(str(row["task_relpath"]))
        task_dir = _safe_join(self.root, task_relpath)
        names = {
            "input": tuple(item.display_name for item in files if item.kind == "input"),
            "output": tuple(item.display_name for item in files if item.kind == "output"),
        }
        summary = self._summary_from_row(row, names)
        return TaskDetail(
            summary=summary,
            parameters=_json_object(row["parameters_json"]),
            result=_json_object(row["result_json"]),
            task_dir=task_dir,
            input_dir=task_dir / "inputs",
            output_dir=task_dir / "outputs",
            inputs=tuple(item for item in files if item.kind == "input"),
            outputs=tuple(item for item in files if item.kind == "output"),
        )

    def move_to_trash(self, task_id: str) -> Path:
        io_lock = self._task_io_lock(task_id)
        if not io_lock.acquire(blocking=False):
            raise HistoryStoreError("资料仍在整理中，请稍后再试。")
        try:
            lock_path = self._archive_lock_path(task_id, require_running=False)
            try:
                with _exclusive_file_lock(lock_path, blocking=False):
                    return self._move_to_trash_locked(task_id)
            except BlockingIOError as exc:
                raise HistoryStoreError("资料仍在整理中，请稍后再试。") from exc
        finally:
            io_lock.release()

    def _trash_move_marker_path(self, task_id: str) -> Path:
        marker = self.root / f"{TRASH_MOVE_PENDING_PREFIX}{task_id}.json"
        if not _is_trash_move_marker_name(marker.name):
            raise HistoryStoreError("历史任务编号无效。")
        return marker

    def _normalize_pending_trash_moves_without_database(self) -> None:
        """Make a pending move's manifest and directory agree before index rebuild."""
        for marker_path in sorted(self.root.glob(f"{TRASH_MOVE_PENDING_PREFIX}*")):
            if not _is_trash_move_marker_name(marker_path.name):
                raise HistoryStoreError("回收站移动恢复标记名称无效。")
            _assert_storage_entry(marker_path, "回收站移动恢复标记", kind="file", allow_missing=False)
            try:
                payload = json.loads(marker_path.read_text(encoding="utf-8"))
                if not isinstance(payload, dict) or payload.get("schema_version") != SCHEMA_VERSION:
                    raise ValueError("恢复标记版本无效")
                task_id = str(payload["task_id"]).lower()
                if marker_path != self._trash_move_marker_path(task_id):
                    raise ValueError("恢复标记任务编号不一致")
                old_relative = Path(str(payload["old_relpath"]))
                new_relative = Path(str(payload["new_relpath"]))
                if old_relative.is_absolute() or new_relative.is_absolute():
                    raise ValueError("恢复路径无效")
                old_dir = _safe_join(self.root, old_relative)
                new_dir = _safe_join(self.root, new_relative)
                if not _is_valid_record_task_dir(old_dir, self.records_dir, task_id, allow_missing=True):
                    raise ValueError("原历史目录无效")
                if not _is_valid_trash_task_dir(new_dir, self.trash_dir, task_id, allow_missing=True):
                    raise ValueError("回收站目录无效")
                old_exists = old_dir.is_dir() and not _is_link_like(old_dir)
                new_exists = new_dir.is_dir() and not _is_link_like(new_dir)
                if old_exists and not new_exists:
                    continue
                if not new_exists or old_exists:
                    raise ValueError("历史目录位置无法唯一确认")
                manifest_location = _pending_manifest_location(
                    new_dir / MANIFEST_NAME,
                    task_id,
                    old_relative,
                    new_relative,
                )
                if manifest_location == "old":
                    _mkdir_private(old_dir.parent)
                    new_dir.replace(old_dir)
                elif manifest_location != "new":
                    raise ValueError("历史清单位置无法确认")
            except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
                raise HistoryStoreError(f"上次移入回收站的文件无法在索引恢复前核对：{exc}") from exc

    def _recover_pending_trash_moves(self) -> None:
        for marker_path in sorted(self.root.glob(f"{TRASH_MOVE_PENDING_PREFIX}*")):
            if not _is_trash_move_marker_name(marker_path.name):
                raise HistoryStoreError("回收站移动恢复标记名称无效。")
            _assert_storage_entry(marker_path, "回收站移动恢复标记", kind="file", allow_missing=False)
            try:
                payload = json.loads(marker_path.read_text(encoding="utf-8"))
                if not isinstance(payload, dict) or payload.get("schema_version") != SCHEMA_VERSION:
                    raise ValueError("恢复标记版本无效")
                task_id = str(payload["task_id"]).lower()
                if marker_path != self._trash_move_marker_path(task_id):
                    raise ValueError("恢复标记任务编号不一致")
                old_relative = Path(str(payload["old_relpath"]))
                new_relative = Path(str(payload["new_relpath"]))
                if old_relative.is_absolute() or new_relative.is_absolute():
                    raise ValueError("恢复路径无效")
                old_dir = _safe_join(self.root, old_relative)
                new_dir = _safe_join(self.root, new_relative)
                if not _is_valid_record_task_dir(old_dir, self.records_dir, task_id, allow_missing=True):
                    raise ValueError("原历史目录无效")
                if not _is_valid_trash_task_dir(new_dir, self.trash_dir, task_id, allow_missing=True):
                    raise ValueError("回收站目录无效")
                with self._connect() as connection:
                    row = connection.execute(
                        "SELECT task_relpath, deleted_at FROM tasks WHERE id = ?",
                        (task_id,),
                    ).fetchone()
                if row is None:
                    raise ValueError("历史任务索引不存在")
                indexed_path = str(row["task_relpath"])
                deleted_at = row["deleted_at"]
                old_exists = old_dir.is_dir() and not _is_link_like(old_dir)
                new_exists = new_dir.is_dir() and not _is_link_like(new_dir)
                if old_exists == new_exists:
                    raise ValueError("历史目录位置无法唯一确认")
                if new_exists and indexed_path == new_relative.as_posix() and deleted_at is not None:
                    _verify_manifest_task_identity(new_dir / MANIFEST_NAME, task_id)
                    self._write_manifest(task_id)
                elif new_exists and indexed_path == old_relative.as_posix() and deleted_at is None:
                    _verify_manifest_task_identity(new_dir / MANIFEST_NAME, task_id)
                    _mkdir_private(old_dir.parent)
                    new_dir.replace(old_dir)
                elif old_exists and indexed_path == old_relative.as_posix() and deleted_at is None:
                    _verify_manifest_task_identity(old_dir / MANIFEST_NAME, task_id)
                else:
                    raise ValueError("历史索引与文件位置不一致")
                marker_path.unlink()
            except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
                raise HistoryStoreError(f"上次移入回收站的操作无法安全续做：{exc}") from exc

    def _move_to_trash_locked(self, task_id: str) -> Path:
        with self._task_lock(task_id):
            task = self.get_task(task_id)
            if task is None:
                raise HistoryStoreError("历史记录不存在。")
            if task.summary.status == "running":
                raise HistoryStoreError("正在处理的记录不能移到回收站。")
            if not _is_valid_record_task_dir(task.task_dir, self.records_dir, task_id):
                raise HistoryStoreError("历史记录目录无效，为保护其他资料，已停止移动。")
            manifest_path = task.task_dir / MANIFEST_NAME
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                manifest_task_id = str(manifest["task"]["id"])
            except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
                raise HistoryStoreError("历史记录清单无法核对，为保护其他资料，已停止移动。") from exc
            if manifest.get("schema_version") != SCHEMA_VERSION or manifest_task_id != task_id:
                raise HistoryStoreError("历史记录清单与任务不一致，为保护其他资料，已停止移动。")
            target = _unique_destination(self.trash_dir / task.task_dir.name)
            _mkdir_private(target.parent)
            old_prefix = task.task_dir.relative_to(self.root).as_posix()
            new_prefix = target.relative_to(self.root).as_posix()
            move_marker = self._trash_move_marker_path(task_id)
            _atomic_write_text(
                move_marker,
                _json_dumps(
                    {
                        "schema_version": SCHEMA_VERSION,
                        "task_id": task_id,
                        "old_relpath": old_prefix,
                        "new_relpath": new_prefix,
                        "created_at": _utc_now(),
                    },
                    pretty=True,
                )
                + "\n",
            )
            try:
                task.task_dir.replace(target)
            except Exception:
                move_marker.unlink(missing_ok=True)
                raise
            database_updated = False
            try:
                with self._connect() as connection:
                    connection.execute(
                        "UPDATE tasks SET deleted_at = ?, task_relpath = ? WHERE id = ?",
                        (_utc_now(), new_prefix, task_id),
                    )
                    rows = connection.execute(
                        "SELECT id, relative_path FROM files WHERE task_id = ?",
                        (task_id,),
                    ).fetchall()
                    for row in rows:
                        relative_path = str(row["relative_path"])
                        if relative_path == old_prefix or relative_path.startswith(old_prefix + "/"):
                            updated_path = new_prefix + relative_path[len(old_prefix) :]
                            connection.execute(
                                "UPDATE files SET relative_path = ? WHERE id = ?",
                                (updated_path, int(row["id"])),
                            )
                database_updated = True
                self._write_manifest(task_id)
            except Exception as exc:
                rollback_errors: list[str] = []
                if database_updated:
                    try:
                        with self._connect() as connection:
                            connection.execute(
                                "UPDATE tasks SET deleted_at = NULL, task_relpath = ? WHERE id = ?",
                                (old_prefix, task_id),
                            )
                            rows = connection.execute(
                                "SELECT id, relative_path FROM files WHERE task_id = ?",
                                (task_id,),
                            ).fetchall()
                            for row in rows:
                                relative_path = str(row["relative_path"])
                                if relative_path == new_prefix or relative_path.startswith(new_prefix + "/"):
                                    restored_path = old_prefix + relative_path[len(new_prefix) :]
                                    connection.execute(
                                        "UPDATE files SET relative_path = ? WHERE id = ?",
                                        (restored_path, int(row["id"])),
                                    )
                    except Exception as rollback_exc:
                        rollback_errors.append(str(rollback_exc))
                if not rollback_errors:
                    try:
                        if target.exists() and not task.task_dir.exists():
                            target.replace(task.task_dir)
                    except Exception as rollback_exc:
                        rollback_errors.append(str(rollback_exc))
                if rollback_errors:
                    raise HistoryStoreError(
                        "移入回收站未完成，自动恢复也遇到问题；资料仍保留，请关闭其他窗口后重试。"
                    ) from exc
                move_marker.unlink(missing_ok=True)
                raise
            move_marker.unlink(missing_ok=True)
            return target

    def _discover_archive_manifests(self) -> tuple[list[tuple[Path, bool]], list[str], list[str]]:
        manifest_entries: list[tuple[Path, bool]] = []
        missing_manifests: list[str] = []
        unexpected_entries: list[str] = []

        def relative_name(path: Path) -> str:
            try:
                return path.absolute().relative_to(self.root.absolute()).as_posix()
            except ValueError:
                return path.name

        def children(path: Path) -> list[Path]:
            try:
                return sorted(path.iterdir(), key=lambda item: item.name.casefold())
            except OSError as exc:
                raise HistoryStoreError(f"无法读取历史资料目录 {relative_name(path)}：{exc}") from exc

        for year_dir in children(self.records_dir):
            if _is_ignorable_archive_metadata(year_dir):
                continue
            if (
                _is_link_like(year_dir)
                or not year_dir.is_dir()
                or len(year_dir.name) != 4
                or not year_dir.name.isdigit()
            ):
                unexpected_entries.append(relative_name(year_dir))
                continue
            for month_dir in children(year_dir):
                if _is_ignorable_archive_metadata(month_dir):
                    continue
                if (
                    _is_link_like(month_dir)
                    or not month_dir.is_dir()
                    or len(month_dir.name) != 2
                    or not month_dir.name.isdigit()
                    or not 1 <= int(month_dir.name) <= 12
                ):
                    unexpected_entries.append(relative_name(month_dir))
                    continue
                for task_dir in children(month_dir):
                    if _is_ignorable_archive_metadata(task_dir):
                        continue
                    task_id = _task_id_from_archive_directory_name(task_dir.name)
                    if task_id is None or not _is_valid_record_task_dir(task_dir, self.records_dir, task_id):
                        unexpected_entries.append(relative_name(task_dir))
                        continue
                    manifest_path = task_dir / MANIFEST_NAME
                    if _is_link_like(manifest_path) or not manifest_path.is_file():
                        missing_manifests.append(relative_name(task_dir))
                        continue
                    manifest_entries.append((manifest_path, False))

        for task_dir in children(self.trash_dir):
            if _is_ignorable_archive_metadata(task_dir):
                continue
            task_id = _task_id_from_archive_directory_name(task_dir.name, allow_copy_suffix=True)
            if task_id is None or not _is_valid_trash_task_dir(task_dir, self.trash_dir, task_id):
                unexpected_entries.append(relative_name(task_dir))
                continue
            manifest_path = task_dir / MANIFEST_NAME
            if _is_link_like(manifest_path) or not manifest_path.is_file():
                missing_manifests.append(relative_name(task_dir))
                continue
            manifest_entries.append((manifest_path, True))

        return (
            sorted(manifest_entries, key=lambda item: item[0].as_posix()),
            sorted(set(missing_manifests)),
            sorted(set(unexpected_entries)),
        )

    def rebuild_index_from_manifests(self) -> int:
        restored = 0
        validated = 0
        errors: list[dict[str, str]] = []
        manifest_entries, missing_manifests, unexpected_entries = self._discover_archive_manifests()
        seen_task_ids: set[str] = set()
        for manifest_path, deleted in manifest_entries:
            try:
                payload = json.loads(manifest_path.read_text(encoding="utf-8"))
                if not isinstance(payload, dict) or payload.get("schema_version") != SCHEMA_VERSION:
                    raise ValueError("历史清单版本无效")
                task = payload["task"]
                if not isinstance(task, dict):
                    raise ValueError("历史任务信息无效")
                task_id = str(task["id"])
                if len(task_id) != 32 or uuid.UUID(hex=task_id).hex != task_id.lower():
                    raise ValueError("历史任务编号无效")
                task_id = task_id.lower()
                if task_id in seen_task_ids:
                    raise ValueError("历史任务编号重复")
                seen_task_ids.add(task_id)
                task_dir = manifest_path.parent
                if deleted:
                    valid_task_dir = _is_valid_trash_task_dir(task_dir, self.trash_dir, task_id)
                else:
                    valid_task_dir = _is_valid_record_task_dir(task_dir, self.records_dir, task_id)
                if not valid_task_dir:
                    raise ValueError("历史任务目录层级无效")
                task_relpath = task_dir.absolute().relative_to(self.root.absolute()).as_posix()
                status = str(task["status"])
                if status not in ALLOWED_STATUSES:
                    raise ValueError("历史任务状态无效")
                with self._connect() as connection:
                    active_row = connection.execute(
                        "SELECT status, owner_pid FROM tasks WHERE id = ?",
                        (task_id,),
                    ).fetchone()
                if (
                    active_row is not None
                    and str(active_row["status"]) == "running"
                    and active_row["owner_pid"] is not None
                    and _pid_is_alive(int(active_row["owner_pid"]))
                ):
                    continue
                finished_at = task.get("finished_at")
                error_message = task.get("error_message")
                if status == "running":
                    status = "failed"
                    finished_at = _utc_now()
                    error_message = "程序上次意外关闭，本次处理未正常完成。"
                deleted_at = None
                if deleted:
                    deleted_at = str(task.get("deleted_at") or finished_at or task.get("started_at") or _utc_now())

                validated_files: list[dict[str, Any]] = []
                raw_files = payload.get("files", [])
                if not isinstance(raw_files, list):
                    raise ValueError("历史文件列表无效")
                for item in raw_files:
                    if not isinstance(item, dict):
                        raise ValueError("历史文件信息无效")
                    kind = str(item["kind"])
                    if kind not in {"input", "output"}:
                        raise ValueError("历史文件类型无效")
                    relative_path = Path(str(item["relative_path"]))
                    if relative_path.is_absolute() or ".." in relative_path.parts:
                        raise ValueError("历史文件路径无效")
                    expected_relative_root = Path(task_relpath) / ("inputs" if kind == "input" else "outputs")
                    try:
                        inside_relative = relative_path.relative_to(expected_relative_root)
                    except ValueError as exc:
                        raise ValueError("历史文件词法路径不属于当前任务") from exc
                    if not inside_relative.parts or not _path_components_are_real(self.root, relative_path):
                        raise ValueError("历史文件路径包含链接或无效层级")
                    archived_path = _safe_join(self.root, relative_path)
                    expected_root = task_dir / ("inputs" if kind == "input" else "outputs")
                    if not _is_within(archived_path, expected_root):
                        raise ValueError("历史文件不属于当前任务")
                    if _is_link_like(archived_path) or not archived_path.is_file():
                        raise ValueError("历史文件不存在或不安全")
                    size_bytes = int(item["size_bytes"])
                    if size_bytes < 0 or archived_path.stat().st_size != size_bytes:
                        raise ValueError("历史文件大小不一致")
                    digest = str(item["sha256"]).lower()
                    if len(digest) != 64 or any(character not in "0123456789abcdef" for character in digest):
                        raise ValueError("历史文件校验值无效")
                    if _sha256_file(archived_path) != digest:
                        raise ValueError("历史文件校验不一致")
                    validated_files.append(
                        {
                            "kind": kind,
                            "role": str(item.get("role") or "main")[:100],
                            "display_name": Path(str(item.get("display_name") or archived_path.name)).name,
                            "original_path": Path(str(item.get("original_path") or archived_path.name)).name,
                            "relative_path": relative_path.as_posix(),
                            "size_bytes": size_bytes,
                            "sha256": digest,
                            "modified_ns": max(0, int(item.get("modified_ns", 0))),
                        }
                    )

                changed = False
                with self._connect() as connection:
                    existing = connection.execute(
                        "SELECT task_relpath FROM tasks WHERE id = ?",
                        (task_id,),
                    ).fetchone()
                    if existing is None:
                        connection.execute(
                            """
                            INSERT INTO tasks (
                                id, tool_id, tool_name, mode, app_version, status,
                                started_at, finished_at, parameters_json, result_json,
                                error_message, owner_pid, task_relpath, deleted_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, ?, ?)
                            """,
                            (
                                task_id,
                                str(task["tool_id"]),
                                str(task["tool_name"]),
                                task.get("mode"),
                                str(task["app_version"]),
                                status,
                                str(task["started_at"]),
                                finished_at,
                                _json_dumps(task.get("parameters", {})),
                                _json_dumps(task.get("result", {})),
                                error_message,
                                task_relpath,
                                deleted_at,
                            ),
                        )
                        changed = True
                    elif str(existing["task_relpath"]) != task_relpath:
                        raise ValueError("历史任务编号与目录不一致")
                    for item in validated_files:
                        cursor = connection.execute(
                            """
                            INSERT OR IGNORE INTO files (
                                task_id, kind, role, display_name, original_path,
                                relative_path, size_bytes, sha256, modified_ns
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                task_id,
                                item["kind"],
                                item["role"],
                                item["display_name"],
                                item["original_path"],
                                item["relative_path"],
                                int(item["size_bytes"]),
                                item["sha256"],
                                int(item.get("modified_ns", 0)),
                            ),
                        )
                        changed = changed or cursor.rowcount > 0
                validated += 1
                if changed:
                    restored += 1
            except (OSError, KeyError, TypeError, ValueError, sqlite3.Error, json.JSONDecodeError) as exc:
                try:
                    manifest_name = manifest_path.relative_to(self.root).as_posix()
                except ValueError:
                    manifest_name = manifest_path.name
                errors.append(
                    {
                        "manifest": manifest_name,
                        "error": _clean_error(str(exc)) or exc.__class__.__name__,
                    }
                )
                continue
        self.last_rebuild_report = {
            "discovered": len(manifest_entries) + len(missing_manifests),
            "scanned": len(manifest_entries),
            "validated": validated,
            "restored": restored,
            "errors": errors,
            "missing_manifests": missing_manifests,
            "unexpected_entries": unexpected_entries,
        }
        return restored

    def integrity_check(self) -> bool:
        with self._connect() as connection:
            integrity = connection.execute("PRAGMA integrity_check").fetchone()[0]
            foreign_keys = connection.execute("PRAGMA foreign_key_check").fetchall()
        return integrity == "ok" and not foreign_keys

    def storage_stats(self) -> dict[str, int]:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    COALESCE(SUM(files.size_bytes), 0) AS total_bytes,
                    COALESCE(SUM(CASE WHEN tasks.deleted_at IS NOT NULL THEN files.size_bytes ELSE 0 END), 0)
                        AS trash_bytes
                FROM files
                JOIN tasks ON tasks.id = files.task_id
                """
            ).fetchone()
        return {
            "total_bytes": int(row["total_bytes"]),
            "trash_bytes": int(row["trash_bytes"]),
            "free_bytes": int(shutil.disk_usage(self.root).free),
        }

    def verify_task_files(self, task_id: str, *, kind: str | None = None) -> bool:
        if kind is not None and kind not in {"input", "output"}:
            raise ValueError("历史文件类型无效。")
        task = self.get_task(task_id)
        if task is None:
            raise HistoryStoreError("历史记录不存在。")
        records = (*task.inputs, *task.outputs)
        for record in records:
            if kind is not None and record.kind != kind:
                continue
            path = record.archived_path
            if _is_link_like(path) or not path.is_file():
                raise HistoryStoreError(f"历史原件已被移动或删除：{record.display_name}")
            if path.stat().st_size != record.size_bytes or _sha256_file(path) != record.sha256:
                raise HistoryStoreError(f"历史原件已发生变化，不能继续使用：{record.display_name}")
        return True

    def _iter_source_items(
        self,
        sources: Iterable[SourceSpec],
    ) -> Iterator[tuple[Path, str, tuple[str, ...]]]:
        for spec in sources:
            source = Path(spec.path).expanduser()
            if _is_link_like(source) or not source.exists():
                raise HistoryStoreError(f"原始文件不存在或是链接：{source.name or source}")
            source = source.resolve()
            if source.is_dir() and _is_within(self.root, source):
                raise HistoryStoreError("所选文件夹包含资料库，不能作为原始资料归档。")
            if _is_within(source, self.root):
                if not _is_reusable_archived_source(source, self.records_dir):
                    raise HistoryStoreError("不能把资料库自身作为原始资料归档。")
            suffixes = None if spec.suffixes is None else {suffix.lower() for suffix in spec.suffixes}
            if source.is_file():
                if suffixes is None or source.suffix.lower() in suffixes:
                    yield source, spec.role, (*_source_context_parts(source.parent, 3), source.name)
                continue
            if not source.is_dir():
                raise HistoryStoreError(f"无法读取原始资料：{source.name}")
            for child in _walk_regular_files(source):
                if suffixes is not None and child.suffix.lower() not in suffixes:
                    continue
                relative = child.relative_to(source)
                yield child, spec.role, (*_source_context_parts(source.parent, 2), source.name, *relative.parts)

    def _ensure_free_space(self, required_bytes: int) -> None:
        reserve = max(MIN_FREE_SPACE_BYTES, int(required_bytes * FREE_SPACE_RATIO))
        try:
            available = shutil.disk_usage(self.root).free
        except OSError as exc:
            raise HistoryStoreError(f"无法检查资料库剩余空间：{exc}") from exc
        if available < required_bytes + reserve:
            needed_gb = (required_bytes + reserve) / 1024 / 1024 / 1024
            available_gb = available / 1024 / 1024 / 1024
            raise HistoryStoreError(
                f"资料库空间不足，需要约 {needed_gb:.1f} GB，当前可用 {available_gb:.1f} GB。"
            )

    def _copy_and_hash(
        self,
        source: Path,
        destination: Path,
        task_id: str,
        kind: str,
        role: str,
        *,
        cancelled: Callable[[], bool] | None = None,
    ) -> dict[str, Any]:
        _mkdir_private(destination.parent)
        partial = destination.with_name(f".{destination.name}.{uuid.uuid4().hex}.partial")
        expected_archive = self._archived_source_expectation(source)
        before = source.stat()
        digest = hashlib.sha256()
        written = 0
        try:
            with source.open("rb") as input_handle, partial.open("xb") as output_handle:
                while True:
                    _raise_if_cancelled(cancelled)
                    chunk = input_handle.read(COPY_BUFFER_BYTES)
                    if not chunk:
                        break
                    output_handle.write(chunk)
                    digest.update(chunk)
                    written += len(chunk)
                output_handle.flush()
                os.fsync(output_handle.fileno())
            after = source.stat()
            if written != before.st_size or after.st_size != before.st_size or after.st_mtime_ns != before.st_mtime_ns:
                raise HistoryStoreError(f"复制期间文件发生变化，请关闭文件后重试：{source.name}")
            if expected_archive is not None:
                expected_size, expected_digest = expected_archive
                if written != expected_size or digest.hexdigest() != expected_digest:
                    raise HistoryStoreError(f"历史原件校验不一致，不能再次使用：{source.name}")
            os.utime(partial, ns=(after.st_atime_ns, after.st_mtime_ns))
            partial.replace(destination)
            _make_private(destination)
        except Exception:
            partial.unlink(missing_ok=True)
            raise
        return {
            "task_id": task_id,
            "kind": kind,
            "role": role,
            "display_name": source.name,
            "original_path": source.name,
            "archived_path": str(destination),
            "relative_path": destination.relative_to(self.root).as_posix(),
            "size_bytes": written,
            "sha256": digest.hexdigest(),
            "modified_ns": before.st_mtime_ns,
        }

    def _archived_source_expectation(self, source: Path) -> tuple[int, str] | None:
        if not _is_within(source, self.records_dir):
            return None
        relative_path = source.resolve().relative_to(self.root).as_posix()
        with self._connect() as connection:
            row = connection.execute(
                "SELECT size_bytes, sha256 FROM files WHERE relative_path = ? AND kind = 'input'",
                (relative_path,),
            ).fetchone()
        if row is None:
            raise HistoryStoreError(f"历史原件缺少校验记录，不能再次使用：{source.name}")
        return int(row["size_bytes"]), str(row["sha256"])

    def _insert_files(self, task_id: str, items: list[dict[str, Any]]) -> tuple[FileRecord, ...]:
        if not items:
            return ()
        inserted_ids: list[int] = []
        with self._connect() as connection:
            for item in items:
                cursor = connection.execute(
                    """
                    INSERT INTO files (
                        task_id, kind, role, display_name, original_path,
                        relative_path, size_bytes, sha256, modified_ns
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        task_id,
                        item["kind"],
                        item["role"],
                        item["display_name"],
                        item["original_path"],
                        item["relative_path"],
                        item["size_bytes"],
                        item["sha256"],
                        item["modified_ns"],
                    ),
                )
                inserted_ids.append(int(cursor.lastrowid))
            placeholders = ",".join("?" for _ in inserted_ids)
            rows = connection.execute(
                f"SELECT * FROM files WHERE id IN ({placeholders}) ORDER BY id",
                inserted_ids,
            ).fetchall()
        return tuple(self._file_from_row(row) for row in rows)

    def _file_from_row(self, row: sqlite3.Row) -> FileRecord:
        relative_path = str(row["relative_path"])
        return FileRecord(
            id=int(row["id"]),
            task_id=str(row["task_id"]),
            kind=str(row["kind"]),
            role=str(row["role"]),
            display_name=str(row["display_name"]),
            original_path=str(row["original_path"]),
            archived_path=_safe_join(self.root, Path(relative_path)),
            relative_path=relative_path,
            size_bytes=int(row["size_bytes"]),
            sha256=str(row["sha256"]),
            modified_ns=int(row["modified_ns"]),
        )

    def _summary_from_row(self, row: sqlite3.Row, names: dict[str, tuple[str, ...]]) -> TaskSummary:
        return TaskSummary(
            id=str(row["id"]),
            tool_id=str(row["tool_id"]),
            tool_name=str(row["tool_name"]),
            mode=None if row["mode"] is None else str(row["mode"]),
            app_version=str(row["app_version"]),
            status=str(row["status"]),
            started_at=str(row["started_at"]),
            finished_at=None if row["finished_at"] is None else str(row["finished_at"]),
            input_names=tuple(names.get("input", ())),
            output_names=tuple(names.get("output", ())),
            error_message=None if row["error_message"] is None else str(row["error_message"]),
            deleted_at=None if row["deleted_at"] is None else str(row["deleted_at"]),
        )

    def _file_names_for_tasks(
        self,
        connection: sqlite3.Connection,
        task_ids: list[str],
    ) -> dict[str, dict[str, tuple[str, ...]]]:
        if not task_ids:
            return {}
        placeholders = ",".join("?" for _ in task_ids)
        rows = connection.execute(
            f"SELECT task_id, kind, display_name FROM files WHERE task_id IN ({placeholders}) ORDER BY id",
            task_ids,
        ).fetchall()
        grouped: dict[str, dict[str, list[str]]] = {}
        for row in rows:
            task_id = str(row["task_id"])
            grouped.setdefault(task_id, {"input": [], "output": []})[str(row["kind"])].append(str(row["display_name"]))
        return {
            task_id: {kind: tuple(values) for kind, values in kinds.items()}
            for task_id, kinds in grouped.items()
        }

    def _write_manifest(
        self,
        task_id: str,
        *,
        pending_items: Iterable[dict[str, Any]] = (),
    ) -> None:
        with self._task_lock(task_id):
            task = self.get_task(task_id)
            if task is None:
                return
            manifest_lock = task.task_dir / ".manifest.lock"
            if _is_link_like(manifest_lock):
                raise HistoryStoreError("历史记录锁文件不安全。")
            with _exclusive_file_lock(manifest_lock):
                task = self.get_task(task_id)
                if task is None:
                    return
                self._write_manifest_locked(task, pending_items)

    def _write_manifest_locked(
        self,
        task: TaskDetail,
        pending_items: Iterable[dict[str, Any]],
    ) -> None:
        file_items = [
            {
                "kind": item.kind,
                "role": item.role,
                "display_name": item.display_name,
                "original_path": item.original_path,
                "relative_path": item.relative_path,
                "size_bytes": item.size_bytes,
                "sha256": item.sha256,
                "modified_ns": item.modified_ns,
            }
            for item in (*task.inputs, *task.outputs)
        ]
        known_paths = {str(item["relative_path"]) for item in file_items}
        for pending in pending_items:
            relative_path = str(pending["relative_path"])
            if relative_path in known_paths:
                continue
            file_items.append(
                {
                    "kind": str(pending["kind"]),
                    "role": str(pending["role"]),
                    "display_name": str(pending["display_name"]),
                    "original_path": str(pending["original_path"]),
                    "relative_path": relative_path,
                    "size_bytes": int(pending["size_bytes"]),
                    "sha256": str(pending["sha256"]),
                    "modified_ns": int(pending["modified_ns"]),
                }
            )
            known_paths.add(relative_path)
        payload = {
            "schema_version": SCHEMA_VERSION,
            "task": {
                "id": task.summary.id,
                "tool_id": task.summary.tool_id,
                "tool_name": task.summary.tool_name,
                "mode": task.summary.mode,
                "app_version": task.summary.app_version,
                "status": task.summary.status,
                "started_at": task.summary.started_at,
                "finished_at": task.summary.finished_at,
                "parameters": task.parameters,
                "result": task.result,
                "error_message": task.summary.error_message,
                "deleted_at": task.summary.deleted_at,
            },
            "files": file_items,
        }
        _atomic_write_text(task.task_dir / MANIFEST_NAME, _json_dumps(payload, pretty=True) + "\n")

    def _task_lock(self, task_id: str) -> threading.RLock:
        with self._locks_guard:
            return self._task_locks.setdefault(task_id, threading.RLock())

    def _task_io_lock(self, task_id: str) -> threading.RLock:
        with self._locks_guard:
            return self._task_io_locks.setdefault(task_id, threading.RLock())

    def _archive_lock_path(self, task_id: str, *, require_running: bool = True) -> Path:
        task = self.get_task(task_id)
        if task is None:
            raise HistoryStoreError("历史记录不存在。")
        if require_running and task.summary.status != "running":
            raise HistoryStoreError("本次历史记录已经结束，不能继续写入资料。")
        if not _is_valid_record_task_dir(task.task_dir, self.records_dir, task_id):
            raise HistoryStoreError("历史记录目录无效。")
        # Keep the archive lock outside the task directory.  Windows refuses
        # to rename a directory while a file inside that directory is still
        # open, which made a valid move-to-trash operation fail with
        # ``WinError 5`` even though the lock itself was correctly held.
        _mkdir_private(self.archive_locks_dir)
        lock_path = self.archive_locks_dir / f"{task_id.lower()}.archive.lock"
        if _is_link_like(lock_path):
            raise HistoryStoreError("历史记录锁文件不安全。")
        return lock_path


def _raise_if_cancelled(cancelled: Callable[[], bool] | None) -> None:
    if cancelled is not None and cancelled():
        raise HistoryStoreError("本次处理已停止。")


def _startup_lock_path(root: Path) -> Path:
    digest = hashlib.sha256(os.fsencode(str(root))).hexdigest()[:24]
    return root.parent / f".hrtoolkit-startup-{digest}.lock"


def _process_file_lock(path: Path) -> threading.RLock:
    key = os.path.normcase(os.path.abspath(os.fspath(path)))
    with _PROCESS_FILE_LOCKS_GUARD:
        lock = _PROCESS_FILE_LOCKS.get(key)
        if lock is None:
            lock = threading.RLock()
            _PROCESS_FILE_LOCKS[key] = lock
        return lock


@contextmanager
def _exclusive_file_lock(
    path: Path,
    *,
    blocking: bool = True,
    private_parent: bool = True,
) -> Iterator[None]:
    if private_parent:
        _mkdir_private(path.parent)
    else:
        parent_existed = path.parent.exists()
        path.parent.mkdir(parents=True, exist_ok=True)
        if _is_link_like(path.parent) or not path.parent.is_dir():
            raise HistoryStoreError("历史记录锁目录不安全。")
        if not parent_existed:
            _make_private(path.parent, directory=True)
    if _is_link_like(path):
        raise HistoryStoreError("历史记录锁文件不安全。")
    process_lock = _process_file_lock(path)
    if not process_lock.acquire(blocking=blocking):
        raise BlockingIOError(f"文件锁正在被当前进程占用：{path}")
    flags = os.O_RDWR | os.O_CREAT
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    try:
        file_descriptor = os.open(path, flags, 0o600)
    except BaseException:
        process_lock.release()
        raise
    locked = False
    try:
        if os.fstat(file_descriptor).st_size == 0:
            os.write(file_descriptor, b"0")
            os.fsync(file_descriptor)
        os.lseek(file_descriptor, 0, os.SEEK_SET)
        if os.name == "nt":
            import msvcrt

            lock_mode = msvcrt.LK_LOCK if blocking else msvcrt.LK_NBLCK
            try:
                msvcrt.locking(file_descriptor, lock_mode, 1)
            except OSError as exc:
                if not blocking:
                    raise BlockingIOError(str(exc)) from exc
                raise
        else:
            import fcntl

            lock_mode = fcntl.LOCK_EX | (0 if blocking else fcntl.LOCK_NB)
            fcntl.flock(file_descriptor, lock_mode)
        locked = True
        _make_private(path)
        yield
    finally:
        if locked:
            try:
                if os.name == "nt":
                    import msvcrt

                    os.lseek(file_descriptor, 0, os.SEEK_SET)
                    msvcrt.locking(file_descriptor, msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl

                    fcntl.flock(file_descriptor, fcntl.LOCK_UN)
            except OSError:
                pass
        os.close(file_descriptor)
        process_lock.release()


def _pid_is_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if os.name == "nt":
        import ctypes

        process_query_limited_information = 0x1000
        still_active = 259
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.argtypes = (ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong)
        kernel32.OpenProcess.restype = ctypes.c_void_p
        kernel32.GetExitCodeProcess.argtypes = (ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong))
        kernel32.GetExitCodeProcess.restype = ctypes.c_int
        kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
        kernel32.CloseHandle.restype = ctypes.c_int
        handle = kernel32.OpenProcess(process_query_limited_information, False, pid)
        if not handle:
            return ctypes.get_last_error() == 5  # ERROR_ACCESS_DENIED means the process exists.
        exit_code = ctypes.c_ulong()
        try:
            if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                return False
            return exit_code.value == still_active
        finally:
            kernel32.CloseHandle(handle)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


def _validate_data_root(path: Path) -> Path:
    if path.exists() and _is_link_like(path):
        raise HistoryStoreError("资料库位置不能是链接目录。")
    resolved = path.resolve()
    home = Path.home().resolve()
    if resolved == Path(resolved.anchor) or resolved == home:
        raise HistoryStoreError("资料库位置过于宽泛，请选择专用文件夹。")
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).resolve().parent
        if resolved == app_dir or _is_within(resolved, app_dir):
            raise HistoryStoreError("资料库不能放在程序安装目录内。")
    return resolved


def _mkdir_private(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    _make_private(path, directory=True)


def _make_private(path: Path, *, directory: bool = False) -> None:
    if os.name == "nt":
        return
    try:
        path.chmod(0o700 if directory or path.is_dir() else 0o600)
    except OSError:
        pass


def _walk_regular_files(root: Path) -> Iterator[Path]:
    for current, directories, files in os.walk(root, followlinks=False):
        current_path = Path(current)
        directories[:] = [name for name in sorted(directories) if not _is_link_like(current_path / name)]
        for name in sorted(files):
            path = current_path / name
            if _is_link_like(path) or not path.is_file():
                continue
            yield path.resolve()


def _is_link_like(path: Path) -> bool:
    if path.is_symlink():
        return True
    is_junction = getattr(path, "is_junction", None)
    if is_junction is not None:
        try:
            if bool(is_junction()):
                return True
        except OSError:
            return True
    if os.name == "nt":
        try:
            attributes = getattr(path.lstat(), "st_file_attributes", 0)
            return bool(attributes & 0x400)  # FILE_ATTRIBUTE_REPARSE_POINT
        except OSError:
            return False
    return False


def _assert_storage_entry(path: Path, label: str, *, kind: str, allow_missing: bool) -> None:
    try:
        entry_stat = path.lstat()
    except FileNotFoundError:
        if allow_missing:
            return
        raise HistoryStoreError(f"{label}不存在。")
    except OSError as exc:
        raise HistoryStoreError(f"无法核对{label}：{exc}") from exc
    is_reparse_point = bool(getattr(entry_stat, "st_file_attributes", 0) & 0x400)
    if stat_module.S_ISLNK(entry_stat.st_mode) or is_reparse_point:
        raise HistoryStoreError(f"{label}不能是链接或系统重定向目录。")
    if kind == "file" and not stat_module.S_ISREG(entry_stat.st_mode):
        raise HistoryStoreError(f"{label}不是普通文件。")
    if kind == "directory" and not stat_module.S_ISDIR(entry_stat.st_mode):
        raise HistoryStoreError(f"{label}不是普通目录。")


def _database_artifact_paths(database_path: Path) -> tuple[Path, ...]:
    return tuple(
        database_path.with_name(database_path.name + suffix)
        for suffix in ("", "-wal", "-shm", "-journal")
    )


def _normalize_schema_sql(value: Any) -> str:
    return "".join(str(value or "").split()).casefold()


def _database_schema_fingerprint(connection: sqlite3.Connection) -> tuple[tuple[str, str, str, str], ...]:
    rows = connection.execute(
        """
        SELECT type, name, tbl_name, sql
        FROM sqlite_master
        WHERE name NOT LIKE 'sqlite_%'
        ORDER BY type, name
        """
    ).fetchall()
    return tuple(
        (
            str(row[0]),
            str(row[1]),
            str(row[2]),
            _normalize_schema_sql(row[3]),
        )
        for row in rows
    )


@lru_cache(maxsize=1)
def _expected_database_schema_fingerprint() -> tuple[tuple[str, str, str, str], ...]:
    connection = sqlite3.connect(":memory:")
    try:
        connection.executescript(SCHEMA_SQL)
        return _database_schema_fingerprint(connection)
    finally:
        connection.close()


def _wal_checksum_bytes(
    data: bytes,
    byte_order: str,
    checksum_1: int = 0,
    checksum_2: int = 0,
) -> tuple[int, int]:
    if len(data) % 8:
        raise ValueError("WAL 校验数据长度无效")
    prefix = "<" if byte_order == "little" else ">"
    values = struct.unpack(f"{prefix}{len(data) // 4}I", data)
    for index in range(0, len(values), 2):
        checksum_1 = (checksum_1 + values[index] + checksum_2) & 0xFFFFFFFF
        checksum_2 = (checksum_2 + values[index + 1] + checksum_1) & 0xFFFFFFFF
    return checksum_1, checksum_2


def _validated_wal_frame_count(path: Path) -> int | None:
    if _is_link_like(path) or not path.is_file():
        return None
    try:
        with path.open("rb") as handle:
            header = handle.read(32)
            if len(header) != 32:
                return None
            magic = int.from_bytes(header[0:4], "big")
            checksum_order = {0x377F0682: "little", 0x377F0683: "big"}.get(magic)
            if checksum_order is None or int.from_bytes(header[4:8], "big") != 3_007_000:
                return None
            page_size = int.from_bytes(header[8:12], "big")
            if page_size < 512 or page_size > 65_536 or page_size & (page_size - 1):
                return None
            checksums = _wal_checksum_bytes(header[:24], checksum_order)
            stored_header_checksums = (
                int.from_bytes(header[24:28], "big"),
                int.from_bytes(header[28:32], "big"),
            )
            if checksums != stored_header_checksums:
                return None
            salt = header[16:24]
            frame_count = 0
            committed_frame_count = 0
            while True:
                frame_header = handle.read(24)
                if not frame_header:
                    return committed_frame_count
                if len(frame_header) < 24:
                    return None
                commit_page_count = int.from_bytes(frame_header[4:8], "big")
                page = handle.read(page_size)
                if len(page) < page_size:
                    return None if commit_page_count else committed_frame_count
                if int.from_bytes(frame_header[0:4], "big") == 0:
                    return None
                if frame_header[8:16] != salt:
                    # After WAL restart SQLite may reuse the beginning while
                    # leaving older-cycle frames in the physical tail.
                    return committed_frame_count if committed_frame_count else None
                checksums = _wal_checksum_bytes(
                    frame_header[:8] + page,
                    checksum_order,
                    checksums[0],
                    checksums[1],
                )
                stored_frame_checksums = (
                    int.from_bytes(frame_header[16:20], "big"),
                    int.from_bytes(frame_header[20:24], "big"),
                )
                if checksums != stored_frame_checksums:
                    return None if commit_page_count else committed_frame_count
                frame_count += 1
                if commit_page_count:
                    committed_frame_count = frame_count
    except (OSError, ValueError, struct.error):
        return None


def _is_probe_directory_name(name: str) -> bool:
    prefix = ".database-probe-"
    token = name[len(prefix) :].lower() if name.startswith(prefix) else ""
    return len(token) == 32 and all(character in "0123456789abcdef" for character in token)


def _is_rebuild_database_name(name: str) -> bool:
    prefix = ".history-rebuild-"
    suffix = ".db"
    if not name.startswith(prefix) or not name.endswith(suffix):
        return False
    token = name[len(prefix) : -len(suffix)].lower()
    return len(token) == 32 and all(character in "0123456789abcdef" for character in token)


def _is_database_corruption(exc: sqlite3.DatabaseError) -> bool:
    if isinstance(exc, _CorruptHistoryDatabase):
        return True
    error_code = getattr(exc, "sqlite_errorcode", None)
    if isinstance(error_code, int):
        primary_code = error_code & 0xFF
        if primary_code in {
            getattr(sqlite3, "SQLITE_CORRUPT", 11),
            getattr(sqlite3, "SQLITE_NOTADB", 26),
        }:
            return True
    message = str(exc).casefold()
    return any(
        marker in message
        for marker in (
            "database disk image is malformed",
            "file is not a database",
            "database corruption",
        )
    )


def _path_components_are_real(root: Path, relative: Path) -> bool:
    current = root.absolute()
    if _is_link_like(current):
        return False
    for part in relative.parts:
        current = current / part
        if _is_link_like(current):
            return False
    return True


def _is_owned_staging_name(name: str) -> bool:
    if not name.startswith("."):
        return False
    parts = name.rsplit(".", 2)
    if len(parts) != 3 or parts[2] not in {"tmp", "partial"}:
        return False
    token = parts[1].lower()
    return len(token) == 32 and all(character in "0123456789abcdef" for character in token)


def _source_context_parts(parent: Path, limit: int) -> tuple[str, ...]:
    generic_parts = {
        Path.home().name.casefold(),
        "users",
        "home",
        "downloads",
        "desktop",
        "documents",
        "tmp",
        "private",
        "var",
        "桌面",
        "下载",
        "文档",
    }
    candidates = [
        part
        for part in parent.parts
        if part and part != parent.anchor and part.casefold() not in generic_parts
    ]
    return tuple(candidates[-max(0, limit) :])


def _safe_component(value: str) -> str:
    normalized = unicodedata.normalize("NFC", str(value)).strip()
    invalid = '<>:"/\\|?*\0'
    for character in invalid:
        normalized = normalized.replace(character, "_")
    normalized = normalized.rstrip(" .")
    if normalized in {"", ".", ".."}:
        normalized = "_"
    stem = normalized.split(".", 1)[0].upper()
    if stem in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        normalized = "_" + normalized
    if len(normalized) > 160:
        suffix = Path(normalized).suffix
        normalized = normalized[: max(1, 160 - len(suffix))] + suffix
    return normalized


def _safe_join(root: Path, relative: Path) -> Path:
    if relative.is_absolute():
        raise HistoryStoreError("资料库路径无效。")
    # ``realpath`` on Windows can return a short/long-name variant when the
    # final directory does not exist yet.  Comparing that result with a root
    # resolved through a different Win32 call makes a valid concurrent task
    # path look outside the data root.  Normalize the lexical path first,
    # then inspect every existing component for symlinks/reparse points.
    root_path = Path(os.path.abspath(os.fspath(root)))
    target = Path(os.path.abspath(os.path.join(os.fspath(root_path), os.fspath(relative))))
    if not _is_lexically_within(target, root_path):
        raise HistoryStoreError("资料库路径越界。")
    try:
        target_relative = Path(os.path.relpath(os.fspath(target), os.fspath(root_path)))
    except ValueError as exc:
        raise HistoryStoreError("资料库路径越界。") from exc
    if not _path_components_are_real(root_path, target_relative):
        raise HistoryStoreError("资料库路径包含链接或系统重定向目录。")
    return target


def _is_lexically_within(path: Path, root: Path) -> bool:
    try:
        path_value = os.path.normcase(os.path.abspath(os.fspath(path)))
        root_value = os.path.normcase(os.path.abspath(os.fspath(root)))
        return os.path.commonpath((path_value, root_value)) == root_value
    except (OSError, ValueError):
        return False


def _is_within(path: Path, root: Path) -> bool:
    try:
        path_value = os.path.normcase(os.path.realpath(os.fspath(path)))
        root_value = os.path.normcase(os.path.realpath(os.fspath(root)))
        return os.path.commonpath((path_value, root_value)) == root_value
    except (OSError, ValueError):
        return False


def _paths_overlap(first: Path, second: Path) -> bool:
    return _is_within(first, second) or _is_within(second, first)


def _is_reusable_archived_source(path: Path, records_dir: Path) -> bool:
    try:
        relative = path.resolve().relative_to(records_dir.resolve())
    except ValueError:
        return False
    parts = relative.parts
    return len(parts) >= 5 and parts[3] == "inputs"


def _task_id_from_archive_directory_name(name: str, *, allow_copy_suffix: bool = False) -> str | None:
    base_name = name
    if allow_copy_suffix and name.endswith(")") and " (" in name:
        candidate, suffix = name.rsplit(" (", 1)
        number_text = suffix[:-1]
        if (
            not number_text.isdigit()
            or number_text.startswith("0")
            or not 2 <= int(number_text) <= 9_999
        ):
            return None
        base_name = candidate
    parts = base_name.split("_")
    if len(parts) != 3:
        return None
    day_text, time_text, task_id = parts
    if not (len(day_text) == 8 and day_text.isdigit() and len(time_text) == 6 and time_text.isdigit()):
        return None
    try:
        datetime.strptime(day_text + time_text, "%Y%m%d%H%M%S")
    except ValueError:
        return None
    task_id = task_id.lower()
    if len(task_id) != 32 or any(character not in "0123456789abcdef" for character in task_id):
        return None
    return task_id


def _is_ignorable_archive_metadata(path: Path) -> bool:
    return (
        not _is_link_like(path)
        and path.is_file()
        and path.name.casefold() in {".ds_store", "thumbs.db", "desktop.ini"}
    )


def _is_valid_record_task_dir(
    task_dir: Path,
    records_dir: Path,
    task_id: str,
    *,
    allow_missing: bool = False,
) -> bool:
    try:
        if _is_link_like(task_dir) or (not allow_missing and not task_dir.is_dir()):
            return False
        if allow_missing and task_dir.exists() and not task_dir.is_dir():
            return False
        relative = task_dir.absolute().relative_to(records_dir.absolute())
    except (OSError, ValueError):
        return False
    if not _path_components_are_real(records_dir, relative):
        return False
    parts = relative.parts
    if len(parts) != 3:
        return False
    year, month, directory_name = parts
    parsed_task_id = _task_id_from_archive_directory_name(directory_name)
    if parsed_task_id is None:
        return False
    day_text = directory_name.split("_", 1)[0]
    if not (year.isdigit() and len(year) == 4 and month.isdigit() and len(month) == 2):
        return False
    if day_text[:4] != year or day_text[4:6] != month or not 1 <= int(month) <= 12:
        return False
    task_id_lower = task_id.lower()
    if len(task_id_lower) != 32 or any(character not in "0123456789abcdef" for character in task_id_lower):
        return False
    return parsed_task_id == task_id_lower


def _is_valid_trash_task_dir(
    task_dir: Path,
    trash_dir: Path,
    task_id: str,
    *,
    allow_missing: bool = False,
) -> bool:
    try:
        if _is_link_like(task_dir) or (not allow_missing and not task_dir.is_dir()):
            return False
        if allow_missing and task_dir.exists() and not task_dir.is_dir():
            return False
        relative = task_dir.absolute().relative_to(trash_dir.absolute())
    except (OSError, ValueError):
        return False
    if len(relative.parts) != 1 or not _path_components_are_real(trash_dir, relative):
        return False
    parsed_task_id = _task_id_from_archive_directory_name(task_dir.name, allow_copy_suffix=True)
    return parsed_task_id == task_id.lower()


def _is_trash_move_marker_name(name: str) -> bool:
    suffix = ".json"
    if not name.startswith(TRASH_MOVE_PENDING_PREFIX) or not name.endswith(suffix):
        return False
    task_id = name[len(TRASH_MOVE_PENDING_PREFIX) : -len(suffix)].lower()
    return len(task_id) == 32 and all(character in "0123456789abcdef" for character in task_id)


def _verify_manifest_task_identity(manifest_path: Path, task_id: str) -> None:
    if _is_link_like(manifest_path) or not manifest_path.is_file():
        raise ValueError("历史清单不存在或不安全")
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    if (
        not isinstance(payload, dict)
        or payload.get("schema_version") != SCHEMA_VERSION
        or not isinstance(payload.get("task"), dict)
        or str(payload["task"].get("id", "")).lower() != task_id.lower()
    ):
        raise ValueError("历史清单与任务不一致")


def _pending_manifest_location(
    manifest_path: Path,
    task_id: str,
    old_relative: Path,
    new_relative: Path,
) -> str:
    _verify_manifest_task_identity(manifest_path, task_id)
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    raw_files = payload.get("files", [])
    if not isinstance(raw_files, list):
        raise ValueError("历史文件列表无效")
    locations: set[str] = set()
    for item in raw_files:
        if not isinstance(item, dict):
            raise ValueError("历史文件信息无效")
        kind = str(item.get("kind") or "")
        if kind not in {"input", "output"}:
            raise ValueError("历史文件类型无效")
        relative_path = Path(str(item.get("relative_path") or ""))
        if relative_path.is_absolute() or ".." in relative_path.parts:
            raise ValueError("历史文件路径无效")
        child_name = "inputs" if kind == "input" else "outputs"
        matched = False
        for location, task_relative in (("old", old_relative), ("new", new_relative)):
            try:
                inside = relative_path.relative_to(task_relative / child_name)
            except ValueError:
                continue
            if inside.parts:
                locations.add(location)
                matched = True
                break
        if not matched:
            raise ValueError("历史文件不属于待恢复任务")
    if len(locations) == 1:
        return next(iter(locations))
    if len(locations) > 1:
        raise ValueError("历史文件路径同时指向两个位置")
    task_payload = payload.get("task")
    if not isinstance(task_payload, dict):
        raise ValueError("历史任务信息无效")
    return "new" if task_payload.get("deleted_at") else "old"


def _is_safe_active_task_storage(task: TaskDetail, records_dir: Path) -> bool:
    if not _is_valid_record_task_dir(task.task_dir, records_dir, task.summary.id):
        return False
    for child_name in ("inputs", "outputs"):
        child = task.task_dir / child_name
        if _is_link_like(child) or not child.is_dir():
            return False
        if not _path_components_are_real(task.task_dir, Path(child_name)):
            return False
    return True


def _unique_destination(path: Path) -> Path:
    if not path.exists() and not path.with_name(path.name + ".partial").exists():
        return path
    suffixes = "".join(path.suffixes)
    base = path.name[: -len(suffixes)] if suffixes else path.name
    for index in range(2, 10_000):
        candidate = path.with_name(f"{base} ({index}){suffixes}")
        if not candidate.exists() and not candidate.with_name(candidate.name + ".partial").exists():
            return candidate
    raise HistoryStoreError(f"同名文件过多，无法归档：{path.name}")


def _atomic_write_text(path: Path, text: str) -> None:
    _mkdir_private(path.parent)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        with temporary.open("x", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        temporary.replace(path)
        _make_private(path)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def _copy_file_durable(source: Path, destination: Path) -> tuple[int, str]:
    _mkdir_private(destination.parent)
    temporary = destination.with_name(f".{destination.name}.{uuid.uuid4().hex}.partial")
    before = source.stat()
    digest = hashlib.sha256()
    written = 0
    try:
        with source.open("rb") as input_handle, temporary.open("xb") as output_handle:
            while True:
                chunk = input_handle.read(COPY_BUFFER_BYTES)
                if not chunk:
                    break
                output_handle.write(chunk)
                digest.update(chunk)
                written += len(chunk)
            output_handle.flush()
            os.fsync(output_handle.fileno())
        after = source.stat()
        if (
            written != before.st_size
            or after.st_size != before.st_size
            or after.st_mtime_ns != before.st_mtime_ns
        ):
            raise HistoryStoreError("备份期间历史索引发生变化。")
        temporary.replace(destination)
        _make_private(destination)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    return written, digest.hexdigest()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(COPY_BUFFER_BYTES), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _json_dumps(value: Any, *, pretty: bool = False) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        indent=2 if pretty else None,
        sort_keys=pretty,
        default=str,
    )


def _json_object(value: Any) -> dict[str, Any]:
    try:
        parsed = json.loads(str(value or "{}"))
    except (TypeError, ValueError, json.JSONDecodeError):
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _clean_error(value: str | None) -> str | None:
    if value is None:
        return None
    cleaned = " ".join(str(value).split())
    return cleaned[:2000]
