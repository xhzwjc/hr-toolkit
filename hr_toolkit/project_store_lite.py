"""Lightweight Project Workspace Manager (ProjectStoreLite).

A streamlined, self-contained reference implementation for departmental automation
toolkits. Provides safe sandbox file copy, SHA-256 hashing, batch manifests, and
result isolation in ~350 lines of clean Python.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Sequence


PROJECT_METADATA_DIR = ".toolkit"
PROJECT_CONFIG_FILE = "project.json"
MANIFESTS_DIR = "manifests"
CATEGORY_UPLOADS = "uploads"
CATEGORY_RESULTS = "results"
CATEGORY_COMMON = "common"
COPY_BUFFER_SIZE = 1024 * 1024


class ProjectStoreLiteError(RuntimeError):
    """Raised when a workspace operation fails."""


@dataclass(frozen=True)
class StoredFile:
    file_id: str
    display_name: str
    relative_path: str
    size_bytes: int
    sha256: str
    category: str


@dataclass(frozen=True)
class BatchRecord:
    batch_id: str
    tool_id: str
    tool_name: str
    created_at: str
    status: str
    input_files: tuple[StoredFile, ...] = ()
    output_files: tuple[StoredFile, ...] = ()
    parameters: dict[str, Any] = field(default_factory=dict)
    summary: dict[str, Any] = field(default_factory=dict)
    error_message: str | None = None


def _calculate_sha256(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as f:
        while chunk := f.read(COPY_BUFFER_SIZE):
            hasher.update(chunk)
    return hasher.hexdigest()


class ProjectStoreLite:
    """Lightweight project store managing workspace directory, inputs, and results."""

    def __init__(self, root: Path, project_id: str, name: str, created_at: str) -> None:
        self.root = root.resolve()
        self.project_id = project_id
        self.name = name
        self.created_at = created_at
        self.metadata_dir = self.root / PROJECT_METADATA_DIR
        self.manifests_dir = self.metadata_dir / MANIFESTS_DIR
        self.uploads_dir = self.root / CATEGORY_UPLOADS
        self.results_dir = self.root / CATEGORY_RESULTS
        self.common_dir = self.root / CATEGORY_COMMON

    @classmethod
    def create(cls, root: str | Path, name: str) -> "ProjectStoreLite":
        project_root = Path(root).expanduser().resolve()
        project_name = name.strip()
        if not project_name:
            raise ProjectStoreLiteError("项目名称不能为空")

        if project_root.exists():
            if not project_root.is_dir():
                raise ProjectStoreLiteError(f"目标位置不是有效目录：{project_root}")
            if any(project_root.iterdir()):
                raise ProjectStoreLiteError(f"新建项目必须使用空目录：{project_root}")
        else:
            project_root.mkdir(parents=True, exist_ok=True)

        metadata_dir = project_root / PROJECT_METADATA_DIR
        metadata_dir.mkdir(parents=True, exist_ok=True)
        (metadata_dir / MANIFESTS_DIR).mkdir(parents=True, exist_ok=True)
        (project_root / CATEGORY_UPLOADS).mkdir(parents=True, exist_ok=True)
        (project_root / CATEGORY_RESULTS).mkdir(parents=True, exist_ok=True)
        (project_root / CATEGORY_COMMON).mkdir(parents=True, exist_ok=True)

        project_id = uuid.uuid4().hex
        created_at = datetime.now(timezone.utc).isoformat()
        config = {
            "format_version": 1,
            "project_id": project_id,
            "name": project_name,
            "created_at": created_at,
        }
        (metadata_dir / PROJECT_CONFIG_FILE).write_text(
            json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return cls(project_root, project_id, project_name, created_at)

    @classmethod
    def open(cls, root: str | Path) -> "ProjectStoreLite":
        project_root = Path(root).expanduser().resolve()
        if not project_root.is_dir():
            raise ProjectStoreLiteError(f"项目目录不存在：{project_root}")
        config_path = project_root / PROJECT_METADATA_DIR / PROJECT_CONFIG_FILE
        if not config_path.is_file():
            raise ProjectStoreLiteError(f"未找到项目标记文件，非有效项目：{project_root}")
        try:
            config = json.loads(config_path.read_text(encoding="utf-8"))
            return cls(
                root=project_root,
                project_id=config["project_id"],
                name=config["name"],
                created_at=config["created_at"],
            )
        except Exception as exc:
            raise ProjectStoreLiteError(f"无法读取项目配置：{exc}") from exc

    def import_file(self, source_path: str | Path, category: str = CATEGORY_UPLOADS) -> StoredFile:
        source = Path(source_path).expanduser().resolve()
        if not source.is_file():
            raise ProjectStoreLiteError(f"源文件不存在：{source}")

        target_dir = self.root / category
        target_dir.mkdir(parents=True, exist_ok=True)

        target_name = source.name
        counter = 1
        target_file = target_dir / target_name
        while target_file.exists():
            stem = source.stem
            suffix = source.suffix
            target_name = f"{stem}_{counter}{suffix}"
            target_file = target_dir / target_name
            counter += 1

        shutil.copyfile(source, target_file)
        file_sha256 = _calculate_sha256(target_file)
        size_bytes = target_file.stat().st_size
        relative_path = target_file.relative_to(self.root).as_posix()

        return StoredFile(
            file_id=uuid.uuid4().hex,
            display_name=target_name,
            relative_path=relative_path,
            size_bytes=size_bytes,
            sha256=file_sha256,
            category=category,
        )

    def prepare_result_path(self, filename: str) -> Path:
        self.results_dir.mkdir(parents=True, exist_ok=True)
        return self.results_dir / filename

    def record_batch(
        self,
        tool_id: str,
        tool_name: str,
        input_files: Sequence[StoredFile],
        output_files: Sequence[StoredFile],
        status: str = "success",
        parameters: dict[str, Any] | None = None,
        summary: dict[str, Any] | None = None,
        error_message: str | None = None,
    ) -> BatchRecord:
        batch_id = uuid.uuid4().hex
        created_at = datetime.now(timezone.utc).isoformat()
        record = BatchRecord(
            batch_id=batch_id,
            tool_id=tool_id,
            tool_name=tool_name,
            created_at=created_at,
            status=status,
            input_files=tuple(input_files),
            output_files=tuple(output_files),
            parameters=parameters or {},
            summary=summary or {},
            error_message=error_message,
        )

        manifest_data = {
            "batch_id": batch_id,
            "tool_id": tool_id,
            "tool_name": tool_name,
            "created_at": created_at,
            "status": status,
            "parameters": record.parameters,
            "summary": record.summary,
            "error_message": record.error_message,
            "inputs": [
                {
                    "file_id": f.file_id,
                    "display_name": f.display_name,
                    "relative_path": f.relative_path,
                    "size_bytes": f.size_bytes,
                    "sha256": f.sha256,
                    "category": f.category,
                }
                for f in record.input_files
            ],
            "outputs": [
                {
                    "file_id": f.file_id,
                    "display_name": f.display_name,
                    "relative_path": f.relative_path,
                    "size_bytes": f.size_bytes,
                    "sha256": f.sha256,
                    "category": f.category,
                }
                for f in record.output_files
            ],
        }

        manifest_file = self.manifests_dir / f"{batch_id}.json"
        manifest_file.write_text(
            json.dumps(manifest_data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return record

    def list_batches(self) -> list[dict[str, Any]]:
        batches = []
        if not self.manifests_dir.exists():
            return []
        for file in sorted(self.manifests_dir.glob("*.json")):
            try:
                data = json.loads(file.read_text(encoding="utf-8"))
                batches.append(data)
            except Exception:
                continue
        return sorted(batches, key=lambda x: x.get("created_at", ""), reverse=True)
