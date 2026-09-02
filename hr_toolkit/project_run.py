"""Framework-neutral helpers for running tools inside a project workspace.

The desktop front ends must not implement their own variations of project
snapshotting or argument rebasing.  These helpers are deliberately independent
of Tk and Qt so both front ends preserve the exact same source-retention and
business-call semantics.
"""

from __future__ import annotations

import inspect
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from .history_store import SourceSpec


PRIMARY_PATH_ARGUMENTS = frozenset(
    {"input_path", "input_dir", "summary_input", "summary_path"}
)
SUPPORTING_PATH_ARGUMENTS = frozenset(
    {
        "roster_path",
        "roster_source",
        "report_staff_path",
        "existing_summary_path",
        "template_path",
        "analysis_template_path",
        "excel_path",
        "target_path",
        "existing_archive_path",
    }
)
PATH_ARGUMENTS = PRIMARY_PATH_ARGUMENTS | SUPPORTING_PATH_ARGUMENTS


def serializable(value: Any) -> Any:
    """Return the stable JSON-compatible representation used by run metadata."""

    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, (list, tuple)):
        return [serializable(item) for item in value]
    if isinstance(value, dict):
        return {str(key): serializable(item) for key, item in value.items()}
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def context_from_call(
    tool_func: Callable[..., Any],
    args: Iterable[Any],
    kwargs: dict[str, Any],
) -> tuple[list[SourceSpec], dict[str, object], Path | None]:
    """Describe source files and parameters for one business-tool invocation."""

    bound = inspect.signature(tool_func).bind_partial(*args, **kwargs)
    bound.apply_defaults()
    parameters: dict[str, object] = {}
    sources: list[SourceSpec] = []
    output_dir: Path | None = None
    for name, value in bound.arguments.items():
        if name in {"progress_callback", "cancelled"}:
            continue
        if name == "output_dir" and value is not None:
            output_dir = Path(value).expanduser()
            parameters[name] = output_dir.name
            continue
        if name == "root_dir" and value is not None:
            root_dir = Path(value).expanduser()
            parameters[name] = root_dir.name
            if root_dir.exists():
                sources.append(
                    SourceSpec(
                        path=root_dir,
                        role="input_path",
                        suffixes=None,
                        preserve_directories=True,
                    )
                )
                output_dir = root_dir
            continue
        if name == "library_dir" and value is not None:
            parameters[name] = str(value)
            continue
        if name == "roster_source" and not _is_file_roster_source(value):
            parameters[name] = serializable(value)
            continue
        if name not in PATH_ARGUMENTS or value is None:
            parameters[name] = serializable(value)
            continue
        raw_paths = value if isinstance(value, (list, tuple)) else [value]
        parameters[name] = [Path(item).name for item in raw_paths if item is not None]
        if not isinstance(value, (list, tuple)):
            parameters[name] = parameters[name][0] if parameters[name] else None
        role = "input_path" if name in PRIMARY_PATH_ARGUMENTS else name
        for raw_path in raw_paths:
            if raw_path is None:
                continue
            path = Path(raw_path).expanduser()
            if path.exists():
                sources.append(SourceSpec(path=path, role=role))
    return sources, parameters, output_dir


def _is_file_roster_source(value: Any) -> bool:
    return isinstance(value, Path) or (
        isinstance(value, str) and Path(value).expanduser().is_file()
    )


def project_source_replacement(
    store: Any,
    batch_id: str,
    source: SourceSpec,
    records: Iterable[Any],
    *,
    source_was_file: bool,
) -> Path:
    """Resolve one imported source to the shape expected by the business tool."""

    materialized = list(records)
    if not materialized:
        raise RuntimeError(f"没有可留存的资料：{source.path.name}")
    paths = [record.path(store.workspace) for record in materialized]
    if source_was_file:
        if len(paths) != 1:
            raise RuntimeError(f"资料快照不完整：{source.path.name}")
        return paths[0]
    detail = store.get_batch(batch_id)
    if detail is None:
        raise RuntimeError("本次处理批次无法读取。")
    upload_root = detail.directories["uploads"]
    top_names: set[str] = set()
    for path in paths:
        relative = path.relative_to(upload_root)
        if not relative.parts:
            raise RuntimeError(f"资料快照位置无效：{source.path.name}")
        top_names.add(relative.parts[0])
    if len(top_names) != 1:
        raise RuntimeError(f"文件夹快照不完整：{source.path.name}")
    return upload_root / next(iter(top_names))


def import_project_run_sources(
    store: Any,
    batch_id: str,
    sources: Iterable[SourceSpec],
    cancel_event: Any,
    *,
    on_progress: Callable[[Any], None] | None = None,
) -> dict[str, list[Path]]:
    """Copy a run's sources into the project without resolving link targets."""

    replacements: dict[str, list[Path]] = {}
    project_root = Path(store.root).absolute()
    for source in sources:
        source_path = Path(source.path).expanduser().absolute()
        source_was_file = source_path.is_file()
        try:
            source_path.relative_to(project_root)
            is_project_source = True
        except ValueError:
            is_project_source = False
        if source.preserve_directories:
            if source_was_file:
                raise RuntimeError(f"需要文件夹资料：{source.path.name}")
            method_name = (
                "copy_project_directory_snapshot"
                if is_project_source
                else "import_directory_snapshot"
            )
            snapshotter = getattr(store, method_name, None)
            if not callable(snapshotter):
                raise RuntimeError("当前版本无法安全留存文件夹结构。")
            replacement = snapshotter(
                batch_id,
                source_path,
                category="uploads",
                role=source.role,
                cancelled=cancel_event.is_set,
                on_progress=on_progress,
            )
        else:
            if is_project_source:
                copier = getattr(store, "copy_project_sources", None)
                if not callable(copier):
                    raise RuntimeError(
                        "当前版本暂时不能复用项目内资料，请从原文件位置重新选择。"
                    )
            else:
                copier = store.import_sources
            records = copier(
                batch_id,
                [source_path],
                category="uploads",
                role=source.role,
                cancelled=cancel_event.is_set,
                on_progress=on_progress,
            )
            replacement = project_source_replacement(
                store,
                batch_id,
                source,
                records,
                source_was_file=source_was_file,
            )
        replacements.setdefault(source.role, []).append(replacement)
    return replacements


def rebase_project_replacements(
    replacements: dict[str, list[Path]],
    old_upload_root: Path,
    new_upload_root: Path,
) -> dict[str, list[Path]]:
    return {
        role: [new_upload_root / path.relative_to(old_upload_root) for path in paths]
        for role, paths in replacements.items()
    }


def call_with_project_inputs(
    tool_func: Callable[..., Any],
    args: Iterable[Any],
    kwargs: dict[str, Any],
    replacements: dict[str, list[Path]],
    result_dir: Path,
    store: Any,
    batch_id: str,
) -> tuple[tuple[Any, ...], dict[str, Any]]:
    """Rebind a call from external inputs to immutable project snapshots."""

    bound = inspect.signature(tool_func).bind_partial(*args, **kwargs)
    bound.apply_defaults()
    for name, value in tuple(bound.arguments.items()):
        if name == "output_dir":
            bound.arguments[name] = result_dir
            continue
        if name == "library_dir":
            continue
        if name == "roster_source" and not _is_file_roster_source(value):
            continue
        if value is None or name not in PATH_ARGUMENTS | {"root_dir"}:
            continue
        role = "input_path" if name in PRIMARY_PATH_ARGUMENTS or name == "root_dir" else name
        copied_paths = replacements.get(role, [])
        if not copied_paths:
            raise RuntimeError(f"没有完整保存 {name} 对应的原始资料。")
        original_values = value if isinstance(value, (list, tuple)) else [value]
        original_paths = [Path(item).expanduser() for item in original_values if item is not None]
        original_was_directory = len(original_paths) == 1 and original_paths[0].is_dir()
        if name == "root_dir":
            if len(copied_paths) != 1 or not copied_paths[0].is_dir():
                raise RuntimeError("人员资料文件夹快照不完整。")
            copier = getattr(store, "create_result_working_copy", None)
            if not callable(copier):
                raise RuntimeError("当前版本无法安全建立文件夹处理副本。")
            replacement: Any = copier(batch_id, copied_paths[0])
        elif name == "template_path" and original_was_directory:
            replacement = copied_paths[0]
        elif isinstance(value, (list, tuple)) or original_was_directory:
            replacement = copied_paths
        else:
            replacement = copied_paths[0]
        bound.arguments[name] = replacement
    return bound.args, bound.kwargs


def project_batch_is_closed(store: Any, batch_id: str) -> bool:
    detail = store.get_batch(batch_id)
    if detail is not None:
        return detail.summary.status in {"success", "failed", "stopped"}
    return any(summary.id == batch_id for summary in store.list_trash())
