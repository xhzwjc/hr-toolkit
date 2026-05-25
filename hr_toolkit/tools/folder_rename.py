from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


TOOL_NAME = "需求8-人员资料文件夹改名"
MODE_APPEND = "append"
MODE_REMOVE = "remove"
MODE_REPLACE = "replace"
MODES = {MODE_APPEND, MODE_REMOVE, MODE_REPLACE}
WINDOWS_INVALID_CHARS = set('<>:"/\\|?*')


@dataclass(frozen=True)
class FolderRenameOperation:
    source: Path
    target: Path

    def to_dict(self) -> dict[str, str]:
        return {
            "source": str(self.source),
            "target": str(self.target),
            "source_name": self.source.name,
            "target_name": self.target.name,
        }


@dataclass
class FolderRenameResult:
    root_dir: Path
    mode: str
    dry_run: bool = False
    operations: list[FolderRenameOperation] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def operation_count(self) -> int:
        return len(self.operations)

    def to_dict(self) -> dict[str, Any]:
        return {
            "tool_name": TOOL_NAME,
            "root_dir": str(self.root_dir),
            "mode": self.mode,
            "dry_run": self.dry_run,
            "operation_count": self.operation_count,
            "operations": [operation.to_dict() for operation in self.operations],
            "warnings": self.warnings,
        }


def rename_person_folders(
    root_dir: str | Path,
    *,
    mode: str,
    text: str = "",
    target_name: str = "",
    replacement_name: str = "",
    dry_run: bool = False,
) -> FolderRenameResult:
    root_dir = Path(root_dir).expanduser().resolve()
    if mode not in MODES:
        raise ValueError(f"不支持的改名模式：{mode}")
    if not root_dir.exists() or not root_dir.is_dir():
        raise FileNotFoundError(f"文件夹不存在：{root_dir}")

    operations, warnings = _plan_operations(
        root_dir=root_dir,
        mode=mode,
        text=text,
        target_name=target_name,
        replacement_name=replacement_name,
    )
    result = FolderRenameResult(
        root_dir=root_dir,
        mode=mode,
        dry_run=dry_run,
        operations=operations,
        warnings=warnings,
    )
    if dry_run:
        return result

    completed: list[FolderRenameOperation] = []
    runtime_warnings = list(warnings)
    for operation in operations:
        try:
            operation.source.rename(operation.target)
        except OSError as exc:
            runtime_warnings.append(f"{operation.source.name} 改名失败：{exc}")
            continue
        completed.append(operation)
    result.operations = completed
    result.warnings = runtime_warnings
    return result


def _plan_operations(
    *,
    root_dir: Path,
    mode: str,
    text: str,
    target_name: str,
    replacement_name: str,
) -> tuple[list[FolderRenameOperation], list[str]]:
    text = text.strip()
    target_name = target_name.strip()
    replacement_name = replacement_name.strip()
    warnings: list[str] = []

    if mode == MODE_APPEND:
        if not text:
            raise ValueError("请填写要追加的内容，例如：劳动合同 或 -劳动合同")
        suffix = _normalize_append_text(text)
        candidates = _iter_target_dirs(root_dir, target_name)
        planned = [
            _build_operation(path, path.name + suffix)
            for path in candidates
            if not _already_has_suffix(path.name, suffix, warnings)
        ]
    elif mode == MODE_REMOVE:
        if not text:
            raise ValueError("请填写要删除的结尾文字，例如：劳动合同、-劳动合同 或 _身份证")
        candidates = _iter_target_dirs(root_dir, target_name)
        planned = []
        suffixes = _remove_suffix_candidates(text)
        for path in candidates:
            suffix = _matching_remove_suffix(path.name, suffixes)
            if suffix is None:
                continue
            new_name = path.name[: -len(suffix)]
            if not new_name:
                warnings.append(f"{path.name} 删除后名称为空，已跳过")
                continue
            planned.append(_build_operation(path, new_name))
    else:
        if not target_name:
            raise ValueError("请填写原文件夹名，例如：张三")
        if not replacement_name:
            raise ValueError("请填写替换后的文件夹名，例如：章五")
        source = root_dir / target_name
        if not source.exists() or not source.is_dir():
            raise FileNotFoundError(f"未找到要替换的文件夹：{target_name}")
        planned = [_build_operation(source, replacement_name)]

    operations = _filter_invalid_operations(planned, warnings)
    return operations, warnings


def _iter_target_dirs(root_dir: Path, target_name: str) -> list[Path]:
    if target_name:
        target = root_dir / target_name
        if target.exists() and target.is_dir():
            return [target]
        return [
            path
            for path in sorted(root_dir.iterdir())
            if path.is_dir() and path.name.startswith(target_name)
        ]
    return sorted(path for path in root_dir.iterdir() if path.is_dir())


def _normalize_append_text(text: str) -> str:
    if text.startswith(("-", "_")):
        return text
    return "-" + text


def _remove_suffix_candidates(text: str) -> list[str]:
    if text.startswith(("-", "_")):
        base = text[1:]
        candidates = [text, "-" + base, "_" + base, base]
    else:
        candidates = [text, "-" + text, "_" + text]
    return sorted(set(candidates), key=len, reverse=True)


def _matching_remove_suffix(folder_name: str, suffixes: list[str]) -> str | None:
    for suffix in suffixes:
        if suffix and folder_name.endswith(suffix):
            return suffix
    return None


def _already_has_suffix(folder_name: str, suffix: str, warnings: list[str]) -> bool:
    if folder_name.endswith(suffix):
        warnings.append(f"{folder_name} 已包含后缀，已跳过")
        return True
    return False


def _build_operation(source: Path, target_name: str) -> FolderRenameOperation:
    _validate_folder_name(target_name)
    return FolderRenameOperation(source=source, target=source.with_name(target_name))


def _validate_folder_name(name: str) -> None:
    if not name.strip():
        raise ValueError("文件夹名称不能为空")
    if name in {".", ".."}:
        raise ValueError(f"文件夹名称不合法：{name}")
    if any(char in WINDOWS_INVALID_CHARS for char in name):
        raise ValueError(f"文件夹名称包含 Windows 不支持的字符：{name}")


def _filter_invalid_operations(
    operations: list[FolderRenameOperation],
    warnings: list[str],
) -> list[FolderRenameOperation]:
    valid: list[FolderRenameOperation] = []
    seen_targets: set[Path] = set()
    for operation in operations:
        if operation.source == operation.target:
            warnings.append(f"{operation.source.name} 改名前后相同，已跳过")
            continue
        if operation.target in seen_targets:
            warnings.append(f"{operation.target.name} 目标名称重复，已跳过")
            continue
        if operation.target.exists():
            warnings.append(f"{operation.target.name} 已存在，{operation.source.name} 已跳过")
            continue
        seen_targets.add(operation.target)
        valid.append(operation)
    return valid
