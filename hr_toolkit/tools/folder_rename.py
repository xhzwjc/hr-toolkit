from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from hr_toolkit.common.excel_compat import ensure_xlsx_workbook, is_supported_excel_file


TOOL_NAME = "需求8-人员资料文件夹改名"
MODE_APPEND = "append"
MODE_REMOVE = "remove"
MODE_REPLACE = "replace"
MODES = {MODE_APPEND, MODE_REMOVE, MODE_REPLACE}
WINDOWS_INVALID_CHARS = set('<>:"/\\|?*')

# 文件类型分组
FILE_TYPE_FOLDER = "folder"
FILE_TYPE_PDF = "pdf"
FILE_TYPE_IMAGE = "image"
FILE_TYPE_DOCUMENT = "document"
FILE_TYPE_ALL = "all"
FILE_TYPE_EXTENSIONS: dict[str, list[str]] = {
    FILE_TYPE_FOLDER: [],
    FILE_TYPE_PDF: [".pdf"],
    FILE_TYPE_IMAGE: [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"],
    FILE_TYPE_DOCUMENT: [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt"],
    FILE_TYPE_ALL: [],
}


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
    file_type: str = FILE_TYPE_FOLDER,
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
        file_type=file_type,
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


def rename_files_by_excel(
    root_dir: str | Path,
    excel_path: str | Path,
    *,
    name_column: str = "姓名",
    header_row: int = 1,
    file_extensions: list[str] | None = None,
    dry_run: bool = False,
) -> FolderRenameResult:
    """按照Excel名字列批量改名文件（支持PDF、图片等任意格式）

    Args:
        root_dir: 包含待改名文件的文件夹路径
        excel_path: Excel文件路径，包含名字列
        name_column: 名字列的表头名称，默认"姓名"
        header_row: 表头行号，默认1
        file_extensions: 要改名的文件扩展名列表，如 [".pdf", ".jpg"]，None表示所有文件
        dry_run: 是否只预览不执行
    """
    root_dir = Path(root_dir).expanduser().resolve()
    excel_path = Path(excel_path).expanduser().resolve()
    if not root_dir.exists() or not root_dir.is_dir():
        raise FileNotFoundError(f"文件夹不存在：{root_dir}")
    if not excel_path.exists() or not excel_path.is_file():
        raise FileNotFoundError(f"Excel文件不存在：{excel_path}")
    if not is_supported_excel_file(excel_path):
        raise ValueError("仅支持 .xlsx 或 .xls 文件。")

    # 读取Excel名字列
    names = _read_names_from_excel(excel_path, name_column, header_row)
    if not names:
        raise ValueError(f"Excel文件中未找到“{name_column}”列或该列无数据。")

    # 获取待改名的文件列表
    files = _list_files_for_rename(root_dir, file_extensions)
    if not files:
        raise ValueError(f"文件夹中未找到可改名的文件。")

    # 生成改名操作
    operations, warnings = _plan_excel_batch_operations(root_dir, files, names)

    result = FolderRenameResult(
        root_dir=root_dir,
        mode=MODE_EXCEL_BATCH,
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


def _read_names_from_excel(excel_path: Path, name_column: str, header_row: int) -> list[str]:
    """从Excel文件中读取名字列"""
    import tempfile
    with tempfile.TemporaryDirectory() as temp_dir:
        working_path = ensure_xlsx_workbook(excel_path, Path(temp_dir))
        wb = load_workbook(working_path, data_only=True, read_only=True)
        try:
            ws = wb.active
            # 查找名字列
            name_col = None
            max_col = min(ws.max_column or 0, 50)
            for col in range(1, max_col + 1):
                header = _normalize_text(ws.cell(header_row, col).value)
                if header == name_column:
                    name_col = col
                    break
            if name_col is None:
                # 尝试模糊匹配
                for col in range(1, max_col + 1):
                    header = _normalize_text(ws.cell(header_row, col).value)
                    if name_column in header or header in ("姓名", "名字", "名称"):
                        name_col = col
                        break
            if name_col is None:
                return []
            # 读取名字列
            names: list[str] = []
            for row in range(header_row + 1, (ws.max_row or 0) + 1):
                value = ws.cell(row, name_col).value
                if value is not None:
                    name = str(value).strip()
                    if name:
                        names.append(name)
            return names
        finally:
            wb.close()


def _list_files_for_rename(root_dir: Path, file_extensions: list[str] | None) -> list[Path]:
    """列出待改名的文件"""
    files: list[Path] = []
    for path in sorted(root_dir.iterdir()):
        if not path.is_file():
            continue
        # 跳过隐藏文件和临时文件
        if path.name.startswith((".", "~$")):
            continue
        if file_extensions is not None:
            if path.suffix.lower() not in file_extensions:
                continue
        files.append(path)
    return files


def _plan_excel_batch_operations(
    root_dir: Path,
    files: list[Path],
    names: list[str],
) -> tuple[list[FolderRenameOperation], list[str]]:
    """生成Excel批量改名操作"""
    operations: list[FolderRenameOperation] = []
    warnings: list[str] = []

    # 文件数量和名字数量不匹配时警告
    if len(files) != len(names):
        warnings.append(f"文件数量（{len(files)}）与名字数量（{len(names)}）不匹配，将按顺序改名到数量较少的一方。")

    count = min(len(files), len(names))
    for i in range(count):
        source = files[i]
        new_name = names[i]
        # 保留原文件扩展名
        suffix = source.suffix
        target_name = f"{new_name}{suffix}"
        target = source.parent / target_name

        if source == target:
            warnings.append(f"{source.name} 改名前后相同，已跳过")
            continue
        if target.exists():
            warnings.append(f"{target.name} 已存在，{source.name} 已跳过")
            continue

        operations.append(FolderRenameOperation(source=source, target=target))

    return operations, warnings


def _normalize_text(value: Any) -> str:
    """规范化文本：移除空白字符"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).replace("\xa0", "")


def _plan_operations(
    *,
    root_dir: Path,
    mode: str,
    text: str,
    target_name: str,
    replacement_name: str,
    file_type: str = FILE_TYPE_FOLDER,
) -> tuple[list[FolderRenameOperation], list[str]]:
    text = text.strip()
    target_name = target_name.strip()
    replacement_name = replacement_name.strip()
    warnings: list[str] = []

    if mode == MODE_APPEND:
        if not text:
            raise ValueError("请填写要追加的内容，例如：劳动合同 或 -劳动合同")
        suffix = _normalize_append_text(text)
        candidates = _iter_target_items(root_dir, target_name, file_type)
        planned = []
        for path in candidates:
            # 获取不含扩展名的名称（文件夹没有扩展名）
            stem = path.stem if path.is_file() else path.name
            ext = path.suffix if path.is_file() else ""
            if _already_has_suffix(stem, suffix, warnings):
                continue
            new_name = f"{stem}{suffix}{ext}"
            planned.append(_build_operation(path, new_name))
    elif mode == MODE_REMOVE:
        if not text:
            raise ValueError("请填写要删除的结尾文字，例如：劳动合同、-劳动合同 或 _身份证")
        candidates = _iter_target_items(root_dir, target_name, file_type)
        planned = []
        suffixes = _remove_suffix_candidates(text)
        for path in candidates:
            # 获取不含扩展名的名称
            stem = path.stem if path.is_file() else path.name
            ext = path.suffix if path.is_file() else ""
            suffix = _matching_remove_suffix(stem, suffixes)
            if suffix is None:
                continue
            new_stem = stem[: -len(suffix)]
            if not new_stem:
                warnings.append(f"{path.name} 删除后名称为空，已跳过")
                continue
            new_name = f"{new_stem}{ext}"
            planned.append(_build_operation(path, new_name))
    else:
        if not target_name:
            raise ValueError("请填写原名称，例如：张三")
        if not replacement_name:
            raise ValueError("请填写替换后的名称，例如：章五")
        source = root_dir / target_name
        if not source.exists():
            raise FileNotFoundError(f"未找到要替换的项目：{target_name}")
        # 对于文件，保留原扩展名
        if source.is_file():
            ext = source.suffix
            if not replacement_name.endswith(ext):
                replacement_name = f"{replacement_name}{ext}"
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


def _iter_target_items(root_dir: Path, target_name: str, file_type: str) -> list[Path]:
    """根据文件类型返回待改名的项目列表"""
    extensions = FILE_TYPE_EXTENSIONS.get(file_type, [])

    def _matches_type(path: Path) -> bool:
        if file_type == FILE_TYPE_FOLDER:
            return path.is_dir()
        if file_type == FILE_TYPE_ALL:
            return path.is_dir() or path.is_file()
        # 按扩展名匹配文件
        return path.is_file() and path.suffix.lower() in extensions

    if target_name:
        target = root_dir / target_name
        if target.exists() and _matches_type(target):
            return [target]
        return [
            path
            for path in sorted(root_dir.iterdir())
            if _matches_type(path) and target_name in path.name
        ]
    return sorted(path for path in root_dir.iterdir() if _matches_type(path) and not path.name.startswith((".", "~$")))


def _normalize_append_text(text: str) -> str:
    return text


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
