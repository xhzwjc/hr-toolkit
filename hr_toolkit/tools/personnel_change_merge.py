from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from hr_toolkit.common.excel import clone_style


TOOL_NAME = "需求6-异动表汇总"
OUTPUT_FILENAME = "异动汇总表.xlsx"
TARGET_SHEETS = ("增员", "减员", "转正", "调动", "奖罚扣补")
HEADER_SERIAL = "序号"


@dataclass
class ChangeRow:
    sheet_name: str
    values: list[Any]
    source_file: str
    source_row: int


@dataclass
class PersonnelChangeMergeResult:
    input_dir: Path
    output_dir: Path
    output_file: Path | None = None
    dry_run: bool = False
    source_files: list[str] = field(default_factory=list)
    sheet_counts: dict[str, int] = field(default_factory=dict)
    record_count: int = 0
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "tool_name": TOOL_NAME,
            "input_dir": str(self.input_dir),
            "output_dir": str(self.output_dir),
            "output_file": None if self.output_file is None else str(self.output_file),
            "dry_run": self.dry_run,
            "source_files": self.source_files,
            "source_file_count": len(self.source_files),
            "sheet_counts": self.sheet_counts,
            "record_count": self.record_count,
            "warnings": self.warnings,
        }


@dataclass(frozen=True)
class ChangeSheetLayout:
    sheet_name: str
    header_row: int
    data_start_row: int
    max_column: int


def merge_personnel_changes(
    input_dir: str | Path,
    output_dir: str | Path,
    *,
    template_path: str | Path | None = None,
    dry_run: bool = False,
) -> PersonnelChangeMergeResult:
    input_dir = Path(input_dir).expanduser().resolve()
    output_dir = Path(output_dir).expanduser().resolve()
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"异动表文件夹不存在：{input_dir}")

    source_files = _find_change_files(input_dir)
    if not source_files:
        raise ValueError("未在所选文件夹中找到 .xlsx 异动表")

    base_template = _resolve_template_path(template_path, source_files)
    rows_by_sheet: dict[str, list[ChangeRow]] = {sheet_name: [] for sheet_name in TARGET_SHEETS}
    warnings: list[str] = []
    used_files: list[str] = []

    for file_path in source_files:
        file_rows, file_warnings = _read_change_file(file_path)
        warnings.extend(file_warnings)
        if any(file_rows.values()):
            used_files.append(str(file_path))
        for sheet_name, rows in file_rows.items():
            rows_by_sheet[sheet_name].extend(rows)

    sheet_counts = {sheet_name: len(rows) for sheet_name, rows in rows_by_sheet.items()}
    record_count = sum(sheet_counts.values())
    result = PersonnelChangeMergeResult(
        input_dir=input_dir,
        output_dir=output_dir,
        dry_run=dry_run,
        source_files=used_files,
        sheet_counts=sheet_counts,
        record_count=record_count,
        warnings=warnings,
    )
    if dry_run:
        return result

    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / OUTPUT_FILENAME
    _write_summary_workbook(base_template, output_file, rows_by_sheet)
    result.output_file = output_file
    return result


def _find_change_files(input_dir: Path) -> list[Path]:
    return sorted(
        path
        for path in input_dir.glob("*.xlsx")
        if path.is_file()
        and not path.name.startswith("~$")
        and path.name != OUTPUT_FILENAME
    )


def _resolve_template_path(template_path: str | Path | None, source_files: list[Path]) -> Path:
    if template_path:
        path = Path(template_path).expanduser().resolve()
        if not path.exists() or not path.is_file():
            raise FileNotFoundError(f"异动表模板不存在：{path}")
        return path
    return source_files[0]


def _read_change_file(file_path: Path) -> tuple[dict[str, list[ChangeRow]], list[str]]:
    warnings: list[str] = []
    rows_by_sheet: dict[str, list[ChangeRow]] = {sheet_name: [] for sheet_name in TARGET_SHEETS}
    workbook = load_workbook(file_path, data_only=True)
    try:
        for sheet_name in TARGET_SHEETS:
            if sheet_name not in workbook.sheetnames:
                warnings.append(f"{file_path.name} 缺少工作表：{sheet_name}")
                continue
            ws = workbook[sheet_name]
            layout = _detect_sheet_layout(ws)
            rows_by_sheet[sheet_name].extend(_read_data_rows(ws, layout, file_path.name))
    finally:
        workbook.close()
    return rows_by_sheet, warnings


def _detect_sheet_layout(ws: Worksheet) -> ChangeSheetLayout:
    header_row = _find_header_row(ws)
    return ChangeSheetLayout(
        sheet_name=ws.title,
        header_row=header_row,
        data_start_row=header_row + 1,
        max_column=_last_header_column(ws, header_row),
    )


def _find_header_row(ws: Worksheet) -> int:
    for row_index in range(1, min(ws.max_row, 20) + 1):
        for col_index in range(1, ws.max_column + 1):
            value = ws.cell(row_index, col_index).value
            if str(value or "").strip() == HEADER_SERIAL:
                return row_index
    raise ValueError(f"{ws.title} 未在前 20 行找到“{HEADER_SERIAL}”表头")


def _last_header_column(ws: Worksheet, header_row: int) -> int:
    last_column = 1
    for col_index in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col_index).value
        if value not in (None, ""):
            last_column = col_index
    return last_column


def _read_data_rows(ws: Worksheet, layout: ChangeSheetLayout, source_file: str) -> list[ChangeRow]:
    rows: list[ChangeRow] = []
    for row_index in range(layout.data_start_row, ws.max_row + 1):
        values = [ws.cell(row_index, col_index).value for col_index in range(1, layout.max_column + 1)]
        if not _is_filled_change_row(values):
            continue
        rows.append(
            ChangeRow(
                sheet_name=layout.sheet_name,
                values=values,
                source_file=source_file,
                source_row=row_index,
            )
        )
    return rows


def _is_filled_change_row(values: list[Any]) -> bool:
    # 模板中常预填“序号”，但其他列为空；这种行不是有效异动记录。
    return any(_has_value(value) for value in values[1:])


def _has_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip() != ""
    return True


def _write_summary_workbook(
    template_path: Path,
    output_file: Path,
    rows_by_sheet: dict[str, list[ChangeRow]],
) -> None:
    workbook = load_workbook(template_path)
    try:
        missing_sheets = [sheet_name for sheet_name in TARGET_SHEETS if sheet_name not in workbook.sheetnames]
        if missing_sheets:
            raise ValueError(f"异动表模板缺少工作表：{'、'.join(missing_sheets)}")
        for sheet_name in TARGET_SHEETS:
            ws = workbook[sheet_name]
            layout = _detect_sheet_layout(ws)
            _clear_data_area(ws, layout)
            _write_sheet_rows(ws, layout, rows_by_sheet.get(sheet_name, []))
        workbook.save(output_file)
    finally:
        workbook.close()


def _clear_data_area(ws: Worksheet, layout: ChangeSheetLayout) -> None:
    for row_index in range(layout.data_start_row, ws.max_row + 1):
        for col_index in range(1, layout.max_column + 1):
            ws.cell(row_index, col_index).value = None


def _write_sheet_rows(ws: Worksheet, layout: ChangeSheetLayout, rows: list[ChangeRow]) -> None:
    template_row = layout.data_start_row
    for offset, row in enumerate(rows):
        target_row = layout.data_start_row + offset
        if target_row != template_row:
            _copy_row_style(ws, template_row, target_row, layout.max_column)
        ws.cell(target_row, 1).value = offset + 1
        for col_index in range(2, layout.max_column + 1):
            ws.cell(target_row, col_index).value = row.values[col_index - 1]


def _copy_row_style(ws: Worksheet, source_row: int, target_row: int, max_column: int) -> None:
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col_index in range(1, max_column + 1):
        source = ws.cell(source_row, col_index)
        target = ws.cell(target_row, col_index)
        clone_style(source, target)
