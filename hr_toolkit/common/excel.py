from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from typing import Any

from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class CellSnapshot:
    value: Any
    font: Any
    fill: Any
    border: Any
    alignment: Any
    number_format: str
    protection: Any


@dataclass(frozen=True)
class RowSnapshot:
    source_row: int
    height: float | None
    cells: list[CellSnapshot]


def snapshot_row(ws: Worksheet, row_index: int, max_column: int) -> RowSnapshot:
    cells: list[CellSnapshot] = []
    for col_index in range(1, max_column + 1):
        cell = ws.cell(row_index, col_index)
        cells.append(
            CellSnapshot(
                value=cell.value,
                font=copy(cell.font),
                fill=copy(cell.fill),
                border=copy(cell.border),
                alignment=copy(cell.alignment),
                number_format=cell.number_format,
                protection=copy(cell.protection),
            )
        )
    return RowSnapshot(
        source_row=row_index,
        height=ws.row_dimensions[row_index].height,
        cells=cells,
    )


def apply_row_snapshot(
    ws: Worksheet,
    target_row: int,
    snapshot: RowSnapshot,
    *,
    translate_formulas: bool = True,
) -> None:
    ws.row_dimensions[target_row].height = snapshot.height
    for col_index, snap in enumerate(snapshot.cells, start=1):
        cell = ws.cell(target_row, col_index)
        cell.font = copy(snap.font)
        cell.fill = copy(snap.fill)
        cell.border = copy(snap.border)
        cell.alignment = copy(snap.alignment)
        cell.number_format = snap.number_format
        cell.protection = copy(snap.protection)

        value = snap.value
        if translate_formulas and isinstance(value, str) and value.startswith("="):
            origin = f"{get_column_letter(col_index)}{snapshot.source_row}"
            destination = f"{get_column_letter(col_index)}{target_row}"
            try:
                value = Translator(value, origin=origin).translate_formula(destination)
            except Exception:
                value = _translate_same_row_formula(value, snapshot.source_row, target_row)
        cell.value = value


def clone_style(source_cell, target_cell) -> None:
    target_cell.font = copy(source_cell.font)
    target_cell.fill = copy(source_cell.fill)
    target_cell.border = copy(source_cell.border)
    target_cell.alignment = copy(source_cell.alignment)
    target_cell.number_format = source_cell.number_format
    target_cell.protection = copy(source_cell.protection)


def unmerge_ranges_from_row(ws: Worksheet, min_row: int) -> None:
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row >= min_row:
            ws.unmerge_cells(str(merged_range))


def _translate_same_row_formula(formula: str, source_row: int, target_row: int) -> str:
    return formula.replace(str(source_row), str(target_row))

