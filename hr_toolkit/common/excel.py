from __future__ import annotations

import re
from copy import copy
from dataclasses import dataclass
from typing import Any, NamedTuple

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


class _GridCell(NamedTuple):
    value: Any


class SheetGrid:
    """把工作表一次性读入内存的轻量值网格。

    openpyxl 的 read_only 模式只适合顺序读取：ws.cell(row, col) 每次都会
    从头重新解析工作表 XML，随机访问会退化成 O(行数²)，大文件要跑几分钟。
    先用 iter_rows 单遍读完，之后在内存里随机访问，行列号仍为 1 起始。

    同时提供 ws.cell(row, col).value 形式的兼容接口，便于原有按
    Worksheet 编写的读取函数直接换用。
    """

    __slots__ = ("title", "max_row", "max_column", "_rows")

    def __init__(self, ws: Any) -> None:
        self.title: str = ws.title
        self._rows: list[tuple[Any, ...]] = [tuple(row) for row in ws.iter_rows(values_only=True)]
        self.max_row: int = len(self._rows)
        self.max_column: int = max((len(row) for row in self._rows), default=0)

    def value(self, row_index: int, col_index: int) -> Any:
        if not 1 <= row_index <= self.max_row or col_index < 1:
            return None
        row = self._rows[row_index - 1]
        return row[col_index - 1] if col_index <= len(row) else None

    def cell(self, row_index: int, col_index: int) -> _GridCell:
        return _GridCell(self.value(row_index, col_index))


def insert_rows(ws: Worksheet, idx: int, amount: int = 1) -> None:
    """在第 idx 行前插入 amount 个空行，内存占用与已用单元格数成正比。

    openpyxl 自带的 ``ws.insert_rows`` 会先 ``list(iter_rows(min_row=idx))``，
    把插入位置以下、所有列的空单元格全部物化成对象再逐个搬动。几千行 × 几十列
    时会瞬间生成上百万个单元格对象而 MemoryError（见 openpyxl
    ``worksheet._move_cells``）。这里只搬动实际存在的稀疏单元格，行为与
    openpyxl 保持一致——同样不改动已有公式引用、不移动合并单元格与行高。
    """
    _shift_cells(ws, min_index=idx, amount=amount, is_row=True)
    ws._current_row = ws.max_row


def insert_cols(ws: Worksheet, idx: int, amount: int = 1) -> None:
    """在第 idx 列前插入 amount 个空列，是 :func:`insert_rows` 的列向版本。"""
    _shift_cells(ws, min_index=idx, amount=amount, is_row=False)


def _shift_cells(ws: Worksheet, *, min_index: int, amount: int, is_row: bool) -> None:
    if amount <= 0:
        return
    cells = ws._cells
    # 只挑出受影响的已存在单元格；从远端向 min_index 方向搬，避免覆盖目标位置
    axis = 0 if is_row else 1
    affected = [key for key in cells if key[axis] >= min_index]
    affected.sort(key=lambda key: key[axis], reverse=True)
    for row, col in affected:
        cell = cells.pop((row, col))
        if is_row:
            row += amount
        else:
            col += amount
        cell.row = row
        cell.column = col
        cells[(row, col)] = cell


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


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
    cell_ref_pattern = re.compile(r"(?<![A-Za-z0-9_])(\$?[A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])")

    def replace_row(match: re.Match[str]) -> str:
        column_ref, row_anchor, row_number = match.groups()
        if row_anchor or int(row_number) != source_row:
            return match.group(0)
        return f"{column_ref}{target_row}"

    return cell_ref_pattern.sub(replace_row, formula)
