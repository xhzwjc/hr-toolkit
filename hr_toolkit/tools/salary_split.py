from __future__ import annotations

import json
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from ..common.excel import (
    RowSnapshot,
    apply_row_snapshot,
    snapshot_row,
    unmerge_ranges_from_row,
)
from ..common.filenames import safe_filename


TOOL_NAME = "需求4-工资表按入职公司拆分"
DETAIL_SHEET_KEYWORD = "明细"
SUMMARY_SHEET_KEYWORD = "汇总"
HEADER_COMPANY = "入职公司"
HEADER_PROJECT = "项目"
HEADER_NAME = "姓名"
HEADER_ID_CARD = "身份证号码"
HEADER_SEQ = "序号"
DETAIL_NON_TOTAL_HEADERS = (
    "区域",
    "项目",
    "岗位",
    "入职公司",
    "卡号",
    "开户行",
    "手机号",
    "专项扣款",
)


@dataclass
class EmployeeRow:
    source_row: int
    company: str
    section: str
    snapshot: RowSnapshot


@dataclass(frozen=True)
class DetailSection:
    label: str
    source_start_row: int
    source_end_row: int
    subtotal_row: int
    subtotal_snapshot: RowSnapshot


@dataclass(frozen=True)
class RenderedSection:
    label: str
    summary_name: str
    employee_count: int
    data_start_row: int
    data_end_row: int
    subtotal_row: int


@dataclass
class CompanyOutput:
    company: str
    employee_count: int
    sections: list[str]
    file_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "company": self.company,
            "employee_count": self.employee_count,
            "sections": self.sections,
            "projects": self.sections,
            "file_path": self.file_path,
        }


@dataclass
class SalarySplitResult:
    input_path: Path
    output_dir: Path
    dry_run: bool
    outputs: list[CompanyOutput] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "tool_name": TOOL_NAME,
            "input_path": str(self.input_path),
            "output_dir": str(self.output_dir),
            "dry_run": self.dry_run,
            "company_count": len(self.outputs),
            "employee_count": sum(item.employee_count for item in self.outputs),
            "outputs": [item.to_dict() for item in self.outputs],
        }


@dataclass(frozen=True)
class SalarySheetLayout:
    detail_sheet_name: str
    summary_sheet_name: str
    header_row: int
    data_start_row: int
    max_column: int
    seq_col: int
    name_col: int
    id_card_col: int
    project_col: int
    company_col: int
    amount_end_col: int


def split_salary_by_company(
    input_path: str | Path,
    output_dir: str | Path,
    *,
    dry_run: bool = False,
    write_manifest: bool = False,
) -> SalarySplitResult:
    """Split one salary workbook into one workbook per hiring company."""
    input_path = Path(input_path).expanduser().resolve()
    output_dir = Path(output_dir).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")
    if input_path.suffix.lower() != ".xlsx":
        raise ValueError("当前工资拆分工具仅支持 .xlsx 文件")

    workbook = load_workbook(input_path, data_only=False)
    layout = _detect_layout(workbook)
    detail_ws = workbook[layout.detail_sheet_name]
    sections = _detect_detail_sections(detail_ws, layout)
    employees = _collect_employees(detail_ws, layout, sections)
    groups = _group_by_company(employees)

    result = SalarySplitResult(input_path=input_path, output_dir=output_dir, dry_run=dry_run)
    for company, rows in groups.items():
        result.outputs.append(
            CompanyOutput(
                company=company,
                employee_count=len(rows),
                sections=list(_group_sections(rows).keys()),
            )
        )

    if dry_run:
        return result

    output_dir.mkdir(parents=True, exist_ok=True)
    for company_output in result.outputs:
        rows = groups[company_output.company]
        output_path = output_dir / f"{safe_filename(company_output.company)}-工资表.xlsx"
        _write_company_workbook(input_path, layout, company_output.company, rows, output_path)
        company_output.file_path = str(output_path)

    if write_manifest:
        manifest_path = output_dir / "_salary_split_manifest.json"
        manifest_path.write_text(
            json.dumps(result.to_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    return result


def _detect_layout(workbook) -> SalarySheetLayout:
    detail_sheet_name = _find_sheet_name(workbook.sheetnames, DETAIL_SHEET_KEYWORD)
    summary_sheet_name = _find_sheet_name(workbook.sheetnames, SUMMARY_SHEET_KEYWORD)
    detail_ws = workbook[detail_sheet_name]

    header_row = _find_header_row(detail_ws, HEADER_COMPANY)
    headers = _read_headers(detail_ws, header_row)
    try:
        return SalarySheetLayout(
            detail_sheet_name=detail_sheet_name,
            summary_sheet_name=summary_sheet_name,
            header_row=header_row,
            data_start_row=_find_data_start_row(detail_ws, header_row),
            max_column=detail_ws.max_column,
            seq_col=headers[HEADER_SEQ],
            name_col=headers[HEADER_NAME],
            id_card_col=headers[HEADER_ID_CARD],
            project_col=headers[HEADER_PROJECT],
            company_col=headers[HEADER_COMPANY],
            amount_end_col=_find_amount_end_col(headers, detail_ws.max_column),
        )
    except KeyError as exc:
        raise ValueError(f"明细表缺少必要字段：{exc.args[0]}") from exc


def _find_sheet_name(sheetnames: list[str], keyword: str) -> str:
    for sheetname in sheetnames:
        if keyword in sheetname:
            return sheetname
    raise ValueError(f"未找到包含“{keyword}”的工作表")


def _find_header_row(ws: Worksheet, required_header: str) -> int:
    for row_index in range(1, min(ws.max_row, 20) + 1):
        values = [str(ws.cell(row_index, col).value or "").strip() for col in range(1, ws.max_column + 1)]
        if required_header in values:
            return row_index
    raise ValueError(f"未在明细表前 20 行找到字段：{required_header}")


def _read_headers(ws: Worksheet, header_row: int) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col_index in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col_index).value
        if value is None:
            continue
        text = str(value).strip()
        if text and text not in headers:
            headers[text] = col_index
    return headers


def _find_amount_end_col(headers: dict[str, int], max_column: int) -> int:
    non_total_cols = [
        col_index
        for header, col_index in headers.items()
        if header in DETAIL_NON_TOTAL_HEADERS
    ]
    if not non_total_cols:
        return max_column
    return min(non_total_cols) - 1


def _find_data_start_row(ws: Worksheet, header_row: int) -> int:
    bottom = header_row
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= header_row <= merged_range.max_row:
            bottom = max(bottom, merged_range.max_row)
    return bottom + 1


def _collect_employees(
    ws: Worksheet,
    layout: SalarySheetLayout,
    sections: list[DetailSection],
) -> list[EmployeeRow]:
    employees: list[EmployeeRow] = []
    for row_index in range(layout.data_start_row, ws.max_row + 1):
        company = _cell_text(ws, row_index, layout.company_col)
        name = _cell_text(ws, row_index, layout.name_col)
        id_card = _cell_text(ws, row_index, layout.id_card_col)
        if not company:
            continue
        if not name and not id_card:
            continue
        section = _find_section_for_row(row_index, sections)
        employees.append(
            EmployeeRow(
                source_row=row_index,
                company=company,
                section=section.label,
                snapshot=snapshot_row(ws, row_index, layout.max_column),
            )
        )
    if not employees:
        raise ValueError("未识别到可拆分的员工数据，请检查明细表的“入职公司”列")
    return employees


def _group_by_company(employees: list[EmployeeRow]) -> OrderedDict[str, list[EmployeeRow]]:
    groups: OrderedDict[str, list[EmployeeRow]] = OrderedDict()
    for employee in employees:
        groups.setdefault(employee.company, []).append(employee)
    return groups


def _group_sections(rows: list[EmployeeRow]) -> OrderedDict[str, list[EmployeeRow]]:
    sections: OrderedDict[str, list[EmployeeRow]] = OrderedDict()
    for row in rows:
        sections.setdefault(row.section, []).append(row)
    return sections


def _write_company_workbook(
    input_path: Path,
    layout: SalarySheetLayout,
    company: str,
    rows: list[EmployeeRow],
    output_path: Path,
) -> None:
    workbook = load_workbook(input_path, data_only=False)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True

    detail_ws = workbook[layout.detail_sheet_name]
    summary_ws = workbook[layout.summary_sheet_name]
    rendered_sections = _rebuild_detail_sheet(detail_ws, layout, rows)
    _rebuild_summary_sheet(summary_ws, layout, rendered_sections)
    workbook.save(output_path)


def _rebuild_detail_sheet(
    ws: Worksheet,
    layout: SalarySheetLayout,
    rows: list[EmployeeRow],
) -> list[RenderedSection]:
    sections = _detect_detail_sections(ws, layout)
    total_template_row = _find_last_summary_row(ws, layout.data_start_row)
    total_snapshot = snapshot_row(ws, total_template_row, layout.max_column)
    total_label = ws.cell(total_template_row, 1).value

    unmerge_ranges_from_row(ws, layout.data_start_row)
    ws.delete_rows(layout.data_start_row, ws.max_row - layout.data_start_row + 1)

    current_row = layout.data_start_row
    employee_index = 1
    rendered_sections: list[RenderedSection] = []
    rows_by_section = _group_sections(rows)
    for section in sections:
        section_rows = rows_by_section.get(section.label, [])
        data_start_row = current_row
        for employee in section_rows:
            apply_row_snapshot(ws, current_row, employee.snapshot, translate_formulas=True)
            ws.cell(current_row, layout.seq_col).value = employee_index
            current_row += 1
            employee_index += 1

        data_end_row = current_row - 1
        subtotal_row = current_row
        apply_row_snapshot(ws, subtotal_row, section.subtotal_snapshot, translate_formulas=False)
        ws.merge_cells(start_row=subtotal_row, start_column=1, end_row=subtotal_row, end_column=3)
        ws.cell(subtotal_row, 1).value = section.label
        _write_detail_section_total_formulas(ws, layout, subtotal_row, data_start_row, data_end_row)
        rendered_sections.append(
            RenderedSection(
                label=section.label,
                summary_name=section.label,
                employee_count=len(section_rows),
                data_start_row=data_start_row,
                data_end_row=data_end_row,
                subtotal_row=subtotal_row,
            )
        )
        current_row += 1

    total_row = current_row
    apply_row_snapshot(ws, total_row, total_snapshot, translate_formulas=False)
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    ws.cell(total_row, 1).value = total_label
    _write_detail_grand_total_formulas(
        ws,
        layout,
        total_row,
        [section.subtotal_row for section in rendered_sections],
    )
    return rendered_sections


def _detect_detail_sections(ws: Worksheet, layout: SalarySheetLayout) -> list[DetailSection]:
    sections: list[DetailSection] = []
    section_start = layout.data_start_row
    for row_index in range(layout.data_start_row, ws.max_row + 1):
        label = _cell_text(ws, row_index, 1)
        if not label or "合计" not in label:
            continue
        if "总计" in label:
            break
        sections.append(
            DetailSection(
                label=label,
                source_start_row=section_start,
                source_end_row=row_index - 1,
                subtotal_row=row_index,
                subtotal_snapshot=snapshot_row(ws, row_index, layout.max_column),
            )
        )
        section_start = row_index + 1
    if not sections:
        raise ValueError("未识别到明细表中的分段小计行，例如“河源无线代维合计”")
    return sections


def _find_section_for_row(row_index: int, sections: list[DetailSection]) -> DetailSection:
    for section in sections:
        if section.source_start_row <= row_index <= section.source_end_row:
            return section
    raise ValueError(f"员工数据行 {row_index} 未落在任何分段小计范围内")


def _find_last_summary_row(ws: Worksheet, data_start_row: int) -> int:
    for row_index in range(ws.max_row, data_start_row - 1, -1):
        value = ws.cell(row_index, 1).value
        if isinstance(value, str) and "总计" in value:
            return row_index
    for row_index in range(ws.max_row, data_start_row - 1, -1):
        value = ws.cell(row_index, 1).value
        if isinstance(value, str) and "合计" in value:
            return row_index
    return data_start_row


def _write_detail_section_total_formulas(
    ws: Worksheet,
    layout: SalarySheetLayout,
    subtotal_row: int,
    first_data_row: int,
    last_data_row: int,
) -> None:
    for col_index in range(4, layout.max_column + 1):
        col_letter = get_column_letter(col_index)
        cell = ws.cell(subtotal_row, col_index)
        if col_index == layout.id_card_col:
            cell.value = None
        elif col_index <= layout.amount_end_col:
            if last_data_row >= first_data_row:
                cell.value = f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})"
            else:
                cell.value = 0
        else:
            cell.value = None


def _write_detail_grand_total_formulas(
    ws: Worksheet,
    layout: SalarySheetLayout,
    total_row: int,
    subtotal_rows: list[int],
) -> None:
    for col_index in range(4, layout.max_column + 1):
        col_letter = get_column_letter(col_index)
        cell = ws.cell(total_row, col_index)
        if col_index == layout.id_card_col:
            cell.value = None
        elif col_index <= layout.amount_end_col:
            if subtotal_rows:
                cell.value = "=" + "+".join(f"{col_letter}{row}" for row in subtotal_rows)
            else:
                cell.value = 0
        else:
            cell.value = None


def _rebuild_summary_sheet(
    ws: Worksheet,
    layout: SalarySheetLayout,
    rendered_sections: list[RenderedSection],
) -> None:
    project_template_row = 6
    total_template_row = _find_summary_total_row(ws, project_template_row)
    sign_template_row = total_template_row + 1
    project_snapshots = [
        snapshot_row(ws, row_index, ws.max_column)
        for row_index in range(project_template_row, total_template_row)
    ]
    project_names = [
        _cell_text(ws, row_index, 1) or section.label
        for row_index, section in zip(
            range(project_template_row, project_template_row + len(rendered_sections)),
            rendered_sections,
        )
    ]
    total_snapshot = snapshot_row(ws, total_template_row, ws.max_column)
    sign_snapshot = snapshot_row(ws, sign_template_row, ws.max_column)
    signature_text = ws.cell(sign_template_row, 1).value

    unmerge_ranges_from_row(ws, project_template_row)
    ws.delete_rows(project_template_row, ws.max_row - project_template_row + 1)

    start_row = project_template_row
    current_row = start_row
    for index, section in enumerate(rendered_sections):
        snapshot = project_snapshots[index] if index < len(project_snapshots) else project_snapshots[-1]
        summary_name = project_names[index] if index < len(project_names) else section.label
        rendered_section = RenderedSection(
            label=section.label,
            summary_name=summary_name,
            employee_count=section.employee_count,
            data_start_row=section.data_start_row,
            data_end_row=section.data_end_row,
            subtotal_row=section.subtotal_row,
        )
        apply_row_snapshot(ws, current_row, snapshot, translate_formulas=False)
        ws.cell(current_row, 1).value = rendered_section.summary_name
        _write_summary_section_formulas(ws, current_row, layout.detail_sheet_name, rendered_section)
        current_row += 1

    total_row = current_row
    apply_row_snapshot(ws, total_row, total_snapshot, translate_formulas=False)
    ws.cell(total_row, 1).value = "合计"
    _write_summary_total_formulas(ws, start_row, total_row)

    signature_row = total_row + 1
    apply_row_snapshot(ws, signature_row, sign_snapshot, translate_formulas=False)
    ws.merge_cells(start_row=signature_row, start_column=1, end_row=signature_row, end_column=21)
    ws.cell(signature_row, 1).value = signature_text


def _find_summary_total_row(ws: Worksheet, start_row: int) -> int:
    for row_index in range(start_row, ws.max_row + 1):
        if _cell_text(ws, row_index, 1) == "合计":
            return row_index
    raise ValueError("汇总表中未找到“合计”行")


def _write_summary_section_formulas(
    ws: Worksheet,
    row_index: int,
    detail_sheet_name: str,
    section: RenderedSection,
) -> None:
    detail = _formula_sheet_name(detail_sheet_name)
    subtotal_row = section.subtotal_row
    if section.employee_count and section.data_end_row >= section.data_start_row:
        count_formula: str | int = (
            f"=COUNT({detail}!$A${section.data_start_row}:$A${section.data_end_row})"
        )
    else:
        count_formula = 0
    formulas = {
        2: count_formula,
        3: f"={detail}!P{subtotal_row}",
        4: f"={detail}!U{subtotal_row}",
        5: f"={detail}!M{subtotal_row}+{detail}!N{subtotal_row}+{detail}!O{subtotal_row}+{detail}!S{subtotal_row}",
        6: f"={detail}!R{subtotal_row}",
        7: f"={detail}!X{subtotal_row}",
        8: f"={detail}!Y{subtotal_row}",
        9: f"={detail}!AA{subtotal_row}",
        10: f"={detail}!AB{subtotal_row}",
        11: f"={detail}!AD{subtotal_row}",
        12: f"={detail}!AE{subtotal_row}",
        13: f"={detail}!AG{subtotal_row}",
        14: f"={detail}!AH{subtotal_row}",
        15: f"={detail}!AJ{subtotal_row}",
        16: f"={detail}!AM{subtotal_row}",
        17: f"={detail}!AN{subtotal_row}",
        18: f"={detail}!AO{subtotal_row}",
        19: f"={detail}!AK{subtotal_row}",
        20: f"={detail}!AP{subtotal_row}",
        21: f"=SUM(D{row_index}:T{row_index})",
    }
    for col_index, formula in formulas.items():
        ws.cell(row_index, col_index).value = formula


def _write_summary_total_formulas(ws: Worksheet, first_row: int, total_row: int) -> None:
    last_row = total_row - 1
    for col_index in range(2, 22):
        col_letter = get_column_letter(col_index)
        if last_row >= first_row:
            ws.cell(total_row, col_index).value = f"=SUM({col_letter}{first_row}:{col_letter}{last_row})"
        else:
            ws.cell(total_row, col_index).value = 0


def _cell_text(ws: Worksheet, row_index: int, col_index: int) -> str:
    value = ws.cell(row_index, col_index).value
    if value is None:
        return ""
    return str(value).strip()


def _formula_sheet_name(sheet_name: str) -> str:
    escaped = sheet_name.replace("'", "''")
    return f"'{escaped}'"
