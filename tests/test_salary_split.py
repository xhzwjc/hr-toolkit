from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from hr_toolkit.tools.salary_split import split_salary_by_company


class SalarySplitTest(unittest.TestCase):
    def test_split_sample_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sample = root / "脱敏工资表.xlsx"
            out_dir = root / "output"
            _write_salary_split_sample(sample)
            result = split_salary_by_company(sample, out_dir)
            payload = result.to_dict()

            self.assertEqual(payload["company_count"], 3)
            self.assertEqual(payload["employee_count"], 23)
            self.assertFalse((out_dir / "_salary_split_manifest.json").exists())

            companies = {item["company"]: item for item in payload["outputs"]}
            self.assertEqual(companies["春苗北京"]["employee_count"], 18)
            self.assertEqual(companies["唐人"]["employee_count"], 4)
            self.assertEqual(companies["岩亨"]["employee_count"], 1)

            wb = load_workbook(companies["春苗北京"]["file_path"], data_only=False)
            detail = wb["明细表"]
            self.assertEqual(detail["A18"].value, "河源无线代维合计")
            self.assertEqual(detail["A25"].value, "河源传输代维合计")
            self.assertEqual(detail["A26"].value, "广东分公司（河源项目部）总计")
            self.assertEqual(detail["B6"].value, "员工1")
            self.assertEqual(detail["B24"].value, "员工23")
            self.assertEqual(detail["AU6"].value, "春苗北京")
            self.assertEqual(detail["P18"].value, "=SUM(P6:P17)")
            self.assertEqual(detail["P25"].value, "=SUM(P19:P24)")
            self.assertEqual(detail["P26"].value, "=P18+P25")

            summary = wb["汇总表"]
            self.assertEqual(summary["A6"].value, "广东河源市2026年4月移动基站代维项目")
            self.assertEqual(summary["A7"].value, "广东河源市2026年4月移动线路代维项目")
            self.assertEqual(summary["C6"].value, "='明细表'!P18")
            self.assertEqual(summary["C7"].value, "='明细表'!P25")
            wb.close()

            empty_section_wb = load_workbook(companies["岩亨"]["file_path"], data_only=False)
            empty_detail = empty_section_wb["明细表"]
            detail_labels = [empty_detail.cell(row, 1).value for row in range(1, empty_detail.max_row + 1)]
            self.assertIn("河源无线代维合计", detail_labels)
            self.assertNotIn("河源传输代维合计", detail_labels)
            empty_summary = empty_section_wb["汇总表"]
            summary_labels = [empty_summary.cell(row, 1).value for row in range(1, empty_summary.max_row + 1)]
            self.assertIn("广东河源市2026年4月移动基站代维项目", summary_labels)
            self.assertNotIn("广东河源市2026年4月移动线路代维项目", summary_labels)
            self.assertEqual(empty_summary["A7"].value, "合计")
            self.assertEqual(empty_summary["C7"].value, "=SUM(C6:C6)")
            empty_section_wb.close()

    def test_manifest_is_optional(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sample = root / "脱敏工资表.xlsx"
            out_dir = root / "output"
            _write_salary_split_sample(sample)
            split_salary_by_company(sample, out_dir, write_manifest=True)
            manifest_path = out_dir / "_salary_split_manifest.json"
            self.assertTrue(manifest_path.exists())

            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["tool_name"], "需求4-工资表按入职公司拆分")


def _write_salary_split_sample(path: Path) -> None:
    """Create a sanitized two-section salary sheet matching the production layout."""
    workbook = Workbook()
    detail = workbook.active
    detail.title = "明细表"
    summary = workbook.create_sheet("汇总表")

    detail["A1"] = "广东分公司（河源项目部）工资表"
    headers = [f"金额{index}" for index in range(1, 46)] + ["项目", "入职公司"]
    headers[0] = "序号"
    headers[1] = "姓名"
    headers[3] = "身份证号码"
    headers[15] = "应发小计"
    for column, header in enumerate(headers, start=1):
        detail.cell(5, column).value = header

    section_one = [
        *(f"员工{index}" for index in range(1, 13)),
        "唐人员工13",
        "唐人员工14",
        "唐人员工15",
        "岩亨员工16",
    ]
    section_two = ["唐人员工17", *(f"员工{index}" for index in range(18, 24))]
    current_row = 6
    for section_name, names in (("河源无线代维合计", section_one), ("河源传输代维合计", section_two)):
        for name in names:
            detail.cell(current_row, 1).value = current_row - 5
            detail.cell(current_row, 2).value = name
            detail.cell(current_row, 4).value = f"44010019900101{current_row:04d}"
            detail.cell(current_row, 16).value = 100
            detail.cell(current_row, 46).value = section_name.replace("合计", "项目")
            if name.startswith("员工"):
                detail.cell(current_row, 47).value = "春苗北京"
            elif name.startswith("唐人"):
                detail.cell(current_row, 47).value = "唐人"
            else:
                detail.cell(current_row, 47).value = "岩亨"
            current_row += 1
        detail.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        detail.cell(current_row, 1).value = section_name
        detail.cell(current_row, 16).value = "=SUM(P6:P6)"
        current_row += 1

    detail.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
    detail.cell(current_row, 1).value = "广东分公司（河源项目部）总计"
    detail.cell(current_row, 16).value = "=SUM(P6:P6)"

    summary["A1"] = "工资汇总表"
    summary["A6"] = "广东河源市2026年4月移动基站代维项目"
    summary["A7"] = "广东河源市2026年4月移动线路代维项目"
    summary["A8"] = "合计"
    summary.merge_cells(start_row=9, start_column=1, end_row=9, end_column=21)
    summary["A9"] = "制表："
    workbook.save(path)
    workbook.close()


if __name__ == "__main__":
    unittest.main()
