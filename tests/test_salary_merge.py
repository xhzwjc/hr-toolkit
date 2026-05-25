from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

from hr_toolkit.tools.salary_merge import AMOUNT_NUMBER_FORMAT, SUMMARY_TITLE, merge_monthly_salary


class SalaryMergeTest(unittest.TestCase):
    def test_merge_monthly_salary_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            output_dir = root / "output"
            input_dir.mkdir()

            _write_salary_file(
                input_dir / "河源项目部工资表_202604.xlsx",
                [
                    ("员工1", "44162219901007667X", 500),
                    ("员工2", "441622198404210312", 600),
                ],
            )
            _write_salary_file(
                input_dir / "河源项目部工资表_202605.xlsx",
                [
                    ("员工1", "44162219901007667X", 700),
                    ("员工3", "44162219800516649X", 800),
                ],
            )

            result = merge_monthly_salary(input_dir, output_dir, year=2026)
            payload = result.to_dict()

            self.assertEqual(payload["source_file_count"], 2)
            self.assertEqual(payload["employee_count"], 3)
            self.assertEqual(payload["record_count"], 4)
            self.assertEqual(payload["applied_record_count"], 4)
            self.assertEqual(payload["skipped_record_count"], 0)
            self.assertEqual(payload["months"][0], "202601")
            self.assertEqual(payload["months"][-1], "202612")
            self.assertTrue(result.output_file and result.output_file.exists())

            wb = load_workbook(result.output_file, data_only=True)
            ws = wb["汇总"]
            self.assertEqual(ws["A1"].value, SUMMARY_TITLE)
            self.assertEqual(ws["A1"].fill.fill_type, None)
            self.assertEqual(ws["D3"].value, 202601)
            self.assertEqual(ws["D3"].fill.fill_type, None)
            self.assertEqual(ws["D5"].number_format, AMOUNT_NUMBER_FORMAT)
            rows = {
                ws.cell(row, 3).value: [ws.cell(row, col).value for col in range(1, 16)]
                for row in range(5, 8)
            }
            self.assertEqual(rows["44162219901007667X"][3:8], [0, 0, 0, 500, 700])
            self.assertEqual(ws["D6"].value, 0)
            self.assertEqual(ws["D6"].number_format, AMOUNT_NUMBER_FORMAT)
            self.assertEqual(rows["441622198404210312"][3:8], [0, 0, 0, 600, 0])
            self.assertEqual(rows["44162219800516649X"][3:8], [0, 0, 0, 0, 800])

    def test_append_to_existing_summary_and_skip_existing_month(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first_input_dir = root / "first_input"
            first_output_dir = root / "first_output"
            second_input_dir = root / "second_input"
            second_output_dir = root / "second_output"
            first_input_dir.mkdir()
            second_input_dir.mkdir()

            _write_salary_file(
                first_input_dir / "河源项目部工资表_202601.xlsx",
                [("员工1", "44162219901007667X", 100)],
            )
            _write_salary_file(
                first_input_dir / "河源项目部工资表_202602.xlsx",
                [
                    ("员工1", "44162219901007667X", 200),
                    ("员工2", "441622198404210312", 300),
                ],
            )
            first_result = merge_monthly_salary(first_input_dir, first_output_dir, year=2026)

            _write_salary_file(
                second_input_dir / "河源项目部工资表_202601.xlsx",
                [("员工1", "44162219901007667X", 999)],
            )
            _write_salary_file(
                second_input_dir / "河源项目部工资表_202603.xlsx",
                [
                    ("员工1", "44162219901007667X", 400),
                    ("员工3", "44162219800516649X", 500),
                ],
            )

            result = merge_monthly_salary(
                second_input_dir,
                second_output_dir,
                existing_summary_path=first_result.output_file,
            )
            payload = result.to_dict()

            self.assertEqual(payload["source_file_count"], 2)
            self.assertEqual(payload["employee_count"], 3)
            self.assertEqual(payload["record_count"], 3)
            self.assertEqual(payload["applied_record_count"], 2)
            self.assertEqual(payload["skipped_record_count"], 1)
            self.assertTrue(any("已存在金额，未覆盖" in warning for warning in payload["warnings"]))

            wb = load_workbook(result.output_file, data_only=True)
            ws = wb["汇总"]
            rows = {
                ws.cell(row, 3).value: [ws.cell(row, col).value for col in range(1, 16)]
                for row in range(5, 8)
            }
            self.assertEqual(rows["44162219901007667X"][3:7], [100, 200, 400, 0])
            self.assertEqual(rows["441622198404210312"][3:7], [0, 300, 0, 0])
            self.assertEqual(rows["44162219800516649X"][3:7], [0, 0, 500, 0])

    def test_detect_month_from_cell_and_accept_id_card_alias(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            output_dir = root / "output"
            input_dir.mkdir()

            _write_salary_file(
                input_dir / "河源项目部工资表.xlsx",
                [("员工1", "44162219901007667X", 1200)],
                id_header="身份证号",
                month_date=date(2026, 6, 1),
            )

            result = merge_monthly_salary(input_dir, output_dir)

            self.assertEqual(result.months, [f"2026{month:02d}" for month in range(1, 13)])
            self.assertEqual(result.employee_count, 1)
            self.assertEqual(result.record_count, 1)
            wb = load_workbook(result.output_file, data_only=True)
            ws = wb["汇总"]
            self.assertEqual(ws["I3"].value, 202606)
            self.assertEqual(ws["I5"].value, 1200)


def _write_salary_file(
    path: Path,
    employees: list[tuple[str, str, int]],
    *,
    id_header: str = "身份证号码",
    month_date: date | None = None,
) -> None:
    workbook = Workbook()
    ws = workbook.active
    ws.title = "明细表"
    if month_date is not None:
        ws["A3"] = month_date
    ws["A4"] = "序号"
    ws["B4"] = "姓名"
    ws["D4"] = id_header
    ws["P4"] = "应发小计"
    for index, (name, id_card, amount) in enumerate(employees, start=1):
        row = index + 4
        ws.cell(row, 1).value = index
        ws.cell(row, 2).value = name
        ws.cell(row, 4).value = id_card
        ws.cell(row, 16).value = amount
    workbook.save(path)


if __name__ == "__main__":
    unittest.main()
