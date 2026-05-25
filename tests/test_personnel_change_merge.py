from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from hr_toolkit.tools.personnel_change_merge import merge_personnel_changes


class PersonnelChangeMergeTest(unittest.TestCase):
    def test_merge_multiple_project_change_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            output_dir = root / "output"
            input_dir.mkdir()

            _write_change_file(
                input_dir / "项目A异动表.xlsx",
                {
                    "增员": [
                        ["员工1", "44162219901007667X", "生产人员"],
                        ["员工2", "441622198404210312", "管理人员"],
                    ],
                    "减员": [["员工3", "44162219800516649X", "离职"]],
                },
            )
            _write_change_file(
                input_dir / "项目B异动表.xlsx",
                {
                    "增员": [["员工4", "44132419860927333X", "生产人员"]],
                    "奖罚扣补": [["员工5", 0, 100, 0, 50]],
                },
            )

            result = merge_personnel_changes(input_dir, output_dir)
            payload = result.to_dict()

            self.assertEqual(payload["source_file_count"], 2)
            self.assertEqual(payload["record_count"], 5)
            self.assertEqual(payload["sheet_counts"]["增员"], 3)
            self.assertEqual(payload["sheet_counts"]["减员"], 1)
            self.assertEqual(payload["sheet_counts"]["奖罚扣补"], 1)
            self.assertTrue(result.output_file and result.output_file.exists())

            wb = load_workbook(result.output_file, data_only=True)
            add_ws = wb["增员"]
            self.assertEqual([add_ws.cell(3, col).value for col in range(1, 5)], [1, "员工1", "44162219901007667X", "生产人员"])
            self.assertEqual([add_ws.cell(5, col).value for col in range(1, 5)], [3, "员工4", "44132419860927333X", "生产人员"])
            self.assertIsNone(add_ws.cell(6, 2).value)

            leave_ws = wb["减员"]
            self.assertEqual([leave_ws.cell(3, col).value for col in range(1, 5)], [1, "员工3", "44162219800516649X", "离职"])

            reward_ws = wb["奖罚扣补"]
            self.assertEqual([reward_ws.cell(3, col).value for col in range(1, 7)], [1, "员工5", 0, 100, 0, 50])

    def test_dry_run_does_not_write_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            output_dir = root / "output"
            input_dir.mkdir()
            _write_change_file(input_dir / "项目A异动表.xlsx", {"增员": [["员工1", "44162219901007667X", "生产人员"]]})

            result = merge_personnel_changes(input_dir, output_dir, dry_run=True)

            self.assertEqual(result.record_count, 1)
            self.assertIsNone(result.output_file)
            self.assertFalse(output_dir.exists())

    def test_missing_template_sheet_reports_clear_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            output_dir = root / "output"
            template_path = root / "缺少调动模板.xlsx"
            input_dir.mkdir()
            _write_change_file(input_dir / "项目A异动表.xlsx", {"增员": [["员工1", "44162219901007667X", "生产人员"]]})
            _write_change_file(template_path, {}, omit_sheets={"调动"})

            with self.assertRaisesRegex(ValueError, "异动表模板缺少工作表：调动"):
                merge_personnel_changes(input_dir, output_dir, template_path=template_path)


def _write_change_file(path: Path, rows_by_sheet: dict[str, list[list]], omit_sheets: set[str] | None = None) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    omit_sheets = omit_sheets or set()
    sheet_headers = {
        "增员": ["序号", "姓名", "身份证号码", "人员分类"],
        "减员": ["序号", "姓名", "身份证号码", "备注"],
        "转正": ["序号", "姓名", "入职日期", "转正日期"],
        "调动": ["序号", "姓名", "原部门", "现部门"],
        "奖罚扣补": ["序号", "姓名", "罚（元）", "奖（元）", "扣（元）", "补（元）"],
    }
    for sheet_name, headers in sheet_headers.items():
        if sheet_name in omit_sheets:
            continue
        ws = workbook.create_sheet(sheet_name)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        ws["A1"] = f"2026年4月{sheet_name}表"
        for col_index, header in enumerate(headers, start=1):
            ws.cell(2, col_index).value = header
        for row_index in range(3, 8):
            ws.cell(row_index, 1).value = row_index - 2

        rows = rows_by_sheet.get(sheet_name, [])
        for offset, values in enumerate(rows):
            row_index = 3 + offset
            ws.cell(row_index, 1).value = offset + 1
            for col_index, value in enumerate(values, start=2):
                ws.cell(row_index, col_index).value = value
    workbook.save(path)


if __name__ == "__main__":
    unittest.main()
