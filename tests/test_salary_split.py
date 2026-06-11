from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from hr_toolkit.tools.salary_split import split_salary_by_company


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "附件" / "问题4-薪资表模板(1).xlsx"


class SalarySplitTest(unittest.TestCase):
    def test_split_sample_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp)
            result = split_salary_by_company(SAMPLE, out_dir)
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
            out_dir = Path(tmp)
            split_salary_by_company(SAMPLE, out_dir, write_manifest=True)
            manifest_path = out_dir / "_salary_split_manifest.json"
            self.assertTrue(manifest_path.exists())

            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["tool_name"], "需求4-工资表按入职公司拆分")


if __name__ == "__main__":
    unittest.main()
