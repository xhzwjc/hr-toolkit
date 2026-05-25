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
            self.assertTrue((out_dir / "_salary_split_manifest.json").exists())

            companies = {item["company"]: item for item in payload["outputs"]}
            self.assertEqual(companies["春苗北京"]["employee_count"], 18)
            self.assertEqual(companies["唐人"]["employee_count"], 4)
            self.assertEqual(companies["岩亨"]["employee_count"], 1)

            manifest = json.loads((out_dir / "_salary_split_manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["tool_name"], "需求4-工资表按入职公司拆分")

            wb = load_workbook(companies["春苗北京"]["file_path"], data_only=False)
            detail = wb["明细表"]
            self.assertEqual(detail["A24"].value, "春苗北京合计")
            self.assertEqual(detail["B6"].value, "员工1")
            self.assertEqual(detail["B23"].value, "员工23")
            self.assertEqual(detail["AU6"].value, "春苗北京")
            self.assertEqual(detail["AS6"].value, "未填写项目")


if __name__ == "__main__":
    unittest.main()
