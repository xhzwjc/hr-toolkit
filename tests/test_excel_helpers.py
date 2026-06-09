from __future__ import annotations

import unittest
import shutil
import tempfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from hr_toolkit.common import resources as package_resources
from hr_toolkit.common.excel import _translate_same_row_formula
from hr_toolkit.common.excel_compat import ensure_xlsx_workbook
from hr_toolkit.common.resources import open_template_resource


class ExcelHelperTest(unittest.TestCase):
    def test_translate_same_row_formula_only_rewrites_cell_reference_rows(self) -> None:
        formula = "=A1+B10+SUM(C1:D1)+E$1+$F1+LOG10(100)"

        self.assertEqual(
            _translate_same_row_formula(formula, source_row=1, target_row=12),
            "=A12+B10+SUM(C12:D12)+E$1+$F12+LOG10(100)",
        )

    def test_ensure_xlsx_workbook_converts_xls_to_temp_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            original_xlsx = root / "原始.xlsx"
            xls_file = root / "原始.xls"
            wb = Workbook()
            wb.active["A1"] = "测试"
            wb.save(original_xlsx)
            shutil.copyfile(original_xlsx, xls_file)

            def fake_convert(source: Path, output_path: Path) -> None:
                shutil.copyfile(source, output_path)

            with patch("hr_toolkit.common.excel_compat._convert_xls_to_xlsx", side_effect=fake_convert):
                converted = ensure_xlsx_workbook(xls_file, root / "temp")

            self.assertEqual(converted.suffix, ".xlsx")
            self.assertNotEqual(converted, xls_file)
            loaded = load_workbook(converted, data_only=True)
            self.assertEqual(loaded.active["A1"].value, "测试")
            loaded.close()

    def test_ensure_xlsx_workbook_converts_renamed_xls_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            renamed_xls = root / "改后缀.xlsx"
            renamed_xls.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1fake")

            def fake_convert(source: Path, output_path: Path) -> None:
                wb = Workbook()
                wb.active["A1"] = "已转换"
                wb.save(output_path)

            with patch("hr_toolkit.common.excel_compat._convert_xls_to_xlsx", side_effect=fake_convert):
                converted = ensure_xlsx_workbook(renamed_xls, root / "temp")

            self.assertNotEqual(converted, renamed_xls)
            loaded = load_workbook(converted, data_only=True)
            self.assertEqual(loaded.active["A1"].value, "已转换")
            loaded.close()

    def test_open_template_resource_falls_back_without_files_api(self) -> None:
        with patch.object(package_resources.resources, "files", None):
            with open_template_resource("data_statistics_template.xlsx") as handle:
                self.assertEqual(handle.read(2), b"PK")


if __name__ == "__main__":
    unittest.main()
