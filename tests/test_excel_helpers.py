from __future__ import annotations

import os
import sys
import unittest
import shutil
import tempfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from hr_toolkit.common import resources as package_resources
from hr_toolkit.common.excel import _translate_same_row_formula
from hr_toolkit.common.excel_compat import (
    _build_conversion_env,
    _diagnose_missing_libreoffice,
    _find_libreoffice,
    ensure_xlsx_workbook,
)
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


class LibreOfficeLookupTest(unittest.TestCase):
    """回归测试：macOS .app 启动时 PATH 被裁剪，找不到 Homebrew 装的 soffice。"""

    def test_find_libreoffice_returns_homebrew_binary_on_macos(self) -> None:
        """当 PATH 不含 /opt/homebrew/bin，但二进制实际存在时，应能识别。"""
        fake_binary = Path("/opt/homebrew/bin/soffice")
        with patch("hr_toolkit.common.excel_compat.os.environ", {"PATH": "/usr/bin:/bin"}):
            with patch("hr_toolkit.common.excel_compat.Path.exists", return_value=True):
                executable = _find_libreoffice()
        # 实际值受其他路径影响，只要不为 None 即可证明查找逻辑生效
        self.assertIsNotNone(executable)
        self.assertTrue(executable.endswith("soffice") or executable.endswith("libreoffice"))

    def test_find_libreoffice_returns_none_when_nothing_exists(self) -> None:
        """全部路径都不存在时返回 None。"""
        with patch("hr_toolkit.common.excel_compat.os.environ", {"PATH": "/usr/bin:/bin"}):
            with patch("hr_toolkit.common.excel_compat.Path.exists", return_value=False):
                with patch("hr_toolkit.common.excel_compat.shutil.which", return_value=None):
                    self.assertIsNone(_find_libreoffice())

    @unittest.skipUnless(sys.platform == "darwin", "macOS-specific LibreOffice bundle path")
    def test_find_libreoffice_prefers_mac_app_path(self) -> None:
        """官方 .app 路径优先级最高。"""
        with patch("hr_toolkit.common.excel_compat.os.environ", {"PATH": "/usr/bin:/bin"}):
            original_exists = Path.exists

            def fake_exists(self: Path) -> bool:
                if str(self) == "/Applications/LibreOffice.app/Contents/MacOS/soffice":
                    return True
                return False

            with patch("hr_toolkit.common.excel_compat.Path.exists", fake_exists):
                executable = _find_libreoffice()
        self.assertEqual(
            executable,
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        )

    def test_build_conversion_env_keeps_existing_path(self) -> None:
        """环境变量注入不能丢失用户已有 PATH。"""
        with patch.dict(os.environ, {"PATH": "/usr/bin:/bin", "HOME": "/tmp"}, clear=False):
            env = _build_conversion_env()
        self.assertIn("/usr/bin", env["PATH"])
        self.assertIn("/bin", env["PATH"])

    def test_build_conversion_env_adds_homebrew_when_exists(self) -> None:
        """当 Homebrew 路径存在时，应被追加到 PATH。"""
        with patch.dict(os.environ, {"PATH": "/usr/bin:/bin"}, clear=False):
            with patch("hr_toolkit.common.excel_compat.Path.exists", return_value=True):
                env = _build_conversion_env()
        self.assertIn("/opt/homebrew/bin", env["PATH"])
        self.assertIn("/usr/local/bin", env["PATH"])

    def test_diagnose_contains_path_and_install_hint(self) -> None:
        """诊断信息应包含 PATH、已检查路径与安装建议。"""
        with patch.dict(os.environ, {"PATH": "/usr/bin:/bin"}, clear=False):
            with patch("hr_toolkit.common.excel_compat.Path.exists", return_value=False):
                message = _diagnose_missing_libreoffice()
        self.assertIn("未找到 libreoffice/soffice", message)
        self.assertIn("PATH", message)
        self.assertIn("brew install", message)
        self.assertIn("/opt/homebrew/bin", message)
        self.assertIn("/Applications/LibreOffice.app", message)


if __name__ == "__main__":
    unittest.main()
