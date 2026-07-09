from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import Workbook

from hr_toolkit.common.inputs import extract_zip_excel_files, normalize_input_paths


def _make_xlsx_bytes() -> bytes:
    workbook = Workbook()
    workbook.active["A1"] = "测试"
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as handle:
        temp_path = Path(handle.name)
    try:
        workbook.save(temp_path)
        return temp_path.read_bytes()
    finally:
        temp_path.unlink(missing_ok=True)


class NormalizeInputPathsTest(unittest.TestCase):
    def test_single_path_is_wrapped(self) -> None:
        paths = normalize_input_paths("some.xlsx", "请选择文件。")
        self.assertEqual(len(paths), 1)
        self.assertTrue(paths[0].is_absolute())

    def test_empty_list_raises_with_message(self) -> None:
        with self.assertRaises(ValueError) as context:
            normalize_input_paths([], "请选择文件。")
        self.assertEqual(str(context.exception), "请选择文件。")


class ExtractZipExcelFilesTest(unittest.TestCase):
    def test_extracts_excel_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "input.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("目录/工资表.xlsx", _make_xlsx_bytes())
                archive.writestr("说明.txt", "忽略")
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual([path.name for path in files], ["工资表.xlsx"])
            self.assertEqual(warnings, [])

    def test_restores_gbk_encoded_filenames(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "gbk.zip"
            # Windows 资源管理器/WinRAR 用 GBK 存储文件名且不设置 UTF-8 标志。
            # zipfile 写入时无法直接生成这种 zip，因此先用等长 ASCII 名占位，
            # 再把归档里的文件名字节替换成 GBK 字节。
            gbk_name_bytes = "春苗5月考勤.xlsx".encode("gbk")
            placeholder = b"A" * (len(gbk_name_bytes) - 5) + b".xlsx"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr(placeholder.decode("ascii"), _make_xlsx_bytes())
            zip_path.write_bytes(zip_path.read_bytes().replace(placeholder, gbk_name_bytes))
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual([path.name for path in files], ["春苗5月考勤.xlsx"])
            self.assertEqual(warnings, [])

    def test_skips_path_traversal_members(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "evil.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("../escape.xlsx", _make_xlsx_bytes())
                archive.writestr("正常.xlsx", _make_xlsx_bytes())
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual([path.name for path in files], ["正常.xlsx"])
            self.assertEqual(len(warnings), 1)
            self.assertIn("不安全路径", warnings[0])
            self.assertFalse((temp_dir / "escape.xlsx").exists())

    def test_subdir_places_files_under_named_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "social.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("明细.xlsx", _make_xlsx_bytes())
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings, subdir="social")
            self.assertEqual(len(files), 1)
            self.assertEqual(files[0].parent.name, "social")

    def test_bad_zip_returns_empty_with_warning(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "broken.zip"
            zip_path.write_bytes(b"not a zip")
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual(files, [])
            self.assertEqual(len(warnings), 1)
            self.assertIn("解压失败", warnings[0])


if __name__ == "__main__":
    unittest.main()
