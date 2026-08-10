from __future__ import annotations

import tempfile
import unittest
import warnings as warnings_module
import zipfile
from pathlib import Path
from unittest.mock import patch

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
                archive.writestr("部门/../overwrite.xlsx", _make_xlsx_bytes())
                archive.writestr("正常.xlsx", _make_xlsx_bytes())
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual([path.name for path in files], ["正常.xlsx"])
            self.assertEqual(len(warnings), 2)
            self.assertTrue(all("不安全路径" in warning for warning in warnings))
            self.assertFalse((temp_dir / "escape.xlsx").exists())
            self.assertFalse(any(temp_dir.rglob("overwrite.xlsx")))

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

    def test_rejects_zip_with_too_many_members_before_extracting(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "many.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("1.xlsx", b"one")
                archive.writestr("2.xlsx", b"two")
            warnings: list[str] = []
            with patch("hr_toolkit.common.inputs.ZIP_MAX_MEMBERS", 1):
                files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual(files, [])
            self.assertIn("文件数量超过", warnings[0])
            self.assertFalse(any(temp_dir.glob("zip_*")))

    def test_rejects_zip_whose_expanded_size_exceeds_limit(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "large.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("large.xlsx", b"x" * 128)
            warnings: list[str] = []
            with patch("hr_toolkit.common.inputs.ZIP_MAX_TOTAL_BYTES", 64):
                files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual(files, [])
            self.assertIn("总大小超过", warnings[0])

    def test_rejects_zip_when_temporary_disk_space_is_insufficient(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "space.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("data.xlsx", b"data")
            disk_usage = type("DiskUsage", (), {"total": 1, "used": 1, "free": 0})()
            warnings: list[str] = []
            with patch("hr_toolkit.common.inputs.shutil.disk_usage", return_value=disk_usage):
                files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual(files, [])
            self.assertIn("临时磁盘空间不足", warnings[0])

    def test_rejects_abnormal_compression_ratio(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "ratio.zip"
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("compressed.xlsx", b"x" * 4096)
            warnings: list[str] = []
            with (
                patch("hr_toolkit.common.inputs.ZIP_RATIO_CHECK_MIN_BYTES", 1),
                patch("hr_toolkit.common.inputs.ZIP_MAX_COMPRESSION_RATIO", 2),
            ):
                files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual(files, [])
            self.assertIn("压缩比例异常", warnings[0])

    def test_skips_symbolic_link_members(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "link.zip"
            link = zipfile.ZipInfo("outside.xlsx")
            link.create_system = 3
            link.external_attr = 0o120777 << 16
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr(link, "../outside.xlsx")
                archive.writestr("normal.xlsx", _make_xlsx_bytes())
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual([path.name for path in files], ["normal.xlsx"])
            self.assertTrue(any("链接文件" in warning for warning in warnings))

    def test_rejects_duplicate_member_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "duplicate.zip"
            with warnings_module.catch_warnings():
                warnings_module.simplefilter("ignore", UserWarning)
                with zipfile.ZipFile(zip_path, "w") as archive:
                    archive.writestr("工资表.xlsx", b"first")
                    archive.writestr("工资表.xlsx", b"second")
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual(files, [])
            self.assertIn("重复路径", warnings[0])
            self.assertFalse(any(temp_dir.glob("zip_*")))

    def test_rejects_duplicate_paths_with_dot_alias(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            temp_dir = Path(temp_root)
            zip_path = temp_dir / "duplicate-alias.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("部门/工资表.xlsx", b"first")
                archive.writestr("部门/./工资表.xlsx", b"second")
            warnings: list[str] = []
            files = extract_zip_excel_files(zip_path, temp_dir, warnings)
            self.assertEqual(files, [])
            self.assertIn("重复路径", warnings[0])


if __name__ == "__main__":
    unittest.main()
