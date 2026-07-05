from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from hr_toolkit.tools.folder_rename import (
    MODE_APPEND,
    MODE_REMOVE,
    MODE_REPLACE,
    rename_person_folders,
)


class FolderRenameTest(unittest.TestCase):
    def test_append_suffix_to_all_folders(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "张三").mkdir()
            (root / "李四").mkdir()
            (root / "说明.txt").write_text("ignore", encoding="utf-8")

            # “-劳动合同”带前缀追加;右侧注释验证“劳动合同”不带前缀时直传
            preview = rename_person_folders(root, mode=MODE_APPEND, text="-劳动合同", dry_run=True)
            self.assertEqual(preview.operation_count, 2)
            self.assertTrue((root / "张三").exists())

            result = rename_person_folders(root, mode=MODE_APPEND, text="-劳动合同")

            self.assertEqual(result.operation_count, 2)
            self.assertTrue((root / "张三-劳动合同").exists())
            self.assertTrue((root / "李四-劳动合同").exists())
            self.assertTrue((root / "说明.txt").exists())

    def test_append_text_passed_through_unchanged(self) -> None:
        """bug4 修复:用户输入什么就追加什么，不自动加分隔符"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "张三").mkdir()

            result = rename_person_folders(root, mode=MODE_APPEND, text="劳动合同")

            self.assertEqual(result.operation_count, 1)
            self.assertTrue((root / "张三劳动合同").exists())
            self.assertFalse((root / "张三-劳动合同").exists())

    def test_append_suffix_to_one_person(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "张三").mkdir()
            (root / "李四").mkdir()

            result = rename_person_folders(root, mode=MODE_APPEND, text="-身份证", target_name="张三")

            self.assertEqual(result.operation_count, 1)
            self.assertTrue((root / "张三-身份证").exists())
            self.assertTrue((root / "李四").exists())

    def test_remove_suffix_variants_from_all_folders(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "张三-劳动合同").mkdir()
            (root / "李四_劳动合同").mkdir()
            (root / "赵露思劳动合同").mkdir()
            (root / "王五-身份证").mkdir()

            result = rename_person_folders(root, mode=MODE_REMOVE, text="_劳动合同")

            self.assertEqual(result.operation_count, 3)
            self.assertTrue((root / "张三").exists())
            self.assertTrue((root / "李四").exists())
            self.assertTrue((root / "赵露思").exists())
            self.assertTrue((root / "王五-身份证").exists())

    def test_remove_suffix_from_one_person(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "张三_身份证").mkdir()
            (root / "李四_身份证").mkdir()

            result = rename_person_folders(root, mode=MODE_REMOVE, text="身份证", target_name="张三")

            self.assertEqual(result.operation_count, 1)
            self.assertTrue((root / "张三").exists())
            self.assertTrue((root / "李四_身份证").exists())

    def test_replace_one_folder_name(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "张三").mkdir()
            (root / "李四").mkdir()

            result = rename_person_folders(
                root,
                mode=MODE_REPLACE,
                target_name="张三",
                replacement_name="章五",
            )

            self.assertEqual(result.operation_count, 1)
            self.assertFalse((root / "张三").exists())
            self.assertTrue((root / "章五").exists())
            self.assertTrue((root / "李四").exists())

    def test_skip_existing_target(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "张三").mkdir()
            (root / "张三-劳动合同").mkdir()

            result = rename_person_folders(root, mode=MODE_APPEND, text="-劳动合同")

            self.assertEqual(result.operation_count, 0)
            self.assertTrue(any("已存在" in warning or "已包含后缀" in warning for warning in result.warnings))


if __name__ == "__main__":
    unittest.main()
