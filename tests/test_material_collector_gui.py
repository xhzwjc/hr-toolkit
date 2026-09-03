from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from hr_toolkit.gui_qt.controller import NAV_GROUPS
from hr_toolkit.cli import build_parser
from hr_toolkit.tools.material_collector import (
    LIBRARY_MODE_FLAT_OCR,
    LIBRARY_MODE_LABELS,
    LIBRARY_MODE_PERSON_FOLDER,
    MODE_BY_EMPLOYEE,
    MODE_BY_MATERIAL,
    MODE_FLAT,
    MODE_LABELS,
)


class MaterialCollectorGUITest(unittest.TestCase):

    def test_navigation_and_workspace_registration(self) -> None:
        hr_group_tools = dict(next(tools for group, tools in NAV_GROUPS if group == "人员与档案"))
        self.assertIn("material_collector", hr_group_tools)
        self.assertEqual(hr_group_tools["material_collector"], "员工资料打包")

    def test_mode_labels_coverage(self) -> None:
        self.assertEqual(MODE_LABELS["按员工归类（每人一个文件夹）"], MODE_BY_EMPLOYEE)
        self.assertEqual(MODE_LABELS["按材料归类（每类材料一个文件夹）"], MODE_BY_MATERIAL)
        self.assertEqual(MODE_LABELS["平铺输出（所有文件在同一文件夹）"], MODE_FLAT)

    def test_library_mode_is_independent_from_existing_output_modes(self) -> None:
        self.assertEqual(
            LIBRARY_MODE_LABELS["按人员文件夹查找（原模式）"],
            LIBRARY_MODE_PERSON_FOLDER,
        )
        self.assertEqual(
            LIBRARY_MODE_LABELS["无序平铺资料库（OCR 索引）"],
            LIBRARY_MODE_FLAT_OCR,
        )
        self.assertEqual(set(MODE_LABELS.values()), {MODE_BY_EMPLOYEE, MODE_BY_MATERIAL, MODE_FLAT})

    def test_cli_defaults_to_old_library_mode_and_accepts_flat_ocr(self) -> None:
        parser = build_parser()
        default_args = parser.parse_args([
            "material-collector", "-l", "/tmp/lib", "-r", "张三", "-o", "/tmp/out",
        ])
        self.assertEqual(default_args.library_mode, LIBRARY_MODE_PERSON_FOLDER)

        flat_args = parser.parse_args([
            "material-collector", "-l", "/tmp/lib", "-r", "张三", "-o", "/tmp/out",
            "--library-mode", LIBRARY_MODE_FLAT_OCR,
            "--mode", MODE_BY_MATERIAL,
        ])
        self.assertEqual(flat_args.library_mode, LIBRARY_MODE_FLAT_OCR)
        self.assertEqual(flat_args.mode, MODE_BY_MATERIAL)


if __name__ == "__main__":
    unittest.main()
