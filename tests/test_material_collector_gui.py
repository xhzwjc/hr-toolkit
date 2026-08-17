from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from hr_toolkit.gui import (
    HRToolkitApp,
    TOOL_NAV_LABELS,
    NAV_GROUPS,
    WORKSPACE_TOOL_PATHS,
)
from hr_toolkit.tools.material_collector import (
    MODE_BY_EMPLOYEE,
    MODE_BY_MATERIAL,
    MODE_FLAT,
    MODE_LABELS,
)


class MaterialCollectorGUITest(unittest.TestCase):

    def test_navigation_and_workspace_registration(self) -> None:
        self.assertIn("material_collector", TOOL_NAV_LABELS)
        self.assertEqual(TOOL_NAV_LABELS["material_collector"], "员工资料打包")

        # Check in NAV_GROUPS
        hr_group_tools = next(tools for group, tools in NAV_GROUPS if group == "人员与档案")
        self.assertIn("material_collector", hr_group_tools)

        # Check in WORKSPACE_TOOL_PATHS
        self.assertIn("material_collector", WORKSPACE_TOOL_PATHS)
        self.assertEqual(WORKSPACE_TOOL_PATHS["material_collector"], ("人员与档案", "员工资料打包"))

    def test_mode_labels_coverage(self) -> None:
        self.assertEqual(MODE_LABELS["按员工归类（每人一个文件夹）"], MODE_BY_EMPLOYEE)
        self.assertEqual(MODE_LABELS["按材料归类（每类材料一个文件夹）"], MODE_BY_MATERIAL)
        self.assertEqual(MODE_LABELS["平铺输出（所有文件在同一文件夹）"], MODE_FLAT)


if __name__ == "__main__":
    unittest.main()
