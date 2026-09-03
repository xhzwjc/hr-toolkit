from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from hr_toolkit.desktop_helpers import (
    default_workspace_project_name,
    workspace_project_create_error_message,
    workspace_project_creation_target,
    workspace_project_name_error,
)
from hr_toolkit.project_store import ProjectStoreError, validate_project_name


class _Value:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class ProjectCreationValidationTests(unittest.TestCase):
    @staticmethod
    def _backend_name_error(value: str) -> str | None:
        try:
            validate_project_name(value)
        except ProjectStoreError as exc:
            return str(exc)
        return None

    def test_default_name_and_current_gui_name_rules(self) -> None:
        self.assertEqual(
            default_workspace_project_name(date(2026, 8, 10)),
            "2026年8月人事月度工作",
        )
        self.assertIsNone(workspace_project_name_error("  华东人事项目  "))
        self.assertEqual(workspace_project_name_error("  "), "项目名称不能为空。")
        for value in (".", "..", "工资/社保", "工资:社保", "工资*社保"):
            with self.subTest(value=value):
                self.assertEqual(
                    workspace_project_name_error(value),
                    self._backend_name_error(value),
                )

    def test_gui_name_validation_matches_backend_for_all_portability_rules(self) -> None:
        invalid_names = [
            "",
            "   ",
            ".",
            "..",
            "项目.",
            "项目.   ",
            "人事\n项目",
            "人事\x00项目",
            "工资/社保",
            "工资\\社保",
            "工资:社保",
            "工资*社保",
            "工资?社保",
            '工资"社保',
            "工资<社保",
            "工资>社保",
            "工资|社保",
            ".hrtoolkit",
            ".HRTOOLKIT",
            "项目" * 61,
            "A" * 121,
            "CON",
            "con.txt",
            "PrN.xlsx",
            "AUX",
            "NUL.log",
            *(f"COM{index}" for index in range(1, 10)),
            *(f"lpt{index}.csv" for index in range(1, 10)),
        ]
        for value in invalid_names:
            with self.subTest(value=repr(value)):
                backend_error = self._backend_name_error(value)
                self.assertIsNotNone(backend_error)
                self.assertEqual(workspace_project_name_error(value), backend_error)

        valid_names = (
            "  华东人事项目  ",
            "月度项目   ",
            "COM0",
            "LPT10",
            ".hrtoolkit备份",
            "A" * 120,
        )
        for value in valid_names:
            with self.subTest(value=repr(value)):
                self.assertIsNone(self._backend_name_error(value))
                self.assertIsNone(workspace_project_name_error(value))

        self.assertEqual(validate_project_name("月度项目   "), "月度项目")

    def test_target_preview_allows_new_or_empty_folder_but_never_overwrites_content(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            parent = Path(temp_root)
            target, error = workspace_project_creation_target(str(parent), "八月人事")
            self.assertEqual(target, parent / "八月人事")
            self.assertIsNone(error)

            target.mkdir()
            _target, error = workspace_project_creation_target(str(parent), "八月人事")
            self.assertIsNone(error)

            (target / "已有资料.xlsx").write_bytes(b"keep")
            _target, error = workspace_project_creation_target(str(parent), "八月人事")
            self.assertIn("已有同名文件夹", error or "")
            self.assertEqual((target / "已有资料.xlsx").read_bytes(), b"keep")

    def test_creation_error_messages_are_hr_readable(self) -> None:
        self.assertIn("没有权限", workspace_project_create_error_message(PermissionError()))
        self.assertIn("不可用", workspace_project_create_error_message(FileNotFoundError()))
        self.assertIn("磁盘空间不足", workspace_project_create_error_message(OSError(28, "full")))
        self.assertEqual(
            workspace_project_create_error_message(RuntimeError("共享盘不能作为活动项目位置")),
            "共享盘不能作为活动项目位置",
        )


if __name__ == "__main__":
    unittest.main()
