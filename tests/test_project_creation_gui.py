from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from hr_toolkit.gui import (
    HRToolkitApp,
    _default_workspace_project_name,
    _workspace_project_create_error_message,
    _workspace_project_creation_target,
    _workspace_project_name_error,
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
            _default_workspace_project_name(date(2026, 8, 10)),
            "2026年8月人事月度工作",
        )
        self.assertIsNone(_workspace_project_name_error("  华东人事项目  "))
        self.assertEqual(_workspace_project_name_error("  "), "项目名称不能为空。")
        for value in (".", "..", "工资/社保", "工资:社保", "工资*社保"):
            with self.subTest(value=value):
                self.assertEqual(
                    _workspace_project_name_error(value),
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
                self.assertEqual(_workspace_project_name_error(value), backend_error)

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
                self.assertIsNone(_workspace_project_name_error(value))

        self.assertEqual(validate_project_name("月度项目   "), "月度项目")

    def test_target_preview_allows_new_or_empty_folder_but_never_overwrites_content(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            parent = Path(temp_root)
            target, error = _workspace_project_creation_target(str(parent), "八月人事")
            self.assertEqual(target, parent / "八月人事")
            self.assertIsNone(error)

            target.mkdir()
            _target, error = _workspace_project_creation_target(str(parent), "八月人事")
            self.assertIsNone(error)

            (target / "已有资料.xlsx").write_bytes(b"keep")
            _target, error = _workspace_project_creation_target(str(parent), "八月人事")
            self.assertIn("已有同名文件夹", error or "")
            self.assertEqual((target / "已有资料.xlsx").read_bytes(), b"keep")

    def test_creation_error_messages_are_hr_readable(self) -> None:
        self.assertIn("没有权限", _workspace_project_create_error_message(PermissionError()))
        self.assertIn("不可用", _workspace_project_create_error_message(FileNotFoundError()))
        self.assertIn("磁盘空间不足", _workspace_project_create_error_message(OSError(28, "full")))
        self.assertEqual(
            _workspace_project_create_error_message(RuntimeError("共享盘不能作为活动项目位置")),
            "共享盘不能作为活动项目位置",
        )


class ProjectCreationControllerTests(unittest.TestCase):
    def _app_for_submit(self, parent: Path, project_name: str, creator: Mock):
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._project_create_busy = False
        app._project_create_name_var = _Value(project_name)
        app._project_create_parent_var = _Value(str(parent))
        app._project_create_preview_var = _Value()
        app._project_create_status_var = _Value()
        app._project_create_name_entry = Mock()
        app._project_create_location_button = Mock()
        app._project_create_cancel_button = Mock()
        app._project_create_submit_button = Mock()
        app._project_create_status_label = Mock()
        app._project_create_window = Mock()
        app._project_store_class = SimpleNamespace(create=creator)
        app._set_workspace_project = Mock()
        app._close_workspace_project_create_dialog = Mock()
        return app

    def test_entry_opens_custom_dialog_without_opening_system_picker(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.root = object()
        app._project_store_class = object()
        app._project_change_is_blocked = Mock(return_value=False)
        app._open_workspace_project_create_dialog = Mock()

        with patch("hr_toolkit.gui.filedialog.askdirectory") as chooser:
            app._create_workspace_project()

        chooser.assert_not_called()
        app._open_workspace_project_create_dialog.assert_called_once_with()

    def test_existing_dialog_is_focused_instead_of_duplicated(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        window = Mock()
        window.winfo_exists.return_value = True
        app._project_create_window = window

        with patch("hr_toolkit.gui.Toplevel") as create_window:
            app._open_workspace_project_create_dialog()

        create_window.assert_not_called()
        window.lift.assert_called_once_with()
        window.focus_force.assert_called_once_with()

    def test_only_location_action_opens_picker_and_cancel_keeps_previous_value(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            app = HRToolkitApp.__new__(HRToolkitApp)
            app._project_create_busy = False
            app._project_create_window = Mock()
            app._project_create_parent_var = _Value(temp_root)
            app._default_workspace_project_parent = Mock(return_value=Path(temp_root))

            with patch("hr_toolkit.gui.filedialog.askdirectory", return_value="") as chooser:
                app._choose_workspace_project_parent()

            self.assertEqual(app._project_create_parent_var.get(), temp_root)
            chooser.assert_called_once()
            self.assertIs(chooser.call_args.kwargs["parent"], app._project_create_window)

    def test_live_refresh_updates_preview_and_submit_state(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            app = HRToolkitApp.__new__(HRToolkitApp)
            app._project_create_busy = False
            app._project_create_name_var = _Value("八月人事")
            app._project_create_parent_var = _Value(temp_root)
            app._project_create_preview_var = _Value()
            app._project_create_status_var = _Value()
            app._project_create_status_label = Mock()
            app._project_create_submit_button = Mock()

            app._refresh_workspace_project_create_form()
            self.assertEqual(app._project_create_preview_var.get(), str(Path(temp_root) / "八月人事"))
            app._project_create_submit_button.configure.assert_called_with(state="normal")

            app._project_create_name_var.set("")
            app._refresh_workspace_project_create_form()
            self.assertEqual(app._project_create_status_var.get(), "项目名称不能为空。")
            app._project_create_submit_button.configure.assert_called_with(state="disabled")

    def test_submit_creates_once_and_reuses_set_workspace_project(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            parent = Path(temp_root)
            project_root = parent / "八月人事"
            project = SimpleNamespace(
                workspace=SimpleNamespace(
                    name="八月人事",
                    root=project_root,
                    writable=True,
                )
            )
            creator = Mock(return_value=project)
            app = self._app_for_submit(parent, "  八月人事  ", creator)

            app._submit_workspace_project_create()

            creator.assert_called_once_with(project_root, "八月人事")
            app._set_workspace_project.assert_called_once_with(
                "八月人事",
                project_root,
                read_only=False,
                store=project,
            )
            app._close_workspace_project_create_dialog.assert_called_once_with(force=True)

    def test_busy_state_rejects_duplicate_submit(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            creator = Mock()
            app = self._app_for_submit(Path(temp_root), "八月人事", creator)
            app._project_create_busy = True

            app._submit_workspace_project_create()

            creator.assert_not_called()

    def test_failure_keeps_values_and_allows_retry(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            parent = Path(temp_root)
            creator = Mock(side_effect=PermissionError("denied"))
            app = self._app_for_submit(parent, "八月人事", creator)

            with patch("hr_toolkit.gui.runlog.log_exception"):
                app._submit_workspace_project_create()

            self.assertFalse(app._project_create_busy)
            self.assertEqual(app._project_create_name_var.get(), "八月人事")
            self.assertEqual(app._project_create_parent_var.get(), str(parent))
            self.assertIn("没有权限", app._project_create_status_var.get())
            app._close_workspace_project_create_dialog.assert_not_called()
            self.assertEqual(app._project_create_name_entry.configure.call_args_list[-1].kwargs["state"], "normal")
            self.assertEqual(app._project_create_submit_button.configure.call_args_list[-1].kwargs["state"], "normal")


if __name__ == "__main__":
    unittest.main()
