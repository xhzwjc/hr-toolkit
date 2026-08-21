"""Tests for cross-platform file dialog path memory and persistence."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from hr_toolkit.gui.app import (
    EXCEL_ARCHIVE_FILE_DIALOG_PATTERN,
    HRToolkitApp,
    _is_excel_or_archive_file,
)
from hr_toolkit.material_preferences import MaterialPreferences


class FileDialogMemoryTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.app = HRToolkitApp.__new__(HRToolkitApp)
        self.app.root = Mock()
        self.app._workspace_width_units = 320
        self.app._workspace_preferred_expanded = True
        self.app.current_project_path = None
        self.app._workspace_recent_projects = []
        self.app._workspace_last_project_path = None
        self.app._last_selected_dir = None

    def test_excel_archive_dialog_accepts_every_supported_archive_suffix(self) -> None:
        for suffix in (
            ".zip",
            ".rar",
            ".7z",
            ".tar",
            ".tar.gz",
            ".tgz",
            ".tar.bz2",
            ".tbz2",
            ".tar.xz",
            ".txz",
        ):
            with self.subTest(suffix=suffix):
                self.assertIn(f"*{suffix}", EXCEL_ARCHIVE_FILE_DIALOG_PATTERN)
                self.assertTrue(_is_excel_or_archive_file(Path(f"资料{suffix.upper()}")))
        self.assertTrue(_is_excel_or_archive_file(Path("资料.xlsx")))
        self.assertFalse(_is_excel_or_archive_file(Path("资料.gz")))

    def test_file_dialog_initial_dir_hierarchy(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            last_dir = temp_path / "last_selected"
            last_dir.mkdir()
            project_dir = temp_path / "project_root"
            project_dir.mkdir()

            # 1. When last_selected_dir is set, it takes top priority
            self.app._last_selected_dir = last_dir
            self.app.current_project_path = project_dir
            self.assertEqual(Path(self.app._file_dialog_initial_dir()).resolve(), last_dir)

            # 2. When last_selected_dir does not exist, fall back to current_project_path
            self.app._last_selected_dir = temp_path / "non_existent_folder"
            self.assertEqual(Path(self.app._file_dialog_initial_dir()).resolve(), project_dir)

            # 3. When neither exists, fall back to desktop or home
            self.app.current_project_path = None
            initial = Path(self.app._file_dialog_initial_dir()).resolve()
            self.assertTrue(initial.is_dir())

    def test_remember_file_dialog_path_with_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            sample_file = temp_path / "data.xlsx"
            sample_file.write_text("test")

            self.app._save_workspace_preferences = Mock()
            self.app._remember_file_dialog_path(str(sample_file))

            self.assertEqual(self.app._last_selected_dir, temp_path)
            self.app._save_workspace_preferences.assert_called_once()

    def test_remember_file_dialog_path_with_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            sub_dir = temp_path / "subfolder"
            sub_dir.mkdir()

            self.app._save_workspace_preferences = Mock()
            self.app._remember_file_dialog_path(str(sub_dir))

            self.assertEqual(self.app._last_selected_dir, sub_dir)
            self.app._save_workspace_preferences.assert_called_once()

    def test_remember_file_dialog_path_with_multiple_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            file1 = temp_path / "file1.xlsx"
            file2 = temp_path / "file2.xlsx"
            file1.write_text("1")
            file2.write_text("2")

            self.app._save_workspace_preferences = Mock()
            self.app._remember_file_dialog_path([str(file1), str(file2)])

            self.assertEqual(self.app._last_selected_dir, temp_path)
            self.app._save_workspace_preferences.assert_called_once()

    def test_remember_file_dialog_path_ignores_empty(self) -> None:
        self.app._save_workspace_preferences = Mock()
        self.app._remember_file_dialog_path("")
        self.app._remember_file_dialog_path([])
        self.app._remember_file_dialog_path(None)
        self.assertIsNone(self.app._last_selected_dir)
        self.app._save_workspace_preferences.assert_not_called()

    def test_preferences_save_and_load_persistence(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            custom_settings = temp_path / "workspace-ui.json"
            chosen_dir = temp_path / "my_work"
            chosen_dir.mkdir()

            with patch.object(self.app, "_workspace_settings_path", return_value=custom_settings):
                self.app._last_selected_dir = chosen_dir
                self.app._material_preferences = MaterialPreferences(
                    custom_materials=["户口本"]
                )
                self.app._material_preferences.save_preset(
                    "补充入职",
                    ["身份证", "户口本"],
                )
                self.app._save_workspace_preferences()

                self.assertTrue(custom_settings.is_file())
                saved_data = json.loads(custom_settings.read_text(encoding="utf-8"))
                self.assertEqual(saved_data.get("last_selected_dir"), str(chosen_dir))
                self.assertEqual(
                    saved_data["material_preferences"]["custom_materials"],
                    ["户口本"],
                )

                # Now simulate restart / new app loading preferences
                new_app = HRToolkitApp.__new__(HRToolkitApp)
                with patch.object(new_app, "_workspace_settings_path", return_value=custom_settings):
                    new_app._load_workspace_preferences()
                    self.assertEqual(new_app._last_selected_dir, chosen_dir)
                    self.assertEqual(
                        new_app._material_preferences.get_preset("补充入职"),
                        ("身份证", "户口本"),
                    )

    def test_preferences_load_ignores_missing_dir(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            custom_settings = temp_path / "workspace-ui.json"
            deleted_dir = temp_path / "deleted_work"
            custom_settings.write_text(json.dumps({"last_selected_dir": str(deleted_dir)}), encoding="utf-8")

            with patch.object(self.app, "_workspace_settings_path", return_value=custom_settings):
                self.app._load_workspace_preferences()
                self.assertIsNone(self.app._last_selected_dir)

    def test_askopenfilename_wraps_initialdir_and_remembers_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            sample_file = temp_path / "salary.xlsx"
            sample_file.write_text("data")

            with patch("hr_toolkit.gui.app.filedialog.askopenfilename", return_value=str(sample_file)) as mock_dialog:
                with patch.object(self.app, "_save_workspace_preferences"):
                    result = self.app._askopenfilename(title="选择工资表")

                    self.assertEqual(result, str(sample_file))
                    mock_dialog.assert_called_once()
                    self.assertIn("initialdir", mock_dialog.call_args.kwargs)
                    self.assertEqual(self.app._last_selected_dir, temp_path)

    def test_askopenfilenames_wraps_initialdir_and_remembers_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            sample_file = temp_path / "file1.xlsx"
            sample_file.write_text("data")

            with patch("hr_toolkit.gui.app.filedialog.askopenfilenames", return_value=(str(sample_file),)) as mock_dialog:
                with patch.object(self.app, "_save_workspace_preferences"):
                    result = self.app._askopenfilenames(title="选择多个文件")

                    self.assertEqual(result, (str(sample_file),))
                    mock_dialog.assert_called_once()
                    self.assertIn("initialdir", mock_dialog.call_args.kwargs)
                    self.assertEqual(self.app._last_selected_dir, temp_path)

    def test_set_change_input_paths_remembers_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            file1 = temp_path / "sample.xlsx"
            file1.write_text("x")

            self.app.change_input_paths = None
            self.app.input_path = Mock()
            self.app.output_dir_user_selected = True
            self.app._refresh_upload_card = Mock()
            self.app._save_workspace_preferences = Mock()

            self.app._set_change_input_paths([file1])

            self.assertEqual(self.app._last_selected_dir, temp_path)
            self.app._save_workspace_preferences.assert_called_once()

    def test_askdirectory_wraps_initialdir_and_remembers_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            sample_dir = temp_path / "payroll_folder"
            sample_dir.mkdir()

            with patch("hr_toolkit.gui.app.filedialog.askdirectory", return_value=str(sample_dir)) as mock_dialog:
                with patch.object(self.app, "_save_workspace_preferences"):
                    result = self.app._askdirectory(title="选择工资表文件夹")

                    self.assertEqual(result, str(sample_dir))
                    mock_dialog.assert_called_once()
                    self.assertIn("initialdir", mock_dialog.call_args.kwargs)
                    self.assertEqual(self.app._last_selected_dir, sample_dir)

    def test_tool_choose_methods_use_wrapped_filedialog(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir).resolve()
            file1 = temp_path / "salary.xlsx"
            file1.write_text("salary")

            self.app.change_input_paths = None
            self.app.input_path = Mock()
            self.app.output_dir_user_selected = True
            self.app._refresh_upload_card = Mock()
            self.app._save_workspace_preferences = Mock()

            with patch.object(self.app, "_askopenfilenames", return_value=(str(file1),)) as mock_ask:
                self.app._choose_salary_files_or_zip()
                mock_ask.assert_called_once()
                self.assertEqual(self.app._last_selected_dir, temp_path)


if __name__ == "__main__":
    unittest.main()
