from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from hr_toolkit.gui.app import HRToolkitApp
from hr_toolkit.material_preferences import (
    BUILTIN_MATERIALS,
    BUILTIN_MATERIAL_PRESETS,
    MaterialPreferences,
)


class _Value:
    def __init__(self, value=None) -> None:
        self.value = value

    def get(self):
        return self.value

    def set(self, value) -> None:
        self.value = value


class MaterialPreferencesTest(unittest.TestCase):
    def test_builtin_materials_and_presets(self) -> None:
        preferences = MaterialPreferences()

        self.assertEqual(preferences.available_materials, BUILTIN_MATERIALS)
        self.assertEqual(
            preferences.get_preset("入职材料"),
            ("身份证", "劳动合同"),
        )
        self.assertEqual(
            preferences.get_preset("证书材料"),
            ("特种证书", "资格证书"),
        )
        self.assertEqual(
            preferences.preset_names,
            tuple(BUILTIN_MATERIAL_PRESETS),
        )

    def test_custom_material_and_preset_crud(self) -> None:
        preferences = MaterialPreferences()
        preferences.add_material("户口本")
        preferences.save_preset("补充入职", ["身份证", "户口本"])
        self.assertEqual(
            preferences.get_preset("补充入职"),
            ("身份证", "户口本"),
        )

        preferences.save_preset(
            "补充入职",
            ["劳动合同", "户口本"],
            replacing="补充入职",
        )
        self.assertEqual(
            preferences.get_preset("补充入职"),
            ("劳动合同", "户口本"),
        )

        renamed = preferences.rename_preset("补充入职", "完整入职")
        self.assertEqual(renamed, "完整入职")
        self.assertIsNone(preferences.get_preset("补充入职"))
        self.assertEqual(
            preferences.get_preset("完整入职"),
            ("劳动合同", "户口本"),
        )

        preferences.delete_preset("完整入职")
        self.assertIsNone(preferences.get_preset("完整入职"))
        self.assertIn("户口本", preferences.available_materials)

    def test_delete_referenced_material_updates_and_removes_presets(self) -> None:
        preferences = MaterialPreferences(custom_materials=["户口本", "体检报告"])
        preferences.save_preset("补充材料", ["身份证", "户口本"])
        preferences.save_preset("只有户口本", ["户口本"])
        preferences.save_preset("体检材料", ["体检报告"])

        result = preferences.remove_material("户口本")

        self.assertEqual(result.updated_presets, ("补充材料",))
        self.assertEqual(result.removed_presets, ("只有户口本",))
        self.assertEqual(preferences.get_preset("补充材料"), ("身份证",))
        self.assertIsNone(preferences.get_preset("只有户口本"))
        self.assertEqual(preferences.get_preset("体检材料"), ("体检报告",))

    def test_roundtrip_and_stale_reference_cleanup(self) -> None:
        preferences = MaterialPreferences(custom_materials=["户口本"])
        preferences.save_preset("补充材料", ["身份证", "户口本"])

        restored = MaterialPreferences.from_payload(preferences.to_payload())
        self.assertEqual(restored.custom_materials, ("户口本",))
        self.assertEqual(
            restored.get_preset("补充材料"),
            ("身份证", "户口本"),
        )

        stale = MaterialPreferences.from_payload(
            {
                "custom_materials": [],
                "custom_presets": {
                    "仍可使用": ["身份证", "已删除材料"],
                    "已经为空": ["已删除材料"],
                },
            }
        )
        self.assertEqual(stale.get_preset("仍可使用"), ("身份证",))
        self.assertIsNone(stale.get_preset("已经为空"))

    def test_rejects_duplicates_invalid_names_and_builtin_mutation(self) -> None:
        preferences = MaterialPreferences()
        with self.assertRaisesRegex(ValueError, "已经存在"):
            preferences.add_material("身份证")
        with self.assertRaisesRegex(ValueError, "不能包含"):
            preferences.add_material("户口本/扫描件")
        with self.assertRaisesRegex(ValueError, "安全的文件夹名称"):
            preferences.add_material("CON")
        with self.assertRaisesRegex(ValueError, "内置预设不能覆盖"):
            preferences.save_preset("入职材料", ["身份证"])
        with self.assertRaisesRegex(ValueError, "内置材料不能删除"):
            preferences.remove_material("身份证")
        with self.assertRaisesRegex(ValueError, "内置预设不能删除"):
            preferences.delete_preset("入职材料")


class MaterialPreferencesGUITest(unittest.TestCase):
    def _make_app(self) -> HRToolkitApp:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._material_preferences = MaterialPreferences(custom_materials=["户口本"])
        app.material_types_selected = {
            name: _Value(True)
            for name in app._material_preferences.available_materials
        }
        app.material_preset_name = _Value("入职材料")
        app.material_collect_all = _Value(True)
        app._on_material_collect_all_changed = Mock()
        return app

    def test_apply_builtin_preset_then_manual_change_controls_final_selection(self) -> None:
        app = self._make_app()

        app._apply_material_preset()

        self.assertFalse(app.material_collect_all.get())
        self.assertEqual(app._selected_material_names(), ["身份证", "劳动合同"])
        app.material_types_selected["劳动合同"].set(False)
        app.material_types_selected["户口本"].set(True)
        self.assertEqual(app._selected_material_names(), ["身份证", "户口本"])
        app._on_material_collect_all_changed.assert_called_once()

    def test_select_and_deselect_include_custom_materials(self) -> None:
        app = self._make_app()
        app._deselect_all_material_types()
        self.assertEqual(app._selected_material_names(), [])

        app._select_all_material_types()
        self.assertEqual(
            app._selected_material_names(),
            list(app._material_preferences.available_materials),
        )

    def test_run_uses_final_checkbox_state_for_builtin_and_custom_materials(self) -> None:
        app = self._make_app()
        app.material_collect_all.set(False)
        for variable in app.material_types_selected.values():
            variable.set(False)
        app.material_types_selected["身份证"].set(True)
        app.material_types_selected["户口本"].set(True)
        app.material_target_input = _Value("张三")
        app.summary_path = _Value("")
        app.material_create_zip = _Value(False)
        app.material_use_ocr_cache = _Value(False)
        app._prepare_result_output_dir = Mock()
        app._begin_tool_run = Mock()
        app._clear_log = Mock()
        app._write_log = Mock()
        app._start_tool_worker = Mock()

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            library = root / "资料库"
            output = root / "输出"
            result_dir = root / "本次结果"
            library.mkdir()
            output.mkdir()
            app.input_path = _Value(str(library))
            app.output_dir = _Value(str(output))
            app._prepare_result_output_dir.return_value = result_dir

            with patch("hr_toolkit.gui.app.messagebox") as mocked_messagebox:
                app._run_material_collector()

        mocked_messagebox.showwarning.assert_not_called()
        app._start_tool_worker.assert_called_once()
        call = app._start_tool_worker.call_args
        self.assertEqual(call.kwargs["material_types"], ["身份证", "户口本"])
        self.assertFalse(call.kwargs["collect_all"])


if __name__ == "__main__":
    unittest.main()
