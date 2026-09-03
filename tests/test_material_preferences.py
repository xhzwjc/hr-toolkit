from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from hr_toolkit.material_preferences import (
    BUILTIN_MATERIALS,
    BUILTIN_MATERIAL_PRESETS,
    MaterialPreferences,
)
from hr_toolkit.tools.material_collector import (
    LIBRARY_MODE_FLAT_OCR,
    LIBRARY_MODE_PERSON_FOLDER,
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


if __name__ == "__main__":
    unittest.main()
