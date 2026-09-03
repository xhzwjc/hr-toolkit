from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from hr_toolkit.gui_qt.form_specs import (
    FormValidationError,
    build_invocation,
    default_values,
    spec_for,
)


class QtFormSpecTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.output = self.root / "project"
        self.output.mkdir()
        self.file_a = self.root / "a.xlsx"
        self.file_b = self.root / "b.xls"
        self.file_a.touch()
        self.file_b.touch()
        self.folder = self.root / "folder"
        self.folder.mkdir()

    def tearDown(self) -> None:
        self.temp.cleanup()

    def invocation(self, nav_id: str, variant: str = "default", **changes):
        spec = spec_for(nav_id, variant)
        values = default_values(spec)
        values.update(changes.pop("values", {}))
        return build_invocation(
            spec,
            input_paths=changes.pop("input_paths", [self.file_a]),
            support_text=changes.pop("support_text", str(self.file_b)),
            values=values,
            output_dir=self.output,
            **changes,
        )

    def test_all_business_tools_map_to_existing_functions_and_project_names(self) -> None:
        cases = (
            ("social_security", "default", "hr_toolkit.tools.social_security", "generate_social_security_reports", "社保明细与汇总"),
            ("insurance_ledger", "default", "hr_toolkit.tools.insurance_ledger", "generate_insurance_ledger", "保险台账与预警"),
            ("data_statistics", "default", "hr_toolkit.tools.data_statistics", "generate_data_statistics_reports", "考勤与周月报"),
            ("salary_split", "default", "hr_toolkit.tools.salary_split", "split_salary_by_company", "工资表拆分"),
            ("salary_merge", "default", "hr_toolkit.tools.salary_merge", "merge_monthly_salary", "多月工资合并"),
            ("personnel_change_merge", "merge", "hr_toolkit.tools.personnel_change_merge", "merge_personnel_changes", "异动汇总"),
            ("personnel_change_merge", "roster", "hr_toolkit.tools.personnel_change_merge", "update_roster_from_change_summaries", "花名册更新"),
            ("archive_import", "import", "hr_toolkit.tools.archive_import", "import_archive_transfers", "档案入库"),
            ("archive_import", "export", "hr_toolkit.tools.archive_import", "export_company_archive_tables", "档案表生成"),
        )
        for nav_id, variant, module, function, project_name in cases:
            with self.subTest(nav_id=nav_id, variant=variant):
                support_text = "" if (nav_id, variant) in {
                    ("salary_split", "default"),
                    ("salary_merge", "default"),
                    ("personnel_change_merge", "merge"),
                    ("archive_import", "import"),
                    ("archive_import", "export"),
                } else str(self.file_b)
                invocation = self.invocation(
                    nav_id,
                    variant,
                    support_text=support_text,
                )
                self.assertEqual(invocation.function_module, module)
                self.assertEqual(invocation.function_name, function)
                self.assertEqual(invocation.tool_name, project_name)
                self.assertTrue(callable(invocation.resolve_function()))

    def test_data_statistics_resolves_dates_identically_before_call(self) -> None:
        defaults = default_values(spec_for("data_statistics"))
        self.assertEqual(defaults["week_start"], "")
        self.assertEqual(defaults["week_end"], "")
        self.assertEqual(defaults["month_start"], "")
        self.assertEqual(defaults["month_end"], "")
        invocation = self.invocation(
            "data_statistics",
            values={
                "week_start": "2026-08-03",
                "week_end": "2026-08-31",
                "month_start": "2026-08-01",
                "month_end": "2026-08-31",
                "remark_unit": "hour",
                "include_business_trip": True,
                "include_workday_business_trip": True,
            },
        )
        self.assertEqual(invocation.kwargs["week_start"], date(2026, 8, 3))
        self.assertEqual(invocation.kwargs["week_end"], date(2026, 8, 31))
        self.assertEqual(invocation.kwargs["month_start"], date(2026, 8, 1))
        self.assertEqual(invocation.kwargs["month_end"], date(2026, 8, 31))
        self.assertEqual(invocation.kwargs["remark_unit"], "hour")
        self.assertTrue(invocation.kwargs["include_business_trip"])
        self.assertTrue(invocation.kwargs["include_workday_business_trip"])

    def test_material_collector_preserves_direct_target_and_ocr_rules(self) -> None:
        invocation = self.invocation(
            "material_collector",
            input_paths=[self.folder],
            support_text="",
            values={
                "target_input": "张三, 李四",
                "library_mode": "person_folder",
                "collect_all": True,
                "create_zip": True,
                "use_ocr_cache": True,
                "material_types": [],
            },
        )
        self.assertEqual(invocation.kwargs["roster_source"], "张三, 李四")
        self.assertTrue(invocation.kwargs["collect_all"])
        self.assertTrue(invocation.kwargs["create_zip"])
        self.assertFalse(invocation.kwargs["use_ocr_cache"])
        self.assertIsNone(invocation.kwargs["material_types"])

    def test_hidden_optional_support_does_not_block_direct_workflows(self) -> None:
        missing = self.root / "previously-selected-but-now-missing.xlsx"
        material = self.invocation(
            "material_collector",
            input_paths=[self.folder],
            support_text=str(missing),
            values={
                "target_input": "张三",
                "library_mode": "person_folder",
                "collect_all": True,
            },
        )
        self.assertEqual(material.kwargs["roster_source"], "张三")

        rename = self.invocation(
            "folder_rename",
            input_paths=[self.folder],
            support_text=str(missing),
            values={"rename_mode": "append", "rename_text": "_已核对"},
            preview=True,
        )
        self.assertNotIn("excel_path", rename.kwargs)

    def test_folder_rename_preview_and_confirmed_excel_call_match(self) -> None:
        spec = spec_for("folder_rename")
        values = default_values(spec)
        values.update({"rename_mode": "excel", "file_type": "image"})
        preview = build_invocation(
            spec,
            input_paths=[self.folder],
            support_text=str(self.file_a),
            values=values,
            output_dir=self.output,
            preview=True,
        )
        self.assertTrue(preview.preview)
        self.assertTrue(preview.kwargs["dry_run"])
        self.assertEqual(preview.tool_name, "资料文件夹改名")
        self.assertEqual(preview.description, "资料文件夹改名-按 Excel 人名顺序批量重命名")

        confirmed = build_invocation(
            spec,
            input_paths=[self.folder],
            support_text=str(self.file_a),
            values=values,
            output_dir=self.output,
            preview=False,
            preview_result={
                "operations": [{"source": "/tmp/1.jpg", "target": "/tmp/张三.jpg"}],
                "warnings": ["提醒"],
            },
        )
        self.assertFalse(confirmed.kwargs["dry_run"])
        self.assertEqual(confirmed.kwargs["expected_operations"], [("1.jpg", "张三.jpg")])
        self.assertEqual(confirmed.kwargs["expected_warnings"], ["提醒"])

    def test_invalid_inputs_fail_before_business_code_runs(self) -> None:
        with self.assertRaisesRegex(FormValidationError, "只支持 .xlsx"):
            unsupported = self.root / "bad.txt"
            unsupported.touch()
            self.invocation("salary_split", input_paths=[unsupported], support_text="")
        with self.assertRaisesRegex(FormValidationError, "只支持选择一个"):
            self.invocation(
                "salary_split",
                input_paths=[self.file_a, self.file_b],
                support_text="",
            )

    def test_material_collector_invocation_collect_all_and_specific_checkboxes(self) -> None:
        # 1. collect_all is True: material_types is None, OCR cache disabled if target_text given in person_folder mode
        inv_all = self.invocation(
            "material_collector",
            input_paths=[self.folder],
            support_text="",
            values={
                "target_input": "张三",
                "collect_all": True,
                "material_types": ["身份证"],
                "library_mode": "person_folder",
            },
        )
        self.assertIsNone(inv_all.kwargs["material_types"])
        self.assertTrue(inv_all.kwargs["collect_all"])
        self.assertFalse(inv_all.kwargs["use_ocr_cache"])
        self.assertEqual(inv_all.kwargs["roster_source"], "张三")

        # 2. Specific checkboxes when collect_all is False
        inv_specific = self.invocation(
            "material_collector",
            input_paths=[self.folder],
            support_text=str(self.file_a),
            values={
                "target_input": "",
                "collect_all": False,
                "material_types": ["身份证", "户口本"],
                "library_mode": "person_folder",
                "use_ocr_cache": True,
            },
        )
        self.assertEqual(inv_specific.kwargs["material_types"], ["身份证", "户口本"])
        self.assertFalse(inv_specific.kwargs["collect_all"])
        self.assertTrue(inv_specific.kwargs["use_ocr_cache"])
        self.assertEqual(inv_specific.kwargs["roster_source"], self.file_a)

        # 3. flat_ocr mode forces use_ocr_cache=True
        inv_flat = self.invocation(
            "material_collector",
            input_paths=[self.folder],
            support_text="李四",
            values={
                "target_input": "李四",
                "collect_all": True,
                "library_mode": "flat_ocr",
                "use_ocr_cache": False,
            },
        )
        self.assertTrue(inv_flat.kwargs["use_ocr_cache"])

        # 4. Validates neither collect_all nor materials
        with self.assertRaisesRegex(FormValidationError, "请至少选择一种材料"):
            self.invocation(
                "material_collector",
                input_paths=[self.folder],
                support_text="王五",
                values={
                    "target_input": "王五",
                    "collect_all": False,
                    "material_types": [],
                },
            )

        # 5. Validates missing roster source
        with self.assertRaisesRegex(FormValidationError, "请输入员工姓名"):
            self.invocation(
                "material_collector",
                input_paths=[self.folder],
                support_text="",
                values={
                    "target_input": "",
                    "collect_all": True,
                },
            )


if __name__ == "__main__":
    unittest.main()
