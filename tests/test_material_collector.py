from __future__ import annotations

import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from hr_toolkit.tools.material_collector import (
    MODE_BY_EMPLOYEE,
    MODE_BY_MATERIAL,
    MODE_FLAT,
    TargetEmployee,
    collect_employee_materials,
    parse_employee_roster,
    _match_folder_to_employee,
    _resolve_material_text,
    _scan_folder_index,
    _is_valid_person_name,
    _classify_material_type,
)


class TestResolveMAterialText(unittest.TestCase):
    def test_single_keyword(self) -> None:
        self.assertEqual(_resolve_material_text("身份证"), ["身份证"])

    def test_multiple_comma_separated(self) -> None:
        result = _resolve_material_text("身份证，合同")
        self.assertIn("身份证", result)
        self.assertIn("劳动合同", result)

    def test_empty(self) -> None:
        self.assertEqual(_resolve_material_text(""), [])

    def test_unrecognized_kept_as_is(self) -> None:
        result = _resolve_material_text("户口本")
        self.assertEqual(result, ["户口本"])

    def test_synonym_resolves_to_canonical(self) -> None:
        result = _resolve_material_text("毕业证")
        self.assertEqual(result, ["学历证明"])


class TestParseEmployeeRoster(unittest.TestCase):
    def test_plain_text_single_name(self) -> None:
        employees = parse_employee_roster("张三")
        self.assertEqual(len(employees), 1)
        self.assertEqual(employees[0].name, "张三")

    def test_plain_text_comma_separated(self) -> None:
        employees = parse_employee_roster("张三, 李四，王五")
        self.assertEqual(len(employees), 3)
        self.assertEqual(employees[0].name, "张三")
        self.assertEqual(employees[1].name, "李四")
        self.assertEqual(employees[2].name, "王五")

    def test_plain_text_with_id(self) -> None:
        employees = parse_employee_roster("张三 440111199001011234")
        self.assertEqual(len(employees), 1)
        self.assertEqual(employees[0].name, "张三")
        self.assertEqual(employees[0].id_card, "440111199001011234")

    def test_plain_text_id_only(self) -> None:
        employees = parse_employee_roster("440111199001011234")
        self.assertEqual(len(employees), 1)
        self.assertEqual(employees[0].id_card, "440111199001011234")

    def test_plain_text_with_materials(self) -> None:
        employees = parse_employee_roster("张三 身份证, 李四 合同")
        self.assertEqual(len(employees), 2)
        self.assertEqual(employees[0].name, "张三")
        self.assertEqual(employees[0].per_person_materials, ("身份证",))
        self.assertEqual(employees[1].name, "李四")
        self.assertEqual(employees[1].per_person_materials, ("劳动合同",))

    def test_dict_list_with_materials(self) -> None:
        data = [
            {"姓名": "张三", "材料": "身份证，合同"},
            {"姓名": "李四", "材料": ""},
        ]
        employees = parse_employee_roster(data)
        self.assertEqual(len(employees), 2)
        self.assertIn("身份证", employees[0].per_person_materials)
        self.assertIn("劳动合同", employees[0].per_person_materials)
        self.assertEqual(employees[1].per_person_materials, ())

    def test_dedup_by_name_and_id(self) -> None:
        employees = parse_employee_roster(["张三", "张三", "李四"])
        self.assertEqual(len(employees), 2)

    def test_noise_words_filtered(self) -> None:
        self.assertFalse(_is_valid_person_name("序号"))
        self.assertFalse(_is_valid_person_name("姓名"))
        self.assertFalse(_is_valid_person_name("合计"))
        self.assertFalse(_is_valid_person_name("123"))
        self.assertTrue(_is_valid_person_name("张三"))

    def test_excel_single_person_does_not_parse_extra(self) -> None:
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as td:
            wb = Workbook()
            ws = wb.active
            ws["A1"] = "序号"
            ws["B1"] = "姓名"
            ws["A2"] = "1"
            ws["B2"] = "张三"
            path = Path(td) / "roster.xlsx"
            wb.save(path)
            wb.close()

            employees = parse_employee_roster(path)
            self.assertEqual(len(employees), 1)
            self.assertEqual(employees[0].name, "张三")


class TestExclusiveClassification(unittest.TestCase):
    """测试严密互斥分类：绝不把身份证、体检表等非合同文件作为劳动合同输出。"""

    def test_non_contract_files_not_collected_as_contract(self) -> None:
        """当员工目录下只有身份证、体检表、离职证明时，勾选劳动合同应如实报缺失，绝不重命名输出！"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            w = lib / "王京川"
            w.mkdir()
            (w / "王京川_身份证正面.jpg").write_text("id front")
            (w / "王京川_体检报告.pdf").write_text("health report")
            (w / "王京川_离职证明.doc").write_text("resignation")

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="王京川",
                material_types=["劳动合同"],
            )
            self.assertEqual(result.total_employees, 1)
            # 劳动合同应判定为缺失，0 个匹配！
            self.assertEqual(result.matched_file_count, 0)
            self.assertIn("王京川", result.missing_records)
            self.assertIn("劳动合同", result.missing_records["王京川"])

    def test_contract_identified_by_filename(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            w = lib / "王京川"
            w.mkdir()
            (w / "王京川_劳动合同.docx").write_text("contract doc")
            (w / "王京川_身份证.jpg").write_text("id")

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="王京川",
                material_types=["劳动合同"],
            )
            # 只有劳动合同被提取，身份证不被作为劳动合同！
            self.assertEqual(result.matched_file_count, 1)
            self.assertEqual(result.matches[0].material_type, "劳动合同")
            self.assertTrue(result.matches[0].target_filename.startswith("王京川_劳动合同"))

    def test_contract_identified_by_docx_content(self) -> None:
        """非标准命名的 docx 文件（如 01.docx），内部含有劳动合同正文时精准识别。"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            w = lib / "王京川"
            w.mkdir()
            docx_path = w / "01.docx"
            # 创建合法 docx 内部结构
            with zipfile.ZipFile(docx_path, "w") as zf:
                zf.writestr(
                    "word/document.xml",
                    '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>劳动合同书 甲方与乙方自愿订立本用工合同，约定工作内容与劳动报酬。</w:t></w:r></w:p></w:body></w:document>',
                )

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="王京川",
                material_types=["劳动合同"],
            )
            self.assertEqual(result.matched_file_count, 1)
            self.assertEqual(result.matches[0].material_type, "劳动合同")


class TestCollectAll(unittest.TestCase):
    def test_single_person_only_copies_target_folder(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            (lib / "张三").mkdir()
            (lib / "张三" / "身份证.pdf").write_text("id")
            (lib / "李四").mkdir()
            (lib / "李四" / "李四资料.pdf").write_text("lisi")

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="张三",
                collect_all=True,
            )
            self.assertEqual(result.total_employees, 1)
            self.assertEqual(result.matched_file_count, 1)
            self.assertTrue((out / "张三").is_dir())
            self.assertFalse((out / "李四").exists())


if __name__ == "__main__":
    unittest.main()
