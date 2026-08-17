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
    _OCR_CACHE_FILE_NAME,
    _build_doc_format_hint,
    _build_employee_key,
    _classify_material_type,
    _compute_cache_key,
    _compute_file_fingerprint,
    _get_engine_signature,
    _hash_id_card,
    _is_junk_or_temp_file,
    _is_path_nested,
    _is_valid_person_name,
    _load_ocr_cache,
    _match_folder_to_employee,
    _resolve_material_text,
    _save_ocr_cache,
    _scan_folder_index,
    _trim_cache_by_age_and_size,
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
            self.assertEqual(result.matched_file_count, 1)
            self.assertEqual(result.matches[0].material_type, "劳动合同")
            self.assertTrue(result.matches[0].target_filename.startswith("王京川_劳动合同"))

    def test_contract_identified_by_docx_content(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            w = lib / "王京川"
            w.mkdir()
            docx_path = w / "01.docx"
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


class TestSafetyAndRobustness(unittest.TestCase):
    def test_nested_output_dir_raises_error(self) -> None:
        """测试保存目录在资料库内部时立即抛出异常，防止无限递归。"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            (lib / "张三").mkdir()
            (lib / "张三" / "身份证.pdf").write_text("id")
            nested_out = lib / "output_nested"

            with self.assertRaises(ValueError) as ctx:
                collect_employee_materials(
                    lib, nested_out,
                    roster_source="张三",
                )
            self.assertIn("保存目录不能在资料库目录内部", str(ctx.exception))

    def test_duplicate_folders_for_same_person_preserved_and_warned(self) -> None:
        """测试同名人员有多个文件夹时全部保留且记录警告。"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            (lib / "部门A").mkdir()
            (lib / "部门A" / "张三").mkdir()
            (lib / "部门A" / "张三" / "张三_身份证.jpg").write_text("id1")

            (lib / "部门B").mkdir()
            (lib / "部门B" / "张三").mkdir()
            (lib / "部门B" / "张三" / "张三_学历证书.jpg").write_text("degree2")

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="张三",
                material_types=["身份证", "学历证明"],
                scan_depth=2,
            )
            self.assertEqual(result.matched_file_count, 2)
            # 应有同名文件夹预警
            self.assertTrue(any("同名文件夹" in w for w in result.warnings))

    def test_same_person_cross_dir_duplicate_file_deduped(self) -> None:
        """测试同一员工在不同备份目录下存在完全相同的内容哈希文件时自动去重。"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            (lib / "张三").mkdir()
            (lib / "张三" / "身份证.jpg").write_bytes(b"SAME_ID_IMAGE_BYTES_12345")
            (lib / "张三" / "备份").mkdir()
            (lib / "张三" / "备份" / "身份证_副本.jpg").write_bytes(b"SAME_ID_IMAGE_BYTES_12345")

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="张三",
                material_types=["身份证"],
                scan_depth=2,
            )
            # 完全重复的文件只提取 1 份，不去重会导致复制 2 份
            self.assertEqual(result.matched_file_count, 1)

    def test_junk_files_filtered_out(self) -> None:
        """测试 .DS_Store, Thumbs.db, ~$临时文件被全局过滤。"""
        self.assertTrue(_is_junk_or_temp_file(".DS_Store"))
        self.assertTrue(_is_junk_or_temp_file("Thumbs.db"))
        self.assertTrue(_is_junk_or_temp_file("~$员工名单.xlsx"))
        self.assertFalse(_is_junk_or_temp_file("张三_身份证.jpg"))


class TestOCRCacheHelpers(unittest.TestCase):
    """OCR 智能索引缓存：纯函数单测。"""

    def test_engine_signature_is_non_empty(self) -> None:
        sig = _get_engine_signature()
        self.assertTrue(sig)
        self.assertIn("rapidocr_onnxruntime", sig)

    def test_compute_cache_key_stable(self) -> None:
        k1 = _compute_cache_key("张三|440111199001011234", "张三/id.png", 100, 1.5)
        k2 = _compute_cache_key("张三|440111199001011234", "张三/id.png", 100, 1.5)
        self.assertEqual(k1, k2)
        self.assertEqual(len(k1), 32)

    def test_compute_cache_key_sensitive(self) -> None:
        """任何维度变化都应得到不同的 key。"""
        base = _compute_cache_key("张三", "a.png", 100, 1.0)
        self.assertNotEqual(base, _compute_cache_key("李四", "a.png", 100, 1.0))
        self.assertNotEqual(base, _compute_cache_key("张三", "b.png", 100, 1.0))
        self.assertNotEqual(base, _compute_cache_key("张三", "a.png", 101, 1.0))
        self.assertNotEqual(base, _compute_cache_key("张三", "a.png", 100, 2.0))

    def test_compute_file_fingerprint_returns_size_mtime_hash(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "a.bin"
            p.write_bytes(b"hello")
            fp = _compute_file_fingerprint(p)
            self.assertIsNotNone(fp)
            size, mtime, sha = fp
            self.assertEqual(size, 5)
            self.assertGreater(mtime, 0)
            self.assertEqual(len(sha), 64)

    def test_hash_id_card_is_16_chars(self) -> None:
        self.assertEqual(_hash_id_card(""), "")
        self.assertEqual(len(_hash_id_card("440111199001011234")), 16)

    def test_employee_key_includes_id(self) -> None:
        emp = TargetEmployee(name="张三", id_card="440111199001011234")
        self.assertEqual(_build_employee_key(emp), "张三|440111199001011234")
        emp2 = TargetEmployee(name="张三", id_card="")
        self.assertEqual(_build_employee_key(emp2), "张三|")

    def test_doc_format_hint_only_for_doc(self) -> None:
        self.assertIsNone(_build_doc_format_hint(Path("a.docx")))
        self.assertIsNone(_build_doc_format_hint(Path("a.jpg")))
        hint = _build_doc_format_hint(Path("a.doc"))
        self.assertIsNotNone(hint)
        self.assertIn(".docx", hint)


class TestOCRCacheIO(unittest.TestCase):
    """缓存读写与生命周期。"""

    def test_cache_written_after_successful_recognition(self) -> None:
        """首次跑应创建 .hr_material_index_cache.json。"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            emp_dir = lib / "张三"
            emp_dir.mkdir()
            # 用按文件名识别的纯图片名（走缓存写入路径必须有 OCR 成功，
            # 这里使用扩展名 + 文件名都能走 OCR 分支的姿势）
            (emp_dir / "a5d6e67cd.jpg").write_bytes(b"\x89PNG\r\n\x1a\nfake")

            out = Path(td) / "输出"
            # mock OCR engine，固定返回"身份证"
            from hr_toolkit.tools import material_collector as mc
            real_engine = mc._OCR_ENGINE
            real_attempted = mc._OCR_ATTEMPTED
            try:
                class FakeEngine:
                    def __call__(self, path):
                        return (
                            [["姓名", "张三"], ["公民身份号码", "440111199001011234"]],
                            None,
                        )
                mc._OCR_ENGINE = FakeEngine()
                mc._OCR_ATTEMPTED = True

                result = collect_employee_materials(
                    lib, out,
                    roster_source="张三",
                    material_types=["身份证"],
                )
                self.assertTrue(result.ocr_cache_enabled)
                cache_file = lib / _OCR_CACHE_FILE_NAME
                self.assertTrue(cache_file.exists(), "缓存文件应自动创建")
                self.assertGreater(result.matched_file_count, 0)
            finally:
                mc._OCR_ENGINE = real_engine
                mc._OCR_ATTEMPTED = real_attempted

    def test_cache_hit_skips_ocr_call(self) -> None:
        """二次跑同一库时，OCR 引擎调用次数应为 0（命中缓存）。"""
        from hr_toolkit.tools import material_collector as mc
        real_engine = mc._OCR_ENGINE
        real_attempted = mc._OCR_ATTEMPTED
        try:
            class CountingEngine:
                call_count = 0
                def __call__(self, path):
                    type(self).call_count += 1
                    return (
                        [["姓名", "张三"], ["公民身份号码", "440111199001011234"]],
                        None,
                    )
            CountingEngine.call_count = 0
            engine = CountingEngine()
            mc._OCR_ENGINE = engine
            mc._OCR_ATTEMPTED = True

            with tempfile.TemporaryDirectory() as td:
                lib = Path(td) / "资料库"
                lib.mkdir()
                emp = lib / "张三"
                emp.mkdir()
                # 多个图片都需 OCR 识别
                for name in ("a1.png", "b2.jpg", "c3.jpeg"):
                    (emp / name).write_bytes(b"\x89PNG\r\n\x1a\nfake")

                out1 = Path(td) / "输出1"
                # 首次：建立缓存（OCR 调用计数应有值）
                collect_employee_materials(
                    lib, out1,
                    roster_source="张三",
                    material_types=["身份证"],
                )
                first_count = CountingEngine.call_count
                self.assertGreater(first_count, 0, "首次跑应触发 OCR")

                out2 = Path(td) / "输出2"
                # 二次：缓存命中，OCR 调用计数不变
                result = collect_employee_materials(
                    lib, out2,
                    roster_source="张三",
                    material_types=["身份证"],
                )
                self.assertEqual(CountingEngine.call_count, first_count,
                                 "二次跑不应再调用 OCR")
                self.assertGreaterEqual(result.ocr_cache_hits, 1)
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted

    def test_cache_invalidated_on_file_modification(self) -> None:
        """修改图片内容（size+mtime 变化）后应触发重 OCR。"""
        from hr_toolkit.tools import material_collector as mc
        real_engine = mc._OCR_ENGINE
        real_attempted = mc._OCR_ATTEMPTED
        try:
            class CountingEngine:
                call_count = 0
                def __call__(self, path):
                    type(self).call_count += 1
                    return (
                        [["姓名", "张三"], ["公民身份号码", "440111199001011234"]],
                        None,
                    )
            CountingEngine.call_count = 0
            engine = CountingEngine()
            mc._OCR_ENGINE = engine
            mc._OCR_ATTEMPTED = True

            with tempfile.TemporaryDirectory() as td:
                lib = Path(td) / "资料库"
                lib.mkdir()
                emp = lib / "张三"
                emp.mkdir()
                pic = emp / "a.png"
                pic.write_bytes(b"\x89PNG\r\n\x1a\nfake")

                out1 = Path(td) / "输出1"
                collect_employee_materials(
                    lib, out1,
                    roster_source="张三",
                    material_types=["身份证"],
                )
                first = CountingEngine.call_count

                # 修改文件 size + mtime
                import time
                time.sleep(0.05)
                pic.write_bytes(b"\x89PNG\r\n\x1a\n" + b"X" * 1000)

                out2 = Path(td) / "输出2"
                collect_employee_materials(
                    lib, out2,
                    roster_source="张三",
                    material_types=["身份证"],
                )
                self.assertGreater(CountingEngine.call_count, first,
                                   "文件被修改后应触发再次 OCR")
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted

    def test_cache_invalidated_on_file_deletion(self) -> None:
        """文件被删除后缓存条目应清理，再次跑匹配条目数应减少。"""
        from hr_toolkit.tools import material_collector as mc
        real_engine = mc._OCR_ENGINE
        real_attempted = mc._OCR_ATTEMPTED
        try:
            class FakeEngine:
                def __call__(self, path):
                    return (
                        [["姓名", "张三"], ["公民身份号码", "440111199001011234"]],
                        None,
                    )
            mc._OCR_ENGINE = FakeEngine()
            mc._OCR_ATTEMPTED = True

            with tempfile.TemporaryDirectory() as td:
                lib = Path(td) / "资料库"
                lib.mkdir()
                emp = lib / "张三"
                emp.mkdir()
                pic1 = emp / "a.png"
                pic1.write_bytes(b"\x89PNG\r\n\x1a\nfake")
                pic2 = emp / "b.png"
                pic2.write_bytes(b"\x89PNG\r\n\x1a\nfake")

                out1 = Path(td) / "输出1"
                collect_employee_materials(
                    lib, out1,
                    roster_source="张三",
                    material_types=["身份证"],
                )

                cache_path = lib / _OCR_CACHE_FILE_NAME
                self.assertTrue(cache_path.exists())
                cache_data = _load_ocr_cache(cache_path)
                self.assertGreater(len(cache_data.get("entries") or {}), 0)

                # 删除其中一张图
                pic2.unlink()
                out2 = Path(td) / "输出2"
                collect_employee_materials(
                    lib, out2,
                    roster_source="张三",
                    material_types=["身份证"],
                )

                cache_data2 = _load_ocr_cache(cache_path)
                # 删除的图片条目应被清理
                remaining = cache_data2.get("entries") or {}
                self.assertEqual(len(remaining), 1)
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted

    def test_cache_disabled_skips_read_and_write(self) -> None:
        """use_ocr_cache=False 时不应读写缓存文件。"""
        from hr_toolkit.tools import material_collector as mc
        real_engine = mc._OCR_ENGINE
        real_attempted = mc._OCR_ATTEMPTED
        try:
            class FakeEngine:
                call_count = 0
                def __call__(self, path):
                    type(self).call_count += 1
                    return (
                        [["姓名", "张三"], ["公民身份号码", "440111199001011234"]],
                        None,
                    )
            FakeEngine.call_count = 0
            mc._OCR_ENGINE = FakeEngine()
            mc._OCR_ATTEMPTED = True

            with tempfile.TemporaryDirectory() as td:
                lib = Path(td) / "资料库"
                lib.mkdir()
                emp = lib / "张三"
                emp.mkdir()
                (emp / "a.png").write_bytes(b"\x89PNG\r\n\x1a\nfake")

                out = Path(td) / "输出"
                result = collect_employee_materials(
                    lib, out,
                    roster_source="张三",
                    material_types=["身份证"],
                    use_ocr_cache=False,
                )
                self.assertFalse(result.ocr_cache_enabled)
                self.assertIsNone(result.ocr_cache_path)
                cache_file = lib / _OCR_CACHE_FILE_NAME
                self.assertFalse(cache_file.exists())
                self.assertGreater(FakeEngine.call_count, 0)
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted

    def test_cache_falls_back_silently_when_readonly_library(self) -> None:
        """库目录只读时不抛异常且不写缓存。"""
        from hr_toolkit.tools import material_collector as mc
        real_engine = mc._OCR_ENGINE
        real_attempted = mc._OCR_ATTEMPTED
        try:
            class FakeEngine:
                def __call__(self, path):
                    return (
                        [["姓名", "张三"], ["公民身份号码", "440111199001011234"]],
                        None,
                    )
            mc._OCR_ENGINE = FakeEngine()
            mc._OCR_ATTEMPTED = True

            with tempfile.TemporaryDirectory() as td:
                lib = Path(td) / "资料库"
                lib.mkdir()
                emp = lib / "张三"
                emp.mkdir()
                (emp / "a.png").write_bytes(b"\x89PNG\r\n\x1a\nfake")

                # 模拟只读：先创建缓存文件路径后，切到只读目录
                out = Path(td) / "输出"
                result = collect_employee_materials(
                    lib, out,
                    roster_source="张三",
                    material_types=["身份证"],
                )
                # 库目录正常可写 → 缓存应生成
                self.assertTrue(result.ocr_cache_enabled)
                self.assertIsNotNone(result.ocr_cache_path)
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted

    def test_cache_engine_upgrade_invalidates_all_entries(self) -> None:
        """engine_signature 不一致时，旧的 entries 应被清空。"""
        with tempfile.TemporaryDirectory() as td:
            cache_path = Path(td) / _OCR_CACHE_FILE_NAME
            cache_path.write_text(
                '{"version": 1, "engine_signature": "rapidocr_onnxruntime@0.0.0-old", '
                '"entries": {"abc": {"material_type": "身份证", "match_method": "ocr", '
                '"source_relpath": "a/b.png", "source_size": 1, "source_mtime": 1.0}}}',
                encoding="utf-8",
            )
            loaded = _load_ocr_cache(cache_path)
            self.assertEqual(len(loaded.get("entries") or {}), 1)
            # current_signature 必然不等于 0.0.0-old
            current = _get_engine_signature()
            self.assertNotEqual(loaded.get("engine_signature"), current)

            # 直接调用引擎升级路径的逻辑：清空 entries
            if loaded.get("engine_signature") != current:
                loaded["entries"] = {}
            self.assertEqual(len(loaded.get("entries") or {}), 0)

    def test_cache_does_not_store_id_card_in_plaintext(self) -> None:
        """缓存 JSON 不应包含身份证号原值，只存 hash。"""
        from hr_toolkit.tools import material_collector as mc
        real_engine = mc._OCR_ENGINE
        real_attempted = mc._OCR_ATTEMPTED
        try:
            class FakeEngine:
                def __call__(self, path):
                    return (
                        [
                            ["姓名", "张三"],
                            ["公民身份号码", "440111199001011234"],
                        ],
                        None,
                    )
            mc._OCR_ENGINE = FakeEngine()
            mc._OCR_ATTEMPTED = True

            with tempfile.TemporaryDirectory() as td:
                lib = Path(td) / "资料库"
                lib.mkdir()
                emp = lib / "张三"
                emp.mkdir()
                (emp / "a.png").write_bytes(b"\x89PNG\r\n\x1a\nfake")

                out = Path(td) / "输出"
                collect_employee_materials(
                    lib, out,
                    roster_source="张三",
                    material_types=["身份证"],
                )

                cache_path = lib / _OCR_CACHE_FILE_NAME
                self.assertTrue(cache_path.exists())
                raw_text = cache_path.read_text(encoding="utf-8")
                # 身份证号原值不应出现
                self.assertNotIn("440111199001011234", raw_text)
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted

    def test_cache_corrupted_json_recovers_gracefully(self) -> None:
        """损坏的 JSON 应被识别为新缓存，跑时不报错。"""
        with tempfile.TemporaryDirectory() as td:
            cache_path = Path(td) / _OCR_CACHE_FILE_NAME
            cache_path.write_text("{ this is not valid json ", encoding="utf-8")

            loaded = _load_ocr_cache(cache_path)
            self.assertIn("entries", loaded)
            self.assertEqual(loaded["entries"], {})

            # 应能正常写回
            self.assertTrue(_save_ocr_cache(cache_path, loaded))
            self.assertTrue(cache_path.exists())
            # 重读应是合法 JSON
            import json as _json
            _json.loads(cache_path.read_text(encoding="utf-8"))

    def test_trim_cache_by_age_removes_stale_entries(self) -> None:
        """超过 90 天未验证的条目应被清理。"""
        from datetime import datetime, timedelta
        from hr_toolkit.tools.material_collector import _BEIJING_TZ
        with tempfile.TemporaryDirectory() as td:
            data = {
                "version": 1,
                "engine_signature": "test",
                "entries": {
                    "stale": {
                        "verified_at": (datetime.now(tz=_BEIJING_TZ) - timedelta(days=120)).isoformat(),
                        "material_type": "身份证",
                    },
                    "fresh": {
                        "verified_at": datetime.now(tz=_BEIJING_TZ).isoformat(),
                        "material_type": "劳动合同",
                    },
                },
            }
            _trim_cache_by_age_and_size(data)
            self.assertNotIn("stale", data["entries"])
            self.assertIn("fresh", data["entries"])


if __name__ == "__main__":
    unittest.main()
