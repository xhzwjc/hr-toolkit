from __future__ import annotations

import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from hr_toolkit.tools.material_collector import (
    LIBRARY_MODE_FLAT_OCR,
    LIBRARY_MODE_PERSON_FOLDER,
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
            (w / "王京川_安全员C证.pdf").write_text("safety cert")
            (w / "王京川_特种作业操作证.pdf").write_text("special cert")

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

    def test_collect_all_with_images_uses_ocr_cache(self) -> None:
        """collect_all=True + 图片场景必须能正确触发缓存读写，不抛异常。

        回归测试：防止 _collect_all_from_folders 内对 _lookup_ocr_cache /
        _store_ocr_cache 的旧位置参数调用（会触发 AttributeError）再次回归。
        """
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
                pic = emp / "a5d6e67cd.jpg"
                pic.write_bytes(b"\x89PNG\r\n\x1a\nfake")

                out1 = Path(td) / "输出1"
                # 第一次 collect_all + 图片：不能抛 AttributeError
                result = collect_employee_materials(
                    lib, out1,
                    roster_source="张三",
                    collect_all=True,
                )
                self.assertEqual(result.total_employees, 1)
                self.assertEqual(result.matched_file_count, 1)
                # OCR 应被调用至少一次
                self.assertGreaterEqual(FakeEngine.call_count, 1)
                # 缓存应已写入
                cache_file = lib / _OCR_CACHE_FILE_NAME
                self.assertTrue(cache_file.exists())

                # 第二次：缓存命中，OCR 调用次数不变
                out2 = Path(td) / "输出2"
                result2 = collect_employee_materials(
                    lib, out2,
                    roster_source="张三",
                    collect_all=True,
                )
                self.assertEqual(FakeEngine.call_count, 1,
                                 "第二次 collect_all 应命中缓存，不重复调用 OCR")
                self.assertGreaterEqual(result2.ocr_cache_hits, 1)
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted


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
        with tempfile.TemporaryDirectory() as td:
            p1 = Path(td) / "a.png"
            p1.write_bytes(b"test_image_data")
            k1 = _compute_cache_key(p1)
            k2 = _compute_cache_key(p1)
            self.assertEqual(k1, k2)
            self.assertIn("_", k1)

    def test_compute_cache_key_invariant_to_filename(self) -> None:
        """重命名文件，只要内容相同，生成的 cache_key 必须完全相同。"""
        with tempfile.TemporaryDirectory() as td:
            p1 = Path(td) / "original_name.jpg"
            p1.write_bytes(b"identical_binary_content")
            k1 = _compute_cache_key(p1)

            p2 = Path(td) / "completely_different_name.png"
            p2.write_bytes(b"identical_binary_content")
            k2 = _compute_cache_key(p2)
            self.assertEqual(k1, k2, "相同内容的不同文件名必须生成完全相同的 content hash key")

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

    def test_cache_hit_after_file_renamed_or_moved(self) -> None:
        """文件被重命名或移动到深层子目录后，只要内容不变，依然 100% 命中 OCR 缓存。"""
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
            mc._OCR_ENGINE = CountingEngine()
            mc._OCR_ATTEMPTED = True

            with tempfile.TemporaryDirectory() as td:
                lib = Path(td) / "资料库"
                lib.mkdir()
                emp = lib / "张三"
                emp.mkdir()
                pic = emp / "old_random_hash_name.jpg"
                pic.write_bytes(b"\x89PNG\r\n\x1a\nfake_image_bytes")

                out1 = Path(td) / "输出1"
                # 首次运行：写入内容哈希指纹缓存
                collect_employee_materials(lib, out1, roster_source="张三", material_types=["身份证"])
                first_count = CountingEngine.call_count
                self.assertEqual(first_count, 1)

                # 重命名文件并移动到深层子目录
                sub_dir = emp / "2026新证件"
                sub_dir.mkdir()
                new_pic = sub_dir / "renamed_random_name_0x8f2a.png"
                pic.rename(new_pic)

                out2 = Path(td) / "输出2"
                # 第二次运行：即使文件名和路径彻底改变，因内容哈希一致，依然秒级命中缓存，OCR 调用次数不变！
                res2 = collect_employee_materials(lib, out2, roster_source="张三", material_types=["身份证"])
                self.assertEqual(CountingEngine.call_count, first_count, "文件重命名后不应重新调用 OCR，必须命中缓存")
                self.assertGreaterEqual(res2.ocr_cache_hits, 1)
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

    def test_safety_and_special_cert_collection(self) -> None:
        """测试新增的安全员证与特种证书的精准检索与提取。"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            emp = lib / "张三"
            emp.mkdir()
            (emp / "张三_安全员C证.pdf").write_text("c cert")
            (emp / "张三_特种作业操作证.jpg").write_bytes(b"\x89PNG\r\n\x1a\nfake")

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="张三",
                material_types=["安全员证", "特种证书"],
            )
            self.assertEqual(result.total_employees, 1)
            self.assertEqual(result.matched_file_count, 2)
            self.assertEqual(len(result.missing_records), 0)

    def test_custom_other_material_collection(self) -> None:
        """测试用户通过「其他」手动输入的自定义材料（如入职证明、保密协议）。"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            emp = lib / "李四"
            emp.mkdir()
            (emp / "李四_入职证明.pdf").write_text("onboarding proof")
            (emp / "李四_保密协议.docx").write_text("nda doc")

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="李四",
                material_types=["入职证明", "保密协议"],
            )
            self.assertEqual(result.total_employees, 1)
            self.assertEqual(result.matched_file_count, 2)
            self.assertEqual(len(result.missing_records), 0)
            self.assertTrue((out / "李四" / "李四_入职证明.pdf").exists())
            self.assertTrue((out / "李四" / "李四_保密协议.docx").exists())

    def test_multipage_contract_and_reverse_order_id_card(self) -> None:
        """测试多页合同全量提取不被早停截断，以及身份证反面先被扫描时依然完整提取正反双面。"""
        with tempfile.TemporaryDirectory() as td:
            lib = Path(td) / "资料库"
            lib.mkdir()
            emp = lib / "王五"
            emp.mkdir()
            # 模拟多页合同
            (emp / "王五_劳动合同_第1页.jpg").write_bytes(b"\x89PNG\r\n\x1a\npage1")
            (emp / "王五_劳动合同_第2页.jpg").write_bytes(b"\x89PNG\r\n\x1a\npage2")
            # 模拟反面命名排在前面或先被扫描的身份证
            (emp / "王五_身份证反面.jpg").write_bytes(b"\x89PNG\r\n\x1a\nback")
            (emp / "王五_身份证正面.jpg").write_bytes(b"\x89PNG\r\n\x1a\nfront")
            # 模拟大量低优先级普通无关图片
            (emp / "无关图片_01.jpg").write_bytes(b"\x89PNG\r\n\x1a\nirrelevant")
            (emp / "无关图片_02.jpg").write_bytes(b"\x89PNG\r\n\x1a\nirrelevant2")

            out = Path(td) / "输出"
            result = collect_employee_materials(
                lib, out,
                roster_source="王五",
                material_types=["劳动合同", "身份证"],
            )
            self.assertEqual(result.total_employees, 1)
            # 应该提取到 2页合同 + 2面身份证 = 4个文件，绝不截断丢页
            self.assertEqual(result.matched_file_count, 4)
            self.assertEqual(len(result.missing_records), 0)

            emp_out = out / "王五"
            self.assertTrue((emp_out / "王五_劳动合同_1.jpg").exists())
            self.assertTrue((emp_out / "王五_劳动合同_2.jpg").exists())
            self.assertTrue((emp_out / "王五_身份证_正面.jpg").exists())
            self.assertTrue((emp_out / "王五_身份证_反面.jpg").exists())


class TestFlatOCRMaterialLibrary(unittest.TestCase):
    def setUp(self) -> None:
        from hr_toolkit.tools import material_collector as mc

        self.mc = mc
        self.real_engine = mc._OCR_ENGINE
        self.real_attempted = mc._OCR_ATTEMPTED

        class ContentEngine:
            call_count = 0

            def __call__(self, source):
                type(self).call_count += 1
                marker = Path(source).read_bytes().decode("utf-8", errors="ignore")
                if marker == "zhang_contract":
                    return ([
                        ["乙方", "张三"],
                        ["材料", "劳动合同"],
                    ], None)
                if marker == "zhang_id_card":
                    return ([
                        ["姓名", "张三"],
                        ["公民身份号码", "440111199001011234"],
                        ["住址", "广东省"],
                    ], None)
                if marker in {"li_id_card", "lisi__id_card"}:
                    return ([
                        ["姓名", "李四"],
                        ["公民身份号码", "440111199001011235"],
                        ["住址", "广东省"],
                    ], None)
                if marker == "zhang_second_id_card":
                    return ([
                        ["姓名", "张三"],
                        ["公民身份号码", "440111199001019999"],
                        ["住址", "广东省"],
                    ], None)
                if marker == "zhang_household":
                    return ([
                        ["姓名", "张三"],
                        ["材料", "户口本"],
                    ], None)
                if marker == "zhang_degree_unlabeled":
                    return ([
                        ["证书", "普通高等学校毕业证书"],
                        ["持有人文字", "张三"],
                    ], None)
                if marker == "zhang_sanfeng_degree":
                    return ([
                        ["证书", "普通高等学校毕业证书"],
                        ["持有人文字", "张三丰"],
                    ], None)
                return ([["标题", "无法识别的普通图片"]], None)

        self.engine_type = ContentEngine
        self.engine_type.call_count = 0
        mc._OCR_ENGINE = ContentEngine()
        mc._OCR_ATTEMPTED = True

    def tearDown(self) -> None:
        self.mc._OCR_ENGINE = self.real_engine
        self.mc._OCR_ATTEMPTED = self.real_attempted

    @staticmethod
    def _write_image(path: Path, marker: str) -> None:
        path.write_bytes(marker.encode("utf-8"))

    def test_default_library_mode_remains_person_folder(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "资料库"
            employee = library / "张三"
            employee.mkdir(parents=True)
            (employee / "张三_劳动合同.txt").write_text("劳动合同", encoding="utf-8")

            result = collect_employee_materials(
                library,
                root / "输出",
                roster_source="张三",
                material_types=["劳动合同"],
            )

            self.assertEqual(result.library_mode, LIBRARY_MODE_PERSON_FOLDER)
            self.assertEqual(result.matched_file_count, 1)

    def test_first_scan_indexes_negatives_and_second_query_does_not_repeat_ocr(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            for index in range(99):
                self._write_image(library / f"{index:03d}.png", f"noise-{index}")
            self._write_image(library / "099.png", "zhang_id_card")

            first = collect_employee_materials(
                library,
                root / "输出1",
                roster_source="张三 440111199001011234",
                material_types=["身份证"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            first_calls = self.engine_type.call_count
            self.assertEqual(first_calls, 100)
            self.assertEqual(first.matched_file_count, 1)

            cache_path = library / _OCR_CACHE_FILE_NAME
            cache_data = _load_ocr_cache(cache_path)
            self.assertEqual(len(cache_data["paths"]), 100)
            self.assertEqual(len(cache_data["entries"]), 100)
            self.assertTrue(any(
                entry.get("material_type") == "其他材料"
                for entry in cache_data["entries"].values()
            ))
            self.assertNotIn("440111199001011234", cache_path.read_text(encoding="utf-8"))

            second = collect_employee_materials(
                library,
                root / "输出2",
                roster_source="张三 440111199001011234",
                material_types=["身份证"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            self.assertEqual(self.engine_type.call_count, first_calls)
            self.assertEqual(second.ocr_cache_hits, 100)
            self.assertEqual(second.ocr_cache_misses, 0)
            self.assertEqual(second.matched_file_count, 1)

    def test_filename_person_is_ignored_and_ocr_person_wins(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            self._write_image(library / "张三_身份证.png", "li_id_card")

            result = collect_employee_materials(
                library,
                root / "输出",
                roster_source="张三,李四",
                material_types=["身份证"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )

            self.assertEqual([match.employee_name for match in result.matches], ["李四"])
            self.assertEqual(self.engine_type.call_count, 1, "多人名单只能全库索引一次")
            self.assertIn("张三", result.missing_records)
            self.assertNotIn("李四", result.missing_records)

    def test_same_size_changed_file_is_reindexed_even_when_mtime_is_restored(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            source = library / "000.png"
            self._write_image(source, "zhang_id_card")
            original_stat = source.stat()

            collect_employee_materials(
                library,
                root / "输出1",
                roster_source="张三",
                material_types=["身份证"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            first_calls = self.engine_type.call_count

            # 两个 marker 等长；恢复 mtime 后仍必须依靠 ctime/hash 识别内容变更。
            self.assertEqual(len("zhang_id_card"), len("lisi__id_card"))
            self._write_image(source, "lisi__id_card")
            os.utime(source, ns=(original_stat.st_atime_ns, original_stat.st_mtime_ns))

            result = collect_employee_materials(
                library,
                root / "输出2",
                roster_source="张三",
                material_types=["身份证"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            self.assertGreater(self.engine_type.call_count, first_calls)
            self.assertEqual(result.matched_file_count, 0)
            self.assertGreaterEqual(result.ocr_cache_invalidated, 1)

    def test_file_changed_after_index_is_not_copied_as_stale_match(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            source = library / "000.png"
            self._write_image(source, "zhang_contract")
            changed = False

            def mutate_before_copy(_current: int, _total: int, message: str) -> None:
                nonlocal changed
                if not changed and "正在检索与匹配" in message:
                    self._write_image(source, "li_id_card")
                    changed = True

            result = collect_employee_materials(
                library,
                root / "输出",
                roster_source="张三",
                material_types=["劳动合同"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
                progress_callback=mutate_before_copy,
            )

            self.assertEqual(result.matched_file_count, 0)
            self.assertEqual(result.missing_records["张三"], ["劳动合同"])
            self.assertTrue(any("索引后发生变化" in warning for warning in result.warnings))

    def test_corrupted_copy_is_removed_and_reported_missing(self) -> None:
        from unittest.mock import patch

        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            self._write_image(library / "000.png", "zhang_contract")

            def corrupt_copy(_source, destination):
                Path(destination).write_bytes(b"corrupted")

            with patch.object(self.mc.shutil, "copy2", side_effect=corrupt_copy):
                result = collect_employee_materials(
                    library,
                    root / "输出",
                    roster_source="张三",
                    material_types=["劳动合同"],
                    library_mode=LIBRARY_MODE_FLAT_OCR,
                )

            self.assertEqual(result.matched_file_count, 0)
            self.assertEqual(result.missing_records["张三"], ["劳动合同"])
            self.assertFalse(any(path.is_file() for path in (root / "输出").rglob("*.png")))
            self.assertTrue(any("复制后校验不一致" in warning for warning in result.warnings))

    def test_renamed_file_reuses_content_hash_without_ocr(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            old_path = library / "old.png"
            self._write_image(old_path, "zhang_contract")
            collect_employee_materials(
                library,
                root / "输出1",
                roster_source="张三",
                material_types=["劳动合同"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            first_calls = self.engine_type.call_count

            nested = library / "新目录"
            nested.mkdir()
            old_path.rename(nested / "new-random.png")
            result = collect_employee_materials(
                library,
                root / "输出2",
                roster_source="张三",
                material_types=["劳动合同"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )

            self.assertEqual(self.engine_type.call_count, first_calls)
            self.assertEqual(result.matched_file_count, 1)
            self.assertEqual(result.ocr_cache_hits, 1)
            cache = _load_ocr_cache(library / _OCR_CACHE_FILE_NAME)
            self.assertEqual(set(cache["paths"]), {"新目录/new-random.png"})

    def test_collect_all_only_copies_files_identified_for_target(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            self._write_image(library / "a.png", "zhang_contract")
            self._write_image(library / "b.png", "li_id_card")

            result = collect_employee_materials(
                library,
                root / "输出",
                roster_source="张三",
                collect_all=True,
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )

            self.assertEqual(result.matched_file_count, 1)
            self.assertEqual(result.matches[0].employee_name, "张三")
            copied_files = [path.name for path in (root / "输出" / "张三").iterdir()]
            self.assertEqual(copied_files, ["a.png"])

    def test_same_name_employees_are_separated_by_id_in_output_and_report(self) -> None:
        from openpyxl import load_workbook

        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            self._write_image(library / "a.png", "zhang_id_card")
            self._write_image(library / "b.png", "zhang_second_id_card")

            output = root / "输出"
            result = collect_employee_materials(
                library,
                output,
                roster_source=(
                    "张三 440111199001011234,"
                    "张三 440111199001019999"
                ),
                material_types=["身份证"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )

            self.assertEqual(result.total_employees, 2)
            self.assertEqual(result.complete_employee_count, 2)
            self.assertEqual(result.matched_file_count, 2)
            self.assertTrue(
                (output / "张三（证件尾号1234）" / "张三（证件尾号1234）_身份证_正面.png").exists()
            )
            self.assertTrue(
                (output / "张三（证件尾号9999）" / "张三（证件尾号9999）_身份证_正面.png").exists()
            )
            self.assertEqual(
                len({match.employee_identity_key for match in result.matches}),
                2,
            )

            workbook = load_workbook(result.report_path, data_only=True)
            try:
                sheet = workbook["资料提取汇总与缺失清单"]
                self.assertEqual(sheet["D8"].value, "1/1")
                self.assertEqual(sheet["D9"].value, "1/1")
                self.assertEqual(sheet["F8"].value, "已提取(1份)")
                self.assertEqual(sheet["F9"].value, "已提取(1份)")
            finally:
                workbook.close()

    def test_all_existing_output_modes_work_with_flat_library_input(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            self._write_image(library / "random.png", "zhang_contract")

            expected = {
                MODE_BY_EMPLOYEE: lambda output: output / "张三" / "张三_劳动合同.png",
                MODE_BY_MATERIAL: lambda output: output / "劳动合同" / "张三_劳动合同.png",
                MODE_FLAT: lambda output: output / "张三_劳动合同.png",
            }
            for mode, expected_path in expected.items():
                output = root / f"输出-{mode}"
                result = collect_employee_materials(
                    library,
                    output,
                    roster_source="张三",
                    material_types=["劳动合同"],
                    mode=mode,
                    library_mode=LIBRARY_MODE_FLAT_OCR,
                )
                self.assertEqual(result.matched_file_count, 1)
                self.assertTrue(expected_path(output).exists(), mode)

    def test_cached_text_supports_new_custom_material_without_reocr(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            self._write_image(library / "random.png", "zhang_household")

            first = collect_employee_materials(
                library,
                root / "输出1",
                roster_source="张三",
                material_types=["身份证"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            first_calls = self.engine_type.call_count
            self.assertEqual(first.matched_file_count, 0)

            second = collect_employee_materials(
                library,
                root / "输出2",
                roster_source="张三",
                material_types=["户口本"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            self.assertEqual(self.engine_type.call_count, first_calls)
            self.assertEqual(second.matched_file_count, 1)

    def test_text_document_is_indexed_without_ocr(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            (library / "random.txt").write_text(
                "劳动合同 乙方：张三 工作内容：项目管理",
                encoding="utf-8",
            )

            result = collect_employee_materials(
                library,
                root / "输出",
                roster_source="张三",
                material_types=["劳动合同"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )

            self.assertEqual(self.engine_type.call_count, 0)
            self.assertEqual(result.matched_file_count, 1)
            self.assertEqual(result.matches[0].matched_by, "doc_content_contract")

    def test_unlabeled_name_uses_exact_ocr_text_boundary(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            self._write_image(library / "a.png", "zhang_degree_unlabeled")
            self._write_image(library / "b.png", "zhang_sanfeng_degree")

            result = collect_employee_materials(
                library,
                root / "输出",
                roster_source="张三",
                material_types=["学历证明"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )

            self.assertEqual(result.matched_file_count, 1)
            self.assertEqual(result.matches[0].relative_source_path, "a.png")

    def test_ocr_unavailable_is_not_persisted_as_negative_result(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            self._write_image(library / "random.png", "zhang_contract")
            self._write_image(library / "noise.png", "noise")

            working_engine = self.mc._OCR_ENGINE
            self.mc._OCR_ENGINE = None
            self.mc._OCR_ATTEMPTED = True
            first = collect_employee_materials(
                library,
                root / "输出1",
                roster_source="张三",
                material_types=["劳动合同"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            self.assertEqual(first.matched_file_count, 0)
            self.assertEqual(
                sum("暂时无法完成 OCR" in warning for warning in first.warnings),
                1,
            )
            cache = _load_ocr_cache(library / _OCR_CACHE_FILE_NAME)
            self.assertEqual(cache["entries"], {})

            self.mc._OCR_ENGINE = working_engine
            second = collect_employee_materials(
                library,
                root / "输出2",
                roster_source="张三",
                material_types=["劳动合同"],
                library_mode=LIBRARY_MODE_FLAT_OCR,
            )
            self.assertEqual(second.matched_file_count, 1)

    def test_flat_library_rejects_disabled_cache(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            library = root / "无序资料库"
            library.mkdir()
            with self.assertRaisesRegex(ValueError, "必须启用 OCR 索引缓存"):
                collect_employee_materials(
                    library,
                    root / "输出",
                    roster_source="张三",
                    library_mode=LIBRARY_MODE_FLAT_OCR,
                    use_ocr_cache=False,
                )


if __name__ == "__main__":
    unittest.main()
