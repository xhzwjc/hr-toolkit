from __future__ import annotations

import hashlib
import unittest
from pathlib import Path

from hr_toolkit.tools import material_collector as mc


class MaterialDocumentGroupsTest(unittest.TestCase):
    def page(self, tag: str, text: str, person: str = "") -> mc._FlatIndexedFile:
        name = hashlib.sha256(tag.encode()).hexdigest()[:32] + ".png"
        path = Path("library") / name
        return mc._FlatIndexedFile(
            path, name, tag, "劳动合同" if "劳动合同" in text else "其他材料",
            "ocr", "", (person,) if person else (), "", text_snippet=text,
        )

    def test_unordered_contract_keeps_continuations_without_identity(self):
        first = self.page("first", "劳动合同 第一条 合同期限", "张三")
        second = self.page("second", "第三条 工作时间 第六条 劳动保护")
        third = self.page("third", "第九条 劳动合同变更")
        last = self.page("last", "第十二条 劳动争议 乙方（签名）", "张三")
        result = mc._enrich_flat_index_with_document_groups([third, first, last, second], [], ["劳动合同"])
        self.assertEqual([item.source_path for item in result], [item.source_path for item in (first, second, third, last)])
        self.assertTrue(all(item.document_group_id for item in result))
        self.assertTrue(all("完整性待确认" in item.document_warning for item in result))

    def test_competing_contracts_do_not_share_anonymous_pages(self):
        pages = [
            self.page("a", "劳动合同 第一条 合同期限", "张三"),
            self.page("b", "劳动合同 第一条 合同期限", "李四"),
            self.page("middle", "第三条 工作时间"),
            self.page("end_a", "第十二条 劳动争议 乙方（签名）", "张三"),
            self.page("end_b", "第十二条 劳动争议 乙方（签名）", "李四"),
        ]
        result = mc._enrich_flat_index_with_document_groups(pages, [], ["劳动合同"])
        middle = next(item for item in result if item.source_path == pages[2].source_path)
        self.assertFalse(middle.document_group_id)
        self.assertFalse(middle.extracted_names)
        self.assertIn("待确认", middle.document_warning)

    def test_portrait_filename_and_employee_folder_use_full_names(self):
        self.assertEqual(mc._portrait_filename_identity("张三_证件照.png")[0], ("张三",))
        self.assertEqual(mc._portrait_filename_identity("证件照.png"), ((), "", ""))
        self.assertIsNone(mc._match_folder_to_employee("张三丰", mc.TargetEmployee("张三")))
        self.assertEqual(mc._match_folder_to_employee("张三资料", mc.TargetEmployee("张三")), "name")

    def test_work_schedule_without_contract_clause_is_not_absorbed(self):
        first = self.page("first", "劳动合同 第1页 共2页", "张三")
        last = self.page("last", "乙方签字 第2页 共2页", "张三")
        unrelated = self.page("schedule", "工作时间：9:00-18:00 公司活动安排")
        result = mc._enrich_flat_index_with_document_groups([first, unrelated, last], [], ["劳动合同"])
        schedule = next(item for item in result if item.source_path == unrelated.source_path)
        self.assertIs(schedule, unrelated)
        self.assertFalse(schedule.document_group_id)
        self.assertFalse(schedule.document_warning)
        self.assertTrue(all(item.document_group_id for item in result if item is not schedule))

    def test_footer_and_random_filename_are_not_scan_sequence(self):
        self.assertEqual(mc._footer_page_marker("第一条 工作内容\n4"), (4, None))
        self.assertIsNone(mc._footer_page_marker("签署日期\n2026年9月5日"))
        self.assertIsNone(mc._filename_page_sequence(Path("390a48679aa88a83ca0bbe5b491d8820.png")))

    def test_large_ambiguous_block_honors_cancellation(self):
        pages = [self.page(str(index), "劳动合同 第一条 工作内容") for index in range(300)]
        with self.assertRaises(mc.MaterialCollectionCancelled):
            mc._enrich_flat_index_with_document_groups(pages, [], ["劳动合同"], cancelled=lambda: True)
