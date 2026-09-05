from __future__ import annotations

import unittest
from unittest.mock import patch

from hr_toolkit.tools import material_collector as mc
from hr_toolkit.tools.material_progress import MaterialProgress


class MaterialOCRProgressTest(unittest.TestCase):
    def test_only_finished_files_advance_and_pdf_details_keep_file_total(self):
        events = []
        progress = MaterialProgress(lambda c, t, message: events.append((c, t, message)))
        progress.begin("识别资料", 3, "个文件")
        items = progress.items(["a", "b", "c"], str)
        self.assertEqual(next(items), "a")
        self.assertEqual(progress.current, 0)
        progress.detail(8, 8, "正在处理 PDF 第 8/8 页")
        self.assertEqual((progress.current, progress.total), (0, 3))
        self.assertEqual(next(items), "b")
        self.assertEqual(progress.current, 1)
        items.close()
        self.assertEqual(progress.current, 1)
        self.assertFalse(any(c == t == 3 for c, t, _ in events))

    def test_last_completed_item_emits_exact_total(self):
        events = []
        progress = MaterialProgress(lambda c, t, text: events.append((c, t)))
        progress.begin("提取资料", 1000, "个文件")
        for _item in progress.items(range(1000), str):
            pass
        self.assertEqual(events[-1], (1000, 1000))

    def test_low_memory_stops_without_waiting_or_advancing(self):
        with patch.object(mc, "_available_ocr_memory", return_value=511 * 1024**2):
            with self.assertRaises(mc.OCRMemoryPressureError):
                mc._ensure_ocr_memory_headroom()

    def test_thread_budget_keeps_recognition_batch_unchanged(self):
        with patch.object(mc.os, "cpu_count", return_value=4):
            self.assertEqual(mc._ocr_runtime_options(), {"intra_op_num_threads": 2, "inter_op_num_threads": 1})
        with patch.object(mc.os, "cpu_count", return_value=2):
            self.assertEqual(mc._ocr_runtime_options()["intra_op_num_threads"], 1)
