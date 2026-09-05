from __future__ import annotations

import unittest
from unittest.mock import patch

from hr_toolkit.tools import material_collector as mc


class MaterialOCRProgressTest(unittest.TestCase):
    def test_low_memory_stops_without_waiting_or_advancing(self):
        with patch.object(mc, "_available_ocr_memory", return_value=511 * 1024**2):
            with self.assertRaises(mc.OCRMemoryPressureError):
                mc._ensure_ocr_memory_headroom()

    def test_thread_budget_keeps_recognition_batch_unchanged(self):
        with patch.object(mc.os, "cpu_count", return_value=4):
            self.assertEqual(mc._ocr_runtime_options(), {"intra_op_num_threads": 2, "inter_op_num_threads": 1})
        with patch.object(mc.os, "cpu_count", return_value=2):
            self.assertEqual(mc._ocr_runtime_options()["intra_op_num_threads"], 1)
