from __future__ import annotations

import tempfile
import threading
import unittest
from pathlib import Path

from hr_toolkit.background_process import (
    PROCESS_FILE_THRESHOLD_BYTES,
    run_business_process,
    should_use_process,
)


class BackgroundProcessTests(unittest.TestCase):
    def test_load_policy_is_bounded_and_does_not_scan_directories(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            tiny = root / "tiny.xlsx"
            tiny.write_bytes(b"x")
            large = root / "large.xlsx"
            with large.open("wb") as handle:
                handle.truncate(PROCESS_FILE_THRESHOLD_BYTES)
            self.assertFalse(should_use_process("salary_split", (tiny, root / "out"), {}))
            self.assertTrue(should_use_process("salary_split", (large, root / "out"), {}))
            self.assertTrue(should_use_process("salary_merge", ([tiny] * 5, root / "out"), {}))
            self.assertTrue(should_use_process("salary_merge", (root, root / "out"), {}))
            self.assertTrue(should_use_process("material_collector", (tiny, root / "out"), {}))

    def test_spawned_call_returns_serializable_payload_and_progress(self) -> None:
        progress = []
        result = run_business_process(
            module_name="hr_toolkit.background_process",
            function_name="_process_smoke_probe",
            args=(Path("input.xlsx"), Path("output")),
            kwargs={},
            cancel_event=threading.Event(),
            on_progress=lambda current, total, message: progress.append((current, total, message)),
        )
        self.assertEqual(result.payload["input_path"], "input.xlsx")
        self.assertEqual(result.payload["output_dir"], "output")
        self.assertGreaterEqual(result.elapsed_seconds, 0.0)
        self.assertIn((1, 1, "完成"), progress)


if __name__ == "__main__":
    unittest.main()
