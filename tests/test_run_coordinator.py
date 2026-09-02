from __future__ import annotations

import tempfile
import threading
import unittest
from pathlib import Path

from hr_toolkit.project_store import ProjectStore
from hr_toolkit.run_coordinator import (
    PROGRESS_UI_INTERVAL_SECONDS,
    ProjectRunCoordinator,
    RunCallbacks,
    RunRequest,
)


def _copy_probe(
    input_path,
    output_dir,
    *,
    cancelled=None,
    progress_callback=None,
):
    if cancelled is not None and cancelled():
        raise RuntimeError("cancelled")
    output = Path(output_dir) / "same.txt"
    output.write_bytes(Path(input_path).read_bytes())
    if progress_callback is not None:
        progress_callback(1, 1, "完成")
    return {"output_file": str(output), "value": Path(input_path).read_text(encoding="utf-8")}


class ProjectRunCoordinatorTests(unittest.TestCase):
    def test_business_progress_is_coalesced_without_losing_completion(self) -> None:
        def noisy_probe(*, progress_callback=None):
            for current in range(1, 1001):
                if progress_callback is not None:
                    progress_callback(current, 1000, f"处理 {current}")
            return {"count": 1000}

        progress = []
        request = RunRequest(
            tool_id="probe",
            tool_name="进度测试",
            group_name="测试",
            description="进度测试",
            function=noisy_probe,
            args=(),
            kwargs={},
        )
        payload, isolated = ProjectRunCoordinator._business_call(
            request,
            (),
            {},
            threading.Event(),
            RunCallbacks(progress=lambda *values: progress.append(values)),
        )

        self.assertEqual(PROGRESS_UI_INTERVAL_SECONDS, 0.1)
        self.assertEqual(payload, {"count": 1000})
        self.assertFalse(isolated)
        self.assertGreaterEqual(len(progress), 1)
        self.assertLessEqual(len(progress), 2)
        self.assertEqual(progress[-1], (1000, 1000, "处理 1000"))

    def test_project_snapshot_business_output_and_result_registration(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.txt"
            source.write_text("unchanged-business-value", encoding="utf-8")
            project_root = root / "project"
            store = ProjectStore.create(project_root, "测试项目")
            coordinator = ProjectRunCoordinator()
            finished = threading.Event()
            success = []
            errors = []
            request = RunRequest(
                tool_id="probe",
                tool_name="测试工具",
                group_name="测试",
                description="测试工具",
                function=_copy_probe,
                args=(source, project_root),
                kwargs={},
            )
            callbacks = RunCallbacks(
                success=lambda *payload: success.append(payload),
                error=lambda error: errors.append(error),
                finished=finished.set,
            )
            try:
                self.assertTrue(coordinator.start(store, request, callbacks))
                self.assertTrue(finished.wait(10))
                self.assertFalse(errors)
                self.assertEqual(len(success), 1)
                payload, result_dir, _elapsed, _isolated = success[0]
                self.assertEqual(payload["value"], "unchanged-business-value")
                self.assertEqual((Path(result_dir) / "same.txt").read_text(encoding="utf-8"), "unchanged-business-value")
                summaries = store.list_batches()
                self.assertEqual(len(summaries), 1)
                self.assertEqual(summaries[0].status, "success")
                self.assertEqual(source.read_text(encoding="utf-8"), "unchanged-business-value")
            finally:
                store.close()


if __name__ == "__main__":
    unittest.main()
