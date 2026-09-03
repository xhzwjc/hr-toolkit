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


def _fake_folder_rename(root_dir, *, mode, cancelled=None, progress_callback=None):
    target = Path(root_dir) / "张三-已核对"
    (Path(root_dir) / "张三").rename(target)
    return {"count": 1, "output_dir": str(root_dir)}


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

    def test_coordinator_folder_rename_operates_on_project_copy_and_preserves_customer_source(self) -> None:
        from hr_toolkit.project_store import CATEGORY_RESULTS

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            customer_source = root / "人员资料"
            person_folder = customer_source / "张三"
            person_folder.mkdir(parents=True)
            (person_folder / "说明.txt").write_text("record", encoding="utf-8")

            project_root = root / "project"
            store = ProjectStore.create(project_root, "测试项目")
            coordinator = ProjectRunCoordinator()
            finished = threading.Event()
            success = []
            errors = []

            request = RunRequest(
                tool_id="folder_rename",
                tool_name="资料文件夹改名",
                group_name="人员与档案",
                description="改名测试",
                function=_fake_folder_rename,
                args=(customer_source,),
                kwargs={"mode": "append"},
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

                # Customer source folder MUST REMAIN UNTOUCHED
                self.assertTrue((customer_source / "张三" / "说明.txt").is_file())
                self.assertFalse((customer_source / "张三-已核对").exists())

                # Project results copy MUST HAVE THE RENAMED FOLDER
                batches = store.list_batches()
                self.assertEqual(len(batches), 1)
                batch_detail = store.get_batch(batches[0].id)
                assert batch_detail is not None
                results_dir = batch_detail.directories[CATEGORY_RESULTS]
                found = list(results_dir.glob("**/张三-已核对"))
                self.assertTrue(len(found) > 0, f"Expected renamed folder in results: {list(results_dir.rglob('*'))}")
            finally:
                store.close()


if __name__ == "__main__":
    unittest.main()
