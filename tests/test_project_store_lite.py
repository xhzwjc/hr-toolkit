from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from hr_toolkit.project_store_lite import ProjectStoreLite, ProjectStoreLiteError


class ProjectStoreLiteTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_create_and_open_project(self) -> None:
        proj_root = self.tmp_dir / "TestProject"
        store = ProjectStoreLite.create(proj_root, "测试项目")
        self.assertEqual(store.name, "测试项目")
        self.assertTrue((proj_root / ".toolkit" / "project.json").is_file())

        opened = ProjectStoreLite.open(proj_root)
        self.assertEqual(opened.name, "测试项目")
        self.assertEqual(opened.project_id, store.project_id)

    def test_create_requires_empty_dir(self) -> None:
        proj_root = self.tmp_dir / "NonEmpty"
        proj_root.mkdir()
        (proj_root / "dummy.txt").write_text("hello")
        with self.assertRaises(ProjectStoreLiteError):
            ProjectStoreLite.create(proj_root, "测试")

    def test_import_file_and_record_batch(self) -> None:
        proj_root = self.tmp_dir / "BatchProject"
        store = ProjectStoreLite.create(proj_root, "批次项目")

        dummy_input = self.tmp_dir / "input.xlsx"
        dummy_input.write_bytes(b"dummy excel content")

        stored_in = store.import_file(dummy_input)
        self.assertEqual(stored_in.display_name, "input.xlsx")
        self.assertTrue((proj_root / stored_in.relative_path).is_file())

        res_path = store.prepare_result_path("output.xlsx")
        res_path.write_bytes(b"result content")
        stored_out = store.import_file(res_path, category="results")

        record = store.record_batch(
            tool_id="salary_split",
            tool_name="工资表拆分",
            input_files=[stored_in],
            output_files=[stored_out],
            summary={"count": 5},
        )
        self.assertEqual(record.status, "success")

        batches = store.list_batches()
        self.assertEqual(len(batches), 1)
        self.assertEqual(batches[0]["tool_id"], "salary_split")
        self.assertEqual(batches[0]["summary"]["count"], 5)


if __name__ == "__main__":
    unittest.main()
