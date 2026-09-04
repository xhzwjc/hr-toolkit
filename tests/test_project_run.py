from __future__ import annotations

import hashlib
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

from hr_toolkit.history_store import SourceSpec
from hr_toolkit.project_run import import_project_run_sources
from hr_toolkit.project_store import ProjectStore


class ProjectRunSourceImportTests(unittest.TestCase):
    def test_hundred_selected_files_use_one_streaming_project_transaction(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source_root = root / "sources"
            source_root.mkdir()
            sources = []
            source_hashes = {}
            for index in range(100):
                source = source_root / f"社保清单-{index:03d}.xlsx"
                source.write_bytes((f"row-{index}\n" * 32).encode("utf-8"))
                sources.append(source)
                source_hashes[source] = hashlib.sha256(source.read_bytes()).hexdigest()

            store = ProjectStore.create(root / "project", "百文件事务测试")
            draft = store.create_draft(
                group_name="社保与保险",
                tool_id="social_security",
                tool_name="社保明细与汇总",
                business_description="百文件事务测试",
                business_period="2026-09-04",
            )
            progress = []
            specs = [SourceSpec(path=path, role="input_path") for path in sources]
            try:
                with patch.object(
                    store,
                    "import_sources",
                    wraps=store.import_sources,
                ) as import_sources:
                    replacements = import_project_run_sources(
                        store,
                        draft.summary.id,
                        specs,
                        threading.Event(),
                        on_progress=progress.append,
                    )

                import_sources.assert_called_once()
                self.assertEqual(len(replacements["input_path"]), 100)
                self.assertEqual(
                    [path.name for path in replacements["input_path"]],
                    [path.name for path in sources],
                )
                self.assertEqual(len(store.get_batch(draft.summary.id).files), 100)
                self.assertTrue(
                    any(
                        event.phase == "finalizing"
                        and event.files_completed == 100
                        and event.files_total == 100
                        for event in progress
                    )
                )
                self.assertEqual(
                    source_hashes,
                    {
                        path: hashlib.sha256(path.read_bytes()).hexdigest()
                        for path in sources
                    },
                )
            finally:
                store.close()

    def test_batched_duplicate_names_keep_existing_collision_order(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first = root / "first" / "工资表.xlsx"
            second = root / "second" / "工资表.xlsx"
            first.parent.mkdir()
            second.parent.mkdir()
            first.write_bytes(b"first")
            second.write_bytes(b"second")

            store = ProjectStore.create(root / "project", "重名文件测试")
            draft = store.create_draft(
                group_name="薪酬管理",
                tool_id="salary_merge",
                tool_name="多月工资合并",
                business_description="重名文件测试",
                business_period="2026-09-04",
            )
            try:
                replacements = import_project_run_sources(
                    store,
                    draft.summary.id,
                    [
                        SourceSpec(path=first, role="input_path"),
                        SourceSpec(path=second, role="input_path"),
                    ],
                    threading.Event(),
                )["input_path"]

                self.assertEqual(
                    [path.name for path in replacements],
                    ["工资表.xlsx", "工资表 (2).xlsx"],
                )
                self.assertEqual(replacements[0].read_bytes(), b"first")
                self.assertEqual(replacements[1].read_bytes(), b"second")
            finally:
                store.close()


if __name__ == "__main__":
    unittest.main()
