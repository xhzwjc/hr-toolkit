from __future__ import annotations

import hashlib
import json
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unittest
import uuid
from contextlib import closing
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

from hr_toolkit.history_store import (
    DATA_DIR_ENV,
    SCHEMA_SQL,
    SCHEMA_VERSION,
    SQLITE_APPLICATION_ID,
    HistoryStore,
    HistoryStoreError,
    SourceSpec,
    default_history_root,
)


class HistoryStoreTests(unittest.TestCase):
    def setUp(self) -> None:
        self._temporary = tempfile.TemporaryDirectory()
        self.base = Path(self._temporary.name)
        self.store = HistoryStore(self.base / "library")
        self.source_dir = self.base / "source"
        self.source_dir.mkdir()

    def tearDown(self) -> None:
        self._temporary.cleanup()

    def _start(self, name: str = "工资表拆分") -> str:
        return self.store.start_task(
            tool_id="salary_split",
            tool_name=name,
            app_version="0.2.4",
            parameters={"demo": True},
        )

    def _open_store_in_processes(self, root: Path, count: int = 8) -> list[subprocess.CompletedProcess[str]]:
        command = [
            sys.executable,
            "-c",
            "from hr_toolkit.history_store import HistoryStore; import sys; "
            "HistoryStore(sys.argv[1]); print('ok')",
            str(root),
        ]
        processes = [
            subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            for _ in range(count)
        ]
        results: list[subprocess.CompletedProcess[str]] = []
        for process in processes:
            stdout, stderr = process.communicate(timeout=20)
            results.append(subprocess.CompletedProcess(command, process.returncode, stdout, stderr))
        return results

    def test_complete_lifecycle_persists_inputs_outputs_and_manifest(self) -> None:
        original = self.source_dir / "7月工资.xlsx"
        original.write_bytes(b"salary-data")
        result_dir = self.base / "result"
        result_dir.mkdir()
        output = result_dir / "公司A-工资表.xlsx"
        output.write_bytes(b"output-data")

        task_id = self._start()
        inputs = self.store.archive_sources(task_id, [SourceSpec(original, role="input_path")])
        outputs = self.store.archive_output_directory(task_id, result_dir)
        self.assertTrue(self.store.mark_success(task_id, {"employee_count": 3}))

        self.assertEqual(len(inputs), 1)
        self.assertEqual(len(outputs), 1)
        detail = self.store.get_task(task_id)
        self.assertIsNotNone(detail)
        assert detail is not None
        self.assertEqual(detail.summary.status, "success")
        self.assertEqual(detail.summary.input_names, ("7月工资.xlsx",))
        self.assertEqual(detail.summary.output_names, ("公司A-工资表.xlsx",))
        self.assertEqual(detail.inputs[0].archived_path.read_bytes(), original.read_bytes())
        self.assertEqual(detail.outputs[0].archived_path.read_bytes(), output.read_bytes())
        self.assertEqual(detail.inputs[0].sha256, hashlib.sha256(original.read_bytes()).hexdigest())
        self.assertEqual(detail.result, {"employee_count": 3})

        manifest = json.loads((detail.task_dir / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["task"]["status"], "success")
        self.assertEqual(len(manifest["files"]), 2)
        self.assertFalse((detail.task_dir / "manifest.json.tmp").exists())

        reopened = HistoryStore(self.store.root)
        reopened_detail = reopened.get_task(task_id)
        self.assertIsNotNone(reopened_detail)
        assert reopened_detail is not None
        self.assertEqual(reopened_detail.summary.status, "success")
        self.assertTrue(reopened.integrity_check())

    def test_list_is_paginated_and_searches_file_names(self) -> None:
        for index in range(4):
            path = self.source_dir / f"工资-{index}.xlsx"
            path.write_bytes(str(index).encode())
            task_id = self._start()
            self.store.archive_sources(task_id, [SourceSpec(path, role="input_paths")])
            self.store.mark_success(task_id)

        first_page, total = self.store.list_tasks(limit=2)
        second_page, second_total = self.store.list_tasks(limit=2, offset=2)
        matching, matching_total = self.store.list_tasks(search="工资-2")
        self.assertEqual(total, 4)
        self.assertEqual(second_total, 4)
        self.assertEqual(len(first_page), 2)
        self.assertEqual(len(second_page), 2)
        self.assertEqual(matching_total, 1)
        self.assertEqual(matching[0].input_names, ("工资-2.xlsx",))

    def test_directory_archive_filters_suffixes_and_skips_links(self) -> None:
        nested = self.source_dir / "项目" / "本月"
        nested.mkdir(parents=True)
        excel = nested / "异动.xlsx"
        excel.write_bytes(b"excel")
        archive_names = (
            "资料.zip",
            "资料.rar",
            "资料.7z",
            "资料.tar",
            "资料.tar.gz",
            "资料.tgz",
            "资料.tar.bz2",
            "资料.tbz2",
            "资料.tar.xz",
            "资料.txz",
        )
        for archive_name in archive_names:
            (nested / archive_name).write_bytes(archive_name.encode("utf-8"))
        (nested / "单文件.gz").write_bytes(b"not supported")
        (nested / "说明.txt").write_text("not archived", encoding="utf-8")
        link = nested / "链接.xlsx"
        try:
            link.symlink_to(excel)
        except (OSError, NotImplementedError):
            link = None

        task_id = self._start("异动汇总")
        records = self.store.archive_sources(task_id, [SourceSpec(self.source_dir, role="input_paths")])
        self.assertEqual(
            {record.display_name for record in records},
            {"异动.xlsx", *archive_names},
        )
        self.assertTrue(all("项目" in record.relative_path for record in records))
        self.assertNotIn("单文件.gz", {record.display_name for record in records})
        if link is not None:
            self.assertNotIn("链接.xlsx", {record.display_name for record in records})

    def test_duplicate_names_never_overwrite(self) -> None:
        first_dir = self.source_dir / "a"
        second_dir = self.source_dir / "b"
        first_dir.mkdir()
        second_dir.mkdir()
        first = first_dir / "同名.xlsx"
        second = second_dir / "同名.xlsx"
        first.write_bytes(b"first")
        second.write_bytes(b"second")
        task_id = self._start()

        records = self.store.archive_sources(
            task_id,
            [SourceSpec(first, role="input_paths"), SourceSpec(second, role="input_paths")],
        )
        self.assertEqual(len(records), 2)
        self.assertNotEqual(records[0].archived_path, records[1].archived_path)
        self.assertEqual({record.archived_path.read_bytes() for record in records}, {b"first", b"second"})

    def test_stopped_status_is_not_overwritten_by_late_success(self) -> None:
        result_dir = self.base / "late-result"
        result_dir.mkdir()
        (result_dir / "结果.xlsx").write_bytes(b"late")
        task_id = self._start()
        self.assertTrue(self.store.mark_stopped(task_id))
        with self.assertRaisesRegex(HistoryStoreError, "已经结束"):
            self.store.archive_output_directory(task_id, result_dir)
        self.assertFalse(self.store.mark_success(task_id))
        detail = self.store.get_task(task_id)
        assert detail is not None
        self.assertEqual(detail.summary.status, "stopped")
        self.assertEqual(detail.summary.output_names, ())

    def test_terminal_manifest_failure_does_not_leave_false_success(self) -> None:
        task_id = self._start()
        with patch.object(self.store, "_write_manifest", side_effect=OSError("manifest full")):
            with self.assertRaisesRegex(OSError, "manifest full"):
                self.store.mark_success(task_id, {"count": 1})
        detail = self.store.get_task(task_id)
        assert detail is not None
        self.assertEqual(detail.summary.status, "running")
        self.assertEqual(detail.result, {})
        self.assertTrue(self.store.mark_failed(task_id, "留存失败"))

    def test_running_tasks_are_recovered_as_failed_after_restart(self) -> None:
        task_id = self._start()
        with patch("hr_toolkit.history_store._pid_is_alive", return_value=False):
            reopened = HistoryStore(self.store.root)
        detail = reopened.get_task(task_id)
        assert detail is not None
        self.assertEqual(detail.summary.status, "failed")
        self.assertIn("意外关闭", detail.summary.error_message or "")

    def test_second_live_instance_does_not_interrupt_active_task(self) -> None:
        task_id = self._start()
        reopened = HistoryStore(self.store.root)
        detail = reopened.get_task(task_id)
        assert detail is not None
        self.assertEqual(detail.summary.status, "running")

    def test_future_database_version_is_rejected_without_rewrite(self) -> None:
        with closing(sqlite3.connect(self.store.database_path)) as connection:
            connection.execute("PRAGMA journal_mode = DELETE")
            connection.execute("PRAGMA user_version = 99")
            connection.commit()
        before = self.store.database_path.read_bytes()
        with self.assertRaisesRegex(HistoryStoreError, "版本不兼容"):
            HistoryStore(self.store.root)
        self.assertEqual(self.store.database_path.read_bytes(), before)
        with closing(sqlite3.connect(self.store.database_path)) as connection:
            self.assertEqual(connection.execute("PRAGMA user_version").fetchone()[0], 99)

    @unittest.skipIf(os.name == "nt", "symlink creation may require elevated Windows privileges")
    def test_database_symlink_is_rejected_without_touching_target(self) -> None:
        for suffix in ("", "-wal", "-shm", "-journal"):
            (self.store.root / f"history.db{suffix}").unlink(missing_ok=True)
        outside = self.base / "outside.db"
        outside.write_bytes(b"keep-outside")
        self.store.database_path.symlink_to(outside)
        with self.assertRaisesRegex(HistoryStoreError, "不能是链接"):
            HistoryStore(self.store.root)
        self.assertEqual(outside.read_bytes(), b"keep-outside")

    def test_move_to_trash_is_recoverable_and_hides_record(self) -> None:
        task_id = self._start()
        self.store.mark_success(task_id)
        trash_path = self.store.move_to_trash(task_id)
        tasks, total = self.store.list_tasks()
        self.assertEqual(tasks, ())
        self.assertEqual(total, 0)
        self.assertTrue(trash_path.is_dir())
        self.assertTrue((trash_path / "manifest.json").is_file())
        manifest = json.loads((trash_path / "manifest.json").read_text(encoding="utf-8"))
        self.assertTrue(all(item["relative_path"].startswith("trash/") for item in manifest["files"]))
        self.assertIsNotNone(manifest["task"]["deleted_at"])

    def test_move_to_trash_manifest_failure_rolls_back_index_and_directory(self) -> None:
        task_id = self._start()
        self.store.mark_success(task_id)
        before = self.store.get_task(task_id)
        assert before is not None
        with patch.object(self.store, "_write_manifest", side_effect=OSError("disk full")):
            with self.assertRaisesRegex(OSError, "disk full"):
                self.store.move_to_trash(task_id)

        after = self.store.get_task(task_id)
        assert after is not None
        self.assertEqual(after.task_dir, before.task_dir)
        self.assertIsNone(after.summary.deleted_at)
        self.assertTrue(after.task_dir.is_dir())
        self.assertEqual(self.store.list_tasks()[1], 1)
        self.assertFalse(any(self.store.root.glob(".trash-move-*.json")))

    def test_pending_trash_move_is_rolled_back_after_crash(self) -> None:
        task_id = self._start()
        self.store.mark_success(task_id)
        detail = self.store.get_task(task_id)
        assert detail is not None
        target = self.store.trash_dir / detail.task_dir.name
        marker = self.store._trash_move_marker_path(task_id)
        marker.write_text(
            json.dumps(
                {
                    "schema_version": SCHEMA_VERSION,
                    "task_id": task_id,
                    "old_relpath": detail.task_dir.relative_to(self.store.root).as_posix(),
                    "new_relpath": target.relative_to(self.store.root).as_posix(),
                }
            ),
            encoding="utf-8",
        )
        detail.task_dir.replace(target)

        reopened = HistoryStore(self.store.root)
        recovered = reopened.get_task(task_id)
        assert recovered is not None
        self.assertTrue(recovered.task_dir.is_dir())
        self.assertTrue(recovered.task_dir.is_relative_to(reopened.records_dir))
        self.assertFalse(marker.exists())

    def test_pending_trash_move_is_normalized_before_database_rebuild(self) -> None:
        task_id = self._start()
        self.store.mark_success(task_id)
        detail = self.store.get_task(task_id)
        assert detail is not None
        target = self.store.trash_dir / detail.task_dir.name
        marker = self.store._trash_move_marker_path(task_id)
        marker.write_text(
            json.dumps(
                {
                    "schema_version": SCHEMA_VERSION,
                    "task_id": task_id,
                    "old_relpath": detail.task_dir.relative_to(self.store.root).as_posix(),
                    "new_relpath": target.relative_to(self.store.root).as_posix(),
                }
            ),
            encoding="utf-8",
        )
        detail.task_dir.replace(target)
        for suffix in ("-wal", "-shm", "-journal"):
            self.store.database_path.with_name(self.store.database_path.name + suffix).unlink(missing_ok=True)
        self.store.database_path.write_bytes(b"broken-during-trash-move")

        reopened = HistoryStore(self.store.root)
        recovered = reopened.get_task(task_id)
        assert recovered is not None
        self.assertTrue(recovered.task_dir.is_relative_to(reopened.records_dir))
        self.assertFalse(marker.exists())

    def test_move_to_trash_rejects_broad_or_malformed_task_directory(self) -> None:
        normal_task = self._start()
        self.store.mark_success(normal_task)
        normal_detail = self.store.get_task(normal_task)
        assert normal_detail is not None
        rogue_id = "f" * 32
        with self.store._connect() as connection:
            connection.execute(
                """
                INSERT INTO tasks (
                    id, tool_id, tool_name, app_version, status, started_at,
                    finished_at, task_relpath
                ) VALUES (?, 'test', '伪造任务', '0.2.4', 'success', ?, ?, 'records')
                """,
                (rogue_id, "2026-01-01T00:00:00+00:00", "2026-01-01T00:00:01+00:00"),
            )
        with self.assertRaisesRegex(HistoryStoreError, "目录无效"):
            self.store.move_to_trash(rogue_id)
        self.assertTrue(normal_detail.task_dir.is_dir())

    def test_rebuild_index_recovers_records_from_manifests(self) -> None:
        original = self.source_dir / "源.xlsx"
        original.write_bytes(b"source")
        task_id = self._start()
        self.store.archive_sources(task_id, [SourceSpec(original)])
        self.store.mark_success(task_id)

        for suffix in ("", "-wal", "-shm"):
            (self.store.root / f"history.db{suffix}").unlink(missing_ok=True)
        rebuilt = HistoryStore(self.store.root)
        self.assertEqual(rebuilt.rebuild_index_from_manifests(), 0)
        detail = rebuilt.get_task(task_id)
        self.assertIsNotNone(detail)
        assert detail is not None
        self.assertEqual(detail.summary.status, "success")
        self.assertEqual(detail.summary.input_names, ("源.xlsx",))

    def test_rejects_symlink_source(self) -> None:
        target = self.source_dir / "target.xlsx"
        target.write_bytes(b"data")
        link = self.source_dir / "link.xlsx"
        try:
            link.symlink_to(target)
        except (OSError, NotImplementedError):
            self.skipTest("symlinks are unavailable")
        task_id = self._start()
        with self.assertRaisesRegex(HistoryStoreError, "链接"):
            self.store.archive_sources(task_id, [SourceSpec(link)])

    def test_rejects_source_directory_that_contains_history_root(self) -> None:
        task_id = self._start()
        with self.assertRaisesRegex(HistoryStoreError, "包含资料库"):
            self.store.archive_sources(task_id, [SourceSpec(self.base)])

    def test_archived_input_can_be_reused_without_modifying_original_record(self) -> None:
        original = self.source_dir / "复用.xlsx"
        original.write_bytes(b"reusable")
        first_task = self._start()
        first_record = self.store.archive_sources(first_task, [SourceSpec(original)])[0]
        self.store.mark_success(first_task)

        second_task = self._start()
        second_record = self.store.archive_sources(second_task, [SourceSpec(first_record.archived_path)])[0]
        self.assertEqual(second_record.archived_path.read_bytes(), b"reusable")
        self.assertEqual(first_record.archived_path.read_bytes(), b"reusable")

    def test_tampered_archived_input_cannot_be_reused(self) -> None:
        original = self.source_dir / "原件.xlsx"
        original.write_bytes(b"original")
        first_task = self._start()
        first_record = self.store.archive_sources(first_task, [SourceSpec(original)])[0]
        self.store.mark_success(first_task)
        first_record.archived_path.write_bytes(b"tampered")

        second_task = self._start()
        with self.assertRaisesRegex(HistoryStoreError, "校验不一致"):
            self.store.archive_sources(second_task, [SourceSpec(first_record.archived_path)])
        second_detail = self.store.get_task(second_task)
        assert second_detail is not None
        self.assertEqual(second_detail.inputs, ())

    def test_same_task_archives_are_serialized(self) -> None:
        first_dir = self.source_dir / "one"
        second_dir = self.source_dir / "two"
        first_dir.mkdir()
        second_dir.mkdir()
        (first_dir / "同名.xlsx").write_bytes(b"one")
        (second_dir / "同名.xlsx").write_bytes(b"two")
        task_id = self._start()

        with ThreadPoolExecutor(max_workers=2) as executor:
            records = list(
                executor.map(
                    lambda path: self.store.archive_sources(task_id, [SourceSpec(path)]),
                    (first_dir / "同名.xlsx", second_dir / "同名.xlsx"),
                )
            )
        archived = [record for batch in records for record in batch]
        self.assertEqual(len(archived), 2)
        self.assertEqual({item.archived_path.read_bytes() for item in archived}, {b"one", b"two"})
        self.assertFalse(any(self.store.root.rglob("*.partial")))

    def test_same_task_is_serialized_across_store_instances(self) -> None:
        first_dir = self.source_dir / "cross-one"
        second_dir = self.source_dir / "cross-two"
        first_dir.mkdir()
        second_dir.mkdir()
        first = first_dir / "同名.xlsx"
        second = second_dir / "同名.xlsx"
        first.write_bytes(b"one")
        second.write_bytes(b"two")
        task_id = self._start()
        second_store = HistoryStore(self.store.root)

        with ThreadPoolExecutor(max_workers=2) as executor:
            batches = list(
                executor.map(
                    lambda item: item[0].archive_sources(task_id, [SourceSpec(item[1])]),
                    ((self.store, first), (second_store, second)),
                )
            )
        archived = [record for batch in batches for record in batch]
        self.assertEqual({item.archived_path.read_bytes() for item in archived}, {b"one", b"two"})
        self.assertTrue(all(item.archived_path.is_file() for item in archived))
        self.assertTrue(self.store.integrity_check())

    def test_task_directory_collision_never_deletes_existing_task(self) -> None:
        fixed_id = uuid.UUID("12345678-1234-5678-1234-567812345678")
        fixed_now = datetime(2026, 8, 24, 12, 0, tzinfo=timezone.utc)
        with (
            patch("hr_toolkit.history_store.uuid.uuid4", return_value=fixed_id),
            patch("hr_toolkit.history_store.datetime", wraps=datetime) as mocked_datetime,
        ):
            mocked_datetime.now.return_value = fixed_now
            task_id = self._start()
            detail = self.store.get_task(task_id)
            assert detail is not None
            sentinel = detail.task_dir / "keep.txt"
            sentinel.write_text("keep", encoding="utf-8")
            with self.assertRaises(FileExistsError):
                self._start()
        self.assertEqual(sentinel.read_text(encoding="utf-8"), "keep")

    def test_start_task_manifest_failure_rolls_back_hidden_running_record(self) -> None:
        with patch.object(self.store, "_write_manifest", side_effect=OSError("disk full")):
            with self.assertRaisesRegex(OSError, "disk full"):
                self._start()
        tasks, total = self.store.list_tasks()
        self.assertEqual(tasks, ())
        self.assertEqual(total, 0)
        self.assertFalse(any(path.is_dir() for path in self.store.records_dir.glob("*/*/*")))

    def test_single_file_archive_keeps_business_parent_context(self) -> None:
        context_dir = self.source_dir / "2026年7月" / "华东社保账套"
        context_dir.mkdir(parents=True)
        source = context_dir / "缴费明细.xlsx"
        source.write_bytes(b"social-security")
        task_id = self._start("社保报表")
        record = self.store.archive_sources(task_id, [SourceSpec(source)])[0]
        self.assertIn("2026年7月", record.relative_path)
        self.assertIn("华东社保账套", record.relative_path)

    def test_terminal_task_owned_staging_can_be_cleaned_after_worker_exit(self) -> None:
        task_id = self._start()
        detail = self.store.get_task(task_id)
        assert detail is not None
        partial = detail.input_dir / ".工资.xlsx.12345678123456781234567812345678.partial"
        partial.write_bytes(b"partial")
        self.store.mark_stopped(task_id)
        self.store.cleanup_task_staging(task_id)
        self.assertFalse(partial.exists())

    def test_corrupt_database_is_backed_up_and_rebuilt_from_manifests(self) -> None:
        source = self.source_dir / "可恢复.xlsx"
        source.write_bytes(b"recoverable")
        task_id = self._start()
        self.store.archive_sources(task_id, [SourceSpec(source)])
        self.store.mark_success(task_id)
        with closing(sqlite3.connect(self.store.database_path)) as connection:
            connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        for suffix in ("-wal", "-shm", "-journal"):
            self.store.database_path.with_name(self.store.database_path.name + suffix).unlink(missing_ok=True)
        corrupt_bytes = b"not-a-sqlite-database"
        self.store.database_path.write_bytes(corrupt_bytes)

        recovered = HistoryStore(self.store.root)
        self.assertIsNotNone(recovered.recovered_database_backup)
        assert recovered.recovered_database_backup is not None
        self.assertEqual(
            (recovered.recovered_database_backup / "history.db").read_bytes(),
            corrupt_bytes,
        )
        self.assertTrue((recovered.recovered_database_backup / "complete.json").is_file())
        self.assertTrue((recovered.recovered_database_backup / "recovery-report.json").is_file())
        self.assertFalse((self.store.root / ".database-recovery-pending.json").exists())
        detail = recovered.get_task(task_id)
        assert detail is not None
        self.assertEqual(detail.summary.status, "success")
        self.assertTrue(recovered.integrity_check())

    def test_zero_byte_database_with_records_is_backed_up_and_recovered(self) -> None:
        source = self.source_dir / "零字节恢复.xlsx"
        source.write_bytes(b"zero-byte")
        task_id = self._start()
        self.store.archive_sources(task_id, [SourceSpec(source)])
        self.store.mark_success(task_id)
        with closing(sqlite3.connect(self.store.database_path)) as connection:
            connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        for suffix in ("-wal", "-shm", "-journal"):
            self.store.database_path.with_name(self.store.database_path.name + suffix).unlink(missing_ok=True)
        self.store.database_path.write_bytes(b"")

        recovered = HistoryStore(self.store.root)
        detail = recovered.get_task(task_id)
        self.assertIsNotNone(detail)
        assert recovered.recovered_database_backup is not None
        self.assertEqual((recovered.recovered_database_backup / "history.db").read_bytes(), b"")

    def test_trash_only_history_is_recovered_after_database_corruption(self) -> None:
        source = self.source_dir / "回收站恢复.xlsx"
        source.write_bytes(b"trash-data")
        task_id = self._start()
        self.store.archive_sources(task_id, [SourceSpec(source)])
        self.store.mark_success(task_id)
        trash_path = self.store.move_to_trash(task_id)
        for suffix in ("-wal", "-shm", "-journal"):
            self.store.database_path.with_name(self.store.database_path.name + suffix).unlink(missing_ok=True)
        self.store.database_path.write_bytes(b"broken-trash-index")

        recovered = HistoryStore(self.store.root)
        detail = recovered.get_task(task_id)
        self.assertIsNotNone(detail)
        assert detail is not None
        self.assertIsNotNone(detail.summary.deleted_at)
        self.assertEqual(detail.task_dir, trash_path)
        self.assertEqual(recovered.list_tasks()[1], 0)
        self.assertEqual(recovered.storage_stats()["trash_bytes"], len(b"trash-data"))

    def test_missing_manifest_blocks_partial_automatic_rebuild(self) -> None:
        task_id = self._start()
        self.store.mark_success(task_id)
        detail = self.store.get_task(task_id)
        assert detail is not None
        (detail.task_dir / "manifest.json").unlink()
        for suffix in ("-wal", "-shm", "-journal"):
            self.store.database_path.with_name(self.store.database_path.name + suffix).unlink(missing_ok=True)
        broken = b"broken-with-missing-manifest"
        self.store.database_path.write_bytes(broken)

        with self.assertRaisesRegex(HistoryStoreError, "缺少清单"):
            HistoryStore(self.store.root)
        self.assertEqual(self.store.database_path.read_bytes(), broken)
        self.assertTrue((self.store.root / ".database-recovery-pending.json").is_file())

    def test_exact_schema_fingerprint_rejects_extra_trigger_without_rewrite(self) -> None:
        for suffix in ("", "-wal", "-shm", "-journal"):
            self.store.database_path.with_name(self.store.database_path.name + suffix).unlink(missing_ok=True)
        with closing(sqlite3.connect(self.store.database_path)) as connection:
            connection.executescript(SCHEMA_SQL)
            connection.execute(
                "CREATE TRIGGER unexpected_trigger AFTER INSERT ON tasks BEGIN SELECT 1; END"
            )
            connection.execute(f"PRAGMA application_id = {SQLITE_APPLICATION_ID}")
            connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
            connection.commit()
        before = self.store.database_path.read_bytes()

        with self.assertRaisesRegex(HistoryStoreError, "结构无法识别"):
            HistoryStore(self.store.root)
        self.assertEqual(self.store.database_path.read_bytes(), before)
        self.assertFalse(self.store.database_backups_dir.exists())

    def test_future_version_in_wal_is_rejected_without_touching_database_artifacts(self) -> None:
        with closing(sqlite3.connect(self.store.database_path)) as connection:
            connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        connection = sqlite3.connect(self.store.database_path)
        try:
            connection.execute("PRAGMA journal_mode = WAL")
            connection.execute("PRAGMA wal_autocheckpoint = 0")
            connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION + 98}")
            connection.commit()
            artifacts = [
                path
                for path in (
                    self.store.database_path,
                    self.store.database_path.with_name("history.db-wal"),
                    self.store.database_path.with_name("history.db-shm"),
                )
                if path.exists()
            ]
            before = {
                path.name: (path.read_bytes(), path.stat().st_mtime_ns)
                for path in artifacts
            }
            with self.assertRaisesRegex(HistoryStoreError, "版本不兼容"):
                self.store._probe_existing_database()
            after = {
                path.name: (path.read_bytes(), path.stat().st_mtime_ns)
                for path in artifacts
            }
            self.assertEqual(after, before)
        finally:
            connection.close()

    def test_interrupted_corrupt_database_recovery_resumes_from_verified_backup(self) -> None:
        source = self.source_dir / "断点恢复.xlsx"
        source.write_bytes(b"resume")
        task_id = self._start()
        self.store.archive_sources(task_id, [SourceSpec(source)])
        self.store.mark_success(task_id)
        for suffix in ("-wal", "-shm", "-journal"):
            self.store.database_path.with_name(self.store.database_path.name + suffix).unlink(missing_ok=True)
        self.store.database_path.write_bytes(b"broken-before-resume")

        with patch.object(
            HistoryStore,
            "_rebuild_and_publish_database",
            side_effect=RuntimeError("simulated interruption"),
        ):
            with self.assertRaisesRegex(RuntimeError, "simulated interruption"):
                HistoryStore(self.store.root)
        self.assertTrue((self.store.root / ".database-recovery-pending.json").is_file())

        recovered = HistoryStore(self.store.root)
        detail = recovered.get_task(task_id)
        self.assertIsNotNone(detail)
        self.assertFalse((self.store.root / ".database-recovery-pending.json").exists())
        self.assertTrue(recovered.integrity_check())

    def test_failed_atomic_publish_keeps_official_main_database(self) -> None:
        official_before = self.store.database_path.read_bytes()
        rebuild_path = self.store.root / f".history-rebuild-{uuid.uuid4().hex}.db"
        rebuild_path.write_bytes(b"replacement")
        with patch("pathlib.Path.replace", side_effect=OSError("publish failed")):
            with self.assertRaisesRegex(OSError, "publish failed"):
                self.store._publish_rebuilt_database(rebuild_path, self.store.database_path)
        self.assertEqual(self.store.database_path.read_bytes(), official_before)
        self.assertEqual(rebuild_path.read_bytes(), b"replacement")

    def test_unknown_database_is_rejected_without_modifying_it(self) -> None:
        for path in (
            self.store.database_path,
            self.store.database_path.with_name("history.db-wal"),
            self.store.database_path.with_name("history.db-shm"),
            self.store.database_path.with_name("history.db-journal"),
        ):
            path.unlink(missing_ok=True)
        with closing(sqlite3.connect(self.store.database_path)) as connection:
            connection.execute("PRAGMA application_id = 123456")
            connection.execute("CREATE TABLE unrelated(secret TEXT)")
            connection.execute("INSERT INTO unrelated VALUES ('keep')")
            connection.commit()
        before = self.store.database_path.read_bytes()
        with self.assertRaisesRegex(HistoryStoreError, "不是 HRToolkit"):
            HistoryStore(self.store.root)
        self.assertEqual(self.store.database_path.read_bytes(), before)
        self.assertFalse(self.store.database_backups_dir.exists())

    def test_concurrent_first_start_is_serialized(self) -> None:
        fresh_root = self.base / "concurrent-first-start"

        def open_store(_index: int) -> bool:
            return HistoryStore(fresh_root).integrity_check()

        with ThreadPoolExecutor(max_workers=8) as executor:
            results = list(executor.map(open_store, range(8)))
        self.assertEqual(results, [True] * 8)
        self.assertEqual(len(list(fresh_root.glob("history.db"))), 1)

    def test_concurrent_first_start_is_serialized_across_processes(self) -> None:
        fresh_root = self.base / "concurrent-process-first-start"
        results = self._open_store_in_processes(fresh_root)
        self.assertEqual(
            [(result.returncode, result.stdout.strip(), result.stderr) for result in results],
            [(0, "ok", "")] * len(results),
        )

    def test_concurrent_corrupt_recovery_creates_one_backup_across_processes(self) -> None:
        task_id = self._start()
        self.store.mark_success(task_id)
        for suffix in ("-wal", "-shm", "-journal"):
            self.store.database_path.with_name(self.store.database_path.name + suffix).unlink(missing_ok=True)
        self.store.database_path.write_bytes(b"concurrent-corrupt-index")

        results = self._open_store_in_processes(self.store.root)
        self.assertEqual(
            [(result.returncode, result.stdout.strip(), result.stderr) for result in results],
            [(0, "ok", "")] * len(results),
        )
        backups = [path for path in self.store.database_backups_dir.iterdir() if path.is_dir()]
        self.assertEqual(len(backups), 1)
        recovered = HistoryStore(self.store.root).get_task(task_id)
        self.assertIsNotNone(recovered)

    def test_crash_recovery_indexes_completed_orphan_copy(self) -> None:
        task_id = self._start()
        detail = self.store.get_task(task_id)
        assert detail is not None
        orphan = detail.input_dir / "input_path" / "孤儿.xlsx"
        orphan.parent.mkdir(parents=True)
        orphan.write_bytes(b"orphan")
        business_tmp = detail.input_dir / "input_path" / "important.tmp"
        business_tmp.write_bytes(b"business-temp")
        partial = detail.input_dir / ".未完成.xlsx.12345678123456781234567812345678.partial"
        partial.write_bytes(b"partial")
        with patch("hr_toolkit.history_store._pid_is_alive", return_value=False):
            reopened = HistoryStore(self.store.root)
        recovered = reopened.get_task(task_id)
        assert recovered is not None
        self.assertEqual(recovered.summary.status, "failed")
        self.assertEqual(set(recovered.summary.input_names), {"孤儿.xlsx", "important.tmp"})
        self.assertEqual(
            {item.archived_path.read_bytes() for item in recovered.inputs},
            {b"orphan", b"business-temp"},
        )
        self.assertFalse(partial.exists())

    def test_recovery_never_deletes_registered_business_tmp_file(self) -> None:
        business_file = self.source_dir / "important.tmp"
        business_file.write_bytes(b"important")
        task_id = self._start()
        record = self.store.archive_sources(task_id, [SourceSpec(business_file, suffixes=None)])[0]
        with patch("hr_toolkit.history_store._pid_is_alive", return_value=False):
            reopened = HistoryStore(self.store.root)
        self.assertTrue(record.archived_path.is_file())
        self.assertEqual(record.archived_path.read_bytes(), b"important")
        recovered = reopened.get_task(task_id)
        assert recovered is not None
        self.assertEqual(recovered.summary.input_names, ("important.tmp",))

    def test_manifest_rebuild_rejects_path_outside_task(self) -> None:
        original = self.source_dir / "安全.xlsx"
        original.write_bytes(b"safe")
        task_id = self._start()
        self.store.archive_sources(task_id, [SourceSpec(original)])
        self.store.mark_success(task_id)
        detail = self.store.get_task(task_id)
        assert detail is not None
        manifest_path = detail.task_dir / "manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["files"][0]["relative_path"] = "history.db"
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        for suffix in ("", "-wal", "-shm"):
            (self.store.root / f"history.db{suffix}").unlink(missing_ok=True)

        with self.assertRaisesRegex(HistoryStoreError, "未通过完整校验"):
            HistoryStore(self.store.root)
        self.assertTrue((self.store.root / ".database-recovery-pending.json").is_file())

    @unittest.skipIf(os.name == "nt", "symlink creation may require elevated Windows privileges")
    def test_manifest_rebuild_rejects_linked_task_subdirectory(self) -> None:
        task_a = self._start("A")
        task_b = self._start("B")
        source = self.source_dir / "B.xlsx"
        source.write_bytes(b"b")
        record_b = self.store.archive_sources(task_b, [SourceSpec(source)])[0]
        self.store.mark_success(task_a)
        self.store.mark_success(task_b)
        detail_a = self.store.get_task(task_a)
        detail_b = self.store.get_task(task_b)
        assert detail_a is not None and detail_b is not None
        detail_a.input_dir.rmdir()
        detail_a.input_dir.symlink_to(detail_b.input_dir, target_is_directory=True)
        linked_path = detail_a.input_dir / record_b.archived_path.relative_to(detail_b.input_dir)
        manifest_path = detail_a.task_dir / "manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["files"] = [
            {
                "kind": "input",
                "role": "input_path",
                "display_name": "B.xlsx",
                "original_path": "B.xlsx",
                "relative_path": linked_path.relative_to(self.store.root).as_posix(),
                "size_bytes": record_b.size_bytes,
                "sha256": record_b.sha256,
                "modified_ns": record_b.modified_ns,
            }
        ]
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        self.assertEqual(self.store.rebuild_index_from_manifests(), 0)
        refreshed_a = self.store.get_task(task_a)
        assert refreshed_a is not None
        self.assertEqual(refreshed_a.inputs, ())

    def test_file_metadata_does_not_store_absolute_source_path(self) -> None:
        original = self.source_dir / "隐私.xlsx"
        original.write_bytes(b"private")
        task_id = self._start()
        record = self.store.archive_sources(task_id, [SourceSpec(original)])[0]
        self.assertEqual(record.original_path, "隐私.xlsx")

    def test_insufficient_space_fails_before_copying(self) -> None:
        original = self.source_dir / "large.xlsx"
        original.write_bytes(b"x" * 1024)
        task_id = self._start()
        disk_usage = shutil.disk_usage(self.store.root)
        fake_usage = type("DiskUsage", (), {"total": 1024, "used": 1024, "free": 0})()
        self.assertGreaterEqual(disk_usage.total, 1)  # keep the real path exercised
        with patch("hr_toolkit.history_store.shutil.disk_usage", return_value=fake_usage):
            with self.assertRaisesRegex(HistoryStoreError, "空间不足"):
                self.store.archive_sources(task_id, [SourceSpec(original)])
        detail = self.store.get_task(task_id)
        assert detail is not None
        self.assertEqual(detail.inputs, ())

    def test_concurrent_task_writes_keep_database_consistent(self) -> None:
        def create_task(index: int) -> str:
            task_id = self.store.start_task(
                tool_id="data_statistics",
                tool_name="考勤与周月报",
                app_version="0.2.4",
                parameters={"index": index},
            )
            self.store.mark_success(task_id, {"index": index})
            return task_id

        with ThreadPoolExecutor(max_workers=8) as executor:
            task_ids = list(executor.map(create_task, range(24)))
        tasks, total = self.store.list_tasks(limit=50)
        self.assertEqual(total, 24)
        self.assertEqual({task.id for task in tasks}, set(task_ids))
        self.assertTrue(self.store.integrity_check())

    @unittest.skipIf(os.name == "nt", "POSIX permissions are not meaningful on Windows")
    def test_data_files_are_private_on_posix(self) -> None:
        task_id = self._start()
        detail = self.store.get_task(task_id)
        assert detail is not None
        self.assertEqual(self.store.root.stat().st_mode & 0o777, 0o700)
        self.assertEqual(self.store.database_path.stat().st_mode & 0o777, 0o600)
        self.assertEqual((detail.task_dir / "manifest.json").stat().st_mode & 0o777, 0o600)


class DefaultHistoryRootTests(unittest.TestCase):
    def test_environment_override_is_used(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "custom"
            with patch.dict(os.environ, {DATA_DIR_ENV: str(path)}):
                self.assertEqual(default_history_root(), path)

    def test_nonempty_unmarked_directory_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "occupied"
            root.mkdir()
            (root / "unrelated.txt").write_text("keep", encoding="utf-8")
            with self.assertRaisesRegex(HistoryStoreError, "专用文件夹"):
                HistoryStore(root)
            self.assertEqual((root / "unrelated.txt").read_text(encoding="utf-8"), "keep")


if __name__ == "__main__":
    unittest.main()
