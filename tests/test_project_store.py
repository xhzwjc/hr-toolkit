from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch

import hr_toolkit.project_store as project_store_module
from hr_toolkit.project_store import (
    CATEGORY_RESULTS,
    CATEGORY_SUPPLEMENTS,
    CATEGORY_UPLOADS,
    COMMON_VISIBLE_DIR,
    PROJECT_FILE_NAME,
    PROJECT_FORMAT_VERSION,
    PROJECT_METADATA_DIR,
    ImportCancelled,
    ImportProgress,
    ProjectStore,
    ProjectStoreError,
    TrashBatchDetail,
    validate_project_name,
)


class ProjectStoreTests(unittest.TestCase):
    def setUp(self) -> None:
        self._temporary = tempfile.TemporaryDirectory()
        self.base = Path(self._temporary.name).resolve()
        self.project_root = self.base / "测试项目"
        self.store = ProjectStore.create(self.project_root, "人事月度处理")
        self.sources = self.base / "外部资料"
        self.sources.mkdir()

    def tearDown(self) -> None:
        self.store.close()
        self._temporary.cleanup()

    def _draft(self, *, group_name: str = "甲公司", tool_name: str = "工资表拆分"):
        return self.store.create_draft(
            group_name=group_name,
            tool_id="salary_split",
            tool_name=tool_name,
            business_description="7月工资",
            business_period="2026-07",
        )

    def _running(self, *, now: datetime | None = None):
        draft = self._draft()
        return self.store.start_batch(draft.summary.id, now=now)

    def _trashed_successful_batch(
        self,
        *,
        file_name: str = "回收.xlsx",
        content: bytes = b"restore",
    ):
        source = self.sources / file_name
        source.write_bytes(content)
        draft = self._draft()
        self.store.import_sources(draft.summary.id, [source])
        running = self.store.start_batch(draft.summary.id)
        self.store.mark_success(running.summary.id)
        trash_path = self.store.move_to_trash(running.summary.id)
        return running, trash_path

    def test_create_only_builds_common_material_and_hidden_metadata(self) -> None:
        visible = sorted(path.name for path in self.project_root.iterdir() if path.name != PROJECT_METADATA_DIR)
        self.assertEqual(visible, [COMMON_VISIBLE_DIR])
        marker = self.project_root / PROJECT_METADATA_DIR / PROJECT_FILE_NAME
        payload = json.loads(marker.read_text(encoding="utf-8"))
        self.assertEqual(payload["format_version"], PROJECT_FORMAT_VERSION)
        self.assertEqual(payload["project_id"], self.store.workspace.project_id)
        self.assertTrue(self.store.writable)
        self.assertTrue(self.store.integrity_check())

    def test_list_batch_locations_matches_batch_detail_without_materializing_files(self) -> None:
        draft = self._draft()
        source = self.sources / "名单.xlsx"
        source.write_bytes(b"roster")
        self.store.import_sources(draft.summary.id, [source])

        locations = self.store.list_batch_locations()
        detail = self.store.get_batch(draft.summary.id)

        self.assertIsNotNone(detail)
        self.assertEqual(len(locations), 1)
        summary, directories = locations[0]
        self.assertEqual(summary, detail.summary)
        self.assertEqual(directories, detail.directories)

    def test_create_rejects_nonempty_folder_without_touching_it(self) -> None:
        occupied = self.base / "已有资料"
        occupied.mkdir()
        original = occupied / "不能覆盖.xlsx"
        original.write_bytes(b"keep")
        with self.assertRaisesRegex(ProjectStoreError, "空文件夹"):
            ProjectStore.create(occupied, "危险项目")
        self.assertEqual(original.read_bytes(), b"keep")
        self.assertFalse((occupied / PROJECT_METADATA_DIR).exists())

    def test_project_name_rules_are_portable_across_windows_and_macos(self) -> None:
        self.assertEqual(validate_project_name("  华东人事项目  "), "华东人事项目")
        for value in (
            "",
            ".",
            "..",
            "工资/社保",
            "工资:社保",
            "工资.",
            "换行\n项目",
            ".hrtoolkit",
            "CON",
            "con.txt",
            "LPT9",
            "x" * 121,
        ):
            with self.subTest(value=value):
                with self.assertRaises(ProjectStoreError):
                    validate_project_name(value)

    def test_create_rejects_project_path_reached_through_linked_parent(self) -> None:
        real_parent = self.base / "真实位置"
        real_parent.mkdir()
        linked_parent = self.base / "位置别名"
        try:
            linked_parent.symlink_to(real_parent, target_is_directory=True)
        except (OSError, NotImplementedError):
            self.skipTest("当前文件系统不支持符号链接")

        with self.assertRaisesRegex(ProjectStoreError, "链接|重定向"):
            ProjectStore.create(linked_parent / "新项目", "新项目")
        self.assertFalse((real_parent / "新项目").exists())

    def test_draft_layout_is_group_tool_batch_and_supplements_are_on_demand(self) -> None:
        detail = self._draft(group_name="华东事业部", tool_name="社保明细与汇总")
        batch_root = self.project_root / "华东事业部" / "社保明细与汇总" / detail.summary.directory_name
        self.assertEqual(detail.summary.status, "draft")
        self.assertEqual(detail.summary.group_name, "华东事业部")
        self.assertTrue((batch_root / "上传资料").is_dir())
        self.assertTrue((batch_root / "处理结果").is_dir())
        self.assertFalse((batch_root / "补充资料").exists())

        manifest_path = self.project_root / PROJECT_METADATA_DIR / "manifests" / f"{detail.summary.id}.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(manifest["batch"]["tool_id"], "salary_split")
        self.assertEqual(manifest["batch"]["tool_name"], "社保明细与汇总")
        self.assertEqual(manifest["batch"]["group_name"], "华东事业部")
        self.assertFalse(any(Path(value).is_absolute() for value in manifest["directories"].values()))

    def test_same_second_batches_are_unique_and_cross_day_names_follow_start_day(self) -> None:
        local_zone = timezone(timedelta(hours=8))
        same_time = datetime(2026, 8, 3, 10, 20, 30, tzinfo=local_zone)
        first = self.store.start_batch(self._draft().summary.id, now=same_time)
        second = self.store.start_batch(self._draft().summary.id, now=same_time)
        next_day = self.store.start_batch(
            self._draft().summary.id,
            now=datetime(2026, 8, 4, 9, 0, 0, tzinfo=local_zone),
        )
        self.assertEqual(first.summary.directory_name, "20260803_102030_7月工资_2026-07")
        self.assertEqual(second.summary.directory_name, "20260803_102030_7月工资_2026-07_2")
        self.assertTrue(next_day.summary.directory_name.startswith("20260804_090000_"))
        self.assertNotEqual(first.summary.id, second.summary.id)

    def test_import_preserves_source_and_renames_second_top_folder_as_a_unit(self) -> None:
        first_parent = self.sources / "第一处"
        second_parent = self.sources / "第二处"
        first = first_parent / "资料"
        second = second_parent / "资料"
        first.mkdir(parents=True)
        second.mkdir(parents=True)
        (first / "一月.xlsx").write_bytes(b"one")
        (first / "子目录").mkdir()
        (first / "子目录" / "说明.txt").write_bytes(b"note")
        (second / "二月.xlsx").write_bytes(b"two")
        source_snapshot = {
            path: (path.read_bytes(), path.stat().st_mtime_ns)
            for path in (first / "一月.xlsx", first / "子目录" / "说明.txt", second / "二月.xlsx")
        }

        draft = self._draft()
        records = self.store.import_sources(draft.summary.id, [first, second])
        upload_root = draft.directories[CATEGORY_UPLOADS]
        self.assertEqual(len(records), 3)
        self.assertTrue((upload_root / "资料" / "一月.xlsx").is_file())
        self.assertTrue((upload_root / "资料" / "子目录" / "说明.txt").is_file())
        self.assertTrue((upload_root / "资料 (2)" / "二月.xlsx").is_file())
        self.assertFalse((upload_root / "资料" / "二月.xlsx").exists())
        for path, (content, modified_ns) in source_snapshot.items():
            self.assertEqual(path.read_bytes(), content)
            self.assertEqual(path.stat().st_mtime_ns, modified_ns)

    def test_duplicate_file_names_never_overwrite(self) -> None:
        first = self.sources / "甲" / "工资.xlsx"
        second = self.sources / "乙" / "工资.xlsx"
        first.parent.mkdir()
        second.parent.mkdir()
        first.write_bytes(b"first")
        second.write_bytes(b"second")
        draft = self._draft()
        self.store.import_sources(draft.summary.id, [first, second])
        upload_root = draft.directories[CATEGORY_UPLOADS]
        self.assertEqual((upload_root / "工资.xlsx").read_bytes(), b"first")
        self.assertEqual((upload_root / "工资 (2).xlsx").read_bytes(), b"second")

    def test_project_sources_snapshot_old_uploads_and_completed_results(self) -> None:
        source = self.sources / "上月花名册.xlsx"
        source.write_bytes(b"old-upload")
        old_draft = self._draft()
        self.store.import_sources(old_draft.summary.id, [source])
        old_running = self.store.start_batch(old_draft.summary.id)
        old_result = self.store.result_directory(old_running.summary.id) / "上月结果.xlsx"
        old_result.write_bytes(b"old-result")
        self.store.register_results(old_running.summary.id, old_result)
        old_finished = self.store.mark_success(old_running.summary.id)
        old_upload = next(
            item.path(self.store.workspace)
            for item in old_finished.files
            if item.category == CATEGORY_UPLOADS
        )

        new_draft = self._draft()
        copied = self.store.copy_project_sources(
            new_draft.summary.id,
            [old_upload, old_result],
        )
        copied_paths = {item.display_name: item.path(self.store.workspace) for item in copied}
        self.assertEqual(copied_paths["上月花名册.xlsx"].read_bytes(), b"old-upload")
        self.assertEqual(copied_paths["上月结果.xlsx"].read_bytes(), b"old-result")
        self.assertNotEqual(copied_paths["上月花名册.xlsx"], old_upload)
        self.assertNotEqual(copied_paths["上月花名册.xlsx"].stat().st_ino, old_upload.stat().st_ino)
        self.assertEqual(copied_paths["上月花名册.xlsx"].parent, new_draft.directories[CATEGORY_UPLOADS])

        unfinished = self._running()
        unfinished_result = self.store.result_directory(unfinished.summary.id) / "未完成.xlsx"
        unfinished_result.write_bytes(b"unfinished")
        self.store.register_results(unfinished.summary.id, unfinished_result)
        with self.assertRaisesRegex(ProjectStoreError, "成功完成"):
            self.store.copy_project_sources(new_draft.summary.id, [unfinished_result])
        with self.assertRaisesRegex(ProjectStoreError, "隐藏"):
            self.store.copy_project_sources(
                new_draft.summary.id,
                [self.project_root / PROJECT_METADATA_DIR / PROJECT_FILE_NAME],
            )
        with self.assertRaisesRegex(ProjectStoreError, "目标批次自身"):
            self.store.copy_project_sources(
                new_draft.summary.id,
                [new_draft.directories[CATEGORY_UPLOADS]],
            )

    def test_project_source_reuse_rejects_changed_or_unregistered_results(self) -> None:
        running = self._running()
        result = self.store.result_directory(running.summary.id) / "上月结果.xlsx"
        result.write_bytes(b"registered")
        self.store.register_results(running.summary.id, result)
        finished = self.store.mark_success(running.summary.id)
        new_draft = self._draft()

        result.write_bytes(b"changed")
        with self.assertRaisesRegex(ProjectStoreError, "原清单|不一致|不能复用"):
            self.store.copy_project_sources(new_draft.summary.id, [result])

        result.write_bytes(b"registered")
        unregistered = finished.directories[CATEGORY_RESULTS] / "手工新增.xlsx"
        unregistered.write_bytes(b"not registered")
        with self.assertRaisesRegex(ProjectStoreError, "已登记|清单"):
            self.store.copy_project_sources(new_draft.summary.id, [unregistered])
        with self.assertRaisesRegex(ProjectStoreError, "清单"):
            self.store.copy_project_sources(
                new_draft.summary.id,
                [finished.directories[CATEGORY_RESULTS]],
            )

    def test_result_working_copy_uses_only_verified_upload_manifest(self) -> None:
        source_dir = self.sources / "人员资料"
        (source_dir / "张三").mkdir(parents=True)
        (source_dir / "张三" / "说明.txt").write_bytes(b"record")
        draft = self._draft()
        records = self.store.import_sources(draft.summary.id, [source_dir])
        draft_upload_source = records[0].path(self.store.workspace).parents[1]
        relative_source = draft_upload_source.relative_to(draft.directories[CATEGORY_UPLOADS])
        running = self.store.start_batch(draft.summary.id)
        upload_source = running.directories[CATEGORY_UPLOADS] / relative_source

        injected = upload_source / "未登记.txt"
        injected.write_bytes(b"injected")
        with self.assertRaisesRegex(ProjectStoreError, "清单"):
            self.store.create_result_working_copy(running.summary.id, upload_source)
        injected.unlink()

        secret = self.sources / "项目外秘密.txt"
        secret.write_bytes(b"secret")
        link = upload_source / "外部链接.txt"
        try:
            link.symlink_to(secret)
        except (OSError, NotImplementedError):
            link = None
        if link is not None:
            with self.assertRaisesRegex(ProjectStoreError, "链接"):
                self.store.create_result_working_copy(running.summary.id, upload_source)
            link.unlink()

        uploaded_file = upload_source / "张三" / "说明.txt"
        uploaded_file.write_bytes(b"changed")
        with self.assertRaisesRegex(ProjectStoreError, "发生变化"):
            self.store.create_result_working_copy(running.summary.id, upload_source)
        uploaded_file.write_bytes(b"record")

        copied = self.store.create_result_working_copy(running.summary.id, upload_source)
        self.assertEqual((copied / "张三" / "说明.txt").read_bytes(), b"record")

    def test_directory_snapshot_preserves_empty_folders_and_locks_topology(self) -> None:
        source_dir = self.sources / "testname"
        (source_dir / "54").mkdir(parents=True)
        (source_dir / "4343" / "空子目录").mkdir(parents=True)
        (source_dir / "2331221").mkdir(parents=True)
        metadata = source_dir / ".DS_Store"
        metadata.write_bytes(b"finder metadata")
        source_before = sorted(
            path.relative_to(source_dir).as_posix()
            for path in source_dir.rglob("*")
        )
        metadata_before = hashlib.sha256(metadata.read_bytes()).hexdigest()

        draft = self._draft(tool_name="人员资料文件夹改名")
        draft_snapshot = self.store.import_directory_snapshot(
            draft.summary.id,
            source_dir,
            role="input_path",
        )
        relative_snapshot = draft_snapshot.relative_to(
            draft.directories[CATEGORY_UPLOADS]
        )
        running = self.store.start_batch(draft.summary.id)
        upload_snapshot = running.directories[CATEGORY_UPLOADS] / relative_snapshot

        self.assertEqual(
            sorted(path.name for path in upload_snapshot.iterdir()),
            ["2331221", "4343", "54"],
        )
        self.assertTrue((upload_snapshot / "4343" / "空子目录").is_dir())
        self.assertFalse((upload_snapshot / ".DS_Store").exists())

        injected = upload_snapshot / "预览后新增"
        injected.mkdir()
        with self.assertRaisesRegex(ProjectStoreError, "文件夹结构|清单"):
            self.store.create_result_working_copy(running.summary.id, upload_snapshot)
        injected.rmdir()

        copied = self.store.create_result_working_copy(
            running.summary.id,
            upload_snapshot,
        )
        self.assertEqual(
            sorted(
                path.relative_to(copied).as_posix()
                for path in copied.rglob("*")
            ),
            ["2331221", "4343", "4343/空子目录", "54"],
        )
        self.assertEqual(
            sorted(path.relative_to(source_dir).as_posix() for path in source_dir.rglob("*")),
            source_before,
        )
        self.assertEqual(
            hashlib.sha256(metadata.read_bytes()).hexdigest(),
            metadata_before,
        )

    def test_empty_directory_snapshot_survives_batch_move_restore_and_reuse(self) -> None:
        source_dir = self.sources / "空人员资料"
        (source_dir / "001" / "空子目录").mkdir(parents=True)
        draft = self._draft(tool_name="人员资料文件夹改名")
        draft_snapshot = self.store.import_directory_snapshot(
            draft.summary.id,
            source_dir,
            role="input_path",
        )
        relative_snapshot = draft_snapshot.relative_to(
            draft.directories[CATEGORY_UPLOADS]
        )
        running = self.store.start_batch(draft.summary.id)
        self.store.mark_success(running.summary.id)
        self.store.move_to_trash(running.summary.id)
        restored = self.store.restore_from_trash(running.summary.id)
        restored_snapshot = (
            restored.directories[CATEGORY_UPLOADS] / relative_snapshot
        )
        self.assertTrue((restored_snapshot / "001" / "空子目录").is_dir())

        reused_draft = self._draft(tool_name="人员资料文件夹改名")
        reused_snapshot = self.store.copy_project_directory_snapshot(
            reused_draft.summary.id,
            restored_snapshot,
            role="input_path",
        )
        self.assertTrue((reused_snapshot / "001" / "空子目录").is_dir())

    def test_import_ignores_office_temporary_files_and_rejects_scripts_before_copy(self) -> None:
        folder = self.sources / "批量"
        folder.mkdir()
        (folder / "有效.xlsx").write_bytes(b"valid")
        (folder / "~$有效.xlsx").write_bytes(b"lock")
        (folder / ".DS_Store").write_bytes(b"system")
        draft = self._draft()
        records = self.store.import_sources(draft.summary.id, [folder])
        self.assertEqual([record.display_name for record in records], ["有效.xlsx"])

        dangerous = self.sources / "危险"
        dangerous.mkdir()
        (dangerous / "另一个.xlsx").write_bytes(b"valid-two")
        (dangerous / "运行.cmd").write_text("echo unsafe", encoding="utf-8")
        before = set(draft.directories[CATEGORY_UPLOADS].rglob("*"))
        with self.assertRaisesRegex(ProjectStoreError, "不能导入"):
            self.store.import_sources(draft.summary.id, [dangerous])
        self.assertEqual(set(draft.directories[CATEGORY_UPLOADS].rglob("*")), before)

    def test_import_rejects_symbolic_link_without_partial_batch(self) -> None:
        folder = self.sources / "带链接"
        folder.mkdir()
        target = folder / "原件.xlsx"
        target.write_bytes(b"data")
        link = folder / "链接.xlsx"
        try:
            link.symlink_to(target)
        except (OSError, NotImplementedError):
            self.skipTest("当前文件系统不支持符号链接")
        draft = self._draft()
        with self.assertRaisesRegex(ProjectStoreError, "包含链接"):
            self.store.import_sources(draft.summary.id, [folder])
        self.assertFalse(any(draft.directories[CATEGORY_UPLOADS].iterdir()))

    def test_import_rejects_regular_file_reached_through_linked_parent(self) -> None:
        real_folder = self.sources / "真实目录"
        real_folder.mkdir()
        source = real_folder / "秘密.xlsx"
        source.write_bytes(b"private")
        linked_parent = self.sources / "目录链接"
        try:
            linked_parent.symlink_to(real_folder, target_is_directory=True)
        except (OSError, NotImplementedError):
            self.skipTest("当前文件系统不支持符号链接")
        draft = self._draft()
        with self.assertRaisesRegex(ProjectStoreError, "路径不能经过链接"):
            self.store.import_sources(draft.summary.id, [linked_parent / source.name])
        self.assertFalse(any(draft.directories[CATEGORY_UPLOADS].iterdir()))

    def test_source_change_during_stream_copy_rolls_back_visible_file(self) -> None:
        source = self.sources / "大文件.xlsx"
        source.write_bytes(b"a" * (2 * 1024 * 1024))
        changed = False

        def mutate_source(_copied: int, _total: int, _name: str) -> None:
            nonlocal changed
            if not changed:
                changed = True
                with source.open("ab") as handle:
                    handle.write(b"changed")

        draft = self._draft()
        with self.assertRaisesRegex(ProjectStoreError, "发生变化"):
            self.store.import_sources(draft.summary.id, [source], progress=mutate_source)
        self.assertEqual(self.store.get_batch(draft.summary.id).files, ())
        self.assertFalse(any(draft.directories[CATEGORY_UPLOADS].iterdir()))

    def test_structured_progress_keeps_legacy_callback_and_orders_phases(self) -> None:
        source = self.sources / "进度.xlsx"
        source.write_bytes(b"progress-data")
        events: list[ImportProgress] = []
        legacy: list[tuple[int, int, str]] = []

        imported = self.store.import_common_sources(
            [source],
            progress=lambda copied, total, name: legacy.append((copied, total, name)),
            on_progress=events.append,
        )

        self.assertEqual(imported[0].read_bytes(), b"progress-data")
        self.assertEqual(events[0].phase, "checking")
        self.assertIn("copying", [event.phase for event in events])
        self.assertEqual(events[-1].phase, "finalizing")
        copying = [event for event in events if event.phase == "copying"]
        self.assertEqual(
            [event.bytes_copied for event in copying],
            sorted(event.bytes_copied for event in copying),
        )
        self.assertEqual(copying[-1].files_completed, 1)
        self.assertEqual(copying[-1].files_total, 1)
        self.assertTrue(legacy)
        self.assertEqual(legacy[-1], (len(b"progress-data"), len(b"progress-data"), source.name))

    def test_cancel_during_folder_scan_leaves_no_visible_or_staged_files(self) -> None:
        folder = self.sources / "大量资料"
        folder.mkdir()
        for index in range(5):
            (folder / f"资料{index}.xlsx").write_bytes(str(index).encode())
        cancel_requested = False

        def observe(event: ImportProgress) -> None:
            nonlocal cancel_requested
            if event.phase == "checking" and event.files_scanned == 1:
                cancel_requested = True

        with self.assertRaises(ImportCancelled):
            self.store.import_common_sources(
                [folder],
                cancelled=lambda: cancel_requested,
                on_progress=observe,
            )

        self.assertFalse(any(self.store.workspace.common_root.iterdir()))
        self.assertFalse(any(self.store.staging_dir.iterdir()))

    def test_cancel_during_copy_rolls_back_before_finalizing(self) -> None:
        source = self.sources / "取消复制.xlsx"
        source.write_bytes(b"x" * (2 * 1024 * 1024))
        cancel_requested = False
        phases: list[str] = []

        def observe(event: ImportProgress) -> None:
            nonlocal cancel_requested
            phases.append(event.phase)
            if event.phase == "copying" and event.bytes_copied:
                cancel_requested = True

        with self.assertRaises(ImportCancelled):
            self.store.import_common_sources(
                [source],
                cancelled=lambda: cancel_requested,
                on_progress=observe,
            )

        self.assertNotIn("finalizing", phases)
        self.assertFalse((self.store.workspace.common_root / source.name).exists())
        self.assertFalse(any(self.store.staging_dir.iterdir()))

    def test_cancel_requested_after_finalizing_does_not_interrupt_publication(self) -> None:
        source = self.sources / "完成保存.xlsx"
        source.write_bytes(b"finish")
        cancel_requested = False

        def observe(event: ImportProgress) -> None:
            nonlocal cancel_requested
            if event.phase == "finalizing":
                cancel_requested = True

        imported = self.store.import_common_sources(
            [source],
            cancelled=lambda: cancel_requested,
            on_progress=observe,
        )

        self.assertTrue(cancel_requested)
        self.assertEqual(imported[0].read_bytes(), b"finish")
        self.assertFalse(any(self.store.staging_dir.iterdir()))

    def test_empty_or_only_temporary_folder_is_not_reported_as_success(self) -> None:
        empty = self.sources / "空文件夹"
        empty.mkdir()
        with self.assertRaisesRegex(ProjectStoreError, "没有可导入"):
            self.store.import_common_sources([empty])

        temporary = self.sources / "临时文件"
        temporary.mkdir()
        (temporary / "~$花名册.xlsx").write_bytes(b"lock")
        (temporary / ".DS_Store").write_bytes(b"metadata")
        with self.assertRaisesRegex(ProjectStoreError, "没有可导入"):
            self.store.import_common_sources([temporary])
        self.assertFalse(any(self.store.workspace.common_root.iterdir()))

    def test_insufficient_space_fails_before_copying(self) -> None:
        source = self.sources / "空间.xlsx"
        source.write_bytes(b"data")
        draft = self._draft()
        usage = type("DiskUsage", (), {"total": 1, "used": 1, "free": 0})()
        with patch("hr_toolkit.project_store.shutil.disk_usage", return_value=usage):
            with self.assertRaisesRegex(ProjectStoreError, "空间不足"):
                self.store.import_sources(draft.summary.id, [source])
        self.assertFalse(any(draft.directories[CATEGORY_UPLOADS].iterdir()))

    def test_results_are_registered_in_place_without_second_copy(self) -> None:
        running = self._running()
        output = self.store.result_directory(running.summary.id) / "结果.xlsx"
        output.write_bytes(b"result")
        before_stat = output.stat()
        records = self.store.register_results(running.summary.id, output)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].path(self.store.workspace), output)
        self.assertEqual(output.stat().st_ino, before_stat.st_ino)
        self.assertEqual(records[0].sha256, hashlib.sha256(b"result").hexdigest())
        finished = self.store.mark_success(running.summary.id)
        self.assertEqual(finished.summary.status, "success")

    def test_mark_success_rejects_unregistered_result(self) -> None:
        running = self._running()
        output = self.store.result_directory(running.summary.id) / "半成品.xlsx"
        output.write_bytes(b"partial")
        with self.assertRaisesRegex(ProjectStoreError, "未登记"):
            self.store.mark_success(running.summary.id)
        self.assertEqual(self.store.get_batch(running.summary.id).summary.status, "running")
        self.assertTrue(output.exists())

    def test_failed_batch_quarantines_unregistered_results(self) -> None:
        running = self._running()
        output = self.store.result_directory(running.summary.id) / "失败半成品.xlsx"
        output.write_bytes(b"partial")
        failed = self.store.mark_failed(running.summary.id, "测试失败")
        self.assertEqual(failed.summary.status, "failed")
        self.assertFalse(output.exists())
        quarantined = list((self.project_root / PROJECT_METADATA_DIR / "quarantine").rglob("失败半成品.xlsx"))
        self.assertEqual(len(quarantined), 1)
        self.assertEqual(quarantined[0].read_bytes(), b"partial")

    def test_startup_recovery_stops_batch_and_hides_unregistered_results(self) -> None:
        running = self._running()
        output = self.store.result_directory(running.summary.id) / "崩溃半成品.xlsx"
        output.write_bytes(b"partial")
        self.store.close()
        self.store = ProjectStore.open(self.project_root)
        detail = self.store.get_batch(running.summary.id)
        self.assertEqual(detail.summary.status, "stopped")
        self.assertFalse(output.exists())
        self.assertTrue(any((self.project_root / PROJECT_METADATA_DIR / "quarantine").rglob("崩溃半成品.xlsx")))

    def test_second_instance_is_read_only_until_writer_closes(self) -> None:
        second = ProjectStore.open(self.project_root)
        try:
            self.assertFalse(second.writable)
            self.assertIn("另一个窗口", second.workspace.read_only_reason)
            with self.assertRaisesRegex(ProjectStoreError, "只读|另一个窗口"):
                second.create_draft(
                    group_name="甲公司",
                    tool_id="salary_split",
                    tool_name="工资表拆分",
                )
        finally:
            second.close()
        self.store.close()
        third = ProjectStore.open(self.project_root)
        try:
            self.assertTrue(third.writable)
        finally:
            third.close()

    def test_second_process_also_falls_back_to_read_only(self) -> None:
        command = [
            sys.executable,
            "-c",
            (
                "from hr_toolkit.project_store import ProjectStore; import sys; "
                "store=ProjectStore.open(sys.argv[1]); "
                "print('writable' if store.writable else 'readonly'); store.close()"
            ),
            str(self.project_root),
        ]
        result = subprocess.run(command, capture_output=True, text=True, timeout=20, check=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), "readonly")

    def test_future_version_opens_read_only_without_mutating_project(self) -> None:
        self.store.close()
        marker = self.project_root / PROJECT_METADATA_DIR / PROJECT_FILE_NAME
        payload = json.loads(marker.read_text(encoding="utf-8"))
        payload["format_version"] = PROJECT_FORMAT_VERSION + 10
        marker.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        before = {
            path.relative_to(self.project_root).as_posix(): (path.read_bytes(), path.stat().st_mtime_ns)
            for path in self.project_root.rglob("*")
            if path.is_file()
        }
        future = ProjectStore.open(self.project_root)
        try:
            self.assertFalse(future.writable)
            self.assertIn("更高版本", future.workspace.read_only_reason)
            self.assertEqual(future.list_batches(), ())
        finally:
            future.close()
        after = {
            path.relative_to(self.project_root).as_posix(): (path.read_bytes(), path.stat().st_mtime_ns)
            for path in self.project_root.rglob("*")
            if path.is_file()
        }
        self.assertEqual(after, before)
        # Keep tearDown idempotent without reopening the intentionally future project.
        self.store = future

    def test_manifest_path_escape_is_rejected(self) -> None:
        draft = self._draft()
        manifest_path = self.project_root / PROJECT_METADATA_DIR / "manifests" / f"{draft.summary.id}.json"
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
        payload["files"].append(
            {
                "id": "1" * 32,
                "batch_id": draft.summary.id,
                "category": CATEGORY_UPLOADS,
                "role": "main",
                "display_name": "越界.xlsx",
                "relative_path": "../越界.xlsx",
                "size_bytes": 1,
                "sha256": "0" * 64,
                "modified_ns": 0,
            }
        )
        manifest_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        with self.assertRaisesRegex(ProjectStoreError, "相对路径|不属于"):
            self.store.get_batch(draft.summary.id)
        self.assertEqual(self.store.list_batch_locations(), ())

    def test_project_can_move_and_reopen_using_only_relative_paths(self) -> None:
        source = self.sources / "迁移.xlsx"
        source.write_bytes(b"portable")
        draft = self._draft()
        self.store.import_sources(draft.summary.id, [source])
        self.store.close()
        moved = self.base / "搬家后的项目"
        self.project_root.rename(moved)
        self.project_root = moved
        self.store = ProjectStore.open(moved)
        detail = self.store.get_batch(draft.summary.id)
        self.assertEqual(detail.files[0].path(self.store.workspace).read_bytes(), b"portable")
        self.assertTrue(str(detail.files[0].path(self.store.workspace)).startswith(str(moved)))
        self.assertTrue(self.store.integrity_check())

    def test_hidden_recycle_bin_restores_batch_without_overwrite(self) -> None:
        running, trash_path = self._trashed_successful_batch()
        original_root = running.directories[CATEGORY_UPLOADS].parent
        self.assertFalse(original_root.exists())
        self.assertIsNone(self.store.get_batch(running.summary.id))
        self.assertTrue((trash_path / "manifest.json").is_file())
        restored = self.store.restore_from_trash(running.summary.id)
        self.assertEqual(restored.files[0].path(self.store.workspace).read_bytes(), b"restore")
        self.assertTrue(restored.directories[CATEGORY_UPLOADS].is_dir())
        self.assertEqual(self.store.list_trash(), ())

    def test_trash_details_expose_aggregate_business_data_without_hidden_paths(self) -> None:
        upload = self.sources / "员工明细.xlsx"
        supplement = self.sources / "补充证明.pdf"
        upload.write_bytes(b"upload")
        supplement.write_bytes(b"supplement")
        draft = self._draft(group_name="华东事业部", tool_name="社保明细与汇总")
        self.store.import_sources(draft.summary.id, [upload])
        self.store.import_sources(
            draft.summary.id,
            [supplement],
            category=CATEGORY_SUPPLEMENTS,
        )
        running = self.store.start_batch(draft.summary.id)
        result = running.directories[CATEGORY_RESULTS] / "社保汇总.xlsx"
        result.write_bytes(b"result")
        self.store.register_results(running.summary.id, result)
        self.store.mark_success(running.summary.id)
        expected_relative = running.directories[CATEGORY_UPLOADS].parent.relative_to(
            self.project_root
        ).as_posix()
        self.store.move_to_trash(running.summary.id)

        details = self.store.list_trash_details()

        self.assertEqual(len(details), 1)
        detail = details[0]
        self.assertIsInstance(detail, TrashBatchDetail)
        self.assertEqual(detail.summary.business_period, "2026-07")
        self.assertEqual(detail.summary.tool_name, "社保明细与汇总")
        self.assertIsNotNone(detail.summary.deleted_at)
        self.assertEqual(detail.original_relative_path, expected_relative)
        self.assertEqual(detail.upload_count, 1)
        self.assertEqual(detail.result_count, 1)
        self.assertEqual(detail.supplement_count, 1)
        self.assertEqual(
            detail.total_size_bytes,
            len(b"upload") + len(b"supplement") + len(b"result"),
        )
        self.assertFalse(Path(detail.original_relative_path).is_absolute())
        self.assertNotIn(PROJECT_METADATA_DIR, Path(detail.original_relative_path).parts)
        self.assertEqual(self.store.list_trash(), (detail.summary,))

    def test_restore_collision_uses_parenthesized_name_without_overwrite(self) -> None:
        running, _trash_path = self._trashed_successful_batch()
        original_root = running.directories[CATEGORY_UPLOADS].parent
        original_root.mkdir(parents=True)
        sentinel = original_root / "现有资料.txt"
        sentinel.write_bytes(b"keep")

        restored = self.store.restore_from_trash(running.summary.id)

        self.assertEqual(restored.summary.directory_name, f"{running.summary.directory_name} (2)")
        self.assertEqual(sentinel.read_bytes(), b"keep")
        self.assertEqual(
            restored.directories[CATEGORY_UPLOADS].parent.name,
            f"{running.summary.directory_name} (2)",
        )

    def test_restore_rejects_invalid_batch_id_before_reading_trash(self) -> None:
        with self.assertRaisesRegex(ProjectStoreError, "批次编号无效"):
            self.store.restore_from_trash("../manifests/unsafe")

    def test_restore_rejects_mismatched_manifest_identity(self) -> None:
        running, trash_path = self._trashed_successful_batch()
        manifest_path = trash_path / "manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["files"][0]["batch_id"] = "f" * 32
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")

        with self.assertRaisesRegex(ProjectStoreError, "项目文件与批次不一致"):
            self.store.restore_from_trash(running.summary.id)

        self.assertFalse(self.store._manifest_path(running.summary.id).exists())
        self.assertTrue((trash_path / "batch" / "上传资料" / "回收.xlsx").is_file())

    def test_restore_rejects_registered_file_hash_change(self) -> None:
        running, trash_path = self._trashed_successful_batch(content=b"restore")
        stored = trash_path / "batch" / "上传资料" / "回收.xlsx"
        stored.write_bytes(b"changed")

        with self.assertRaisesRegex(ProjectStoreError, "发生变化"):
            self.store.restore_from_trash(running.summary.id)

        manifest = json.loads((trash_path / "manifest.json").read_text(encoding="utf-8"))
        self.assertIsNone(manifest.get("pending_restore"))
        self.assertIsNone(self.store.get_batch(running.summary.id))
        self.assertTrue(stored.is_file())

    def test_restore_rejects_unregistered_extra_file(self) -> None:
        running, trash_path = self._trashed_successful_batch()
        extra = trash_path / "batch" / "处理结果" / "未登记结果.xlsx"
        extra.write_bytes(b"extra")

        with self.assertRaisesRegex(ProjectStoreError, "未登记文件"):
            self.store.restore_from_trash(running.summary.id)

        manifest = json.loads((trash_path / "manifest.json").read_text(encoding="utf-8"))
        self.assertIsNone(manifest.get("pending_restore"))
        self.assertTrue(extra.is_file())

    def test_read_only_project_can_list_trash_but_cannot_restore(self) -> None:
        running, _trash_path = self._trashed_successful_batch()
        read_only = ProjectStore.open(self.project_root)
        try:
            self.assertFalse(read_only.writable)
            self.assertEqual(read_only.list_trash_details()[0].summary.id, running.summary.id)
            with self.assertRaisesRegex(ProjectStoreError, "另一个窗口|只读"):
                read_only.restore_from_trash(running.summary.id)
        finally:
            read_only.close()
        self.assertIsNone(self.store.get_batch(running.summary.id))

    def test_pending_restore_is_completed_after_reopen(self) -> None:
        running, trash_path = self._trashed_successful_batch()

        class SimulatedProcessStop(BaseException):
            pass

        with patch.object(
            self.store,
            "_recover_pending_restore",
            side_effect=SimulatedProcessStop("模拟写入恢复记录后中断"),
        ):
            with self.assertRaises(SimulatedProcessStop):
                self.store.restore_from_trash(running.summary.id)

        pending = json.loads((trash_path / "manifest.json").read_text(encoding="utf-8"))
        self.assertIsInstance(pending.get("pending_restore"), dict)
        self.store.close()
        self.store = ProjectStore.open(self.project_root)

        restored = self.store.get_batch(running.summary.id)
        self.assertIsNotNone(restored)
        self.assertEqual(restored.files[0].path(self.store.workspace).read_bytes(), b"restore")
        self.assertEqual(self.store.list_trash(), ())

    def test_restore_after_directory_move_is_completed_after_reopen(self) -> None:
        running, trash_path = self._trashed_successful_batch()
        active_manifest = self.store._manifest_path(running.summary.id)
        real_write_json = project_store_module._write_json

        class SimulatedProcessStop(BaseException):
            pass

        def stop_before_active_manifest(path: Path, payload: dict) -> None:
            if path == active_manifest and payload.get("pending_restore") is None:
                raise SimulatedProcessStop("模拟目录移动后中断")
            real_write_json(path, payload)

        with patch.object(
            project_store_module,
            "_write_json",
            side_effect=stop_before_active_manifest,
        ):
            with self.assertRaises(SimulatedProcessStop):
                self.store.restore_from_trash(running.summary.id)

        pending = json.loads((trash_path / "manifest.json").read_text(encoding="utf-8"))
        target_uploads = self.project_root / pending["pending_restore"]["target_directories"][CATEGORY_UPLOADS]
        self.assertTrue(target_uploads.is_dir())
        self.assertFalse((trash_path / "batch").exists())
        self.store.close()
        self.store = ProjectStore.open(self.project_root)

        restored = self.store.get_batch(running.summary.id)
        self.assertIsNotNone(restored)
        self.assertEqual(restored.files[0].path(self.store.workspace).read_bytes(), b"restore")
        self.assertEqual(self.store.list_trash(), ())

    def test_pending_import_is_completed_after_reopen(self) -> None:
        source = self.sources / "恢复导入.xlsx"
        source.write_bytes(b"recover")
        draft = self._draft()
        real_write = self.store._write_active_manifest
        saw_pending = False

        def fail_final(batch_id, manifest):
            nonlocal saw_pending
            if manifest.get("pending_import"):
                saw_pending = True
                return real_write(batch_id, manifest)
            if saw_pending:
                raise OSError("模拟发布后断电")
            return real_write(batch_id, manifest)

        with patch.object(self.store, "_write_active_manifest", side_effect=fail_final):
            with self.assertRaises(OSError):
                self.store.import_sources(draft.summary.id, [source])
        self.store.close()
        self.store = ProjectStore.open(self.project_root)
        recovered = self.store.get_batch(draft.summary.id)
        self.assertEqual(len(recovered.files), 1)
        self.assertEqual(recovered.files[0].path(self.store.workspace).read_bytes(), b"recover")

    def test_pending_empty_directory_snapshot_is_completed_after_reopen(self) -> None:
        source = self.sources / "待恢复空目录"
        (source / "001" / "空子目录").mkdir(parents=True)
        draft = self._draft(tool_name="人员资料文件夹改名")
        real_write = self.store._write_active_manifest
        saw_pending = False

        def fail_final(batch_id, manifest):
            nonlocal saw_pending
            if manifest.get("pending_import"):
                saw_pending = True
                return real_write(batch_id, manifest)
            if saw_pending:
                raise OSError("模拟文件夹发布后断电")
            return real_write(batch_id, manifest)

        with patch.object(self.store, "_write_active_manifest", side_effect=fail_final):
            with self.assertRaises(OSError):
                self.store.import_directory_snapshot(
                    draft.summary.id,
                    source,
                    role="input_path",
                )
        self.store.close()
        self.store = ProjectStore.open(self.project_root)

        recovered = self.store.get_batch(draft.summary.id)
        assert recovered is not None
        recovered_source = recovered.directories[CATEGORY_UPLOADS] / source.name
        self.assertTrue((recovered_source / "001" / "空子目录").is_dir())
        self.assertTrue(self.store.verify_batch_files(draft.summary.id))

    def test_visible_directory_import_journal_completes_after_reopen(self) -> None:
        source = self.sources / "待恢复资料"
        source.mkdir()
        (source / "一.xlsx").write_bytes(b"one")
        (source / "二.xlsx").write_bytes(b"two")
        real_publish = project_store_module._publish_no_replace
        publish_count = 0

        class SimulatedProcessStop(BaseException):
            pass

        def stop_during_second_publish(staging: Path, destination: Path) -> None:
            nonlocal publish_count
            publish_count += 1
            if publish_count == 2:
                raise SimulatedProcessStop("模拟发布期间进程中断")
            real_publish(staging, destination)

        with patch.object(
            project_store_module,
            "_publish_no_replace",
            side_effect=stop_during_second_publish,
        ):
            with self.assertRaises(SimulatedProcessStop):
                self.store.import_common_sources([source])

        journals = list(self.store.staging_dir.glob("*/operation.json"))
        self.assertEqual(len(journals), 1)
        visible_root = self.store.workspace.common_root / source.name
        self.assertEqual(len(list(visible_root.glob("*.xlsx"))), 1)

        self.store.close()
        self.store = ProjectStore.open(self.project_root)

        self.assertEqual((visible_root / "一.xlsx").read_bytes(), b"one")
        self.assertEqual((visible_root / "二.xlsx").read_bytes(), b"two")
        self.assertFalse(any(self.store.staging_dir.iterdir()))

    def test_visible_directory_import_journal_is_also_recovered_by_refresh(self) -> None:
        source = self.sources / "刷新恢复.xlsx"
        source.write_bytes(b"refresh")

        class SimulatedProcessStop(BaseException):
            pass

        with patch.object(
            self.store,
            "_recover_visible_directory_import",
            side_effect=SimulatedProcessStop("模拟写入日志后中断"),
        ):
            with self.assertRaises(SimulatedProcessStop):
                self.store.import_common_sources([source])

        self.assertEqual(len(list(self.store.staging_dir.glob("*/operation.json"))), 1)
        self.store.refresh()
        self.assertEqual((self.store.workspace.common_root / source.name).read_bytes(), b"refresh")
        self.assertFalse(any(self.store.staging_dir.iterdir()))

    def test_visible_directory_recovery_conflict_never_overwrites(self) -> None:
        source = self.sources / "冲突恢复"
        source.mkdir()
        (source / "一.xlsx").write_bytes(b"one")
        (source / "二.xlsx").write_bytes(b"two")
        real_publish = project_store_module._publish_no_replace
        publish_count = 0

        class SimulatedProcessStop(BaseException):
            pass

        def stop_during_second_publish(staging: Path, destination: Path) -> None:
            nonlocal publish_count
            publish_count += 1
            if publish_count == 2:
                raise SimulatedProcessStop("模拟发布期间进程中断")
            real_publish(staging, destination)

        with patch.object(
            project_store_module,
            "_publish_no_replace",
            side_effect=stop_during_second_publish,
        ):
            with self.assertRaises(SimulatedProcessStop):
                self.store.import_common_sources([source])

        journal_path = next(self.store.staging_dir.glob("*/operation.json"))
        journal = json.loads(journal_path.read_text(encoding="utf-8"))
        published = [
            self.project_root / item["destination_relative"]
            for item in journal["items"]
            if (self.project_root / item["destination_relative"]).exists()
        ]
        self.assertEqual(len(published), 1)
        published[0].write_bytes(b"external-change")
        self.store.close()

        with self.assertRaisesRegex(ProjectStoreError, "冲突|未覆盖"):
            ProjectStore.open(self.project_root)

        self.assertEqual(published[0].read_bytes(), b"external-change")
        self.assertTrue(journal_path.is_file())

    def test_visible_directory_recovery_rejects_journal_path_escape(self) -> None:
        source = self.sources / "越界.xlsx"
        source.write_bytes(b"private")

        class SimulatedProcessStop(BaseException):
            pass

        with patch.object(
            self.store,
            "_recover_visible_directory_import",
            side_effect=SimulatedProcessStop("模拟写入日志后中断"),
        ):
            with self.assertRaises(SimulatedProcessStop):
                self.store.import_common_sources([source])

        journal_path = next(self.store.staging_dir.glob("*/operation.json"))
        journal = json.loads(journal_path.read_text(encoding="utf-8"))
        journal["items"][0]["destination_relative"] = "../项目外.xlsx"
        journal_path.write_text(json.dumps(journal, ensure_ascii=False), encoding="utf-8")
        self.store.close()

        with self.assertRaisesRegex(ProjectStoreError, "相对路径|越界"):
            ProjectStore.open(self.project_root)

        self.assertFalse((self.base / "项目外.xlsx").exists())
        self.assertTrue(journal_path.is_file())

    def test_new_folder_and_common_import_cannot_write_results_or_metadata(self) -> None:
        department = self.store.new_folder(COMMON_VISIBLE_DIR, "制度资料")
        source = self.sources / "制度.pdf"
        source.write_bytes(b"policy")
        imported = self.store.import_common_sources([source], subdirectory="制度资料")
        self.assertEqual(imported[0], department / "制度.pdf")
        self.assertEqual(imported[0].read_bytes(), b"policy")

        draft = self._draft()
        supplement_root = draft.directories[CATEGORY_SUPPLEMENTS]
        self.assertFalse(supplement_root.exists())
        supplement_folder = self.store.new_folder(supplement_root, "补交证明")
        self.assertEqual(supplement_folder, supplement_root / "补交证明")
        self.assertTrue(supplement_folder.is_dir())

        running = self._running()
        with self.assertRaisesRegex(ProjectStoreError, "处理结果"):
            self.store.import_to_directory(running.directories[CATEGORY_RESULTS], [source])
        with self.assertRaisesRegex(ProjectStoreError, "隐藏"):
            self.store.new_folder(PROJECT_METADATA_DIR, "错误目录")
        with self.assertRaisesRegex(ProjectStoreError, "保留名称|隐藏"):
            self.store.create_draft(
                group_name=PROJECT_METADATA_DIR,
                tool_id="unsafe",
                tool_name="测试工具",
            )

    def test_supplement_directory_is_created_only_on_first_import(self) -> None:
        draft = self._draft()
        self.assertFalse(draft.directories[CATEGORY_SUPPLEMENTS].exists())
        source = self.sources / "补充.xlsx"
        source.write_bytes(b"supplement")
        records = self.store.import_sources(
            draft.summary.id,
            [source],
            category=CATEGORY_SUPPLEMENTS,
        )
        self.assertEqual(len(records), 1)
        self.assertTrue(draft.directories[CATEGORY_SUPPLEMENTS].is_dir())

    def test_supplements_can_be_added_after_every_terminal_status(self) -> None:
        success_running = self._running()
        terminal_batches = [self.store.mark_success(success_running.summary.id)]
        failed_running = self._running()
        terminal_batches.append(self.store.mark_failed(failed_running.summary.id, "测试失败"))
        stopped_running = self._running()
        terminal_batches.append(self.store.mark_stopped(stopped_running.summary.id))

        for index, terminal in enumerate(terminal_batches, start=1):
            with self.subTest(status=terminal.summary.status):
                source = self.sources / f"补交材料{index}.xlsx"
                content = f"supplement-{index}".encode()
                source.write_bytes(content)
                records = self.store.import_sources(
                    terminal.summary.id,
                    [source],
                    category=CATEGORY_SUPPLEMENTS,
                    role="manual_supplement",
                )
                self.assertEqual(len(records), 1)
                self.assertEqual(records[0].sha256, hashlib.sha256(content).hexdigest())
                self.assertEqual(records[0].path(self.store.workspace).read_bytes(), content)
                refreshed = self.store.get_batch(terminal.summary.id)
                self.assertEqual(refreshed.summary.status, terminal.summary.status)
                self.assertTrue(self.store.verify_batch_files(terminal.summary.id))


if __name__ == "__main__":
    unittest.main()
