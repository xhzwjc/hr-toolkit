from __future__ import annotations

import json
import os
from pathlib import Path
import tempfile
import threading
import unittest
from datetime import date
from types import SimpleNamespace
from unittest.mock import patch


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("HR_TOOLKIT_SKIP_UPDATE", "1")

try:
    from hr_toolkit.gui_qt.compat import QCoreApplication
    from hr_toolkit.gui_qt.controller import AppController
    from hr_toolkit.gui_qt.models import LogModel
except ImportError:
    QCoreApplication = None
    AppController = None
    LogModel = None


@unittest.skipUnless(AppController is not None, "PySide GUI runtime is not installed")
class QtControllerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.application = QCoreApplication.instance() or QCoreApplication([])

    def controller(self):
        value = AppController()
        value._save_workspace_preferences = lambda: None
        return value

    def test_date_presets_keep_legacy_calendar_ranges(self) -> None:
        controller = self.controller()
        controller.selectTool("data_statistics")
        controller.applyDatePreset("week", "this_week")
        state = controller._form_states[("data_statistics", "default")]
        today = date.today()
        start = date.fromisoformat(state["week_start"])
        end = date.fromisoformat(state["week_end"])
        self.assertEqual(start.weekday(), 0)
        self.assertEqual(end.weekday(), 6)
        self.assertLessEqual(start, today)
        self.assertGreaterEqual(end, today)

        controller.applyDatePreset("month", "this_month")
        self.assertEqual(date.fromisoformat(state["month_start"]).day, 1)
        self.assertEqual(date.fromisoformat(state["month_end"]).month, today.month)
        controller.applyDatePreset("month", "clear")
        self.assertEqual(state["month_start"], "")
        self.assertEqual(state["month_end"], "")
        controller.close()

    def test_legacy_history_reuse_accepts_every_current_navigation_tool(self) -> None:
        for tool_id in (
            "social_security",
            "insurance_ledger",
            "data_statistics",
            "salary_split",
            "salary_merge",
            "material_collector",
        ):
            with self.subTest(tool_id=tool_id):
                controller = self.controller()
                controller._history_selected = SimpleNamespace(
                    summary=SimpleNamespace(tool_id=tool_id, mode=None),
                    inputs=(),
                )
                controller.reuseHistory()
                self.assertEqual(controller.currentTool, tool_id)
                controller.close()

    def test_run_button_text_tracks_tool_switches(self) -> None:
        controller = self.controller()
        changes = []
        controller.runButtonTextChanged.connect(lambda: changes.append(controller.runButtonText))

        self.assertEqual(controller.runButtonText, "生成报表")
        controller.selectTool("material_collector")

        self.assertEqual(controller.runButtonText, "开始打包")
        self.assertIn("开始打包", changes)
        controller.close()

    def test_run_without_project_reports_required_next_step(self) -> None:
        controller = self.controller()
        notifications = []
        controller.notificationRequested.connect(
            lambda *args: notifications.append(args)
        )

        controller.runOrCancel()

        self.assertFalse(controller.busy)
        self.assertEqual(len(notifications), 1)
        self.assertEqual(notifications[0][0], "请先打开工作项目")
        self.assertIn("新建或打开", notifications[0][1])
        controller.close()

    def test_workspace_selection_details_follow_model_refresh(self) -> None:
        controller = self.controller()
        selected = {
            "name": "名单.xlsx",
            "path": str(Path.cwd() / "名单.xlsx"),
            "isDir": False,
            "detail": "XLSX",
        }
        controller._workspace_items = [selected]
        controller.selectWorkspaceRow(0)

        self.assertTrue(controller.workspaceSelectionAvailable)
        self.assertEqual(controller.workspaceSelectedName, "名单.xlsx")
        self.assertEqual(controller.workspaceSelectedDetail, "XLSX")

        controller._workspace_generation = 2
        controller._apply_workspace_items(2, [])
        self.assertFalse(controller.workspaceSelectionAvailable)
        controller.close()

    def test_workspace_folder_toggle_updates_rows_without_model_reset(self) -> None:
        controller = self.controller()
        root_path = Path.cwd() / "共用资料"
        child_path = root_path / "名单.xlsx"
        sibling_path = Path.cwd() / "处理结果"
        controller._workspace_generation = 7
        controller._workspace_items = [
            {
                "name": "共用资料",
                "path": str(root_path),
                "isDir": True,
                "depth": 0,
                "expanded": True,
                "hasChildren": True,
                "detail": "文件夹",
            },
            {
                "name": "名单.xlsx",
                "path": str(child_path),
                "isDir": False,
                "depth": 1,
                "expanded": False,
                "hasChildren": False,
                "detail": "XLSX",
            },
            {
                "name": "处理结果",
                "path": str(sibling_path),
                "isDir": True,
                "depth": 0,
                "expanded": False,
                "hasChildren": True,
                "detail": "文件夹",
            },
        ]
        controller.workspaceModel.set_items(controller._workspace_items)
        resets = []
        removals = []
        insertions = []
        controller.workspaceModel.modelReset.connect(lambda: resets.append(True))
        controller.workspaceModel.rowsRemoved.connect(lambda *_args: removals.append(True))
        controller.workspaceModel.rowsInserted.connect(lambda *_args: insertions.append(True))

        controller.toggleWorkspaceRow(0)

        self.assertEqual(resets, [])
        self.assertEqual(removals, [True])
        self.assertEqual(controller.workspaceModel.rowCount(), 2)
        self.assertFalse(controller.workspaceModel.item_at(0)["expanded"])
        self.assertEqual(controller.workspaceModel.item_at(1)["name"], "处理结果")

        with patch("hr_toolkit.gui_qt.controller.threading.Thread") as thread:
            controller.toggleWorkspaceRow(0)
            thread.return_value.start.assert_called_once_with()
        controller._apply_workspace_children(
            7,
            0,
            str(root_path),
            0,
            [
                {
                    "name": "名单.xlsx",
                    "path": str(child_path),
                    "isDir": False,
                    "depth": 1,
                    "expanded": False,
                    "hasChildren": False,
                    "detail": "XLSX",
                }
            ],
        )

        self.assertEqual(resets, [])
        self.assertEqual(insertions, [True])
        self.assertEqual(controller.workspaceModel.rowCount(), 3)
        self.assertTrue(controller.workspaceModel.item_at(0)["expanded"])
        self.assertEqual(controller.workspaceModel.item_at(1)["name"], "名单.xlsx")
        controller.close()

    def test_qt_tutorial_uses_the_complete_legacy_tk_content(self) -> None:
        controller = self.controller()
        groups = controller.tutorialGroups
        items = [item for group in groups for item in group["items"]]
        self.assertEqual(
            [item["label"] for item in items],
            [
                "社保明细与汇总",
                "保险台账与预警",
                "考勤与周月报",
                "工资表拆分",
                "多月工资合并",
                "异动表汇总",
                "花名册更新",
                "档案入库",
                "档案表生成",
                "员工资料打包",
                "资料文件夹改名",
            ],
        )
        statistics = next(item for item in items if item["toolId"] == "data_statistics")
        copy = [line["text"] for line in statistics["lines"]]
        self.assertIn("容易疑惑1：", copy[6])
        self.assertIn("容易疑惑2：", copy[7])
        self.assertEqual(statistics["lines"][-1]["style"], "warning")
        controller.close()

    def test_close_requires_confirmation_while_background_work_is_active(self) -> None:
        controller = self.controller()
        prompts = []
        controller.confirmationRequested.connect(lambda *args: prompts.append(args))
        controller._busy = True
        self.assertFalse(controller.requestClose())
        self.assertEqual(len(prompts), 1)
        self.assertFalse(controller.requestClose())
        self.assertEqual(len(prompts), 1)
        token = prompts[0][2]
        controller.confirmAction(token, False)
        controller._busy = False
        self.assertTrue(controller.requestClose())
        self.assertTrue(controller._closed)

    def test_close_waits_for_project_and_storage_workers(self) -> None:
        for attribute in ("_project_opening", "_history_busy", "_trash_busy"):
            with self.subTest(attribute=attribute):
                controller = self.controller()
                prompts = []
                controller.confirmationRequested.connect(lambda *args: prompts.append(args))
                setattr(controller, attribute, True)
                self.assertFalse(controller.requestClose())
                self.assertEqual(len(prompts), 1)
                controller.confirmAction(prompts[0][2], False)
                setattr(controller, attribute, False)
                controller.close()

    def test_cancelling_workspace_import_does_not_cancel_update_download(self) -> None:
        controller = self.controller()
        workspace_cancel = SimpleNamespace(called=False)
        update_cancel = SimpleNamespace(called=False)
        workspace_cancel.set = lambda: setattr(workspace_cancel, "called", True)
        update_cancel.set = lambda: setattr(update_cancel, "called", True)
        controller._workspace_cancel_event = workspace_cancel
        controller._update_cancel_event = update_cancel
        controller.cancelWorkspaceImport()
        self.assertTrue(workspace_cancel.called)
        self.assertFalse(update_cancel.called)
        controller._workspace_cancel_event = None
        controller._update_cancel_event = None
        controller.close()

    def test_log_model_append_batch_preserves_order_and_truncation(self) -> None:
        model = LogModel()
        items = [{"time": "12:00:00", "text": f"msg_{i}", "level": "info"} for i in range(15)]
        model.append_batch(items[:10], maximum=10)
        self.assertEqual(len(model), 10)
        self.assertEqual(model.item_at(0)["text"], "msg_0")
        self.assertEqual(model.item_at(9)["text"], "msg_9")

        # Adding 5 more with maximum=10 should drop oldest 5 and keep newest 10
        model.append_batch(items[10:], maximum=10)
        self.assertEqual(len(model), 10)
        self.assertEqual(model.item_at(0)["text"], "msg_5")
        self.assertEqual(model.item_at(9)["text"], "msg_14")

    def test_controller_log_batching_and_synchronous_flush(self) -> None:
        controller = self.controller()
        controller._clear_logs()
        self.assertEqual(len(controller.logModel), 0)

        # Emitting 10 logs buffers them before flush
        for i in range(10):
            controller._append_log(f"test_log_{i}", "info")
        self.assertEqual(len(controller._log_buffer), 10)
        self.assertEqual(len(controller.logModel), 0)

        # Synchronous flush empties buffer and populates model in FIFO order
        controller._flush_logs()
        self.assertEqual(len(controller._log_buffer), 0)
        self.assertEqual(len(controller.logModel), 10)
        self.assertEqual(controller.logModel.item_at(0)["text"], "test_log_0")
        self.assertEqual(controller.logModel.item_at(9)["text"], "test_log_9")

        # _apply_run_finished forces synchronous flush
        controller._append_log("trailing_message", "success")
        self.assertEqual(len(controller._log_buffer), 1)
        controller._apply_run_finished()
        self.assertEqual(len(controller._log_buffer), 0)
        self.assertEqual(len(controller.logModel), 11)
        self.assertEqual(controller.logModel.item_at(10)["text"], "trailing_message")
        controller.close()

    def test_controller_log_timer_flush(self) -> None:
        import time

        controller = self.controller()
        controller._clear_logs()
        controller._append_log("timer_msg", "info")
        self.assertEqual(len(controller.logModel), 0)
        self.assertEqual(len(controller._log_buffer), 1)

        # Let Qt event loop process events until timer fires
        deadline = time.monotonic() + 0.5
        while time.monotonic() < deadline and len(controller.logModel) == 0:
            QCoreApplication.processEvents()
            time.sleep(0.01)

        self.assertEqual(len(controller.logModel), 1)
        self.assertEqual(controller.logModel.item_at(0)["text"], "timer_msg")
        controller.close()

    def test_controller_file_dialog_directory_memory_hierarchy_and_cancellation(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            saved_dir = tmp_path / "saved"
            saved_dir.mkdir()
            project_dir = tmp_path / "project"
            project_dir.mkdir()
            other_dir = tmp_path / "other"
            other_dir.mkdir()
            sample_file = other_dir / "test.xlsx"
            sample_file.write_text("dummy", encoding="utf-8")

            controller = self.controller()

            # 1. When last_selected_dir is valid, it takes priority
            controller._last_selected_dir = saved_dir
            controller._project_path = project_dir
            self.assertEqual(controller._file_dialog_initial_dir(), str(saved_dir))

            # 2. When last_selected_dir is deleted/invalid, falls back to project_path
            controller._last_selected_dir = tmp_path / "non_existent"
            self.assertEqual(controller._file_dialog_initial_dir(), str(project_dir))

            # 3. When project_path is also None, falls back to desktop / home
            controller._project_path = None
            fallback = controller._file_dialog_initial_dir()
            self.assertTrue(Path(fallback).is_dir())

            # 4. role="new_project" starts at defaultProjectParent
            self.assertEqual(
                controller._file_dialog_initial_dir(role="new_project"),
                str(Path(controller.defaultProjectParent).expanduser().absolute()),
            )

            # 5. Successful selection remembers folder (or parent folder for file)
            controller._remember_file_dialog_path(str(sample_file))
            self.assertEqual(controller._last_selected_dir, other_dir)

            # 6. Cancelled dialog (empty string/list/None) DOES NOT overwrite memory
            controller._remember_file_dialog_path("")
            self.assertEqual(controller._last_selected_dir, other_dir)
            controller._remember_file_dialog_path([])
            self.assertEqual(controller._last_selected_dir, other_dir)
            controller._remember_file_dialog_path(None)
            self.assertEqual(controller._last_selected_dir, other_dir)

            controller.close()

    def test_controller_file_dialog_directory_memory_persists_in_settings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            settings_file = tmp_path / "workspace-ui.json"
            chosen_dir = tmp_path / "chosen"
            chosen_dir.mkdir()

            with patch.object(AppController, "_settings_path", return_value=settings_file):
                controller = AppController()
                controller._remember_file_dialog_path(chosen_dir)

                self.assertTrue(settings_file.is_file())
                data = json.loads(settings_file.read_text(encoding="utf-8"))
                self.assertEqual(data.get("last_selected_dir"), str(chosen_dir))

                # New controller instance loads it on start()
                controller2 = AppController()
                controller2.start()
                self.assertEqual(controller2._last_selected_dir, chosen_dir)
                self.assertEqual(controller2._file_dialog_initial_dir(), str(chosen_dir))

                controller.close()
                controller2.close()

    def test_controller_trash_read_only_protection_and_restore_selection(self) -> None:
        controller = self.controller()
        fake_store = SimpleNamespace(writable=False, list_trash_details=lambda: [])
        controller._project_store = fake_store
        controller._trash_selected_id = "batch-1"

        with patch("threading.Thread") as mock_thread:
            controller.restoreSelectedTrash()
            # Read-only store prevents launching restore worker thread
            mock_thread.assert_not_called()

        # Selection retention test
        fake_detail = SimpleNamespace(
            summary=SimpleNamespace(
                id="batch-1",
                group_name="薪酬管理",
                tool_name="工资表拆分",
                business_description="七月",
                business_period="2026-07",
                directory_name="dir",
                status="success",
                deleted_at="2026-08-01T00:00:00Z",
            ),
            original_relative_path="rel",
            upload_count=1,
            result_count=1,
            supplement_count=0,
            total_size_bytes=100,
        )
        controller._apply_trash_list(controller._trash_generation, [fake_detail], "")
        self.assertEqual(controller.trashSelectedId, "batch-1")

        # Simulate restore failure
        notifications = []
        controller.notificationRequested.connect(lambda *args: notifications.append(args))
        controller._apply_trash_action(False, "权限不足")
        self.assertEqual(notifications[-1][0], "恢复没有完成")
        self.assertEqual(notifications[-1][1], "权限不足")
        controller._apply_trash_list(controller._trash_generation, [fake_detail], "")
        self.assertFalse(controller.trashBusy)
        self.assertEqual(controller.trashSelectedId, "batch-1")
        controller.close()

    def test_controller_workspace_search_and_scan_threading(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sub = root / "sub"
            sub.mkdir()
            (sub / "file.xlsx").touch()

            controller = self.controller()
            controller._project_path = root

            threads_started = []
            orig_start = threading.Thread.start

            def tracking_start(t_self):
                threads_started.append(t_self.name)
                return orig_start(t_self)

            with patch.object(threading.Thread, "start", side_effect=tracking_start, autospec=True):
                controller.refreshWorkspace()

            self.assertIn("HRToolkit-workspace-scan", threads_started)
            controller.close()

    def test_controller_workspace_import_cancellation_and_state(self) -> None:
        controller = self.controller()
        controller._workspace_busy = True
        cancel_event = threading.Event()
        controller._workspace_cancel_event = cancel_event

        self.assertFalse(cancel_event.is_set())
        controller.cancelWorkspaceImport()
        self.assertTrue(cancel_event.is_set())

        notifications = []
        controller.notificationRequested.connect(lambda *args: notifications.append(args))
        controller._apply_workspace_import_result(False, "资料导入已取消。")
        self.assertFalse(controller.workspaceBusy)
        self.assertIsNone(controller._workspace_cancel_event)
        self.assertEqual(notifications[-1][0], "导入未完成")
        self.assertEqual(notifications[-1][1], "资料导入已取消。")
        controller.close()

    def test_controller_material_preferences_and_presets(self) -> None:
        controller = self.controller()
        controller.selectTool("material_collector")

        state = controller._form_states[("material_collector", "default")]
        self.assertTrue(state.get("collect_all"))

        # Apply preset "入职材料" -> collect_all becomes False, materials updated
        controller.applyMaterialPreset("入职材料")
        self.assertFalse(state.get("collect_all"))
        self.assertEqual(state.get("material_types"), ["身份证", "劳动合同"])

        # Add custom material via submitTextAction
        controller.requestAddCustomMaterial()
        token = controller._pending_text_action
        self.assertIsNotNone(token)
        controller.submitTextAction(token, "体检报告")
        self.assertIn("体检报告", state.get("material_types"))
        self.assertIn("体检报告", controller._material_preferences.custom_materials)

        # Clear and select all
        controller.clearMaterials()
        self.assertEqual(state.get("material_types"), [])
        controller.selectAllMaterials()
        self.assertEqual(
            set(state.get("material_types")),
            set(controller._material_preferences.available_materials),
        )
        controller.close()

    def test_controller_create_project_validation_and_creation(self) -> None:
        import time

        controller = self.controller()
        notifications = []
        controller.notificationRequested.connect(lambda *args: notifications.append(args))

        # 1. Invalid project name rejects before creating thread
        controller.createProject("CON", "/some/path")
        self.assertEqual(notifications[-1][0], "无法创建项目")
        self.assertIn("Windows 系统保留名称", notifications[-1][1])

        # 2. Valid project creates project store in background
        with tempfile.TemporaryDirectory() as tmp:
            parent = Path(tmp)
            opened = []
            controller._projectOpened.connect(lambda gen, store, path: opened.append((path, store)))
            dummy_store = SimpleNamespace(close=lambda: None)
            with patch("hr_toolkit.project_store.ProjectStore.create", return_value=dummy_store):
                controller.createProject("测试项目", str(parent))
                for _ in range(50):
                    if opened:
                        break
                    time.sleep(0.05)
                    QCoreApplication.processEvents()
            self.assertEqual(len(opened), 1)
            target_path, store = opened[0]
            self.assertEqual(Path(target_path), parent / "测试项目")
            self.assertIs(store, dummy_store)

        controller.close()


if __name__ == "__main__":
    unittest.main()
