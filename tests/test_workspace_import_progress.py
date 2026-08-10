from __future__ import annotations

import queue
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from hr_toolkit.gui import HRToolkitApp
from hr_toolkit.project_store import ImportCancelled, ImportProgress


class _Value:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class WorkspaceImportProgressTests(unittest.TestCase):
    def _write_app(self) -> HRToolkitApp:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.root = Mock()
        app.project_store = object()
        app.current_project_path = Path("/tmp/project")
        app._workspace_project_generation = 7
        app._workspace_write_token = 0
        app._workspace_write_tasks = {}
        app._workspace_write_progress = {}
        app._workspace_write_progress_lock = threading.Lock()
        app._workspace_write_callbacks = {}
        app._workspace_queue = queue.Queue()
        app._workspace_recovery_blocked = False
        app._workspace_recovery_error = None
        app._workspace_close_requested = False
        app._tool_running = False
        app._project_batch_by_token = {}
        app._update_workspace_action_states = Mock()
        app._open_workspace_import_progress = Mock()
        return app

    def test_background_progress_keeps_only_latest_snapshot(self) -> None:
        app = self._write_app()
        latest = ImportProgress(phase="copying", files_completed=500, files_total=500)

        def write(_cancelled, report) -> str:
            for index in range(500):
                report(ImportProgress(phase="copying", files_completed=index, files_total=500))
            report(latest)
            return "saved"

        app._workspace_run_write("导入文件", write, progress_target=Path("/tmp/project/共用资料"))
        status, generation, payload = app._workspace_queue.get(timeout=5)

        self.assertEqual((status, generation), ("write_changed", 7))
        self.assertEqual(payload[-1], "saved")
        self.assertEqual(tuple(app._workspace_write_progress.values()), (latest,))
        app._open_workspace_import_progress.assert_called_once()

    def test_import_cancelled_has_its_own_terminal_state(self) -> None:
        app = self._write_app()

        class DerivedImportCancelled(ImportCancelled):
            pass

        def write(_cancelled, _report):
            raise DerivedImportCancelled("本次导入已停止。")

        app._workspace_run_write("导入文件", write, progress_target=Path("/tmp/project/共用资料"))
        status, _generation, _payload = app._workspace_queue.get(timeout=5)
        self.assertEqual(status, "write_cancelled")

    def test_exception_name_alone_does_not_count_as_import_cancellation(self) -> None:
        app = self._write_app()
        fake_cancelled = type("ImportCancelled", (RuntimeError,), {})

        def write(_cancelled, _report):
            raise fake_cancelled("not the backend cancellation type")

        app._workspace_run_write("导入文件", write, progress_target=Path("/tmp/project/共用资料"))
        status, _generation, _payload = app._workspace_queue.get(timeout=5)
        self.assertEqual(status, "write_error")

    def test_cancel_sets_event_but_finalizing_cannot_be_cancelled(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        event = threading.Event()
        app._workspace_import_token = 4
        app._workspace_import_phase = "copying"
        app._workspace_write_tasks = {4: (event, object())}
        app._workspace_write_progress = {}
        app._workspace_write_progress_lock = threading.Lock()
        app._workspace_import_title_var = _Value()
        app._workspace_import_subtitle_var = _Value()
        app._workspace_import_cancel_button = Mock()
        app._workspace_import_stage_labels = [Mock(), Mock(), Mock()]

        app._cancel_workspace_import()
        self.assertTrue(event.is_set())
        self.assertIn("安全停止", app._workspace_import_title_var.get())
        app._workspace_import_cancel_button.configure.assert_called_with(text="正在停止…", state="disabled")

        second = threading.Event()
        app._workspace_write_tasks[4] = (second, object())
        app._workspace_import_phase = "copying"
        app._workspace_write_progress[4] = ImportProgress(phase="finalizing")
        app._cancel_workspace_import()
        self.assertFalse(second.is_set())
        self.assertEqual(app._workspace_import_phase, "finalizing")
        app._workspace_import_cancel_button.configure.assert_called_with(
            text="正在完成…",
            state="disabled",
        )

    def test_finalizing_state_disables_cancel_and_explains_safe_finish(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._workspace_import_token = 9
        app._workspace_import_phase = "copying"
        app._workspace_import_stage_labels = [Mock(), Mock(), Mock()]
        app._workspace_import_cancel_button = Mock()
        app._workspace_import_progress_canvas = None
        app._workspace_import_title_var = _Value()
        app._workspace_import_subtitle_var = _Value()
        app._workspace_import_state_var = _Value()
        app._workspace_import_name_var = _Value()
        app._workspace_import_left_var = _Value()
        app._workspace_import_middle_var = _Value()
        app._workspace_import_elapsed_var = _Value()
        app._workspace_import_safety_title_var = _Value()
        app._workspace_import_safety_text_var = _Value()
        app._workspace_import_started_at = time.monotonic()

        app._render_workspace_import_progress(
            9,
            ImportProgress(
                phase="finalizing",
                files_scanned=2,
                files_completed=2,
                files_total=2,
                bytes_copied=2048,
                bytes_total=2048,
            ),
        )

        self.assertEqual(app._workspace_import_title_var.get(), "正在完成保存")
        self.assertIn("安全登记", app._workspace_import_safety_title_var.get())
        self.assertIn("自动恢复", app._workspace_import_safety_text_var.get())
        app._workspace_import_cancel_button.configure.assert_called_with(text="正在完成…", state="disabled")

    def test_terminal_renders_latest_snapshot_then_shows_success(self) -> None:
        app = self._write_app()
        store = app.project_store
        snapshot = ImportProgress(
            phase="finalizing",
            files_completed=2,
            files_total=2,
            bytes_copied=2048,
            bytes_total=2048,
        )
        app._workspace_write_tasks = {5: (threading.Event(), store)}
        app._workspace_write_callbacks = {5: (None, None, None)}
        app._workspace_queue.put(("write_changed", 7, (5, store, "saved")))
        app._workspace_latest_progress = Mock(return_value=snapshot)
        app._render_workspace_import_progress = Mock()
        app._show_workspace_import_success = Mock()
        app._refresh_workspace_tree = Mock()
        app._update_sidebar_project_summary = Mock()

        app._poll_workspace_queue()

        app._render_workspace_import_progress.assert_called_once_with(5, snapshot)
        app._show_workspace_import_success.assert_called_once_with(5)

    def test_success_state_is_visible_before_delayed_close(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.root = Mock()
        app._workspace_import_token = 11
        app._workspace_import_phase = "copying"
        app._workspace_import_stage_labels = [Mock(), Mock(), Mock()]
        app._workspace_import_cancel_button = Mock()
        app._workspace_import_progress_canvas = None
        app._workspace_import_success_job = None
        app._workspace_import_title_var = _Value()
        app._workspace_import_subtitle_var = _Value()
        app._workspace_import_state_var = _Value()
        app._workspace_import_name_var = _Value()
        app._workspace_import_safety_title_var = _Value()
        app._workspace_import_safety_text_var = _Value()

        app._show_workspace_import_success(11)

        self.assertEqual(app._workspace_import_title_var.get(), "完成保存")
        self.assertIn("安全保存", app._workspace_import_subtitle_var.get())
        self.assertEqual(app._workspace_import_state_var.get(), "已完成")
        app._workspace_import_cancel_button.configure.assert_called_with(
            text="已完成",
            state="disabled",
        )
        app.root.after.assert_called_once()

    def test_finalizing_error_refreshes_store_and_uses_recovered_terminal(self) -> None:
        app = self._write_app()
        store = Mock()
        app.project_store = store

        def write(_cancelled, report):
            report(ImportProgress(phase="finalizing"))
            raise RuntimeError("commit failed")

        app._workspace_run_write("导入文件", write, progress_target=Path("/tmp/project/共用资料"))
        terminal = app._workspace_queue.get(timeout=5)

        self.assertEqual(terminal[0], "write_recovered")
        store.refresh.assert_called_once_with()
        app._workspace_queue.put(terminal)
        app._close_workspace_import_progress = Mock()
        app._refresh_workspace_tree = Mock()
        app._update_sidebar_project_summary = Mock()

        with (
            patch("hr_toolkit.gui.messagebox.showwarning") as showwarning,
            patch("hr_toolkit.gui.messagebox.showerror") as showerror,
            patch("hr_toolkit.gui.runlog.log_exception"),
            patch("hr_toolkit.gui.runlog.log_line"),
        ):
            app._poll_workspace_queue()

        self.assertFalse(app._workspace_recovery_blocked)
        app._refresh_workspace_tree.assert_called_once_with()
        app._update_sidebar_project_summary.assert_called_once_with()
        self.assertEqual(showwarning.call_args.args[0], "项目已恢复到安全状态")
        self.assertIn("请先检查右侧是否已出现资料，再决定是否重试", showwarning.call_args.args[1])
        showerror.assert_not_called()

    def test_failed_finalizing_recovery_blocks_writes_and_project_switching(self) -> None:
        app = self._write_app()
        store = Mock()
        store.refresh.side_effect = OSError("recovery failed")
        app.project_store = store

        def write(_cancelled, report):
            report(ImportProgress(phase="finalizing"))
            raise RuntimeError("commit failed")

        app._workspace_run_write("导入文件", write, progress_target=Path("/tmp/project/共用资料"))
        terminal = app._workspace_queue.get(timeout=5)
        self.assertEqual(terminal[0], "write_recovery_blocked")
        app._workspace_queue.put(terminal)
        app._close_workspace_import_progress = Mock()

        with (
            patch("hr_toolkit.gui.messagebox.showerror") as showerror,
            patch("hr_toolkit.gui.runlog.log_exception"),
        ):
            app._poll_workspace_queue()

        self.assertTrue(app._workspace_recovery_blocked)
        self.assertEqual(app._workspace_recovery_error, "recovery failed")
        self.assertTrue(app._project_change_is_blocked())
        self.assertIn("重新打开当前项目", showerror.call_args.args[1])
        app._update_workspace_action_states.assert_called()

        callback = Mock()
        with patch("hr_toolkit.gui.messagebox.showerror") as blocked_error:
            app._workspace_run_write("新建文件夹", callback)
        callback.assert_not_called()
        blocked_error.assert_called_once()

        with patch("hr_toolkit.gui.messagebox.showerror") as tool_error:
            app._run_current_tool()
        tool_error.assert_called_once()

    def test_cancelled_poll_never_shows_failure_dialog(self) -> None:
        app = self._write_app()
        store = app.project_store
        event = threading.Event()
        app._workspace_write_tasks = {3: (event, store)}
        app._workspace_write_callbacks = {3: (None, Mock(), None)}
        app._workspace_queue.put(
            ("write_cancelled", 7, (3, store, "导入文件", ImportCancelled("stopped")))
        )
        app._close_workspace_import_progress = Mock()
        app._refresh_workspace_tree = Mock()
        app._update_sidebar_project_summary = Mock()

        with (
            patch("hr_toolkit.gui.messagebox.showerror") as showerror,
            patch("hr_toolkit.gui.runlog.log_line"),
        ):
            app._poll_workspace_queue()

        showerror.assert_not_called()
        self.assertFalse(app._workspace_write_tasks)
        app._close_workspace_import_progress.assert_called_once_with(token=3, force=True)


if __name__ == "__main__":
    unittest.main()
