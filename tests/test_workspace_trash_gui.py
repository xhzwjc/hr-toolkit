from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from hr_toolkit.gui import (
    HRToolkitApp,
    _workspace_trash_deleted_text,
    _workspace_trash_dialog_height,
    _workspace_trash_group_tool,
    _workspace_trash_ignore_enter,
    _workspace_trash_matches,
    _workspace_trash_period_label,
    _workspace_trash_restore_location,
    _workspace_trash_title,
)


class _Value:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


def _trash_detail(
    batch_id: str = "batch-1",
    *,
    period: str = "2026-07",
    description: str = "七月薪酬",
    tool_name: str = "工资表拆分",
):
    return SimpleNamespace(
        summary=SimpleNamespace(
            id=batch_id,
            group_name="薪酬管理",
            business_period=period,
            business_description=description,
            tool_name=tool_name,
            directory_name="2026年7月",
            status="success",
            deleted_at="2026-08-10T04:30:00+00:00",
        ),
        original_relative_path="薪酬管理/工资表拆分/20260810_180000_七月薪酬_2026-07",
        upload_count=2,
        result_count=3,
        supplement_count=1,
        total_size_bytes=4096,
    )


class WorkspaceTrashPresentationTests(unittest.TestCase):
    def test_business_labels_and_search_are_hr_readable(self) -> None:
        detail = _trash_detail()

        self.assertEqual(_workspace_trash_period_label("2026-07"), "2026年7月")
        self.assertEqual(_workspace_trash_period_label("待确认"), "待确认")
        self.assertEqual(_workspace_trash_title(detail), "七月薪酬")
        self.assertEqual(_workspace_trash_group_tool(detail), "薪酬管理 · 工资表拆分")
        self.assertEqual(_workspace_trash_restore_location(detail), "薪酬管理 / 工资表拆分")
        for query in ("2026-07", "七月薪酬", "工资表", "薪酬管理"):
            with self.subTest(query=query):
                self.assertTrue(_workspace_trash_matches(detail, query))
        self.assertFalse(_workspace_trash_matches(detail, "社保台账"))
        self.assertRegex(_workspace_trash_deleted_text(detail.summary.deleted_at), r"^2026-08-10 \d{2}:\d{2}$")

        without_description = _trash_detail(description="")
        self.assertEqual(_workspace_trash_title(without_description), "2026年7月")

    def test_dialog_height_grows_when_possible_and_caps_on_small_screens(self) -> None:
        self.assertEqual(_workspace_trash_dialog_height(487, 548, 828), 548)
        self.assertEqual(_workspace_trash_dialog_height(487, 548, 500), 500)
        self.assertEqual(_workspace_trash_dialog_height(487, 420, 828), 487)

    def test_search_enter_is_consumed_without_restoring(self) -> None:
        self.assertEqual(_workspace_trash_ignore_enter(), "break")

    def test_cards_separate_business_identity_group_tool_and_status(self) -> None:
        detail = _trash_detail()
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._workspace_trash_list_body = Mock()
        app._workspace_trash_list_body.winfo_children.return_value = []
        app._workspace_trash_search_var = _Value()
        app._workspace_trash_details = (detail,)
        app._workspace_trash_selected_id = detail.summary.id
        app._select_workspace_trash_detail = Mock()
        app._format_history_size = Mock(return_value="4 KB")
        app.section_font = ("TkDefaultFont", 10)
        app.tiny_font = ("TkDefaultFont", 8)
        app.base_font = ("TkDefaultFont", 10)
        app._pad = lambda value, *_rest: value
        app._px = lambda value: value

        card = Mock()
        title_row = Mock()
        label_widgets = [Mock() for _index in range(5)]
        with (
            patch("hr_toolkit.gui.Frame", side_effect=(card, title_row)),
            patch("hr_toolkit.gui.Label", side_effect=label_widgets) as label_class,
        ):
            app._render_workspace_trash_cards()

        texts = [call.kwargs.get("text") for call in label_class.call_args_list]
        self.assertIn("七月薪酬", texts)
        self.assertIn("薪酬管理 · 工资表拆分", texts)
        self.assertIn("已完成", texts)
        moved_text = next(text for text in texts if str(text).startswith("移入时间："))
        self.assertNotIn("已完成", moved_text)

    def test_empty_and_search_empty_states_use_distinct_copy(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._workspace_trash_list_body = Mock()
        app._workspace_trash_list_body.winfo_children.return_value = []
        app._workspace_trash_search_var = _Value()
        app._workspace_trash_details = ()
        app._render_workspace_trash_detail = Mock()
        app.small_font = ("TkDefaultFont", 9)
        app._pad = lambda value, *_rest: value
        app._px = lambda value: value

        empty_label = Mock()
        with patch("hr_toolkit.gui.Label", return_value=empty_label) as label_class:
            app._render_workspace_trash_cards()

        self.assertIn("回收站是空的", label_class.call_args.kwargs["text"])
        empty_label.pack.assert_called_once()
        app._render_workspace_trash_detail.assert_called_once_with(None)

        app._workspace_trash_details = (_trash_detail(),)
        app._workspace_trash_search_var.set("没有这个记录")
        search_empty_label = Mock()
        with patch("hr_toolkit.gui.Label", return_value=search_empty_label) as label_class:
            app._render_workspace_trash_cards()

        self.assertEqual(label_class.call_args.kwargs["text"], "没有找到匹配的处理记录。")
        search_empty_label.pack.assert_called_once()


class WorkspaceTrashControllerTests(unittest.TestCase):
    def _controller(self, detail=None) -> HRToolkitApp:
        detail = detail or _trash_detail()
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.root = object()
        app.current_project_path = Path("/tmp/hr-project")
        app.project_store = Mock()
        app._workspace_project_read_only = False
        app._workspace_recovery_blocked = False
        app._tool_running = False
        app._project_batch_by_token = {}
        app._workspace_trash_details = (detail,)
        app._workspace_trash_selected_id = str(detail.summary.id)
        app._workspace_trash_search_var = _Value()
        app._workspace_trash_status_var = _Value()
        app._workspace_trash_restore_in_progress = False
        app._workspace_trash_restore_button = Mock()
        app._workspace_write_in_progress = Mock(return_value=False)
        return app

    def test_existing_dialog_is_focused_instead_of_duplicated(self) -> None:
        app = self._controller()
        window = Mock()
        window.winfo_exists.return_value = True
        app._workspace_trash_window = window

        with patch("hr_toolkit.gui.Toplevel") as create_window:
            app._open_workspace_trash_dialog()

        create_window.assert_not_called()
        window.lift.assert_called_once_with()
        window.focus_force.assert_called_once_with()

    def test_reload_keeps_selection_and_uses_backend_order_for_default(self) -> None:
        first = _trash_detail("first")
        second = _trash_detail("second", period="2026-06")
        app = self._controller(first)
        app.project_store.list_trash_details.return_value = (first, second)
        app._workspace_trash_selected_id = "second"
        app._render_workspace_trash_cards = Mock()

        app._reload_workspace_trash_details()
        self.assertEqual(app._workspace_trash_selected_id, "second")

        app._workspace_trash_selected_id = "missing"
        app._reload_workspace_trash_details()
        self.assertEqual(app._workspace_trash_selected_id, "first")

    def test_reload_failure_keeps_existing_items_and_selection(self) -> None:
        detail = _trash_detail()
        app = self._controller(detail)
        app.project_store.list_trash_details.side_effect = RuntimeError("磁盘暂时不可用")
        app._render_workspace_trash_cards = Mock()

        with patch("hr_toolkit.gui.runlog.log_exception"):
            app._reload_workspace_trash_details()

        self.assertEqual(app._workspace_trash_details, (detail,))
        self.assertEqual(app._workspace_trash_selected_id, detail.summary.id)
        self.assertIn("磁盘暂时不可用", app._workspace_trash_status_var.get())

    def test_read_only_project_can_view_but_cannot_restore(self) -> None:
        app = self._controller()
        app._workspace_project_read_only = True

        app._update_workspace_trash_restore_state()

        app._workspace_trash_restore_button.configure.assert_called_with(
            text="恢复到项目",
            state="disabled",
        )
        app._restore_selected_workspace_trash()
        self.assertIn("只读", app._workspace_trash_status_var.get())
        app.project_store.restore_from_trash.assert_not_called()

    def test_recovery_blocked_project_cannot_restore(self) -> None:
        app = self._controller()
        app._workspace_recovery_blocked = True

        app._update_workspace_trash_restore_state()

        app._workspace_trash_restore_button.configure.assert_called_with(
            text="恢复到项目",
            state="disabled",
        )
        app._restore_selected_workspace_trash()
        self.assertIn("重新打开", app._workspace_trash_status_var.get())
        app.project_store.restore_from_trash.assert_not_called()

    def test_detail_uses_business_location_without_timestamp_directory(self) -> None:
        detail = _trash_detail()
        app = self._controller(detail)
        app._workspace_trash_detail_title = Mock()
        app._workspace_trash_restore_path_var = _Value()
        app._workspace_trash_project_var = _Value()
        app._workspace_trash_notice_var = _Value()
        app.workspace_project_name = _Value("2026年7月人事工作")
        app.project_store.workspace = SimpleNamespace(name="2026年7月人事工作")
        app._update_workspace_trash_restore_state = Mock()

        app._render_workspace_trash_detail(detail)

        self.assertEqual(app._workspace_trash_restore_path_var.get(), "薪酬管理 / 工资表拆分")
        self.assertNotIn("20260810", app._workspace_trash_restore_path_var.get())
        self.assertIn("不会覆盖", app._workspace_trash_notice_var.get())

    def test_restore_uses_workspace_write_and_failure_keeps_dialog_selection(self) -> None:
        detail = _trash_detail()
        app = self._controller(detail)
        app._workspace_run_write = Mock()
        app._update_workspace_trash_restore_state = Mock()

        app._restore_selected_workspace_trash()

        self.assertTrue(app._workspace_trash_restore_in_progress)
        action, callback = app._workspace_run_write.call_args.args
        self.assertEqual(action, "恢复项目资料")
        callback(False)
        app.project_store.restore_from_trash.assert_called_once_with(detail.summary.id)

        failure = app._workspace_run_write.call_args.kwargs["on_error"]
        failure(RuntimeError("资料校验不一致"))
        self.assertFalse(app._workspace_trash_restore_in_progress)
        self.assertEqual(app._workspace_trash_selected_id, detail.summary.id)
        self.assertIn("资料校验不一致", app._workspace_trash_status_var.get())

    def test_restore_success_closes_dialog_and_reports_business_location(self) -> None:
        app = self._controller()
        app._workspace_run_write = Mock()
        app._update_workspace_trash_restore_state = Mock()
        app._close_workspace_trash_dialog = Mock()
        restored_root = app.current_project_path / "薪酬管理" / "工资表拆分" / "2026年7月 (2)"
        result = SimpleNamespace(directories={"uploads": restored_root / "上传资料"})

        app._restore_selected_workspace_trash()
        success = app._workspace_run_write.call_args.kwargs["on_success"]
        with patch("hr_toolkit.gui.messagebox.showinfo") as showinfo:
            success(result)

        app._close_workspace_trash_dialog.assert_called_once_with(force=True)
        self.assertIn("薪酬管理 / 工资表拆分", showinfo.call_args.args[1])
        self.assertNotIn("20260810", showinfo.call_args.args[1])
        self.assertIn("没有被覆盖", showinfo.call_args.args[1])

    def test_close_is_blocked_during_restore_unless_forced(self) -> None:
        app = self._controller()
        window = Mock()
        app._workspace_trash_window = window
        app._workspace_trash_restore_in_progress = True

        app._close_workspace_trash_dialog()
        window.destroy.assert_not_called()

        app._close_workspace_trash_dialog(force=True)
        window.destroy.assert_called_once_with()

    def test_batch_root_match_is_exact(self) -> None:
        with tempfile.TemporaryDirectory() as temp_root:
            batch_root = Path(temp_root) / "薪酬管理" / "工资表拆分" / "2026年7月"
            detail = SimpleNamespace(directories={"uploads": batch_root / "上传资料"})
            summary = SimpleNamespace(id="batch-1")
            app = HRToolkitApp.__new__(HRToolkitApp)
            app.project_store = SimpleNamespace(
                list_batches=lambda: (summary,),
                get_batch=lambda _batch_id: detail,
            )

            self.assertEqual(app._workspace_batch_root_for_path(batch_root), (summary, detail))
            self.assertIsNone(app._workspace_batch_root_for_path(batch_root / "上传资料"))
            self.assertIsNone(app._workspace_batch_root_for_path(batch_root / "上传资料" / "名单.xlsx"))

    def test_move_to_trash_only_moves_complete_batch_after_confirmation(self) -> None:
        detail = SimpleNamespace(directories={})
        summary = SimpleNamespace(
            id="batch-1",
            status="success",
            business_description="七月薪酬",
            business_period="2026-07",
            directory_name="2026年7月",
        )
        app = self._controller()
        app._selected_workspace_path = Mock(return_value=Path("/tmp/hr-project/完整批次"))
        app._workspace_batch_root_for_path = Mock(return_value=(summary, detail))
        app._workspace_run_write = Mock()

        with patch("hr_toolkit.gui.messagebox.askyesno", return_value=True):
            app._move_selected_workspace_batch_to_trash()

        action, callback = app._workspace_run_write.call_args.args
        self.assertEqual(action, "移到项目回收站")
        callback(False)
        app.project_store.move_to_trash.assert_called_once_with(summary.id)

    def test_recycle_entry_is_disabled_while_a_workspace_write_is_active(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.current_project_path = Path("/tmp/hr-project")
        app.project_store = object()
        app._workspace_project_read_only = False
        app._tool_running = False
        app._project_batch_by_token = {}
        app._workspace_write_in_progress = Mock(return_value=True)
        app._project_change_is_blocked = Mock(return_value=True)
        app._selected_workspace_path = Mock(return_value=None)
        app._update_workspace_trash_restore_state = Mock()
        app.workspace_add_button = Mock()
        app.workspace_refresh_button = Mock()
        app.workspace_open_project_button = Mock()
        app.workspace_trash_button = Mock()
        app.workspace_switch_button = Mock()
        app.workspace_search_entry = Mock()
        app.workspace_open_item_button = Mock()
        app.workspace_reveal_item_button = Mock()

        app._update_workspace_action_states()

        app.workspace_trash_button.configure.assert_called_once_with(state="disabled")


if __name__ == "__main__":
    unittest.main()
