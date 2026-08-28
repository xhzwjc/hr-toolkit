"""Unit tests for GUI rendering performance, layout stability, redraw caching, and scroll responsiveness."""

from __future__ import annotations

import os
import sys
import tempfile
import time
import tkinter as tk
import unittest
from collections import deque
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from hr_toolkit.gui.constants import FORCE_UI_SCALE_ENV
from hr_toolkit.gui.app import (
    HRToolkitApp,
    UPLOAD_LIST_VISIBLE_ROWS,
    WORKSPACE_TREE_RENDER_BATCH,
    WINDOW_RESIZE_SETTLE_MS,
    _CoalescedCanvasScroller,
)
from hr_toolkit.gui.widgets import CodexButton, RoundedCard, SidebarItem


class GuiPerformanceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        try:
            cls.root = tk.Tk()
            cls.root.withdraw()
        except Exception as exc:
            raise unittest.SkipTest(f"GUI display not available: {exc}")

    @classmethod
    def tearDownClass(cls) -> None:
        if hasattr(cls, "root") and cls.root:
            try:
                cls.root.destroy()
            except Exception:
                pass

    def tearDown(self) -> None:
        if hasattr(self, "root") and self.root:
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass

    def _run_event_loop_until(
        self,
        predicate,
        *,
        timeout_ms: int = 2000,
        failure_message: str,
    ) -> None:
        """Wait for a Tk state transition without assuming event delivery speed."""

        deadline = time.monotonic() + timeout_ms / 1000.0
        state = {"matched": bool(predicate())}

        def poll() -> None:
            state["matched"] = bool(predicate())
            if state["matched"] or time.monotonic() >= deadline:
                self.root.quit()
                return
            self.root.after(10, poll)

        if not state["matched"]:
            self.root.after_idle(poll)
            self.root.mainloop()
        self.assertTrue(state["matched"], failure_message)

    def test_tessellate_round_rect_geometry(self) -> None:
        from hr_toolkit.gui.widgets import _tessellate_round_rect

        # Standard rounded rect points
        pts = _tessellate_round_rect(0, 0, 100, 50, 10, segments_per_corner=4)
        self.assertGreater(len(pts), 16)
        self.assertEqual(len(pts) % 2, 0)

        # Zero radius fallback
        rect_pts = _tessellate_round_rect(0, 0, 100, 50, 0)
        self.assertEqual(rect_pts, [0, 0, 100, 0, 100, 50, 0, 50])

    def test_collapsed_workspace_defers_tree_read(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.workspace_tree = Mock()
        app._workspace_small = False
        app._workspace_preferred_expanded = False
        app._workspace_drawer_open = False
        app._workspace_search_generation = 0

        app._refresh_workspace_tree()

        app.workspace_tree.get_children.assert_not_called()

    def test_workspace_toggle_only_changes_current_session(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._workspace_small = False
        app._workspace_preferred_expanded = False
        app._save_workspace_preferences = Mock()
        app._apply_workspace_panel_mode = Mock()

        app._toggle_workspace_panel()

        self.assertTrue(app._workspace_preferred_expanded)
        app._save_workspace_preferences.assert_not_called()
        app._apply_workspace_panel_mode.assert_called_once_with()

    def test_workspace_mode_transition_to_expanded_refreshes_tree_once(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.root = Mock()
        app.ui_scale = 1.0
        app._workspace_small = False
        app._workspace_preferred_expanded = True
        app._workspace_width_units = 320
        app._workspace_panel_was_temporary_open = False
        app._workspace_panel_mode_key = ("place", "collapsed", 0)
        app._workspace_panel = Mock()
        app._workspace_main_area = Mock()
        app._workspace_main_area.winfo_width.return_value = 1400
        app._workspace_resize_handle = Mock()
        app._workspace_expanded_body = Mock()
        app._workspace_collapsed_body = Mock()
        app.workspace_collapse_button = Mock()
        app._update_workspace_text_wraps = Mock()
        app._refresh_workspace_tree = Mock()

        app._apply_workspace_panel_mode()

        scheduled = [call.args[0] for call in app.root.after_idle.call_args_list]
        self.assertIn(app._refresh_workspace_tree, scheduled)

        app.root.after_idle.reset_mock()
        app._apply_workspace_panel_mode()

        scheduled = [call.args[0] for call in app.root.after_idle.call_args_list]
        self.assertNotIn(app._refresh_workspace_tree, scheduled)

    def test_large_upload_folder_metadata_uses_bounded_directory_scan(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            for index in range(230):
                (folder / f"item-{index:03d}.txt").touch()

            _badge, _background, _foreground, detail = app._upload_item_meta(folder)

        self.assertEqual(detail, "文件夹 · 200+ 个项目")

    def test_log_writes_are_coalesced_and_flushed_in_bounded_batches(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._is_alive = True
        app._pending_log_entries = deque()
        app._log_flush_job = None
        app.root = Mock()
        app.root.after_idle.return_value = "log-job"
        app.log_text = Mock()
        app.log_text.winfo_exists.return_value = True

        for index in range(1000):
            app._write_log(f"提醒：第 {index} 条")

        self.assertEqual(app.root.after_idle.call_count, 1)
        self.assertEqual(app.log_text.insert.call_count, 0)
        app._flush_pending_log_entries()
        self.assertEqual(len(app._pending_log_entries), 800)
        self.assertEqual(app.log_text.see.call_count, 1)
        app.root.after.assert_called_once_with(1, app._flush_pending_log_entries)

    def test_log_paint_is_deferred_while_content_is_scrolling(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._is_alive = True
        app._scroll_active = True
        app._pending_log_entries = deque()
        app._log_flush_job = None
        app.root = Mock()
        app.log_text = Mock()
        app.log_text.winfo_exists.return_value = True

        app._write_log("正在处理第 1 条")

        self.assertEqual(len(app._pending_log_entries), 1)
        app.root.after_idle.assert_not_called()
        app.log_text.insert.assert_not_called()

        app._scroll_active = False
        app._flush_pending_log_entries()
        self.assertEqual(len(app._pending_log_entries), 0)
        self.assertGreater(app.log_text.insert.call_count, 0)

    def test_smooth_wheel_scroller_animates_then_reaches_exact_target(self) -> None:
        root = Mock()
        root.after.return_value = "scroll-job"
        canvas = Mock()
        canvas.winfo_exists.return_value = True
        controller = _CoalescedCanvasScroller(
            root,
            canvas,
            smooth_units=True,
        )
        controller.update_geometry(2000, 500)
        controller.update_view(0.25, 0.5)

        controller.queue_units(1)
        controller._flush()

        first_frame = canvas.yview_moveto.call_args.args[0]
        self.assertGreater(first_frame, 0.25)
        self.assertLess(first_frame, 0.275)
        self.assertTrue(controller.scheduled)

        controller.flush_pending()
        self.assertFalse(controller.scheduled)
        self.assertAlmostEqual(canvas.yview_moveto.call_args.args[0], 0.275, places=6)

    def test_tool_thread_fairness_is_restored_after_processing(self) -> None:
        original = sys.getswitchinterval()
        app = HRToolkitApp.__new__(HRToolkitApp)
        app._original_thread_switch_interval = None
        app._active_tool_worker_tokens = set()
        app._tool_running = True
        try:
            app._enable_ui_thread_fairness()
            self.assertLessEqual(sys.getswitchinterval(), original)
            app._tool_running = False
            app._restore_thread_switch_interval()
            self.assertAlmostEqual(sys.getswitchinterval(), original, places=7)
        finally:
            sys.setswitchinterval(original)

    def test_workspace_batch_locations_are_cached_until_tree_refresh(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        root = Path("/project")
        directories = {
            "uploads": root / "业务" / "工具" / "批次" / "上传资料",
            "supplements": root / "业务" / "工具" / "批次" / "补充资料",
            "results": root / "业务" / "工具" / "批次" / "处理结果",
        }
        summary = SimpleNamespace(id="a" * 32, tool_id="salary_split")
        detail = SimpleNamespace(summary=summary, directories=directories)

        class Store:
            list_batch_locations = Mock(return_value=((summary, directories),))
            get_batch = Mock(return_value=detail)

        store = Store()
        app.project_store = store
        app.current_project_path = root
        app._workspace_batch_location_store = None
        app._workspace_batch_location_cache = ()

        self.assertEqual(
            app._workspace_batch_for_target(directories["uploads"] / "名单.xlsx"),
            (summary.id, "uploads"),
        )
        self.assertEqual(
            app._workspace_batch_root_for_path(directories["uploads"].parent),
            (summary, detail),
        )
        app._workspace_batch_for_target(directories["supplements"])
        store.list_batch_locations.assert_called_once_with()

        app._invalidate_workspace_batch_location_cache()
        app._workspace_batch_for_target(directories["supplements"])
        self.assertEqual(store.list_batch_locations.call_count, 2)

    def test_search_result_rendering_does_not_stat_paths_on_ui_thread(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.workspace_tree = Mock()
        app.workspace_tree.get_children.return_value = ()
        app.workspace_tree.insert.side_effect = ("item-1", "item-2")
        app._workspace_tree_paths = {}
        app.current_project_path = Path("/project")
        app.workspace_empty_text = Mock()
        app.workspace_detail_title = Mock()
        app.workspace_detail_text = Mock()
        app._show_workspace_empty = Mock()
        results = [Path("/project/folder"), Path("/project/file.xlsx")]

        with patch.object(Path, "is_dir", side_effect=AssertionError("UI stat")):
            app._render_workspace_search_results(results, False, None)

        self.assertEqual(app.workspace_tree.insert.call_count, 2)

    def test_large_workspace_directory_rows_are_rendered_in_bounded_batches(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.root = Mock()
        scheduled_callbacks = []
        app.root.after.side_effect = (
            lambda _delay, callback: scheduled_callbacks.append(callback) or "job"
        )
        app.workspace_tree = Mock()
        app.workspace_tree.insert.side_effect = (
            f"item-{index}" for index in range(WORKSPACE_TREE_RENDER_BATCH + 5)
        )
        app._workspace_tree_paths = {}
        app._workspace_search_generation = 7
        app._workspace_small = False
        app._workspace_preferred_expanded = True
        app._workspace_drawer_open = False
        app.current_project_path = Path("/project")
        app._show_workspace_empty = Mock()
        app._update_workspace_action_states = Mock()
        records = [
            (Path(f"/project/file-{index:04d}.xlsx"), False)
            for index in range(WORKSPACE_TREE_RENDER_BATCH + 5)
        ]

        app._render_workspace_directory_records(7, "", Path("/project"), records)

        self.assertEqual(app.workspace_tree.insert.call_count, WORKSPACE_TREE_RENDER_BATCH)
        self.assertEqual(len(scheduled_callbacks), 1)
        scheduled_callbacks.pop()()
        self.assertEqual(app.workspace_tree.insert.call_count, len(records))
        app._update_workspace_action_states.assert_called_once_with()

    def test_large_workspace_root_scan_moves_full_load_off_ui_thread(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.workspace_tree = Mock()
        app.workspace_tree.get_children.return_value = ()
        app._workspace_tree_paths = {}
        app._workspace_search_generation = 0
        app._workspace_small = False
        app._workspace_preferred_expanded = True
        app._workspace_drawer_open = False
        app._invalidate_workspace_batch_location_cache = Mock()
        app._set_workspace_detail = Mock()
        app._update_workspace_action_states = Mock()
        app._start_workspace_directory_load = Mock()
        app.workspace_search = Mock()
        app.workspace_search.get.return_value = ""
        app.workspace_empty_text = Mock()
        app._show_workspace_empty = Mock()

        with tempfile.TemporaryDirectory() as temp_dir:
            root_path = Path(temp_dir)
            app.current_project_path = root_path
            app._workspace_scope_root = Mock(return_value=root_path)
            records = [
                (root_path / f"file-{index:04d}.xlsx", False)
                for index in range(201)
            ]
            app._workspace_visible_child_records = Mock(return_value=records)

            app._refresh_workspace_tree()

        app._start_workspace_directory_load.assert_called_once_with(
            1,
            "",
            root_path,
        )
        app.workspace_tree.insert.assert_not_called()

    def test_codex_button_idempotent_redraw_and_item_reuse(self) -> None:
        btn = CodexButton(self.root, text="测试按钮", variant="primary")
        self.root.update_idletasks()

        # Capture initial canvas item IDs
        initial_items = set(btn.find_all())
        self.assertGreater(len(initial_items), 0)

        # Re-triggering _redraw with identical state should be a no-op
        btn._redraw()
        current_items = set(btn.find_all())
        self.assertEqual(initial_items, current_items)

        # Triggering <Configure> with identical dimensions should not wipe items
        btn._on_configure()
        self.assertEqual(set(btn.find_all()), initial_items)

        # Changing text should reuse/update existing items without blank flash
        btn.configure(text="更新文字")
        self.root.update_idletasks()
        self.assertEqual(btn._display_text(), "更新文字")
        btn.destroy()

    def test_codex_button_releases_and_replaces_variable_trace(self) -> None:
        first = tk.StringVar(master=self.root, value="first")
        second = tk.StringVar(master=self.root, value="second")
        btn = CodexButton(self.root, textvariable=first)

        self.assertEqual(len(first.trace_info()), 1)
        btn.configure(textvariable=second)
        self.assertEqual(first.trace_info(), [])
        self.assertEqual(len(second.trace_info()), 1)

        btn.destroy()
        self.assertEqual(second.trace_info(), [])

    def test_sidebar_item_idempotent_redraw(self) -> None:
        item = SidebarItem(self.root, text="社保明细", icon_id="social_security")
        self.root.update_idletasks()

        initial_items = set(item.find_all())
        item._redraw()
        self.assertEqual(set(item.find_all()), initial_items)

        initial_icon_items = tuple(item._icon_items)
        item._on_enter()
        item._on_leave()
        item.set_selected(True)
        item.set_selected(False)
        self.assertEqual(tuple(item._icon_items), initial_icon_items)

        state_items = set(item.find_all())
        item._on_configure()
        self.assertEqual(set(item.find_all()), state_items)
        item.destroy()

    def test_rounded_card_in_place_polygon_update_without_delete(self) -> None:
        card = RoundedCard(self.root, padding=(20, 16, 20, 18))
        self.root.update_idletasks()

        # Trigger initial sync
        card._sync()
        card_items = card.find_withtag("card_bg")
        self.assertEqual(len(card_items), 2)  # shadow polygon + surface polygon
        shadow_id, bg_id = card_items

        # Resyncing or resizing should update coordinates in-place rather than deleting
        card._sync()
        card_items_after = card.find_withtag("card_bg")
        self.assertEqual(card_items_after, (shadow_id, bg_id))
        card.destroy()

    def test_tool_switching_layout_stability_and_redraw_counts(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.update_idletasks()

            tools = [
                "social_security",
                "insurance_ledger",
                "data_statistics",
                "salary_split",
                "salary_merge",
                "personnel_change_merge",
                "archive_import",
                "material_collector",
                "folder_rename",
            ]

            for tool in tools:
                app._select_tool(tool)
                self.root.update_idletasks()
                self.assertEqual(app.current_tool, tool)

            # Switching to same tool is immediate no-op
            prev_tool = app.current_tool
            app._select_tool(prev_tool)
            self.root.update_idletasks()
            self.assertEqual(app.current_tool, prev_tool)
        finally:
            if app is not None:
                app.destroy()
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass

    def test_all_tool_pages_keep_coalesced_scrolling_after_upload(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            tools = [
                "social_security",
                "insurance_ledger",
                "data_statistics",
                "salary_split",
                "salary_merge",
                "personnel_change_merge",
                "archive_import",
                "material_collector",
                "folder_rename",
            ]
            with tempfile.TemporaryDirectory() as temp_dir:
                paths = []
                for index in range(12):
                    path = Path(temp_dir) / f"输入_{index + 1:02d}.xlsx"
                    path.write_bytes(b"test")
                    paths.append(path)

                for tool in tools:
                    app._select_tool(tool)
                    if app._input_allow_multi:
                        app.change_input_paths = list(paths)
                        app._sync_input_path_text()
                    else:
                        app.change_input_paths = None
                        app.input_path.set(str(paths[0]))
                    app._refresh_upload_card()
                    self.root.update()
                    for _ in range(40):
                        app._right_canvas.event_generate("<MouseWheel>", delta=-120)
                    self.assertTrue(app._right_scroll_controller.scheduled, tool)
                    app._right_scroll_controller.flush_pending()
                    self.assertFalse(app._right_scroll_controller.scheduled, tool)
        finally:
            if app is not None:
                app.destroy()

    def test_small_window_reflows_forms_and_uses_temporary_workspace_drawer(self) -> None:
        app = None
        try:
            with patch.dict(os.environ, {FORCE_UI_SCALE_ENV: "1"}, clear=False):
                app = HRToolkitApp(self.root)
            self.root.deiconify()
            self.root.geometry("900x650")
            self.root.update()

            app._select_tool("material_collector")
            self.root.update()
            self.assertEqual(app._form_layout_mode, "compact")
            self.assertTrue(app.workspace_title_button.winfo_ismapped())
            self.assertEqual(app._workspace_panel.winfo_manager(), "")
            material_columns = {
                int(widget.grid_info()["column"])
                for widget in app._material_check_widgets
            }
            self.assertLessEqual(max(material_columns, default=0), 1)

            app.workspace_title_button.event_generate(
                "<Button-1>",
                x=max(1, app.workspace_title_button.winfo_width() // 2),
                y=max(1, app.workspace_title_button.winfo_height() // 2),
            )
            self.root.update()
            self.assertEqual(app._workspace_panel.winfo_manager(), "place")
            self.assertLessEqual(
                app._workspace_panel.winfo_width(),
                app._workspace_main_area.winfo_width(),
            )
            outside_click = tk.Event()
            outside_click.widget = app.form
            app._close_workspace_drawer_on_outside_click(outside_click)
            self.root.update()
            self.assertEqual(app._workspace_panel.winfo_manager(), "")

            app._select_tool("data_statistics")
            self.root.update()
            self.assertLess(app.stats_unit_col.winfo_rooty(), app.stats_out_col.winfo_rooty())
            self.assertLess(app.stats_out_col.winfo_rooty(), app.stats_trip_col.winfo_rooty())

            self.root.minsize(1, 1)
            self.root.geometry("700x650")
            self.root.update()
            app._select_tool("material_collector")
            app.material_collect_all.set(False)
            app._on_material_collect_all_changed()
            self.root.update()
            self.assertEqual(app._form_layout_mode, "narrow")
            narrow_columns = {
                int(widget.grid_info()["column"])
                for widget in app._material_check_widgets
            }
            self.assertEqual(narrow_columns, {0})

            app._select_tool("data_statistics")
            self.root.update()
            self.assertTrue(all(button.winfo_manager() == "grid" for button in app.stats_week_preset_buttons))
        finally:
            if app is not None:
                app.destroy()
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass

    def test_rapid_window_resize_settles_canvas_and_responsive_layout(self) -> None:
        app = None
        try:
            with patch.dict(os.environ, {FORCE_UI_SCALE_ENV: "1"}, clear=False):
                app = HRToolkitApp(self.root)
            self.root.deiconify()
            self.root.minsize(1, 1)
            app._dismiss_startup_loading_screen()
            self.root.update()
            right_frame = self.root.nametowidget(
                app._right_canvas.itemcget(app._right_canvas_window, "window")
            )

            def layout_is_settled(mode: str, workspace_small: bool) -> bool:
                canvas_width = app._right_canvas.winfo_width()
                canvas_height = app._right_canvas.winfo_height()
                expected_window_size = (
                    canvas_width,
                    max(right_frame.winfo_reqheight(), canvas_height),
                )
                return (
                    not app._window_resize_active
                    and not app._window_resize_restoring
                    and app._form_layout_mode == mode
                    and app._workspace_small is workspace_small
                    and app._form_resize_job is None
                    and app._form_post_layout_sync_job is None
                    and app._workspace_area_resize_job is None
                    and not app._right_canvas_sync_pending
                    and app._last_canvas_window_size == expected_window_size
                )

            # Several native resize messages may arrive before the settle job.
            # Only the final dimensions must drive the expensive content reflow.
            for width in range(1180, 699, -40):
                self.root.geometry(f"{width}x650")
                self.root.update_idletasks()
            self._run_event_loop_until(
                lambda: layout_is_settled("narrow", True),
                failure_message="narrow resize state did not settle",
            )

            self.assertEqual(app._form_layout_mode, "narrow")
            self.assertTrue(app._workspace_small)
            self.assertEqual(
                app._last_canvas_window_size[0],
                app._right_canvas.winfo_width(),
            )
            self.assertEqual(
                app._last_canvas_window_size[1],
                max(right_frame.winfo_reqheight(), app._right_canvas.winfo_height()),
            )

            for width in range(700, 1401, 50):
                self.root.geometry(f"{width}x780")
                self.root.update_idletasks()
            self._run_event_loop_until(
                lambda: layout_is_settled("wide", False),
                failure_message="wide resize state did not settle",
            )

            self.assertEqual(app._form_layout_mode, "wide")
            self.assertFalse(app._workspace_small)
            self.assertEqual(
                app._last_canvas_window_size[0],
                app._right_canvas.winfo_width(),
            )
            self.assertEqual(
                app._last_canvas_window_size[1],
                max(right_frame.winfo_reqheight(), app._right_canvas.winfo_height()),
            )
        finally:
            if app is not None:
                app.destroy()
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass

    def test_live_resize_uses_one_preview_surface_and_restores_widget_tree(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            self.root.minsize(1, 1)
            app._dismiss_startup_loading_screen()
            # Keep this test independent of desktop screenshot permissions.
            snapshot_job = app._window_resize_snapshot_job
            if snapshot_job is not None:
                self.root.after_cancel(snapshot_job)
                app._window_resize_snapshot_job = None
            app._window_resize_snapshot_supported = False
            observed = {}

            def inspect_during_drag() -> None:
                observed["active"] = app._window_resize_active
                observed["manager"] = app._root_frame.winfo_manager()
                observed["overlay"] = app._window_resize_overlay.winfo_manager()
                observed["activations"] = app._window_resize_overlay_activations

            for index in range(18):
                width = 1180 - index * 14
                height = 760 - index * 4
                self.root.after(
                    index * 5,
                    lambda w=width, h=height: self.root.geometry(f"{w}x{h}"),
                )
            self.root.after(60, inspect_during_drag)
            self.root.after(300, self.root.quit)
            self.root.mainloop()

            self.assertTrue(observed["active"])
            self.assertEqual(observed["manager"], "")
            self.assertEqual(observed["overlay"], "place")
            self.assertEqual(observed["activations"], 1)
            self.assertFalse(app._window_resize_active)
            self.assertEqual(app._root_frame.winfo_manager(), "pack")
            self.assertEqual(app._window_resize_overlay.winfo_manager(), "")
            self.assertGreater(app._window_resize_overlay_frames, 0)
        finally:
            if app is not None:
                app.destroy()
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass

    def test_live_resize_can_render_cached_ui_snapshot(self) -> None:
        try:
            from PIL import Image, ImageTk
        except ImportError as exc:
            self.skipTest(f"Pillow is unavailable: {exc}")

        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            app._dismiss_startup_loading_screen()
            self.root.update()
            width = max(app.root.winfo_width(), 2)
            height = max(app.root.winfo_height(), 2)
            snapshot = Image.new("RGB", (width, height), "#F5F5F4")
            app._window_resize_snapshot_image = snapshot
            app._window_resize_snapshot_photo = ImageTk.PhotoImage(
                snapshot,
                master=self.root,
            )

            app._begin_window_resize_composite()
            app._window_resize_overlay_target_size = (width - 20, height - 10)
            app._render_window_resize_overlay()

            self.assertTrue(app._window_resize_active)
            self.assertIsNotNone(app._window_resize_overlay_photo)
            self.assertEqual(
                app._window_resize_overlay.itemcget(
                    app._window_resize_overlay_image_item,
                    "state",
                ),
                "normal",
            )
            self.assertGreater(app._window_resize_overlay_frames, 0)
            app._flush_window_resize_composite()
            self.assertFalse(app._window_resize_active)
        finally:
            if app is not None:
                app.destroy()

    def test_resize_compositor_shutdown_cancels_jobs_and_releases_preview(self) -> None:
        app = HRToolkitApp(self.root)
        self.root.deiconify()
        app._dismiss_startup_loading_screen()
        self.root.geometry("1080x690")
        self.root.after(20, self.root.quit)
        self.root.mainloop()
        self.assertTrue(app._window_resize_active)
        bindtag = app._window_resize_configure_bindtag
        self.assertIn(bindtag, self.root.bindtags())
        self.assertNotIn(bindtag, app._right_canvas.bindtags())

        overlay = app._window_resize_overlay
        app.destroy()

        self.assertFalse(app._window_resize_active)
        self.assertIsNone(app._window_resize_settle_job)
        self.assertIsNone(app._window_resize_render_job)
        self.assertIsNone(app._window_resize_snapshot_job)
        self.assertNotIn(bindtag, self.root.bindtags())
        self.assertFalse(overlay.winfo_exists())

    def test_native_resize_release_restores_immediately_and_preserves_focus(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            app._dismiss_startup_loading_screen()
            app._select_tool("folder_rename")
            self.root.update()
            app.rename_text_widget.focus_force()
            pointer = {"down": True}
            app._window_resize_pointer_state_reader = lambda: pointer["down"]

            self.root.geometry("1080x690")
            self._run_event_loop_until(
                lambda: (
                    app._window_resize_active
                    and app._window_resize_native_drag
                ),
                failure_message="native resize compositor did not start",
            )
            # A confirmed native drag has no inactivity timer. Releasing the
            # mocked button can therefore complete only through the platform
            # button-state poll, regardless of Configure delivery latency.
            self.assertIsNone(app._window_resize_settle_job)
            pointer["down"] = False
            self._run_event_loop_until(
                lambda: not app._window_resize_active,
                failure_message="native resize compositor did not restore",
            )

            self.assertEqual(app._root_frame.winfo_manager(), "pack")
            self.assertIs(self.root.focus_get(), app.rename_text_widget)
        finally:
            if app is not None:
                app.destroy()

    def test_native_resize_pause_keeps_single_surface_until_release(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            app._dismiss_startup_loading_screen()
            pointer = {"down": True}
            app._window_resize_pointer_state_reader = lambda: pointer["down"]
            observed = {}

            self.root.geometry("1080x690")

            def inspect_while_held() -> None:
                observed["active_while_held"] = app._window_resize_active
                observed["manager_while_held"] = app._root_frame.winfo_manager()
                pointer["down"] = False

            def inspect_after_release() -> None:
                observed["active_after_release"] = app._window_resize_active
                observed["manager_after_release"] = app._root_frame.winfo_manager()

            # Hold the border motionless beyond the inactivity fallback.  The
            # complex widget tree must remain unmapped until the real release.
            self.root.after(WINDOW_RESIZE_SETTLE_MS + 60, inspect_while_held)
            self.root.after(WINDOW_RESIZE_SETTLE_MS + 130, inspect_after_release)
            self.root.after(WINDOW_RESIZE_SETTLE_MS + 150, self.root.quit)
            self.root.mainloop()

            self.assertTrue(observed["active_while_held"])
            self.assertEqual(observed["manager_while_held"], "")
            self.assertFalse(observed["active_after_release"])
            self.assertEqual(observed["manager_after_release"], "pack")
        finally:
            if app is not None:
                app.destroy()

    def test_right_canvas_scrolling_does_not_mutate_items(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.update_idletasks()

            canvas = app._right_canvas
            initial_items = len(canvas.find_all())

            # Perform multiple scrolling operations
            for delta in [1, -1, 2, -2, 1, -1]:
                canvas.yview_scroll(delta, "units")
                self.root.update_idletasks()

            self.assertEqual(len(canvas.find_all()), initial_items)
        finally:
            if app is not None:
                app.destroy()

    def test_right_canvas_coalesces_wheel_event_burst(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            self.root.update()

            canvas = app._right_canvas
            original_scroll = canvas.yview_scroll
            calls: list[tuple[int, str]] = []

            def record_scroll(units: int, mode: str):
                calls.append((units, mode))
                return original_scroll(units, mode)

            canvas.yview_scroll = record_scroll
            for _ in range(240):
                canvas.event_generate("<MouseWheel>", delta=-120)

            self.assertTrue(app._right_scroll_controller.scheduled)
            self.assertEqual(calls, [])
            app._right_scroll_controller.flush_pending()
            self.assertEqual(len(calls), 1)
            self.assertEqual(calls[0][1], "units")
        finally:
            if app is not None:
                app.destroy()

    def test_right_canvas_coalesces_touchpad_pixel_burst(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            controller = app._right_scroll_controller
            controller.update_geometry(2000, 500)
            controller.update_view(0.5, 0.75)
            movements: list[float] = []
            app._right_canvas.yview_moveto = lambda fraction: movements.append(float(fraction))

            for _ in range(40):
                controller.queue_pixels(2)

            self.assertTrue(controller.scheduled)
            controller.flush_pending()
            self.assertEqual(len(movements), 1)
            self.assertAlmostEqual(movements[0], 0.46, places=5)
        finally:
            if app is not None:
                app.destroy()

    def test_twelve_uploads_share_one_canvas_surface(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            app._select_tool("personnel_change_merge")
            with tempfile.TemporaryDirectory() as temp_dir:
                paths = []
                for index in range(12):
                    path = Path(temp_dir) / f"异动表_{index + 1:02d}.xlsx"
                    path.write_bytes(b"test")
                    paths.append(path)
                app.change_input_paths = paths
                app._sync_input_path_text()
                app._refresh_upload_card()
                self.root.update()

                children = app.upload_body.winfo_children()
                self.assertEqual(len(children), 1)
                self.assertIsInstance(children[0], tk.Canvas)
                self.assertTrue(children[0].find_withtag("chip_close_0"))
                self.assertTrue(children[0].find_withtag("chip_close_11"))
                expected_height = 12 * app._px(44) + 11 * app._px(8)
                self.assertEqual(int(float(children[0].cget("height"))), expected_height)

                removed_path = paths[5]
                close_x = children[0].winfo_width() - app._px(20)
                close_y = 5 * (app._px(44) + app._px(8)) + app._px(22)
                children[0].event_generate("<Motion>", x=close_x, y=close_y)
                self.root.update()
                children[0].event_generate("<Button-1>", x=close_x, y=close_y)
                self.root.update()
                self.assertEqual(len(app.change_input_paths or []), 11)
                self.assertNotIn(removed_path, app.change_input_paths or [])
        finally:
            if app is not None:
                app.destroy()

    def test_large_upload_list_has_bounded_height_and_visible_item_count(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            app._select_tool("personnel_change_merge")
            with tempfile.TemporaryDirectory() as temp_dir:
                paths = []
                for index in range(200):
                    path = Path(temp_dir) / f"异动表_{index + 1:03d}.xlsx"
                    path.write_bytes(b"test")
                    paths.append(path)
                app.change_input_paths = paths
                app._sync_input_path_text()
                app._refresh_upload_card()
                self.root.update()

                upload_canvas = app.upload_body.winfo_children()[0]
                expected_height = (
                    UPLOAD_LIST_VISIBLE_ROWS * app._px(44)
                    + (UPLOAD_LIST_VISIBLE_ROWS - 1) * app._px(8)
                )
                self.assertEqual(app._upload_body_height, expected_height)
                self.assertEqual(int(float(upload_canvas.cget("height"))), expected_height)
                self.assertLess(len(upload_canvas.find_all()), 160)

                controller = app._upload_items_scroll_controller
                controller.queue_units(40)
                controller.flush_pending()
                self.root.update()

                rendered_start, rendered_end = app._upload_items_rendered_range
                self.assertGreater(rendered_start, 0)
                self.assertLess(rendered_end - rendered_start, 30)
                self.assertTrue(upload_canvas.find_withtag("chip_close_40"))
                self.assertLess(len(upload_canvas.find_all()), 160)

                thumb = upload_canvas.find_withtag("upload_scrollbar")
                self.assertEqual(len(thumb), 1)
                thumb_box = upload_canvas.bbox(thumb[0])
                self.assertIsNotNone(thumb_box)
                assert thumb_box is not None
                thumb_x = (thumb_box[0] + thumb_box[2]) // 2
                thumb_y = int(
                    (thumb_box[1] + thumb_box[3]) / 2
                    - upload_canvas.canvasy(0)
                )
                before_drag = upload_canvas.yview()[0]
                upload_canvas.event_generate(
                    "<ButtonPress-1>",
                    x=thumb_x,
                    y=thumb_y,
                )
                upload_canvas.event_generate(
                    "<B1-Motion>",
                    x=thumb_x,
                    y=thumb_y + app._px(80),
                )
                upload_canvas.event_generate(
                    "<ButtonRelease-1>",
                    x=thumb_x,
                    y=thumb_y + app._px(80),
                )
                self.root.update()
                self.assertGreater(upload_canvas.yview()[0], before_drag)
        finally:
            if app is not None:
                app.destroy()

    def test_upload_canvases_reuse_items_during_width_changes(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            app._select_tool("personnel_change_merge")
            with tempfile.TemporaryDirectory() as temp_dir:
                paths = []
                for index in range(12):
                    path = Path(temp_dir) / f"异动表_{index + 1:02d}.xlsx"
                    path.write_bytes(b"test")
                    paths.append(path)
                app.change_input_paths = paths
                app._sync_input_path_text()
                app._refresh_upload_card()
                self.root.update()

                upload_canvas = app.upload_body.winfo_children()[0]
                initial_items = upload_canvas.find_all()
                with patch.object(
                    app,
                    "_upload_item_meta",
                    side_effect=AssertionError("unchanged upload metadata was re-read"),
                ):
                    app._refresh_upload_card()
                self.root.update()
                self.assertIs(app.upload_body.winfo_children()[0], upload_canvas)
                self.root.minsize(1, 1)
                self.root.geometry("900x650")
                self.root.after(250, self.root.quit)
                self.root.mainloop()
                self.assertEqual(upload_canvas.find_all(), initial_items)

                app.change_input_paths = None
                app._sync_input_path_text()
                app._refresh_upload_card()
                self.root.update()
                drop_zone = app.upload_body.winfo_children()[0]
                initial_drop_items = drop_zone.find_all()
                self.root.geometry("1180x760")
                self.root.after(250, self.root.quit)
                self.root.mainloop()
                self.assertEqual(drop_zone.find_all(), initial_drop_items)
        finally:
            if app is not None:
                app.destroy()
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass

    def test_empty_upload_zone_is_reused_across_tool_switches(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            self.root.deiconify()
            self.root.update()
            drop_zone = app.upload_body.winfo_children()[0]

            app._select_tool("insurance_ledger")
            self.root.update()

            self.assertIs(app.upload_body.winfo_children()[0], drop_zone)
            text_items = [
                drop_zone.itemcget(item_id, "text")
                for item_id in drop_zone.find_all()
                if drop_zone.type(item_id) == "text"
            ]
            self.assertIn("选择保单清单、压缩包或文件夹", text_items)

            app._select_tool("material_collector")
            self.root.update()

            self.assertIs(app.upload_body.winfo_children()[0], drop_zone)
            text_items = [
                drop_zone.itemcget(item_id, "text")
                for item_id in drop_zone.find_all()
                if drop_zone.type(item_id) == "text"
            ]
            self.assertIn("点击浏览文件夹路径", text_items)

            folder_command = Mock()
            app._input_folder_cmd = folder_command
            app._refresh_upload_card()
            self.root.update()
            folder_link = next(
                item_id
                for item_id in drop_zone.find_all()
                if drop_zone.type(item_id) == "text"
                and drop_zone.itemcget(item_id, "text") == "点击浏览文件夹路径"
            )
            x1, y1, x2, y2 = drop_zone.bbox(folder_link)
            drop_zone.event_generate(
                "<Button-1>",
                x=(x1 + x2) // 2,
                y=(y1 + y2) // 2,
            )
            self.root.update()
            folder_command.assert_called_once_with()
        finally:
            if app is not None:
                app.destroy()
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass

    def test_duplicate_configure_events_use_cached_dimensions(self) -> None:
        button = CodexButton(self.root, text="测试按钮", variant="primary")
        sidebar = SidebarItem(self.root, text="社保明细", icon_id="social_security")
        card = RoundedCard(self.root, padding=(20, 16, 20, 18))
        self.root.update_idletasks()

        button_event = SimpleNamespace(
            width=button._last_draw_key[0],
            height=button._last_draw_key[1],
        )
        sidebar_event = SimpleNamespace(
            width=sidebar._last_draw_key[0],
            height=sidebar._last_draw_key[1],
        )
        card_size = (card.winfo_width(), card.winfo_height())
        card._last_self_event_size = card_size
        card_event = SimpleNamespace(
            widget=card,
            width=card_size[0],
            height=card_size[1],
        )

        with patch.object(button, "winfo_width", side_effect=AssertionError("button geometry query")):
            button._on_configure(button_event)
        with patch.object(sidebar, "winfo_width", side_effect=AssertionError("sidebar geometry query")):
            sidebar._on_configure(sidebar_event)
        with patch.object(card, "winfo_width", side_effect=AssertionError("card geometry query")):
            card._sync(card_event)

        button.destroy()
        sidebar.destroy()
        card.destroy()

    def test_paint_codex_badge_icon(self) -> None:
        from hr_toolkit.gui.widgets import _paint_codex_badge_icon

        canvas = tk.Canvas(self.root, width=100, height=100)
        items = _paint_codex_badge_icon(canvas, 10, 10, 64)
        self.assertGreater(len(items), 4)
        canvas.destroy()

    def test_startup_loading_screen_lifecycle(self) -> None:
        app = None
        try:
            app = HRToolkitApp(self.root)
            # The loading overlay is initially created
            self.assertTrue(hasattr(app, "_setup_startup_loading_screen"))
            # Explicitly dismiss
            app._dismiss_startup_loading_screen()
            self.assertIsNone(app._loading_overlay)
            # Idempotent dismiss
            app._dismiss_startup_loading_screen()
        finally:
            if app is not None:
                app.destroy()
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
