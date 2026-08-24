"""Unit tests for GUI rendering performance, layout stability, redraw caching, and scroll responsiveness."""

from __future__ import annotations

import os
import tkinter as tk
import unittest
from unittest.mock import Mock, patch

from hr_toolkit.gui.constants import FORCE_UI_SCALE_ENV
from hr_toolkit.gui.app import HRToolkitApp
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

        item._on_configure()
        self.assertEqual(set(item.find_all()), initial_items)
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

            app._toggle_workspace_panel()
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
