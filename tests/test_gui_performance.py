"""Unit tests for GUI rendering performance, layout stability, redraw caching, and scroll responsiveness."""

from __future__ import annotations

import tkinter as tk
import unittest

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
            for child in self.root.winfo_children():
                try:
                    child.destroy()
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
