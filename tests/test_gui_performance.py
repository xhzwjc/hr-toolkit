"""Unit tests for GUI rendering performance, layout stability, redraw caching, and scroll responsiveness."""

from __future__ import annotations

import tkinter as tk
import unittest

from hr_toolkit.gui.app import HRToolkitApp
from hr_toolkit.gui.widgets import CodexButton, RoundedCard, SidebarItem


class GuiPerformanceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.root = tk.Tk()
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls) -> None:
        cls.root.destroy()

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
        app = HRToolkitApp(self.root)
        self.root.update()

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
            self.root.update()
            self.assertEqual(app.current_tool, tool)

        # Switching to same tool is immediate no-op
        prev_tool = app.current_tool
        app._select_tool(prev_tool)
        self.root.update()
        self.assertEqual(app.current_tool, prev_tool)

    def test_right_canvas_scrolling_does_not_mutate_items(self) -> None:
        app = HRToolkitApp(self.root)
        self.root.update()

        canvas = app._right_canvas
        initial_items = len(canvas.find_all())

        # Perform multiple scrolling operations
        for delta in [1, -1, 2, -2, 1, -1]:
            canvas.yview_scroll(delta, "units")
            self.root.update_idletasks()

        self.assertEqual(len(canvas.find_all()), initial_items)


if __name__ == "__main__":
    unittest.main()
