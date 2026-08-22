from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest.mock import patch

from hr_toolkit import gui
from hr_toolkit.gui.app import HRToolkitApp


class FontScalingTests(unittest.TestCase):
    def test_macos_tk_86_keeps_legacy_font_compensation(self) -> None:
        with (
            patch.object(gui.sys, "platform", "darwin"),
            patch.object(gui, "TkVersion", 8.6),
        ):
            self.assertEqual(
                [gui._font_size(size) for size in (1, 7, 10, 18)],
                [1, 9, 13, 24],
            )

    def test_macos_tk_87_and_newer_use_native_point_size(self) -> None:
        for tk_version in (8.7, 9.0):
            with (
                self.subTest(tk_version=tk_version),
                patch.object(gui.sys, "platform", "darwin"),
                patch.object(gui, "TkVersion", tk_version),
            ):
                self.assertEqual(gui._font_size(10), 10)

    def test_non_macos_font_size_is_unchanged(self) -> None:
        for platform in ("win32", "linux"):
            for tk_version in (8.6, 9.0):
                with (
                    self.subTest(platform=platform, tk_version=tk_version),
                    patch.object(gui.sys, "platform", platform),
                    patch.object(gui, "TkVersion", tk_version),
                ):
                    self.assertEqual(gui._font_size(10), 10)


class ResponsiveLayoutTests(unittest.TestCase):
    def test_window_size_is_capped_at_work_area_even_at_100_percent(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.root = SimpleNamespace(
            winfo_screenwidth=lambda: 1366,
            winfo_screenheight=lambda: 768,
        )
        app.ui_scale = 1.0
        with patch("hr_toolkit.gui.app._windows_work_area_for_root", return_value=(0, 0, 1366, 728)):
            self.assertEqual(app._window_size(1400, 780), (1350, 712))
            self.assertEqual(app._window_size(900, 600), (900, 600))

    def test_window_size_respects_scaled_margin(self) -> None:
        app = HRToolkitApp.__new__(HRToolkitApp)
        app.root = SimpleNamespace(
            winfo_screenwidth=lambda: 1366,
            winfo_screenheight=lambda: 768,
        )
        app.ui_scale = 1.5
        with patch("hr_toolkit.gui.app._windows_work_area_for_root", return_value=(0, 0, 1366, 728)):
            self.assertEqual(app._window_size(1400, 780), (1342, 704))

    def test_supported_windows_size_and_scale_matrix_stays_inside_work_area(self) -> None:
        cases = (
            (1366, 728, 1.0, (1350, 712)),
            (1366, 728, 1.25, (1346, 708)),
            (1366, 728, 1.5, (1342, 704)),
            (1280, 680, 1.0, (1264, 664)),
            (1280, 680, 1.25, (1260, 660)),
            (1920, 1040, 1.5, (1896, 1016)),
        )
        for work_width, work_height, scale, expected in cases:
            with self.subTest(size=(work_width, work_height), scale=scale):
                app = HRToolkitApp.__new__(HRToolkitApp)
                app.root = SimpleNamespace(
                    winfo_screenwidth=lambda: work_width,
                    winfo_screenheight=lambda: work_height,
                )
                app.ui_scale = scale
                work_area = (0, 0, work_width, work_height)
                with patch("hr_toolkit.gui.app._windows_work_area_for_root", return_value=work_area):
                    self.assertEqual(app._window_size(1400, 780), expected)

    def test_form_layout_modes_follow_actual_usable_width(self) -> None:
        cases = (
            (900, gui.LAYOUT_MODE_WIDE),
            (760, gui.LAYOUT_MODE_WIDE),
            (759, gui.LAYOUT_MODE_COMPACT),
            (520, gui.LAYOUT_MODE_COMPACT),
            (519, gui.LAYOUT_MODE_NARROW),
            (320, gui.LAYOUT_MODE_NARROW),
        )
        for width, expected in cases:
            with self.subTest(width=width):
                self.assertEqual(gui._responsive_layout_mode(width), expected)

    def test_material_checkbox_columns_follow_layout_mode(self) -> None:
        self.assertEqual(gui._responsive_checkbox_columns(gui.LAYOUT_MODE_WIDE), 4)
        self.assertEqual(gui._responsive_checkbox_columns(gui.LAYOUT_MODE_COMPACT), 2)
        self.assertEqual(gui._responsive_checkbox_columns(gui.LAYOUT_MODE_NARROW), 1)

    def test_temporary_drawer_never_exceeds_available_width(self) -> None:
        self.assertEqual(gui._responsive_drawer_width(800, 320), 320)
        self.assertEqual(gui._responsive_drawer_width(300, 320), 288)
        self.assertEqual(gui._responsive_drawer_width(8, 320), 1)

if __name__ == "__main__":
    unittest.main()
