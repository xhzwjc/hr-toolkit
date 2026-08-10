from __future__ import annotations

import unittest
from unittest.mock import patch

from hr_toolkit import gui


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


if __name__ == "__main__":
    unittest.main()
