from __future__ import annotations

import unittest
from unittest.mock import Mock, patch

from hr_toolkit import gui


def _layout_elements(layout) -> list[str]:
    elements: list[str] = []
    for element, options in layout:
        elements.append(element)
        elements.extend(_layout_elements(options.get("children", [])))
    return elements


class ScrollbarStyleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = gui.HRToolkitApp.__new__(gui.HRToolkitApp)
        self.app.root = Mock()
        self.app.ui_scale = 1.0

    def test_global_layout_uses_rounded_thumb_and_invisible_arrows(self) -> None:
        style = Mock()
        style.element_names.return_value = ()
        images = (Mock(), Mock(), Mock())
        transparent_arrow = Mock()

        with (
            patch.object(
                self.app,
                "_scrollbar_thumb_image",
                side_effect=images,
            ),
            patch.object(gui, "PhotoImage", return_value=transparent_arrow),
        ):
            self.app._configure_scrollbar_style(style)

        element_calls = {
            call.args[0]: call for call in style.element_create.call_args_list
        }
        element_call = element_calls["HRToolkit.Scrollbar.thumb"]
        self.assertEqual(element_call.args[:2], ("HRToolkit.Scrollbar.thumb", "image"))
        self.assertIs(element_call.args[2], images[0])
        self.assertEqual(element_call.args[3], ("pressed", images[2]))
        self.assertEqual(element_call.args[4], ("active", images[1]))
        self.assertEqual(element_call.kwargs["border"], 5)
        arrow_elements = {
            "HRToolkit.Vertical.Scrollbar.uparrow",
            "HRToolkit.Vertical.Scrollbar.downarrow",
            "HRToolkit.Horizontal.Scrollbar.leftarrow",
            "HRToolkit.Horizontal.Scrollbar.rightarrow",
        }
        self.assertEqual(set(element_calls), {"HRToolkit.Scrollbar.thumb", *arrow_elements})
        for element in arrow_elements:
            self.assertIs(element_calls[element].args[2], transparent_arrow)

        layouts = {call.args[0]: call.args[1] for call in style.layout.call_args_list}
        self.assertEqual(set(layouts), {"Vertical.TScrollbar", "Horizontal.TScrollbar"})
        for layout in layouts.values():
            elements = _layout_elements(layout)
            self.assertIn("HRToolkit.Scrollbar.thumb", elements)
            self.assertTrue(any(element in arrow_elements for element in elements))

        configured_styles = {call.args[0] for call in style.configure.call_args_list}
        self.assertEqual(configured_styles, {"Vertical.TScrollbar", "Horizontal.TScrollbar"})
        for call in style.configure.call_args_list:
            self.assertEqual(call.kwargs["arrowsize"], 12)
            self.assertEqual(call.kwargs["troughcolor"], gui.COLOR_SURFACE)

    def test_existing_elements_are_reused(self) -> None:
        style = Mock()
        style.element_names.return_value = (
            "HRToolkit.Scrollbar.thumb",
            "HRToolkit.Vertical.Scrollbar.uparrow",
            "HRToolkit.Vertical.Scrollbar.downarrow",
            "HRToolkit.Horizontal.Scrollbar.leftarrow",
            "HRToolkit.Horizontal.Scrollbar.rightarrow",
        )

        with (
            patch.object(self.app, "_scrollbar_thumb_image") as create_image,
            patch.object(gui, "PhotoImage") as create_transparent_image,
        ):
            self.app._configure_scrollbar_style(style)

        create_image.assert_not_called()
        create_transparent_image.assert_not_called()
        style.element_create.assert_not_called()
        self.assertEqual(style.layout.call_count, 2)


if __name__ == "__main__":
    unittest.main()
