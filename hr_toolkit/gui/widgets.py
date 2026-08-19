"""Custom Canvas-rendered GUI widgets with anti-aliased rounded corners and macOS/Notion styling."""

from __future__ import annotations

from tkinter import Canvas, StringVar, ttk

from .constants import (
    COLOR_BG,
    COLOR_BORDER,
    COLOR_DISABLED,
    COLOR_NAV_HOVER,
    COLOR_NAV_SELECTED,
    COLOR_NAV_TEXT,
    COLOR_NAV_TEXT_SELECTED,
    COLOR_MUTED,
    COLOR_PRIMARY,
    COLOR_PRIMARY_ACTIVE,
    COLOR_SIDEBAR,
    COLOR_SURFACE,
    COLOR_SURFACE_ALT,
    COLOR_SURFACE_PRESSED,
)
from .scaling import _font_size, _scale_float, _scale_px, _widget_ui_scale


class CodexButton(Canvas):
    def __init__(
        self,
        master,
        *,
        text: str = "",
        command=None,
        textvariable: StringVar | None = None,
        icon: str = "",
        variant: str = "secondary",
        width: int | None = None,
        height: int = 34,
        min_width: int = 92,
    ) -> None:
        self._scale = _widget_ui_scale(master)
        self._text = text
        self._command = command
        self._textvariable = textvariable
        self._icon = icon
        self._variant = variant
        self._state = "normal"
        self._hover = False
        self._height = self._px(height)
        self._min_width = self._px(min_width)
        self._variable_trace: str | None = None
        display_text = self._display_text()
        initial_width = self._px(width) if width is not None else self._measure_width(display_text, icon, self._min_width)
        self._canvas_bg = self._resolve_parent_bg(master)
        super().__init__(
            master,
            width=initial_width,
            height=self._height,
            bg=self._canvas_bg,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        if textvariable is not None:
            self._variable_trace = textvariable.trace_add("write", lambda *_args: self._redraw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Configure>", lambda _event: self._redraw())
        self._redraw()

    def configure(self, cnf=None, **kwargs):  # type: ignore[override]
        if cnf:
            kwargs.update(cnf)
        if "text" in kwargs:
            self._text = kwargs.pop("text")
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
            super().configure(cursor="" if self._state == "disabled" else "hand2")
        if "textvariable" in kwargs:
            self._textvariable = kwargs.pop("textvariable")
        if "icon" in kwargs:
            self._icon = kwargs.pop("icon")
        if "variant" in kwargs:
            self._variant = kwargs.pop("variant")
        if kwargs:
            super().configure(**kwargs)
        self._redraw()

    config = configure

    @staticmethod
    def _resolve_parent_bg(master) -> str:
        try:
            background = master.cget("background")
            if background:
                return background
        except Exception:
            pass
        try:
            style_name = master.cget("style")
            if style_name:
                background = ttk.Style(master).lookup(style_name, "background")
                if background:
                    return background
        except Exception:
            pass
        return COLOR_BG

    def _px(self, value: int | float) -> int:
        return _scale_px(value, self._scale)

    def _pxf(self, value: int | float) -> float:
        return _scale_float(value, self._scale)

    def _display_text(self) -> str:
        if self._textvariable is not None:
            return self._textvariable.get()
        return self._text

    def _measure_width(self, text: str, icon: str, min_width: int) -> int:
        text_units = sum(2 if ord(char) > 127 else 1 for char in text)
        width = self._px(28) + text_units * self._px(7)
        if icon:
            width += self._px(20)
        return max(min_width, width)

    def _palette(self) -> tuple[str, str, str, str]:
        if self._state == "disabled":
            return COLOR_SURFACE_PRESSED, COLOR_SURFACE_PRESSED, COLOR_DISABLED, COLOR_BORDER
        if self._variant == "primary":
            return COLOR_PRIMARY, COLOR_PRIMARY_ACTIVE, "#ffffff", COLOR_PRIMARY
        if self._variant == "tonal":
            return COLOR_SURFACE_PRESSED, COLOR_NAV_SELECTED, COLOR_NAV_TEXT, COLOR_SURFACE_PRESSED
        return COLOR_SURFACE, COLOR_SURFACE_ALT, COLOR_NAV_TEXT, COLOR_BORDER

    def _on_enter(self, _event=None) -> None:
        self._hover = True
        self._redraw()

    def _on_leave(self, _event=None) -> None:
        self._hover = False
        self._redraw()

    def _on_click(self, _event=None) -> None:
        if self._state == "disabled" or self._command is None:
            return
        self._command()

    def _redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), int(float(self.cget("width"))))
        height = max(self.winfo_height(), self._height)
        text = self._display_text()
        family = self.master.winfo_toplevel().tk.call("font", "actual", "TkDefaultFont", "-family")
        if self._variant == "link":
            if self._state == "disabled":
                foreground = COLOR_DISABLED
            else:
                foreground = COLOR_PRIMARY_ACTIVE if self._hover else COLOR_PRIMARY
            font = (family, _font_size(10))
            content = f"{self._icon} {text}".strip() if self._icon else text
            self.create_text(width / 2, height / 2, text=content, fill=foreground, font=font)
            return
        normal, active, foreground, border = self._palette()
        fill = active if self._hover and self._state != "disabled" else normal
        inset = self._pxf(1)
        self._draw_round_rect(inset, inset, width - inset, height - inset, self._pxf(9), fill=fill, outline=border, width=self._pxf(1))
        font = (family, _font_size(10), "bold") if self._variant == "primary" else (family, _font_size(10))
        if self._icon:
            content_width = self._measure_width(text, self._icon, 0) - self._px(28)
            start_x = max((width - content_width) / 2, self._pxf(12))
            self.create_text(start_x + self._pxf(7), height / 2, text=self._icon, fill=foreground, font=font, anchor="center")
            self.create_text(start_x + self._pxf(22), height / 2, text=text, fill=foreground, font=font, anchor="w")
        else:
            self.create_text(width / 2, height / 2, text=text, fill=foreground, font=font)

    def _draw_round_rect(self, x1: float, y1: float, x2: float, y2: float, radius: float, **kwargs) -> None:
        radius = max(0, min(radius, (x2 - x1) / 2, (y2 - y1) / 2))
        self.create_polygon(
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1,
            smooth=True,
            splinesteps=16,
            **kwargs,
        )


def _paint_tool_icon(canvas: Canvas, icon_id: str, color: str, x: float, y: float, size: float, line_width: float) -> None:
    """在 size×size 的方框内绘制工具线性图标（坐标按设计稿 14×14 视图换算）。"""

    def px(value: float) -> float:
        return x + value * size / 14.0

    def py(value: float) -> float:
        return y + value * size / 14.0

    line = {"fill": color, "width": line_width, "capstyle": "round", "joinstyle": "round"}
    if icon_id == "social_security":
        canvas.create_rectangle(px(2.5), py(1.5), px(11.5), py(12.5), outline=color, width=line_width)
        canvas.create_line(px(5), py(5), px(9), py(5), **line)
        canvas.create_line(px(5), py(8), px(9), py(8), **line)
    elif icon_id == "insurance_ledger":
        canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width)
        canvas.create_line(px(4.6), py(7), px(6.3), py(8.7), px(9.6), py(5.4), **line)
    elif icon_id == "data_statistics":
        canvas.create_line(px(2.5), py(12), px(2.5), py(7), **line)
        canvas.create_line(px(7), py(12), px(7), py(2.5), **line)
        canvas.create_line(px(11.5), py(12), px(11.5), py(5), **line)
    elif icon_id == "salary_split":
        canvas.create_line(px(7), py(2), px(7), py(6), **line)
        canvas.create_line(px(7), py(6), px(3), py(11), **line)
        canvas.create_line(px(7), py(6), px(11), py(11), **line)
    elif icon_id == "salary_merge":
        canvas.create_line(px(3), py(3), px(7), py(8), **line)
        canvas.create_line(px(11), py(3), px(7), py(8), **line)
        canvas.create_line(px(7), py(8), px(7), py(12), **line)
    elif icon_id == "personnel_change_merge":
        canvas.create_line(px(2), py(4.5), px(10), py(4.5), **line)
        canvas.create_line(px(8), py(2), px(10.5), py(4.5), px(8), py(7), **line)
        canvas.create_line(px(12), py(9.5), px(4), py(9.5), **line)
        canvas.create_line(px(6), py(7), px(3.5), py(9.5), px(6), py(12), **line)
    elif icon_id == "archive_import":
        canvas.create_rectangle(px(2), py(4.5), px(12), py(12), outline=color, width=line_width)
        canvas.create_line(px(2), py(7), px(12), py(7), **line)
        canvas.create_line(px(7), py(4.5), px(7), py(2.5), **line)
    elif icon_id == "material_collector":
        canvas.create_rectangle(px(2.5), py(3), px(11.5), py(11.5), outline=color, width=line_width)
        canvas.create_line(px(2.5), py(6), px(11.5), py(6), **line)
        canvas.create_line(px(7), py(6), px(7), py(11.5), **line)
    elif icon_id == "folder_rename":
        canvas.create_line(
            px(2), py(10.5), px(2), py(4), px(3.5), py(2.5), px(6), py(2.5), px(7.2), py(4),
            px(10.5), py(4), px(12), py(5.5), px(12), py(10.5), px(10.5), py(12), px(3.5), py(12), px(2), py(10.5),
            **line,
        )
    elif icon_id == "tutorial":
        canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width)
        canvas.create_line(px(7), py(6.5), px(7), py(10), **line)
        canvas.create_line(px(7), py(4), px(7), py(4.45), **line)
    elif icon_id == "clock":
        canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width)
        canvas.create_line(px(7), py(4), px(7), py(7.2), px(9), py(8.6), **line)
    else:
        canvas.create_oval(px(3), py(3), px(11), py(11), outline=color, width=line_width)


class SidebarItem(Canvas):
    """侧边栏导航条目：圆角底 + 线性图标 + 文字（对应设计稿导航行）。"""

    def __init__(
        self,
        master,
        *,
        text: str,
        icon_id: str,
        command=None,
        height: int = 32,
        muted: bool = False,
    ) -> None:
        self._scale = _widget_ui_scale(master)
        self._text = text
        self._icon_id = icon_id
        self._command = command
        self._muted = muted
        self._selected = False
        self._hover = False
        super().__init__(
            master,
            height=_scale_px(height, self._scale),
            bg=COLOR_SIDEBAR,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Configure>", lambda _event: self._redraw())
        self._redraw()

    def _px(self, value: int | float) -> int:
        return _scale_px(value, self._scale)

    def _pxf(self, value: int | float) -> float:
        return _scale_float(value, self._scale)

    def set_selected(self, selected: bool) -> None:
        if self._selected != selected:
            self._selected = selected
            self._redraw()

    def _on_enter(self, _event=None) -> None:
        self._hover = True
        self._redraw()

    def _on_leave(self, _event=None) -> None:
        self._hover = False
        self._redraw()

    def _on_click(self, _event=None) -> None:
        if self._command is not None:
            self._command()

    def _redraw(self) -> None:
        self.delete("all")
        width = max(self.winfo_width(), 1)
        height = max(self.winfo_height(), 1)
        if self._selected:
            fill = COLOR_NAV_SELECTED
        elif self._hover:
            fill = COLOR_NAV_HOVER
        else:
            fill = COLOR_SIDEBAR
        if fill != COLOR_SIDEBAR:
            radius = self._pxf(7)
            CodexButton._draw_round_rect(self, 0, 0, width, height, radius, fill=fill, outline="")
        if self._selected:
            foreground = COLOR_NAV_TEXT_SELECTED
        elif self._muted:
            foreground = COLOR_MUTED
        else:
            foreground = COLOR_NAV_TEXT
        icon_size = self._pxf(15)
        icon_x = self._pxf(9)
        icon_y = (height - icon_size) / 2
        _paint_tool_icon(self, self._icon_id, foreground, icon_x, icon_y, icon_size, max(1.0, self._pxf(1.4)))
        family = self.winfo_toplevel().tk.call("font", "actual", "TkDefaultFont", "-family")
        font = (family, _font_size(10), "bold") if self._selected else (family, _font_size(10))
        self.create_text(icon_x + icon_size + self._pxf(9), height / 2, text=self._text, fill=foreground, font=font, anchor="w")


class RoundedCard(Canvas):
    """白色圆角卡片容器（设计稿 border-radius:14px 的卡片分区）。"""

    def __init__(
        self,
        master,
        *,
        padding: tuple[int, int, int, int] = (20, 16, 20, 18),
        radius: int = 14,
        fill_height: bool = False,
        min_height: int = 0,
    ) -> None:
        self._scale = _widget_ui_scale(master)
        self._radius = _scale_float(radius, self._scale)
        self._fill_height = fill_height
        page_bg = CodexButton._resolve_parent_bg(master)
        super().__init__(
            master,
            bg=page_bg,
            highlightthickness=0,
            bd=0,
            height=_scale_px(min_height, self._scale) if min_height else 1,
        )
        self.inner = ttk.Frame(self, style="InputWrap.TFrame")
        self._pads = (0, 0, 0, 0)
        self._window = self.create_window(0, 0, window=self.inner, anchor="nw")
        self._last_bg_size = (0, 0)
        self.set_padding(padding, sync=False)
        self.inner.bind("<Configure>", self._sync)
        self.bind("<Configure>", self._sync)

    def set_padding(self, padding: tuple[int, int, int, int], *, sync: bool = True) -> None:
        self._pads = tuple(_scale_px(value, self._scale) for value in padding)
        self.coords(self._window, self._pads[0], self._pads[1])
        if sync:
            self._sync()

    def _sync(self, _event=None) -> None:
        left, top, right, bottom = self._pads
        width = max(self.winfo_width(), 1)
        inner_width = max(width - left - right, 1)
        if self._fill_height:
            height = max(self.winfo_height(), 1)
            self.itemconfigure(self._window, width=inner_width, height=max(height - top - bottom, 1))
        else:
            self.itemconfigure(self._window, width=inner_width)
            height = self.inner.winfo_reqheight() + top + bottom
            if int(float(self.cget("height"))) != height:
                self.configure(height=height)
        if (width, height) != self._last_bg_size:
            self._last_bg_size = (width, height)
            self._redraw_bg(width, height)

    def _redraw_bg(self, width: int, height: int) -> None:
        self.delete("card_bg")
        offset = max(1.0, _scale_float(1.5, self._scale))
        CodexButton._draw_round_rect(
            self, offset, offset * 1.6, width - offset * 0.4, height, self._radius, fill="#ECE9E2", outline="", tags="card_bg"
        )
        CodexButton._draw_round_rect(
            self, 0, 0, width - offset, height - offset, self._radius, fill=COLOR_SURFACE, outline="", tags="card_bg"
        )
        self.tag_lower("card_bg")


__all__ = [
    "CodexButton",
    "SidebarItem",
    "RoundedCard",
    "_paint_tool_icon",
]
