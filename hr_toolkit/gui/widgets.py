"""Custom Canvas-rendered GUI widgets with anti-aliased rounded corners and macOS/Notion styling."""

from __future__ import annotations

import math

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

_DEFAULT_FONT_FAMILY: str | None = None


def _get_default_font_family(master) -> str:
    global _DEFAULT_FONT_FAMILY
    if _DEFAULT_FONT_FAMILY is None:
        try:
            _DEFAULT_FONT_FAMILY = master.winfo_toplevel().tk.call("font", "actual", "TkDefaultFont", "-family")
        except Exception:
            _DEFAULT_FONT_FAMILY = "Helvetica"
    return _DEFAULT_FONT_FAMILY


def _tessellate_round_rect(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    radius: float,
    segments_per_corner: int = 6,
) -> list[float]:
    """Return a dense list of (x, y) points approximating a rounded rectangle.

    Tk's ``create_polygon(smooth=True, splinesteps=16)`` is one of the slowest
    drawing primitives in Tkinter (12 vertices × 16 splinesteps ≈ 192 anti-aliased
    line segments per polygon) and its anti-aliasing is fragile under rapid
    redraws — visible as TV-static / ghosting during the initial render.

    Pre-tessellating the corners in Python produces a polygon of plain straight
    segments that Tk renders cheaply with no smoothing artifacts. At the small
    corner radii used by CodexButton (9px), RoundedCard (14px) and SidebarItem
    (7px), 6 straight segments per quadrant is visually indistinguishable from
    the smoothed curve.
    """
    if segments_per_corner < 2:
        segments_per_corner = 2
    radius = max(0.0, min(radius, (x2 - x1) / 2.0, (y2 - y1) / 2.0))
    if radius <= 0:
        return [x1, y1, x2, y1, x2, y2, x1, y2]

    pts: list[float] = []

    # 1. Top-right corner (from (x2 - radius, y1) to (x2, y1 + radius))
    cx, cy = x2 - radius, y1 + radius
    for i in range(segments_per_corner + 1):
        angle = (math.pi / 2.0) * (1.0 - i / segments_per_corner)
        pts.extend([cx + radius * math.cos(angle), cy - radius * math.sin(angle)])

    # 2. Bottom-right corner (from (x2, y2 - radius) to (x2 - radius, y2))
    cx, cy = x2 - radius, y2 - radius
    for i in range(1, segments_per_corner + 1):
        angle = (math.pi / 2.0) * (i / segments_per_corner)
        pts.extend([cx + radius * math.cos(angle), cy + radius * math.sin(angle)])

    # 3. Bottom-left corner (from (x1 + radius, y2) to (x1, y2 - radius))
    cx, cy = x1 + radius, y2 - radius
    for i in range(1, segments_per_corner + 1):
        angle = (math.pi / 2.0) * (1.0 - i / segments_per_corner)
        pts.extend([cx - radius * math.cos(angle), cy + radius * math.sin(angle)])

    # 4. Top-left corner (from (x1, y1 + radius) to (x1 + radius, y1))
    cx, cy = x1 + radius, y1 + radius
    for i in range(1, segments_per_corner):
        angle = (math.pi / 2.0) * (i / segments_per_corner)
        pts.extend([cx - radius * math.cos(angle), cy - radius * math.sin(angle)])

    return pts


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
        self._last_draw_key: tuple | None = None
        self._poly_id: int | None = None
        self._icon_item_id: int | None = None
        self._text_id: int | None = None
        display_text = self._display_text()
        initial_width = self._px(width) if width is not None else self._measure_width(display_text, icon, self._min_width)
        self._requested_width = initial_width
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
        self._set_textvariable(textvariable)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Configure>", self._on_configure)
        self._redraw()

    def _on_configure(self, _event=None) -> None:
        if _event is not None:
            w = max(int(getattr(_event, "width", 1)), 1)
            h = max(int(getattr(_event, "height", self._height)), 1)
        else:
            w = max(self.winfo_width(), self._requested_width)
            h = max(self.winfo_height(), self._height)
        if self._last_draw_key is None or (w, h) != (self._last_draw_key[0], self._last_draw_key[1]):
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
            self._set_textvariable(kwargs.pop("textvariable"))
        if "icon" in kwargs:
            self._icon = kwargs.pop("icon")
        if "variant" in kwargs:
            self._variant = kwargs.pop("variant")
        if "width" in kwargs:
            try:
                self._requested_width = max(1, int(float(kwargs["width"])))
            except (TypeError, ValueError):
                pass
        if kwargs:
            super().configure(**kwargs)
        self._redraw()

    config = configure

    def _set_textvariable(self, variable: StringVar | None) -> None:
        if self._textvariable is not None and self._variable_trace is not None:
            try:
                self._textvariable.trace_remove("write", self._variable_trace)
            except Exception:
                pass
        self._textvariable = variable
        self._variable_trace = None
        if variable is not None:
            self._variable_trace = variable.trace_add("write", lambda *_args: self._redraw())

    def destroy(self) -> None:
        self._set_textvariable(None)
        super().destroy()

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
        width = max(self.winfo_width(), int(float(self.cget("width"))))
        height = max(self.winfo_height(), self._height)
        text = self._display_text()
        key = (width, height, text, self._state, self._hover, self._variant, self._icon, self._canvas_bg)
        if key == self._last_draw_key:
            return
        self._last_draw_key = key

        family = _get_default_font_family(self)
        if self._variant == "link":
            if self._state == "disabled":
                foreground = COLOR_DISABLED
            else:
                foreground = COLOR_PRIMARY_ACTIVE if self._hover else COLOR_PRIMARY
            font = (family, _font_size(10))
            content = f"{self._icon} {text}".strip() if self._icon else text
            if self._poly_id is not None:
                self.delete(self._poly_id)
                self._poly_id = None
            if self._icon_item_id is not None:
                self.delete(self._icon_item_id)
                self._icon_item_id = None
            if self._text_id is None or not self.find_withtag(self._text_id):
                self._text_id = self.create_text(width / 2, height / 2, text=content, fill=foreground, font=font, anchor="center")
            else:
                self.coords(self._text_id, width / 2, height / 2)
                self.itemconfigure(self._text_id, text=content, fill=foreground, font=font, anchor="center")
            return

        normal, active, foreground, border = self._palette()
        fill = active if self._hover and self._state != "disabled" else normal
        inset = self._pxf(1)
        poly_pts = _tessellate_round_rect(inset, inset, width - inset, height - inset, self._pxf(9))

        if self._poly_id is None or not self.find_withtag(self._poly_id):
            self._poly_id = self.create_polygon(*poly_pts, fill=fill, outline=border, width=self._pxf(1))
            self.tag_lower(self._poly_id)
        else:
            self.coords(self._poly_id, *poly_pts)
            self.itemconfigure(self._poly_id, fill=fill, outline=border, width=self._pxf(1))

        font = (family, _font_size(10), "bold") if self._variant == "primary" else (family, _font_size(10))
        if self._icon:
            content_width = self._measure_width(text, self._icon, 0) - self._px(28)
            start_x = max((width - content_width) / 2, self._pxf(12))
            icon_x = start_x + self._pxf(7)
            text_x = start_x + self._pxf(22)

            if self._icon_item_id is None or not self.find_withtag(self._icon_item_id):
                self._icon_item_id = self.create_text(icon_x, height / 2, text=self._icon, fill=foreground, font=font, anchor="center")
            else:
                self.coords(self._icon_item_id, icon_x, height / 2)
                self.itemconfigure(self._icon_item_id, text=self._icon, fill=foreground, font=font, anchor="center")

            if self._text_id is None or not self.find_withtag(self._text_id):
                self._text_id = self.create_text(text_x, height / 2, text=text, fill=foreground, font=font, anchor="w")
            else:
                self.coords(self._text_id, text_x, height / 2)
                self.itemconfigure(self._text_id, text=text, fill=foreground, font=font, anchor="w")
        else:
            if self._icon_item_id is not None:
                self.delete(self._icon_item_id)
                self._icon_item_id = None
            if self._text_id is None or not self.find_withtag(self._text_id):
                self._text_id = self.create_text(width / 2, height / 2, text=text, fill=foreground, font=font, anchor="center")
            else:
                self.coords(self._text_id, width / 2, height / 2)
                self.itemconfigure(self._text_id, text=text, fill=foreground, font=font, anchor="center")

    def _draw_round_rect(self, x1: float, y1: float, x2: float, y2: float, radius: float, **kwargs) -> None:
        pts = _tessellate_round_rect(x1, y1, x2, y2, radius)
        self.create_polygon(*pts, **kwargs)


def _paint_tool_icon(canvas: Canvas, icon_id: str, color: str, x: float, y: float, size: float, line_width: float) -> list[int]:
    """在 size×size 的方框内绘制工具线性图标（坐标按设计稿 14×14 视图换算）。"""

    item_ids: list[int] = []

    def track(item_id: int) -> int:
        item_ids.append(item_id)
        return item_id

    def px(value: float) -> float:
        return x + value * size / 14.0

    def py(value: float) -> float:
        return y + value * size / 14.0

    line = {"fill": color, "width": line_width, "capstyle": "round", "joinstyle": "round"}
    if icon_id == "social_security":
        track(canvas.create_rectangle(px(2.5), py(1.5), px(11.5), py(12.5), outline=color, width=line_width))
        track(canvas.create_line(px(5), py(5), px(9), py(5), **line))
        track(canvas.create_line(px(5), py(8), px(9), py(8), **line))
    elif icon_id == "insurance_ledger":
        track(canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width))
        track(canvas.create_line(px(4.6), py(7), px(6.3), py(8.7), px(9.6), py(5.4), **line))
    elif icon_id == "data_statistics":
        track(canvas.create_line(px(2.5), py(12), px(2.5), py(7), **line))
        track(canvas.create_line(px(7), py(12), px(7), py(2.5), **line))
        track(canvas.create_line(px(11.5), py(12), px(11.5), py(5), **line))
    elif icon_id == "salary_split":
        track(canvas.create_line(px(7), py(2), px(7), py(6), **line))
        track(canvas.create_line(px(7), py(6), px(3), py(11), **line))
        track(canvas.create_line(px(7), py(6), px(11), py(11), **line))
    elif icon_id == "salary_merge":
        track(canvas.create_line(px(3), py(3), px(7), py(8), **line))
        track(canvas.create_line(px(11), py(3), px(7), py(8), **line))
        track(canvas.create_line(px(7), py(8), px(7), py(12), **line))
    elif icon_id == "personnel_change_merge":
        track(canvas.create_line(px(2), py(4.5), px(10), py(4.5), **line))
        track(canvas.create_line(px(8), py(2), px(10.5), py(4.5), px(8), py(7), **line))
        track(canvas.create_line(px(12), py(9.5), px(4), py(9.5), **line))
        track(canvas.create_line(px(6), py(7), px(3.5), py(9.5), px(6), py(12), **line))
    elif icon_id == "archive_import":
        track(canvas.create_rectangle(px(2), py(4.5), px(12), py(12), outline=color, width=line_width))
        track(canvas.create_line(px(2), py(7), px(12), py(7), **line))
        track(canvas.create_line(px(7), py(4.5), px(7), py(2.5), **line))
    elif icon_id == "material_collector":
        track(canvas.create_rectangle(px(2.5), py(3), px(11.5), py(11.5), outline=color, width=line_width))
        track(canvas.create_line(px(2.5), py(6), px(11.5), py(6), **line))
        track(canvas.create_line(px(7), py(6), px(7), py(11.5), **line))
    elif icon_id == "folder_rename":
        track(canvas.create_line(
            px(2), py(10.5), px(2), py(4), px(3.5), py(2.5), px(6), py(2.5), px(7.2), py(4),
            px(10.5), py(4), px(12), py(5.5), px(12), py(10.5), px(10.5), py(12), px(3.5), py(12), px(2), py(10.5),
            **line,
        ))
    elif icon_id == "tutorial":
        track(canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width))
        track(canvas.create_line(px(7), py(6.5), px(7), py(10), **line))
        track(canvas.create_line(px(7), py(4), px(7), py(4.45), **line))
    elif icon_id == "clock":
        track(canvas.create_oval(px(1.5), py(1.5), px(12.5), py(12.5), outline=color, width=line_width))
        track(canvas.create_line(px(7), py(4), px(7), py(7.2), px(9), py(8.6), **line))
    else:
        track(canvas.create_oval(px(3), py(3), px(11), py(11), outline=color, width=line_width))
    return item_ids


def _paint_codex_badge_icon(
    canvas: Canvas,
    x: float,
    y: float,
    size: float,
    *,
    badge_bg: str = "#E8E6E1",
    badge_border: str = "#DBD8D1",
    icon_color: str = "#6E6C68",
    scale: float = 1.0,
) -> list[int]:
    """在 size×size 的方框内绘制灰色 Codex / 工具箱风格品牌圆角徽标与线性图标。"""
    item_ids: list[int] = []

    def track(item_id: int) -> int:
        item_ids.append(item_id)
        return item_id

    # 1. 浅灰圆角卡片底板 + 柔和阴影
    radius = size * 0.22
    shadow_offset = max(1.0, _scale_float(1.5, scale))
    shadow_pts = _tessellate_round_rect(
        x + shadow_offset,
        y + shadow_offset * 1.5,
        x + size + shadow_offset,
        y + size + shadow_offset * 1.5,
        radius,
    )
    track(canvas.create_polygon(*shadow_pts, fill="#DCD9D2", outline=""))

    badge_pts = _tessellate_round_rect(x, y, x + size, y + size, radius)
    track(canvas.create_polygon(*badge_pts, fill=badge_bg, outline=badge_border, width=max(1.0, _scale_float(1.0, scale))))

    # 2. 居中的工具箱/Codex 线性图标（按 48×48 设计坐标换算）
    def gx(v: float) -> float:
        return x + (v / 48.0) * size

    def gy(v: float) -> float:
        return y + (v / 48.0) * size

    line_w = max(1.5, _scale_float(2.0, scale))
    line_opts = {"fill": icon_color, "width": line_w, "capstyle": "round", "joinstyle": "round"}

    # 工具箱提手拱形
    track(canvas.create_line(gx(18), gy(17), gx(18), gy(12.5), gx(30), gy(12.5), gx(30), gy(17), **line_opts))
    # 工具箱主体圆角轮廓
    body_r = size * 0.08
    body_pts = _tessellate_round_rect(gx(9.5), gy(17), gx(38.5), gy(37.5), body_r, segments_per_corner=3)
    track(canvas.create_polygon(*body_pts, fill="", outline=icon_color, width=line_w))
    # 工具箱水平开合中缝
    track(canvas.create_line(gx(9.5), gy(26), gx(38.5), gy(26), **line_opts))
    # 工具箱中央卡扣锁
    track(canvas.create_rectangle(
        gx(21.5), gy(23.5), gx(26.5), gy(28.5),
        fill=badge_bg, outline=icon_color, width=max(1.0, _scale_float(1.2, scale))
    ))

    return item_ids


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
        self._last_draw_key: tuple | None = None
        self._bg_poly_id: int | None = None
        self._text_item_id: int | None = None
        self._icon_items: list[int] = []
        self._icon_draw_key: tuple | None = None
        self._icon_color: str | None = None
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
        self.bind("<Configure>", self._on_configure)
        self._redraw()

    def _on_configure(self, _event=None) -> None:
        if _event is not None:
            w = max(int(getattr(_event, "width", 1)), 1)
            h = max(int(getattr(_event, "height", 1)), 1)
        else:
            w = max(self.winfo_width(), 1)
            h = max(self.winfo_height(), 1)
        if self._last_draw_key is None or (w, h) != (self._last_draw_key[0], self._last_draw_key[1]):
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
        width = max(self.winfo_width(), 1)
        height = max(self.winfo_height(), 1)
        key = (width, height, self._selected, self._hover, self._muted, self._text, self._icon_id)
        if key == self._last_draw_key:
            return
        self._last_draw_key = key

        if self._selected:
            fill = COLOR_NAV_SELECTED
        elif self._hover:
            fill = COLOR_NAV_HOVER
        else:
            fill = COLOR_SIDEBAR
        show_bg = fill != COLOR_SIDEBAR
        if show_bg:
            radius = self._pxf(7)
            pts = _tessellate_round_rect(0, 0, width, height, radius)
            if self._bg_poly_id is None or not self.find_withtag(self._bg_poly_id):
                self._bg_poly_id = self.create_polygon(*pts, fill=fill, outline="")
                self.tag_lower(self._bg_poly_id)
            else:
                self.coords(self._bg_poly_id, *pts)
                self.itemconfigure(self._bg_poly_id, fill=fill, state="normal")
        else:
            if self._bg_poly_id is not None and self.find_withtag(self._bg_poly_id):
                self.itemconfigure(self._bg_poly_id, state="hidden")

        if self._selected:
            foreground = COLOR_NAV_TEXT_SELECTED
        elif self._muted:
            foreground = COLOR_MUTED
        else:
            foreground = COLOR_NAV_TEXT
        icon_size = self._pxf(15)
        icon_x = self._pxf(9)
        icon_y = (height - icon_size) / 2
        icon_draw_key = (height, self._icon_id)
        icon_items_exist = bool(self._icon_items) and all(
            self.find_withtag(item_id) for item_id in self._icon_items
        )
        if not icon_items_exist or self._icon_draw_key != icon_draw_key:
            for item_id in self._icon_items:
                self.delete(item_id)
            self._icon_items = _paint_tool_icon(
                self,
                self._icon_id,
                foreground,
                icon_x,
                icon_y,
                icon_size,
                max(1.0, self._pxf(1.4)),
            )
            self._icon_draw_key = icon_draw_key
            self._icon_color = foreground
        elif self._icon_color != foreground:
            for item_id in self._icon_items:
                item_type = self.type(item_id)
                if item_type == "line":
                    self.itemconfigure(item_id, fill=foreground)
                else:
                    self.itemconfigure(item_id, outline=foreground)
            self._icon_color = foreground
        family = _get_default_font_family(self)
        font = (family, _font_size(10), "bold") if self._selected else (family, _font_size(10))
        if self._text_item_id is None or not self.find_withtag(self._text_item_id):
            self._text_item_id = self.create_text(
                icon_x + icon_size + self._pxf(9),
                height / 2,
                text=self._text,
                fill=foreground,
                font=font,
                anchor="w",
            )
        else:
            self.coords(self._text_item_id, icon_x + icon_size + self._pxf(9), height / 2)
            self.itemconfigure(self._text_item_id, text=self._text, fill=foreground, font=font)


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
        self._last_inner_size = (0, 0)
        self._shadow_poly: int | None = None
        self._card_poly: int | None = None
        self._syncing = False
        self._sync_job = None
        self._last_self_event_size = (0, 0)
        self._last_inner_event_size = (0, 0)
        self.set_padding(padding, sync=False)
        self.inner.bind("<Configure>", self._schedule_sync)
        self.bind("<Configure>", self._schedule_sync)

    def set_padding(self, padding: tuple[int, int, int, int], *, sync: bool = True) -> None:
        self._pads = tuple(_scale_px(value, self._scale) for value in padding)
        self.coords(self._window, self._pads[0], self._pads[1])
        if sync:
            self._sync()

    def _schedule_sync(self, _event=None) -> None:
        if _event is not None:
            event_size = (
                max(int(getattr(_event, "width", 1)), 1),
                max(int(getattr(_event, "height", 1)), 1),
            )
            if getattr(_event, "widget", None) is self.inner:
                if event_size == self._last_inner_event_size:
                    return
                self._last_inner_event_size = event_size
            else:
                if event_size == self._last_self_event_size:
                    return
                self._last_self_event_size = event_size
        if self._sync_job is not None:
            return
        try:
            self._sync_job = self.after_idle(self._run_scheduled_sync)
        except Exception:
            self._sync_job = None

    def _run_scheduled_sync(self) -> None:
        self._sync_job = None
        self._sync()

    def _sync(self, _event=None) -> None:
        if self._syncing:
            return
        self._syncing = True
        try:
            left, top, right, bottom = self._pads
            width = max(self.winfo_width(), 1)
            inner_width = max(width - left - right, 1)
            if self._fill_height:
                height = max(self.winfo_height(), 1)
                inner_height = max(height - top - bottom, 1)
                if (inner_width, inner_height) != self._last_inner_size:
                    self._last_inner_size = (inner_width, inner_height)
                    self.itemconfigure(self._window, width=inner_width, height=inner_height)
            else:
                if inner_width != self._last_inner_size[0]:
                    self._last_inner_size = (inner_width, 0)
                    self.itemconfigure(self._window, width=inner_width)
                req_h = max(self.inner.winfo_reqheight() + top + bottom, 1)
                try:
                    curr_h = int(float(self.cget("height")))
                except Exception:
                    curr_h = 0
                if curr_h != req_h:
                    self.configure(height=req_h)
                height = req_h
            if (width, height) != self._last_bg_size:
                self._last_bg_size = (width, height)
                self._redraw_bg(width, height)
        finally:
            self._syncing = False

    def destroy(self) -> None:
        job = self._sync_job
        self._sync_job = None
        if job is not None:
            try:
                self.after_cancel(job)
            except Exception:
                pass
        super().destroy()

    def _redraw_bg(self, width: int, height: int) -> None:
        offset = max(1.0, _scale_float(1.5, self._scale))
        shadow_pts = _tessellate_round_rect(offset, offset * 1.6, width - offset * 0.4, height, self._radius)
        card_pts = _tessellate_round_rect(0, 0, width - offset, height - offset, self._radius)

        if self._shadow_poly is None or not self.find_withtag(self._shadow_poly):
            self._shadow_poly = self.create_polygon(
                *shadow_pts, fill="#ECE9E2", outline="", tags="card_bg"
            )
            self._card_poly = self.create_polygon(
                *card_pts, fill=COLOR_SURFACE, outline="", tags="card_bg"
            )
            self.tag_lower("card_bg")
        else:
            self.coords(self._shadow_poly, *shadow_pts)
            self.coords(self._card_poly, *card_pts)


__all__ = [
    "CodexButton",
    "SidebarItem",
    "RoundedCard",
    "_paint_tool_icon",
    "_paint_codex_badge_icon",
    "_get_default_font_family",
]
