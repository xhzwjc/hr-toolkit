"""生成应用图标资源。

图标是品牌绿圆角方块 + 白色 "HR" 字标，与主界面侧栏的品牌标识一致。
仅用标准库实现：带符号距离场（SDF）抗锯齿光栅化 + 手写 PNG/ICO 编码，
不引入 Pillow 等图像依赖。

输出：
- hr_toolkit/_icon_data.py        运行时窗口图标（base64 PNG，Tk iconphoto 使用）
- packaging/windows/HRToolkit.ico Windows exe 图标（PyInstaller --icon 使用）
- release/app_icon_preview.png    256px 预览图，便于人工检查

用法：python scripts/generate_app_icons.py
"""

from __future__ import annotations

import base64
import math
import struct
import sys
import zlib
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
ICON_DATA_FILE = REPO_ROOT / "hr_toolkit" / "_icon_data.py"
ICO_FILE = REPO_ROOT / "packaging" / "windows" / "HRToolkit.ico"
PREVIEW_FILE = REPO_ROOT / "release" / "app_icon_preview.png"

# 从大到小排列：Tk 在 macOS 上只取第一张作为 Dock 图标，
# 必须让最大尺寸排在最前，否则会拿小图放大导致模糊
RUNTIME_PNG_SIZES = (512, 256, 128, 64, 32, 16)
ICO_BMP_SIZES = (16, 24, 32, 48, 64)
ICO_PNG_SIZES = (256,)

# 品牌绿渐变（上浅下深，中值即主界面的 COLOR_PRIMARY #007f5f）
GRADIENT_TOP = (0, 148, 111)
GRADIENT_BOTTOM = (0, 102, 76)
LETTER_COLOR = (255, 255, 255)

# 以下坐标均在 64x64 设计网格中
CORNER_RADIUS = 14.0

# 字标几何参数。小尺寸（任务栏 16px）单独做视觉修正：
# 笔画更粗、字距更大，否则 H 和 R 会糊成一团。
_GLYPHS_DEFAULT = {
    "stroke": 3.3,  # 笔画半宽，总宽约 6.6
    "h_left_x": 16.0,
    "h_right_x": 27.5,
    "h_top_y": 19.5,
    "h_bottom_y": 44.5,
    "h_bar_y": 31.5,
    "r_stem_x": 38.5,
    "r_bowl_radius": 6.2,
    "r_leg_start": (40.5, 33.0),
    "r_leg_end": (47.5, 44.5),
}
_GLYPHS_SMALL = {
    "stroke": 3.6,
    "h_left_x": 13.5,
    "h_right_x": 25.0,
    "h_top_y": 18.0,
    "h_bottom_y": 46.0,
    "h_bar_y": 32.0,
    "r_stem_x": 41.0,
    "r_bowl_radius": 7.6,
    "r_leg_start": (43.5, 35.0),
    "r_leg_end": (50.0, 46.0),
}
_SMALL_GLYPH_MAX_SIZE = 20


def _rounded_box_sdf(x: float, y: float, cx: float, cy: float, half_w: float, half_h: float, radius: float) -> float:
    qx = abs(x - cx) - (half_w - radius)
    qy = abs(y - cy) - (half_h - radius)
    outside = math.hypot(max(qx, 0.0), max(qy, 0.0))
    inside = min(max(qx, qy), 0.0)
    return outside + inside - radius


def _capsule_sdf(x: float, y: float, ax: float, ay: float, bx: float, by: float, radius: float) -> float:
    px, py = x - ax, y - ay
    dx, dy = bx - ax, by - ay
    t = max(0.0, min(1.0, (px * dx + py * dy) / (dx * dx + dy * dy)))
    return math.hypot(px - dx * t, py - dy * t) - radius


def _half_ring_sdf(x: float, y: float, cx: float, cy: float, radius: float, half_width: float) -> float:
    """右半圆环（带圆头端点），用于 R 的弧。"""
    if x >= cx:
        return abs(math.hypot(x - cx, y - cy) - radius) - half_width
    top = math.hypot(x - cx, y - (cy - radius))
    bottom = math.hypot(x - cx, y - (cy + radius))
    return min(top, bottom) - half_width


def _letters_sdf(x: float, y: float, glyphs: dict) -> float:
    stroke = glyphs["stroke"]
    top, bottom = glyphs["h_top_y"], glyphs["h_bottom_y"]
    h_left = _capsule_sdf(x, y, glyphs["h_left_x"], top, glyphs["h_left_x"], bottom, stroke)
    h_right = _capsule_sdf(x, y, glyphs["h_right_x"], top, glyphs["h_right_x"], bottom, stroke)
    h_bar = _capsule_sdf(x, y, glyphs["h_left_x"], glyphs["h_bar_y"], glyphs["h_right_x"], glyphs["h_bar_y"], stroke)
    r_stem = _capsule_sdf(x, y, glyphs["r_stem_x"], top, glyphs["r_stem_x"], bottom, stroke)
    r_bowl = _half_ring_sdf(x, y, glyphs["r_stem_x"], top + glyphs["r_bowl_radius"], glyphs["r_bowl_radius"], stroke)
    r_leg = _capsule_sdf(x, y, *glyphs["r_leg_start"], *glyphs["r_leg_end"], stroke)
    return min(h_left, h_right, h_bar, r_stem, r_bowl, r_leg)


def _coverage(distance: float, pixel_size: float) -> float:
    return max(0.0, min(1.0, 0.5 - distance / pixel_size))


def render_icon(size: int) -> list[list[tuple[int, int, int, int]]]:
    """在 64 单位设计网格中按 SDF 采样，返回 RGBA 像素行。"""
    glyphs = _GLYPHS_SMALL if size <= _SMALL_GLYPH_MAX_SIZE else _GLYPHS_DEFAULT
    # 小尺寸靠超采样保细节；大尺寸 SDF 自带抗锯齿已足够，降低采样省时间
    supersample = 3 if size <= 64 else 1
    scale = 64.0 / size
    sample_step = scale / supersample
    rows: list[list[tuple[int, int, int, int]]] = []
    for py in range(size):
        row: list[tuple[int, int, int, int]] = []
        for px in range(size):
            acc_r = acc_g = acc_b = acc_a = 0.0
            for sy in range(supersample):
                for sx in range(supersample):
                    x = (px + (sx + 0.5) / supersample) * scale
                    y = (py + (sy + 0.5) / supersample) * scale
                    bg_alpha = _coverage(_rounded_box_sdf(x, y, 32.0, 32.0, 32.0, 32.0, CORNER_RADIUS), sample_step)
                    if bg_alpha <= 0.0:
                        continue
                    t = y / 64.0
                    r = GRADIENT_TOP[0] + (GRADIENT_BOTTOM[0] - GRADIENT_TOP[0]) * t
                    g = GRADIENT_TOP[1] + (GRADIENT_BOTTOM[1] - GRADIENT_TOP[1]) * t
                    b = GRADIENT_TOP[2] + (GRADIENT_BOTTOM[2] - GRADIENT_TOP[2]) * t
                    letter_alpha = _coverage(_letters_sdf(x, y, glyphs), sample_step)
                    if letter_alpha > 0.0:
                        r += (LETTER_COLOR[0] - r) * letter_alpha
                        g += (LETTER_COLOR[1] - g) * letter_alpha
                        b += (LETTER_COLOR[2] - b) * letter_alpha
                    acc_r += r * bg_alpha
                    acc_g += g * bg_alpha
                    acc_b += b * bg_alpha
                    acc_a += bg_alpha
            samples = supersample * supersample
            alpha = acc_a / samples
            if alpha <= 0.0:
                row.append((0, 0, 0, 0))
                continue
            # acc_* 为 alpha 加权累计值，除以 acc_a 还原直通（非预乘）颜色
            row.append(
                (
                    int(round(acc_r / acc_a)),
                    int(round(acc_g / acc_a)),
                    int(round(acc_b / acc_a)),
                    int(round(alpha * 255)),
                )
            )
        rows.append(row)
    return rows


def encode_png(rows: list[list[tuple[int, int, int, int]]]) -> bytes:
    size = len(rows)

    def chunk(tag: bytes, payload: bytes) -> bytes:
        data = tag + payload
        return struct.pack(">I", len(payload)) + data + struct.pack(">I", zlib.crc32(data))

    raw = b"".join(b"\x00" + bytes(channel for pixel in row for channel in pixel) for row in rows)
    return b"".join(
        (
            b"\x89PNG\r\n\x1a\n",
            chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)),
            chunk(b"IDAT", zlib.compress(raw, 9)),
            chunk(b"IEND", b""),
        )
    )


def encode_ico(bmp_sizes: tuple[int, ...], png_sizes: tuple[int, ...]) -> bytes:
    entries: list[tuple[int, bytes]] = []
    for size in bmp_sizes:
        entries.append((size, _encode_ico_bmp(render_icon(size))))
    for size in png_sizes:
        entries.append((size, encode_png(render_icon(size))))

    count = len(entries)
    header = struct.pack("<HHH", 0, 1, count)
    directory = b""
    offset = len(header) + count * 16
    for size, payload in entries:
        directory += struct.pack(
            "<BBBBHHII",
            size % 256,
            size % 256,
            0,
            0,
            1,
            32,
            len(payload),
            offset,
        )
        offset += len(payload)
    return header + directory + b"".join(payload for _size, payload in entries)


def _encode_ico_bmp(rows: list[list[tuple[int, int, int, int]]]) -> bytes:
    """32 位 BGRA DIB（老版本 Windows 对小尺寸 PNG 条目兼容性差）。"""
    size = len(rows)
    header = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0, size * size * 4, 0, 0, 0, 0)
    xor_data = b"".join(
        bytes(channel for r, g, b, a in row for channel in (b, g, r, a))
        for row in reversed(rows)
    )
    and_stride = ((size + 31) // 32) * 4
    and_mask = b"\x00" * (and_stride * size)
    return header + xor_data + and_mask


def write_icon_data_module(path: Path) -> None:
    lines = [
        '"""应用窗口图标数据。',
        "",
        "由 scripts/generate_app_icons.py 生成，请勿手工修改；",
        "调整图标请改脚本后重新生成。",
        '"""',
        "",
        "# fmt: off",
        "APP_ICON_PNGS_BASE64 = {",
    ]
    for size in RUNTIME_PNG_SIZES:
        encoded = base64.b64encode(encode_png(render_icon(size))).decode("ascii")
        lines.append(f'    {size}: "{encoded}",')
    lines.append("}")
    lines.append("# fmt: on")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    write_icon_data_module(ICON_DATA_FILE)
    print(f"已生成运行时图标数据：{ICON_DATA_FILE}")

    ICO_FILE.parent.mkdir(parents=True, exist_ok=True)
    ICO_FILE.write_bytes(encode_ico(ICO_BMP_SIZES, ICO_PNG_SIZES))
    print(f"已生成 Windows 图标：{ICO_FILE}")

    PREVIEW_FILE.parent.mkdir(parents=True, exist_ok=True)
    PREVIEW_FILE.write_bytes(encode_png(render_icon(256)))
    print(f"已生成预览图：{PREVIEW_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
