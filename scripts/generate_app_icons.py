"""从正式品牌源图生成应用图标资源。

品牌源图是 F1 暖陶土色抽象标识。脚本只使用标准库解析 8 位 RGBA PNG，
再以预乘 Alpha 超采样缩放生成各平台尺寸，不依赖 Pillow 等图像库。

输出：
- hr_toolkit/_icon_data.py        运行时窗口图标（base64 PNG，Tk iconphoto 使用）
- packaging/windows/HRToolkit.ico Windows exe 图标（PyInstaller --icon 使用）
- release/app_icon_preview.png    256px 预览图，便于人工检查

用法：python scripts/generate_app_icons.py
"""

from __future__ import annotations

import base64
import hashlib
import math
import struct
import sys
import zlib
from functools import lru_cache
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE_ICON_FILE = REPO_ROOT / "packaging" / "icons" / "HRToolkit-source.png"
ICON_DATA_FILE = REPO_ROOT / "hr_toolkit" / "_icon_data.py"
ICO_FILE = REPO_ROOT / "packaging" / "windows" / "HRToolkit.ico"
PREVIEW_FILE = REPO_ROOT / "release" / "app_icon_preview.png"
SOURCE_ICON_SHA256 = "f567181bc26f828657cc0dc53f4c226caac1461f33e5a8b6303e823519324c9e"

# 从大到小排列：Tk 在 macOS 上只取第一张作为 Dock 图标，
# 必须让最大尺寸排在最前，否则会拿小图放大导致模糊
RUNTIME_PNG_SIZES = (512, 256, 128, 64, 32, 16)
# 侧栏品牌标识与启动页会按 Windows DPI 选择最接近的原生像素尺寸，
# 避免 Tk 对位图做整数倍缩放而产生模糊或裁切。
BRAND_MARK_PNG_SIZES = (26, 32, 39, 46, 52, 64, 78, 96, 112, 128, 144, 160, 192)
ICO_BMP_SIZES = (16, 24, 32, 48, 64)
ICO_PNG_SIZES = (256,)
MIN_SOURCE_COMPONENT_PIXELS = 1_000


def _paeth_predictor(left: int, above: int, upper_left: int) -> int:
    estimate = left + above - upper_left
    distance_left = abs(estimate - left)
    distance_above = abs(estimate - above)
    distance_upper_left = abs(estimate - upper_left)
    if distance_left <= distance_above and distance_left <= distance_upper_left:
        return left
    if distance_above <= distance_upper_left:
        return above
    return upper_left


def _remove_isolated_fragments(
    rows: tuple[bytes, ...],
    width: int,
    height: int,
) -> tuple[bytes, ...]:
    """去掉生成图周围与主体不相连的零散像素，不改变主体轮廓。"""
    pending = bytearray(width * height)
    for y, row in enumerate(rows):
        row_offset = y * width
        for x in range(width):
            if row[x * 4 + 3] > 0:
                pending[row_offset + x] = 1

    keep = bytearray(width * height)
    for start, state in enumerate(pending):
        if state != 1:
            continue
        pending[start] = 2
        component = [start]
        cursor = 0
        while cursor < len(component):
            position = component[cursor]
            cursor += 1
            y, x = divmod(position, width)
            neighbours = (
                position - 1 if x > 0 else -1,
                position + 1 if x + 1 < width else -1,
                position - width if y > 0 else -1,
                position + width if y + 1 < height else -1,
            )
            for neighbour in neighbours:
                if neighbour >= 0 and pending[neighbour] == 1:
                    pending[neighbour] = 2
                    component.append(neighbour)
        if len(component) >= MIN_SOURCE_COMPONENT_PIXELS:
            for position in component:
                keep[position] = 1

    cleaned_rows: list[bytes] = []
    for y, row in enumerate(rows):
        cleaned = bytearray(row)
        row_offset = y * width
        for x in range(width):
            if not keep[row_offset + x]:
                pixel_offset = x * 4
                cleaned[pixel_offset : pixel_offset + 4] = b"\x00\x00\x00\x00"
        cleaned_rows.append(bytes(cleaned))
    return tuple(cleaned_rows)


@lru_cache(maxsize=1)
def _load_source_icon() -> tuple[int, int, tuple[bytes, ...]]:
    """读取固定的 8 位 RGBA、非隔行 PNG 品牌源图。"""
    data = SOURCE_ICON_FILE.read_bytes()
    digest = hashlib.sha256(data).hexdigest()
    if digest != SOURCE_ICON_SHA256:
        raise ValueError(f"品牌源图校验失败：{SOURCE_ICON_FILE}")
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValueError(f"品牌源图不是有效 PNG：{SOURCE_ICON_FILE}")

    width = height = 0
    idat_parts: list[bytes] = []
    offset = 8
    while offset + 12 <= len(data):
        payload_size = struct.unpack(">I", data[offset : offset + 4])[0]
        chunk_type = data[offset + 4 : offset + 8]
        payload_start = offset + 8
        payload_end = payload_start + payload_size
        crc_end = payload_end + 4
        if crc_end > len(data):
            raise ValueError("品牌源图 PNG 数据不完整")
        payload = data[payload_start:payload_end]
        expected_crc = struct.unpack(">I", data[payload_end:crc_end])[0]
        if zlib.crc32(chunk_type + payload) & 0xFFFFFFFF != expected_crc:
            raise ValueError("品牌源图 PNG 分块校验失败")
        if chunk_type == b"IHDR":
            if len(payload) != 13:
                raise ValueError("品牌源图 PNG 头长度异常")
            width, height, bit_depth, color_type, compression, filter_method, interlace = struct.unpack(
                ">IIBBBBB", payload
            )
            if (bit_depth, color_type, compression, filter_method, interlace) != (8, 6, 0, 0, 0):
                raise ValueError("品牌源图必须是 8 位 RGBA 非隔行 PNG")
        elif chunk_type == b"IDAT":
            idat_parts.append(payload)
        elif chunk_type == b"IEND":
            break
        offset = crc_end

    if width <= 0 or height <= 0 or not idat_parts:
        raise ValueError("品牌源图 PNG 缺少尺寸或像素数据")
    inflated = zlib.decompress(b"".join(idat_parts))
    row_width = width * 4
    expected_size = (row_width + 1) * height
    if len(inflated) != expected_size:
        raise ValueError("品牌源图 PNG 解压尺寸异常")

    rows: list[bytes] = []
    previous = bytes(row_width)
    cursor = 0
    for _row_index in range(height):
        filter_type = inflated[cursor]
        cursor += 1
        encoded_row = inflated[cursor : cursor + row_width]
        cursor += row_width
        decoded = bytearray(row_width)
        for index, value in enumerate(encoded_row):
            left = decoded[index - 4] if index >= 4 else 0
            above = previous[index]
            upper_left = previous[index - 4] if index >= 4 else 0
            if filter_type == 0:
                predictor = 0
            elif filter_type == 1:
                predictor = left
            elif filter_type == 2:
                predictor = above
            elif filter_type == 3:
                predictor = (left + above) // 2
            elif filter_type == 4:
                predictor = _paeth_predictor(left, above, upper_left)
            else:
                raise ValueError(f"品牌源图 PNG 使用了未知过滤器：{filter_type}")
            decoded[index] = (value + predictor) & 0xFF
        previous = bytes(decoded)
        rows.append(previous)
    return width, height, _remove_isolated_fragments(tuple(rows), width, height)


def render_icon(size: int) -> list[bytes]:
    """以预乘 Alpha 超采样缩放 F1 源图，返回 RGBA 像素行。"""
    if size <= 0:
        raise ValueError("图标尺寸必须大于 0")
    source_width, source_height, source_rows = _load_source_icon()
    scale_x = source_width / size
    scale_y = source_height / size
    # 输出越小，单个像素覆盖的源图区域越大；按面积比例增加采样，
    # 但封顶 6×6，兼顾 16px 边缘质量和打包速度。
    sample_grid = max(1, min(6, int(math.ceil(math.sqrt(max(scale_x, scale_y))))))
    sample_count = sample_grid * sample_grid
    rows: list[bytes] = []

    for output_y in range(size):
        output_row = bytearray(size * 4)
        for output_x in range(size):
            accumulated_r = accumulated_g = accumulated_b = accumulated_alpha = 0.0
            for sample_y in range(sample_grid):
                source_y = (output_y + (sample_y + 0.5) / sample_grid) * scale_y - 0.5
                source_y = max(0.0, min(source_y, source_height - 1.0))
                y0 = int(source_y)
                y1 = min(y0 + 1, source_height - 1)
                fy = source_y - y0
                wy0, wy1 = 1.0 - fy, fy
                row0, row1 = source_rows[y0], source_rows[y1]

                for sample_x in range(sample_grid):
                    source_x = (output_x + (sample_x + 0.5) / sample_grid) * scale_x - 0.5
                    source_x = max(0.0, min(source_x, source_width - 1.0))
                    x0 = int(source_x)
                    x1 = min(x0 + 1, source_width - 1)
                    fx = source_x - x0
                    wx0, wx1 = 1.0 - fx, fx
                    index0, index1 = x0 * 4, x1 * 4

                    weights = (wy0 * wx0, wy0 * wx1, wy1 * wx0, wy1 * wx1)
                    pixels = (
                        (row0, index0),
                        (row0, index1),
                        (row1, index0),
                        (row1, index1),
                    )
                    for weight, (pixel_row, pixel_index) in zip(weights, pixels):
                        alpha = pixel_row[pixel_index + 3]
                        accumulated_alpha += alpha * weight
                        accumulated_r += pixel_row[pixel_index] * alpha * weight
                        accumulated_g += pixel_row[pixel_index + 1] * alpha * weight
                        accumulated_b += pixel_row[pixel_index + 2] * alpha * weight

            output_index = output_x * 4
            if accumulated_alpha <= 0.0:
                continue
            output_row[output_index] = max(0, min(255, int(round(accumulated_r / accumulated_alpha))))
            output_row[output_index + 1] = max(0, min(255, int(round(accumulated_g / accumulated_alpha))))
            output_row[output_index + 2] = max(0, min(255, int(round(accumulated_b / accumulated_alpha))))
            output_row[output_index + 3] = max(
                0,
                min(255, int(round(accumulated_alpha / sample_count))),
            )
        rows.append(bytes(output_row))
    return rows


def encode_png(rows: list[bytes]) -> bytes:
    size = len(rows)
    if size <= 0 or any(len(row) != size * 4 for row in rows):
        raise ValueError("RGBA 像素行必须组成正方形图标")

    def chunk(tag: bytes, payload: bytes) -> bytes:
        data = tag + payload
        return struct.pack(">I", len(payload)) + data + struct.pack(">I", zlib.crc32(data))

    raw = b"".join(b"\x00" + row for row in rows)
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


def _encode_ico_bmp(rows: list[bytes]) -> bytes:
    """32 位 BGRA DIB（老版本 Windows 对小尺寸 PNG 条目兼容性差）。"""
    size = len(rows)
    header = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0, size * size * 4, 0, 0, 0, 0)
    xor_data = bytearray()
    for row in reversed(rows):
        for index in range(0, len(row), 4):
            xor_data.extend((row[index + 2], row[index + 1], row[index], row[index + 3]))
    and_stride = ((size + 31) // 32) * 4
    and_mask = b"\x00" * (and_stride * size)
    return header + bytes(xor_data) + and_mask


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
    encoded_icons: dict[int, str] = {}

    def encoded_icon(size: int) -> str:
        if size not in encoded_icons:
            encoded_icons[size] = base64.b64encode(encode_png(render_icon(size))).decode("ascii")
        return encoded_icons[size]

    for size in RUNTIME_PNG_SIZES:
        lines.append(f'    {size}: "{encoded_icon(size)}",')
    lines.append("}")
    lines.append("")
    lines.append("BRAND_MARK_PNGS_BASE64 = {")
    for size in BRAND_MARK_PNG_SIZES:
        lines.append(f'    {size}: "{encoded_icon(size)}",')
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
