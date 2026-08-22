from __future__ import annotations

import base64
import hashlib
import struct
import unittest

from hr_toolkit._icon_data import APP_ICON_PNGS_BASE64, BRAND_MARK_PNGS_BASE64


class AppIconDataTests(unittest.TestCase):
    @staticmethod
    def _png_size(data: bytes) -> tuple[int, int]:
        if not data.startswith(b"\x89PNG\r\n\x1a\n"):
            raise AssertionError("数据不是 PNG")
        return struct.unpack(">II", data[16:24])

    def test_runtime_icons_are_valid_pngs_with_expected_sizes(self) -> None:
        self.assertEqual(sorted(APP_ICON_PNGS_BASE64), [16, 32, 64, 128, 256, 512])
        # macOS Dock 只用 iconphoto 的第一张图，数据必须按从大到小排列
        self.assertEqual(list(APP_ICON_PNGS_BASE64), sorted(APP_ICON_PNGS_BASE64, reverse=True))
        for size, encoded in APP_ICON_PNGS_BASE64.items():
            data = base64.b64decode(encoded)
            self.assertEqual(self._png_size(data), (size, size))

    def test_brand_mark_has_native_dpi_sizes(self) -> None:
        from scripts.generate_app_icons import BRAND_MARK_PNG_SIZES

        self.assertEqual(tuple(BRAND_MARK_PNGS_BASE64), BRAND_MARK_PNG_SIZES)
        for size, encoded in BRAND_MARK_PNGS_BASE64.items():
            self.assertEqual(self._png_size(base64.b64decode(encoded)), (size, size))

    def test_selected_f1_source_and_generated_data_are_in_sync(self) -> None:
        from scripts.generate_app_icons import (
            SOURCE_ICON_FILE,
            SOURCE_ICON_SHA256,
            encode_png,
            render_icon,
        )

        source = SOURCE_ICON_FILE.read_bytes()
        self.assertEqual(hashlib.sha256(source).hexdigest(), SOURCE_ICON_SHA256)
        self.assertEqual(self._png_size(source), (1254, 1254))
        generated_64 = encode_png(render_icon(64))
        self.assertEqual(generated_64, base64.b64decode(APP_ICON_PNGS_BASE64[64]))

        rows = render_icon(64)
        center_alpha = rows[32][32 * 4 + 3]
        self.assertLess(center_alpha, 16, "F1 标识中心应保持镂空")
        self.assertTrue(any(row[index + 3] >= 240 for row in rows for index in range(0, len(row), 4)))

    def test_windows_ico_structure(self) -> None:
        from scripts.generate_app_icons import ICO_BMP_SIZES, ICO_PNG_SIZES, ICO_FILE

        data = ICO_FILE.read_bytes()
        reserved, image_type, count = struct.unpack("<HHH", data[:6])
        self.assertEqual((reserved, image_type), (0, 1))
        self.assertEqual(count, len(ICO_BMP_SIZES) + len(ICO_PNG_SIZES))
        sizes = []
        for index in range(count):
            entry = data[6 + index * 16 : 6 + (index + 1) * 16]
            width = entry[0] or 256
            height = entry[1] or 256
            self.assertEqual(width, height)
            sizes.append(width)
            payload_size, offset = struct.unpack("<II", entry[8:16])
            self.assertLessEqual(offset + payload_size, len(data), "ICO 条目越界")
        self.assertEqual(sizes, [*ICO_BMP_SIZES, *ICO_PNG_SIZES])


if __name__ == "__main__":
    unittest.main()
