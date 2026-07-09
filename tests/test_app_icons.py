from __future__ import annotations

import base64
import struct
import unittest

from hr_toolkit._icon_data import APP_ICON_PNGS_BASE64


class AppIconDataTests(unittest.TestCase):
    def test_runtime_icons_are_valid_pngs_with_expected_sizes(self) -> None:
        self.assertEqual(sorted(APP_ICON_PNGS_BASE64), [16, 32, 64, 128, 256, 512])
        # macOS Dock 只用 iconphoto 的第一张图，数据必须按从大到小排列
        self.assertEqual(list(APP_ICON_PNGS_BASE64), sorted(APP_ICON_PNGS_BASE64, reverse=True))
        for size, encoded in APP_ICON_PNGS_BASE64.items():
            data = base64.b64decode(encoded)
            self.assertTrue(data.startswith(b"\x89PNG\r\n\x1a\n"), f"{size}px 不是 PNG")
            # IHDR 位于固定偏移：8 字节签名 + 4 长度 + 4 类型
            width, height = struct.unpack(">II", data[16:24])
            self.assertEqual((width, height), (size, size))

    def test_windows_ico_structure(self) -> None:
        from scripts.generate_app_icons import ICO_BMP_SIZES, ICO_PNG_SIZES, ICO_FILE

        data = ICO_FILE.read_bytes()
        reserved, image_type, count = struct.unpack("<HHH", data[:6])
        self.assertEqual((reserved, image_type), (0, 1))
        self.assertEqual(count, len(ICO_BMP_SIZES) + len(ICO_PNG_SIZES))
        for index in range(count):
            entry = data[6 + index * 16 : 6 + (index + 1) * 16]
            payload_size, offset = struct.unpack("<II", entry[8:16])
            self.assertLessEqual(offset + payload_size, len(data), "ICO 条目越界")


if __name__ == "__main__":
    unittest.main()
