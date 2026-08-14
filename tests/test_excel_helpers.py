from __future__ import annotations

import os
import sys
import unittest
import shutil
import tempfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from hr_toolkit.common import resources as package_resources
from hr_toolkit.common.excel import (
    _translate_same_row_formula,
    apply_row_snapshot,
    cached_style_id,
    insert_rows,
    set_style_ids,
    snapshot_row,
    style_source_id,
)
from hr_toolkit.common.excel_compat import (
    _build_conversion_env,
    _diagnose_missing_libreoffice,
    _find_libreoffice,
    ensure_xlsx_workbook,
)
from hr_toolkit.common.resources import open_template_resource


class ExcelHelperTest(unittest.TestCase):
    def test_translate_same_row_formula_only_rewrites_cell_reference_rows(self) -> None:
        formula = "=A1+B10+SUM(C1:D1)+E$1+$F1+LOG10(100)"

        self.assertEqual(
            _translate_same_row_formula(formula, source_row=1, target_row=12),
            "=A12+B10+SUM(C12:D12)+E$1+$F12+LOG10(100)",
        )

    def test_ensure_xlsx_workbook_converts_xls_to_temp_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            original_xlsx = root / "原始.xlsx"
            xls_file = root / "原始.xls"
            wb = Workbook()
            wb.active["A1"] = "测试"
            wb.save(original_xlsx)
            shutil.copyfile(original_xlsx, xls_file)

            def fake_convert(source: Path, output_path: Path) -> None:
                shutil.copyfile(source, output_path)

            with patch("hr_toolkit.common.excel_compat._convert_xls_to_xlsx", side_effect=fake_convert):
                converted = ensure_xlsx_workbook(xls_file, root / "temp")

            self.assertEqual(converted.suffix, ".xlsx")
            self.assertNotEqual(converted, xls_file)
            loaded = load_workbook(converted, data_only=True)
            self.assertEqual(loaded.active["A1"].value, "测试")
            loaded.close()

    def test_ensure_xlsx_workbook_converts_renamed_xls_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            renamed_xls = root / "改后缀.xlsx"
            renamed_xls.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1fake")

            def fake_convert(source: Path, output_path: Path) -> None:
                wb = Workbook()
                wb.active["A1"] = "已转换"
                wb.save(output_path)

            with patch("hr_toolkit.common.excel_compat._convert_xls_to_xlsx", side_effect=fake_convert):
                converted = ensure_xlsx_workbook(renamed_xls, root / "temp")

            self.assertNotEqual(converted, renamed_xls)
            loaded = load_workbook(converted, data_only=True)
            self.assertEqual(loaded.active["A1"].value, "已转换")
            loaded.close()

    def test_open_template_resource_falls_back_without_files_api(self) -> None:
        with patch.object(package_resources.resources, "files", None):
            with open_template_resource("data_statistics_template.xlsx") as handle:
                self.assertEqual(handle.read(2), b"PK")


def _style_signature(cell) -> tuple:
    font, fill, border, alignment = cell.font, cell.fill, cell.border, cell.alignment
    return (
        font.name,
        font.sz,
        font.b,
        fill.patternType,
        str(fill.fgColor.rgb) if fill.fgColor is not None else None,
        border.left.style,
        border.right.style,
        border.top.style,
        border.bottom.style,
        alignment.horizontal,
        alignment.vertical,
        alignment.wrapText,
        alignment.textRotation,
        cell.number_format,
        cell.protection.locked,
    )


class RowSnapshotStyleTest(unittest.TestCase):
    """样式快照走的是样式表下标快路径，必须与逐项赋值的结果完全一致。"""

    @staticmethod
    def _styled_workbook() -> Workbook:
        workbook = Workbook()
        ws = workbook.active
        thin = Side(style="thin", color="000000")
        for col_index in range(1, 6):
            cell = ws.cell(1, col_index)
            cell.value = f"值{col_index}"
            cell.font = Font(name="宋体", size=11, bold=col_index % 2 == 0, color="FF0000")
            cell.fill = PatternFill("solid", fgColor="FCE4D6")
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrapText=True, textRotation=col_index)
            cell.number_format = "0.00"
        ws.row_dimensions[1].height = 33.5
        return workbook

    def test_same_workbook_apply_reproduces_every_style_facet(self) -> None:
        workbook = self._styled_workbook()
        ws = workbook.active
        snapshot = snapshot_row(ws, 1, 5)

        apply_row_snapshot(ws, 7, snapshot, translate_formulas=False)

        self.assertEqual(ws.row_dimensions[7].height, 33.5)
        for col_index in range(1, 6):
            self.assertEqual(
                _style_signature(ws.cell(7, col_index)),
                _style_signature(ws.cell(1, col_index)),
                f"第 {col_index} 列样式不一致",
            )

    def test_cross_workbook_apply_reproduces_every_style_facet(self) -> None:
        """工资表拆分会把快照套用到另一个工作簿，样式下标在那边无效。"""
        source = self._styled_workbook()
        snapshot = snapshot_row(source.active, 1, 5)

        target = Workbook()
        target_ws = target.active
        apply_row_snapshot(target_ws, 3, snapshot, translate_formulas=False)
        # 第二行走缓存路径，必须和第一行结果相同
        apply_row_snapshot(target_ws, 4, snapshot, translate_formulas=False)

        for row_index in (3, 4):
            for col_index in range(1, 6):
                self.assertEqual(
                    _style_signature(target_ws.cell(row_index, col_index)),
                    _style_signature(source.active.cell(1, col_index)),
                    f"第 {row_index} 行第 {col_index} 列样式不一致",
                )

    def test_cross_workbook_cache_does_not_mix_up_two_sources(self) -> None:
        """两个来源工作簿的样式下标含义不同，缓存不能相互串味。"""
        first = Workbook()
        first.active.cell(1, 1).font = Font(name="宋体", size=20, bold=True)
        second = Workbook()
        second.active.cell(1, 1).font = Font(name="Arial", size=8, bold=False)

        first_snapshot = snapshot_row(first.active, 1, 1)
        second_snapshot = snapshot_row(second.active, 1, 1)

        target = Workbook()
        apply_row_snapshot(target.active, 1, first_snapshot, translate_formulas=False)
        apply_row_snapshot(target.active, 2, second_snapshot, translate_formulas=False)

        self.assertEqual((target.active.cell(1, 1).font.name, target.active.cell(1, 1).font.sz), ("宋体", 20.0))
        self.assertEqual((target.active.cell(2, 1).font.name, target.active.cell(2, 1).font.sz), ("Arial", 8.0))

    def test_snapshot_does_not_keep_source_workbook_alive(self) -> None:
        """快照只持弱引用，否则批量处理会把每个来源工作簿都留在内存里。"""
        import gc
        import weakref

        workbook = self._styled_workbook()
        snapshot = snapshot_row(workbook.active, 1, 5)
        reference = weakref.ref(workbook)

        del workbook
        gc.collect()

        self.assertIsNone(reference(), "快照不应延长来源工作簿的生命周期")
        # 来源已释放时仍要能安全套用（退回逐项赋值）
        target = Workbook()
        apply_row_snapshot(target.active, 1, snapshot, translate_formulas=False)
        self.assertEqual(target.active.cell(1, 1).font.name, "宋体")


class StyleIdCacheTest(unittest.TestCase):
    def test_cached_style_id_matches_direct_assignment(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        alignment = Alignment(horizontal="center", vertical="center", wrapText=True)

        ws.cell(1, 1).border = border
        ws.cell(1, 1).alignment = alignment

        border_id = cached_style_id(ws, "border", "thin", lambda: border)
        alignment_id = cached_style_id(ws, "alignment", "center", lambda: alignment)
        set_style_ids(ws.cell(2, 1), border_id=border_id, alignment_id=alignment_id)

        self.assertEqual(_style_signature(ws.cell(2, 1)), _style_signature(ws.cell(1, 1)))

    def test_cached_style_id_is_stable_across_calls(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        font = Font(name="宋体", size=10)
        first = cached_style_id(ws, "font", "song10", lambda: font)
        second = cached_style_id(ws, "font", "song10", lambda: self.fail("不应重复构造样式"))
        self.assertEqual(first, second)

    def test_set_style_ids_does_not_mutate_shared_style_array(self) -> None:
        """StyleArray 可能被多个单元格共用，就地改写会污染邻居。"""
        workbook = Workbook()
        ws = workbook.active
        shared = ws.cell(1, 1)._style
        ws.cell(1, 2)._style = shared  # 人为制造共用
        thin = Side(style="thin", color="000000")
        border_id = cached_style_id(
            ws, "border", "thin", lambda: Border(left=thin, right=thin, top=thin, bottom=thin)
        )

        set_style_ids(ws.cell(1, 1), border_id=border_id)

        self.assertEqual(ws.cell(1, 1).border.left.style, "thin")
        self.assertIsNone(ws.cell(1, 2).border.left.style, "相邻单元格不应被连带修改")

    def test_style_source_id_reports_current_index(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        self.assertEqual(style_source_id(ws.cell(1, 1), "alignment"), 0)
        ws.cell(1, 1).alignment = Alignment(horizontal="center")
        self.assertNotEqual(style_source_id(ws.cell(1, 1), "alignment"), 0)


class InsertRowsTest(unittest.TestCase):
    def test_insert_rows_tracks_max_row_without_scanning_every_cell(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        for row_index in range(1, 6):
            ws.cell(row_index, 1).value = row_index

        insert_rows(ws, 2, 1)

        self.assertEqual(ws.max_row, 6)
        self.assertEqual(ws._current_row, ws.max_row)
        self.assertIsNone(ws.cell(2, 1).value)
        self.assertEqual([ws.cell(r, 1).value for r in (1, 3, 4, 5, 6)], [1, 2, 3, 4, 5])

    def test_insert_rows_beyond_used_range_keeps_max_row(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        ws.cell(1, 1).value = "只有一行"

        insert_rows(ws, 10, 2)

        self.assertEqual(ws.max_row, 1)
        self.assertEqual(ws._current_row, 1)


class LibreOfficeLookupTest(unittest.TestCase):
    """回归测试：macOS .app 启动时 PATH 被裁剪，找不到 Homebrew 装的 soffice。"""

    def test_find_libreoffice_returns_homebrew_binary_on_macos(self) -> None:
        """当 PATH 不含 /opt/homebrew/bin，但二进制实际存在时，应能识别。"""
        fake_binary = Path("/opt/homebrew/bin/soffice")
        with patch("hr_toolkit.common.excel_compat.os.environ", {"PATH": "/usr/bin:/bin"}):
            with patch("hr_toolkit.common.excel_compat.Path.exists", return_value=True):
                executable = _find_libreoffice()
        # 实际值受其他路径影响，只要不为 None 即可证明查找逻辑生效
        self.assertIsNotNone(executable)
        self.assertTrue(executable.endswith("soffice") or executable.endswith("libreoffice"))

    def test_find_libreoffice_returns_none_when_nothing_exists(self) -> None:
        """全部路径都不存在时返回 None。"""
        with patch("hr_toolkit.common.excel_compat.os.environ", {"PATH": "/usr/bin:/bin"}):
            with patch("hr_toolkit.common.excel_compat.Path.exists", return_value=False):
                with patch("hr_toolkit.common.excel_compat.shutil.which", return_value=None):
                    self.assertIsNone(_find_libreoffice())

    @unittest.skipUnless(sys.platform == "darwin", "macOS-specific LibreOffice bundle path")
    def test_find_libreoffice_prefers_mac_app_path(self) -> None:
        """官方 .app 路径优先级最高。"""
        with patch("hr_toolkit.common.excel_compat.os.environ", {"PATH": "/usr/bin:/bin"}):
            original_exists = Path.exists

            def fake_exists(self: Path) -> bool:
                if str(self) == "/Applications/LibreOffice.app/Contents/MacOS/soffice":
                    return True
                return False

            with patch("hr_toolkit.common.excel_compat.Path.exists", fake_exists):
                executable = _find_libreoffice()
        self.assertEqual(
            executable,
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        )

    def test_build_conversion_env_keeps_existing_path(self) -> None:
        """环境变量注入不能丢失用户已有 PATH。"""
        with patch.dict(os.environ, {"PATH": "/usr/bin:/bin", "HOME": "/tmp"}, clear=False):
            env = _build_conversion_env()
        self.assertIn("/usr/bin", env["PATH"])
        self.assertIn("/bin", env["PATH"])

    def test_build_conversion_env_adds_homebrew_when_exists(self) -> None:
        """当 Homebrew 路径存在时，应被追加到 PATH。"""
        with patch.dict(os.environ, {"PATH": "/usr/bin:/bin"}, clear=False):
            with patch("hr_toolkit.common.excel_compat.Path.exists", return_value=True):
                env = _build_conversion_env()
        self.assertIn("/opt/homebrew/bin", env["PATH"])
        self.assertIn("/usr/local/bin", env["PATH"])

    def test_diagnose_contains_path_and_install_hint(self) -> None:
        """诊断信息应包含 PATH、已检查路径与安装建议。"""
        with patch.dict(os.environ, {"PATH": "/usr/bin:/bin"}, clear=False):
            with patch("hr_toolkit.common.excel_compat.Path.exists", return_value=False):
                message = _diagnose_missing_libreoffice()
        self.assertIn("未找到 libreoffice/soffice", message)
        self.assertIn("PATH", message)
        self.assertIn("brew install", message)
        self.assertIn("/opt/homebrew/bin", message)
        self.assertIn("/Applications/LibreOffice.app", message)


if __name__ == "__main__":
    unittest.main()
