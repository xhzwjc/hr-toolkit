from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from hr_toolkit.tools.archive_import import import_archive_transfers


class ArchiveImportTest(unittest.TestCase):
    def test_import_archive_transfers_by_company_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            output_dir = root / "output"
            input_dir.mkdir()
            target = root / "档案汇总.xlsx"
            _write_transfer_file(input_dir / "茂名项目部人事档案移交表.xlsx")
            _write_summary_file(target)
            _write_summary_file(input_dir / "档案汇总.xlsx")

            result = import_archive_transfers(input_dir, target, output_dir)

            self.assertEqual(result.source_record_count, 3)
            self.assertEqual(result.inserted_count, 2)
            self.assertEqual(result.updated_count, 1)
            self.assertTrue(any("不是档案移交表" in warning for warning in result.warnings))
            self.assertTrue(result.output_file and result.output_file.exists())

            wb = load_workbook(result.output_file, data_only=False)
            ws1 = wb["公司1"]
            self.assertEqual(ws1.cell(5, 2).value, "张三")
            self.assertEqual(ws1.cell(5, 1).value, "11")
            self.assertEqual(ws1.cell(5, 4).value, '=MIDB(C5,7,4)&"-"&MIDB(C5,11,2)&"-"&MIDB(C5,13,2)')
            self.assertEqual(ws1.cell(5, 9).value, '=A5&"-"&TEXT(G5,"00000000")&"-"&TEXT(J5,"00")&"-"&H5')
            self.assertEqual(ws1.cell(5, 12).value, "√")
            self.assertEqual(ws1.cell(5, 19).value, 4)
            self.assertIn("驾照复印件", ws1.cell(5, 30).value)
            self.assertIn("解除合同协议书", ws1.cell(5, 30).value)

            ws2 = wb["公司2"]
            self.assertEqual(ws2.cell(4, 2).value, "已存在")
            self.assertEqual(ws2.cell(4, 12).value, "√")
            self.assertEqual(ws2.cell(4, 19).value, 4)

    def test_dry_run_does_not_write_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_file = root / "茂名项目部人事档案移交表.xlsx"
            target = root / "档案汇总.xlsx"
            output_dir = root / "output"
            _write_transfer_file(input_file)
            _write_summary_file(target)

            result = import_archive_transfers(input_file, target, output_dir, dry_run=True)

            self.assertEqual(result.source_record_count, 3)
            self.assertEqual(result.inserted_count, 3)
            self.assertIsNone(result.output_file)
            self.assertFalse(output_dir.exists())

    def test_duplicate_new_id_card_in_same_batch_is_skipped(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_file = root / "茂名项目部人事档案移交表.xlsx"
            target = root / "档案汇总.xlsx"
            output_dir = root / "output"
            _write_transfer_file(input_file, duplicate_first=True)
            _write_summary_file(target)

            result = import_archive_transfers(input_file, target, output_dir)

            self.assertEqual(result.source_record_count, 4)
            self.assertEqual(result.inserted_count, 2)
            self.assertEqual(result.updated_count, 1)
            self.assertEqual(result.skipped_count, 1)
            self.assertTrue(any("在本次导入中重复" in warning for warning in result.warnings))

            wb = load_workbook(result.output_file, data_only=False)
            ws = wb["公司1"]
            names = [ws.cell(row, 2).value for row in range(4, 8)]
            self.assertEqual(names.count("张三"), 1)
            wb.close()


def _write_transfer_file(path: Path, *, duplicate_first: bool = False) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "移交表"
    ws["A1"] = "茂名项目部人事档案移交表"
    headers = [
        "公司",
        "姓名",
        "身份证",
        "出生日期",
        "年齡",
        "入职时间",
        "入职公式",
        "出生年月公式",
        "电子照片",
        "入职登记表",
        "劳动合同",
        "保密协议",
        "入职须知",
        "员工三级安全教育",
        "身份证复印件",
        "银行卡复印件",
        "驾照复印件",
        "解除合同协议书",
        "备注",
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(2, col).value = header
    rows = [
        ["公司1", "张三", "4600271987030XXXXX", None, None, "2026-04-09", None, None, "√", "√", 4, 2, 2, "√", "√", "√", "√", "√", "补充说明"],
        ["公司2", "已存在", "440921198009XXXXXX", None, None, "2026-04-14", None, None, None, "√", 4, None, None, None, None, None, None, None, None],
        ["公司3", "张五", "4409211994103XXXXX", None, None, "2026-03-13", None, None, None, None, None, None, None, None, None, None, None, None, None],
    ]
    if duplicate_first:
        rows.append(["公司1", "张三重复", "4600271987030XXXXX", None, None, "2026-04-09", None, None, None, None, None, None, None, None, None, None, None, None, None])
    for row_index, row in enumerate(rows, start=3):
        for col_index, value in enumerate(row, start=1):
            ws.cell(row_index, col_index).value = value
    wb.save(path)


def _write_summary_file(path: Path) -> None:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "公司1"
    ws2 = wb.create_sheet("公司2")
    for ws in (ws1, ws2):
        ws["A1"] = "人事档案编号表"
        headers = [
            "编号",
            "姓名",
            "身份证",
            "出生日期",
            "年齡",
            "入职时间",
            "入职公式",
            "序号",
            "档案号",
            "出生年月公式",
            "档案柜号",
            "员工入职表",
            "身份证复印件",
            "银行卡复印件",
            "体检报告单",
            "学历证书",
            "学位证书",
            "相关资格证书",
            "劳动合同",
            "照片",
            "离职证明",
            "入职员工须知",
            "员工手册签收单",
            "安全生产责任书",
            "保密协议",
            "竞业协议",
            "三级安全教育登记（登记卡+试卷）",
            "员工健康情况调查表",
            "员工进场记录",
            "其他",
            "员工异动审批表",
            "入职考试试卷",
            "员工转正审批表",
            "转正考试试卷",
            "增购社保申请单",
            "离职申请单",
        ]
        for col, header in enumerate(headers, start=1):
            ws.cell(3, col).value = header
        ws.cell(4, 2).value = "模板行"
        ws.cell(4, 3).value = "000000000000000000"
        ws.cell(5, 1).value = "对应行2的序号，如是抚州项目则标02"
    ws2.cell(4, 2).value = "已存在"
    ws2.cell(4, 3).value = "440921198009XXXXXX"
    ws2.cell(4, 12).value = None
    wb.save(path)


if __name__ == "__main__":
    unittest.main()
