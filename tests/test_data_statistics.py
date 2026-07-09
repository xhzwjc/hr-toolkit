from __future__ import annotations

import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from hr_toolkit.tools.data_statistics import (
    generate_data_statistics_reports,
    parse_report_date,
    resolve_week_range,
)


class DataStatisticsTest(unittest.TestCase):
    def test_generate_data_statistics_reports(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            output_dir = root / "output"
            input_dir.mkdir()
            _write_attendance_file(input_dir / "考勤结果.xlsx")
            _write_weekly_file(input_dir / "【汇报】唐人周报04.01-05.04.xlsx")
            _write_monthly_file(input_dir / "【汇报】唐人月报04.01-05.04.xlsx")

            result = generate_data_statistics_reports(input_dir, output_dir)

            self.assertEqual(result.attendance_source_count, 4)
            self.assertEqual(result.attendance_person_count, 1)
            self.assertEqual(result.attendance_exception_count, 5)
            self.assertEqual(result.weekly_record_count, 8)
            self.assertEqual(result.monthly_record_count, 2)
            self.assertEqual(result.report_person_count, 2)
            self.assertEqual(result.report_exception_count, 2)
            self.assertTrue(result.output_file and result.output_file.exists())

            wb = load_workbook(result.output_file, data_only=False)
            attendance = wb["考勤统计"]
            self.assertEqual(attendance.cell(3, 2).value, "总部")
            self.assertEqual(attendance.cell(3, 3).value, "运营部")
            self.assertEqual(attendance.cell(3, 4).value, "王小丽")
            self.assertEqual(attendance.cell(3, 8).value, 0.5)
            self.assertEqual(attendance.cell(3, 12).value, 0.5)
            self.assertEqual(attendance.cell(3, 15).value, 3)
            self.assertIn("4.10上班未打卡", attendance.cell(3, 17).value)
            self.assertIn("4.15晚上加班0.5天", attendance.cell(3, 17).value)

            report = wb["周月报统计"]
            self.assertEqual(report.max_column, 10)
            rows = {report.cell(row, 4).value: [report.cell(row, col).value for col in range(1, 11)] for row in (3, 4)}
            self.assertEqual(rows["黄五"][7:10], [1, None, "月报超时（17:31提交）"])
            self.assertEqual(rows["黄三"][5:10], [1, None, None, None, "第四周周报超时（18:37提交）"])
            self.assertEqual(report.cell(5, 2).value, "总计（周报截止时间2026.5.4 17:00）")
            self.assertIn("审批", report.cell(7, 1).value)
            self.assertEqual(report.cell(9, 3).value, "汇报规则：")
            self.assertIn("2026.4.13、2026.4.20、2026.4.27、2026.5.4", report.cell(9, 4).value)
            self.assertEqual(report.cell(10, 3).value, "月报规则：")
            self.assertIn("2026.5.2 17:00", report.cell(10, 4).value)

            detail = wb["周月报异常明细"]
            self.assertEqual(detail.cell(2, 5).value, "2026年4月")
            self.assertEqual(detail.cell(3, 5).value, "第四周")
            self.assertIsNone(detail.cell(2, 9).value)
            wb.close()

    def test_report_staff_list_counts_missing_people(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            output_dir = root / "output"
            staff_file = root / "应汇报人员名单.xlsx"
            input_dir.mkdir()
            _write_weekly_file(input_dir / "【汇报】唐人周报04.01-05.04.xlsx")
            _write_monthly_file(input_dir / "【汇报】唐人月报04.01-05.04.xlsx")
            _write_staff_file(staff_file)

            result = generate_data_statistics_reports(input_dir, output_dir, report_staff_path=staff_file)

            self.assertEqual(result.expected_reporter_count, 3)
            self.assertEqual(result.report_person_count, 3)
            self.assertEqual(result.report_exception_count, 7)

            wb = load_workbook(result.output_file, data_only=True)
            report = wb["周月报统计"]
            rows = {report.cell(row, 4).value: [report.cell(row, col).value for col in range(1, 11)] for row in range(3, 6)}
            self.assertEqual(rows["黄六"][4:10], [4, None, 1, None, None, "第二周未写周报；第三周未写周报；第四周未写周报；第五周未写周报；未写月报"])
            self.assertEqual(rows["黄六"][1], "总部")
            self.assertEqual(rows["黄六"][2], "财务部")
            wb.close()

    def test_report_deadline_allows_170059(self) -> None:
        self.assertFalse(
            _generate_deadline_case(datetime(2026, 5, 2, 17, 0, 59)).report_exception_count,
        )
        self.assertEqual(
            _generate_deadline_case(datetime(2026, 5, 2, 17, 1, 0)).report_exception_count,
            1,
        )

    def test_weekly_deadline_allows_170059(self) -> None:
        self.assertFalse(
            _generate_weekly_deadline_case(datetime(2026, 4, 13, 17, 0, 59)).report_exception_count,
        )
        self.assertEqual(
            _generate_weekly_deadline_case(datetime(2026, 4, 13, 17, 1, 0)).report_exception_count,
            1,
        )

    def test_week_range_excludes_previous_month_monday(self) -> None:
        # 2026年6月1日是周一：不选日期时，6.1 会作为第一个周报截止日，
        # 统计到 5 月最后一周的周报；选择 6.2-6.30 后不再统计。
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            input_dir.mkdir()
            _write_june_weekly_file(input_dir / "【汇报】唐人周报06.01-06.30.xlsx")

            # 无日期范围：6.1 的周报超时 + 未写月报，共 2 条异常
            default_result = generate_data_statistics_reports(input_dir, root / "out1", dry_run=True)
            self.assertEqual(default_result.report_exception_count, 2)

            # 选择 6.2-6.30：只剩未写月报 1 条，5 月最后一周不再统计
            ranged_result = generate_data_statistics_reports(
                input_dir,
                root / "out2",
                week_start="2026-06-02",
                week_end="2026-06-30",
            )
            self.assertEqual(ranged_result.report_exception_count, 1)
            self.assertEqual(ranged_result.week_range_start.isoformat(), "2026-06-02")

            wb = load_workbook(ranged_result.output_file, data_only=True)
            report = wb["周月报统计"]
            self.assertIn("2026.6.8、2026.6.15、2026.6.22、2026.6.29", report.cell(9, 4).value)
            self.assertNotIn("2026.6.1、", report.cell(9, 4).value)
            wb.close()

    def test_week_range_late_annotation_with_cross_day_time(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            input_dir.mkdir()
            _write_report_boundary_file(
                input_dir / "【汇报】唐人周报04.13-04.13.xlsx",
                datetime(2026, 4, 14, 9, 5),
                "周报",
            )
            output_result = generate_data_statistics_reports(input_dir, root / "out2")
            wb = load_workbook(output_result.output_file, data_only=True)
            report = wb["周月报统计"]
            remarks = report.cell(3, 10).value
            self.assertIn("周报超时（4月14日9:05提交）", remarks)
            wb.close()

    def test_midweek_makeup_counts_as_previous_week_late(self) -> None:
        # 周三补交：算上一期超时（写明日期时间），下一周没交照样记未写
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            input_dir.mkdir()
            _write_weekly_rows_file(
                input_dir / "【汇报】唐人周报06.01-06.30.xlsx",
                [
                    datetime(2026, 6, 8, 15, 0),  # 6.8 周一按时
                    datetime(2026, 6, 17, 9, 5),  # 6.15 那期的周三补交
                ],
            )
            result = generate_data_statistics_reports(
                input_dir,
                root / "out",
                week_start="2026-06-02",
                week_end="2026-06-30",
            )
            wb = load_workbook(result.output_file, data_only=True)
            remarks = wb["周月报统计"].cell(3, 10).value
            wb.close()
            self.assertIn("第二周周报超时（6月17日9:05提交）", remarks)
            self.assertIn("第三周未写周报", remarks)
            self.assertIn("第四周未写周报", remarks)

    def test_early_submission_rolls_to_next_week_when_already_reported(self) -> None:
        # 上周一已按时交过，周四又交一份（提前交下期）：算下一期，不记未写
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            input_dir.mkdir()
            _write_weekly_rows_file(
                input_dir / "【汇报】唐人周报06.15-06.28.xlsx",
                [
                    datetime(2026, 6, 15, 15, 0),  # 6.15 周一按时
                    datetime(2026, 6, 18, 16, 0),  # 周四提前交 6.22 那期
                ],
            )
            result = generate_data_statistics_reports(
                input_dir,
                root / "out",
                week_start="2026-06-15",
                week_end="2026-06-28",
                dry_run=True,
            )
            # 只剩未写月报，两个周一都算已交
            self.assertEqual(result.report_exception_count, 1)

    def test_friday_submission_counts_for_next_monday(self) -> None:
        # 几维周末双休：周五下班交的周报算下周一截止那期，按时
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            input_dir.mkdir()
            _write_weekly_rows_file(
                input_dir / "【汇报】几维周报06.16-06.28.xlsx",
                [datetime(2026, 6, 19, 18, 0)],  # 周五
            )
            result = generate_data_statistics_reports(
                input_dir,
                root / "out",
                week_start="2026-06-16",
                week_end="2026-06-28",
                dry_run=True,
            )
            self.assertEqual(result.report_exception_count, 1)  # 仅未写月报

    def test_submission_for_period_outside_range_is_ignored(self) -> None:
        # 6.26（周五）交的属于 6.29 那期；范围只到 6.24 时不参与本次统计，
        # 也不能被强行算到 6.22 头上记超时
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            input_dir = root / "input"
            input_dir.mkdir()
            _write_weekly_rows_file(
                input_dir / "【汇报】唐人周报06.02-06.24.xlsx",
                [
                    datetime(2026, 6, 8, 15, 0),
                    datetime(2026, 6, 26, 10, 0),
                ],
            )
            result = generate_data_statistics_reports(
                input_dir,
                root / "out",
                week_start="2026-06-02",
                week_end="2026-06-24",
            )
            wb = load_workbook(result.output_file, data_only=True)
            remarks = wb["周月报统计"].cell(3, 10).value
            wb.close()
            self.assertNotIn("超时", remarks.replace("月报超时", ""))
            self.assertIn("第二周未写周报", remarks)
            self.assertIn("第三周未写周报", remarks)

    def test_week_range_requires_both_dates(self) -> None:
        with self.assertRaises(ValueError):
            resolve_week_range("2026-06-02", None)
        with self.assertRaises(ValueError):
            resolve_week_range(None, "2026-06-30")
        with self.assertRaises(ValueError):
            resolve_week_range("2026-06-30", "2026-06-02")
        self.assertIsNone(resolve_week_range(None, None))

    def test_parse_report_date_formats(self) -> None:
        from datetime import date

        expected = date(2026, 6, 2)
        for text in ("2026-06-02", "2026-6-2", "2026/6/2", "2026.6.2", "2026年6月2日"):
            self.assertEqual(parse_report_date(text), expected, text)
        with self.assertRaises(ValueError):
            parse_report_date("6.2")
        with self.assertRaises(ValueError):
            parse_report_date("2026-13-01")

    def test_generate_from_zip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "考勤结果.xlsx"
            archive = root / "数据统计.zip"
            output_dir = root / "output"
            _write_attendance_file(source)
            with zipfile.ZipFile(archive, "w") as zip_file:
                zip_file.write(source, arcname=source.name)

            result = generate_data_statistics_reports([archive], output_dir)

            self.assertEqual(result.attendance_source_count, 4)
            self.assertEqual(result.attendance_person_count, 1)
            self.assertTrue(result.output_file and result.output_file.exists())


def _write_attendance_file(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "日结果"
    headers = [
        "姓名",
        "部门名称",
        "确认状态",
        "日期",
        "是否异常处理",
        "漏打卡次数",
        "应出勤小时数",
        "实出勤小时数",
        "迟到次数",
        "迟到分钟数",
        "早退次数",
        "早退分钟数",
        "旷工天数",
        "旷工次数",
        "外出",
        "工作日出差",
        "休息日出差天数",
        "事假",
        "病假天数",
        "婚假",
        "产假天数",
        "陪护假",
        "丧假",
        "探亲假",
        "工伤",
        "年假天数",
        "调休",
        "加班计调休时长",
        "计划上下班时间",
        "当日刷卡记录",
        "缺卡记录",
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col).value = header
    rows = [
        ["王小丽", "运营部", "否", datetime(2026, 4, 10), "否", 1, 7, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "08:30,12:00|14:00,17:30", "08:30,17:51", None],
        ["王小丽", "运营部", "否", datetime(2026, 4, 15), "否", 1, 7, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3.5, "08:30~12:00|14:00~17:30", "08:19,17:32,17:54,21:00", None],
        ["王小丽", "运营部", "否", datetime(2026, 4, 18), "否", 0, 3.5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3.5, 0, "08:30~12:00", None, None],
        ["王小丽", "运营部", "否", datetime(2026, 4, 30), "否", 1, 7, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "08:30,12:00|14:00,17:30", "08:23,17:30", None],
    ]
    for row_index, row in enumerate(rows, start=2):
        for col_index, value in enumerate(row, start=1):
            ws.cell(row_index, col_index).value = value
    wb.save(path)
    wb.close()


def _write_weekly_file(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "工作表1"
    headers = ["汇报编号", "汇报时间", "汇报人", "汇报人部门", "汇报对象"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col).value = header
    rows = [
        ["1", datetime(2026, 4, 11, 15, 0), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
        ["2", datetime(2026, 4, 11, 15, 0), "黄五", "唐人数智科技股份有限公司/唐人数智/行政人事中心/办公室", "罗一一"],
        ["3", datetime(2026, 4, 18, 15, 0), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
        ["4", datetime(2026, 4, 18, 15, 0), "黄五", "唐人数智科技股份有限公司/唐人数智/行政人事中心/办公室", "罗一一"],
        ["5", datetime(2026, 4, 27, 18, 37), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
        ["6", datetime(2026, 4, 25, 15, 0), "黄五", "唐人数智科技股份有限公司/唐人数智/行政人事中心/办公室", "罗一一"],
        ["7", datetime(2026, 5, 2, 11, 0), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
        ["8", datetime(2026, 5, 2, 11, 0), "黄五", "唐人数智科技股份有限公司/唐人数智/行政人事中心/办公室", "罗一一"],
    ]
    for row_index, row in enumerate(rows, start=2):
        for col_index, value in enumerate(row, start=1):
            ws.cell(row_index, col_index).value = value
    wb.save(path)
    wb.close()


def _write_monthly_file(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "工作表1"
    headers = ["汇报编号", "汇报时间", "汇报人", "汇报人部门", "汇报对象", "评论"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col).value = header
    rows = [
        ["1", datetime(2026, 4, 30, 11, 30), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一", None],
        ["2", datetime(2026, 5, 2, 17, 31), "黄五", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一", None],
    ]
    for row_index, row in enumerate(rows, start=2):
        for col_index, value in enumerate(row, start=1):
            ws.cell(row_index, col_index).value = value
    wb.save(path)
    wb.close()


def _write_weekly_rows_file(path: Path, report_times: list[datetime], name: str = "黄三") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "工作表1"
    headers = ["汇报编号", "汇报时间", "汇报人", "汇报人部门", "汇报对象"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col).value = header
    for row_index, report_time in enumerate(report_times, start=2):
        row = [str(row_index - 1), report_time, name, "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"]
        for col_index, value in enumerate(row, start=1):
            ws.cell(row_index, col_index).value = value
    wb.save(path)
    wb.close()


def _write_june_weekly_file(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "工作表1"
    headers = ["汇报编号", "汇报时间", "汇报人", "汇报人部门", "汇报对象"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col).value = header
    rows = [
        # 6.1（周一）提交的是 5 月最后一周的周报，且超时
        ["1", datetime(2026, 6, 1, 18, 0), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
        ["2", datetime(2026, 6, 8, 15, 0), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
        ["3", datetime(2026, 6, 15, 15, 0), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
        ["4", datetime(2026, 6, 22, 15, 0), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
        ["5", datetime(2026, 6, 29, 15, 0), "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"],
    ]
    for row_index, row in enumerate(rows, start=2):
        for col_index, value in enumerate(row, start=1):
            ws.cell(row_index, col_index).value = value
    wb.save(path)
    wb.close()


def _write_staff_file(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "应汇报人员"
    headers = ["姓名", "部门（片区）"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col).value = header
    rows = [
        ["黄三", "装备事业部"],
        ["黄五", "行政人事中心/办公室"],
        ["黄六", "财务部"],
    ]
    for row_index, row in enumerate(rows, start=2):
        for col_index, value in enumerate(row, start=1):
            ws.cell(row_index, col_index).value = value
    wb.save(path)
    wb.close()


def _generate_deadline_case(monthly_time: datetime):
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        input_dir = root / "input"
        output_dir = root / "output"
        input_dir.mkdir()
        wb = Workbook()
        ws = wb.active
        ws.title = "工作表1"
        headers = ["汇报编号", "汇报时间", "汇报人", "汇报人部门", "汇报对象"]
        for col, header in enumerate(headers, start=1):
            ws.cell(1, col).value = header
        row = ["1", monthly_time, "黄三", "唐人数智科技股份有限公司/唐人数智/装备事业部", "罗一一"]
        for col_index, value in enumerate(row, start=1):
            ws.cell(2, col_index).value = value
        wb.save(input_dir / "【汇报】唐人月报04.01-05.04.xlsx")
        wb.close()
        return generate_data_statistics_reports(input_dir, output_dir)


def _generate_weekly_deadline_case(weekly_time: datetime):
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        input_dir = root / "input"
        output_dir = root / "output"
        input_dir.mkdir()
        _write_report_boundary_file(
            input_dir / "【汇报】唐人周报04.13-04.13.xlsx",
            weekly_time,
            "周报",
        )
        _write_report_boundary_file(
            input_dir / "【汇报】唐人月报04.01-05.04.xlsx",
            datetime(2026, 4, 30, 11, 0, 0),
            "月报",
        )
        return generate_data_statistics_reports(input_dir, output_dir)


def _write_report_boundary_file(path: Path, report_time: datetime, report_kind: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "工作表1"
    headers = ["汇报编号", "汇报时间", "汇报人", "汇报人部门", "汇报对象"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col).value = header
    row = [
        f"{report_kind}1",
        report_time,
        "黄三",
        "唐人数智科技股份有限公司/唐人数智/装备事业部",
        "罗一一",
    ]
    for col_index, value in enumerate(row, start=1):
        ws.cell(2, col_index).value = value
    wb.save(path)
    wb.close()


if __name__ == "__main__":
    unittest.main()
