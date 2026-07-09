from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

from hr_toolkit import runlog


class RunLogTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.log_file = Path(self._tmp.name) / "app.log"
        os.environ[runlog.RUN_LOG_ENV] = str(self.log_file)

    def tearDown(self) -> None:
        os.environ.pop(runlog.RUN_LOG_ENV, None)
        self._tmp.cleanup()

    def test_log_line_writes_timestamped_entry(self) -> None:
        runlog.log_line("开始 数据统计：a.xlsx(160KB)")
        content = self.log_file.read_text(encoding="utf-8")
        self.assertRegex(content, r"^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] 开始 数据统计：a\.xlsx\(160KB\)\n$")

    def test_log_exception_includes_traceback(self) -> None:
        try:
            raise ValueError("测试异常")
        except ValueError as exc:
            runlog.log_exception("数据统计 失败", exc)
        content = self.log_file.read_text(encoding="utf-8")
        self.assertIn("数据统计 失败", content)
        self.assertIn("Traceback", content)
        self.assertIn("ValueError: 测试异常", content)

    def test_log_is_trimmed_when_oversized(self) -> None:
        self.log_file.write_bytes(b"[old] line\n" * 200_000)  # ~2 MB
        runlog.log_line("新的一行")
        data = self.log_file.read_bytes()
        self.assertLess(len(data), 512 * 1024)
        self.assertTrue(data.startswith(b"(...earlier log trimmed...)\n"))
        self.assertIn("新的一行".encode("utf-8"), data)

    def test_describe_value_summarizes_without_content(self) -> None:
        excel = Path(self._tmp.name) / "几维6月考勤.xlsx"
        excel.write_bytes(b"x" * 2048)
        self.assertEqual(runlog.describe_value(excel), "几维6月考勤.xlsx(2KB)")
        self.assertEqual(runlog.describe_value(None), "无")
        many = [excel] * 7
        summary = runlog.describe_value(many)
        self.assertIn("等共7项", summary)
        self.assertEqual(summary.count("几维6月考勤"), 5)

    def test_describe_call_combines_args_and_kwargs(self) -> None:
        excel = Path(self._tmp.name) / "输入.xlsx"
        excel.write_bytes(b"x")
        text = runlog.describe_call((excel,), {"dry_run": True, "staff": None})
        self.assertIn("输入.xlsx", text)
        self.assertIn("dry_run=True", text)
        self.assertNotIn("staff", text)  # None 参数不记录

    def test_log_failure_is_silent(self) -> None:
        os.environ[runlog.RUN_LOG_ENV] = "/nonexistent-root/no-way/app.log"
        runlog.log_line("不应抛出异常")  # 只要不抛异常即通过


if __name__ == "__main__":
    unittest.main()
