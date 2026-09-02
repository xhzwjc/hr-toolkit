from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from hr_toolkit import runlog


class RunLogTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.log_file = Path(self._tmp.name) / "app.log"
        os.environ[runlog.RUN_LOG_ENV] = str(self.log_file)
        os.environ.pop(runlog.RUN_LOG_JSON_ENV, None)

    def tearDown(self) -> None:
        os.environ.pop(runlog.RUN_LOG_ENV, None)
        os.environ.pop(runlog.RUN_LOG_JSON_ENV, None)
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

    def test_log_event_plain_and_json(self) -> None:
        # Plain event
        runlog.log_event("task_start", tool="salary_split", count=10)
        content = self.log_file.read_text(encoding="utf-8")
        self.assertIn("[EVENT: task_start]", content)
        self.assertIn("tool=salary_split", content)
        self.assertIn("count=10", content)

        # JSON event with standard and non-standard JSON types (datetime, set, custom)
        from datetime import datetime
        os.environ[runlog.RUN_LOG_JSON_ENV] = "1"
        now = datetime(2026, 8, 19, 10, 30, 0)
        runlog.log_event(
            "task_finish",
            tool="salary_split",
            status="success",
            started_at=now,
            tags={"tag1", "tag2"},
        )
        lines = self.log_file.read_text(encoding="utf-8").strip().splitlines()
        last_line = lines[-1]
        record = json.loads(last_line)
        self.assertEqual(record["event"], "task_finish")
        self.assertEqual(record["tool"], "salary_split")
        self.assertEqual(record["status"], "success")
        self.assertIn("2026-08-19", str(record["started_at"]))
        self.assertIn("timestamp", record)

    def test_frozen_macos_log_uses_writable_user_log_directory(self) -> None:
        os.environ.pop(runlog.RUN_LOG_ENV, None)
        user_home = Path(self._tmp.name) / "user"
        with patch.object(runlog.sys, "frozen", True, create=True), patch.object(
            runlog.sys, "platform", "darwin"
        ), patch.object(runlog.Path, "home", return_value=user_home):
            self.assertEqual(
                runlog.run_log_path(),
                user_home / "Library" / "Logs" / "HRToolkit" / runlog.RUN_LOG_FILE,
            )

    def test_frozen_windows_log_uses_local_app_data(self) -> None:
        os.environ.pop(runlog.RUN_LOG_ENV, None)
        local_app_data = Path(self._tmp.name) / "LocalAppData"
        with patch.object(runlog.sys, "frozen", True, create=True), patch.object(
            runlog.sys, "platform", "win32"
        ), patch.dict(os.environ, {"LOCALAPPDATA": str(local_app_data)}):
            self.assertEqual(
                runlog.run_log_path(),
                local_app_data / "HRToolkit" / "logs" / runlog.RUN_LOG_FILE,
            )


if __name__ == "__main__":
    unittest.main()
