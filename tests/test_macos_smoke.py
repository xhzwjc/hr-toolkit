from __future__ import annotations

import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from hr_toolkit import __version__
from scripts import verify_macos_bundle


class MacSmokeDiagnosticsTests(unittest.TestCase):
    def _run_checks(self, failing_command=None, returncode=0, failure_output=""):
        expected = {
            "--version": __version__,
            "--smoke-test": f"HRToolkit {__version__} smoke-test OK",
            "--update-smoke-test": f"HRToolkit {__version__} update-smoke-test OK; latest={__version__}",
            "--qt-smoke-test": "HRToolkit Qt smoke-test OK",
        }
        calls = []

        def run(command, **kwargs):
            flag = command[-1]
            calls.append(flag)
            failed = flag == failing_command
            output = Path(kwargs["env"]["HR_TOOLKIT_CHECK_OUTPUT"])
            output.write_text(failure_output if failed else expected[flag], encoding="utf-8")
            # Windowed PyInstaller executables can have empty stdout/stderr.
            return subprocess.CompletedProcess(command, returncode if failed else 0, "", "")

        with tempfile.TemporaryDirectory() as tmp:
            app = Path(tmp) / "HRToolkit.app"
            launcher = app / "Contents" / "MacOS" / "HRToolkit"
            launcher.parent.mkdir(parents=True)
            launcher.touch()
            launcher.chmod(0o755)
            with patch("scripts.verify_macos_bundle.subprocess.run", side_effect=run):
                verify_macos_bundle.run_headless_smoke_tests(app, __version__)
        return calls

    def test_all_four_packaged_checks_remain_required(self):
        self.assertEqual(
            self._run_checks(),
            ["--version", "--smoke-test", "--update-smoke-test", "--qt-smoke-test"],
        )

    def test_nonzero_exit_preserves_result_file_when_stderr_is_empty(self):
        for flag in ("--version", "--smoke-test", "--update-smoke-test", "--qt-smoke-test"):
            with self.subTest(flag=flag), self.assertRaisesRegex(
                verify_macos_bundle.MacBundleVerificationError, "更新配置中没有 macos 平台"
            ):
                self._run_checks(flag, 1, "Traceback\n更新配置中没有 macos 平台")

    def test_wrong_output_still_fails_and_preserves_diagnostics(self):
        for flag in ("--version", "--smoke-test", "--update-smoke-test", "--qt-smoke-test"):
            with self.subTest(flag=flag), self.assertRaisesRegex(
                verify_macos_bundle.MacBundleVerificationError, "unexpected-result"
            ):
                self._run_checks(flag, 0, "unexpected-result")

    def test_large_logs_are_bounded_and_native_crashes_remain_visible(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "result.txt"
            output.write_text("discard-prefix" + "x" * 20000 + "result-tail", encoding="utf-8")
            Path(str(output) + ".native.log").write_text("native-crash", encoding="utf-8")
            result = subprocess.CompletedProcess([], 1, "stdout-info", "stderr-info")
            details = verify_macos_bundle._smoke_failure_details(result, output)
            self.assertNotIn("discard-prefix", details)
            for marker in ("result-tail", "native-crash", "stdout-info", "stderr-info"):
                self.assertIn(marker, details)
            self.assertLess(len(details), 8400)

    def test_missing_logs_do_not_hide_original_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = subprocess.CompletedProcess([], 1, None, None)
            details = verify_macos_bundle._smoke_failure_details(result, Path(tmp) / "missing.txt")
            self.assertIn("无运行检查诊断输出", details)


if __name__ == "__main__":
    unittest.main()
