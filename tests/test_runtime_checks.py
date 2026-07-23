from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from hr_toolkit import __version__
from hr_toolkit.runtime_checks import CHECK_OUTPUT_ENV, run_headless_command, smoke_test


class RuntimeChecksTest(unittest.TestCase):
    def test_version_command_is_headless_and_machine_verifiable(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "version.txt"
            os.environ[CHECK_OUTPUT_ENV] = str(output)
            try:
                self.assertEqual(run_headless_command(["--version"]), 0)
            finally:
                os.environ.pop(CHECK_OUTPUT_ENV, None)
            self.assertEqual(output.read_text(encoding="utf-8"), __version__ + "\n")

    def test_smoke_test_reads_all_packaged_templates(self) -> None:
        smoke_test()

    def test_unknown_arguments_are_left_for_cli(self) -> None:
        self.assertIsNone(run_headless_command(["salary-split"]))

    def test_module_entrypoint_reports_version_without_starting_gui(self) -> None:
        completed = subprocess.run(
            [sys.executable, "-m", "hr_toolkit", "--version"],
            check=True,
            capture_output=True,
            text=True,
        )
        self.assertEqual(completed.stdout.strip(), __version__)

    def test_pyinstaller_entrypoint_runs_smoke_test_without_starting_gui(self) -> None:
        entrypoint = Path(__file__).resolve().parents[1] / "hr_toolkit_app.py"
        completed = subprocess.run(
            [sys.executable, str(entrypoint), "--smoke-test"],
            check=True,
            capture_output=True,
            text=True,
        )
        self.assertIn(f"HRToolkit {__version__} smoke-test OK", completed.stdout)


if __name__ == "__main__":
    unittest.main()
