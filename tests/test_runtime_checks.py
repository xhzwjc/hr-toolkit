from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from types import ModuleType
from unittest.mock import patch

from hr_toolkit import __version__
from hr_toolkit.runtime_checks import (
    CHECK_OUTPUT_ENV,
    ocr_runtime_smoke_test,
    run_headless_command,
    smoke_test,
)
from hr_toolkit.update_runner import main as update_runner_main


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

    def test_smoke_test_reads_templates_and_runs_project_lifecycle(self) -> None:
        smoke_test()

    def test_smoke_test_creates_metadata_inside_the_temporary_project(self) -> None:
        class FixedTemporaryDirectory:
            def __init__(self, path: Path) -> None:
                self.path = path

            def __enter__(self) -> str:
                return str(self.path)

            def __exit__(self, *_args: object) -> None:
                return None

        with tempfile.TemporaryDirectory() as tmp:
            temp_root = Path(tmp) / "runtime-smoke"
            temp_root.mkdir()
            with patch(
                "hr_toolkit.runtime_checks.tempfile.TemporaryDirectory",
                return_value=FixedTemporaryDirectory(temp_root),
            ):
                smoke_test()
            project_root = temp_root.resolve() / "project"
            self.assertTrue((project_root / ".hrtoolkit").is_dir())
            self.assertFalse((temp_root / "history").exists())

    def test_ocr_runtime_smoke_executes_inference_with_valid_png(self) -> None:
        calls: list[Path] = []

        class FakeEngine:
            def __call__(self, image_path: str):
                path = Path(image_path)
                calls.append(path)
                self.assert_valid_png(path)
                return None, None

            @staticmethod
            def assert_valid_png(path: Path) -> None:
                payload = path.read_bytes()
                if not payload.startswith(b"\x89PNG\r\n\x1a\n") or b"IHDR" not in payload:
                    raise AssertionError("invalid PNG fixture")

        rapidocr = ModuleType("rapidocr_onnxruntime")
        rapidocr.RapidOCR = FakeEngine  # type: ignore[attr-defined]
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict(sys.modules, {"rapidocr_onnxruntime": rapidocr}):
                ocr_runtime_smoke_test(Path(tmp))
            self.assertEqual(calls, [Path(tmp) / "blank.png"])

    def test_ocr_runtime_smoke_rejects_invalid_result_shape(self) -> None:
        class InvalidEngine:
            def __call__(self, _image_path: str):
                return []

        rapidocr = ModuleType("rapidocr_onnxruntime")
        rapidocr.RapidOCR = InvalidEngine  # type: ignore[attr-defined]
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict(sys.modules, {"rapidocr_onnxruntime": rapidocr}):
                with self.assertRaisesRegex(RuntimeError, "返回格式无效"):
                    ocr_runtime_smoke_test(Path(tmp))

    def test_unknown_arguments_are_left_for_cli(self) -> None:
        self.assertIsNone(run_headless_command(["salary-split"]))

    def test_update_smoke_command_is_headless_and_machine_verifiable(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "update-smoke.txt"
            os.environ[CHECK_OUTPUT_ENV] = str(output)
            try:
                with patch("hr_toolkit.runtime_checks.update_smoke_test", return_value="0.2.1"):
                    self.assertEqual(run_headless_command(["--update-smoke-test"]), 0)
            finally:
                os.environ.pop(CHECK_OUTPUT_ENV, None)
            self.assertEqual(
                output.read_text(encoding="utf-8"),
                f"HRToolkit {__version__} update-smoke-test OK; latest=0.2.1\n",
            )

    def test_updater_smoke_command_is_headless_and_machine_verifiable(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "updater-smoke.txt"
            os.environ[CHECK_OUTPUT_ENV] = str(output)
            try:
                self.assertEqual(update_runner_main(["--smoke-test"]), 0)
            finally:
                os.environ.pop(CHECK_OUTPUT_ENV, None)
            self.assertEqual(
                output.read_text(encoding="utf-8"),
                f"HRToolkitUpdater {__version__} smoke-test OK\n",
            )

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
