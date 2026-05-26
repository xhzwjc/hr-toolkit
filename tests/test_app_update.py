from __future__ import annotations

import tempfile
import unittest
import urllib.request
import zipfile
import os
import sys
from pathlib import Path

from hr_toolkit.app_update import (
    UpdateError,
    check_for_update,
    download_update_package,
    is_newer_version,
    launch_update_replacement,
    parse_update_manifest,
    sha256_file,
    update_manifest_url,
)
from hr_toolkit import update_runner
from hr_toolkit.update_runner import main as update_runner_main


class AppUpdateTests(unittest.TestCase):
    def _run_update_runner(self, args: list[str]) -> int:
        old_cwd = Path.cwd()
        try:
            return update_runner_main(args)
        finally:
            os.chdir(old_cwd)

    def test_version_compare(self) -> None:
        self.assertTrue(is_newer_version("0.2.0", "0.1.9"))
        self.assertTrue(is_newer_version("v1.0.1", "1.0.0"))
        self.assertFalse(is_newer_version("1.0.0", "1.0.0"))
        self.assertFalse(is_newer_version("1.0.0", "1.0.1"))

    def test_parse_platform_manifest(self) -> None:
        manifest = {
            "version": "0.2.0",
            "notes": ["修复问题"],
            "platforms": {
                "windows": {
                    "file_url": "releases/HRToolkit-0.2.0-win.zip",
                    "sha256": "abc123",
                }
            },
        }

        update = parse_update_manifest(
            manifest,
            manifest_url="http://hr.seedlingintl.com/hr-toolkit/latest.json",
            platform="windows",
        )

        self.assertEqual(update.version, "0.2.0")
        self.assertEqual(update.file_url, "http://hr.seedlingintl.com/hr-toolkit/releases/HRToolkit-0.2.0-win.zip")
        self.assertEqual(update.sha256, "abc123")
        self.assertEqual(update.notes, ("修复问题",))

    def test_parse_manifest_requires_platform(self) -> None:
        manifest = {"version": "0.2.0", "platforms": {"macos": {"file_url": "mac.zip", "sha256": "abc"}}}

        with self.assertRaises(UpdateError):
            parse_update_manifest(manifest, manifest_url="http://example.test/latest.json", platform="windows")

    def test_check_for_update_allows_current_version_without_package_fields(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            manifest = Path(tmp) / "latest.json"
            manifest.write_text('{"version": "0.1.0"}', encoding="utf-8")
            manifest_url = "file://" + urllib.request.pathname2url(str(manifest))

            self.assertIsNone(check_for_update("0.1.0", manifest_url=manifest_url, platform="windows"))

    def test_download_package_verifies_sha256(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            source = tmp_dir / "HRToolkit-0.2.0-win.zip"
            source.write_bytes(b"fake zip payload")
            source_url = "file://" + urllib.request.pathname2url(str(source))
            update = parse_update_manifest(
                {
                    "version": "0.2.0",
                    "file_url": source_url,
                    "sha256": sha256_file(source),
                },
                manifest_url="file://" + urllib.request.pathname2url(str(tmp_dir / "latest.json")),
                platform="windows",
            )

            downloaded = download_update_package(update, dest_dir=tmp_dir / "download")

            self.assertEqual(downloaded.read_bytes(), b"fake zip payload")

    def test_update_url_file_overrides_default(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            (tmp_dir / "update_url.txt").write_text(
                "https://gitee.com/company/hr-toolkit/raw/master/release/latest.json\n",
                encoding="utf-8",
            )
            old_cwd = Path.cwd()
            old_env = os.environ.pop("HR_TOOLKIT_UPDATE_URL", None)
            try:
                os.chdir(tmp_dir)
                self.assertEqual(
                    update_manifest_url(),
                    "https://gitee.com/company/hr-toolkit/raw/master/release/latest.json",
                )
            finally:
                os.chdir(old_cwd)
                if old_env is not None:
                    os.environ["HR_TOOLKIT_UPDATE_URL"] = old_env

    def test_launch_update_prefers_updater_from_package(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            app_dir.mkdir()
            updater_name = "HRToolkitUpdater.exe" if sys.platform.startswith("win") else "HRToolkitUpdater"
            (app_dir / updater_name).write_text("old updater", encoding="utf-8")
            launcher = app_dir / ("HRToolkit.exe" if sys.platform.startswith("win") else "HRToolkit")
            launcher.write_text("old app", encoding="utf-8")

            package = tmp_dir / "update.zip"
            with zipfile.ZipFile(package, "w") as archive:
                archive.writestr(updater_name, "new updater")
                archive.writestr(launcher.name, "new app")
                archive.writestr("_internal/data.txt", "data")

            captured: dict[str, object] = {}
            original_popen = __import__("subprocess").Popen

            def fake_popen(args, **kwargs):  # type: ignore[no-untyped-def]
                captured["args"] = args
                captured["kwargs"] = kwargs

                class Process:
                    pid = 123

                return Process()

            try:
                __import__("subprocess").Popen = fake_popen
                launch_update_replacement(package, app_dir=app_dir, launcher_path=launcher, wait_pid=99)
            finally:
                __import__("subprocess").Popen = original_popen

            args = captured["args"]
            self.assertIsInstance(args, list)
            updater_path = Path(args[0])
            self.assertEqual(updater_path.read_text(encoding="utf-8"), "new updater")
            self.assertIn("--log-file", args)
            self.assertEqual(captured["kwargs"].get("cwd"), str(tmp_dir))
            self.assertTrue((tmp_dir / "HRToolkit_update.log").exists())

    def test_update_runner_replaces_app_directory(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            app_dir.mkdir()
            (app_dir / "HRToolkit.exe").write_text("old", encoding="utf-8")

            payload_dir = tmp_dir / "payload"
            payload_dir.mkdir()
            (payload_dir / "HRToolkit.exe").write_text("new", encoding="utf-8")
            (payload_dir / "_internal").mkdir()
            (payload_dir / "_internal" / "data.txt").write_text("data", encoding="utf-8")

            package = tmp_dir / "update.zip"
            with zipfile.ZipFile(package, "w") as archive:
                for file_path in payload_dir.rglob("*"):
                    if file_path.is_file():
                        archive.write(file_path, file_path.relative_to(payload_dir))

            log_file = tmp_dir / "HRToolkit_update.log"
            exit_code = self._run_update_runner([
                "--zip",
                str(package),
                "--app-dir",
                str(app_dir),
                "--launcher",
                "HRToolkit.exe",
                "--log-file",
                str(log_file),
            ])

            self.assertEqual(exit_code, 0)
            self.assertEqual((app_dir / "HRToolkit.exe").read_text(encoding="utf-8"), "new")
            self.assertTrue((app_dir / "_internal" / "data.txt").exists())
            self.assertIn("工作目录已切换到：", log_file.read_text(encoding="utf-8"))

    def test_update_runner_handles_empty_target_reappearing_during_replace(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            app_dir.mkdir()
            (app_dir / "HRToolkit.exe").write_text("old", encoding="utf-8")
            (app_dir / "_internal").mkdir()

            payload_dir = tmp_dir / "payload"
            payload_dir.mkdir()
            (payload_dir / "HRToolkit.exe").write_text("new", encoding="utf-8")
            (payload_dir / "_internal").mkdir()
            (payload_dir / "_internal" / "data.txt").write_text("data", encoding="utf-8")

            package = tmp_dir / "update.zip"
            with zipfile.ZipFile(package, "w") as archive:
                for file_path in payload_dir.rglob("*"):
                    if file_path.is_file():
                        archive.write(file_path, file_path.relative_to(payload_dir))

            original_rename = update_runner.os.rename

            def rename_and_recreate_empty_target(source, target):  # type: ignore[no-untyped-def]
                result = original_rename(source, target)
                if Path(source) == app_dir and "HRToolkit_backup_" in Path(target).name:
                    app_dir.mkdir()
                return result

            try:
                update_runner.os.rename = rename_and_recreate_empty_target
                exit_code = self._run_update_runner([
                    "--zip",
                    str(package),
                    "--app-dir",
                    str(app_dir),
                    "--launcher",
                    "HRToolkit.exe",
                    "--log-file",
                    str(tmp_dir / "HRToolkit_update.log"),
                ])
            finally:
                update_runner.os.rename = original_rename

            self.assertEqual(exit_code, 0)
            self.assertEqual((app_dir / "HRToolkit.exe").read_text(encoding="utf-8"), "new")
            self.assertTrue((app_dir / "_internal" / "data.txt").exists())
            self.assertFalse(any(item.name.startswith("HRToolkit_new_") for item in app_dir.iterdir()))

    def test_update_runner_restores_backup_when_payload_is_invalid(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            app_dir.mkdir()
            (app_dir / "HRToolkit.exe").write_text("old", encoding="utf-8")
            (app_dir / "_internal").mkdir()

            payload_dir = tmp_dir / "bad_payload"
            payload_dir.mkdir()
            (payload_dir / "readme.txt").write_text("bad", encoding="utf-8")

            package = tmp_dir / "bad_update.zip"
            log_file = tmp_dir / "HRToolkit_update.log"
            with zipfile.ZipFile(package, "w") as archive:
                archive.write(payload_dir / "readme.txt", "readme.txt")

            exit_code = self._run_update_runner([
                "--zip",
                str(package),
                "--app-dir",
                str(app_dir),
                "--launcher",
                "HRToolkit.exe",
                "--log-file",
                str(log_file),
            ])

            self.assertEqual(exit_code, 1)
            self.assertEqual((app_dir / "HRToolkit.exe").read_text(encoding="utf-8"), "old")
            self.assertTrue((app_dir / "_internal").exists())
            self.assertIn("更新失败", log_file.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
