from __future__ import annotations

import io
import ssl
import tempfile
import unittest
import urllib.request
import zipfile
import os
import sys
from pathlib import Path
from unittest.mock import patch

from hr_toolkit.app_update import (
    DEFAULT_UPDATE_MANIFEST_URL,
    UpdateError,
    check_for_update,
    cleanup_stale_update_files,
    download_update_package,
    fetch_update_manifest,
    is_newer_version,
    launch_update_replacement,
    parse_update_manifest,
    sha256_file,
    trim_log_file,
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
        self.assertEqual(update.update_mode, "auto")

    def test_macos_defaults_to_manual_update(self) -> None:
        update = parse_update_manifest(
            {
                "version": "0.2.1",
                "platforms": {
                    "macos": {
                        "file_url": "HRToolkit_0.2.1_universal.dmg",
                        "sha256": "abc123",
                    }
                },
            },
            manifest_url="https://github.com/xhzwjc/hr-toolkit/releases/latest/download/latest.json",
            platform="macos",
        )

        self.assertEqual(update.update_mode, "manual")
        with self.assertRaisesRegex(UpdateError, "手动安装包"):
            download_update_package(update)

    def test_macos_manifest_selects_current_architecture_before_generic_entry(self) -> None:
        manifest = {
            "version": "0.2.1",
            "platforms": {
                "macos": {"file_url": "universal.dmg", "sha256": "universal"},
                "macos-arm64": {"file_url": "arm64.dmg", "sha256": "arm64"},
                "macos-x64": {"file_url": "x64.dmg", "sha256": "x64"},
            },
        }

        with patch("hr_toolkit.app_update.platform_module.machine", return_value="arm64"):
            arm_update = parse_update_manifest(manifest, "https://example.test/latest.json", "macos")
        with patch("hr_toolkit.app_update.platform_module.machine", return_value="x86_64"):
            x64_update = parse_update_manifest(manifest, "https://example.test/latest.json", "macos")

        self.assertEqual(arm_update.file_url, "https://example.test/arm64.dmg")
        self.assertEqual(x64_update.file_url, "https://example.test/x64.dmg")

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

    def test_https_manifest_uses_validating_certifi_context(self) -> None:
        response = io.BytesIO(b'{"version": "0.2.1"}')
        with patch("hr_toolkit.app_update.urllib.request.urlopen", return_value=response) as urlopen:
            manifest = fetch_update_manifest("https://example.test/latest.json")

        self.assertEqual(manifest["version"], "0.2.1")
        context = urlopen.call_args.kwargs["context"]
        self.assertIsInstance(context, ssl.SSLContext)
        self.assertEqual(context.verify_mode, ssl.CERT_REQUIRED)
        self.assertTrue(context.check_hostname)
        self.assertGreater(len(context.get_ca_certs()), 0)

    def test_legacy_http_manifest_does_not_receive_tls_context(self) -> None:
        response = io.BytesIO(b'{"version": "0.2.1"}')
        with patch("hr_toolkit.app_update.urllib.request.urlopen", return_value=response) as urlopen:
            fetch_update_manifest("http://hr.seedlingintl.com/hr-toolkit/latest.json")

        self.assertNotIn("context", urlopen.call_args.kwargs)

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

    def test_default_update_url_points_to_public_github_release(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            old_cwd = Path.cwd()
            old_env = os.environ.pop("HR_TOOLKIT_UPDATE_URL", None)
            try:
                os.chdir(tmp)
                self.assertEqual(
                    update_manifest_url(),
                    "https://github.com/xhzwjc/hr-toolkit/releases/latest/download/latest.json",
                )
                self.assertEqual(update_manifest_url(), DEFAULT_UPDATE_MANIFEST_URL)
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
            # 主程序启动更新器时开启进度窗口；直接调用 update_runner 的场景（测试、脚本）默认无界面
            self.assertIn("--ui", args)
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
            # 更新成功后应清理下载的更新包
            self.assertFalse(package.exists())

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

    def test_cleanup_stale_update_files_removes_only_old_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            stale_dir = tmp_dir / "hr_toolkit_update_abc"
            stale_dir.mkdir()
            (stale_dir / "HRToolkit-old.zip").write_bytes(b"zip")
            fresh_dir = tmp_dir / "hr_toolkit_extract_new"
            fresh_dir.mkdir()
            unrelated_dir = tmp_dir / "other_app_temp"
            unrelated_dir.mkdir()
            week_ago = __import__("time").time() - 7 * 86400
            os.utime(stale_dir, (week_ago, week_ago))
            os.utime(unrelated_dir, (week_ago, week_ago))

            removed = cleanup_stale_update_files(max_age_days=3, temp_dir=tmp_dir)

            self.assertEqual(removed, 1)
            self.assertFalse(stale_dir.exists())
            self.assertTrue(fresh_dir.exists())
            self.assertTrue(unrelated_dir.exists())

    def test_trim_log_file_keeps_tail(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            log_file = Path(tmp) / "update.log"
            log_file.write_bytes(b"line\n" * 400_000)  # ~2 MB

            trim_log_file(log_file, max_bytes=1024 * 1024, keep_bytes=64 * 1024)

            data = log_file.read_bytes()
            self.assertLess(len(data), 128 * 1024)
            self.assertTrue(data.startswith(b"(...earlier log trimmed...)\n"))
            self.assertTrue(data.endswith(b"line\n"))

if __name__ == "__main__":
    unittest.main()
