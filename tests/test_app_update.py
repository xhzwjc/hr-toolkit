from __future__ import annotations

import io
import ssl
import tempfile
import unittest
import urllib.request
import zipfile
import os
import stat
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from hr_toolkit.app_update import (
    DEFAULT_UPDATE_MANIFEST_URL,
    DEFAULT_UPDATE_MANIFEST_URLS,
    GITEE_LATEST_RELEASE_API_URL,
    WIN7_UPDATER_APP_LOCAL_RUNTIME_FILES,
    UpdateCancelledError,
    UpdateError,
    check_for_update,
    cleanup_stale_update_files,
    download_update_package,
    fetch_update_manifest,
    is_newer_version,
    launch_update_replacement,
    load_update_manifest,
    parse_update_manifest,
    platform_key,
    resolve_download_url,
    sha256_file,
    trim_log_file,
    update_manifest_url,
    update_manifest_urls,
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

    def test_windows_platform_key_separates_win7_from_modern_windows(self) -> None:
        with (
            patch.object(sys, "platform", "win32"),
            patch.object(
                sys,
                "getwindowsversion",
                return_value=SimpleNamespace(major=6, minor=1),
                create=True,
            ),
        ):
            self.assertEqual(platform_key(), "windows-x64-win7")

        with (
            patch.object(sys, "platform", "win32"),
            patch.object(
                sys,
                "getwindowsversion",
                return_value=SimpleNamespace(major=6, minor=2),
                create=True,
            ),
        ):
            self.assertEqual(platform_key(), "windows-x64-win7")

        with (
            patch.object(sys, "platform", "win32"),
            patch.object(
                sys,
                "getwindowsversion",
                return_value=SimpleNamespace(major=6, minor=3),
                create=True,
            ),
        ):
            self.assertEqual(platform_key(), "windows-x64-modern")

        with (
            patch.object(sys, "platform", "win32"),
            patch.object(
                sys,
                "getwindowsversion",
                return_value=SimpleNamespace(major=10, minor=0),
                create=True,
            ),
        ):
            self.assertEqual(platform_key(), "windows-x64-modern")

        # 兼容性清单可能让旧 GetVersionEx 字段报告 6.1；Python 提供的
        # platform_version 才是实际内核版本，应优先保持现代更新通道。
        with (
            patch.object(sys, "platform", "win32"),
            patch.object(
                sys,
                "getwindowsversion",
                return_value=SimpleNamespace(
                    major=6,
                    minor=1,
                    platform_version=(10, 0, 26100),
                ),
                create=True,
            ),
        ):
            self.assertEqual(platform_key(), "windows-x64-modern")

    def test_windows_update_channels_select_matching_installer(self) -> None:
        manifest = {
            "version": "0.6.0",
            "platforms": {
                "windows": {"file_url": "modern.exe", "sha256": "modern"},
                "windows-x64-modern": {
                    "file_url": "modern-explicit.exe",
                    "sha256": "modern-explicit",
                },
                "windows-x64-win7": {
                    "file_url": "win7.exe",
                    "sha256": "win7",
                },
            },
        }
        modern = parse_update_manifest(
            manifest,
            "https://example.test/latest.json",
            "windows-x64-modern",
        )
        win7 = parse_update_manifest(
            manifest,
            "https://example.test/latest.json",
            "windows-x64-win7",
        )
        legacy_caller = parse_update_manifest(
            manifest,
            "https://example.test/latest.json",
            "windows",
        )

        self.assertEqual(modern.file_url, "https://example.test/modern-explicit.exe")
        self.assertEqual(win7.file_url, "https://example.test/win7.exe")
        self.assertEqual(legacy_caller.file_url, "https://example.test/modern.exe")

    def test_win7_update_channel_never_falls_back_to_modern_installer(self) -> None:
        manifest = {
            "version": "0.6.0",
            "platforms": {
                "windows": {"file_url": "modern.exe", "sha256": "modern"},
            },
        }
        with self.assertRaisesRegex(UpdateError, "windows-x64-win7"):
            parse_update_manifest(
                manifest,
                "https://example.test/latest.json",
                "windows-x64-win7",
            )

    def test_runtime_downloads_keep_only_gitee_from_release_metadata(self) -> None:
        update = parse_update_manifest(
            {
                "version": "0.2.3",
                "file_url": "https://gitee.com/company/hr/releases/download/v0.2.3/update.zip",
                "fallback_urls": [
                    "https://github.com/company/hr/releases/download/v0.2.3/update.zip",
                    "https://gitee.com/company/hr/releases/download/v0.2.3/update.zip",
                ],
                "sha256": "abc123",
            },
            manifest_url="https://gitee.com/company/hr/releases/download/v0.2.3/latest.json",
            platform="windows",
        )

        self.assertEqual(update.file_url, "https://gitee.com/company/hr/releases/download/v0.2.3/update.zip")
        self.assertEqual(
            update.download_urls,
            (
                "https://gitee.com/company/hr/releases/download/v0.2.3/update.zip",
            ),
        )

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

    def test_remote_manifest_cannot_reference_local_file(self) -> None:
        with self.assertRaisesRegex(UpdateError, "不能引用本机文件"):
            parse_update_manifest(
                {"version": "0.2.0", "file_url": "file:///tmp/update.zip", "sha256": "abc"},
                manifest_url="https://example.test/latest.json",
                platform="windows",
            )

    def test_update_network_rejects_unsupported_protocols(self) -> None:
        with self.assertRaisesRegex(UpdateError, "不支持的更新地址协议"):
            fetch_update_manifest("ftp://example.test/latest.json")

    def test_check_for_update_allows_current_version_without_package_fields(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            manifest = Path(tmp) / "latest.json"
            manifest.write_text('{"version": "0.1.0"}', encoding="utf-8")
            manifest_url = "file://" + urllib.request.pathname2url(str(manifest))

            self.assertIsNone(check_for_update("0.1.0", manifest_url=manifest_url, platform="windows"))

    def test_check_for_update_does_not_fall_back_to_github_after_gitee_failure(self) -> None:
        with patch(
            "hr_toolkit.app_update.load_update_manifest",
            side_effect=UpdateError("timed out"),
        ) as loader:
            with self.assertRaisesRegex(UpdateError, "Gitee"):
                check_for_update("0.2.2", platform="windows")

        self.assertEqual(
            [call.args[0] for call in loader.call_args_list],
            list(DEFAULT_UPDATE_MANIFEST_URLS),
        )

    def test_check_for_update_does_not_query_github_when_gitee_succeeds(self) -> None:
        with patch(
            "hr_toolkit.app_update.load_update_manifest",
            return_value=({"version": "0.2.2"}, "https://gitee.com/latest.json"),
        ) as loader:
            self.assertIsNone(check_for_update("0.2.2", platform="windows"))

        loader.assert_called_once_with(GITEE_LATEST_RELEASE_API_URL)

    def test_gitee_latest_release_resolves_attached_manifest(self) -> None:
        release = {
            "tag_name": "v0.2.3",
            "assets": [
                {
                    "name": "latest.json",
                    "browser_download_url": "https://gitee.com/company/hr/releases/download/v0.2.3/latest.json",
                }
            ],
        }
        manifest = {"version": "0.2.3"}
        with patch(
            "hr_toolkit.app_update._fetch_json_object",
            side_effect=[release, manifest],
        ) as fetcher:
            payload, resolved_url = load_update_manifest(GITEE_LATEST_RELEASE_API_URL)

        self.assertEqual(payload, manifest)
        self.assertEqual(resolved_url, release["assets"][0]["browser_download_url"])
        self.assertEqual(fetcher.call_count, 2)

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

    def test_oversized_manifest_is_rejected_with_bounded_read(self) -> None:
        response = io.BytesIO(b'{"version":"0.2.1","padding":"xxxxxxxx"}')
        with (
            patch("hr_toolkit.app_update.UPDATE_MANIFEST_MAX_BYTES", 16),
            patch("hr_toolkit.app_update.urllib.request.urlopen", return_value=response),
            self.assertRaisesRegex(UpdateError, "配置文件过大"),
        ):
            fetch_update_manifest("https://example.test/latest.json")

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

    def test_download_package_falls_back_after_primary_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            missing = tmp_dir / "missing.zip"
            fallback = tmp_dir / "fallback.zip"
            fallback.write_bytes(b"fallback payload")
            update = parse_update_manifest(
                {
                    "version": "0.2.3",
                    "file_url": "file://" + urllib.request.pathname2url(str(missing)),
                    "fallback_urls": [
                        "file://" + urllib.request.pathname2url(str(fallback)),
                    ],
                    "sha256": sha256_file(fallback),
                },
                manifest_url="file://" + urllib.request.pathname2url(str(tmp_dir / "latest.json")),
                platform="windows",
            )

            downloaded = download_update_package(update, dest_dir=tmp_dir / "download")
            self.assertEqual(downloaded.read_bytes(), b"fallback payload")

    def test_download_package_cancelled_by_event(self) -> None:
        import threading
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            pkg = tmp_dir / "package.zip"
            pkg.write_bytes(b"A" * (1024 * 512))
            update = parse_update_manifest(
                {
                    "version": "0.2.3",
                    "file_url": "file://" + urllib.request.pathname2url(str(pkg)),
                    "sha256": sha256_file(pkg),
                },
                manifest_url="file://" + urllib.request.pathname2url(str(tmp_dir / "latest.json")),
                platform="windows",
            )

            cancel_event = threading.Event()
            # 设置取消事件
            cancel_event.set()
            with self.assertRaises(UpdateCancelledError):
                download_update_package(update, dest_dir=tmp_dir / "download", cancel_event=cancel_event)

            # 验证临时文件已被自动清理干净
            download_dir = tmp_dir / "download"
            if download_dir.exists():
                files = list(download_dir.iterdir())
                self.assertEqual(files, [], "取消下载后临时文件必须被完全清理")

    def test_download_package_rejects_body_beyond_size_limit(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            package = tmp_dir / "oversized.zip"
            package.write_bytes(b"0123456789")
            update = parse_update_manifest(
                {
                    "version": "0.2.3",
                    "file_url": package.as_uri(),
                    "sha256": sha256_file(package),
                },
                manifest_url=(tmp_dir / "latest.json").as_uri(),
                platform="windows",
            )

            with (
                patch("hr_toolkit.app_update.UPDATE_PACKAGE_MAX_BYTES", 8),
                self.assertRaisesRegex(UpdateError, "最大体积"),
            ):
                download_update_package(update, dest_dir=tmp_dir / "download")

            self.assertEqual(list((tmp_dir / "download").iterdir()), [])

    def test_manual_download_url_does_not_probe_github(self) -> None:
        update = parse_update_manifest(
            {
                "version": "0.2.3",
                "file_url": "https://gitee.com/company/hr/releases/download/v0.2.3/app.dmg",
                "fallback_urls": ["https://github.com/company/hr/releases/download/v0.2.3/app.dmg"],
                "sha256": "abc123",
                "update_mode": "manual",
            },
            manifest_url="https://gitee.com/company/hr/releases/download/v0.2.3/latest.json",
            platform="macos",
        )
        with patch("hr_toolkit.app_update._open_url", side_effect=OSError("timed out")) as opener:
            with self.assertRaisesRegex(UpdateError, "Gitee"):
                resolve_download_url(update)

        self.assertEqual(opener.call_count, 1)

    def test_update_url_file_overrides_default(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            (tmp_dir / "update_url.txt").write_text(
                "https://gitee.com/company/hr-toolkit/releases/latest\n"
                "https://github.com/company/hr-toolkit/releases/latest/download/latest.json\n",
                encoding="utf-8",
            )
            old_cwd = Path.cwd()
            old_env = os.environ.pop("HR_TOOLKIT_UPDATE_URL", None)
            try:
                os.chdir(tmp_dir)
                self.assertEqual(
                    update_manifest_url(),
                    "https://gitee.com/company/hr-toolkit/releases/latest",
                )
                self.assertEqual(
                    update_manifest_urls(),
                    ("https://gitee.com/company/hr-toolkit/releases/latest",),
                )
            finally:
                os.chdir(old_cwd)
                if old_env is not None:
                    os.environ["HR_TOOLKIT_UPDATE_URL"] = old_env

    def test_default_update_urls_use_gitee_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            old_cwd = Path.cwd()
            old_env = os.environ.pop("HR_TOOLKIT_UPDATE_URL", None)
            try:
                os.chdir(tmp)
                self.assertEqual(
                    update_manifest_url(),
                    GITEE_LATEST_RELEASE_API_URL,
                )
                self.assertEqual(update_manifest_url(), DEFAULT_UPDATE_MANIFEST_URL)
                self.assertEqual(
                    update_manifest_urls(),
                    (GITEE_LATEST_RELEASE_API_URL,),
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

    def test_update_runner_rejects_traversal_and_case_conflicts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            log_file = root / "update.log"
            for name, entries in (
                ("traversal.zip", [("../outside.txt", "bad")]),
                ("case-conflict.zip", [("Data/file.txt", "one"), ("data/FILE.txt", "two")]),
            ):
                package = root / name
                with zipfile.ZipFile(package, "w") as archive:
                    for member_name, content in entries:
                        archive.writestr(member_name, content)
                extract_dir = root / f"extract-{name}"
                extract_dir.mkdir()

                with self.subTest(name=name), self.assertRaisesRegex(RuntimeError, "非法路径|大小写冲突"):
                    update_runner._safe_extract_zip(package, extract_dir, log_file)

            self.assertFalse((root / "outside.txt").exists())

    def test_update_runner_rejects_links_and_oversized_payloads(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            log_file = root / "update.log"
            link_package = root / "link.zip"
            link_info = zipfile.ZipInfo("_internal/link")
            link_info.create_system = 3
            link_info.external_attr = (stat.S_IFLNK | 0o777) << 16
            with zipfile.ZipFile(link_package, "w") as archive:
                archive.writestr(link_info, "target")

            link_extract = root / "link-extract"
            link_extract.mkdir()
            with self.assertRaisesRegex(RuntimeError, "链接或特殊文件"):
                update_runner._safe_extract_zip(link_package, link_extract, log_file)

            large_package = root / "large.zip"
            with zipfile.ZipFile(large_package, "w") as archive:
                archive.writestr("payload.bin", b"12")
            large_extract = root / "large-extract"
            large_extract.mkdir()
            with (
                patch.object(update_runner, "ZIP_MAX_TOTAL_BYTES", 1),
                self.assertRaisesRegex(RuntimeError, "总体积异常"),
            ):
                update_runner._safe_extract_zip(large_package, large_extract, log_file)

    def test_update_runner_rejects_launcher_paths_before_replacement(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            app_dir = root / "HRToolkit"
            app_dir.mkdir()
            launcher = app_dir / "HRToolkit.exe"
            launcher.write_text("old", encoding="utf-8")
            package = root / "update.zip"
            with zipfile.ZipFile(package, "w") as archive:
                archive.writestr("HRToolkit.exe", "new")
                archive.writestr("_internal/data.txt", "data")

            exit_code = self._run_update_runner([
                "--zip", str(package),
                "--app-dir", str(app_dir),
                "--launcher", "../outside.exe",
                "--log-file", str(root / "update.log"),
            ])

            self.assertEqual(exit_code, 1)
            self.assertEqual(launcher.read_text(encoding="utf-8"), "old")
            self.assertTrue(package.exists())

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

    def test_update_runner_runs_installer(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            app_dir.mkdir()
            (app_dir / "HRToolkit.exe").write_text("old", encoding="utf-8")

            installer = tmp_dir / "HRToolkit_setup.exe"
            installer.write_bytes(b"MZfake_setup")
            log_file = tmp_dir / "HRToolkit_update.log"

            with patch("subprocess.run") as mock_run:
                mock_run.return_value.returncode = 0
                exit_code = self._run_update_runner([
                    "--installer",
                    str(installer),
                    "--app-dir",
                    str(app_dir),
                    "--launcher",
                    "HRToolkit.exe",
                    "--log-file",
                    str(log_file),
                ])

                self.assertEqual(exit_code, 0)
                mock_run.assert_called_once_with(
                    [str(installer.resolve()), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"],
                    check=False,
                )
                self.assertFalse(installer.exists())

    def test_launch_update_replacement_with_exe_uses_installer_flag(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            app_dir.mkdir()
            updater_name = "HRToolkitUpdater.exe" if sys.platform.startswith("win") else "HRToolkitUpdater"
            updater = app_dir / updater_name
            updater.write_text("updater", encoding="utf-8")
            installer = tmp_dir / "HRToolkit_0.3.5_x64-setup.exe"
            installer.write_bytes(b"MZ")

            captured: dict = {}

            def fake_popen(args, **kwargs):  # type: ignore[no-untyped-def]
                captured["args"] = args
                captured["kwargs"] = kwargs
                return object()

            with patch("subprocess.Popen", side_effect=fake_popen):
                launch_update_replacement(
                    package_path=installer,
                    app_dir=app_dir,
                    launcher_path=app_dir / "HRToolkit.exe",
                    wait_pid=1234,
                )

            args = captured["args"]
            self.assertIn("--installer", args)
            self.assertIn(str(installer), args)
            self.assertNotIn("--zip", args)

    def test_win7_temporary_updater_keeps_app_local_runtime_beside_exe(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            app_dir.mkdir()
            (app_dir / "HRToolkitUpdater.exe").write_bytes(b"updater")
            for name in WIN7_UPDATER_APP_LOCAL_RUNTIME_FILES:
                (app_dir / name).write_bytes(b"pinned:" + name.encode("ascii"))
            installer = tmp_dir / "HRToolkit_0.6.0_win7_x64-setup.exe"
            installer.write_bytes(b"MZ")
            temp_updater_dir = tmp_dir / "temp-updater"
            temp_updater_dir.mkdir()

            with (
                patch.object(sys, "platform", "win32"),
                patch.object(
                    sys,
                    "getwindowsversion",
                    return_value=SimpleNamespace(major=6, minor=1),
                    create=True,
                ),
                patch(
                    "hr_toolkit.app_update.tempfile.mkdtemp",
                    return_value=str(temp_updater_dir),
                ),
                patch("subprocess.Popen"),
            ):
                launch_update_replacement(
                    package_path=installer,
                    app_dir=app_dir,
                    launcher_path=app_dir / "HRToolkit.exe",
                    wait_pid=1234,
                )

            self.assertEqual(
                (temp_updater_dir / "HRToolkitUpdater.exe").read_bytes(),
                b"updater",
            )
            for name in WIN7_UPDATER_APP_LOCAL_RUNTIME_FILES:
                self.assertEqual(
                    (temp_updater_dir / name).read_bytes(),
                    (app_dir / name).read_bytes(),
                )


if __name__ == "__main__":
    unittest.main()
