from __future__ import annotations

import builtins
import copy
import hashlib
import json
import runpy
import struct
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from scripts import build_update_assets
from scripts import build_macos
from scripts import build_windows
from scripts import build_windows_installers
from scripts import release_windows
from scripts import prepare_macos_x64_runtime
from scripts import prepare_win7_runtime
from scripts import verify_macos_bundle
from hr_toolkit.app_update import WIN7_UPDATER_APP_LOCAL_RUNTIME_FILES
from hr_toolkit.runtime_checks import TEMPLATE_NAMES


class WindowsPackagingTests(unittest.TestCase):
    @property
    def version(self) -> str:
        return build_windows.read_project_version()

    def test_semver_is_canonical_and_must_match_project(self) -> None:
        self.assertEqual(build_windows.validate_stable_semver("0.2.1"), (0, 2, 1))
        for invalid in ("v0.2.1", "0.2", "0.2.1-rc.1", "01.2.3", "1.02.3", "1.2.03"):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                build_windows.validate_stable_semver(invalid)
        with self.assertRaises(ValueError):
            build_windows.validate_build_version("99.99.99")

    def test_packaged_tutorial_contract_does_not_load_legacy_gui(self) -> None:
        command = (
            "import sys; "
            "from hr_toolkit.tutorial_content import tutorial_groups; "
            "assert tutorial_groups(); "
            "assert 'hr_toolkit.gui' not in sys.modules"
        )
        completed = subprocess.run(
            [sys.executable, "-c", command],
            cwd=build_windows.REPO_ROOT,
            check=False,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)

    def test_pyinstaller_commands_are_onedir_onefile_and_resource_whitelisted(self) -> None:
        self.assertEqual(build_windows.RELEASE_TEMPLATE_NAMES, TEMPLATE_NAMES)
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            main, updater = build_windows.pyinstaller_commands(
                version=self.version,
                output_dir=tmp_dir / "dist",
                work_dir=tmp_dir / "build",
                version_file=tmp_dir / "version.txt",
            )
            explicit_modern_main, explicit_modern_updater = build_windows.pyinstaller_commands(
                version=self.version,
                output_dir=tmp_dir / "dist",
                work_dir=tmp_dir / "build",
                version_file=tmp_dir / "version.txt",
                target=build_windows.WINDOWS_TARGET_MODERN,
            )

        self.assertEqual(main, explicit_modern_main)
        self.assertEqual(updater, explicit_modern_updater)
        self.assertIn("--onedir", main)
        self.assertIn("--windowed", main)
        self.assertNotIn("--onefile", main)
        self.assertIn("--onefile", updater)
        self.assertIn("--windowed", updater)
        self.assertNotIn("--onedir", updater)
        self.assertIn(str(build_windows.WINDOWS_MANIFEST), main)
        self.assertNotIn(str(build_windows.WINDOWS_WIN7_MANIFEST), main)
        self.assertNotIn("--add-binary", main)
        self.assertNotIn("--add-binary", updater)
        self.assertIn(
            ["--additional-hooks-dir", str(build_windows.QT_PYINSTALLER_HOOKS_DIR)],
            [main[index : index + 2] for index in range(len(main) - 1)],
        )
        self.assertNotIn("--additional-hooks-dir", updater)
        self.assertEqual(main[-1], str(build_windows.APP_ENTRYPOINT))
        self.assertEqual(updater[-1], str(build_windows.UPDATER_ENTRYPOINT))

        # PyInstaller's setuptools hook aliases the vendored implementation to
        # distutils. Excluding it first makes PyInstaller 6.21 fail while
        # constructing the module graph on both Windows and macOS.
        self.assertNotIn("distutils", build_windows.EXCLUDED_MODULES)
        self.assertNotIn(
            ["--exclude-module", "distutils"],
            [main[index : index + 2] for index in range(len(main) - 1)],
        )
        self.assertNotIn(
            ["--exclude-module", "distutils"],
            [updater[index : index + 2] for index in range(len(updater) - 1)],
        )

        for excluded in build_windows.EXCLUDED_MODULES:
            self.assertIn(excluded, main)
            self.assertIn(excluded, updater)
        self.assertIn("tkinter", build_windows.MAIN_APP_EXCLUDED_MODULES)
        self.assertIn("hr_toolkit.gui", build_windows.MAIN_APP_EXCLUDED_MODULES)
        self.assertIn("PIL.ImageTk", build_windows.MAIN_APP_EXCLUDED_MODULES)
        self.assertIn("PIL.AvifImagePlugin", build_windows.MAIN_APP_EXCLUDED_MODULES)
        self.assertIn("PIL._avif", build_windows.MAIN_APP_EXCLUDED_MODULES)
        self.assertIn(
            ["--exclude-module", "tkinter"],
            [main[index : index + 2] for index in range(len(main) - 1)],
        )
        self.assertNotIn(
            ["--exclude-module", "tkinter"],
            [updater[index : index + 2] for index in range(len(updater) - 1)],
        )
        self.assertEqual(build_windows.QT6_WINDOWS_RUNTIME_ROOT, Path("PySide6"))
        self.assertEqual(build_windows.QT5_WINDOWS_RUNTIME_ROOT, Path("PySide2"))
        controller_source = (
            build_windows.REPO_ROOT / "hr_toolkit" / "gui_qt" / "controller.py"
        ).read_text(encoding="utf-8")
        self.assertIn(
            "from hr_toolkit.tutorial_content import tutorial_groups",
            controller_source,
        )
        self.assertNotIn("from hr_toolkit.gui", controller_source)
        for hidden_import in build_windows.HIDDEN_IMPORTS:
            self.assertIn(["--hidden-import", hidden_import], [main[index : index + 2] for index in range(len(main) - 1)])
        for module in build_windows.COLLECT_ALL_MODULES:
            self.assertIn(["--collect-all", module], [main[index : index + 2] for index in range(len(main) - 1)])
        self.assertIn(
            ["--copy-metadata", "pypdf"],
            [main[index : index + 2] for index in range(len(main) - 1)],
        )
        self.assertNotIn("pypdfium2", main)

        data_values = [main[index + 1] for index, value in enumerate(main[:-1]) if value == "--add-data"]
        self.assertEqual(len(data_values), 3 + len(build_windows.release_template_files()))
        self.assertTrue(any(value.startswith(str(build_windows.README_FILE) + ";") for value in data_values))
        self.assertIn(
            f"{build_windows.QML_DIR};hr_toolkit/gui_qt/qml",
            data_values,
        )
        self.assertIn(
            f"{build_windows.QT_NOTICE};third_party/qt",
            data_values,
        )
        template_sources = {
            value.split(";", 1)[0]
            for value in data_values
            if value.lower().endswith(";hr_toolkit/templates")
        }
        self.assertEqual(template_sources, {str(path) for path in build_windows.release_template_files()})
        self.assertFalse(any(value.startswith(str(build_windows.TEMPLATES_DIR) + ";") for value in data_values))
        self.assertFalse(any("附件" in value or "outputs" in value for value in data_values))

    def test_win7_pyinstaller_lane_is_isolated_and_bundles_compatibility_runtimes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            seven_zip = tmp_dir / "7zip"
            ucrt = tmp_dir / "ucrt"
            vc_runtime = tmp_dir / "vc-runtime"
            seven_zip.mkdir()
            ucrt.mkdir()
            vc_runtime.mkdir()
            for name in build_windows.WIN7_REQUIRED_7ZIP_FILES:
                (seven_zip / name).write_bytes(b"runtime")
            for name in build_windows.WIN7_REQUIRED_UCRT_FILES:
                (ucrt / name).write_bytes(b"runtime")
            for name in build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES:
                (vc_runtime / name).write_bytes(b"runtime")

            main, updater = build_windows.pyinstaller_commands(
                version=self.version,
                output_dir=tmp_dir / "dist",
                work_dir=tmp_dir / "build",
                version_file=tmp_dir / "version.txt",
                target=build_windows.WINDOWS_TARGET_WIN7,
                seven_zip_dir=seven_zip,
                ucrt_dir=ucrt,
                vc_runtime_dir=vc_runtime,
            )

        self.assertIn(str(build_windows.WINDOWS_WIN7_MANIFEST), main)
        expected_hook_option = [
            "--additional-hooks-dir",
            str(build_windows.WIN7_PYINSTALLER_HOOKS_DIR),
        ]
        self.assertIn(
            expected_hook_option,
            [main[index : index + 2] for index in range(len(main) - 1)],
        )
        self.assertIn(
            expected_hook_option,
            [updater[index : index + 2] for index in range(len(updater) - 1)],
        )
        self.assertNotIn("unrar.cffi.rarfile", main)
        self.assertNotIn("unrar", main)
        for hidden_import in build_windows.WIN7_HIDDEN_IMPORTS:
            self.assertIn(
                ["--hidden-import", hidden_import],
                [main[index : index + 2] for index in range(len(main) - 1)],
            )
        for module in build_windows.WIN7_COLLECT_ALL_MODULES:
            self.assertIn(
                ["--collect-all", module],
                [main[index : index + 2] for index in range(len(main) - 1)],
            )
        self.assertIn(
            ["--copy-metadata", "pypdfium2"],
            [main[index : index + 2] for index in range(len(main) - 1)],
        )
        self.assertNotIn(
            ["--copy-metadata", "pypdf"],
            [main[index : index + 2] for index in range(len(main) - 1)],
        )
        self.assertIn(str(seven_zip.resolve() / "7z.exe") + ";third_party/7zip", main)
        self.assertIn(str(seven_zip.resolve() / "7z.dll") + ";third_party/7zip", main)
        for name in build_windows.WIN7_REQUIRED_UCRT_FILES:
            expected = str(ucrt.resolve() / name) + ";."
            self.assertNotIn(expected, main)
            self.assertIn(expected, updater)
        for name in build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES:
            expected = str(vc_runtime.resolve() / name) + ";."
            self.assertNotIn(expected, main)
            self.assertIn(expected, updater)

    def test_win7_app_local_runtimes_are_staged_beside_main_executable(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            internal = app_dir / "_internal"
            ucrt_dir = tmp_dir / "ucrt"
            vc_runtime_dir = tmp_dir / "vc-runtime"
            internal.mkdir(parents=True)
            ucrt_dir.mkdir()
            vc_runtime_dir.mkdir()

            for source_dir, label, names in (
                (ucrt_dir, b"pinned-ucrt:", build_windows.WIN7_REQUIRED_UCRT_FILES),
                (
                    vc_runtime_dir,
                    b"pinned-vc:",
                    build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES,
                ),
            ):
                for name in names:
                    (source_dir / name).write_bytes(label + name.encode("ascii"))
                    (internal / name).write_bytes(b"runner-runtime")

            build_windows.stage_win7_app_local_runtimes(
                app_dir=app_dir,
                ucrt_dir=ucrt_dir,
                vc_runtime_dir=vc_runtime_dir,
            )

            for source_dir, names in (
                (ucrt_dir, build_windows.WIN7_REQUIRED_UCRT_FILES),
                (vc_runtime_dir, build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES),
            ):
                for name in names:
                    self.assertEqual((app_dir / name).read_bytes(), (source_dir / name).read_bytes())
                    self.assertFalse((internal / name).exists())

    def test_modern_vc_runtimes_replace_private_qt_copies(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            pyside_dir = app_dir / "_internal" / "PySide6"
            shiboken_dir = app_dir / "_internal" / "shiboken6"
            runtime_dir = tmp_dir / "system32"
            pyside_dir.mkdir(parents=True)
            shiboken_dir.mkdir(parents=True)
            runtime_dir.mkdir()
            available = (
                *build_windows.MODERN_REQUIRED_VC_RUNTIME_FILES,
                *build_windows.MODERN_OPTIONAL_VC_RUNTIME_FILES[:2],
            )
            for name in available:
                payload = b"current-vc:" + name.encode("ascii")
                (runtime_dir / name).write_bytes(payload)
                (pyside_dir / name).write_bytes(b"old-pyside-vc")
                (shiboken_dir / name).write_bytes(b"old-shiboken-vc")

            staged = build_windows.stage_modern_app_local_vc_runtimes(
                app_dir=app_dir,
                runtime_dir=runtime_dir,
            )

            self.assertEqual(staged, available)
            for name in available:
                self.assertEqual(
                    (app_dir / name).read_bytes(),
                    (runtime_dir / name).read_bytes(),
                )
                self.assertFalse((pyside_dir / name).exists())
                self.assertFalse((shiboken_dir / name).exists())
            build_windows.verify_modern_vc_runtime_payload(app_dir)

            duplicate = pyside_dir / available[0]
            duplicate.write_bytes(b"stale-copy")
            with self.assertRaisesRegex(RuntimeError, "存在重复副本"):
                build_windows.verify_modern_vc_runtime_payload(app_dir)

    def test_modern_runtime_resolution_does_not_import_onnx_or_qt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            runtime_dir = Path(tmp)
            for name in build_windows.MODERN_REQUIRED_VC_RUNTIME_FILES:
                (runtime_dir / name).write_bytes(b"runtime")

            original_import = builtins.__import__

            def guarded_import(name, *args, **kwargs):
                if name == "onnxruntime" or name.startswith("PySide6"):
                    raise AssertionError(f"构建进程不应导入扩展模块：{name}")
                return original_import(name, *args, **kwargs)

            with (
                patch.object(build_windows.sys, "platform", "win32"),
                patch.object(
                    build_windows,
                    "_windows_system_directory",
                    return_value=runtime_dir,
                ),
                patch.object(builtins, "__import__", side_effect=guarded_import),
            ):
                self.assertEqual(
                    build_windows.resolve_modern_vc_runtime_dir(),
                    runtime_dir,
                )

    def test_win7_qt_smoke_uses_one_native_windows_configuration(self) -> None:
        self.assertIs(
            build_windows.WIN7_SOURCE_QT_SMOKE_ENV,
            build_windows.WIN7_PACKAGED_QT_SMOKE_ENV,
        )
        self.assertEqual(
            build_windows.WIN7_SOURCE_QT_SMOKE_ENV["QT_QPA_PLATFORM"],
            "windows",
        )

    def test_win7_pyinstaller_hook_rejects_recursive_host_api_sets(self) -> None:
        fake_dylib = SimpleNamespace(include_library=lambda _name: True)
        with patch.dict(
            "sys.modules",
            {"PyInstaller.depend": SimpleNamespace(dylib=fake_dylib)},
        ):
            runpy.run_path(str(build_windows.WIN7_PYINSTALLER_HOOK))

        self.assertFalse(
            fake_dylib.include_library("C:/Windows/System32/api-ms-win-core-fibers-l1-1-0.dll")
        )
        self.assertFalse(
            fake_dylib.include_library("C:/Windows/System32/ext-ms-win-example-l1-1-0.dll")
        )
        self.assertTrue(fake_dylib.include_library("C:/Windows/System32/KERNEL32.dll"))

    def test_macos_spec_collects_7z_and_embedded_unrar_runtime(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            spec_path = tmp_dir / "HRToolkit.spec"
            build_macos._write_spec(
                spec_path,
                architecture="arm64",
                version=self.version,
                icon_path=tmp_dir / "HRToolkit.icns",
                codesign_identity=None,
                entitlements_file=None,
            )
            spec = spec_path.read_text(encoding="utf-8")
        self.assertIn('collect_all("py7zr")', spec)
        self.assertIn('collect_all("unrar")', spec)
        self.assertIn('copy_metadata("pypdf")', spec)
        self.assertNotIn('copy_metadata("pypdfium2")', spec)
        self.assertNotIn('collect_all("pypdfium2")', spec)
        self.assertIn('"pypdf"', spec)
        self.assertIn("_sevenzip_binaries + _rar_binaries", spec)
        self.assertIn("_sevenzip_hidden + _rar_hidden", spec)
        self.assertNotIn('"distutils"', spec)
        self.assertIn('"tkinter"', spec)
        self.assertIn('"hr_toolkit.gui"', spec)
        self.assertIn('"PIL.ImageTk"', spec)
        self.assertIn('"PIL.AvifImagePlugin"', spec)
        self.assertIn('"PIL._avif"', spec)

    def test_macos_dmg_uses_two_stage_ulmo_compression(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit.app"
            app_dir.mkdir()
            dmg_path = tmp_dir / "output" / f"HRToolkit_{self.version}_arm64.dmg"
            staging_dir = tmp_dir / "work" / "dmg-staging"
            with patch.object(build_macos, "_run") as mocked_run:
                build_macos._create_dmg(
                    app_dir,
                    dmg_path,
                    staging_dir,
                    self.version,
                )

            commands = [call.args[0] for call in mocked_run.call_args_list]
            self.assertEqual(commands[0][0], "ditto")
            self.assertEqual(commands[1][0:2], ["hdiutil", "create"])
            self.assertEqual(commands[1][commands[1].index("-format") + 1], "UDRO")
            self.assertEqual(commands[2][0:2], ["hdiutil", "convert"])
            self.assertEqual(
                commands[2][commands[2].index("-format") + 1],
                verify_macos_bundle.EXPECTED_DMG_FORMAT,
            )
            self.assertNotIn("UDZO", " ".join(" ".join(command) for command in commands))

    def test_macos_dmg_retries_only_transient_hdiutil_resource_busy(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit.app"
            app_dir.mkdir()
            dmg_path = tmp_dir / "output" / f"HRToolkit_{self.version}_x64.dmg"
            staging_dir = tmp_dir / "work" / "dmg-staging"
            uncompressed_dmg = staging_dir.parent / f"{dmg_path.stem}.uncompressed.dmg"
            create_attempts = 0
            partial_file_survived = None

            def fake_run(command, **_kwargs):
                nonlocal create_attempts, partial_file_survived
                if command[0:2] == ["hdiutil", "create"]:
                    create_attempts += 1
                    if create_attempts == 1:
                        uncompressed_dmg.write_bytes(b"partial")
                        raise build_macos.MacBuildError("hdiutil: create failed - Resource busy")
                    partial_file_survived = uncompressed_dmg.exists()
                return subprocess.CompletedProcess(command, 0)

            with (
                patch.object(build_macos, "_run", side_effect=fake_run),
                patch.object(build_macos.time, "sleep") as mocked_sleep,
            ):
                build_macos._create_dmg(
                    app_dir,
                    dmg_path,
                    staging_dir,
                    self.version,
                )

            self.assertEqual(create_attempts, 2)
            self.assertFalse(partial_file_survived)
            mocked_sleep.assert_called_once_with(build_macos.HDIUTIL_BUSY_RETRY_DELAYS[0])

    def test_macos_dmg_does_not_retry_non_transient_hdiutil_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit.app"
            app_dir.mkdir()
            dmg_path = tmp_dir / "output" / f"HRToolkit_{self.version}_x64.dmg"
            staging_dir = tmp_dir / "work" / "dmg-staging"
            create_attempts = 0

            def fake_run(command, **_kwargs):
                nonlocal create_attempts
                if command[0:2] == ["hdiutil", "create"]:
                    create_attempts += 1
                    raise build_macos.MacBuildError("hdiutil: create failed - Permission denied")
                return subprocess.CompletedProcess(command, 0)

            with (
                patch.object(build_macos, "_run", side_effect=fake_run),
                patch.object(build_macos.time, "sleep") as mocked_sleep,
                self.assertRaisesRegex(build_macos.MacBuildError, "Permission denied"),
            ):
                build_macos._create_dmg(
                    app_dir,
                    dmg_path,
                    staging_dir,
                    self.version,
                )

            self.assertEqual(create_attempts, 1)
            mocked_sleep.assert_not_called()

    def test_windows_payload_prunes_only_optional_opencv_video_backend(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            app_dir, _updater = self._fake_app(Path(tmp))
            cv2_dir = app_dir / "_internal" / "cv2"
            cv2_dir.mkdir()
            ffmpeg = cv2_dir / "opencv_videoio_ffmpeg500_64.dll"
            retained = cv2_dir / "opencv_imgcodecs.pyd"
            ffmpeg.write_bytes(b"optional video backend")
            self._write_fake_pe(retained, build_windows.PE_MACHINE_AMD64)
            retained_payload = retained.read_bytes()

            with self.assertRaisesRegex(RuntimeError, "OpenCV 视频后端"):
                build_windows.verify_windows_payload(app_dir)
            removed = build_windows.remove_unused_opencv_videoio_ffmpeg(app_dir)

            self.assertEqual(removed, len(b"optional video backend"))
            self.assertFalse(ffmpeg.exists())
            self.assertEqual(retained.read_bytes(), retained_payload)
            build_windows.verify_windows_payload(app_dir)

    def test_packaging_prunes_only_unused_qt_translations(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            app_dir, _updater = self._fake_app(root / "windows")
            translations = (
                app_dir
                / "_internal"
                / build_windows.QT6_WINDOWS_RUNTIME_ROOT
                / "translations"
            )
            translations.mkdir(parents=True)
            retained = {
                "qt_en.qm": b"english",
                "qtbase_zh_CN.qm": b"simplified",
                "qtbase_zh_TW.qm": b"traditional",
            }
            for name, payload in retained.items():
                (translations / name).write_bytes(payload)
            removed = translations / "qtbase_de.qm"
            removed.write_bytes(b"unused-german")

            with self.assertRaisesRegex(RuntimeError, "未使用的 modern Qt 翻译"):
                build_windows.verify_windows_payload(app_dir)
            removed_bytes = build_windows.remove_unused_qt_translations(app_dir)
            self.assertEqual(removed_bytes, len(b"unused-german"))
            self.assertFalse(removed.exists())
            for name, payload in retained.items():
                self.assertEqual((translations / name).read_bytes(), payload)
            build_windows.verify_windows_payload(app_dir)

            mac_app = root / "HRToolkit.app"
            mac_translations = (
                mac_app
                / "Contents"
                / "Resources"
                / "PySide6"
                / "Qt"
                / "translations"
            )
            mac_translations.mkdir(parents=True)
            for name, payload in retained.items():
                (mac_translations / name).write_bytes(payload)
            mac_removed = mac_translations / "qtbase_fr.qm"
            mac_removed.write_bytes(b"unused-french")
            self.assertEqual(
                build_macos.remove_unused_qt_translations(mac_app),
                len(b"unused-french"),
            )
            self.assertFalse(mac_removed.exists())
            for name, payload in retained.items():
                self.assertEqual((mac_translations / name).read_bytes(), payload)

            windows_tooling = (
                app_dir
                / "_internal"
                / build_windows.QT6_WINDOWS_RUNTIME_ROOT
                / "plugins"
                / "qmltooling"
            )
            windows_tooling.mkdir(parents=True)
            (windows_tooling / "qmldbg.dll").write_bytes(b"development-plugin")
            with self.assertRaisesRegex(RuntimeError, "Qt QML 开发插件"):
                build_windows.verify_windows_payload(app_dir)
            self.assertEqual(
                build_windows.remove_qt_development_plugins(app_dir),
                len(b"development-plugin"),
            )
            self.assertFalse(windows_tooling.exists())
            build_windows.verify_windows_payload(app_dir)

            mac_tooling = (
                mac_app
                / "Contents"
                / "Frameworks"
                / "PySide6"
                / "Qt"
                / "plugins"
                / "qmltooling"
            )
            mac_tooling.mkdir(parents=True)
            (mac_tooling / "libqmldbg.dylib").write_bytes(b"development-plugin-mac")
            self.assertEqual(
                build_macos.remove_qt_development_plugins(mac_app),
                len(b"development-plugin-mac"),
            )
            self.assertFalse(mac_tooling.exists())

            windows_images = (
                app_dir
                / "_internal"
                / build_windows.QT6_WINDOWS_RUNTIME_ROOT
                / "plugins"
                / "imageformats"
            )
            windows_images.mkdir(parents=True)
            windows_retained = windows_images / "qjpeg.dll"
            windows_removed = windows_images / "qtiff.dll"
            self._write_fake_pe(windows_retained, build_windows.PE_MACHINE_AMD64)
            self._write_fake_pe(windows_removed, build_windows.PE_MACHINE_AMD64)
            retained_payload = windows_retained.read_bytes()
            removed_size = windows_removed.stat().st_size
            with self.assertRaisesRegex(RuntimeError, "Qt 图片格式插件"):
                build_windows.verify_windows_payload(app_dir)
            self.assertEqual(
                build_windows.remove_unused_qt_image_format_plugins(app_dir),
                removed_size,
            )
            self.assertEqual(windows_retained.read_bytes(), retained_payload)
            self.assertFalse(windows_removed.exists())
            build_windows.verify_windows_payload(app_dir)

            pillow_dir = app_dir / "_internal" / "PIL"
            pillow_dir.mkdir(parents=True)
            avif_runtime = pillow_dir / "libavif.dll"
            avif_runtime.write_bytes(b"unsupported-avif")
            with self.assertRaisesRegex(RuntimeError, "Pillow AVIF"):
                build_windows.verify_windows_payload(app_dir)
            avif_runtime.unlink()
            build_windows.verify_windows_payload(app_dir)

            mac_images = mac_tooling.parent / "imageformats"
            mac_images.mkdir(parents=True)
            mac_retained = mac_images / "libqwebp.dylib"
            mac_removed = mac_images / "libqtiff.dylib"
            mac_retained.write_bytes(b"retained-webp")
            mac_removed.write_bytes(b"unused-tiff")
            self.assertEqual(
                build_macos.remove_unused_qt_image_format_plugins(mac_app),
                len(b"unused-tiff"),
            )
            self.assertEqual(mac_retained.read_bytes(), b"retained-webp")
            self.assertFalse(mac_removed.exists())

    def test_compact_intel_macos_numpy_runtime_is_strictly_gated(self) -> None:
        compact_configuration = {
            "Build Dependencies": {
                "blas": {"name": "none"},
                "lapack": {"detection method": "internal"},
            }
        }
        self.assertTrue(
            prepare_macos_x64_runtime.configuration_uses_internal_fallback(
                compact_configuration
            )
        )
        external_configuration = {
            "Build Dependencies": {
                "blas": {"name": "openblas"},
                "lapack": {"detection method": "pkgconfig"},
            }
        }
        self.assertFalse(
            prepare_macos_x64_runtime.configuration_uses_internal_fallback(
                external_configuration
            )
        )
        with tempfile.TemporaryDirectory() as tmp:
            wheel = Path(tmp) / prepare_macos_x64_runtime.EXPECTED_WHEEL_NAME
            wheel.write_bytes(b"compact-wheel")
            prepare_macos_x64_runtime.validate_compact_wheel(wheel)
            invalid = Path(tmp) / "numpy-1.26.4-cp312-cp312-macosx_12_0_x86_64.whl"
            invalid.write_bytes(b"wrong-floor")
            with self.assertRaisesRegex(
                prepare_macos_x64_runtime.MacX64RuntimeError,
                "最低 macOS 版本",
            ):
                prepare_macos_x64_runtime.validate_compact_wheel(invalid)

    def test_win7_frozen_smoke_forces_bundled_7zip_without_changing_modern(self) -> None:
        observed_calls: list[tuple[list[str], dict[str, str]]] = []

        def fake_run(command: list[str], *, timeout=None, env=None) -> None:
            del timeout
            self.assertIsNotNone(env)
            assert env is not None
            observed_calls.append((list(command), dict(env)))
            output_path = Path(env["HR_TOOLKIT_CHECK_OUTPUT"])
            if "--version" in command:
                output_path.write_text(self.version, encoding="utf-8")
            elif Path(command[0]).name == "HRToolkitUpdater.exe":
                for runtime_name in (
                    *build_windows.WIN7_REQUIRED_UCRT_FILES,
                    *build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES,
                ):
                    self.assertTrue((Path(command[0]).parent / runtime_name).is_file())
                output_path.write_text(
                    f"HRToolkitUpdater {self.version} smoke-test OK",
                    encoding="utf-8",
                )
            elif "--smoke-test" in command:
                output_path.write_text(
                    f"HRToolkit {self.version} smoke-test OK",
                    encoding="utf-8",
                )
            elif "--qt-smoke-test" in command:
                output_path.write_text("HRToolkit Qt smoke-test OK", encoding="utf-8")
            else:
                output_path.write_text(
                    f"HRToolkit {self.version} update-smoke-test OK; latest={self.version}",
                    encoding="utf-8",
                )

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app = tmp_dir / "HRToolkit.exe"
            updater = tmp_dir / "HRToolkitUpdater.exe"
            self._write_fake_pe(updater, build_windows.PE_MACHINE_AMD64)
            for runtime_name in (
                *build_windows.WIN7_REQUIRED_UCRT_FILES,
                *build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES,
            ):
                (tmp_dir / runtime_name).write_bytes(b"runtime")
            with (
                patch.dict(
                    build_windows.os.environ,
                    {
                        build_windows.WIN7_7ZIP_OVERRIDE_ENV: "C:/build-runtime/7z.exe",
                        "QT_QPA_PLATFORM": "windows",
                        "QT_QUICK_BACKEND": "caller-selected",
                        "QSG_RENDER_LOOP": "threaded",
                    },
                ),
                patch.object(build_windows, "_run", side_effect=fake_run),
            ):
                build_windows.run_runtime_smoke(
                    app,
                    updater,
                    target=build_windows.WINDOWS_TARGET_WIN7,
                )
                win7_calls = list(observed_calls)
                observed_calls.clear()
                build_windows.run_runtime_smoke(app, updater)
                modern_calls = list(observed_calls)

        self.assertEqual(len(win7_calls), 7)
        self.assertEqual(
            sum("--qt-smoke-test" in command for command, _ in win7_calls),
            3,
        )
        self.assertTrue(
            all(
                build_windows.WIN7_7ZIP_OVERRIDE_ENV not in env
                for _, env in win7_calls
            )
        )
        win7_qt_env = next(
            env for command, env in win7_calls if "--qt-smoke-test" in command
        )
        for name, value in build_windows.WIN7_PACKAGED_QT_SMOKE_ENV.items():
            self.assertEqual(win7_qt_env[name], value)

        self.assertEqual(len(modern_calls), 4)
        self.assertTrue(
            all(
                env[build_windows.WIN7_7ZIP_OVERRIDE_ENV]
                == "C:/build-runtime/7z.exe"
                for _, env in modern_calls
            )
        )
        modern_qt_env = next(
            env for command, env in modern_calls if "--qt-smoke-test" in command
        )
        self.assertEqual(modern_qt_env["QT_QPA_PLATFORM"], "windows")
        self.assertEqual(modern_qt_env["QT_QUICK_BACKEND"], "caller-selected")
        self.assertEqual(modern_qt_env["QSG_RENDER_LOOP"], "threaded")

    def test_windows_version_metadata_uses_requested_version(self) -> None:
        payload = build_windows.windows_version_info("0.2.1")
        self.assertIn("filevers=(0, 2, 1, 0)", payload)
        self.assertIn("StringStruct('ProductVersion', '0.2.1')", payload)

    def test_packaged_smoke_timeout_reports_the_last_runtime_stage(self) -> None:
        def fake_run(command: list[str], *, timeout=None, env=None) -> None:
            self.assertIsNotNone(env)
            assert env is not None
            output_path = Path(env["HR_TOOLKIT_CHECK_OUTPUT"])
            if "--version" in command:
                output_path.write_text(self.version, encoding="utf-8")
                return
            output_path.write_text(
                f"HRToolkit {self.version} smoke-test RUNNING: ocr-inference\n",
                encoding="utf-8",
            )
            raise subprocess.TimeoutExpired(command, timeout)

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app = tmp_dir / "HRToolkit.exe"
            updater = tmp_dir / "HRToolkitUpdater.exe"
            self._write_fake_pe(updater, build_windows.PE_MACHINE_AMD64)
            with patch.object(build_windows, "_run", side_effect=fake_run):
                with self.assertRaisesRegex(RuntimeError, "ocr-inference"):
                    build_windows.run_runtime_smoke(app, updater)

    def test_packaged_qt_crash_preserves_native_traceback_and_does_not_retry(self) -> None:
        qt_launches = []

        def fake_run(command, *, timeout=None, env=None):
            output = Path(env["HR_TOOLKIT_CHECK_OUTPUT"])
            if "--version" in command:
                output.write_text(self.version, encoding="utf-8")
            elif "--smoke-test" in command:
                output.write_text(
                    f"HRToolkit {self.version} smoke-test OK", encoding="utf-8"
                )
            elif "--update-smoke-test" in command:
                output.write_text(
                    f"HRToolkit {self.version} update-smoke-test OK; latest={self.version}",
                    encoding="utf-8",
                )
            else:
                qt_launches.append(command)
                output.write_text(
                    "HRToolkit Qt smoke-test RUNNING: qt-qml-load", encoding="utf-8"
                )
                Path(str(output) + ".native.log").write_text(
                    "Windows fatal exception: access violation\nmain.py: engine.load",
                    encoding="utf-8",
                )
                raise subprocess.CalledProcessError(0xC0000005, command)

        with tempfile.TemporaryDirectory() as tmp:
            app = Path(tmp) / "HRToolkit.exe"
            updater = Path(tmp) / "HRToolkitUpdater.exe"
            self._write_fake_pe(updater, build_windows.PE_MACHINE_AMD64)
            with patch.object(build_windows, "_run", side_effect=fake_run):
                with self.assertRaises(RuntimeError) as raised:
                    build_windows.run_runtime_smoke(app, updater, target="win7")
        self.assertEqual(len(qt_launches), 1)
        self.assertIn("qt-qml-load", str(raised.exception))
        self.assertIn("access violation", str(raised.exception))
        self.assertIn("engine.load", str(raised.exception))

    def test_wix_xml_indentation_has_a_python38_fallback(self) -> None:
        root = ET.Element("root")
        child = ET.SubElement(root, "child")
        ET.SubElement(child, "leaf")
        modern_tree = ET.ElementTree(copy.deepcopy(root))
        fallback_tree = ET.ElementTree(copy.deepcopy(root))

        build_windows_installers._indent_xml(modern_tree)
        with patch.object(build_windows_installers.ET, "indent", None):
            build_windows_installers._indent_xml(fallback_tree)

        self.assertEqual(
            ET.tostring(fallback_tree.getroot()),
            ET.tostring(modern_tree.getroot()),
        )

    def test_windows_release_job_forces_utf8_python_output(self) -> None:
        workflow = (build_windows.REPO_ROOT / ".github" / "workflows" / "release.yml").read_text(
            encoding="utf-8"
        )
        windows_job = workflow.split("\n  build-windows:", 1)[1].split(
            "\n  build-macos:", 1
        )[0]
        job_configuration = windows_job.split("\n    steps:", 1)[0]
        self.assertIn('PYTHONUTF8: "1"', job_configuration)

    def test_ci_actions_are_immutable_and_production_dependencies_are_locked(self) -> None:
        workflow_dir = build_windows.REPO_ROOT / ".github" / "workflows"
        workflow_paths = sorted(workflow_dir.glob("*.yml"))
        action_lines: list[str] = []
        for workflow_path in workflow_paths:
            workflow = workflow_path.read_text(encoding="utf-8")
            action_lines.extend(
                line.strip() for line in workflow.splitlines() if "uses: actions/" in line
            )
        self.assertTrue(action_lines)
        for line in action_lines:
            with self.subTest(line=line):
                self.assertRegex(
                    line,
                    r"^uses: actions/[a-z-]+@[0-9a-f]{40} # v\d+(?:\.\d+){0,2}$",
                )

        constraint_path = (
            build_windows.REPO_ROOT / "constraints" / "python312-production.txt"
        )
        constraints = [
            line.strip()
            for line in constraint_path.read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        ]
        self.assertIn("pip==26.2", constraints)
        self.assertIn("rapidocr_onnxruntime==1.4.4", constraints)
        self.assertIn("pypdf==6.16.1", constraints)
        self.assertIn("pyinstaller==6.21.0", constraints)
        self.assertIn(
            'onnxruntime==1.17.3; platform_system == "Darwin" and platform_machine == "x86_64"',
            constraints,
        )
        self.assertIn(
            'onnxruntime==1.29.0; platform_system != "Darwin" or platform_machine != "x86_64"',
            constraints,
        )
        self.assertIn(
            'opencv-python==4.5.5.64; platform_system == "Darwin" and platform_machine == "x86_64"',
            constraints,
        )
        self.assertIn(
            'opencv-python==4.10.0.84; platform_system != "Darwin" or platform_machine != "x86_64"',
            constraints,
        )
        for constraint in constraints:
            with self.subTest(constraint=constraint):
                requirement = constraint.split(";", 1)[0].strip()
                self.assertRegex(requirement, r"^[A-Za-z0-9_.-]+==[^=\s]+$")

        ci = (workflow_dir / "ci.yml").read_text(encoding="utf-8")
        self.assertIn("requirements-audit.txt", ci)
        self.assertIn("python -m pip_audit --local", ci)
        self.assertIn("constraints/python312-production.txt", ci)
        self.assertIn("constraints/python38-win7.txt", ci)
        self.assertIn('python-version: "3.8.10"', ci)
        modern_ci = ci.split("\n  windows-test:", 1)[1].split(
            "\n  windows-win7-compat-test:", 1
        )[0]
        self.assertIn("runs-on: windows-latest", modern_ci)
        win7_ci = ci.split("\n  windows-win7-compat-test:", 1)[1]
        self.assertIn("runs-on: windows-2022", win7_ci)
        for name, value in build_windows.WIN7_SOURCE_QT_SMOKE_ENV.items():
            self.assertIn(f'{name}: "{value}"', win7_ci)

        win7_constraints = (
            build_windows.REPO_ROOT / "constraints" / "python38-win7.txt"
        ).read_text(encoding="utf-8")
        self.assertIn("onnxruntime==1.11.1", win7_constraints)
        self.assertIn("opencv-python==4.8.1.78", win7_constraints)
        self.assertIn("pypdfium2==4.27.0", win7_constraints)
        self.assertNotIn("pypdf==", win7_constraints)
        self.assertIn("pyinstaller==6.21.0", win7_constraints)

    def test_scheduled_windows_package_gate_runs_tests_and_ocr(self) -> None:
        workflow = (
            build_windows.REPO_ROOT / ".github" / "workflows" / "test-build.yml"
        ).read_text(encoding="utf-8")
        self.assertIn('- cron: "0 18 * * 0"', workflow)
        self.assertIn('TARGET_PLATFORM="windows-all"', workflow)
        self.assertIn('RUN_TESTS="true"', workflow)
        self.assertIn("ocr_runtime_smoke_test", workflow)
        self.assertIn("scripts/release_windows.py", workflow)
        self.assertIn("--target win7", workflow)
        self.assertIn("prepare_win7_runtime.py", workflow)
        modern_job = workflow.split("\n  build-windows:", 1)[1].split(
            "\n  build-windows-win7:", 1
        )[0]
        self.assertIn("runs-on: windows-latest", modern_job)
        win7_job = workflow.split("\n  build-windows-win7:", 1)[1].split(
            "\n  build-macos:", 1
        )[0]
        self.assertIn("runs-on: windows-2022", win7_job)

    def test_release_win7_build_uses_stable_windows_2022_runner(self) -> None:
        workflow = (
            build_windows.REPO_ROOT / ".github" / "workflows" / "release.yml"
        ).read_text(encoding="utf-8")
        modern_job = workflow.split("\n  build-windows:", 1)[1].split(
            "\n  build-windows-win7:", 1
        )[0]
        win7_job = workflow.split("\n  build-windows-win7:", 1)[1].split(
            "\n  build-macos:", 1
        )[0]

        self.assertIn("runs-on: windows-latest", modern_job)
        self.assertIn("runs-on: windows-2022", win7_job)

    def test_release_and_test_build_use_real_parallel_macos_architectures(self) -> None:
        workflow_dir = build_windows.REPO_ROOT / ".github" / "workflows"
        release = (workflow_dir / "release.yml").read_text(encoding="utf-8")
        test_build = (workflow_dir / "test-build.yml").read_text(encoding="utf-8")

        for workflow in (release, test_build):
            with self.subTest(workflow=workflow[:30]):
                self.assertIn("runner: macos-15-intel", workflow)
                self.assertIn("architecture: x86_64", workflow)
                self.assertIn("runner: macos-15", workflow)
                self.assertIn("architecture: arm64", workflow)
                self.assertIn('test "$(uname -m)" = "${{ matrix.architecture }}"', workflow)
                self.assertIn("if: matrix.architecture == 'x86_64'", workflow)
                self.assertIn("python scripts/prepare_macos_x64_runtime.py", workflow)

        self.assertIn("needs.build-macos.result == 'success'", release)
        self.assertNotIn("build-macos-universal", release)
        self.assertNotIn("build-macos-fallback", release)
        self.assertNotIn("--architecture universal2", release)

    def test_payload_verification_accepts_only_readme_and_builtin_excel_templates(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            app_dir, _updater = self._fake_app(Path(tmp))
            build_windows.verify_windows_payload(app_dir)

            forbidden = app_dir / "真实工资表.xlsx"
            forbidden.write_bytes(b"private")
            with self.assertRaisesRegex(RuntimeError, "模板目录之外"):
                build_windows.verify_windows_payload(app_dir)
            forbidden.unlink()

            project_db = app_dir / "_internal" / "project.db"
            project_db.write_bytes(b"private project index")
            with self.assertRaisesRegex(RuntimeError, "用户项目数据"):
                build_windows.verify_windows_payload(app_dir)
            project_db.unlink()

            project_metadata = app_dir / "_internal" / ".hrtoolkit"
            for relative in (
                Path("project.json"),
                Path("index.db"),
                Path("manifests") / "batch.json",
                Path("project-write.lock"),
            ):
                project_artifact = project_metadata / relative
                project_artifact.parent.mkdir(parents=True, exist_ok=True)
                project_artifact.write_bytes(b"private project metadata")
                with self.assertRaisesRegex(RuntimeError, "禁止目录"):
                    build_windows.verify_windows_payload(app_dir)
                project_artifact.unlink()

            uploaded_zip = app_dir / "_internal" / "上传资料" / "薪酬管理" / "工资表拆分" / "8月工资" / "工资.zip"
            uploaded_zip.parent.mkdir(parents=True)
            uploaded_zip.write_bytes(b"private upload")
            with self.assertRaisesRegex(RuntimeError, "禁止目录或缓存"):
                build_windows.verify_windows_payload(app_dir)
            uploaded_zip.unlink()

            result_pdf = app_dir / "_internal" / "处理结果" / "社保与保险" / "保险台账" / "8月" / "结果.pdf"
            result_pdf.parent.mkdir(parents=True)
            result_pdf.write_bytes(b"private result")
            with self.assertRaisesRegex(RuntimeError, "禁止目录或缓存"):
                build_windows.verify_windows_payload(app_dir)
            result_pdf.unlink()

            supplement = app_dir / "_internal" / "补充资料" / "人员与档案" / "说明.docx"
            supplement.parent.mkdir(parents=True)
            supplement.write_bytes(b"private supplement")
            with self.assertRaisesRegex(RuntimeError, "禁止目录或缓存"):
                build_windows.verify_windows_payload(app_dir)
            supplement.unlink()

            common_file = app_dir / "_internal" / "共用资料" / "员工花名册.xlsx"
            common_file.parent.mkdir(parents=True)
            common_file.write_bytes(b"private common material")
            with self.assertRaisesRegex(RuntimeError, "禁止目录或缓存"):
                build_windows.verify_windows_payload(app_dir)
            common_file.unlink()

            # 一般依赖可以合法使用这些通用名称；只有项目专用
            # `.hrtoolkit` 和可见的上传/结果目录才是数据泄漏信号。
            dependency_manifest = app_dir / "_internal" / "dependency" / "records" / "trash" / "manifest.json"
            dependency_manifest.parent.mkdir(parents=True)
            dependency_manifest.write_text("{}", encoding="utf-8")
            build_windows.verify_windows_payload(app_dir)
            dependency_manifest.unlink()

            cache = app_dir / "_internal" / "__pycache__" / "module.pyc"
            cache.parent.mkdir()
            cache.write_bytes(b"cache")
            with self.assertRaisesRegex(RuntimeError, "禁止目录或缓存"):
                build_windows.verify_windows_payload(app_dir)

    def test_modern_payload_requires_one_root_vc_runtime_set(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            app_dir, _updater = self._fake_app(Path(tmp))
            build_windows.verify_windows_payload(app_dir)

            missing = app_dir / build_windows.MODERN_REQUIRED_VC_RUNTIME_FILES[0]
            missing.unlink()
            with self.assertRaisesRegex(RuntimeError, r"Visual C\+\+ runtime"):
                build_windows.verify_windows_payload(app_dir)

    def test_macos_resource_verifier_rejects_project_metadata_without_blocking_generic_names(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            app_dir = Path(tmp) / "HRToolkit.app"
            resources = app_dir / "Contents" / "Resources"
            templates = resources / "hr_toolkit" / "templates"
            templates.mkdir(parents=True)
            for template in verify_macos_bundle.DEFAULT_TEMPLATE_DIR.glob("*.xlsx"):
                (templates / template.name).write_bytes(template.read_bytes())
            (resources / "README.md").write_bytes(verify_macos_bundle.DEFAULT_README.read_bytes())
            qml_root = resources / "hr_toolkit" / "gui_qt" / "qml"
            for source in verify_macos_bundle.DEFAULT_QML_DIR.rglob("*.qml"):
                target = qml_root / source.relative_to(verify_macos_bundle.DEFAULT_QML_DIR)
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(source.read_bytes())
            qt_notice = resources / "third_party" / "qt" / verify_macos_bundle.DEFAULT_QT_NOTICE.name
            qt_notice.parent.mkdir(parents=True, exist_ok=True)
            qt_notice.write_bytes(verify_macos_bundle.DEFAULT_QT_NOTICE.read_bytes())
            qt_qml_root = resources / "PySide6" / "Qt" / "qml"
            for relative in verify_macos_bundle.QT6_REQUIRED_QML_FILES:
                path = qt_qml_root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(b"qt-qml")

            dependency_manifest = resources / "dependency" / "records" / "trash" / "manifest.json"
            dependency_manifest.parent.mkdir(parents=True)
            dependency_manifest.write_text("{}", encoding="utf-8")
            verify_macos_bundle.verify_packaged_resources(app_dir)

            project_marker = resources / ".hrtoolkit" / "project.json"
            project_marker.parent.mkdir()
            project_marker.write_text("private", encoding="utf-8")
            with self.assertRaisesRegex(verify_macos_bundle.MacBundleVerificationError, "禁止目录"):
                verify_macos_bundle.verify_packaged_resources(app_dir)

            project_marker.unlink()
            project_marker.parent.rmdir()
            common_file = resources / "共用资料" / "员工花名册.xlsx"
            common_file.parent.mkdir()
            common_file.write_text("private", encoding="utf-8")
            with self.assertRaisesRegex(verify_macos_bundle.MacBundleVerificationError, "禁止目录"):
                verify_macos_bundle.verify_packaged_resources(app_dir)

    def test_pe_machine_verification_requires_amd64(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            executable = Path(tmp) / "app.exe"
            self._write_fake_pe(executable, build_windows.PE_MACHINE_AMD64)
            self.assertEqual(build_windows.read_pe_machine(executable), build_windows.PE_MACHINE_AMD64)
            build_windows.verify_pe_x64(executable)

            self._write_fake_pe(executable, 0x014C)
            with self.assertRaisesRegex(RuntimeError, "不是 x64 PE"):
                build_windows.verify_pe_x64(executable)

    def test_win7_payload_requires_python38_ucrt_and_bundled_7zip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            app_dir, _updater = self._fake_app(
                Path(tmp),
                target=build_windows.WINDOWS_TARGET_WIN7,
            )
            internal = app_dir / "_internal"
            self._write_fake_pe(internal / "python38.dll", build_windows.PE_MACHINE_AMD64)
            self._write_fake_pe(internal / "python3.dll", build_windows.PE_MACHINE_AMD64)
            for name in build_windows.WIN7_REQUIRED_UCRT_FILES:
                self._write_fake_pe(app_dir / name, build_windows.PE_MACHINE_AMD64)
            for name in build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES:
                self._write_fake_pe(app_dir / name, build_windows.PE_MACHINE_AMD64)
            seven_zip = internal / "third_party" / "7zip"
            self._write_fake_pe(seven_zip / "7z.exe", build_windows.PE_MACHINE_AMD64)
            self._write_fake_pe(seven_zip / "7z.dll", build_windows.PE_MACHINE_AMD64)
            (seven_zip / "License.txt").write_text("license", encoding="utf-8")
            (seven_zip / build_windows.WIN7_THIRD_PARTY_NOTICE.name).write_bytes(
                build_windows.WIN7_THIRD_PARTY_NOTICE.read_bytes()
            )

            build_windows.verify_windows_payload(
                app_dir,
                target=build_windows.WINDOWS_TARGET_WIN7,
            )
            forbidden = internal / "api-ms-win-core-path-l1-1-0.dll"
            self._write_fake_pe(forbidden, build_windows.PE_MACHINE_AMD64)
            with self.assertRaisesRegex(RuntimeError, "旁加载伪造"):
                build_windows.verify_windows_payload(
                    app_dir,
                    target=build_windows.WINDOWS_TARGET_WIN7,
                )
            forbidden.unlink()
            unpinned_api_set = internal / "api-ms-win-core-fibers-l1-1-0.dll"
            self._write_fake_pe(unpinned_api_set, build_windows.PE_MACHINE_AMD64)
            with self.assertRaisesRegex(RuntimeError, "未锁定的系统 API Set"):
                build_windows.verify_windows_payload(
                    app_dir,
                    target=build_windows.WINDOWS_TARGET_WIN7,
                )
            unpinned_api_set.unlink()
            self._write_fake_pe(internal / "python39.dll", build_windows.PE_MACHINE_AMD64)
            with self.assertRaisesRegex(RuntimeError, "其他 Python 运行时"):
                build_windows.verify_windows_payload(
                    app_dir,
                    target=build_windows.WINDOWS_TARGET_WIN7,
                )
            (internal / "python39.dll").unlink()
            (internal / "python38.dll").unlink()
            with self.assertRaisesRegex(RuntimeError, "python38.dll"):
                build_windows.verify_windows_payload(
                    app_dir,
                    target=build_windows.WINDOWS_TARGET_WIN7,
                )
            self._write_fake_pe(internal / "python38.dll", build_windows.PE_MACHINE_AMD64)
            (internal / "python3.dll").unlink()
            with self.assertRaisesRegex(RuntimeError, "python3.dll"):
                build_windows.verify_windows_payload(
                    app_dir,
                    target=build_windows.WINDOWS_TARGET_WIN7,
                )

    def test_win7_runtime_integrity_rejects_pyinstaller_substitution(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir, updater = self._fake_app(tmp_dir)
            internal = app_dir / "_internal"
            seven_zip_dir = tmp_dir / "sources" / "7zip"
            ucrt_dir = tmp_dir / "sources" / "ucrt"
            vc_runtime_dir = tmp_dir / "sources" / "vc-runtime"
            archive_payloads: dict[str, bytes] = {}

            for label, source_dir, payload_dir, names in (
                (
                    "7zip",
                    seven_zip_dir,
                    internal / "third_party" / "7zip",
                    build_windows.WIN7_REQUIRED_7ZIP_FILES,
                ),
                (
                    "ucrt",
                    ucrt_dir,
                    app_dir,
                    build_windows.WIN7_REQUIRED_UCRT_FILES,
                ),
                (
                    "vc",
                    vc_runtime_dir,
                    app_dir,
                    build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES,
                ),
            ):
                source_dir.mkdir(parents=True, exist_ok=True)
                payload_dir.mkdir(parents=True, exist_ok=True)
                for name in names:
                    payload = f"{label}:{name}".encode("utf-8")
                    (source_dir / name).write_bytes(payload)
                    (payload_dir / name).write_bytes(payload)
                    if label in {"ucrt", "vc"}:
                        archive_payloads[name] = payload

            class FakeArchiveReader:
                def __init__(self, _path: str):
                    self.toc = {name: object() for name in archive_payloads}

                def extract(self, name: str) -> bytes:
                    return archive_payloads[name]

            kwargs = {
                "app_dir": app_dir,
                "updater": updater,
                "seven_zip_dir": seven_zip_dir,
                "ucrt_dir": ucrt_dir,
                "vc_runtime_dir": vc_runtime_dir,
                "archive_reader_cls": FakeArchiveReader,
            }
            build_windows.verify_win7_runtime_source_integrity(**kwargs)

            archive_payloads["api-ms-win-core-fibers-l1-1-0.dll"] = b"modern host forwarder"
            with self.assertRaisesRegex(RuntimeError, "Updater 内嵌了未锁定"):
                build_windows.verify_win7_runtime_source_integrity(**kwargs)
            del archive_payloads["api-ms-win-core-fibers-l1-1-0.dll"]

            substituted = app_dir / build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES[0]
            original = substituted.read_bytes()
            substituted.write_bytes(b"newer system runtime")
            with self.assertRaisesRegex(RuntimeError, "未使用已锁定"):
                build_windows.verify_win7_runtime_source_integrity(**kwargs)
            substituted.write_bytes(original)

            runtime_name = build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES[0]
            archive_payloads[runtime_name] = b"newer embedded runtime"
            with self.assertRaisesRegex(RuntimeError, "Updater 未使用已锁定"):
                build_windows.verify_win7_runtime_source_integrity(**kwargs)

    def test_win7_pe_gate_rejects_post_win7_imports(self) -> None:
        imported = SimpleNamespace(name=b"PssQuerySnapshot")
        entry = SimpleNamespace(dll=b"KERNEL32.dll", imports=(imported,))

        class FakePE:
            OPTIONAL_HEADER = SimpleNamespace(
                MajorSubsystemVersion=6,
                MinorSubsystemVersion=1,
            )
            DIRECTORY_ENTRY_IMPORT = (entry,)
            DIRECTORY_ENTRY_DELAY_IMPORT = ()

            def parse_data_directories(self, *, directories):
                self.directories = directories

            def close(self):
                pass

        fake_pefile = SimpleNamespace(
            DIRECTORY_ENTRY={
                "IMAGE_DIRECTORY_ENTRY_IMPORT": 1,
                "IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT": 2,
            },
            PE=lambda _path, fast_load: FakePE(),
            PEFormatError=ValueError,
        )
        with patch.dict("sys.modules", {"pefile": fake_pefile}):
            with self.assertRaisesRegex(RuntimeError, "PssQuerySnapshot"):
                build_windows.verify_win7_pe_compatibility((Path("runtime.dll"),))

    def test_win7_pe_gate_rejects_unbundled_api_sets(self) -> None:
        entry = SimpleNamespace(
            dll=b"api-ms-win-core-future-l1-1-0.dll",
            imports=(),
        )

        class FakePE:
            OPTIONAL_HEADER = SimpleNamespace(
                MajorSubsystemVersion=6,
                MinorSubsystemVersion=1,
            )
            DIRECTORY_ENTRY_IMPORT = (entry,)
            DIRECTORY_ENTRY_DELAY_IMPORT = ()

            def parse_data_directories(self, *, directories):
                self.directories = directories

            def close(self):
                pass

        fake_pefile = SimpleNamespace(
            DIRECTORY_ENTRY={
                "IMAGE_DIRECTORY_ENTRY_IMPORT": 1,
                "IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT": 2,
            },
            PE=lambda _path, fast_load: FakePE(),
            PEFormatError=ValueError,
        )
        with patch.dict("sys.modules", {"pefile": fake_pefile}):
            with self.assertRaisesRegex(RuntimeError, "未随包提供"):
                build_windows.verify_win7_pe_compatibility((Path("runtime.dll"),))
            build_windows.verify_win7_pe_compatibility(
                (
                    Path("runtime.dll"),
                    Path("api-ms-win-core-future-l1-1-0.dll"),
                )
            )

    def test_win7_pe_gate_requires_app_local_vc_dependencies(self) -> None:
        entry = SimpleNamespace(
            dll=b"MSVCP140.dll",
            imports=(SimpleNamespace(name=b"?required@@", ordinal=None),),
        )

        class FakePE:
            OPTIONAL_HEADER = SimpleNamespace(
                MajorSubsystemVersion=6,
                MinorSubsystemVersion=1,
            )
            DIRECTORY_ENTRY_IMPORT = (entry,)
            DIRECTORY_ENTRY_DELAY_IMPORT = ()

            def parse_data_directories(self, *, directories):
                self.directories = directories

            def close(self):
                pass

        fake_pefile = SimpleNamespace(
            DIRECTORY_ENTRY={
                "IMAGE_DIRECTORY_ENTRY_IMPORT": 1,
                "IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT": 2,
            },
            PE=lambda _path, fast_load: FakePE(),
            PEFormatError=ValueError,
        )
        with patch.dict("sys.modules", {"pefile": fake_pefile}):
            with self.assertRaisesRegex(RuntimeError, "缺少 app-local MSVCP140.dll"):
                build_windows.verify_win7_pe_compatibility((Path("extension.pyd"),))

    def test_win7_pe_gate_checks_symbols_against_pinned_vc_runtime(self) -> None:
        required_symbol = b"?required@@"

        class FakePE:
            OPTIONAL_HEADER = SimpleNamespace(
                MajorSubsystemVersion=6,
                MinorSubsystemVersion=1,
            )
            DIRECTORY_ENTRY_DELAY_IMPORT = ()

            def __init__(self, path: str):
                if Path(path).name.casefold() == "msvcp140.dll":
                    self.DIRECTORY_ENTRY_IMPORT = ()
                    self.DIRECTORY_ENTRY_EXPORT = SimpleNamespace(
                        symbols=(SimpleNamespace(name=b"?different@@", ordinal=1),)
                    )
                else:
                    self.DIRECTORY_ENTRY_IMPORT = (
                        SimpleNamespace(
                            dll=b"MSVCP140.dll",
                            imports=(SimpleNamespace(name=required_symbol, ordinal=None),),
                        ),
                    )

            def parse_data_directories(self, *, directories):
                self.directories = directories

            def close(self):
                pass

        fake_pefile = SimpleNamespace(
            DIRECTORY_ENTRY={
                "IMAGE_DIRECTORY_ENTRY_IMPORT": 1,
                "IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT": 2,
                "IMAGE_DIRECTORY_ENTRY_EXPORT": 3,
            },
            PE=lambda path, fast_load: FakePE(path),
            PEFormatError=ValueError,
        )
        with patch.dict("sys.modules", {"pefile": fake_pefile}):
            with self.assertRaisesRegex(RuntimeError, "不导出"):
                build_windows.verify_win7_pe_compatibility(
                    (Path("extension.pyd"), Path("msvcp140.dll"))
                )

    def test_update_zip_and_windows_only_bridge_manifest_are_deterministic(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir, updater = self._fake_app(tmp_dir)
            output_dir = tmp_dir / "release-assets"
            setup_name = f"HRToolkit_{self.version}_x64-setup.exe"
            fake_setup = output_dir / setup_name
            output_dir.mkdir(parents=True, exist_ok=True)
            fake_setup.write_bytes(b"MZ" + b"\0" * 32)
            with (
                patch.object(
                    build_update_assets,
                    "stage_windows_payload",
                    side_effect=AssertionError("skip 模式不应重复复制 payload"),
                ),
                patch.object(
                    build_update_assets,
                    "run_runtime_smoke",
                    side_effect=AssertionError("skip 模式不应重复启动程序"),
                ),
            ):
                setup_path, manifest_path = build_update_assets.build_update_assets(
                    version=self.version,
                    app_dir=app_dir,
                    updater=updater,
                    output_dir=output_dir,
                    notes=["桥接 GitHub Release"],
                    runtime_smoke=False,
                )
            first_digest = build_update_assets.sha256_file(setup_path)

            self.assertEqual(setup_path.name, f"HRToolkit_{self.version}_x64-setup.exe")
            self.assertEqual(manifest_path.name, "legacy-server-latest.json")
            self.assertFalse((output_dir / "latest.json").exists())

            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["version"], self.version)
            self.assertEqual(set(manifest["platforms"]), {"windows"})
            windows = manifest["platforms"]["windows"]
            self.assertEqual(windows["version"], self.version)
            self.assertEqual(windows["sha256"], first_digest)
            self.assertEqual(
                windows["file_url"],
                "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/download/"
                f"v{self.version}/{setup_path.name}",
            )
            self.assertEqual(
                windows["fallback_urls"],
                [
                    "https://github.com/xhzwjc/hr-toolkit/releases/download/"
                    f"v{self.version}/{setup_path.name}"
                ],
            )

            # staging 生成不得污染纯 PyInstaller 输出目录。
            self.assertFalse((app_dir / "HRToolkitUpdater.exe").exists())
            self.assertFalse((app_dir / "update_url.txt").exists())

    def test_win7_bridge_manifest_uses_only_win7_update_channel(self) -> None:
        filename = f"HRToolkit_{self.version}_win7_x64-setup.exe"
        manifest = build_update_assets.legacy_server_manifest(
            version=self.version,
            filename=filename,
            sha256="a" * 64,
            notes=None,
            mandatory=True,
            target=build_windows.WINDOWS_TARGET_WIN7,
        )
        self.assertEqual(set(manifest["platforms"]), {"windows-x64-win7"})
        self.assertNotIn("windows", manifest["platforms"])
        self.assertTrue(
            manifest["platforms"]["windows-x64-win7"]["file_url"].endswith(filename)
        )

    def test_win7_staged_payload_allows_only_pinned_root_runtimes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            payload = Path(tmp)
            for name in (
                "HRToolkit.exe",
                "HRToolkitUpdater.exe",
                *build_windows.WIN7_REQUIRED_UCRT_FILES,
                *build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES,
            ):
                (payload / name).write_bytes(b"runtime")
            (payload / "update_url.txt").write_text(
                "\n".join(build_update_assets.UPDATE_MANIFEST_URLS) + "\n",
                encoding="utf-8",
            )

            with (
                patch.object(build_update_assets, "verify_windows_payload"),
                patch.object(build_update_assets, "verify_pe_x64"),
            ):
                build_update_assets.verify_staged_payload(
                    payload,
                    target=build_windows.WINDOWS_TARGET_WIN7,
                )

                extra = payload / "api-ms-win-core-path-l1-1-0.dll"
                extra.write_bytes(b"forbidden")
                with self.assertRaisesRegex(RuntimeError, "\u6839\u6587\u4ef6"):
                    build_update_assets.verify_staged_payload(
                        payload,
                        target=build_windows.WINDOWS_TARGET_WIN7,
                    )

    def test_installer_definitions_are_per_user_and_keep_payload_under_app_subdir(self) -> None:
        build_windows_installers.validate_installer_definitions()
        attributes = (build_windows.REPO_ROOT / ".gitattributes").read_text(encoding="utf-8")
        self.assertIn("packaging/windows/ChineseSimplified.isl text eol=lf", attributes)
        iss = build_windows_installers.INNO_SCRIPT.read_text(encoding="utf-8")
        self.assertIn("PrivilegesRequired=lowest", iss)
        self.assertIn("DefaultDirName={localappdata}\\Programs\\HRToolkit", iss)
        self.assertIn('DestDir: "{app}\\app"', iss)
        self.assertIn('Type: filesandordirs; Name: "{app}\\app"', iss)
        self.assertIn("Check: ExistingPayloadIsWin7", iss)
        self.assertIn("function ExistingPayloadIsWin7: Boolean;", iss)
        self.assertIn("SignTool={#SignToolName}", iss)
        self.assertIn("Compression=lzma2/max", iss)
        self.assertNotIn("Compression=lzma2/ultra64", iss)
        self.assertIn(
            'MessagesFile: "compiler:Default.isl,ChineseSimplified.isl"',
            iss,
        )

        tree = ET.parse(build_windows_installers.WIX_SOURCE)
        root = tree.getroot()
        namespace = {"w": build_windows_installers.WIX_NAMESPACE}
        package = root.find("w:Package", namespace)
        self.assertIsNotNone(package)
        assert package is not None
        self.assertEqual(package.attrib["Scope"], "perUser")
        major_upgrade = root.find(".//w:MajorUpgrade", namespace)
        self.assertIsNotNone(major_upgrade)
        assert major_upgrade is not None
        self.assertEqual(major_upgrade.attrib["AllowSameVersionUpgrades"], "yes")
        media_template = root.find(".//w:MediaTemplate", namespace)
        self.assertIsNotNone(media_template)
        assert media_template is not None
        self.assertEqual(media_template.attrib["CompressionLevel"], "medium")
        app_directory = root.find(".//w:Directory[@Id='APPDIR']", namespace)
        self.assertIsNotNone(app_directory)
        self.assertEqual(app_directory.attrib["Name"], "app")

    def test_generated_wix_payload_fragment_references_only_staged_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir, updater = self._fake_app(tmp_dir)
            staged = tmp_dir / "staged"
            build_update_assets.stage_windows_payload(
                app_dir=app_dir,
                updater=updater,
                target_dir=staged,
            )
            fragment = build_windows_installers.generate_wix_payload_fragment(
                staged,
                tmp_dir / "payload.wxs",
            )
            tree = ET.parse(fragment)
            namespace = {"w": build_windows_installers.WIX_NAMESPACE}
            files = tree.findall(".//w:File", namespace)
            refs = tree.findall(".//w:ComponentRef", namespace)
            staged_files = [path for path in staged.rglob("*") if path.is_file()]
            self.assertEqual(len(files), len(staged_files))
            self.assertEqual(len(refs), len(staged_files))
            self.assertEqual(
                {Path(item.attrib["Source"]).resolve() for item in files},
                {path.resolve() for path in staged_files},
            )

    def test_installer_commands_and_names_are_deterministic_x64(self) -> None:
        exe_name, msi_name = build_windows_installers.installer_asset_names(self.version)
        self.assertEqual(exe_name, f"HRToolkit_{self.version}_x64-setup.exe")
        self.assertEqual(msi_name, f"HRToolkit_{self.version}_x64.msi")
        inno = build_windows_installers.inno_compile_command(
            compiler="ISCC.exe",
            version=self.version,
            payload_dir=Path("C:/payload"),
            output_dir=Path("C:/assets"),
        )
        self.assertIn(f"/DMyAppVersion={self.version}", inno)
        self.assertIn("/DInstallerSuffix=x64", inno)
        self.assertIn("/DMinWindowsVersion=6.3", inno)
        self.assertNotIn("/DWin7Compatibility=1", inno)
        self.assertNotIn("/DCleanExistingPayload=1", inno)
        self.assertEqual(inno[-1], str(build_windows_installers.INNO_SCRIPT))

        wix = build_windows_installers.wix_build_command(
            wix_executable="wix.exe",
            version=self.version,
            payload_fragment=Path("C:/payload.wxs"),
            output_path=Path("C:/assets") / msi_name,
        )
        self.assertIn("x64", wix)
        self.assertIn(f"AppVersion={self.version}", wix)
        self.assertIn("WindowsTarget=modern", wix)
        self.assertEqual(wix[-1], str(Path("C:/assets") / msi_name))

        win7_exe, win7_msi = build_windows_installers.installer_asset_names(
            self.version,
            build_windows.WINDOWS_TARGET_WIN7,
        )
        self.assertEqual(win7_exe, f"HRToolkit_{self.version}_win7_x64-setup.exe")
        self.assertEqual(win7_msi, f"HRToolkit_{self.version}_win7_x64.msi")
        win7_inno = build_windows_installers.inno_compile_command(
            compiler="ISCC.exe",
            version=self.version,
            payload_dir=Path("C:/payload"),
            output_dir=Path("C:/assets"),
            target=build_windows.WINDOWS_TARGET_WIN7,
        )
        self.assertIn("/DInstallerSuffix=win7_x64", win7_inno)
        self.assertIn("/DMinWindowsVersion=6.1sp1", win7_inno)
        self.assertIn("/DWin7Compatibility=1", win7_inno)
        self.assertIn("/DCleanExistingPayload=1", win7_inno)

    def test_win7_installer_smoke_skips_only_on_unsupported_windows(self) -> None:
        modern_windows = SimpleNamespace(
            major=10,
            minor=0,
            platform_version=(10, 0, 26100),
            service_pack_major=0,
        )
        with (
            patch.object(build_windows_installers, "ensure_windows_runtime"),
            patch.object(
                build_windows_installers.sys,
                "getwindowsversion",
                return_value=modern_windows,
                create=True,
            ),
            patch.object(build_windows_installers, "_smoke_test_inno") as smoke_inno,
            patch.object(build_windows_installers, "_smoke_test_msi") as smoke_msi,
        ):
            build_windows_installers.smoke_test_installers(
                Path("win7.exe"),
                Path("win7.msi"),
                target=build_windows.WINDOWS_TARGET_WIN7,
            )
            smoke_inno.assert_not_called()
            smoke_msi.assert_not_called()

            build_windows_installers.smoke_test_installers(
                Path("modern.exe"),
                Path("modern.msi"),
                target=build_windows.WINDOWS_TARGET_MODERN,
            )
            smoke_inno.assert_called_once()
            smoke_msi.assert_called_once()

    def test_win7_installer_smoke_runs_on_windows7_sp1(self) -> None:
        win7_sp1 = SimpleNamespace(
            major=6,
            minor=1,
            platform_version=(6, 1, 7601),
            service_pack_major=1,
        )
        with (
            patch.object(build_windows_installers, "ensure_windows_runtime"),
            patch.object(
                build_windows_installers.sys,
                "getwindowsversion",
                return_value=win7_sp1,
                create=True,
            ),
            patch.object(build_windows_installers, "_smoke_test_inno") as smoke_inno,
            patch.object(build_windows_installers, "_smoke_test_msi") as smoke_msi,
        ):
            build_windows_installers.smoke_test_installers(
                Path("win7.exe"),
                Path("win7.msi"),
                target=build_windows.WINDOWS_TARGET_WIN7,
            )

        smoke_inno.assert_called_once()
        smoke_msi.assert_called_once()

    def test_release_windows_only_orchestrates_three_stages_without_version_bump(self) -> None:
        commands = release_windows.stage_commands(
            version=self.version,
            build_dir=Path("C:/build"),
            work_dir=Path("C:/work"),
            output_dir=Path("C:/assets"),
            notes=["test"],
        )
        self.assertEqual([label.split()[0] for label, _command in commands], ["1/3", "2/3", "3/3"])
        flat = "\n".join(" ".join(command) for _label, command in commands)
        self.assertIn("build_windows.py", flat)
        self.assertIn("build_windows_installers.py", flat)
        self.assertIn("build_update_assets.py", flat)
        self.assertIn("--skip-runtime-smoke", commands[2][1])
        self.assertNotIn("bump_version", flat)
        self.assertNotIn("prepare_gitee_release", flat)
        self.assertNotIn("git add", flat)
        self.assertNotIn("--publish-dir", flat)

        win7_commands = release_windows.stage_commands(
            version=self.version,
            build_dir=Path("C:/build-win7"),
            work_dir=Path("C:/work-win7"),
            output_dir=Path("C:/assets-win7"),
            target=build_windows.WINDOWS_TARGET_WIN7,
            seven_zip_dir=Path("C:/runtime/7zip"),
            ucrt_dir=Path("C:/runtime/ucrt"),
            vc_runtime_dir=Path("C:/runtime/vc-runtime"),
        )
        win7_flat = "\n".join(" ".join(command) for _label, command in win7_commands)
        self.assertEqual(win7_flat.count("--target win7"), 3)
        self.assertIn("--seven-zip-dir", win7_flat)
        self.assertIn("--ucrt-dir", win7_flat)
        self.assertIn("--vc-runtime-dir", win7_flat)

    def test_release_workflow_publishes_only_gitee_source_metadata_after_github(self) -> None:
        workflow = (build_windows.REPO_ROOT / ".github" / "workflows" / "release.yml").read_text(
            encoding="utf-8"
        )
        mirror_job = workflow.split("\n  mirror-gitee:", 1)[1]
        job_configuration = mirror_job.split("\n    steps:", 1)[0]
        self.assertIn("- publish", job_configuration)
        self.assertIn("always()", job_configuration)
        self.assertIn("needs.publish.result == 'success'", job_configuration)
        self.assertIn("secrets.GITEE_TOKEN", mirror_job)
        self.assertIn("python scripts/gitee_git.py fetch gitee", mirror_job)
        self.assertIn("python scripts/gitee_git.py ls-remote gitee", mirror_job)
        self.assertIn("push --atomic gitee", mirror_job)
        self.assertIn("publish_gitee_release.py", mirror_job)
        self.assertIn('gh release download "${TAG}"', mirror_job)
        self.assertIn('--pattern "SHA256SUMS.txt"', mirror_job)
        self.assertIn("--source-metadata-only", mirror_job)
        self.assertIn("--upload-transport urllib", mirror_job)
        self.assertNotIn("gitee-release-assets", mirror_job)
        self.assertNotIn("GITEE_MAX_ASSET_BYTES", mirror_job)
        self.assertNotIn("latest.json", mirror_job)
        self.assertNotIn("x64-setup.exe", mirror_job)
        self.assertNotIn(".msi", mirror_job)
        self.assertNotIn(".dmg", mirror_job)
        self.assertNotIn("git add", mirror_job)
        self.assertNotIn("git commit", mirror_job)

    def test_clean_win7_acceptance_script_checks_installed_runtime_and_smoke(self) -> None:
        script = (
            build_windows.REPO_ROOT / "scripts" / "smoke_win7_installer.ps1"
        ).read_text(encoding="utf-8")
        self.assertIn('Version -notlike "6.1.*"', script)
        self.assertIn("ServicePackMajorVersion", script)
        self.assertIn("python38.dll", script)
        self.assertIn("python3.dll", script)
        self.assertIn("vcruntime140_1.dll", script)
        self.assertIn('(Join-Path $payload "ucrtbase.dll")', script)
        self.assertNotIn('(Join-Path $internal "ucrtbase.dll")', script)
        self.assertIn("api-ms-win-core-path-l1-1-0.dll", script)
        self.assertIn("Remove-Item Env:\\HR_TOOLKIT_7ZIP_EXE", script)
        self.assertIn('ArgumentList "--smoke-test"', script)
        self.assertIn("HRToolkitUpdater $ExpectedVersion smoke-test OK", script)

    def test_gitee_source_sync_is_manual_and_never_publishes_a_release(self) -> None:
        workflow = (
            build_windows.REPO_ROOT / ".github" / "workflows" / "gitee-sync.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("workflow_dispatch:", workflow)
        self.assertIn("secrets.GITEE_TOKEN", workflow)
        self.assertIn("git merge-base --is-ancestor refs/remotes/gitee/main HEAD", workflow)
        self.assertIn("push gitee HEAD:refs/heads/main", workflow)
        self.assertIn("python scripts/gitee_git.py ls-remote gitee refs/heads/main", workflow)
        self.assertNotIn("refs/tags", workflow)
        self.assertNotIn("publish_gitee_release.py", workflow)
        self.assertNotIn("gh release", workflow)

    def test_gitee_release_workflow_publishes_source_archives_and_checksum_only(self) -> None:
        workflow = (
            build_windows.REPO_ROOT / ".github" / "workflows" / "gitee-release.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("workflow_dispatch:", workflow)
        self.assertIn("tag:", workflow)
        self.assertIn("push --atomic gitee", workflow)
        self.assertIn("secrets.GITEE_TOKEN", workflow)
        self.assertIn("ref: main", workflow)
        self.assertIn("path: tagged-source", workflow)
        self.assertIn(
            "--project-version-file tagged-source/hr_toolkit/__init__.py",
            workflow,
        )
        self.assertIn('gh release download "${TAG}"', workflow)
        self.assertIn('--pattern "SHA256SUMS.txt"', workflow)
        self.assertIn("publish_gitee_release.py", workflow)
        self.assertIn("--source-metadata-only", workflow)
        self.assertIn("--upload-transport urllib", workflow)
        self.assertNotIn("gitee-release-assets", workflow)
        self.assertNotIn("GITEE_MAX_ASSET_BYTES", workflow)
        self.assertNotIn("latest.json", workflow)
        self.assertNotIn("x64-setup.exe", workflow)
        self.assertNotIn(".msi", workflow)
        self.assertNotIn(".dmg", workflow)
        self.assertNotIn("build_windows.py", workflow)
        self.assertNotIn("build_macos.py", workflow)

    def test_win7_runtime_download_is_pinned_and_hash_verified(self) -> None:
        self.assertEqual(prepare_win7_runtime.SEVEN_ZIP_VERSION, "26.02")
        self.assertEqual(len(prepare_win7_runtime.SEVEN_ZIP_SHA256), 64)
        self.assertIn("github.com/ip7z/7zip/releases/download/26.02", prepare_win7_runtime.SEVEN_ZIP_URL)
        self.assertEqual(prepare_win7_runtime.UCRT_VERSION, "10.0.14393.795")
        self.assertEqual(len(prepare_win7_runtime.UCRT_SHA256), 64)
        self.assertIn("download.microsoft.com", prepare_win7_runtime.UCRT_URL)
        self.assertIn(
            "948a611cd2aca64b1e5113ffb7b95d5f.cab",
            prepare_win7_runtime.UCRT_URL,
        )
        self.assertEqual(len(build_windows.WIN7_REQUIRED_UCRT_FILES), 41)
        self.assertNotIn(
            "api-ms-win-core-path-l1-1-0.dll",
            build_windows.WIN7_REQUIRED_UCRT_FILES,
        )
        self.assertEqual(prepare_win7_runtime.VC_REDIST_VERSION, "14.29.30157")
        self.assertEqual(len(prepare_win7_runtime.VC_REDIST_SHA256), 64)
        self.assertEqual(
            len(prepare_win7_runtime.VC_REDIST_ATTACHED_CONTAINER_SHA256),
            64,
        )
        self.assertEqual(len(prepare_win7_runtime.VC_REDIST_MINIMUM_CAB_SHA256), 64)
        self.assertIn("download.visualstudio.microsoft.com", prepare_win7_runtime.VC_REDIST_URL)
        self.assertEqual(
            set(WIN7_UPDATER_APP_LOCAL_RUNTIME_FILES),
            set(build_windows.WIN7_REQUIRED_UCRT_FILES)
            | set(build_windows.WIN7_REQUIRED_VC_RUNTIME_FILES),
        )
        self.assertEqual(
            build_windows.WIN7_PINNED_DISTRIBUTIONS["onnxruntime"],
            "1.11.1",
        )
        constraints = (
            build_windows.REPO_ROOT / "constraints" / "python38-win7.txt"
        ).read_text(encoding="utf-8")
        self.assertIn("onnxruntime==1.11.1", constraints)
        self.assertNotIn("onnxruntime==1.14.1", constraints)

    def test_verified_slice_rejects_truncated_or_changed_content(self) -> None:
        payload = b"pinned embedded container"
        expected_sha256 = hashlib.sha256(payload).hexdigest()
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            source = tmp_dir / "bundle.exe"
            destination = tmp_dir / "attached.cab"
            source.write_bytes(b"prefix" + payload + b"signature")
            prepare_win7_runtime._copy_verified_slice(
                source=source,
                destination=destination,
                offset=len(b"prefix"),
                size=len(payload),
                expected_sha256=expected_sha256,
            )
            self.assertEqual(destination.read_bytes(), payload)

            with self.assertRaisesRegex(RuntimeError, "校验失败"):
                prepare_win7_runtime._copy_verified_slice(
                    source=source,
                    destination=destination,
                    offset=len(b"prefix"),
                    size=len(payload),
                    expected_sha256="0" * 64,
                )
            self.assertFalse(destination.exists())

            with self.assertRaisesRegex(RuntimeError, "不完整"):
                prepare_win7_runtime._copy_verified_slice(
                    source=source,
                    destination=destination,
                    offset=source.stat().st_size - 1,
                    size=2,
                    expected_sha256=expected_sha256,
                )
            self.assertFalse(destination.exists())

    def test_vc_redist_extracts_pinned_embedded_minimum_runtime(self) -> None:
        subprocess_calls = []
        minimum_cab = b"pinned minimum cab"

        def fake_download(_url, destination, _expected_sha256):
            destination.write_bytes(b"pinned VC redist")

        def fake_copy_slice(**kwargs):
            kwargs["destination"].write_bytes(b"attached container")

        def fake_run(command, **kwargs):
            subprocess_calls.append((command, kwargs))
            destination = Path(command[-1])
            if command[1] == f"-F:{prepare_win7_runtime.VC_REDIST_MINIMUM_CAB_MEMBER}":
                (destination / prepare_win7_runtime.VC_REDIST_MINIMUM_CAB_MEMBER).write_bytes(
                    minimum_cab
                )
            return SimpleNamespace(returncode=0)

        with tempfile.TemporaryDirectory() as tmp:
            with (
                patch.object(
                    prepare_win7_runtime,
                    "_download_verified",
                    side_effect=fake_download,
                ),
                patch.object(
                    prepare_win7_runtime,
                    "_copy_verified_slice",
                    side_effect=fake_copy_slice,
                ),
                patch.object(
                    prepare_win7_runtime.subprocess,
                    "run",
                    side_effect=fake_run,
                ),
                patch.object(prepare_win7_runtime, "_replace_with_discovered_files"),
                patch.object(
                    prepare_win7_runtime,
                    "VC_REDIST_MINIMUM_CAB_SHA256",
                    hashlib.sha256(minimum_cab).hexdigest(),
                ),
            ):
                prepare_win7_runtime._prepare_vc_runtime(Path(tmp) / "vc-runtime")

        self.assertEqual(subprocess_calls[0][0][0], "expand.exe")
        self.assertEqual(
            subprocess_calls[0][0][1],
            f"-F:{prepare_win7_runtime.VC_REDIST_MINIMUM_CAB_MEMBER}",
        )
        self.assertEqual(subprocess_calls[1][0][0:2], ["expand.exe", "-F:*"])
        self.assertFalse(
            any("/layout" in command or "msiexec.exe" in command for command, _ in subprocess_calls)
        )

    def test_win7_ucrt_archive_files_are_discovered_by_embedded_dll_name(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            extracted = Path(tmp) / "extracted"
            extracted.mkdir()
            expected = {
                "api-ms-win-core-console-l1-1-0.dll": extracted / "fil-api-set",
                "ucrtbase.dll": extracted / "fil-ucrtbase",
            }
            for name, path in expected.items():
                self._write_fake_pe(path, build_windows.PE_MACHINE_AMD64)
                with path.open("ab") as handle:
                    handle.write(name.encode("utf-16le"))

            with patch.object(
                prepare_win7_runtime,
                "WIN7_REQUIRED_UCRT_FILES",
                tuple(expected),
            ):
                selected = prepare_win7_runtime._discover_ucrt_files(extracted)

        self.assertEqual(selected, expected)

    def test_win7_build_rejects_drift_from_compatibility_dependency_versions(self) -> None:
        expected = build_windows.WIN7_PINNED_DISTRIBUTIONS
        with patch.object(
            build_windows.importlib_metadata,
            "version",
            side_effect=lambda distribution: expected[distribution],
        ):
            build_windows._ensure_win7_pinned_distributions()

        with patch.object(
            build_windows.importlib_metadata,
            "version",
            side_effect=lambda distribution: (
                "1.12.1" if distribution == "onnxruntime" else expected[distribution]
            ),
        ):
            with self.assertRaisesRegex(RuntimeError, "onnxruntime=1.12.1"):
                build_windows._ensure_win7_pinned_distributions()

    def test_installer_output_magic_checks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            exe = tmp_dir / "setup.exe"
            msi = tmp_dir / "setup.msi"
            exe.write_bytes(b"MZ" + b"\0" * 16)
            msi.write_bytes(build_windows_installers.MSI_MAGIC + b"\0" * 16)
            build_windows_installers.verify_installer_outputs(exe, msi)

    def test_win7_inno_output_directory_is_absolute_before_it_exists(self) -> None:
        relative_output = Path("artifacts") / "windows-win7-not-created"
        command = build_windows_installers.inno_compile_command(
            compiler="ISCC.exe",
            version=self.version,
            payload_dir=Path("C:/payload"),
            output_dir=relative_output,
            target=build_windows.WINDOWS_TARGET_WIN7,
        )
        output_argument = next(
            value for value in command if value.startswith("/DOutputDir=")
        )
        self.assertTrue(Path(output_argument.split("=", 1)[1]).is_absolute())

    def _fake_app(
        self,
        root: Path,
        *,
        target: str = build_windows.WINDOWS_TARGET_MODERN,
    ) -> tuple[Path, Path]:
        app_dir = root / "HRToolkit"
        templates = app_dir / "_internal" / "hr_toolkit" / "templates"
        templates.mkdir(parents=True)
        self._write_fake_pe(app_dir / "HRToolkit.exe", build_windows.PE_MACHINE_AMD64)
        (app_dir / "_internal" / "README.md").write_bytes(build_windows.README_FILE.read_bytes())
        self._write_fake_pe(
            app_dir / "_internal" / "runtime.dll",
            build_windows.PE_MACHINE_AMD64,
        )
        if target == build_windows.WINDOWS_TARGET_MODERN:
            for name in build_windows.MODERN_REQUIRED_VC_RUNTIME_FILES:
                self._write_fake_pe(
                    app_dir / name,
                    build_windows.PE_MACHINE_AMD64,
                )
        for source in build_windows.release_template_files():
            (templates / source.name).write_bytes(b"template:" + source.name.encode("utf-8"))
        qml_root = app_dir / "_internal" / "hr_toolkit" / "gui_qt" / "qml"
        for source in build_windows.release_qml_files():
            target = qml_root / source.relative_to(build_windows.QML_DIR)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(source.read_bytes())
        qt_notice = app_dir / "_internal" / "third_party" / "qt" / build_windows.QT_NOTICE.name
        qt_notice.parent.mkdir(parents=True, exist_ok=True)
        qt_notice.write_bytes(build_windows.QT_NOTICE.read_bytes())
        for qml_root, required in (
            (
                app_dir
                / "_internal"
                / build_windows.QT6_WINDOWS_RUNTIME_ROOT
                / "qml",
                build_windows.QT6_REQUIRED_QML_FILES,
            ),
            (
                app_dir / "_internal" / "PySide2" / "qml",
                build_windows.QT5_REQUIRED_QML_FILES,
            ),
        ):
            for relative in required:
                path = qml_root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                if path.suffix.lower() == ".dll":
                    self._write_fake_pe(path, build_windows.PE_MACHINE_AMD64)
                else:
                    path.write_bytes(b"qt-qml")
        for relative in build_windows.QT5_REQUIRED_RUNTIME_FILES:
            self._write_fake_pe(
                app_dir / "_internal" / relative,
                build_windows.PE_MACHINE_AMD64,
            )
        updater = root / "HRToolkitUpdater.exe"
        self._write_fake_pe(updater, build_windows.PE_MACHINE_AMD64)
        return app_dir, updater

    @staticmethod
    def _write_fake_pe(path: Path, machine: int) -> None:
        payload = bytearray(512)
        payload[:2] = b"MZ"
        struct.pack_into("<I", payload, 0x3C, 0x80)
        payload[0x80:0x84] = b"PE\0\0"
        struct.pack_into("<H", payload, 0x84, machine)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload)


if __name__ == "__main__":
    unittest.main()
