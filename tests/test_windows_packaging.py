from __future__ import annotations

import json
import struct
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from scripts import build_update_assets
from scripts import build_windows
from scripts import build_windows_installers
from scripts import release_windows
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

        self.assertIn("--onedir", main)
        self.assertIn("--windowed", main)
        self.assertNotIn("--onefile", main)
        self.assertIn("--onefile", updater)
        self.assertIn("--windowed", updater)
        self.assertNotIn("--onedir", updater)
        self.assertEqual(main[-1], str(build_windows.APP_ENTRYPOINT))
        self.assertEqual(updater[-1], str(build_windows.UPDATER_ENTRYPOINT))

        data_values = [main[index + 1] for index, value in enumerate(main[:-1]) if value == "--add-data"]
        self.assertEqual(len(data_values), 1 + len(build_windows.release_template_files()))
        self.assertTrue(any(value.startswith(str(build_windows.README_FILE) + ";") for value in data_values))
        template_sources = {
            value.split(";", 1)[0]
            for value in data_values
            if value.lower().endswith(";hr_toolkit/templates")
        }
        self.assertEqual(template_sources, {str(path) for path in build_windows.release_template_files()})
        self.assertFalse(any(value.startswith(str(build_windows.TEMPLATES_DIR) + ";") for value in data_values))
        self.assertFalse(any("附件" in value or "outputs" in value for value in data_values))

    def test_windows_version_metadata_uses_requested_version(self) -> None:
        payload = build_windows.windows_version_info("0.2.1")
        self.assertIn("filevers=(0, 2, 1, 0)", payload)
        self.assertIn("StringStruct('ProductVersion', '0.2.1')", payload)

    def test_windows_release_job_forces_utf8_python_output(self) -> None:
        workflow = (build_windows.REPO_ROOT / ".github" / "workflows" / "release.yml").read_text(
            encoding="utf-8"
        )
        windows_job = workflow.split("\n  build-windows:", 1)[1].split(
            "\n  build-macos-universal:", 1
        )[0]
        job_configuration = windows_job.split("\n    steps:", 1)[0]
        self.assertIn('PYTHONUTF8: "1"', job_configuration)

    def test_payload_verification_accepts_only_readme_and_builtin_excel_templates(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            app_dir, _updater = self._fake_app(Path(tmp))
            build_windows.verify_windows_payload(app_dir)

            forbidden = app_dir / "真实工资表.xlsx"
            forbidden.write_bytes(b"private")
            with self.assertRaisesRegex(RuntimeError, "模板目录之外"):
                build_windows.verify_windows_payload(app_dir)
            forbidden.unlink()

            cache = app_dir / "_internal" / "__pycache__" / "module.pyc"
            cache.parent.mkdir()
            cache.write_bytes(b"cache")
            with self.assertRaisesRegex(RuntimeError, "禁止目录或缓存"):
                build_windows.verify_windows_payload(app_dir)

    def test_pe_machine_verification_requires_amd64(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            executable = Path(tmp) / "app.exe"
            self._write_fake_pe(executable, build_windows.PE_MACHINE_AMD64)
            self.assertEqual(build_windows.read_pe_machine(executable), build_windows.PE_MACHINE_AMD64)
            build_windows.verify_pe_x64(executable)

            self._write_fake_pe(executable, 0x014C)
            with self.assertRaisesRegex(RuntimeError, "不是 x64 PE"):
                build_windows.verify_pe_x64(executable)

    def test_update_zip_and_windows_only_bridge_manifest_are_deterministic(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir, updater = self._fake_app(tmp_dir)
            output_dir = tmp_dir / "release-assets"
            zip_path, manifest_path = build_update_assets.build_update_assets(
                version=self.version,
                app_dir=app_dir,
                updater=updater,
                output_dir=output_dir,
                notes=["桥接 GitHub Release"],
                runtime_smoke=False,
            )
            first_digest = build_update_assets.sha256_file(zip_path)
            build_update_assets.build_update_assets(
                version=self.version,
                app_dir=app_dir,
                updater=updater,
                output_dir=output_dir,
                notes=["桥接 GitHub Release"],
                runtime_smoke=False,
            )
            self.assertEqual(first_digest, build_update_assets.sha256_file(zip_path))

            self.assertEqual(zip_path.name, f"HRToolkit-{self.version}-win-update.zip")
            self.assertEqual(manifest_path.name, "legacy-server-latest.json")
            self.assertFalse((output_dir / "latest.json").exists())
            with zipfile.ZipFile(zip_path) as archive:
                names = archive.namelist()
                self.assertIn("HRToolkit.exe", names)
                self.assertIn("HRToolkitUpdater.exe", names)
                self.assertIn("update_url.txt", names)
                self.assertTrue(all(info.date_time == build_update_assets.ZIP_EPOCH for info in archive.infolist()))
                update_urls = tuple(
                    line.strip()
                    for line in archive.read("update_url.txt").decode("utf-8").splitlines()
                    if line.strip()
                )
                self.assertEqual(update_urls, build_update_assets.UPDATE_MANIFEST_URLS)
                xlsx_names = [name for name in names if name.lower().endswith(".xlsx")]
                self.assertEqual(
                    {Path(name).name for name in xlsx_names},
                    {path.name for path in build_windows.release_template_files()},
                )
                self.assertTrue(all("hr_toolkit/templates/" in name for name in xlsx_names))
                self.assertFalse(any("tests/" in name or "outputs/" in name or "附件" in name for name in names))

            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["version"], self.version)
            self.assertEqual(set(manifest["platforms"]), {"windows"})
            windows = manifest["platforms"]["windows"]
            self.assertEqual(windows["version"], self.version)
            self.assertEqual(windows["sha256"], first_digest)
            self.assertEqual(
                windows["file_url"],
                "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/download/"
                f"v{self.version}/{zip_path.name}",
            )
            self.assertEqual(
                windows["fallback_urls"],
                [
                    "https://github.com/xhzwjc/hr-toolkit/releases/download/"
                    f"v{self.version}/{zip_path.name}"
                ],
            )

            # staging/zip 生成不得污染纯 PyInstaller 输出目录。
            self.assertFalse((app_dir / "HRToolkitUpdater.exe").exists())
            self.assertFalse((app_dir / "update_url.txt").exists())

    def test_installer_definitions_are_per_user_and_keep_payload_under_app_subdir(self) -> None:
        build_windows_installers.validate_installer_definitions()
        attributes = (build_windows.REPO_ROOT / ".gitattributes").read_text(encoding="utf-8")
        self.assertIn("packaging/windows/ChineseSimplified.isl text eol=lf", attributes)
        iss = build_windows_installers.INNO_SCRIPT.read_text(encoding="utf-8")
        self.assertIn("PrivilegesRequired=lowest", iss)
        self.assertIn("DefaultDirName={localappdata}\\Programs\\HRToolkit", iss)
        self.assertIn('DestDir: "{app}\\app"', iss)
        self.assertIn('Type: filesandordirs; Name: "{app}\\app"', iss)
        self.assertIn("SignTool={#SignToolName}", iss)
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
        self.assertEqual(inno[-1], str(build_windows_installers.INNO_SCRIPT))

        wix = build_windows_installers.wix_build_command(
            wix_executable="wix.exe",
            version=self.version,
            payload_fragment=Path("C:/payload.wxs"),
            output_path=Path("C:/assets") / msi_name,
        )
        self.assertIn("x64", wix)
        self.assertIn(f"AppVersion={self.version}", wix)
        self.assertEqual(wix[-1], str(Path("C:/assets") / msi_name))

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
        self.assertNotIn("bump_version", flat)
        self.assertNotIn("prepare_gitee_release", flat)
        self.assertNotIn("git add", flat)
        self.assertNotIn("--publish-dir", flat)

    def test_release_workflow_mirrors_only_after_github_publish(self) -> None:
        workflow = (build_windows.REPO_ROOT / ".github" / "workflows" / "release.yml").read_text(
            encoding="utf-8"
        )
        mirror_job = workflow.split("\n  mirror-gitee:", 1)[1]
        job_configuration = mirror_job.split("\n    steps:", 1)[0]
        self.assertIn("- publish", job_configuration)
        self.assertIn("always()", job_configuration)
        self.assertIn("needs.publish.result == 'success'", job_configuration)
        self.assertIn("secrets.GITEE_TOKEN", mirror_job)
        self.assertIn("publish_gitee_release.py", mirror_job)
        self.assertIn("--timeout 600", mirror_job)
        self.assertIn("http.postBuffer=1073741824", mirror_job)
        self.assertIn("http.version=HTTP/1.1", mirror_job)
        self.assertIn("push --atomic gitee", mirror_job)
        self.assertNotIn("git add", mirror_job)
        self.assertNotIn("git commit", mirror_job)

    def test_gitee_source_sync_is_manual_and_never_publishes_a_release(self) -> None:
        workflow = (
            build_windows.REPO_ROOT / ".github" / "workflows" / "gitee-sync.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("workflow_dispatch:", workflow)
        self.assertIn("secrets.GITEE_TOKEN", workflow)
        self.assertIn("git merge-base --is-ancestor refs/remotes/gitee/main HEAD", workflow)
        self.assertIn("push gitee HEAD:refs/heads/main", workflow)
        self.assertIn("git ls-remote gitee refs/heads/main", workflow)
        self.assertNotIn("refs/tags", workflow)
        self.assertNotIn("publish_gitee_release.py", workflow)
        self.assertNotIn("gh release", workflow)

    def test_gitee_release_recovery_reuses_github_assets_without_rebuilding(self) -> None:
        workflow = (
            build_windows.REPO_ROOT / ".github" / "workflows" / "gitee-release.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("workflow_dispatch:", workflow)
        self.assertIn("tag:", workflow)
        self.assertIn("gh release download", workflow)
        self.assertIn("publish_gitee_release.py", workflow)
        self.assertIn("push --atomic gitee", workflow)
        self.assertIn("secrets.GITEE_TOKEN", workflow)
        self.assertIn("ref: main", workflow)
        self.assertIn("path: tagged-source", workflow)
        self.assertIn(
            "--project-version-file tagged-source/hr_toolkit/__init__.py",
            workflow,
        )
        self.assertIn("--timeout 600", workflow)
        self.assertNotIn("build_windows.py", workflow)
        self.assertNotIn("build_macos.py", workflow)

    def test_installer_output_magic_checks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            exe = tmp_dir / "setup.exe"
            msi = tmp_dir / "setup.msi"
            exe.write_bytes(b"MZ" + b"\0" * 16)
            msi.write_bytes(build_windows_installers.MSI_MAGIC + b"\0" * 16)
            build_windows_installers.verify_installer_outputs(exe, msi)

    def _fake_app(self, root: Path) -> tuple[Path, Path]:
        app_dir = root / "HRToolkit"
        templates = app_dir / "_internal" / "hr_toolkit" / "templates"
        templates.mkdir(parents=True)
        self._write_fake_pe(app_dir / "HRToolkit.exe", build_windows.PE_MACHINE_AMD64)
        (app_dir / "_internal" / "README.md").write_bytes(build_windows.README_FILE.read_bytes())
        self._write_fake_pe(
            app_dir / "_internal" / "runtime.dll",
            build_windows.PE_MACHINE_AMD64,
        )
        for source in build_windows.release_template_files():
            (templates / source.name).write_bytes(b"template:" + source.name.encode("utf-8"))
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
