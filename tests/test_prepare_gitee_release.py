from __future__ import annotations

import importlib.util
import json
import tempfile
import unittest
import zipfile
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "prepare_gitee_release.py"
SPEC = importlib.util.spec_from_file_location("prepare_gitee_release", SCRIPT_PATH)
prepare_gitee_release = importlib.util.module_from_spec(SPEC)
assert SPEC is not None and SPEC.loader is not None
SPEC.loader.exec_module(prepare_gitee_release)


class PrepareGiteeReleaseTest(unittest.TestCase):
    def test_prepare_release_writes_zip_update_url_and_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            app_dir = tmp_dir / "HRToolkit"
            app_dir.mkdir()
            (app_dir / "HRToolkit.exe").write_text("app", encoding="utf-8")
            updater = tmp_dir / "HRToolkitUpdater.exe"
            updater.write_text("updater", encoding="utf-8")
            output_dir = tmp_dir / "downloads"
            bundle_dir = tmp_dir / "bundle" / "hr-toolkit"
            publish_dir = tmp_dir / "scripthub" / "hr-toolkit"
            manifest_path = prepare_gitee_release.REPO_ROOT / "release" / "latest.json"
            original_manifest = manifest_path.read_text(encoding="utf-8") if manifest_path.exists() else None

            try:
                exit_code = prepare_gitee_release.main([
                    "--platform",
                    "windows",
                    "--version",
                    "9.9.9",
                    "--notes",
                    "测试发布",
                    "--app-dir",
                    str(app_dir),
                    "--updater",
                    str(updater),
                    "--output-dir",
                    str(output_dir),
                    "--bundle-dir",
                    str(bundle_dir),
                    "--publish-dir",
                    str(publish_dir),
                ])
                self.assertEqual(exit_code, 0)
                self.assertEqual((app_dir / "update_url.txt").read_text(encoding="utf-8").strip(), prepare_gitee_release.SCRIPT_HUB_MANIFEST_URL)
                self.assertTrue((app_dir / "HRToolkitUpdater.exe").exists())

                zip_path = output_dir / "HRToolkit-9.9.9-win.zip"
                self.assertTrue(zip_path.exists())
                with zipfile.ZipFile(zip_path) as archive:
                    self.assertIn("HRToolkit.exe", archive.namelist())
                    self.assertIn("HRToolkitUpdater.exe", archive.namelist())
                    self.assertIn("update_url.txt", archive.namelist())

                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                self.assertEqual(manifest["version"], "9.9.9")
                self.assertEqual(manifest["notes"], ["测试发布"])
                self.assertEqual(
                    manifest["platforms"]["windows"]["file_url"],
                    "http://hr.seedlingintl.com/api/static/hr-toolkit/releases/HRToolkit-9.9.9-win.zip",
                )
                self.assertTrue((bundle_dir / "latest.json").exists())
                self.assertTrue((bundle_dir / "releases" / "HRToolkit-9.9.9-win.zip").exists())
                self.assertTrue((publish_dir / "latest.json").exists())
                self.assertTrue((publish_dir / "releases" / "HRToolkit-9.9.9-win.zip").exists())
            finally:
                if original_manifest is None:
                    manifest_path.unlink(missing_ok=True)
                else:
                    manifest_path.write_text(original_manifest, encoding="utf-8")


if __name__ == "__main__":
    unittest.main()
