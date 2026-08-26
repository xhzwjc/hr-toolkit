from __future__ import annotations

import importlib.util
import json
import tempfile
import unittest
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "generate_release_metadata.py"
SPEC = importlib.util.spec_from_file_location("generate_release_metadata", SCRIPT_PATH)
release_metadata = importlib.util.module_from_spec(SPEC)
assert SPEC is not None and SPEC.loader is not None
SPEC.loader.exec_module(release_metadata)


class ReleaseMetadataTests(unittest.TestCase):
    VERSION = "0.2.1"
    TAG = "v0.2.1"
    REPOSITORY = "xhzwjc/hr-toolkit"

    def _write_assets(self, root: Path, mac_variant: str) -> None:
        names = release_metadata.release_asset_names(self.VERSION, mac_variant=mac_variant)
        for name in names:
            (root / name).write_bytes(("payload:" + name).encode("utf-8"))

    def test_validate_version_rejects_non_strict_versions(self) -> None:
        for version in ("v0.2.1", "0.2", "01.2.3", "0.2.1-beta", "0.2.1+build"):
            with self.subTest(version=version):
                with self.assertRaises(release_metadata.ReleaseMetadataError):
                    release_metadata.validate_version(version)

    def test_release_asset_size_gate_is_strict_at_decimal_100_mb(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            asset = Path(temporary) / "asset.bin"
            asset.write_bytes(b"small")
            self.assertEqual(
                release_metadata.require_release_asset_under_limit(asset),
                len(b"small"),
            )

            with asset.open("r+b") as handle:
                handle.truncate(release_metadata.MAX_RELEASE_ASSET_BYTES)
            with self.assertRaisesRegex(
                release_metadata.ReleaseMetadataError,
                "必须严格小于 100000000 字节",
            ):
                release_metadata.require_release_asset_under_limit(asset)

    def test_release_identity_requires_exact_tag_and_project_version(self) -> None:
        release_metadata.validate_release_identity(self.VERSION, self.TAG, self.VERSION)
        with self.assertRaisesRegex(release_metadata.ReleaseMetadataError, "Tag 与版本不一致"):
            release_metadata.validate_release_identity(self.VERSION, "v0.2.2", self.VERSION)
        with self.assertRaisesRegex(release_metadata.ReleaseMetadataError, "__version__"):
            release_metadata.validate_release_identity(self.VERSION, self.TAG, "0.1.32")

    def test_generates_universal_manifest_and_deterministic_checksums(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._write_assets(assets_dir, "universal2")

            latest_path, checksums_path, names = release_metadata.generate_release_metadata(
                assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.REPOSITORY,
                project_version=self.VERSION,
                notes=("一次性桥接版本",),
            )

            self.assertIn("latest.json", names)
            self.assertIn("SHA256SUMS.txt", names)
            manifest = json.loads(latest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["version"], self.VERSION)
            self.assertTrue(manifest["mandatory"])
            self.assertEqual(manifest["notes"], ["一次性桥接版本"])
            self.assertEqual(
                set(manifest["platforms"]),
                {"windows", "windows-x64-modern", "windows-x64-win7", "macos"},
            )
            self.assertEqual(manifest["platforms"]["windows"]["update_mode"], "auto")
            self.assertEqual(
                manifest["platforms"]["windows"],
                manifest["platforms"]["windows-x64-modern"],
            )
            self.assertTrue(
                manifest["platforms"]["windows-x64-win7"]["file_url"].endswith(
                    "_win7_x64-setup.exe"
                )
            )
            self.assertEqual(manifest["platforms"]["macos"]["update_mode"], "manual")
            self.assertEqual(
                manifest["platforms"]["macos"]["file_url"],
                "https://github.com/xhzwjc/hr-toolkit/releases/download/v0.2.1/"
                "HRToolkit_0.2.1_universal.dmg",
            )
            checksum_lines = checksums_path.read_text(encoding="utf-8").splitlines()
            checksum_names = [line.split("  ", 1)[1] for line in checksum_lines]
            expected_checksum_names = sorted(
                release_metadata.release_asset_names(self.VERSION, mac_variant="universal2")
                + ("latest.json",)
            )
            self.assertEqual(checksum_names, expected_checksum_names)
            self.assertNotIn("SHA256SUMS.txt", checksum_names)

    def test_generates_arch_specific_manual_macos_payloads_after_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._write_assets(assets_dir, "split")

            latest_path, _checksums_path, _names = release_metadata.generate_release_metadata(
                assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.REPOSITORY,
                project_version=self.VERSION,
            )

            platforms = json.loads(latest_path.read_text(encoding="utf-8"))["platforms"]
            self.assertEqual(
                set(platforms),
                {
                    "windows",
                    "windows-x64-modern",
                    "windows-x64-win7",
                    "macos-arm64",
                    "macos-x64",
                },
            )
            self.assertEqual(platforms["macos-arm64"]["update_mode"], "manual")
            self.assertTrue(platforms["macos-arm64"]["file_url"].endswith("_arm64.dmg"))
            self.assertTrue(platforms["macos-x64"]["file_url"].endswith("_x64.dmg"))

    def test_generates_gitee_primary_manifest_with_github_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._write_assets(assets_dir, "universal2")

            latest_path, _checksums_path, _names = release_metadata.generate_release_metadata(
                assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.REPOSITORY,
                project_version=self.VERSION,
                download_base_url=(
                    "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/download"
                ),
                release_url=(
                    "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/tag/v0.2.1"
                ),
                fallback_download_base_url=(
                    "https://github.com/xhzwjc/hr-toolkit/releases/download"
                ),
            )

            manifest = json.loads(latest_path.read_text(encoding="utf-8"))
            windows = manifest["platforms"]["windows"]
            win7 = manifest["platforms"]["windows-x64-win7"]
            self.assertTrue(windows["file_url"].startswith("https://gitee.com/"))
            self.assertTrue(win7["file_url"].startswith("https://gitee.com/"))
            self.assertEqual(
                windows["fallback_urls"],
                [
                    "https://github.com/xhzwjc/hr-toolkit/releases/download/v0.2.1/"
                    "HRToolkit_0.2.1_x64-setup.exe"
                ],
            )
            self.assertEqual(
                manifest["release_url"],
                "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/tag/v0.2.1",
            )

    def test_limits_gitee_primary_to_explicit_installer_and_size(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._write_assets(assets_dir, "universal2")
            windows_installer = f"HRToolkit_{self.VERSION}_x64-setup.exe"
            win7_installer = f"HRToolkit_{self.VERSION}_win7_x64-setup.exe"

            latest_path, _checksums_path, _names = release_metadata.generate_release_metadata(
                assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.REPOSITORY,
                project_version=self.VERSION,
                download_base_url=(
                    "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/download"
                ),
                release_url=(
                    "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/tag/v0.2.1"
                ),
                fallback_download_base_url=(
                    "https://github.com/xhzwjc/hr-toolkit/releases/download"
                ),
                primary_download_max_bytes=100_000_000,
                primary_download_asset_names=(windows_installer, win7_installer),
            )

            platforms = json.loads(latest_path.read_text(encoding="utf-8"))["platforms"]
            self.assertTrue(
                platforms["windows"]["file_url"].startswith("https://gitee.com/")
            )
            self.assertEqual(len(platforms["windows"]["fallback_urls"]), 1)
            self.assertTrue(
                platforms["windows-x64-win7"]["file_url"].startswith("https://gitee.com/")
            )
            self.assertTrue(
                platforms["macos"]["file_url"].startswith("https://github.com/")
            )
            self.assertNotIn("fallback_urls", platforms["macos"])

            latest_path, _checksums_path, _names = release_metadata.generate_release_metadata(
                assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.REPOSITORY,
                project_version=self.VERSION,
                download_base_url=(
                    "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/download"
                ),
                release_url=(
                    "https://gitee.com/optimistic-little-sunspot/hr-toolkit/releases/tag/v0.2.1"
                ),
                fallback_download_base_url=(
                    "https://github.com/xhzwjc/hr-toolkit/releases/download"
                ),
                primary_download_max_bytes=1,
                primary_download_asset_names=(windows_installer, win7_installer),
            )
            windows = json.loads(latest_path.read_text(encoding="utf-8"))["platforms"][
                "windows"
            ]
            self.assertTrue(windows["file_url"].startswith("https://github.com/"))
            self.assertNotIn("fallback_urls", windows)

    def test_rejects_incomplete_or_mixed_macos_assets(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._write_assets(assets_dir, "universal2")
            (assets_dir / f"HRToolkit_{self.VERSION}_arm64.dmg").write_bytes(b"arm")
            with self.assertRaisesRegex(release_metadata.ReleaseMetadataError, "同时发布"):
                release_metadata.detect_mac_variant(assets_dir, self.VERSION)

        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            for name in release_metadata.release_asset_names(self.VERSION, mac_variant="split")[:3]:
                (assets_dir / name).write_bytes(b"windows")
            (assets_dir / f"HRToolkit_{self.VERSION}_arm64.dmg").write_bytes(b"arm")
            with self.assertRaisesRegex(release_metadata.ReleaseMetadataError, "同时提供"):
                release_metadata.detect_mac_variant(assets_dir, self.VERSION)

    def test_rejects_non_whitelisted_release_file(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._write_assets(assets_dir, "universal2")
            (assets_dir / "actions-artifact.zip").write_bytes(b"not a direct asset")
            with self.assertRaisesRegex(release_metadata.ReleaseMetadataError, "非白名单"):
                release_metadata.generate_release_metadata(
                    assets_dir,
                    version=self.VERSION,
                    tag=self.TAG,
                    repository=self.REPOSITORY,
                    project_version=self.VERSION,
                )


if __name__ == "__main__":
    unittest.main()
