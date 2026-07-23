from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

METADATA_SPEC = importlib.util.spec_from_file_location(
    "generate_release_metadata",
    SCRIPTS_DIR / "generate_release_metadata.py",
)
release_metadata = importlib.util.module_from_spec(METADATA_SPEC)
assert METADATA_SPEC is not None and METADATA_SPEC.loader is not None
METADATA_SPEC.loader.exec_module(release_metadata)
sys.modules["generate_release_metadata"] = release_metadata

PUBLISH_SPEC = importlib.util.spec_from_file_location(
    "publish_gitee_release",
    SCRIPTS_DIR / "publish_gitee_release.py",
)
gitee_publish = importlib.util.module_from_spec(PUBLISH_SPEC)
assert PUBLISH_SPEC is not None and PUBLISH_SPEC.loader is not None
PUBLISH_SPEC.loader.exec_module(gitee_publish)


class FakeGiteeClient:
    def __init__(self, assets_dir: Path, tag: str) -> None:
        self.assets_dir = assets_dir
        self.tag = tag
        self.release: dict | None = {"id": 7, "tag_name": tag}
        self.attachments: list[dict] = [
            {"id": 1, "name": "stale.bin", "size": 5},
        ]
        self.upload_order: list[str] = []
        self.deleted: list[str] = []
        self.timeout_after_accepting: set[str] = set()

    def get_release_by_tag(self, _repository: str, _tag: str):
        return self.release

    def create_release(self, _repository: str, **kwargs):
        self.release = {"id": 7, "tag_name": kwargs["tag"]}
        return self.release

    def update_release(self, _repository: str, _release_id: str, **kwargs):
        self.release = {"id": 7, "tag_name": kwargs["tag"]}
        return self.release

    def list_attachments(self, _repository: str, _release_id: str):
        return list(self.attachments)

    def delete_attachment(self, _repository: str, _release_id: str, attachment_id: str):
        self.deleted.append(attachment_id)
        self.attachments = [
            item for item in self.attachments if str(item["id"]) != str(attachment_id)
        ]

    def upload_attachment(self, _repository: str, _release_id: str, file_path: Path):
        self.upload_order.append(file_path.name)
        attachment = {
            "id": len(self.attachments) + 10,
            "name": file_path.name,
            "size": file_path.stat().st_size,
            "browser_download_url": (
                f"https://gitee.com/company/hr/releases/download/{self.tag}/{file_path.name}"
            ),
        }
        self.attachments.append(attachment)
        if file_path.name in self.timeout_after_accepting:
            self.timeout_after_accepting.remove(file_path.name)
            raise gitee_publish.GiteeReleaseError("simulated response timeout")
        return attachment

    def get_public_latest_release(self, _repository: str):
        return {"id": 7, "tag_name": self.tag, "assets": list(self.attachments)}

    def get_public_json(self, _url: str):
        return json.loads((self.assets_dir / "latest.json").read_text(encoding="utf-8"))


class GiteeMirrorTests(unittest.TestCase):
    VERSION = "0.2.3"
    TAG = "v0.2.3"
    GITEE_REPOSITORY = "optimistic-little-sunspot/hr-toolkit"
    GITHUB_REPOSITORY = "xhzwjc/hr-toolkit"

    def _build_assets(self, assets_dir: Path) -> None:
        for name in release_metadata.release_asset_names(self.VERSION, mac_variant="universal2"):
            (assets_dir / name).write_bytes(("payload:" + name).encode("utf-8"))
        release_metadata.generate_release_metadata(
            assets_dir,
            version=self.VERSION,
            tag=self.TAG,
            repository=self.GITHUB_REPOSITORY,
            project_version=self.VERSION,
            download_base_url=(
                f"https://gitee.com/{self.GITEE_REPOSITORY}/releases/download"
            ),
            release_url=f"https://gitee.com/{self.GITEE_REPOSITORY}/releases/tag/{self.TAG}",
            fallback_download_base_url=(
                f"https://github.com/{self.GITHUB_REPOSITORY}/releases/download"
            ),
        )

    def test_validates_exact_gitee_assets_and_github_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._build_assets(assets_dir)

            names = gitee_publish.validate_mirror_assets(
                assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.GITEE_REPOSITORY,
            )

            self.assertEqual(set(names), {path.name for path in assets_dir.iterdir()})

    def test_rejects_manifest_without_github_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._build_assets(assets_dir)
            latest_path = assets_dir / "latest.json"
            manifest = json.loads(latest_path.read_text(encoding="utf-8"))
            manifest["platforms"]["windows"].pop("fallback_urls")
            latest_path.write_text(json.dumps(manifest), encoding="utf-8")

            with self.assertRaisesRegex(gitee_publish.GiteeReleaseError, "GitHub 备用"):
                gitee_publish.validate_mirror_assets(
                    assets_dir,
                    version=self.VERSION,
                    tag=self.TAG,
                    repository=self.GITEE_REPOSITORY,
                )

    def test_publish_is_idempotent_and_uploads_latest_json_last(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._build_assets(assets_dir)
            client = FakeGiteeClient(assets_dir, self.TAG)

            _release, names = gitee_publish.publish_gitee_release(
                client,
                assets_dir=assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.GITEE_REPOSITORY,
                target_commitish="a" * 40,
                name=f"HR Toolkit {self.TAG}",
                body="mirror",
            )
            gitee_publish.verify_public_release(
                client,
                assets_dir=assets_dir,
                repository=self.GITEE_REPOSITORY,
                tag=self.TAG,
                names=names,
                attempts=1,
            )

            self.assertEqual(client.deleted, ["1"])
            self.assertEqual(client.upload_order[-1], "latest.json")
            self.assertEqual(set(client.upload_order), set(names))

            client.upload_order.clear()
            gitee_publish.publish_gitee_release(
                client,
                assets_dir=assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.GITEE_REPOSITORY,
                target_commitish="a" * 40,
                name=f"HR Toolkit {self.TAG}",
                body="mirror",
            )
            self.assertEqual(client.upload_order, [])

    def test_upload_timeout_after_server_accepts_does_not_duplicate_asset(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._build_assets(assets_dir)
            client = FakeGiteeClient(assets_dir, self.TAG)
            first_asset = sorted(
                name
                for name in gitee_publish.expected_asset_names(assets_dir, self.VERSION)
                if name != "latest.json"
            )[0]
            client.timeout_after_accepting.add(first_asset)

            _release, names = gitee_publish.publish_gitee_release(
                client,
                assets_dir=assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.GITEE_REPOSITORY,
                target_commitish="a" * 40,
                name=f"HR Toolkit {self.TAG}",
                body="mirror",
                retry_delay=0,
            )

            self.assertEqual(client.upload_order.count(first_asset), 1)
            self.assertEqual(client.upload_order[-1], "latest.json")
            self.assertEqual(
                {item["name"] for item in client.attachments},
                set(names),
            )

    def test_partial_release_resumes_and_reuploads_latest_json_last(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._build_assets(assets_dir)
            client = FakeGiteeClient(assets_dir, self.TAG)
            names = gitee_publish.expected_asset_names(assets_dir, self.VERSION)
            completed_name = next(name for name in names if name not in {"latest.json", "SHA256SUMS.txt"})
            completed_path = assets_dir / completed_name
            client.attachments = [
                {
                    "id": 20,
                    "name": completed_name,
                    "size": completed_path.stat().st_size,
                },
                {
                    "id": 21,
                    "name": "latest.json",
                    "size": (assets_dir / "latest.json").stat().st_size,
                },
            ]

            gitee_publish.publish_gitee_release(
                client,
                assets_dir=assets_dir,
                version=self.VERSION,
                tag=self.TAG,
                repository=self.GITEE_REPOSITORY,
                target_commitish="a" * 40,
                name=f"HR Toolkit {self.TAG}",
                body="mirror",
            )

            self.assertNotIn(completed_name, client.upload_order)
            self.assertEqual(client.upload_order[-1], "latest.json")
            self.assertIn("21", client.deleted)

    def test_dry_run_does_not_require_token(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            assets_dir = Path(temporary)
            self._build_assets(assets_dir)

            exit_code = gitee_publish.main(
                [
                    "--version",
                    self.VERSION,
                    "--tag",
                    self.TAG,
                    "--repository",
                    self.GITEE_REPOSITORY,
                    "--target-commitish",
                    "a" * 40,
                    "--assets-dir",
                    str(assets_dir),
                    "--dry-run",
                ]
            )

            self.assertEqual(exit_code, 0)


if __name__ == "__main__":
    unittest.main()
