from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "release.py"
SPEC = importlib.util.spec_from_file_location("hr_toolkit_release", SCRIPT_PATH)
assert SPEC is not None and SPEC.loader is not None
release = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = release
SPEC.loader.exec_module(release)


START_HEAD = "1" * 40
RELEASE_COMMIT = "2" * 40
TAG_OBJECT = "3" * 40
REMOTE_URL = "https://github.com/xhzwjc/hr-toolkit.git"


def write_version_tree(root: Path, version: str = "0.1.32") -> None:
    package_dir = root / "hr_toolkit"
    package_dir.mkdir(parents=True)
    (package_dir / "__init__.py").write_text(
        f'"""test package"""\n\n__version__ = "{version}"\n',
        encoding="utf-8",
    )
    (root / "package.json").write_text(
        json.dumps(
            {
                "name": "hr-toolkit",
                "private": True,
                "version": version,
                "scripts": {"release": "python3 scripts/release.py"},
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    (root / "package-lock.json").write_text(
        json.dumps(
            {
                "name": "hr-toolkit",
                "version": version,
                "lockfileVersion": 3,
                "requires": True,
                "packages": {"": {"name": "hr-toolkit", "version": version}},
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


class PrepareGit:
    def __init__(self) -> None:
        self.branch_name = "main"
        self.head_oid = START_HEAD
        self.origin_oid = START_HEAD
        self.remote_tag = False
        self.local_tag = False
        self.fetch_count = 0
        self.clean_checks = 0

    def branch(self) -> str:
        return self.branch_name

    def ensure_clean(self) -> None:
        self.clean_checks += 1

    def ensure_no_git_operation(self) -> None:
        pass

    def ensure_origin(self) -> str:
        return REMOTE_URL

    def fetch_origin(self, remote_url: str) -> None:
        assert remote_url == REMOTE_URL
        self.fetch_count += 1

    def head(self) -> str:
        return self.head_oid

    def origin_main(self) -> str:
        return self.origin_oid

    def remote_tag_exists(self, _tag: str, remote_url: str) -> bool:
        assert remote_url == REMOTE_URL
        return self.remote_tag

    def local_tag_exists(self, _tag: str) -> bool:
        return self.local_tag


class ReleaseGit:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.branch_name = "main"
        self.origin_oid = START_HEAD
        self.head_oid = START_HEAD
        self.tag_oid = None
        self.staged = False
        self.commit_message = None
        self.pushed_tag = None
        self.push_error = None
        self.calls: list[tuple] = []
        self.remote_refs = {"refs/heads/main": START_HEAD}

    def ensure_clean(self) -> None:
        self.calls.append(("ensure_clean",))

    def ensure_no_git_operation(self) -> None:
        self.calls.append(("ensure_no_git_operation",))

    def branch(self) -> str:
        return self.branch_name

    def ensure_origin(self) -> str:
        return REMOTE_URL

    def fetch_origin(self, remote_url: str) -> None:
        assert remote_url == REMOTE_URL
        self.calls.append(("fetch", remote_url))

    def origin_main(self) -> str:
        return self.origin_oid

    def remote_tag_exists(self, _tag: str, remote_url: str) -> bool:
        assert remote_url == REMOTE_URL
        return False

    def local_tag_exists(self, _tag: str) -> bool:
        return self.tag_oid is not None

    def staged_paths(self) -> set[str]:
        return set(release.VERSION_FILES) if self.staged else set()

    def unstaged_paths(self) -> set[str]:
        if self.staged or self.head_oid == RELEASE_COMMIT:
            return set()
        if release.read_configured_versions(self.root).current != "0.1.32":
            return set(release.VERSION_FILES)
        return set()

    def untracked_paths(self) -> set[str]:
        return set()

    def stage_version_files(self) -> None:
        self.staged = True
        self.calls.append(("stage", tuple(release.VERSION_FILES)))

    def check_staged_diff(self) -> None:
        self.calls.append(("staged_diff_check",))

    def staged_file_bytes(self, relative: str) -> bytes:
        return (self.root / relative).read_bytes()

    def commit(self, message: str) -> str:
        self.commit_message = message
        self.head_oid = RELEASE_COMMIT
        self.staged = False
        self.calls.append(("commit", message))
        return RELEASE_COMMIT

    def commit_parents(self, _commit: str) -> tuple[str, ...]:
        return (START_HEAD,)

    def commit_paths(self, _commit: str) -> set[str]:
        return set(release.VERSION_FILES)

    def committed_file_bytes(self, _commit: str, relative: str) -> bytes:
        return (self.root / relative).read_bytes()

    def create_annotated_tag(self, tag: str, commit: str) -> None:
        assert commit == RELEASE_COMMIT
        self.tag_oid = TAG_OBJECT
        self.calls.append(("annotated_tag", tag, commit))

    def local_ref_oid(self, ref: str):
        if ref.endswith("^{}"):
            return RELEASE_COMMIT if self.tag_oid is not None else None
        return self.tag_oid

    def push_atomic(self, *, tag: str, remote_url: str, release_commit: str, tag_object: str) -> None:
        assert remote_url == REMOTE_URL
        assert release_commit == RELEASE_COMMIT
        assert tag_object == TAG_OBJECT
        self.pushed_tag = tag
        self.calls.append(("atomic_push", tag, release_commit, tag_object))
        if self.push_error is not None:
            raise self.push_error
        self.remote_refs = {
            "refs/heads/main": RELEASE_COMMIT,
            f"refs/tags/{tag}": TAG_OBJECT,
            f"refs/tags/{tag}^{{}}": RELEASE_COMMIT,
        }

    def head(self) -> str:
        return self.head_oid

    def delete_tag(self, tag: str) -> None:
        self.calls.append(("delete_tag", tag))
        self.tag_oid = None

    def reset_mixed(self, commit: str) -> None:
        self.calls.append(("reset_mixed", commit))
        self.head_oid = commit
        self.staged = False

    def restore_staged_version_files(self) -> None:
        self.calls.append(("restore_staged",))
        self.staged = False

    def remote_release_refs(self, _tag: str, remote_url: str):
        assert remote_url == REMOTE_URL
        return dict(self.remote_refs)


def make_plan(root: Path) -> object:
    return release.ReleasePlan(
        root=root,
        current_version="0.1.32",
        target_version="0.2.1",
        start_head=START_HEAD,
        tag="v0.2.1",
        remote_url=REMOTE_URL,
    )


class StableSemVerTest(unittest.TestCase):
    def test_accepts_canonical_stable_versions(self) -> None:
        self.assertEqual(release.parse_stable_semver("0.2.1"), (0, 2, 1))
        self.assertEqual(release.parse_stable_semver("12.34.56"), (12, 34, 56))

    def test_rejects_noncanonical_or_nonstable_versions(self) -> None:
        invalid = (
            "01.2.3",
            "1.02.3",
            "1.2.03",
            "1.2",
            "1.2.3.4",
            "v1.2.3",
            "1.2.3-alpha",
            "1.2.3+build",
            " 1.2.3",
            "1.2.3 ",
            "",
        )
        for version in invalid:
            with self.subTest(version=version):
                with self.assertRaises(release.ReleaseError):
                    release.parse_stable_semver(version)


class RemoteIdentityTest(unittest.TestCase):
    def test_normalizes_common_github_url_spellings(self) -> None:
        expected = "github.com/xhzwjc/hr-toolkit"
        self.assertEqual(release.canonical_remote_identity(REMOTE_URL), expected)
        self.assertEqual(
            release.canonical_remote_identity("git@github.com:xhzwjc/hr-toolkit.git"),
            expected,
        )

    def test_rejects_fetch_and_push_urls_for_different_repositories(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(
                ["git", "remote", "add", "origin", REMOTE_URL],
                cwd=root,
                check=True,
            )
            subprocess.run(
                [
                    "git",
                    "remote",
                    "set-url",
                    "--push",
                    "origin",
                    "git@github.com:someone-else/hr-toolkit.git",
                ],
                cwd=root,
                check=True,
            )
            with self.assertRaisesRegex(release.ReleaseError, "不是同一个仓库"):
                release.GitRepository(root).ensure_origin()


class ConfiguredVersionTest(unittest.TestCase):
    def test_reads_all_synchronized_version_fields(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)

            configured = release.read_configured_versions(root)

            self.assertEqual(configured.current, "0.1.32")
            self.assertEqual(len(configured.values), 4)
            self.assertEqual(set(configured.values.values()), {"0.1.32"})

    def test_rejects_mismatched_version_fields(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            package = json.loads((root / "package.json").read_text(encoding="utf-8"))
            package["version"] = "0.1.31"
            (root / "package.json").write_text(json.dumps(package), encoding="utf-8")

            with self.assertRaisesRegex(release.ReleaseError, "版本文件不同步"):
                release.read_configured_versions(root)

    def test_render_is_in_memory_and_updates_every_field(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            originals = release.snapshot_version_files(root)

            rendered = release.render_version_files(root, "0.2.1")

            self.assertEqual(release.snapshot_version_files(root), originals)
            release.write_version_files(root, rendered)
            self.assertEqual(release.read_configured_versions(root).current, "0.2.1")
            self.assertEqual(set(rendered), set(release.VERSION_FILES))


class PrepareReleaseTest(unittest.TestCase):
    def test_requires_fetch_and_exact_origin_main_sync(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = PrepareGit()

            plan = release.prepare_release("0.2.1", root, git)

            self.assertEqual(plan.start_head, START_HEAD)
            self.assertEqual(plan.tag, "v0.2.1")
            self.assertEqual(git.fetch_count, 1)
            self.assertEqual(git.clean_checks, 2)

            git.origin_oid = "9" * 40
            with self.assertRaisesRegex(release.ReleaseError, "完全一致"):
                release.prepare_release("0.2.1", root, git)

    def test_requires_target_greater_and_no_local_or_remote_tag(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = PrepareGit()

            with self.assertRaisesRegex(release.ReleaseError, "目标版本必须高于"):
                release.prepare_release("0.1.32", root, git)

            git.remote_tag = True
            with self.assertRaisesRegex(release.ReleaseError, "远端 Tag"):
                release.prepare_release("0.2.1", root, git)

            git.remote_tag = False
            git.local_tag = True
            with self.assertRaisesRegex(release.ReleaseError, "本地 Tag"):
                release.prepare_release("0.2.1", root, git)


class DryRunTest(unittest.TestCase):
    def test_dry_run_runs_checks_without_writing_or_git_mutations(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            before = release.snapshot_version_files(root)
            git = ReleaseGit(root)
            checks: list[Path] = []
            plan = make_plan(root)

            release.execute_release_plan(
                plan,
                git,
                dry_run=True,
                checks=lambda checked_root: checks.append(checked_root),
            )

            self.assertEqual(checks, [root])
            self.assertEqual(release.snapshot_version_files(root), before)
            self.assertGreaterEqual(git.calls.count(("ensure_clean",)), 2)
            self.assertEqual(git.head_oid, START_HEAD)
            self.assertIsNone(git.tag_oid)
            self.assertIsNone(git.pushed_tag)

    def test_dry_run_rechecks_remote_state_after_long_checks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = ReleaseGit(root)

            def advance_remote(_root: Path) -> None:
                git.origin_oid = "9" * 40

            with self.assertRaisesRegex(release.ReleaseError, "origin/main"):
                release.execute_release_plan(
                    make_plan(root),
                    git,
                    dry_run=True,
                    checks=advance_remote,
                )
            self.assertEqual(release.read_configured_versions(root).current, "0.1.32")

    def test_full_checks_include_tests_compileall_and_diff_check(self) -> None:
        captured_commands: list[list[str]] = []

        class FakeRunner:
            def __init__(self, _root: Path) -> None:
                pass

            def run(self, command, *, capture=True, check=True):
                captured_commands.append(list(command))
                return None

        with mock.patch.object(release, "CommandRunner", FakeRunner):
            release.run_full_checks(Path("/unused"))

        self.assertEqual(captured_commands[0][1:5], ["-m", "unittest", "discover", "-s"])
        self.assertIn("compileall", captured_commands[1])
        self.assertEqual(captured_commands[2], ["git", "diff", "--check"])
        self.assertFalse(
            any(
                Path(part).name in {"build_macos.py", "build_windows.py"}
                for command in captured_commands
                for part in command
            )
        )


class NormalReleaseTest(unittest.TestCase):
    def test_normal_flow_stages_only_whitelist_and_uses_atomic_push(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = ReleaseGit(root)
            plan = make_plan(root)

            release.execute_release_plan(plan, git, dry_run=False, checks=lambda _root: None)

            self.assertEqual(release.read_configured_versions(root).current, "0.2.1")
            self.assertEqual(git.commit_message, "chore(recruitment): 发布 0.2.1")
            self.assertIn(("stage", tuple(release.VERSION_FILES)), git.calls)
            self.assertIn(("annotated_tag", "v0.2.1", RELEASE_COMMIT), git.calls)
            self.assertIn(
                ("atomic_push", "v0.2.1", RELEASE_COMMIT, TAG_OBJECT),
                git.calls,
            )

    def test_pre_push_failure_restores_version_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            before = release.snapshot_version_files(root)
            git = ReleaseGit(root)
            plan = make_plan(root)

            def fail_checks(_root: Path) -> None:
                raise release.ReleaseError("test failure")

            with self.assertRaisesRegex(release.ReleaseError, "本地发布状态已回滚"):
                release.execute_release_plan(plan, git, dry_run=False, checks=fail_checks)

            self.assertEqual(release.snapshot_version_files(root), before)
            self.assertEqual(git.head_oid, START_HEAD)
            self.assertIsNone(git.tag_oid)

    def test_staged_blob_mismatch_is_rejected_and_rolled_back(self) -> None:
        class TamperedIndexGit(ReleaseGit):
            def staged_file_bytes(self, relative: str) -> bytes:
                if relative == "package.json":
                    return b'{"version":"0.2.1","scripts":{"release":"tampered"}}\n'
                return super().staged_file_bytes(relative)

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = TamperedIndexGit(root)
            with self.assertRaisesRegex(release.ReleaseError, "本地发布状态已回滚"):
                release.execute_release_plan(
                    make_plan(root), git, dry_run=False, checks=lambda _root: None
                )
            self.assertEqual(release.read_configured_versions(root).current, "0.1.32")

    def test_commit_blob_mismatch_is_rejected_and_rolled_back(self) -> None:
        class TamperedCommitGit(ReleaseGit):
            def committed_file_bytes(self, _commit: str, relative: str) -> bytes:
                if relative == "package.json":
                    return b"tampered\n"
                return super().committed_file_bytes(_commit, relative)

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = TamperedCommitGit(root)
            with self.assertRaisesRegex(release.ReleaseError, "本地发布状态已回滚"):
                release.execute_release_plan(
                    make_plan(root), git, dry_run=False, checks=lambda _root: None
                )
            self.assertEqual(release.read_configured_versions(root).current, "0.1.32")


class PushFailureRecoveryTest(unittest.TestCase):
    def _written_state(self, root: Path, git: ReleaseGit):
        snapshots = release.snapshot_version_files(root)
        rendered = release.render_version_files(root, "0.2.1")
        release.write_version_files(root, rendered)
        git.head_oid = RELEASE_COMMIT
        git.tag_oid = TAG_OBJECT
        return release.LocalReleaseState(
            root=root,
            start_head=START_HEAD,
            tag="v0.2.1",
            snapshots=snapshots,
            rendered=rendered,
            remote_url=REMOTE_URL,
            files_written=True,
            release_commit=RELEASE_COMMIT,
            tag_created=True,
            tag_object=TAG_OBJECT,
        )

    def test_classifies_remote_refs(self) -> None:
        state = release.LocalReleaseState(
            root=Path("/unused"),
            start_head=START_HEAD,
            tag="v0.2.1",
            snapshots={},
            rendered={},
            remote_url=REMOTE_URL,
            release_commit=RELEASE_COMMIT,
            tag_created=True,
            tag_object=TAG_OBJECT,
        )
        self.assertIs(
            release.classify_remote_push_state({"refs/heads/main": START_HEAD}, state),
            release.RemotePushState.UNCHANGED,
        )
        self.assertIs(
            release.classify_remote_push_state(
                {
                    "refs/heads/main": RELEASE_COMMIT,
                    "refs/tags/v0.2.1": TAG_OBJECT,
                    "refs/tags/v0.2.1^{}": RELEASE_COMMIT,
                },
                state,
            ),
            release.RemotePushState.APPLIED,
        )
        self.assertIs(
            release.classify_remote_push_state(
                {"refs/heads/main": RELEASE_COMMIT}, state
            ),
            release.RemotePushState.UNKNOWN,
        )

    def test_failed_push_rolls_back_only_when_both_remote_refs_unchanged(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = ReleaseGit(root)
            state = self._written_state(root, git)
            git.remote_refs = {"refs/heads/main": START_HEAD}

            with self.assertRaisesRegex(release.ReleaseError, "均未变化"):
                release.resolve_failed_push(git, state, RuntimeError("push failed"))

            self.assertEqual(release.read_configured_versions(root).current, "0.1.32")
            self.assertEqual(git.head_oid, START_HEAD)
            self.assertIsNone(git.tag_oid)
            self.assertIn(("delete_tag", "v0.2.1"), git.calls)
            self.assertIn(("reset_mixed", START_HEAD), git.calls)

    def test_execute_plan_does_not_attempt_a_second_rollback_after_push_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = ReleaseGit(root)
            git.push_error = RuntimeError("push failed")
            git.remote_refs = {"refs/heads/main": START_HEAD}
            plan = make_plan(root)

            with self.assertRaisesRegex(release.ReleaseError, "均未变化"):
                release.execute_release_plan(plan, git, dry_run=False, checks=lambda _root: None)

            self.assertEqual(git.calls.count(("delete_tag", "v0.2.1")), 1)
            self.assertEqual(git.calls.count(("reset_mixed", START_HEAD)), 1)
            self.assertEqual(release.read_configured_versions(root).current, "0.1.32")

    def test_partial_or_unknown_remote_state_is_preserved_for_manual_recovery(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = ReleaseGit(root)
            state = self._written_state(root, git)
            git.remote_refs = {"refs/heads/main": RELEASE_COMMIT}

            with self.assertRaisesRegex(release.ManualRecoveryRequired, "部分更新或未知"):
                release.resolve_failed_push(git, state, RuntimeError("push failed"))

            self.assertEqual(release.read_configured_versions(root).current, "0.2.1")
            self.assertEqual(git.head_oid, RELEASE_COMMIT)
            self.assertEqual(git.tag_oid, TAG_OBJECT)
            self.assertNotIn(("delete_tag", "v0.2.1"), git.calls)
            self.assertNotIn(("reset_mixed", START_HEAD), git.calls)

    def test_rollback_refuses_to_overwrite_concurrent_version_file_edit(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_version_tree(root)
            git = ReleaseGit(root)
            state = self._written_state(root, git)
            package = json.loads((root / "package.json").read_text(encoding="utf-8"))
            package["scripts"] = {"release": "user-concurrent-edit"}
            (root / "package.json").write_text(
                json.dumps(package, indent=2) + "\n", encoding="utf-8"
            )

            with self.assertRaisesRegex(
                release.ManualRecoveryRequired, "并发变化"
            ):
                release.rollback_local_release(git, state)

            self.assertEqual(
                json.loads((root / "package.json").read_text(encoding="utf-8"))["scripts"],
                {"release": "user-concurrent-edit"},
            )


if __name__ == "__main__":
    unittest.main()
