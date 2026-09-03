from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
import io
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import call, patch

from scripts import gitee_git


CI_RATE_LIMIT = (
    "error: RPC failed; HTTP 429 curl 22 The requested URL returned error: 429\n"
    "fatal: expected flush after ref listing\n"
)


class GiteeGitTests(unittest.TestCase):
    def test_actual_ci_429_retries_with_backoff_then_succeeds(self) -> None:
        failure = subprocess.CompletedProcess([], 128, "", CI_RATE_LIMIT)
        success = subprocess.CompletedProcess([], 0, "", "From Gitee\n")
        output, errors = io.StringIO(), io.StringIO()
        with patch.object(gitee_git.subprocess, "run", side_effect=[failure, failure, success]) as run:
            with patch.object(gitee_git.time, "sleep") as sleep:
                with redirect_stdout(output), redirect_stderr(errors):
                    result = gitee_git.run_git(["fetch", "gitee", "refs/heads/main:refs/remotes/gitee/main"])
        self.assertEqual(result, 0)
        self.assertEqual(run.call_count, 3)
        self.assertEqual(sleep.call_args_list, [call(15), call(30)])
        self.assertEqual(output.getvalue(), "")
        self.assertIn("HTTP 429", errors.getvalue())
        self.assertIn("retry 2/4", errors.getvalue())
        self.assertIn("http.version=HTTP/1.1", run.call_args.args[0])
        self.assertEqual(run.call_args.kwargs["env"]["GIT_TERMINAL_PROMPT"], "0")

    def test_ls_remote_stdout_contains_only_successful_reference_records(self) -> None:
        reference = "a" * 40 + "\trefs/tags/v0.7.8\n"
        results = [
            subprocess.CompletedProcess([], 128, "discard incomplete response", CI_RATE_LIMIT),
            subprocess.CompletedProcess([], 0, reference, "warning: a server notice\n"),
        ]
        output, errors = io.StringIO(), io.StringIO()
        with patch.object(gitee_git.subprocess, "run", side_effect=results):
            with patch.object(gitee_git.time, "sleep"):
                with redirect_stdout(output), redirect_stderr(errors):
                    self.assertEqual(gitee_git.run_git(["ls-remote", "gitee", "refs/tags/v0.7.8"]), 0)
        self.assertEqual(output.getvalue(), reference)
        self.assertIn("discard incomplete response", errors.getvalue())
        self.assertIn("warning: a server notice", errors.getvalue())

    def test_persistent_rate_limit_exits_nonzero_after_four_attempts(self) -> None:
        with patch.object(gitee_git.subprocess, "run", return_value=subprocess.CompletedProcess([], 128, "", CI_RATE_LIMIT)) as run:
            with patch.object(gitee_git.time, "sleep") as sleep:
                with redirect_stderr(io.StringIO()) as errors:
                    self.assertEqual(gitee_git.run_git(["fetch", "gitee"]), 128)
        self.assertEqual(run.call_count, 4)
        self.assertEqual(sleep.call_args_list, [call(15), call(30), call(60)])
        self.assertIn("source synchronization was not confirmed", errors.getvalue())

    def test_permanent_errors_are_not_retried(self) -> None:
        for detail in (
            "fatal: Authentication failed",
            "The requested URL returned error: 403",
            "fatal: repository not found",
            "! [rejected] main -> main (non-fast-forward)",
            "! [rejected] v0.7.8 -> v0.7.8 (already exists)",
            "fatal: the receiving end does not support --atomic push",
            "fatal: SSL certificate problem: certificate has expired",
            "fatal: unknown error",
        ):
            with self.subTest(detail=detail):
                with patch.object(gitee_git.subprocess, "run", return_value=subprocess.CompletedProcess([], 128, "", detail)) as run:
                    with patch.object(gitee_git.time, "sleep") as sleep:
                        with redirect_stderr(io.StringIO()):
                            self.assertEqual(gitee_git.run_git(["push", "--atomic", "gitee"]), 128)
                self.assertEqual(run.call_count, 1)
                sleep.assert_not_called()

    def test_transient_server_and_connection_errors_retry(self) -> None:
        for detail in ("HTTP 503", "returned error: 502", "Connection reset by peer", "Operation timed out"):
            with self.subTest(detail=detail):
                results = [
                    subprocess.CompletedProcess([], 128, "", detail),
                    subprocess.CompletedProcess([], 0, "", ""),
                ]
                with patch.object(gitee_git.subprocess, "run", side_effect=results):
                    with patch.object(gitee_git.time, "sleep") as sleep:
                        with redirect_stderr(io.StringIO()):
                            self.assertEqual(gitee_git.run_git(["fetch", "gitee"]), 0)
                sleep.assert_called_once_with(15)

    def test_retry_diagnostics_do_not_expose_token(self) -> None:
        with patch.dict(os.environ, {"GITEE_TOKEN": "test-private-token"}):
            with patch.object(gitee_git.subprocess, "run", return_value=subprocess.CompletedProcess([], 128, "", "Authentication failed test-private-token")):
                with redirect_stderr(io.StringIO()) as errors:
                    self.assertEqual(gitee_git.run_git(["fetch", "gitee"]), 128)
        self.assertNotIn("test-private-token", errors.getvalue())
        self.assertIn("***", errors.getvalue())

    def test_local_validation_commands_cannot_be_retried_by_wrapper(self) -> None:
        with patch.object(gitee_git.subprocess, "run") as run:
            with redirect_stderr(io.StringIO()):
                self.assertEqual(gitee_git.run_git(["merge-base", "--is-ancestor", "a", "b"]), 2)
        run.assert_not_called()

    def test_every_gitee_workflow_uses_retry_transport_and_preserves_ref_guards(self) -> None:
        root = Path(__file__).resolve().parents[1]
        for name in ("release.yml", "gitee-release.yml", "gitee-sync.yml"):
            with self.subTest(workflow=name):
                source = (root / ".github/workflows" / name).read_text(encoding="utf-8")
                self.assertIn("python scripts/gitee_git.py fetch gitee", source)
                self.assertIn("python scripts/gitee_git.py ls-remote gitee", source)
                self.assertIn("python scripts/gitee_git.py push", source)
                self.assertIn("git merge-base --is-ancestor refs/remotes/gitee/main", source)
                self.assertNotIn("git fetch gitee", source)
                self.assertNotIn("git ls-remote gitee", source)
                self.assertNotIn("--force", source)
                self.assertNotIn("continue-on-error", source)
                if name != "gitee-sync.yml":
                    self.assertIn('"${REMOTE_TAG}" != "${LOCAL_TAG}"', source)
                    self.assertIn("push --atomic gitee", source)


class GiteeGitLocalRepositoryTests(unittest.TestCase):
    """Exercise real Git; the only push destination is a temporary local repo."""

    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.source = Path(self.temporary.name) / "source"
        self.remote = Path(self.temporary.name) / "remote.git"
        self.source.mkdir()
        self._git("init", "--bare", str(self.remote))
        self._git("init", "-b", "main", str(self.source))
        self._git("commit", "--allow-empty", "-m", "test source")
        self.commit = self._git("rev-parse", "HEAD").strip()
        self._git("tag", "-a", "v0.7.8", "-m", "test release")
        self.tag = self._git("rev-parse", "refs/tags/v0.7.8").strip()
        self.push = ["push", "--atomic", str(self.remote), f"{self.commit}:refs/heads/main", "refs/tags/v0.7.8:refs/tags/v0.7.8"]

    def _git(self, *args: str) -> str:
        return subprocess.run(
            ["git", "-c", "user.name=CI Test", "-c", "user.email=ci@example.invalid", "-c", "commit.gpgsign=false", "-c", "tag.gpgsign=false", *args],
            cwd=self.source, check=True, capture_output=True, text=True,
        ).stdout

    def test_lost_push_response_is_safe_to_repeat_and_preserves_annotated_tag(self) -> None:
        run = subprocess.run
        calls = []

        def transport(command, **kwargs):
            calls.append(command)
            result = run(command, cwd=self.source, **kwargs)
            if len(calls) == 1:
                self.assertEqual(result.returncode, 0)
                return subprocess.CompletedProcess(command, 128, "", "RPC failed; HTTP 502")
            return result

        with patch.object(gitee_git.subprocess, "run", side_effect=transport):
            with patch.object(gitee_git.time, "sleep") as sleep:
                with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
                    self.assertEqual(gitee_git.run_git(self.push), 0)
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[0], calls[1])
        self.assertIn("http.postBuffer=1073741824", calls[0])
        sleep.assert_called_once_with(15)
        refs = self._git("ls-remote", str(self.remote))
        self.assertIn(self.commit + "\trefs/heads/main", refs)
        self.assertIn(self.tag + "\trefs/tags/v0.7.8\n", refs)
        self.assertIn(self.commit + "\trefs/tags/v0.7.8^{}", refs)

    def test_atomic_rejection_never_rolls_back_remote_main_or_creates_tag(self) -> None:
        self._git("commit", "--allow-empty", "-m", "newer remote source")
        newer = self._git("rev-parse", "HEAD").strip()
        self._git("push", str(self.remote), "HEAD:refs/heads/main")
        run = subprocess.run

        def transport(command, **kwargs):
            return run(command, cwd=self.source, **kwargs)

        with patch.object(gitee_git.subprocess, "run", side_effect=transport) as transport_mock:
            with patch.object(gitee_git.time, "sleep") as sleep:
                with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
                    self.assertNotEqual(gitee_git.run_git(self.push), 0)
        self.assertEqual(transport_mock.call_count, 1)
        sleep.assert_not_called()
        refs = self._git("ls-remote", str(self.remote))
        self.assertIn(newer + "\trefs/heads/main", refs)
        self.assertNotIn("refs/tags/v0.7.8", refs)


if __name__ == "__main__":
    unittest.main()
