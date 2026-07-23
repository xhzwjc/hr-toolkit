#!/usr/bin/env python3
"""Create and push an HR Toolkit release commit and tag.

The local release command deliberately does not build installers.  It validates
the repository, synchronizes the three version files, runs the complete local
checks, and atomically pushes ``main`` together with an annotated ``v*`` tag.
GitHub Actions is responsible for all platform builds and publication.
"""

from __future__ import annotations

import argparse
import json
import re
import shlex
import subprocess
import sys
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Callable, Mapping, Optional, Sequence
from urllib.parse import urlsplit


REPO_ROOT = Path(__file__).resolve().parents[1]
MAIN_BRANCH = "main"
ORIGIN = "origin"
VERSION_FILES = (
    "hr_toolkit/__init__.py",
    "package.json",
    "package-lock.json",
)
VERSION_FILE_SET = frozenset(VERSION_FILES)
STABLE_SEMVER_PATTERN = re.compile(
    r"^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$"
)
INIT_VERSION_PATTERN = re.compile(
    r"(?m)^([ \t]*__version__[ \t]*=[ \t]*)([\"'])([^\"']+)([\"'])([ \t]*(?:#.*)?)$"
)


class ReleaseError(RuntimeError):
    """A release cannot continue safely."""


class ManualRecoveryRequired(ReleaseError):
    """Automatic rollback is unsafe because local or remote state is unclear."""


class CommandError(ReleaseError):
    """A child process returned a non-zero exit status."""

    def __init__(
        self,
        command: Sequence[str],
        returncode: int,
        stdout: str = "",
        stderr: str = "",
    ) -> None:
        detail = stderr.strip() or stdout.strip()
        message = f"命令失败（退出码 {returncode}）：{shlex.join(command)}"
        if detail:
            message = f"{message}\n{detail}"
        super().__init__(message)
        self.command = tuple(command)
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


class RemotePushState(Enum):
    UNCHANGED = "unchanged"
    APPLIED = "applied"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class ConfiguredVersions:
    current: str
    values: Mapping[str, str]


@dataclass(frozen=True)
class ReleasePlan:
    root: Path
    current_version: str
    target_version: str
    start_head: str
    tag: str
    remote_url: str


@dataclass
class LocalReleaseState:
    root: Path
    start_head: str
    tag: str
    snapshots: Mapping[str, bytes]
    rendered: Mapping[str, bytes]
    remote_url: str
    files_written: bool = False
    staged: bool = False
    release_commit: Optional[str] = None
    tag_created: bool = False
    tag_object: Optional[str] = None

    @property
    def has_mutations(self) -> bool:
        return self.files_written or self.staged or self.release_commit is not None or self.tag_created


class CommandRunner:
    def __init__(self, root: Path) -> None:
        self.root = root

    def run(
        self,
        command: Sequence[str],
        *,
        capture: bool = True,
        check: bool = True,
    ) -> subprocess.CompletedProcess[str]:
        try:
            result = subprocess.run(
                list(command),
                cwd=self.root,
                text=True,
                stdout=subprocess.PIPE if capture else None,
                stderr=subprocess.PIPE if capture else None,
                check=False,
            )
        except OSError as error:
            raise ReleaseError(f"无法执行命令 {shlex.join(command)}：{error}") from error
        if check and result.returncode != 0:
            raise CommandError(command, result.returncode, result.stdout or "", result.stderr or "")
        return result


class GitRepository:
    def __init__(self, root: Path, runner: Optional[CommandRunner] = None) -> None:
        self.root = root
        self.runner = runner or CommandRunner(root)

    def _run(
        self,
        arguments: Sequence[str],
        *,
        capture: bool = True,
        check: bool = True,
    ) -> subprocess.CompletedProcess[str]:
        return self.runner.run(["git", *arguments], capture=capture, check=check)

    def _capture(self, arguments: Sequence[str]) -> str:
        return (self._run(arguments).stdout or "").strip()

    def branch(self) -> str:
        return self._capture(["branch", "--show-current"])

    def ensure_clean(self) -> None:
        status = self._capture(["status", "--porcelain=v1", "--untracked-files=all"])
        if status:
            raise ReleaseError(f"工作区或暂存区不干净，发布已停止：\n{status}")

    def ensure_origin(self) -> str:
        fetch_urls = _line_tuple(
            self._capture(["remote", "get-url", "--all", ORIGIN])
        )
        push_urls = _line_tuple(
            self._capture(["remote", "get-url", "--push", "--all", ORIGIN])
        )
        if len(fetch_urls) != 1 or len(push_urls) != 1:
            raise ReleaseError(
                "origin 必须各自只有一个 fetch URL 和 push URL。"
                f"\nfetch URLs：{fetch_urls}\npush URLs：{push_urls}"
            )
        fetch_identity = canonical_remote_identity(fetch_urls[0])
        push_identity = canonical_remote_identity(push_urls[0])
        if fetch_identity != push_identity:
            raise ReleaseError(
                "origin 的 fetch/push 目标不是同一个仓库，发布已停止。"
                f"\nfetch：{fetch_urls[0]}\npush：{push_urls[0]}"
            )
        # Pin every later network operation to this already-validated URL so a
        # concurrent pushurl change cannot redirect the release transaction.
        return push_urls[0]

    def fetch_origin(self, remote_url: str) -> None:
        self._run(
            [
                "fetch",
                "--quiet",
                "--no-tags",
                remote_url,
                f"+refs/heads/{MAIN_BRANCH}:refs/remotes/{ORIGIN}/{MAIN_BRANCH}",
            ]
        )

    def head(self) -> str:
        return self._capture(["rev-parse", "HEAD"])

    def origin_main(self) -> str:
        return self._capture(["rev-parse", f"refs/remotes/{ORIGIN}/{MAIN_BRANCH}"])

    def local_tag_exists(self, tag: str) -> bool:
        result = self._run(
            ["show-ref", "--verify", "--quiet", f"refs/tags/{tag}"],
            check=False,
        )
        if result.returncode not in (0, 1):
            raise CommandError(
                ["git", "show-ref", "--verify", "--quiet", f"refs/tags/{tag}"],
                result.returncode,
                result.stdout or "",
                result.stderr or "",
            )
        return result.returncode == 0

    def remote_tag_exists(self, tag: str, remote_url: str) -> bool:
        arguments = [
            "ls-remote",
            "--exit-code",
            "--tags",
            remote_url,
            f"refs/tags/{tag}",
        ]
        result = self._run(arguments, check=False)
        if result.returncode not in (0, 2):
            raise CommandError(
                ["git", *arguments],
                result.returncode,
                result.stdout or "",
                result.stderr or "",
            )
        return result.returncode == 0

    def unstaged_paths(self) -> set[str]:
        output = self._capture(["diff", "--name-only", "--no-renames", "--"])
        return _line_set(output)

    def staged_paths(self) -> set[str]:
        output = self._capture(["diff", "--cached", "--name-only", "--no-renames", "--"])
        return _line_set(output)

    def untracked_paths(self) -> set[str]:
        output = self._capture(["ls-files", "--others", "--exclude-standard"])
        return _line_set(output)

    def stage_version_files(self) -> None:
        self._run(["add", "--", *VERSION_FILES])

    def check_staged_diff(self) -> None:
        self._run(["diff", "--cached", "--check"])

    def staged_file_bytes(self, relative: str) -> bytes:
        result = self._run(["show", f":{relative}"])
        return (result.stdout or "").encode("utf-8")

    def commit(self, message: str) -> str:
        self._run(["commit", "-m", message], capture=False)
        return self.head()

    def commit_parents(self, commit: str) -> tuple[str, ...]:
        fields = self._capture(["rev-list", "--parents", "-n", "1", commit]).split()
        if not fields or fields[0] != commit:
            raise ReleaseError(f"无法读取 release commit 父提交：{commit}")
        return tuple(fields[1:])

    def commit_paths(self, commit: str) -> set[str]:
        output = self._capture(
            ["diff-tree", "--no-commit-id", "--name-only", "--no-renames", "-r", commit]
        )
        return _line_set(output)

    def committed_file_bytes(self, commit: str, relative: str) -> bytes:
        result = self._run(["show", f"{commit}:{relative}"])
        return (result.stdout or "").encode("utf-8")

    def create_annotated_tag(self, tag: str, commit: str) -> None:
        self._run(["tag", "-a", tag, "-m", f"HR Toolkit {tag}", commit])

    def local_ref_oid(self, ref: str) -> Optional[str]:
        result = self._run(["rev-parse", "--verify", "--quiet", ref], check=False)
        if result.returncode == 0:
            return (result.stdout or "").strip()
        if result.returncode == 1:
            return None
        raise CommandError(
            ["git", "rev-parse", "--verify", "--quiet", ref],
            result.returncode,
            result.stdout or "",
            result.stderr or "",
        )

    def push_atomic(
        self,
        *,
        tag: str,
        remote_url: str,
        release_commit: str,
        tag_object: str,
    ) -> None:
        self._run(
            [
                "push",
                "--atomic",
                remote_url,
                f"{release_commit}:refs/heads/{MAIN_BRANCH}",
                f"{tag_object}:refs/tags/{tag}",
            ],
            capture=False,
        )

    def remote_release_refs(self, tag: str, remote_url: str) -> Mapping[str, str]:
        arguments = [
            "ls-remote",
            remote_url,
            f"refs/heads/{MAIN_BRANCH}",
            f"refs/tags/{tag}",
            f"refs/tags/{tag}^{{}}",
        ]
        output = self._capture(arguments)
        refs: dict[str, str] = {}
        for line in output.splitlines():
            fields = line.split("\t", 1)
            if len(fields) != 2:
                raise ReleaseError(f"无法解析远端引用：{line!r}")
            oid, ref = fields
            refs[ref] = oid
        return refs

    def delete_tag(self, tag: str) -> None:
        self._run(["tag", "-d", tag])

    def reset_mixed(self, commit: str) -> None:
        self._run(["reset", "--mixed", commit])

    def restore_staged_version_files(self) -> None:
        self._run(["restore", "--staged", "--", *VERSION_FILES])

    def ensure_no_git_operation(self) -> None:
        operation_refs = ("MERGE_HEAD", "CHERRY_PICK_HEAD", "REVERT_HEAD", "REBASE_HEAD")
        active = []
        for ref in operation_refs:
            result = self._run(["rev-parse", "--verify", "--quiet", ref], check=False)
            if result.returncode == 0:
                active.append(ref)
            elif result.returncode != 1:
                raise CommandError(
                    ["git", "rev-parse", "--verify", "--quiet", ref],
                    result.returncode,
                    result.stdout or "",
                    result.stderr or "",
                )
        if active:
            raise ReleaseError(f"仓库存在进行中的 Git 操作：{', '.join(active)}")


def _line_set(output: str) -> set[str]:
    return {line for line in output.splitlines() if line}


def _line_tuple(output: str) -> tuple[str, ...]:
    return tuple(line.strip() for line in output.splitlines() if line.strip())


def canonical_remote_identity(url: str) -> str:
    """Normalize common Git URL spellings to a repository identity."""

    value = url.strip()
    if not value:
        raise ReleaseError("origin URL 不能为空。")
    if "://" in value:
        parsed = urlsplit(value)
        if parsed.scheme == "file":
            return f"file:{Path(parsed.path).resolve()}"
        if not parsed.hostname:
            raise ReleaseError(f"无法识别 origin URL：{url!r}")
        host = parsed.hostname.lower()
        path = parsed.path
    else:
        scp_match = re.fullmatch(r"(?:[^@/]+@)?([^:/]+):(.+)", value)
        if scp_match:
            host = scp_match.group(1).lower()
            path = scp_match.group(2)
        else:
            return f"file:{Path(value).expanduser().resolve()}"
    normalized_path = path.strip("/")
    if normalized_path.endswith(".git"):
        normalized_path = normalized_path[:-4]
    if not normalized_path:
        raise ReleaseError(f"origin URL 缺少仓库路径：{url!r}")
    return f"{host}/{normalized_path.lower()}"


def parse_stable_semver(version: str) -> tuple[int, int, int]:
    """Parse canonical ``MAJOR.MINOR.PATCH`` without leading zeroes."""

    match = STABLE_SEMVER_PATTERN.fullmatch(version)
    if not match:
        raise ReleaseError(
            f"版本号必须是 canonical stable SemVer（例如 0.2.1）：{version!r}"
        )
    return tuple(int(part) for part in match.groups())  # type: ignore[return-value]


def _load_json_text(text: str, label: str) -> dict:
    try:
        value = json.loads(text)
    except json.JSONDecodeError as error:
        raise ReleaseError(f"{label} 不是有效 JSON：{error}") from error
    if not isinstance(value, dict):
        raise ReleaseError(f"{label} 顶层必须是 JSON object。")
    return value


def _read_init_version(text: str) -> str:
    matches = [match for match in INIT_VERSION_PATTERN.finditer(text) if match.group(2) == match.group(4)]
    if len(matches) != 1:
        raise ReleaseError(
            f"hr_toolkit/__init__.py 中应有且仅有一个 __version__，实际找到 {len(matches)} 个。"
        )
    return matches[0].group(3)


def _configured_versions_from_texts(
    init_text: str,
    package_text: str,
    package_lock_text: str,
) -> ConfiguredVersions:
    package = _load_json_text(package_text, "package.json")
    package_lock = _load_json_text(package_lock_text, "package-lock.json")
    packages = package_lock.get("packages")
    workspace = packages.get("") if isinstance(packages, dict) else None
    if not isinstance(workspace, dict):
        raise ReleaseError('package-lock.json 缺少 packages[""] workspace。')

    values = {
        "hr_toolkit/__init__.py": _read_init_version(init_text),
        "package.json": package.get("version"),
        "package-lock.json": package_lock.get("version"),
        'package-lock.json packages[""]': workspace.get("version"),
    }
    if any(not isinstance(value, str) or not value for value in values.values()):
        raise ReleaseError(f"无法读取全部版本字段：{values}")
    typed_values = {key: str(value) for key, value in values.items()}
    for value in typed_values.values():
        parse_stable_semver(value)
    unique = set(typed_values.values())
    if len(unique) != 1:
        raise ReleaseError(f"版本文件不同步：{typed_values}")
    return ConfiguredVersions(current=next(iter(unique)), values=typed_values)


def read_configured_versions(root: Path = REPO_ROOT) -> ConfiguredVersions:
    return _configured_versions_from_texts(
        (root / VERSION_FILES[0]).read_text(encoding="utf-8"),
        (root / VERSION_FILES[1]).read_text(encoding="utf-8"),
        (root / VERSION_FILES[2]).read_text(encoding="utf-8"),
    )


def _replace_init_version(text: str, version: str) -> str:
    replacements = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal replacements
        if match.group(2) != match.group(4):
            return match.group(0)
        replacements += 1
        return f"{match.group(1)}{match.group(2)}{version}{match.group(4)}{match.group(5)}"

    updated = INIT_VERSION_PATTERN.sub(replace, text)
    if replacements != 1:
        raise ReleaseError(
            f"hr_toolkit/__init__.py 中应有且仅有一个可更新的 __version__，实际找到 {replacements} 个。"
        )
    return updated


def render_version_files(root: Path, version: str) -> Mapping[str, bytes]:
    """Render all version files in memory without changing the workspace."""

    parse_stable_semver(version)
    init_text = (root / VERSION_FILES[0]).read_text(encoding="utf-8")
    package_text = (root / VERSION_FILES[1]).read_text(encoding="utf-8")
    package_lock_text = (root / VERSION_FILES[2]).read_text(encoding="utf-8")

    package = _load_json_text(package_text, "package.json")
    package_lock = _load_json_text(package_lock_text, "package-lock.json")
    packages = package_lock.get("packages")
    workspace = packages.get("") if isinstance(packages, dict) else None
    if not isinstance(workspace, dict):
        raise ReleaseError('package-lock.json 缺少 packages[""] workspace。')

    package["version"] = version
    package_lock["version"] = version
    workspace["version"] = version
    rendered_texts = {
        VERSION_FILES[0]: _replace_init_version(init_text, version),
        VERSION_FILES[1]: json.dumps(package, ensure_ascii=False, indent=2) + "\n",
        VERSION_FILES[2]: json.dumps(package_lock, ensure_ascii=False, indent=2) + "\n",
    }
    rendered_versions = _configured_versions_from_texts(
        rendered_texts[VERSION_FILES[0]],
        rendered_texts[VERSION_FILES[1]],
        rendered_texts[VERSION_FILES[2]],
    )
    if rendered_versions.current != version:
        raise ReleaseError(f"版本渲染校验失败：{rendered_versions.values}")
    return {relative: text.encode("utf-8") for relative, text in rendered_texts.items()}


def snapshot_version_files(root: Path) -> Mapping[str, bytes]:
    return {relative: (root / relative).read_bytes() for relative in VERSION_FILES}


def write_version_files(root: Path, rendered: Mapping[str, bytes]) -> None:
    if set(rendered) != VERSION_FILE_SET:
        raise ReleaseError(f"拒绝写入版本白名单以外的文件：{set(rendered) - VERSION_FILE_SET}")
    for relative in VERSION_FILES:
        (root / relative).write_bytes(rendered[relative])


def restore_version_files(root: Path, snapshots: Mapping[str, bytes]) -> None:
    if set(snapshots) != VERSION_FILE_SET:
        raise ManualRecoveryRequired("版本快照不完整，拒绝自动恢复。")
    for relative in VERSION_FILES:
        (root / relative).write_bytes(snapshots[relative])


def prepare_release(
    target_version: str,
    root: Path = REPO_ROOT,
    git: Optional[GitRepository] = None,
) -> ReleasePlan:
    target_tuple = parse_stable_semver(target_version)
    repository = git or GitRepository(root)

    branch = repository.branch()
    if branch != MAIN_BRANCH:
        raise ReleaseError(f"发布必须在 main 分支执行，当前分支为 {branch or 'detached HEAD'}。")
    repository.ensure_no_git_operation()
    repository.ensure_clean()
    remote_url = repository.ensure_origin()
    repository.fetch_origin(remote_url)
    repository.ensure_clean()

    start_head = repository.head()
    remote_head = repository.origin_main()
    if start_head != remote_head:
        raise ReleaseError(
            "本地 main 必须与 origin/main 完全一致。"
            f"\n本地：{start_head}\n远端：{remote_head}"
        )

    configured = read_configured_versions(root)
    current_tuple = parse_stable_semver(configured.current)
    if target_tuple <= current_tuple:
        raise ReleaseError(
            f"目标版本必须高于当前版本：{configured.current} -> {target_version}"
        )

    tag = f"v{target_version}"
    if repository.remote_tag_exists(tag, remote_url):
        raise ReleaseError(f"远端 Tag 已存在：{tag}")
    if repository.local_tag_exists(tag):
        raise ReleaseError(f"本地 Tag 已存在：{tag}")

    return ReleasePlan(
        root=root,
        current_version=configured.current,
        target_version=target_version,
        start_head=start_head,
        tag=tag,
        remote_url=remote_url,
    )


def revalidate_release_plan(plan: ReleasePlan, git: GitRepository) -> None:
    """Recheck every mutable precondition after interactive confirmation."""

    branch = git.branch()
    if branch != MAIN_BRANCH:
        raise ReleaseError(
            f"确认期间分支发生变化，发布已停止：{branch or 'detached HEAD'}。"
        )
    git.ensure_no_git_operation()
    git.ensure_clean()
    if git.head() != plan.start_head:
        raise ReleaseError("确认期间本地 HEAD 发生变化，发布已停止。")

    remote_url = git.ensure_origin()
    if canonical_remote_identity(remote_url) != canonical_remote_identity(plan.remote_url):
        raise ReleaseError("确认期间 origin 目标发生变化，发布已停止。")
    git.fetch_origin(plan.remote_url)
    git.ensure_clean()
    if git.branch() != MAIN_BRANCH or git.head() != plan.start_head:
        raise ReleaseError("二次检查期间本地 main 或 HEAD 发生变化，发布已停止。")
    if git.origin_main() != plan.start_head:
        raise ReleaseError("确认期间 origin/main 已发生变化，请同步后重新发布。")

    configured = read_configured_versions(plan.root)
    if configured.current != plan.current_version:
        raise ReleaseError(
            "确认期间版本文件发生变化，发布已停止："
            f"{configured.current} != {plan.current_version}"
        )
    if git.remote_tag_exists(plan.tag, plan.remote_url):
        raise ReleaseError(f"确认期间远端 Tag 已出现：{plan.tag}")
    if git.local_tag_exists(plan.tag):
        raise ReleaseError(f"确认期间本地 Tag 已出现：{plan.tag}")


def run_full_checks(root: Path = REPO_ROOT) -> None:
    """Run every local release check without building an installer."""

    runner = CommandRunner(root)
    commands = (
        [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"],
        [
            sys.executable,
            "-m",
            "compileall",
            "-f",
            "-q",
            "hr_toolkit",
            "scripts",
            "tests",
            "hr_toolkit_app.py",
            "hr_toolkit_updater.py",
        ],
        ["git", "diff", "--check"],
    )
    for command in commands:
        print(f"+ {shlex.join(command)}", flush=True)
        runner.run(command, capture=False)


def _assert_only_version_changes(git: GitRepository) -> None:
    staged = git.staged_paths()
    unstaged = git.unstaged_paths()
    untracked = git.untracked_paths()
    if staged or unstaged != VERSION_FILE_SET or untracked:
        raise ReleaseError(
            "版本更新后的文件集合不符合白名单。"
            f"\n暂存：{sorted(staged)}"
            f"\n未暂存：{sorted(unstaged)}"
            f"\n未跟踪：{sorted(untracked)}"
        )


def _assert_exact_staged_files(git: GitRepository) -> None:
    staged = git.staged_paths()
    unstaged = git.unstaged_paths()
    untracked = git.untracked_paths()
    if staged != VERSION_FILE_SET or unstaged or untracked:
        raise ReleaseError(
            "提交前的文件集合不符合版本白名单。"
            f"\n期望暂存：{sorted(VERSION_FILE_SET)}"
            f"\n实际暂存：{sorted(staged)}"
            f"\n未暂存：{sorted(unstaged)}"
            f"\n未跟踪：{sorted(untracked)}"
        )


def _assert_version_file_bytes(
    root: Path,
    expected: Mapping[str, bytes],
    *,
    context: str,
) -> None:
    if set(expected) != VERSION_FILE_SET:
        raise ReleaseError(f"{context}：版本文件身份集合不完整。")
    mismatched = [
        relative
        for relative in VERSION_FILES
        if (root / relative).read_bytes() != expected[relative]
    ]
    if mismatched:
        raise ReleaseError(f"{context}：版本文件内容发生并发变化：{mismatched}")


def _assert_committed_version_bytes(
    git: GitRepository,
    commit: str,
    expected: Mapping[str, bytes],
) -> None:
    mismatched = [
        relative
        for relative in VERSION_FILES
        if git.committed_file_bytes(commit, relative) != expected[relative]
    ]
    if mismatched:
        raise ReleaseError(f"release commit 内容与渲染版本不一致：{mismatched}")


def _assert_staged_version_bytes(
    git: GitRepository,
    expected: Mapping[str, bytes],
) -> None:
    mismatched = [
        relative
        for relative in VERSION_FILES
        if git.staged_file_bytes(relative) != expected[relative]
    ]
    if mismatched:
        raise ReleaseError(f"暂存区内容与渲染版本不一致：{mismatched}")


def classify_remote_push_state(
    refs: Mapping[str, str],
    state: LocalReleaseState,
) -> RemotePushState:
    main_ref = f"refs/heads/{MAIN_BRANCH}"
    tag_ref = f"refs/tags/{state.tag}"
    peeled_tag_ref = f"{tag_ref}^{{}}"
    remote_main = refs.get(main_ref)
    remote_tag = refs.get(tag_ref)
    remote_peeled_tag = refs.get(peeled_tag_ref)

    if (
        remote_main == state.start_head
        and remote_tag is None
        and remote_peeled_tag is None
    ):
        return RemotePushState.UNCHANGED
    if (
        state.release_commit is not None
        and state.tag_object is not None
        and remote_main == state.release_commit
        and remote_tag == state.tag_object
        and remote_peeled_tag == state.release_commit
    ):
        return RemotePushState.APPLIED
    return RemotePushState.UNKNOWN


def rollback_local_release(git: GitRepository, state: LocalReleaseState) -> None:
    """Undo only state created by this invocation, after identity checks."""

    if git.branch() != MAIN_BRANCH:
        raise ManualRecoveryRequired("本地分支已发生变化，拒绝自动回滚。")
    current_head = git.head()
    expected_head = state.release_commit or state.start_head
    if current_head != expected_head:
        raise ManualRecoveryRequired(
            "本地 HEAD 已发生意外变化，拒绝自动回滚。"
            f"\n预期：{expected_head}\n实际：{current_head}"
        )

    if state.tag_created:
        current_tag_object = git.local_ref_oid(f"refs/tags/{state.tag}")
        if state.tag_object is None or current_tag_object != state.tag_object:
            raise ManualRecoveryRequired(
                f"本地 Tag {state.tag} 的对象身份不明确，拒绝自动删除。"
            )

    try:
        _assert_version_file_bytes(
            state.root,
            state.rendered,
            context="回滚前身份检查",
        )
    except Exception as error:
        raise ManualRecoveryRequired(str(error)) from error

    staged = git.staged_paths()
    unstaged = git.unstaged_paths()
    untracked = git.untracked_paths()
    if untracked or staged - VERSION_FILE_SET or unstaged - VERSION_FILE_SET:
        raise ManualRecoveryRequired(
            "回滚前发现本次事务之外的 Git 变化，拒绝自动回滚。"
            f"\n暂存：{sorted(staged)}"
            f"\n未暂存：{sorted(unstaged)}"
            f"\n未跟踪：{sorted(untracked)}"
        )
    if state.release_commit is not None and (staged or unstaged):
        raise ManualRecoveryRequired(
            "release commit 创建后工作区又发生变化，拒绝自动回滚。"
        )

    try:
        if state.tag_created:
            git.delete_tag(state.tag)
        if state.release_commit is not None:
            git.reset_mixed(state.start_head)
        elif state.staged:
            git.restore_staged_version_files()
        if state.files_written:
            restore_version_files(state.root, state.snapshots)
    except BaseException as error:
        raise ManualRecoveryRequired(f"自动回滚未能完整完成：{error}") from error

    if git.head() != state.start_head:
        raise ManualRecoveryRequired("回滚后 HEAD 未恢复到发布前提交。")
    if state.tag_created and git.local_ref_oid(f"refs/tags/{state.tag}") is not None:
        raise ManualRecoveryRequired(f"回滚后 Tag 仍存在：{state.tag}")
    for relative, expected in state.snapshots.items():
        if (state.root / relative).read_bytes() != expected:
            raise ManualRecoveryRequired(f"回滚后文件内容未恢复：{relative}")


def _manual_recovery_message(state: LocalReleaseState, reason: str) -> str:
    return (
        f"{reason}\n"
        "为避免误删已推送或他人创建的引用，已保留本地 release commit/Tag。请人工核对：\n"
        f"  git ls-remote origin refs/heads/{MAIN_BRANCH} refs/tags/{state.tag} refs/tags/{state.tag}^{{}}\n"
        f"  git show --stat {state.release_commit or 'HEAD'}\n"
        f"发布前 HEAD：{state.start_head}\n"
        f"本地 release commit：{state.release_commit or '未知'}\n"
        f"本地 Tag：{state.tag}"
    )


def resolve_failed_push(
    git: GitRepository,
    state: LocalReleaseState,
    push_error: BaseException,
) -> RemotePushState:
    try:
        refs = git.remote_release_refs(state.tag, state.remote_url)
    except BaseException as inspect_error:
        raise ManualRecoveryRequired(
            _manual_recovery_message(
                state,
                f"atomic push 失败，且无法核验远端引用：{push_error}\n核验错误：{inspect_error}",
            )
        ) from push_error

    remote_state = classify_remote_push_state(refs, state)
    if remote_state is RemotePushState.APPLIED:
        print(
            "atomic push 命令虽返回失败，但远端 main 与 annotated tag 均已正确更新；按发布成功处理。",
            file=sys.stderr,
        )
        return remote_state
    if remote_state is RemotePushState.UNCHANGED:
        rollback_local_release(git, state)
        raise ReleaseError(f"atomic push 失败，远端两个引用均未变化；本地发布状态已安全回滚。\n{push_error}") from push_error

    raise ManualRecoveryRequired(
        _manual_recovery_message(
            state,
            f"atomic push 失败，远端引用呈现部分更新或未知状态：{push_error}",
        )
    ) from push_error


def assert_release_refs_ready(git: GitRepository, state: LocalReleaseState) -> None:
    if state.release_commit is None or state.tag_object is None or not state.tag_created:
        raise ReleaseError("发布引用尚未完整创建，拒绝推送。")
    if git.branch() != MAIN_BRANCH or git.head() != state.release_commit:
        raise ManualRecoveryRequired("推送前分支或 HEAD 已发生变化，拒绝推送。")
    git.ensure_no_git_operation()
    git.ensure_clean()
    _assert_version_file_bytes(
        state.root,
        state.rendered,
        context="推送前身份检查",
    )
    _assert_committed_version_bytes(git, state.release_commit, state.rendered)
    if git.local_ref_oid(f"refs/tags/{state.tag}") != state.tag_object:
        raise ManualRecoveryRequired("推送前 annotated Tag object 已发生变化。")
    if git.local_ref_oid(f"refs/tags/{state.tag}^{{}}") != state.release_commit:
        raise ManualRecoveryRequired("推送前 annotated Tag 未指向 release commit。")


def verify_successful_push(git: GitRepository, state: LocalReleaseState) -> None:
    try:
        refs = git.remote_release_refs(state.tag, state.remote_url)
    except BaseException as error:
        raise ManualRecoveryRequired(
            _manual_recovery_message(
                state,
                f"atomic push 返回成功，但无法核验远端引用：{error}",
            )
        ) from error
    if classify_remote_push_state(refs, state) is not RemotePushState.APPLIED:
        raise ManualRecoveryRequired(
            _manual_recovery_message(
                state,
                "atomic push 返回成功，但远端引用与本次 release commit/Tag 不一致。",
            )
        )
    try:
        assert_release_refs_ready(git, state)
    except BaseException as error:
        raise ManualRecoveryRequired(
            _manual_recovery_message(
                state,
                f"远端发布已正确落地，但本地引用在推送后发生变化：{error}",
            )
        ) from error


def execute_release_plan(
    plan: ReleasePlan,
    git: GitRepository,
    *,
    dry_run: bool,
    checks: Callable[[Path], None] = run_full_checks,
) -> None:
    revalidate_release_plan(plan, git)
    rendered = render_version_files(plan.root, plan.target_version)
    if dry_run:
        checks(plan.root)
        revalidate_release_plan(plan, git)
        print(
            f"Dry run 通过：{plan.current_version} -> {plan.target_version}。"
            "未写入版本文件，未创建 commit/Tag，未执行 push。"
        )
        return

    state = LocalReleaseState(
        root=plan.root,
        start_head=plan.start_head,
        tag=plan.tag,
        snapshots=snapshot_version_files(plan.root),
        rendered=rendered,
        remote_url=plan.remote_url,
    )
    try:
        state.files_written = True
        write_version_files(plan.root, rendered)
        configured = read_configured_versions(plan.root)
        if configured.current != plan.target_version:
            raise ReleaseError(f"写入后的版本校验失败：{configured.values}")

        checks(plan.root)
        configured = read_configured_versions(plan.root)
        if configured.current != plan.target_version:
            raise ReleaseError(f"检查执行后版本文件发生变化：{configured.values}")
        if git.branch() != MAIN_BRANCH or git.head() != plan.start_head:
            raise ReleaseError("检查执行期间分支或 HEAD 发生变化，发布已停止。")
        _assert_version_file_bytes(
            plan.root,
            rendered,
            context="暂存前身份检查",
        )
        _assert_only_version_changes(git)
        state.staged = True
        git.stage_version_files()
        _assert_version_file_bytes(
            plan.root,
            rendered,
            context="暂存后身份检查",
        )
        _assert_exact_staged_files(git)
        _assert_staged_version_bytes(git, rendered)
        git.check_staged_diff()

        message = f"chore(recruitment): 发布 {plan.target_version}"
        state.release_commit = git.commit(message)
        state.staged = False
        if git.commit_parents(state.release_commit) != (state.start_head,):
            raise ReleaseError("release commit 必须只有发布前 HEAD 一个父提交。")
        committed_paths = git.commit_paths(state.release_commit)
        if committed_paths != VERSION_FILE_SET:
            raise ReleaseError(
                f"release commit 包含白名单以外的文件：{sorted(committed_paths)}"
            )
        _assert_committed_version_bytes(git, state.release_commit, rendered)
        git.ensure_clean()
        configured = read_configured_versions(plan.root)
        if configured.current != plan.target_version:
            raise ReleaseError(f"release commit 中的版本校验失败：{configured.values}")

        if git.branch() != MAIN_BRANCH or git.head() != state.release_commit:
            raise ReleaseError("创建 Tag 前分支或 HEAD 发生变化。")
        git.create_annotated_tag(plan.tag, state.release_commit)
        state.tag_created = True
        state.tag_object = git.local_ref_oid(f"refs/tags/{plan.tag}")
        if not state.tag_object:
            raise ReleaseError(f"无法读取新建 annotated tag：{plan.tag}")
        if git.local_ref_oid(f"refs/tags/{plan.tag}^{{}}") != state.release_commit:
            raise ReleaseError(f"annotated tag 未指向 release commit：{plan.tag}")
    except BaseException as error:
        if state.has_mutations:
            rollback_local_release(git, state)
            raise ReleaseError(f"发布失败，本地发布状态已回滚：{error}") from error
        raise

    try:
        assert_release_refs_ready(git, state)
        # Resolve origin once more for diagnostics, but push through the URL
        # pinned by the post-confirmation validation above.
        if canonical_remote_identity(git.ensure_origin()) != canonical_remote_identity(plan.remote_url):
            raise ReleaseError("测试期间 origin 目标发生变化，拒绝推送。")
        git.push_atomic(
            tag=plan.tag,
            remote_url=plan.remote_url,
            release_commit=state.release_commit,
            tag_object=state.tag_object,
        )
    except BaseException as push_error:
        remote_state = resolve_failed_push(git, state, push_error)
        if remote_state is not RemotePushState.APPLIED:
            raise AssertionError("未处理的 push 失败状态")
    else:
        verify_successful_push(git, state)

    print(f"发布 {plan.tag} 已原子推送；GitHub Actions 将负责构建与发布。")


def confirm_release(version: str, assume_yes: bool, dry_run: bool) -> bool:
    if dry_run or assume_yes:
        return True
    if not sys.stdin.isatty():
        raise ReleaseError("当前无法交互确认；审核版本后请使用 --yes。")
    answer = input(f"将 HR Toolkit v{version} 发布到 origin，是否继续？[y/N] ")
    return answer.strip().lower() in {"y", "yes"}


def release(
    version: str,
    *,
    dry_run: bool = False,
    assume_yes: bool = False,
    root: Path = REPO_ROOT,
) -> None:
    git = GitRepository(root)
    plan = prepare_release(version, root, git)
    print(
        f"发布计划：{plan.current_version} -> {plan.target_version}，"
        f"分支 {MAIN_BRANCH}，Tag {plan.tag}"
    )
    if not confirm_release(version, assume_yes, dry_run):
        print("发布已取消。")
        return
    execute_release_plan(plan, git, dry_run=dry_run)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="检查、提交并原子推送 HR Toolkit 版本；不在本机构建安装包。"
    )
    parser.add_argument("version", help="canonical stable SemVer，例如 0.2.1")
    parser.add_argument("--dry-run", action="store_true", help="运行全部检查，但不写版本或修改 Git 历史")
    parser.add_argument("--yes", action="store_true", help="跳过正式发布前的交互确认")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        release(args.version, dry_run=args.dry_run, assume_yes=args.yes)
    except ReleaseError as error:
        print(f"发布失败：{error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
