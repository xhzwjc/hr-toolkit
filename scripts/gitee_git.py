"""Run a Gitee Git transport operation with bounded transient-error retries.

Ref validation stays in the workflows. Repeating the same non-forced push is
safe even when the server accepted it but its response was lost; Git still
rejects divergent branches and conflicting tags on the next attempt.
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
import time
from typing import Sequence


RETRY_DELAYS = (15, 30, 60)
TRANSIENT_ERROR = re.compile(
    r"(?:HTTP(?:/[\d.]+)?\s+|returned error:\s*)(?:408|429|500|502|503|504)\b"
    r"|connection (?:reset|timed out)|operation timed out|could not resolve host"
    r"|couldn't connect to server|failed to connect|remote end hung up unexpectedly"
    r"|early EOF|SSL_ERROR_SYSCALL|TLS connection was non-properly terminated",
    re.IGNORECASE,
)
PERMANENT_ERROR = re.compile(
    r"authentication failed|access denied|permission denied|repository not found"
    r"|\[rejected\]|\[remote rejected\]|does not support --atomic",
    re.IGNORECASE,
)


def _write(text: str, stream) -> None:
    token = os.environ.get("GITEE_TOKEN", "")
    if token:
        text = text.replace(token, "***")
    if text:
        print(text, end="" if text.endswith("\n") else "\n", file=stream, flush=True)


def run_git(arguments: Sequence[str]) -> int:
    if not arguments or arguments[0] not in {"fetch", "ls-remote", "push"}:
        _write("Usage: gitee_git.py <fetch|ls-remote|push> [git arguments]", sys.stderr)
        return 2
    command = ["git", "-c", "http.version=HTTP/1.1"]
    if arguments[0] == "push":
        # Preserve the existing source-upload configuration.
        command += ["-c", "http.postBuffer=1073741824"]
    command += list(arguments)
    environment = dict(os.environ, LC_ALL="C", GIT_TERMINAL_PROMPT="0")
    attempts = len(RETRY_DELAYS) + 1
    for attempt in range(attempts):
        result = subprocess.run(
            command,
            env=environment,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if result.returncode == 0:
            # ls-remote stdout is machine-readable. Retry messages and Git
            # warnings must never become part of an object ID parsed by awk.
            _write(result.stdout, sys.stdout)
            _write(result.stderr, sys.stderr)
            return 0
        detail = result.stderr + "\n" + result.stdout
        _write(detail.strip(), sys.stderr)
        if PERMANENT_ERROR.search(detail) or not TRANSIENT_ERROR.search(detail):
            return result.returncode if result.returncode > 0 else 1
        if attempt == attempts - 1:
            _write(
                f"Gitee git {arguments[0]} failed after {attempts} attempts; "
                "source synchronization was not confirmed. Check the Gitee "
                "service, then rerun the source-release workflow for this tag.",
                sys.stderr,
            )
            return result.returncode if result.returncode > 0 else 1
        delay = RETRY_DELAYS[attempt]
        _write(
            f"Gitee git {arguments[0]} hit a transient network error; "
            f"retry {attempt + 2}/{attempts} in {delay}s.",
            sys.stderr,
        )
        time.sleep(delay)
    return 1


def main() -> int:
    return run_git(sys.argv[1:])


if __name__ == "__main__":
    raise SystemExit(main())
