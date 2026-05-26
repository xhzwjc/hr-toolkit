from __future__ import annotations

import re
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INIT_FILE = REPO_ROOT / "hr_toolkit" / "__init__.py"
VERSION_PATTERN = re.compile(r'(__version__\s*=\s*")([^"]+)(")')


def read_project_version(init_file: Path = INIT_FILE) -> str:
    match = VERSION_PATTERN.search(init_file.read_text(encoding="utf-8"))
    if not match:
        raise ValueError(f"未找到版本号：{init_file}")
    return match.group(2)


def bump_version(version: str, bump: str) -> str:
    major, minor, patch = _parse_version(version)
    if bump == "patch":
        patch += 1
    elif bump == "minor":
        minor += 1
        patch = 0
    elif bump == "major":
        major += 1
        minor = 0
        patch = 0
    else:
        raise ValueError("版本类型只能是 patch、minor、major。")
    return f"{major}.{minor}.{patch}"


def write_project_version(version: str, init_file: Path = INIT_FILE) -> None:
    content = init_file.read_text(encoding="utf-8")
    new_content, count = VERSION_PATTERN.subn(rf"\g<1>{version}\3", content, count=1)
    if count != 1:
        raise ValueError(f"未找到可替换的版本号：{init_file}")
    init_file.write_text(new_content, encoding="utf-8")


def bump_project_version(bump: str, init_file: Path = INIT_FILE) -> str:
    new_version = bump_version(read_project_version(init_file), bump)
    write_project_version(new_version, init_file)
    return new_version


def _parse_version(version: str) -> tuple[int, int, int]:
    parts = version.strip().split(".")
    if len(parts) != 3 or not all(part.isdigit() for part in parts):
        raise ValueError(f"版本号格式必须是 x.y.z：{version}")
    return int(parts[0]), int(parts[1]), int(parts[2])
