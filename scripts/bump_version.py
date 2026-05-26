from __future__ import annotations

import argparse

from versioning import bump_project_version


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="自动递增 HR工具箱版本号")
    parser.add_argument("--bump", choices=["patch", "minor", "major"], default="patch")
    args = parser.parse_args(argv)

    new_version = bump_project_version(args.bump)
    print(f"版本号已更新为：{new_version}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
