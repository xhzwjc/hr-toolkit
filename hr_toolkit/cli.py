from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .tools.salary_split import split_salary_by_company


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="hr-toolkit",
        description="人事 Excel 自动化工具箱",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    salary_split = subparsers.add_parser(
        "salary-split",
        help="需求4：将工资表按入职公司拆分为多个工作簿",
    )
    salary_split.add_argument(
        "-i",
        "--input",
        required=True,
        type=Path,
        help="输入工资表 .xlsx，需包含汇总表和明细表",
    )
    salary_split.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
        help="输出目录",
    )
    salary_split.add_argument(
        "--dry-run",
        action="store_true",
        help="只识别分组，不生成 Excel 文件",
    )
    salary_split.add_argument(
        "--json",
        action="store_true",
        help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "salary-split":
        result = split_salary_by_company(
            input_path=args.input,
            output_dir=args.output,
            dry_run=args.dry_run,
            write_manifest=not args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_salary_split_summary(payload)
        return 0

    parser.print_help(sys.stderr)
    return 2


def _print_salary_split_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"输入：{payload['input_path']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别公司数：{payload['company_count']}")
    print(f"识别人员数：{payload['employee_count']}")
    for item in payload["outputs"]:
        file_part = "" if not item.get("file_path") else f" -> {item['file_path']}"
        print(
            f"- {item['company']}: {item['employee_count']} 人，"
            f"{len(item['projects'])} 个项目{file_part}"
        )
