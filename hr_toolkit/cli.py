from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .tools.folder_rename import rename_person_folders
from .tools.personnel_change_merge import merge_personnel_changes
from .tools.salary_merge import merge_monthly_salary
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

    salary_merge = subparsers.add_parser(
        "salary-merge",
        help="需求5：合并多个月工资表，生成个人应发工资汇总",
    )
    salary_merge.add_argument(
        "-i",
        "--input-dir",
        required=True,
        type=Path,
        help="输入文件夹，内含多个 .xlsx 月度工资表",
    )
    salary_merge.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
        help="输出目录",
    )
    salary_merge.add_argument(
        "-s",
        "--summary",
        type=Path,
        help="已有个人薪资汇总表；传入后只追加缺失月份，不覆盖已有金额",
    )
    salary_merge.add_argument(
        "--year",
        type=int,
        help="汇总年份，例如 2026；不填时自动根据工资表月份推断",
    )
    salary_merge.add_argument(
        "--dry-run",
        action="store_true",
        help="只识别文件、月份和人数，不生成 Excel 文件",
    )
    salary_merge.add_argument(
        "--json",
        action="store_true",
        help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成",
    )

    change_merge = subparsers.add_parser(
        "change-merge",
        help="需求6：汇总多个项目异动表",
    )
    change_merge.add_argument(
        "-i",
        "--input-dir",
        required=True,
        type=Path,
        help="输入文件夹，内含多个 .xlsx 项目异动表",
    )
    change_merge.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
        help="输出目录",
    )
    change_merge.add_argument(
        "--template",
        type=Path,
        help="可选异动表模板；不填时使用输入文件夹中的第一份异动表作为模板",
    )
    change_merge.add_argument(
        "--dry-run",
        action="store_true",
        help="只识别异动记录，不生成 Excel 文件",
    )
    change_merge.add_argument(
        "--json",
        action="store_true",
        help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成",
    )

    folder_rename = subparsers.add_parser(
        "folder-rename",
        help="需求8：人员资料文件夹批量改名",
    )
    folder_rename.add_argument(
        "-r",
        "--root",
        required=True,
        type=Path,
        help="需要处理的人员文件夹所在目录",
    )
    folder_rename.add_argument(
        "--mode",
        required=True,
        choices=["append", "remove", "replace"],
        help="append=追加文字，remove=删除结尾文字，replace=修改单个文件夹名",
    )
    folder_rename.add_argument(
        "--text",
        default="",
        help="追加文字或要删除的结尾文字，例如：劳动合同、-劳动合同、_身份证",
    )
    folder_rename.add_argument(
        "--target",
        default="",
        help="指定单个人员/原文件夹名；不填时 append/remove 处理全部子文件夹",
    )
    folder_rename.add_argument(
        "--replacement",
        default="",
        help="replace 模式下的新文件夹名",
    )
    folder_rename.add_argument(
        "--apply",
        action="store_true",
        help="实际执行改名；不加时只预览",
    )
    folder_rename.add_argument(
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

    if args.command == "salary-merge":
        result = merge_monthly_salary(
            input_dir=args.input_dir,
            output_dir=args.output,
            existing_summary_path=args.summary,
            year=args.year,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_salary_merge_summary(payload)
        return 0

    if args.command == "change-merge":
        result = merge_personnel_changes(
            input_dir=args.input_dir,
            output_dir=args.output,
            template_path=args.template,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_change_merge_summary(payload)
        return 0

    if args.command == "folder-rename":
        result = rename_person_folders(
            root_dir=args.root,
            mode=args.mode,
            text=args.text,
            target_name=args.target,
            replacement_name=args.replacement,
            dry_run=not args.apply,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_folder_rename_summary(payload)
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


def _print_salary_merge_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"输入文件夹：{payload['input_dir']}")
    if payload.get("existing_summary_path"):
        print(f"已有汇总表：{payload['existing_summary_path']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别文件数：{payload['source_file_count']}")
    print(f"识别月份：{', '.join(payload['months'])}")
    print(f"识别人员数：{payload['employee_count']}")
    print(f"工资记录数：{payload['record_count']}")
    print(f"本次写入记录数：{payload['applied_record_count']}")
    print(f"已存在未覆盖记录数：{payload['skipped_record_count']}")
    if payload.get("output_file"):
        print(f"输出文件：{payload['output_file']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")


def _print_change_merge_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"输入文件夹：{payload['input_dir']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别文件数：{payload['source_file_count']}")
    print(f"异动记录数：{payload['record_count']}")
    for sheet_name, count in payload["sheet_counts"].items():
        print(f"- {sheet_name}: {count} 条")
    if payload.get("output_file"):
        print(f"输出文件：{payload['output_file']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")


def _print_folder_rename_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"目录：{payload['root_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '执行'}")
    print(f"改名数量：{payload['operation_count']}")
    for operation in payload["operations"]:
        print(f"- {operation['source_name']} -> {operation['target_name']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")
