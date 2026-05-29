from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .tools.folder_rename import rename_person_folders
from .tools.archive_import import export_company_archive_tables, import_archive_transfers
from .tools.personnel_change_merge import merge_personnel_changes, update_roster_from_change_summaries
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
        help="输入工资表 .xlsx 或 .xls，需包含汇总表和明细表",
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
        nargs="+",
        type=Path,
        help="一个或多个 .xlsx/.xls 月度工资表、zip 压缩包，或包含月度工资表/压缩包的文件夹",
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
        nargs="+",
        type=Path,
        help="一个或多个 .xlsx/.xls 项目异动表、zip 压缩包，或包含项目异动表/压缩包的文件夹",
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
        help="可选已有异动汇总表文件或汇总表文件夹；会按异动日期写入对应月份",
    )
    change_merge.add_argument(
        "--analysis-template",
        type=Path,
        help="可选人力资源分析表；传入后会同步更新其中的花名册",
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

    roster_update = subparsers.add_parser(
        "roster-update",
        help="需求6：根据异动汇总表单独更新人力资源花名册",
    )
    roster_update.add_argument(
        "-i",
        "--input",
        required=True,
        nargs="+",
        type=Path,
        help="一个或多个异动汇总表 .xlsx/.xls，或包含异动汇总表的文件夹",
    )
    roster_update.add_argument(
        "-r",
        "--roster",
        required=True,
        type=Path,
        help="人力资源花名册 .xlsx/.xls",
    )
    roster_update.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
        help="输出目录",
    )
    roster_update.add_argument(
        "--dry-run",
        action="store_true",
        help="只识别汇总表记录，不生成花名册",
    )
    roster_update.add_argument(
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

    archive_import = subparsers.add_parser(
        "archive-import",
        help="需求7：将项目档案移交表写入公司档案汇总表",
    )
    archive_import.add_argument(
        "-i",
        "--input",
        required=True,
        type=Path,
        nargs="+",
        help="档案移交表 .xlsx/.xls、.zip，或包含多个移交表/压缩包的文件夹",
    )
    archive_import.add_argument(
        "-t",
        "--target",
        type=Path,
        help="已有档案汇总表 .xlsx/.xls；不传时使用内置空模板",
    )
    archive_import.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
        help="输出目录",
    )
    archive_import.add_argument(
        "--dry-run",
        action="store_true",
        help="只识别记录，不生成 Excel 文件",
    )
    archive_import.add_argument(
        "--json",
        action="store_true",
        help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成",
    )

    archive_export = subparsers.add_parser(
        "archive-export",
        help="需求7：按公司从档案汇总表生成独立档案表",
    )
    archive_export.add_argument(
        "-s",
        "--summary",
        required=True,
        nargs="+",
        type=Path,
        help="一个或多个档案汇总表 .xlsx/.xls，或包含档案汇总表的文件夹",
    )
    archive_export.add_argument(
        "-e",
        "--existing",
        nargs="+",
        type=Path,
        help="可选已有公司档案表文件或文件夹；匹配到公司则追加，未匹配则用内置空模板新建",
    )
    archive_export.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
        help="输出目录",
    )
    archive_export.add_argument(
        "--dry-run",
        action="store_true",
        help="只识别公司，不生成 Excel 文件",
    )
    archive_export.add_argument(
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
            analysis_template_path=args.analysis_template,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_change_merge_summary(payload)
        return 0

    if args.command == "archive-import":
        result = import_archive_transfers(
            input_path=args.input,
            target_path=args.target,
            output_dir=args.output,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_archive_import_summary(payload)
        return 0

    if args.command == "archive-export":
        result = export_company_archive_tables(
            summary_path=args.summary,
            output_dir=args.output,
            existing_archive_path=args.existing,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_archive_export_summary(payload)
        return 0

    if args.command == "roster-update":
        result = update_roster_from_change_summaries(
            summary_input=args.input,
            analysis_template_path=args.roster,
            output_dir=args.output,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_roster_update_summary(payload)
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
    print(f"输入：{payload['input_dir']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别文件数：{payload['source_file_count']}")
    print(f"异动记录数：{payload['record_count']}")
    print(f"写入模式：{'追加到已有汇总表' if payload.get('append_mode') else '新建干净汇总表'}")
    print(f"新增记录数：{payload.get('inserted_count', 0)}")
    print(f"补充已有记录数：{payload.get('updated_count', 0)}")
    print(f"已存在未修改记录数：{payload.get('skipped_count', 0)}")
    for sheet_name, count in payload["sheet_counts"].items():
        print(f"- {sheet_name}: {count} 条")
    if payload.get("output_files"):
        print("输出文件：")
        for output_file in payload["output_files"]:
            print(f"- {output_file}")
    elif payload.get("output_file"):
        print(f"输出文件：{payload['output_file']}")
    if payload.get("roster_output_file"):
        print(f"花名册输出文件：{payload['roster_output_file']}")
        print(f"花名册新增人数：{payload['roster_added_count']}")
        print(f"花名册标记离职人数：{payload['roster_marked_count']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")


def _print_roster_update_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"异动汇总表：{payload['summary_input']}")
    print(f"人力资源花名册：{payload['analysis_template_path']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别汇总表数：{payload['source_file_count']}")
    print(f"识别异动记录数：{payload['record_count']}")
    print(f"花名册新增人数：{payload['roster_added_count']}")
    print(f"花名册标记离职人数：{payload['roster_marked_count']}")
    for sheet_name, count in payload["sheet_counts"].items():
        print(f"- {sheet_name}: {count} 条")
    if payload.get("output_file"):
        print(f"输出文件：{payload['output_file']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")


def _print_archive_import_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"输入：{payload['input_path']}")
    if payload.get("target_path"):
        print(f"档案汇总表：{payload['target_path']}")
    else:
        print("档案汇总表：使用内置空模板")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别文件数：{payload['source_file_count']}")
    print(f"识别记录数：{payload['source_record_count']}")
    print(f"新增记录数：{payload['inserted_count']}")
    print(f"补充已有记录数：{payload['updated_count']}")
    print(f"已存在未修改记录数：{payload['skipped_count']}")
    for company, count in payload["company_counts"].items():
        print(f"- {company}: {count} 条")
    if payload.get("output_file"):
        print(f"输出文件：{payload['output_file']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")


def _print_archive_export_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"档案汇总表：{payload['summary_path']}")
    if payload.get("existing_archive_path"):
        print(f"已有公司档案表：{payload['existing_archive_path']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别公司数：{len(payload['company_counts'])}")
    print(f"新建公司档案表数：{payload['created_count']}")
    print(f"新增记录数：{payload['inserted_count']}")
    print(f"补充已有记录数：{payload['updated_count']}")
    print(f"已存在未修改记录数：{payload['skipped_count']}")
    for company, count in payload["company_counts"].items():
        print(f"- {company}: {count} 条")
    if payload.get("output_files"):
        print("输出文件：")
        for output_file in payload["output_files"]:
            print(f"- {output_file}")
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
