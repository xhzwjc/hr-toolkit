"""Command Line Interface for HR Toolkit."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Callable

from hr_toolkit.runtime_checks import run_headless_command
from hr_toolkit.tools.registry import (
    ToolSpec,
    ensure_default_tools_registered,
    get_all_tools,
    get_tool_by_cli_command,
)

from .tools.folder_rename import (
    FILE_TYPE_FOLDER,
    MODE_EXCEL_BATCH,
    rename_files_by_excel,
    rename_person_folders,
)
from .tools.material_collector import (
    LIBRARY_MODE_FLAT_OCR,
    LIBRARY_MODE_PERSON_FOLDER,
    MODE_BY_EMPLOYEE,
    collect_employee_materials,
)
from .tools.archive_import import export_company_archive_tables, import_archive_transfers
from .tools.data_statistics import generate_data_statistics_reports
from .tools.insurance_ledger import generate_insurance_ledger
from .tools.personnel_change_merge import merge_personnel_changes, update_roster_from_change_summaries
from .tools.salary_merge import merge_monthly_salary
from .tools.salary_split import split_salary_by_company
from .tools.social_security import generate_social_security_reports


def _build_social_security_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("social-security", help="需求1：生成社保明细表和社保汇总表")
    p.add_argument("-i", "--input", required=True, nargs="+", type=Path, help="一个或多个社保缴费清单 .xlsx/.xls、ZIP/RAR/7Z/TAR 压缩包，或包含清单/压缩包的文件夹")
    p.add_argument("-r", "--roster", required=True, type=Path, help="参保人员花名册 .xlsx 或 .xls")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("--dry-run", action="store_true", help="只识别社保缴费记录，不生成 Excel 文件")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_data_statistics_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("data-statistics", help="需求2：生成考勤和周月报统计表")
    p.add_argument("-i", "--input", required=True, nargs="+", type=Path, help="一个或多个考勤/周报/月报 .xlsx/.xls、ZIP/RAR/7Z/TAR 压缩包，或包含这些文件/压缩包的文件夹")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("-s", "--staff", type=Path, help="可选应汇报人员名单 .xlsx/.xls；传入后可统计未写周报/月报人员")
    p.add_argument("--week-start", help="可选周报统计开始日期（如 2026-06-02）；只统计该范围内周一截止的周报，需与 --week-end 同时使用")
    p.add_argument("--week-end", help="可选周报统计结束日期（如 2026-06-30）；需与 --week-start 同时使用")
    p.add_argument("--month-start", help="可选月报统计开始日期（如 2026-06-01）；只统计落在该范围内的月报，需与 --month-end 同时使用")
    p.add_argument("--month-end", help="可选月报统计结束日期（如 2026-06-30）；需与 --month-start 同时使用")
    p.add_argument("--remark-unit", choices=["day", "hour"], default="day", help="考勤统计表备注中加班/调休的单位：day 按天（默认），hour 按小时")
    p.add_argument("--include-business-trip", action="store_true", help="在考勤统计表中新增公出列；默认不勾选（不加该列）")
    p.add_argument(
        "--include-workday-business-trip",
        action="store_true",
        help="在考勤统计表中新增出差列；仅统计工作日出差，不含休息日出差天数",
    )
    p.add_argument("--dry-run", action="store_true", help="只识别记录，不生成 Excel 文件")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_insurance_ledger_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("insurance-ledger", help="需求3：生成保险台账和人员增减预警")
    p.add_argument("-i", "--input", required=True, nargs="+", type=Path, help="一个或多个保单人员清单 .xlsx/.xls、ZIP/RAR/7Z/TAR 压缩包，或包含保单清单/压缩包的文件夹")
    p.add_argument("-r", "--roster", required=True, type=Path, help="需求6人力资源分析表 .xlsx 或 .xls，需包含“花名册”工作表")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("--dry-run", action="store_true", help="只识别保单人员，不生成 Excel 文件")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_salary_split_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("salary-split", help="需求4：将工资表按入职公司拆分为多个工作簿")
    p.add_argument("-i", "--input", required=True, type=Path, help="输入工资表 .xlsx 或 .xls，需包含汇总表和明细表")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("--dry-run", action="store_true", help="只识别分组，不生成 Excel 文件")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_salary_merge_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("salary-merge", help="需求5：合并多个月工资表，生成个人应发工资汇总")
    p.add_argument("-i", "--input-dir", required=True, nargs="+", type=Path, help="一个或多个 .xlsx/.xls 月度工资表、ZIP/RAR/7Z/TAR 压缩包，或包含月度工资表/压缩包的文件夹")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("-s", "--summary", type=Path, help="已有个人薪资汇总表；传入后只追加缺失月份，不覆盖已有金额")
    p.add_argument("--year", type=int, help="汇总年份，例如 2026；不填时自动根据工资表月份推断")
    p.add_argument("--dry-run", action="store_true", help="只识别文件、月份和人数，不生成 Excel 文件")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_change_merge_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("change-merge", help="需求6：汇总多个项目异动表")
    p.add_argument("-i", "--input-dir", required=True, nargs="+", type=Path, help="一个或多个 .xlsx/.xls 项目异动表、ZIP/RAR/7Z/TAR 压缩包，或包含项目异动表/压缩包的文件夹")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("--template", type=Path, help="可选已有异动汇总表文件或汇总表文件夹；会按异动日期写入对应月份")
    p.add_argument("--analysis-template", type=Path, help="可选人力资源分析表；传入后会同步更新其中的花名册")
    p.add_argument("--dry-run", action="store_true", help="只识别异动记录，不生成 Excel 文件")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_roster_update_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("roster-update", help="需求6：根据异动汇总表单独更新人力资源花名册")
    p.add_argument("-i", "--input", required=True, nargs="+", type=Path, help="一个或多个异动汇总表 .xlsx/.xls、ZIP/RAR/7Z/TAR 压缩包，或包含异动汇总表/压缩包的文件夹")
    p.add_argument("-r", "--roster", required=True, type=Path, help="人力资源花名册 .xlsx/.xls")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("--dry-run", action="store_true", help="只识别汇总表记录，不生成花名册")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_folder_rename_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("folder-rename", help="人员资料批量改名（支持 Excel 人名顺序）")
    p.add_argument("-r", "--root", required=True, type=Path, help="需要处理的人员文件夹所在目录")
    p.add_argument("--mode", required=True, choices=["append", "remove", "replace", MODE_EXCEL_BATCH], help="append=追加文字，remove=删除结尾文字，replace=修改名称，excel=按名单顺序改名")
    p.add_argument("--excel", type=Path, help="excel 模式必填：包含姓名列的 .xlsx/.xls 名单")
    p.add_argument("--text", default="", help="追加文字或要删除的结尾文字，例如：劳动合同、-劳动合同、_身份证")
    p.add_argument("--target", default="", help="指定单个项目/原名称；不填时 append/remove 处理全部匹配项")
    p.add_argument("--replacement", default="", help="replace 模式下的新名称")
    p.add_argument("--file-type", default="folder", choices=["folder", "pdf", "image", "document", "all"], help="要改名的类型：folder=文件夹，pdf=PDF，image=图片，document=文档，all=全部")
    p.add_argument("--apply", action="store_true", help="实际执行改名；不加时只预览")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_archive_import_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("archive-import", help="需求7：将项目档案移交表写入公司档案汇总表")
    p.add_argument("-i", "--input", required=True, type=Path, nargs="+", help="档案移交表 .xlsx/.xls、ZIP/RAR/7Z/TAR 压缩包，或包含多个移交表/压缩包的文件夹")
    p.add_argument("-t", "--target", type=Path, help="已有档案汇总表 .xlsx/.xls；不传时使用内置空模板")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("--dry-run", action="store_true", help="只识别记录，不生成 Excel 文件")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_archive_export_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("archive-export", help="需求7：按公司从档案汇总表生成独立档案表")
    p.add_argument("-s", "--summary", required=True, nargs="+", type=Path, help="一个或多个档案汇总表 .xlsx/.xls、ZIP/RAR/7Z/TAR 压缩包，或包含档案汇总表的文件夹")
    p.add_argument("-e", "--existing", nargs="+", type=Path, help="可选已有公司档案表文件、ZIP/RAR/7Z/TAR 压缩包或文件夹；匹配到公司则追加，未匹配则用内置空模板新建")
    p.add_argument("-o", "--output", required=True, type=Path, help="输出目录")
    p.add_argument("--dry-run", action="store_true", help="只识别公司，不生成 Excel 文件")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


def _build_material_collector_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser("material-collector", help="需求9：员工资料自动打包与信息提取")
    p.add_argument("-l", "--library", required=True, type=Path, help="员工资料库根目录")
    p.add_argument("-r", "--roster", required=True, help="员工名单 Excel 文件路径或名单文本")
    p.add_argument("-o", "--output", required=True, type=Path, help="提取文件输出目录")
    p.add_argument("-m", "--materials", nargs="*", help="指定需要提取的材料类型列表（如 身份证 劳动合同 学历证明）")
    p.add_argument("--mode", choices=["by_employee", "by_material", "flat"], default="by_employee", help="归类方式：by_employee（按员工）、by_material（按材料）、flat（平铺）")
    p.add_argument(
        "--library-mode",
        choices=[LIBRARY_MODE_PERSON_FOLDER, LIBRARY_MODE_FLAT_OCR],
        default=LIBRARY_MODE_PERSON_FOLDER,
        help="资料库形式：person_folder（按人员文件夹，原模式）或 flat_ocr（无序平铺 OCR 索引）",
    )
    p.add_argument("--zip", action="store_true", help="自动生成 .zip 压缩包")
    p.add_argument("--no-report", action="store_true", help="不生成 Excel 汇总报告")
    p.add_argument("--no-ocr-cache", action="store_true", help="关闭 OCR 智能索引缓存（默认开启；关闭后所有图片都将实时 OCR）")
    p.add_argument("--json", action="store_true", help="以 JSON 输出执行结果，便于 ScriptHub/Web 集成")


PARSER_BUILDERS: list[Callable[[argparse._SubParsersAction], None]] = [
    _build_social_security_parser,
    _build_data_statistics_parser,
    _build_insurance_ledger_parser,
    _build_salary_split_parser,
    _build_salary_merge_parser,
    _build_change_merge_parser,
    _build_roster_update_parser,
    _build_folder_rename_parser,
    _build_archive_import_parser,
    _build_archive_export_parser,
    _build_material_collector_parser,
]


def build_parser() -> argparse.ArgumentParser:
    ensure_default_tools_registered()
    parser = argparse.ArgumentParser(
        prog="hr-toolkit",
        description="人事 Excel 自动化工具箱",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    for builder in PARSER_BUILDERS:
        builder(subparsers)
    return parser


def main(argv: list[str] | None = None) -> int:
    raw_argv = list(sys.argv[1:] if argv is None else argv)
    headless_result = run_headless_command(raw_argv)
    if headless_result is not None:
        return headless_result
    parser = build_parser()
    args = parser.parse_args(raw_argv)

    if args.command == "social-security":
        result = generate_social_security_reports(
            input_path=args.input,
            roster_path=args.roster,
            output_dir=args.output,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_social_security_summary(payload)
        return 0

    if args.command == "data-statistics":
        result = generate_data_statistics_reports(
            input_path=args.input,
            output_dir=args.output,
            report_staff_path=args.staff,
            week_start=args.week_start,
            week_end=args.week_end,
            month_start=args.month_start,
            month_end=args.month_end,
            remark_unit=args.remark_unit,
            include_business_trip=args.include_business_trip,
            include_workday_business_trip=args.include_workday_business_trip,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_data_statistics_summary(payload)
        return 0

    if args.command == "insurance-ledger":
        result = generate_insurance_ledger(
            input_path=args.input,
            roster_path=args.roster,
            output_dir=args.output,
            dry_run=args.dry_run,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_insurance_ledger_summary(payload)
        return 0

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
        if args.mode == MODE_EXCEL_BATCH:
            if args.excel is None:
                parser.error("folder-rename --mode excel 必须同时提供 --excel 名单文件")
            result = rename_files_by_excel(
                root_dir=args.root,
                excel_path=args.excel,
                file_type=args.file_type,
                dry_run=not args.apply,
            )
        else:
            result = rename_person_folders(
                root_dir=args.root,
                mode=args.mode,
                text=args.text,
                target_name=args.target,
                replacement_name=args.replacement,
                file_type=args.file_type,
                dry_run=not args.apply,
            )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_folder_rename_summary(payload)
        return 0

    if args.command == "material-collector":
        result = collect_employee_materials(
            library_dir=args.library,
            output_dir=args.output,
            roster_source=args.roster,
            material_types=args.materials if args.materials else None,
            mode=args.mode,
            library_mode=args.library_mode,
            create_zip=args.zip,
            generate_report=not args.no_report,
            use_ocr_cache=not args.no_ocr_cache,
        )
        payload = result.to_dict()
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            _print_material_collector_summary(payload)
        return 0

    parser.print_help(sys.stderr)
    return 2


def _print_social_security_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"输入：{payload['input_path']}")
    print(f"参保人员花名册：{payload['roster_path']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别文件数：{payload['source_file_count']}")
    print(f"识别缴费记录数：{payload['source_record_count']}")
    print(f"生成明细行数：{payload['detail_record_count']}")
    print(f"识别人员数：{payload['employee_count']}")
    for account, count in payload["account_counts"].items():
        print(f"- {account}: {count} 人")
    for period, count in payload["period_counts"].items():
        print(f"- {period}: {count} 行")
    if payload.get("detail_output_file"):
        print(f"社保明细表：{payload['detail_output_file']}")
    if payload.get("detail_output_files"):
        print("按参保单位/参保地拆分明细：")
        for output_file in payload["detail_output_files"]:
            print(f"- {output_file}")
    if payload.get("summary_output_file"):
        print(f"社保汇总表：{payload['summary_output_file']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")


def _print_data_statistics_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"输入：{payload['input_dir']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别文件数：{payload['source_file_count']}")
    print(f"考勤原始记录数：{payload['attendance_source_count']}")
    print(f"考勤统计人数：{payload['attendance_person_count']}")
    print(f"考勤异常明细数：{payload['attendance_exception_count']}")
    print(f"周报记录数：{payload['weekly_record_count']}")
    print(f"月报记录数：{payload['monthly_record_count']}")
    if payload.get("report_staff_path"):
        print(f"应汇报人员名单：{payload['report_staff_path']}")
        print(f"应汇报人数：{payload['expected_reporter_count']}")
    print(f"周月报异常人数：{payload['report_person_count']}")
    print(f"周月报异常明细数：{payload['report_exception_count']}")
    if payload.get("output_file"):
        print(f"输出文件：{payload['output_file']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")


def _print_insurance_ledger_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"输入：{payload['input_path']}")
    print(f"人力资源分析表：{payload['roster_path']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"模式：{'预览' if payload['dry_run'] else '生成文件'}")
    print(f"识别文件数：{payload['source_file_count']}")
    print(f"识别保单数：{payload['policy_count']}")
    print(f"保单人员数：{payload['insured_person_count']}")
    print(f"花名册在职人数：{payload['roster_person_count']}")
    print(f"需加保预警：{payload['add_warning_count']}")
    print(f"需减保预警：{payload['reduce_warning_count']}")
    if payload.get("output_file"):
        print(f"输出文件：{payload['output_file']}")
    if payload.get("roster_warning_file"):
        print(f"花名册预警输出文件：{payload['roster_warning_file']}")
    if payload["warnings"]:
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")


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


def _print_material_collector_summary(payload: dict) -> None:
    print(f"工具：{payload['tool_name']}")
    print(f"资料库目录：{payload['library_dir']}")
    print(f"输出目录：{payload['output_dir']}")
    print(f"归类方式：{payload['mode']}")
    print(f"资料库形式：{payload.get('library_mode', LIBRARY_MODE_PERSON_FOLDER)}")
    print(f"目标员工数：{payload['total_employees']} 人")
    print(f"材料齐全员工数：{payload['complete_employee_count']} 人")
    print(f"提取文件总数：{payload['matched_file_count']} 个")
    if payload.get("ocr_cache_enabled"):
        hits = payload.get("ocr_cache_hits", 0)
        misses = payload.get("ocr_cache_misses", 0)
        print(f"OCR 缓存：命中 {hits} 次，实时识别 {misses} 次"
              + (f"（缓存文件：{payload['ocr_cache_path']}）" if payload.get("ocr_cache_path") else ""))
        if payload.get("ocr_cache_skipped_reason"):
            print(f"OCR 缓存说明：{payload['ocr_cache_skipped_reason']}")
    else:
        print("OCR 缓存：已关闭（按 --no-ocr-cache 或 GUI 取消勾选）")
    if payload.get("zip_path"):
        print(f"压缩包文件：{payload['zip_path']}")
    if payload.get("report_path"):
        print(f"汇总报告：{payload['report_path']}")
    if payload.get("missing_records"):
        print("缺件明细：")
        for emp, missing in payload["missing_records"].items():
            print(f"- {emp} 缺少：{', '.join(missing)}")
    if payload.get("warnings"):
        print("提醒：")
        for warning in payload["warnings"]:
            print(f"- {warning}")
