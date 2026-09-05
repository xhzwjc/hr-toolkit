"""Declarative UI contracts and invocation validation for the Qt front end."""

from __future__ import annotations

import importlib
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from hr_toolkit.common.inputs import is_supported_archive_file


EXCEL_SUFFIXES = frozenset({".xlsx", ".xls"})


class FormValidationError(ValueError):
    def __init__(self, title: str, message: str) -> None:
        super().__init__(message)
        self.title = title
        self.message = message


@dataclass(frozen=True)
class ToolUiSpec:
    nav_id: str
    variant: str
    tool_id: str
    group: str
    title: str
    description: str
    input_label: str
    input_hint: str
    input_drop_title: str
    run_text: str
    log_text: str
    input_mode: str = "excel_archive_multi"
    support_id: str = ""
    support_label: str = ""
    support_button: str = "选择"
    support_mode: str = "excel_file"
    support_optional: bool = False
    fields: tuple[dict[str, Any], ...] = ()

    def as_dict(self) -> dict[str, Any]:
        return {
            "navId": self.nav_id,
            "variant": self.variant,
            "toolId": self.tool_id,
            "group": self.group,
            "title": self.title,
            "description": self.description,
            "inputLabel": self.input_label,
            "inputHint": self.input_hint,
            "inputDropTitle": self.input_drop_title,
            "runText": self.run_text,
            "logText": self.log_text,
            "inputMode": self.input_mode,
            "supportId": self.support_id,
            "supportLabel": self.support_label,
            "supportButton": self.support_button,
            "supportMode": self.support_mode,
            "supportOptional": self.support_optional,
            "fields": [dict(field) for field in self.fields],
        }


@dataclass(frozen=True)
class ToolInvocation:
    nav_id: str
    variant: str
    tool_id: str
    tool_name: str
    group_name: str
    function_module: str
    function_name: str
    args: tuple[Any, ...]
    kwargs: dict[str, Any]
    description: str
    preview: bool = False

    def resolve_function(self):
        module = importlib.import_module(self.function_module)
        return getattr(module, self.function_name)


MULTI_HINT = "支持 .xlsx / .xls / ZIP / RAR / 7Z / TAR / 文件夹 · 可多选"


SPECS: tuple[ToolUiSpec, ...] = (
    ToolUiSpec(
        "social_security", "default", "social_security", "社保与保险",
        "社保明细与汇总",
        "选择社保缴费清单、压缩包或文件夹，再选择参保人员花名册，自动生成明细和汇总。",
        "社保缴费清单", MULTI_HINT, "选择缴费清单、压缩包或文件夹", "生成报表",
        "请选择社保缴费清单和参保人员花名册，然后点击“生成报表”。资料和结果会自动留存在当前项目。",
        support_id="roster_path", support_label="参保人员花名册", support_button="选择花名册",
    ),
    ToolUiSpec(
        "insurance_ledger", "default", "insurance_ledger", "社保与保险",
        "保险台账与增减预警",
        "选择各保单人员清单、压缩包或文件夹，再选择需求6的人力资源分析表，自动生成保险台账。",
        "保单人员清单", MULTI_HINT, "选择保单清单、压缩包或文件夹", "生成台账",
        "请选择保单人员清单和人力资源分析表，然后点击“生成台账”。资料和结果会自动留存在当前项目。",
        support_id="roster_path", support_label="人力资源分析表", support_button="选择分析表",
    ),
    ToolUiSpec(
        "data_statistics", "default", "data_statistics", "考勤与统计",
        "考勤与周月报统计",
        "选择考勤结果、周报记录、月报记录，或包含这些文件的文件夹/压缩包，自动生成统计表和异常明细。",
        "考勤与周月报数据", MULTI_HINT, "选择考勤 / 周报 / 月报文件、压缩包或文件夹", "生成统计",
        "请选择考勤结果、周报记录和月报记录，然后点击“生成统计”。应汇报人员名单是可选项，资料和结果会自动留存在当前项目。",
        support_id="report_staff_path", support_label="应汇报人员名单（可选）", support_button="选择名单",
        support_optional=True,
        fields=(
            {
                "id": "week_range",
                "kind": "date_range",
                "label": "周报",
                "startId": "week_start",
                "endId": "week_end",
                "startPlaceholder": "如 2026-06-02",
                "endPlaceholder": "如 2026-06-30",
                "hint": "留空按整月统计",
                "presetGroup": "week",
                "presets": [
                    {"label": "本月", "value": "this_month"},
                    {"label": "上月", "value": "last_month"},
                    {"label": "本周", "value": "this_week"},
                    {"label": "上周", "value": "last_week"},
                    {"label": "清空", "value": "clear"},
                ],
            },
            {
                "id": "month_range",
                "kind": "date_range",
                "label": "月报",
                "startId": "month_start",
                "endId": "month_end",
                "startPlaceholder": "如 2026-06-01",
                "endPlaceholder": "如 2026-06-30",
                "hint": "留空不筛选月报",
                "presetGroup": "month",
                "presets": [
                    {"label": "本月", "value": "this_month"},
                    {"label": "上月", "value": "last_month"},
                    {"label": "清空", "value": "clear"},
                ],
            },
            {"id": "remark_unit", "kind": "choice", "label": "加班/调休单位", "default": "day", "options": [
                {"label": "按天", "value": "day"}, {"label": "按小时", "value": "hour"}
            ]},
            {"id": "include_business_trip", "kind": "check", "label": "新增「公出」列", "default": False},
            {"id": "include_workday_business_trip", "kind": "check", "label": "新增「出差」列", "default": False},
        ),
    ),
    ToolUiSpec(
        "salary_split", "default", "salary_split", "薪酬管理",
        "工资表按入职公司拆分",
        "选择一个包含“汇总表”和“明细表”的工资表，工具会按“入职公司”拆成多个公司文件。",
        "工资表文件", "支持 .xlsx / .xls · 单个文件", "选择工资表文件", "开始拆分",
        "请选择工资表文件，然后点击“开始拆分”。资料和结果会自动留存在当前项目。",
        input_mode="excel_single",
    ),
    ToolUiSpec(
        "salary_merge", "default", "salary_merge", "薪酬管理",
        "多月工资合并",
        "选择工资表文件、压缩包或文件夹；如已有汇总表，可一并选择后追加新月份。",
        "工资表文件", MULTI_HINT, "选择工资表、压缩包或文件夹", "开始合并",
        "请选择工资表文件、压缩包或文件夹，然后点击“开始合并”。已有汇总表是可选项，资料和结果会自动留存在当前项目。",
        support_id="existing_summary_path", support_label="已有汇总表（可选）", support_button="选择汇总表",
        support_optional=True,
    ),
    ToolUiSpec(
        "personnel_change_merge", "merge", "personnel_change_merge", "人员与档案",
        "异动表汇总与花名册",
        "选择异动表、压缩包或文件夹；如已有月度汇总表，可选择后按月份追加。",
        "异动表文件", MULTI_HINT, "选择异动表、压缩包或文件夹", "开始汇总",
        "请选择异动表文件或文件夹，然后点击“开始汇总”。已有汇总表是可选项，资料和结果会自动留存在当前项目。",
        support_id="template_path", support_label="已有汇总表/文件夹（可选）", support_button="选择汇总表",
        support_mode="excel_or_folder", support_optional=True,
    ),
    ToolUiSpec(
        "personnel_change_merge", "roster", "roster_update", "人员与档案",
        "异动表汇总与花名册", "选择异动汇总表和人力资源花名册，单独更新花名册。",
        "异动汇总表", MULTI_HINT, "选择异动汇总表、压缩包或文件夹", "更新花名册",
        "请选择异动汇总表和人力资源花名册，然后点击“更新花名册”。资料和结果会自动留存在当前项目。",
        support_id="analysis_template_path", support_label="人力资源花名册", support_button="选择花名册",
    ),
    ToolUiSpec(
        "archive_import", "import", "archive_import", "人员与档案",
        "档案入库与档案表",
        "选择项目档案移交表、压缩包或文件夹；可选已有档案汇总表，不选则新建。",
        "档案移交表", MULTI_HINT, "选择移交表、压缩包或文件夹", "开始入库",
        "请选择移交表文件、压缩包或文件夹，然后点击“开始入库”。已有档案汇总表是可选项，结果会自动留存在当前项目。",
        support_id="target_path", support_label="已有档案汇总表（可选）", support_button="选择汇总表",
        support_optional=True,
    ),
    ToolUiSpec(
        "archive_import", "export", "archive_export", "人员与档案",
        "档案入库与档案表",
        "选择档案汇总表、压缩包或文件夹，按公司写入已有档案表；没有已有表时自动新建。",
        "档案汇总表", MULTI_HINT, "选择档案汇总表、压缩包或文件夹", "生成档案表",
        "请选择档案汇总表、压缩包或文件夹，然后点击“生成档案表”。已有公司档案表是可选项，结果会自动留存在当前项目。",
        support_id="existing_archive_path", support_label="已有公司档案表（可选）", support_button="选择档案表",
        support_mode="excel_archive_or_folder", support_optional=True,
    ),
    ToolUiSpec(
        "material_collector", "default", "material_collector", "人员与档案",
        "员工资料智能检索与打包",
        "支持按人员文件夹查找，也支持从无序平铺资料库建立 OCR 索引后按人员检索。",
        "员工资料库路径（只读检索）", "仅做本地只读扫描，不复制原资料库，支持上万人超大资料库",
        "选择员工资料库路径（只读扫描）", "开始打包",
        "请选择员工资料库根目录和员工名单 Excel 文件，勾选所需材料类型，然后点击“开始打包”。",
        input_mode="directory_single", support_id="roster_source", support_label="员工名单文件（Excel）",
        support_button="选择名单", support_optional=True,
        fields=(
            {"id": "target_input", "kind": "text", "label": "目标人员", "placeholder": "姓名或身份证，多人用逗号隔开"},
            {"id": "library_mode", "kind": "choice", "label": "资料库形式", "default": "person_folder", "options": [
                {"label": "按人员文件夹查找（原模式）", "value": "person_folder"},
                {"label": "无序平铺资料库（OCR 索引）", "value": "flat_ocr"}
            ]},
            {"id": "collect_all", "kind": "check", "label": "全部（直接拷贝匹配到的人员整个文件夹）", "default": True},
            {"id": "create_zip", "kind": "check", "label": "生成 ZIP 压缩包", "default": False},
            {"id": "use_ocr_cache", "kind": "check", "label": "启用 OCR 缓存", "default": True},
            {"id": "material_types", "kind": "materials", "label": "指定材料", "default": []},
        ),
    ),
    ToolUiSpec(
        "folder_rename", "default", "folder_rename", "人员与档案",
        "人员资料文件夹改名", "选择人员资料目录，先预览，再确认改名。",
        "人员文件夹目录", "只处理所选目录第一层，原目录不会被修改", "选择人员文件夹目录", "预览",
        "请选择人员文件夹目录，填写改名内容，然后点击“预览”。",
        input_mode="directory_single", support_id="excel_path", support_label="人员名单 Excel",
        support_button="选择名单", support_optional=True,
        fields=(
            {"id": "rename_mode", "kind": "choice", "label": "操作", "default": "append", "options": [
                {"label": "追加文字", "value": "append"},
                {"label": "删除结尾文字", "value": "remove"},
                {"label": "修改单人名称", "value": "replace"},
                {"label": "按 Excel 人名顺序批量重命名", "value": "excel"}
            ]},
            {"id": "target_name", "kind": "text", "label": "姓名/原名称", "default": ""},
            {"id": "rename_text", "kind": "text", "label": "追加/删除文字", "default": ""},
            {"id": "replacement_name", "kind": "text", "label": "新名称", "default": ""},
            {"id": "file_type", "kind": "choice", "label": "文件类型", "default": "folder", "options": [
                {"label": "文件夹", "value": "folder"}, {"label": "PDF", "value": "pdf"},
                {"label": "图片（jpg/png/gif等）", "value": "image"},
                {"label": "文档（doc/xls/ppt/txt等）", "value": "document"},
                {"label": "全部", "value": "all"}
            ]},
        ),
    ),
)


_SPEC_MAP = {(spec.nav_id, spec.variant): spec for spec in SPECS}
DEFAULT_VARIANTS = {
    "personnel_change_merge": "merge",
    "archive_import": "import",
}

PROJECT_TOOL_NAMES = {
    "social_security": "社保明细与汇总",
    "insurance_ledger": "保险台账与预警",
    "data_statistics": "考勤与周月报",
    "salary_split": "工资表拆分",
    "salary_merge": "多月工资合并",
    "personnel_change_merge": "异动汇总",
    "roster_update": "花名册更新",
    "archive_import": "档案入库",
    "archive_export": "档案表生成",
    "material_collector": "员工资料打包",
    "folder_rename": "资料文件夹改名",
}


def spec_for(nav_id: str, variant: str | None = None) -> ToolUiSpec:
    selected = variant or DEFAULT_VARIANTS.get(nav_id, "default")
    try:
        return _SPEC_MAP[(nav_id, selected)]
    except KeyError as exc:
        raise KeyError(f"未知界面工具：{nav_id}:{selected}") from exc


def variants_for(nav_id: str) -> tuple[ToolUiSpec, ...]:
    return tuple(spec for spec in SPECS if spec.nav_id == nav_id)


def default_values(spec: ToolUiSpec) -> dict[str, Any]:
    values: dict[str, Any] = {}
    for field in spec.fields:
        if field.get("kind") == "date_range":
            values[str(field["startId"])] = field.get("startDefault", "")
            values[str(field["endId"])] = field.get("endDefault", "")
            continue
        values[str(field["id"])] = field.get("default", "")
    return values


def _validated_inputs(spec: ToolUiSpec, input_paths: list[Path]) -> list[Path]:
    if not input_paths:
        raise FormValidationError("缺少输入", f"请先{spec.input_drop_title}。")
    if spec.input_mode in {"excel_single", "directory_single"} and len(input_paths) != 1:
        raise FormValidationError("输入数量不正确", "当前功能只支持选择一个输入位置。")
    for path in input_paths:
        if not path.exists():
            raise FormValidationError("输入不存在", f"选择的文件或文件夹不存在：{path}")
        if spec.input_mode == "directory_single" and not path.is_dir():
            raise FormValidationError("需要文件夹", "当前功能需要选择一个文件夹。")
        if spec.input_mode == "excel_single" and (
            not path.is_file() or path.suffix.lower() not in EXCEL_SUFFIXES
        ):
            raise FormValidationError("格式不支持", "当前功能只支持 .xlsx 或 .xls 文件。")
        if spec.input_mode == "excel_archive_multi" and path.is_file() and not (
            path.suffix.lower() in EXCEL_SUFFIXES or is_supported_archive_file(path)
        ):
            raise FormValidationError(
                "格式不支持", "输入文件只支持 .xlsx、.xls 或 ZIP/RAR/7Z/TAR 压缩包。"
            )
    return input_paths


def _validated_support(spec: ToolUiSpec, support_text: str) -> Path | None:
    text = str(support_text or "").strip()
    if not text:
        if spec.support_id and not spec.support_optional:
            raise FormValidationError("缺少配套文件", f"请先{spec.support_button}。")
        return None
    path = Path(text).expanduser()
    if not path.exists():
        raise FormValidationError("配套资料不存在", f"选择的配套资料不存在：{path}")
    if spec.support_mode == "excel_file" and (
        not path.is_file() or path.suffix.lower() not in EXCEL_SUFFIXES
    ):
        raise FormValidationError("格式不支持", "配套资料只支持 .xlsx 或 .xls 文件。")
    if spec.support_mode == "excel_or_folder" and path.is_file() and path.suffix.lower() not in EXCEL_SUFFIXES:
        raise FormValidationError("格式不支持", "配套资料只支持 .xlsx、.xls 文件或文件夹。")
    if spec.support_mode == "excel_archive_or_folder" and path.is_file() and not (
        path.suffix.lower() in EXCEL_SUFFIXES or is_supported_archive_file(path)
    ):
        raise FormValidationError("格式不支持", "配套资料只支持 Excel、压缩包或文件夹。")
    return path


def build_invocation(
    spec: ToolUiSpec,
    *,
    input_paths: list[Path],
    support_text: str,
    values: dict[str, Any],
    output_dir: Path,
    preview: bool = False,
    preview_result: dict[str, Any] | None = None,
) -> ToolInvocation:
    """Validate UI state and map it to the exact existing business call."""

    inputs = _validated_inputs(spec, input_paths)
    # Hidden optional controls must not affect the call.  This mirrors the
    # legacy workflow: a direct material target takes precedence over a saved
    # roster path, and non-Excel rename modes ignore the Excel roster field.
    ignore_support = (
        spec.tool_id == "material_collector"
        and bool(str(values.get("target_input") or "").strip())
    ) or (
        spec.tool_id == "folder_rename"
        and str(values.get("rename_mode") or "append") != "excel"
    )
    support = None if ignore_support else _validated_support(spec, support_text)
    project_tool_name = PROJECT_TOOL_NAMES[spec.tool_id]
    description = project_tool_name
    if spec.tool_id == "folder_rename":
        rename_labels = {
            "append": "追加文字",
            "remove": "删除结尾文字",
            "replace": "修改单人名称",
            "excel": "按 Excel 人名顺序批量重命名",
        }
        description = f"{project_tool_name}-{rename_labels.get(str(values.get('rename_mode') or 'append'), '追加文字')}"
    base = dict(
        nav_id=spec.nav_id,
        variant=spec.variant,
        tool_id=spec.tool_id,
        tool_name=project_tool_name,
        group_name=spec.group,
        description=description,
    )
    if spec.tool_id == "social_security":
        return ToolInvocation(**base, function_module="hr_toolkit.tools.social_security", function_name="generate_social_security_reports", args=(inputs, support, output_dir), kwargs={})
    if spec.tool_id == "insurance_ledger":
        return ToolInvocation(**base, function_module="hr_toolkit.tools.insurance_ledger", function_name="generate_insurance_ledger", args=(inputs, support, output_dir), kwargs={})
    if spec.tool_id == "data_statistics":
        from hr_toolkit.tools.data_statistics import resolve_month_range, resolve_week_range

        try:
            week_range = resolve_week_range(values.get("week_start") or None, values.get("week_end") or None)
        except ValueError as exc:
            raise FormValidationError("日期填写有误", str(exc)) from exc
        try:
            month_range = resolve_month_range(values.get("month_start") or None, values.get("month_end") or None)
        except ValueError as exc:
            raise FormValidationError("日期填写有误", str(exc)) from exc
        kwargs = {
            "report_staff_path": support,
            "week_start": None if week_range is None else week_range[0],
            "week_end": None if week_range is None else week_range[1],
            "month_start": None if month_range is None else month_range[0],
            "month_end": None if month_range is None else month_range[1],
            "remark_unit": values.get("remark_unit") or "day",
            "include_business_trip": bool(values.get("include_business_trip")),
            "include_workday_business_trip": bool(values.get("include_workday_business_trip")),
        }
        return ToolInvocation(**base, function_module="hr_toolkit.tools.data_statistics", function_name="generate_data_statistics_reports", args=(inputs, output_dir), kwargs=kwargs)
    if spec.tool_id == "salary_split":
        return ToolInvocation(**base, function_module="hr_toolkit.tools.salary_split", function_name="split_salary_by_company", args=(inputs[0], output_dir), kwargs={})
    if spec.tool_id == "salary_merge":
        return ToolInvocation(**base, function_module="hr_toolkit.tools.salary_merge", function_name="merge_monthly_salary", args=(inputs, output_dir), kwargs={"existing_summary_path": support})
    if spec.tool_id == "personnel_change_merge":
        return ToolInvocation(**base, function_module="hr_toolkit.tools.personnel_change_merge", function_name="merge_personnel_changes", args=(inputs, output_dir), kwargs={"template_path": support})
    if spec.tool_id == "roster_update":
        return ToolInvocation(**base, function_module="hr_toolkit.tools.personnel_change_merge", function_name="update_roster_from_change_summaries", args=(inputs, support, output_dir), kwargs={})
    if spec.tool_id == "archive_import":
        return ToolInvocation(**base, function_module="hr_toolkit.tools.archive_import", function_name="import_archive_transfers", args=(inputs, support, output_dir), kwargs={})
    if spec.tool_id == "archive_export":
        return ToolInvocation(**base, function_module="hr_toolkit.tools.archive_import", function_name="export_company_archive_tables", args=(inputs, output_dir), kwargs={"existing_archive_path": support})
    if spec.tool_id == "material_collector":
        target_text = str(values.get("target_input") or "").strip()
        roster_source: str | Path | None = target_text or support
        if roster_source is None:
            raise FormValidationError("缺少员工信息", "请输入员工姓名/身份证，或选择员工名单 Excel 表格。")
        collect_all = bool(values.get("collect_all", True))
        materials = list(values.get("material_types") or [])
        if not collect_all and not materials:
            raise FormValidationError("未选择材料", "请至少选择一种材料，或者勾选“全部”。")
        library_mode = str(values.get("library_mode") or "person_folder")
        use_ocr_cache = True if library_mode == "flat_ocr" else bool(values.get("use_ocr_cache", True))
        if library_mode == "person_folder" and collect_all and target_text:
            use_ocr_cache = False
        return ToolInvocation(
            **base,
            function_module="hr_toolkit.tools.material_collector",
            function_name="collect_employee_materials",
            args=(inputs[0], output_dir),
            kwargs={
                "roster_source": roster_source,
                "material_types": None if collect_all else materials,
                "mode": "by_employee",
                "library_mode": library_mode,
                "create_zip": bool(values.get("create_zip")),
                "generate_report": True,
                "collect_all": collect_all,
                "use_ocr_cache": use_ocr_cache,
            },
        )
    if spec.tool_id == "folder_rename":
        mode = str(values.get("rename_mode") or "append")
        file_type = str(values.get("file_type") or "folder")
        if mode == "excel":
            if support is None:
                raise FormValidationError("缺少人员名单", "请先选择包含姓名列的 Excel 名单。")
            kwargs: dict[str, Any] = {
                "root_dir": inputs[0],
                "excel_path": support,
                "file_type": file_type,
                "dry_run": preview,
            }
            if not preview and preview_result is not None:
                kwargs["expected_operations"] = [
                    (Path(item["source"]).name, Path(item["target"]).name)
                    for item in preview_result.get("operations", [])
                ]
                kwargs["expected_warnings"] = list(preview_result.get("warnings", []))
            return ToolInvocation(**base, function_module="hr_toolkit.tools.folder_rename", function_name="rename_files_by_excel", args=(), kwargs=kwargs, preview=preview)
        return ToolInvocation(
            **base,
            function_module="hr_toolkit.tools.folder_rename",
            function_name="rename_person_folders",
            args=(),
            kwargs={
                "root_dir": inputs[0], "mode": mode,
                "text": str(values.get("rename_text") or ""),
                "target_name": str(values.get("target_name") or ""),
                "replacement_name": str(values.get("replacement_name") or ""),
                "file_type": file_type, "dry_run": preview,
            },
            preview=preview,
        )
    raise FormValidationError("功能不可用", f"当前功能尚未接入：{spec.tool_id}")
