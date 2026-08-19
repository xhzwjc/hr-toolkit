"""Declarative tool registry for HR Toolkit.

Provides a unified specification (ToolSpec) for tools, allowing CLI and GUI
layers to dynamically discover, inspect, and execute business tools.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable
from pathlib import Path


@dataclass(frozen=True)
class ToolSpec:
    """Specification and metadata for an HR Toolkit tool."""

    tool_id: str
    name: str
    group: str
    cli_command: str
    help_text: str
    entry_point: Callable[..., Any]
    summary_formatter: Callable[[dict[str, Any]], None] | None = None
    multi_input: bool = False
    supports_dry_run: bool = True
    aliases: tuple[str, ...] = ()
    extra_metadata: dict[str, Any] = field(default_factory=dict)


_REGISTRY: dict[str, ToolSpec] = {}
_CLI_COMMAND_MAP: dict[str, ToolSpec] = {}
_INITIALIZED: bool = False


def register_tool(spec: ToolSpec) -> None:
    """Register a tool specification."""
    _REGISTRY[spec.tool_id] = spec
    _CLI_COMMAND_MAP[spec.cli_command] = spec
    for alias in spec.aliases:
        _CLI_COMMAND_MAP[alias] = spec


def ensure_default_tools_registered() -> None:
    """Ensure all standard HR Toolkit tools are registered."""
    global _INITIALIZED
    if _INITIALIZED:
        return
    _INITIALIZED = True

    from .social_security import generate_social_security_reports
    from .insurance_ledger import generate_insurance_ledger
    from .data_statistics import generate_data_statistics_reports
    from .salary_split import split_salary_by_company
    from .salary_merge import merge_monthly_salary
    from .personnel_change_merge import (
        merge_personnel_changes,
        update_roster_from_change_summaries,
    )
    from .archive_import import (
        import_archive_transfers,
        export_company_archive_tables,
    )
    from .material_collector import collect_employee_materials
    from .folder_rename import rename_person_folders

    register_tool(
        ToolSpec(
            tool_id="social_security",
            name="社保明细与汇总",
            group="社保与保险",
            cli_command="social-security",
            help_text="需求1：生成社保明细表和社保汇总表",
            entry_point=generate_social_security_reports,
            multi_input=True,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="insurance_ledger",
            name="保险台账与预警",
            group="社保与保险",
            cli_command="insurance-ledger",
            help_text="需求3：生成保险台账和人员增减预警",
            entry_point=generate_insurance_ledger,
            multi_input=True,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="data_statistics",
            name="考勤与周月报",
            group="考勤与统计",
            cli_command="data-statistics",
            help_text="需求2：生成考勤和周月报统计表",
            entry_point=generate_data_statistics_reports,
            multi_input=True,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="salary_split",
            name="工资表拆分",
            group="薪酬管理",
            cli_command="salary-split",
            help_text="需求4：将工资表按入职公司拆分为多个工作簿",
            entry_point=split_salary_by_company,
            multi_input=False,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="salary_merge",
            name="多月工资合并",
            group="薪酬管理",
            cli_command="salary-merge",
            help_text="需求5：合并多个月工资表，生成个人应发工资汇总",
            entry_point=merge_monthly_salary,
            multi_input=True,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="personnel_change_merge",
            name="异动汇总",
            group="人员与档案",
            cli_command="change-merge",
            help_text="需求6：汇总多个项目异动表",
            entry_point=merge_personnel_changes,
            multi_input=True,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="roster_update",
            name="花名册更新",
            group="人员与档案",
            cli_command="roster-update",
            help_text="需求6：根据异动汇总表单独更新人力资源花名册",
            entry_point=update_roster_from_change_summaries,
            multi_input=True,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="archive_import",
            name="档案入库",
            group="人员与档案",
            cli_command="archive-import",
            help_text="需求7：将项目档案移交表写入公司档案汇总表",
            entry_point=import_archive_transfers,
            multi_input=True,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="archive_export",
            name="档案表生成",
            group="人员与档案",
            cli_command="archive-export",
            help_text="需求7：按公司从档案汇总表生成独立档案表",
            entry_point=export_company_archive_tables,
            multi_input=True,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="material_collector",
            name="员工资料打包",
            group="人员与档案",
            cli_command="material-collector",
            help_text="需求9：员工资料自动打包与信息提取",
            entry_point=collect_employee_materials,
            multi_input=False,
        )
    )
    register_tool(
        ToolSpec(
            tool_id="folder_rename",
            name="资料文件夹改名",
            group="人员与档案",
            cli_command="folder-rename",
            help_text="需求8：人员资料文件夹批量改名",
            entry_point=rename_person_folders,
            multi_input=False,
        )
    )


def get_all_tools() -> tuple[ToolSpec, ...]:
    """Return all registered tools in order of registration."""
    ensure_default_tools_registered()
    return tuple(_REGISTRY.values())


def get_tool_by_id(tool_id: str) -> ToolSpec | None:
    """Look up a tool by its unique tool_id."""
    ensure_default_tools_registered()
    return _REGISTRY.get(tool_id)


def get_tool_by_cli_command(command: str) -> ToolSpec | None:
    """Look up a tool by its CLI command name or alias."""
    ensure_default_tools_registered()
    return _CLI_COMMAND_MAP.get(command)


def get_tools_by_group() -> dict[str, tuple[ToolSpec, ...]]:
    """Group registered tools by their group name."""
    ensure_default_tools_registered()
    groups: dict[str, list[ToolSpec]] = {}
    for spec in _REGISTRY.values():
        groups.setdefault(spec.group, []).append(spec)
    return {group: tuple(tools) for group, tools in groups.items()}


def clear_registry() -> None:
    """Clear all registered tools (primarily used in test suites)."""
    global _INITIALIZED
    _REGISTRY.clear()
    _CLI_COMMAND_MAP.clear()
    _INITIALIZED = False
