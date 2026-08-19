"""Concrete HR workflow tools and declarative registry."""

from .registry import (
    ToolSpec,
    get_all_tools,
    get_tool_by_id,
    get_tool_by_cli_command,
    get_tools_by_group,
    register_tool,
)

__all__ = [
    "ToolSpec",
    "get_all_tools",
    "get_tool_by_id",
    "get_tool_by_cli_command",
    "get_tools_by_group",
    "register_tool",
]
