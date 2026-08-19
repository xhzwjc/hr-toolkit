from __future__ import annotations

import unittest

from hr_toolkit.tools.registry import (
    ToolSpec,
    clear_registry,
    ensure_default_tools_registered,
    get_all_tools,
    get_tool_by_cli_command,
    get_tool_by_id,
    get_tools_by_group,
    register_tool,
)


class ToolRegistryTests(unittest.TestCase):
    def setUp(self) -> None:
        clear_registry()

    def tearDown(self) -> None:
        clear_registry()

    def test_default_tools_registration(self) -> None:
        ensure_default_tools_registered()
        tools = get_all_tools()
        self.assertGreaterEqual(len(tools), 9)

        tool_ids = [t.tool_id for t in tools]
        self.assertIn("social_security", tool_ids)
        self.assertIn("salary_split", tool_ids)
        self.assertIn("material_collector", tool_ids)

    def test_get_tool_by_id(self) -> None:
        spec = get_tool_by_id("salary_split")
        self.assertIsNotNone(spec)
        assert spec is not None
        self.assertEqual(spec.name, "工资表拆分")
        self.assertEqual(spec.cli_command, "salary-split")
        self.assertEqual(spec.group, "薪酬管理")
        self.assertFalse(spec.multi_input)

    def test_get_tool_by_cli_command(self) -> None:
        spec = get_tool_by_cli_command("change-merge")
        self.assertIsNotNone(spec)
        assert spec is not None
        self.assertEqual(spec.tool_id, "personnel_change_merge")

    def test_get_tools_by_group(self) -> None:
        grouped = get_tools_by_group()
        self.assertIn("社保与保险", grouped)
        self.assertIn("薪酬管理", grouped)
        self.assertIn("人员与档案", grouped)

        salary_tool_ids = [t.tool_id for t in grouped["薪酬管理"]]
        self.assertIn("salary_split", salary_tool_ids)
        self.assertIn("salary_merge", salary_tool_ids)

    def test_register_custom_tool(self) -> None:
        custom_spec = ToolSpec(
            tool_id="custom_tool",
            name="自定义工具",
            group="测试分组",
            cli_command="custom-tool",
            help_text="测试帮助说明",
            entry_point=lambda: "custom_done",
            aliases=("custom-alias",),
        )
        register_tool(custom_spec)

        self.assertEqual(get_tool_by_id("custom_tool"), custom_spec)
        self.assertEqual(get_tool_by_cli_command("custom-tool"), custom_spec)
        self.assertEqual(get_tool_by_cli_command("custom-alias"), custom_spec)


if __name__ == "__main__":
    unittest.main()
