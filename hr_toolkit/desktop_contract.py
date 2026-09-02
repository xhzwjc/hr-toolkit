"""Renderer-neutral desktop navigation metadata."""

from __future__ import annotations


TOOL_NAV_ITEMS = (
    ("social_security", "社保明细与汇总"),
    ("insurance_ledger", "保险台账与预警"),
    ("data_statistics", "考勤与周月报"),
    ("salary_split", "工资表拆分"),
    ("salary_merge", "多月工资合并"),
    ("personnel_change_merge", "异动汇总"),
    ("archive_import", "档案入库"),
    ("material_collector", "员工资料打包"),
    ("folder_rename", "资料文件夹改名"),
)

NAV_GROUPS = (
    ("社保与保险", ("social_security", "insurance_ledger")),
    ("考勤与统计", ("data_statistics",)),
    ("薪酬管理", ("salary_split", "salary_merge")),
    (
        "人员与档案",
        (
            "personnel_change_merge",
            "archive_import",
            "material_collector",
            "folder_rename",
        ),
    ),
)


__all__ = ["NAV_GROUPS", "TOOL_NAV_ITEMS"]
