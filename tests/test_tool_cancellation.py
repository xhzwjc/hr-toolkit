from __future__ import annotations

import unittest

from hr_toolkit.tools.data_statistics import generate_data_statistics_reports
from hr_toolkit.tools.folder_rename import MODE_APPEND, rename_files_by_excel, rename_person_folders
from hr_toolkit.tools.insurance_ledger import generate_insurance_ledger
from hr_toolkit.tools.personnel_change_merge import (
    merge_personnel_changes,
    update_roster_from_change_summaries,
)
from hr_toolkit.tools.salary_merge import merge_monthly_salary
from hr_toolkit.tools.salary_split import split_salary_by_company
from hr_toolkit.tools.social_security import generate_social_security_reports


class ToolCancellationTests(unittest.TestCase):
    def test_all_major_tools_honor_pre_start_cancellation(self) -> None:
        stopped = lambda: True
        calls = (
            lambda: merge_monthly_salary("missing", "out", cancelled=stopped),
            lambda: split_salary_by_company("missing", "out", cancelled=stopped),
            lambda: generate_data_statistics_reports("missing", "out", cancelled=stopped),
            lambda: generate_social_security_reports("missing", "roster", "out", cancelled=stopped),
            lambda: generate_insurance_ledger("missing", "roster", "out", cancelled=stopped),
            lambda: merge_personnel_changes("missing", "out", cancelled=stopped),
            lambda: update_roster_from_change_summaries("missing", "roster", "out", cancelled=stopped),
            lambda: rename_person_folders("missing", mode=MODE_APPEND, cancelled=stopped),
            lambda: rename_files_by_excel("missing", "roster", cancelled=stopped),
        )
        for call in calls:
            with self.subTest(call=call), self.assertRaisesRegex(RuntimeError, "已停止"):
                call()


if __name__ == "__main__":
    unittest.main()
