from __future__ import annotations

import unittest

from hr_toolkit.common.excel import _translate_same_row_formula


class ExcelHelperTest(unittest.TestCase):
    def test_translate_same_row_formula_only_rewrites_cell_reference_rows(self) -> None:
        formula = "=A1+B10+SUM(C1:D1)+E$1+$F1+LOG10(100)"

        self.assertEqual(
            _translate_same_row_formula(formula, source_row=1, target_row=12),
            "=A12+B10+SUM(C12:D12)+E$1+$F12+LOG10(100)",
        )


if __name__ == "__main__":
    unittest.main()
