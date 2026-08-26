from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from hr_toolkit.common.paths import path_is_relative_to


class PathCompatibilityTests(unittest.TestCase):
    def test_path_is_relative_to_matches_expected_containment(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            self.assertTrue(path_is_relative_to(root, root))
            self.assertTrue(path_is_relative_to(root / "child" / "file.xlsx", root))
            self.assertFalse(path_is_relative_to(root.parent / "outside.xlsx", root))


if __name__ == "__main__":
    unittest.main()
