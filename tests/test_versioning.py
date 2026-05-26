from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "versioning.py"
SPEC = importlib.util.spec_from_file_location("versioning", SCRIPT_PATH)
versioning = importlib.util.module_from_spec(SPEC)
assert SPEC is not None and SPEC.loader is not None
SPEC.loader.exec_module(versioning)


class VersioningTest(unittest.TestCase):
    def test_bump_version(self) -> None:
        self.assertEqual(versioning.bump_version("0.1.0", "patch"), "0.1.1")
        self.assertEqual(versioning.bump_version("0.1.9", "minor"), "0.2.0")
        self.assertEqual(versioning.bump_version("0.9.9", "major"), "1.0.0")

    def test_bump_project_version_updates_init_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            init_file = Path(tmp) / "__init__.py"
            init_file.write_text('__version__ = "0.1.0"\n', encoding="utf-8")

            new_version = versioning.bump_project_version("patch", init_file)

            self.assertEqual(new_version, "0.1.1")
            self.assertEqual(versioning.read_project_version(init_file), "0.1.1")


if __name__ == "__main__":
    unittest.main()
