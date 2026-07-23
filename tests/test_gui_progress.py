from __future__ import annotations

import unittest

from hr_toolkit.gui import _indeterminate_progress_segment


class UpdateProgressAnimationTests(unittest.TestCase):
    def test_segment_stays_inside_track(self) -> None:
        segment = _indeterminate_progress_segment(300, 150, 75)

        self.assertEqual(segment, (75, 150))

    def test_segment_enters_and_leaves_track_without_overflow(self) -> None:
        entering = _indeterminate_progress_segment(300, 20, 75)
        leaving = _indeterminate_progress_segment(300, 330, 75)

        self.assertEqual(entering, (0.0, 20))
        self.assertEqual(leaving, (255, 300))

    def test_segment_is_absent_while_fully_off_track(self) -> None:
        self.assertIsNone(_indeterminate_progress_segment(300, -1, 75))
        self.assertIsNone(_indeterminate_progress_segment(300, 376, 75))
        self.assertIsNone(_indeterminate_progress_segment(0, 10, 75))
        self.assertIsNone(_indeterminate_progress_segment(300, 10, 0))


if __name__ == "__main__":
    unittest.main()
