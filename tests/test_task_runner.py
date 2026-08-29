from __future__ import annotations

import threading
import time
import unittest
from unittest.mock import MagicMock

from hr_toolkit.gui.task_runner import TaskRunner


class TaskRunnerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.mock_root = MagicMock()
        self.runner = TaskRunner(self.mock_root, poll_interval_ms=10)

    def _wait_for_task(self, task_id: str, *, timeout: float = 5.0) -> None:
        deadline = time.monotonic() + timeout
        while self.runner.is_running(task_id) and time.monotonic() < deadline:
            time.sleep(0.001)
        self.assertFalse(self.runner.is_running(task_id), f"task did not stop: {task_id}")

    def test_submit_success_task(self) -> None:
        completed = []

        def worker_fn(is_cancelled, on_progress):
            on_progress("step 1")
            return "done"

        def on_success(result):
            completed.append(result)

        token = self.runner.submit(
            "test_task",
            worker_fn,
            on_success=on_success,
        )
        self.assertFalse(token.is_cancelled)
        self._wait_for_task("test_task")
        self.runner._poll()
        self.assertEqual(completed, ["done"])

    def test_submit_error_task(self) -> None:
        errors = []

        def worker_fn(is_cancelled, on_progress):
            raise ValueError("failed!")

        def on_error(exc):
            errors.append(str(exc))

        self.runner.submit(
            "err_task",
            worker_fn,
            on_error=on_error,
        )
        self._wait_for_task("err_task")
        self.runner._poll()
        self.assertEqual(errors, ["failed!"])

    def test_cancel_task(self) -> None:
        cancelled = []

        def worker_fn(is_cancelled, on_progress):
            for _ in range(50):
                if is_cancelled():
                    break
                time.sleep(0.01)
            return "should not finish"

        def on_cancel():
            cancelled.append(True)

        token = self.runner.submit("cancel_task", worker_fn, on_cancel=on_cancel)
        token.cancel()
        self._wait_for_task("cancel_task")
        self.runner._poll()
        self.assertEqual(cancelled, [True])


    def test_resubmit_same_task_id_preserves_new_task(self) -> None:
        first_started = threading.Event()
        first_cancelled = threading.Event()
        second_completed = []

        def worker_1(is_cancelled, on_progress):
            first_started.set()
            while True:
                if is_cancelled():
                    first_cancelled.set()
                    return
                time.sleep(0.001)

        def worker_2(is_cancelled, on_progress):
            return "second_done"

        # Submit task 1
        self.runner.submit("dup_task", worker_1)
        self.assertTrue(self.runner.is_running("dup_task"))
        self.assertTrue(first_started.wait(timeout=1.0))

        # Immediately submit task 2 with the same ID
        self.runner.submit("dup_task", worker_2, on_success=lambda res: second_completed.append(res))

        self.assertTrue(first_cancelled.wait(timeout=1.0))
        self._wait_for_task("dup_task")
        self.runner._poll()

        self.assertTrue(first_cancelled.is_set())
        self.assertEqual(second_completed, ["second_done"])
        self.assertFalse(self.runner.is_running("dup_task"))

    def test_stop_runner(self) -> None:
        cancelled = threading.Event()

        def worker(is_cancelled, on_progress):
            for _ in range(50):
                if is_cancelled():
                    cancelled.set()
                    return
                time.sleep(0.01)

        self.runner.submit("long_task", worker)
        self.assertTrue(self.runner.is_running("long_task"))

        self.runner.stop()
        self.assertFalse(self.runner._polling_active)
        self.assertFalse(self.runner.is_running("long_task"))
        self.assertTrue(cancelled.wait(timeout=1.0))

    def test_poll_handles_root_error_gracefully(self) -> None:
        self.mock_root.after.side_effect = RuntimeError("Tk root destroyed")
        self.runner._poll()
        self.assertFalse(self.runner._polling_active)


if __name__ == "__main__":
    unittest.main()
