from __future__ import annotations

import queue
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from hr_toolkit import background_process
from hr_toolkit.background_process import (
    BusinessProcessError,
    BusinessProcessStartError,
    PROCESS_FILE_THRESHOLD_BYTES,
    run_business_process,
    should_use_process,
)


class BackgroundProcessTests(unittest.TestCase):
    def test_load_policy_is_bounded_and_does_not_scan_directories(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            tiny = root / "tiny.xlsx"
            tiny.write_bytes(b"x")
            large = root / "large.xlsx"
            with large.open("wb") as handle:
                handle.truncate(PROCESS_FILE_THRESHOLD_BYTES)
            self.assertFalse(should_use_process("salary_split", (tiny, root / "out"), {}))
            self.assertTrue(should_use_process("salary_split", (large, root / "out"), {}))
            self.assertTrue(should_use_process("salary_merge", ([tiny] * 5, root / "out"), {}))
            self.assertTrue(should_use_process("salary_merge", (root, root / "out"), {}))
            self.assertTrue(should_use_process("material_collector", (tiny, root / "out"), {}))

    def test_spawned_call_returns_serializable_payload_and_progress(self) -> None:
        progress = []
        result = run_business_process(
            module_name="hr_toolkit.background_process",
            function_name="_process_smoke_probe",
            args=(Path("input.xlsx"), Path("output")),
            kwargs={},
            cancel_event=threading.Event(),
            on_progress=lambda current, total, message: progress.append((current, total, message)),
        )
        self.assertEqual(result.payload["input_path"], "input.xlsx")
        self.assertEqual(result.payload["output_dir"], "output")
        self.assertGreaterEqual(result.elapsed_seconds, 0.0)
        self.assertIn((1, 1, "完成"), progress)

    def test_child_flushes_progress_before_sending_terminal_result(self) -> None:
        progress_queue = Mock()
        connection = Mock()
        cancel_event = Mock()
        cancel_event.is_set.return_value = False
        terminal_messages = []

        def capture_terminal(message) -> None:
            if message[0] == "ready":
                terminal_messages.append(message)
                return
            progress_queue.close.assert_called_once_with()
            progress_queue.join_thread.assert_called_once_with()
            terminal_messages.append(message)

        connection.send.side_effect = capture_terminal
        background_process._child_entry(
            connection,
            progress_queue,
            cancel_event,
            "hr_toolkit.background_process",
            "_process_smoke_probe",
            (Path("input.xlsx"), Path("output")),
            {},
        )

        progress_queue.put_nowait.assert_called_once_with((1, 1, "完成"))
        self.assertEqual([message[0] for message in terminal_messages], ["ready", "success"])
        connection.close.assert_called_once_with()

    def test_child_import_failure_is_safe_startup_failure(self) -> None:
        with self.assertRaises(BusinessProcessStartError) as caught:
            run_business_process(
                module_name="hr_toolkit.module_that_does_not_exist",
                function_name="missing",
                args=(),
                kwargs={},
                cancel_event=threading.Event(),
            )

        self.assertIn("ModuleNotFoundError", str(caught.exception))

    def test_business_failure_after_ready_is_not_classified_as_startup(self) -> None:
        with self.assertRaises(BusinessProcessError) as caught:
            run_business_process(
                module_name="hr_toolkit.background_process",
                function_name="_process_smoke_probe",
                args=(),
                kwargs={},
                cancel_event=threading.Event(),
            )

        self.assertIn("TypeError", str(caught.exception))

    def test_spawn_start_failure_is_classified_and_closes_resources(self) -> None:
        context = Mock()
        process_cancel_event = Mock()
        receive_connection = Mock()
        send_connection = Mock()
        progress_queue = Mock()
        process = Mock()
        process.start.side_effect = FileNotFoundError(
            2,
            "系统找不到指定的文件",
            "HRToolkit.exe",
        )
        process.is_alive.return_value = False
        context.Event.return_value = process_cancel_event
        context.Pipe.return_value = (receive_connection, send_connection)
        context.Queue.return_value = progress_queue
        context.Process.return_value = process

        with patch.object(
            background_process.multiprocessing,
            "get_context",
            return_value=context,
        ):
            with self.assertRaises(BusinessProcessStartError) as caught:
                run_business_process(
                    module_name="hr_toolkit.background_process",
                    function_name="_process_smoke_probe",
                    args=(Path("input.xlsx"), Path("output")),
                    kwargs={},
                    cancel_event=threading.Event(),
                )

        self.assertIsInstance(caught.exception.__cause__, FileNotFoundError)
        self.assertIn("无法启动独立后台进程", str(caught.exception))
        receive_connection.close.assert_called_once_with()
        send_connection.close.assert_called_once_with()
        progress_queue.close.assert_called_once_with()
        progress_queue.cancel_join_thread.assert_called_once_with()
        process.terminate.assert_not_called()

    def test_pipe_eof_before_ready_is_safe_startup_failure(self) -> None:
        context = Mock()
        receive_connection = Mock()
        send_connection = Mock()
        progress_queue = Mock()
        process = Mock()
        context.Event.return_value = Mock()
        context.Pipe.return_value = (receive_connection, send_connection)
        context.Queue.return_value = progress_queue
        context.Process.return_value = process
        receive_connection.poll.return_value = True
        receive_connection.recv.side_effect = EOFError()
        progress_queue.get_nowait.side_effect = queue.Empty()
        process.is_alive.return_value = True
        process.exitcode = None

        with patch.object(
            background_process.multiprocessing,
            "get_context",
            return_value=context,
        ):
            with self.assertRaises(BusinessProcessStartError) as caught:
                run_business_process(
                    module_name="hr_toolkit.background_process",
                    function_name="_process_smoke_probe",
                    args=(Path("input.xlsx"), Path("output")),
                    kwargs={},
                    cancel_event=threading.Event(),
                )

        self.assertIn("尚未就绪", str(caught.exception))

    def test_pipe_eof_after_ready_is_not_safe_to_retry(self) -> None:
        context = Mock()
        receive_connection = Mock()
        send_connection = Mock()
        progress_queue = Mock()
        process = Mock()
        context.Event.return_value = Mock()
        context.Pipe.return_value = (receive_connection, send_connection)
        context.Queue.return_value = progress_queue
        context.Process.return_value = process
        receive_connection.poll.return_value = True
        receive_connection.recv.side_effect = [("ready",), EOFError()]
        progress_queue.get_nowait.side_effect = queue.Empty()
        process.is_alive.return_value = True
        process.exitcode = None

        with patch.object(
            background_process.multiprocessing,
            "get_context",
            return_value=context,
        ):
            with self.assertRaises(BusinessProcessError) as caught:
                run_business_process(
                    module_name="hr_toolkit.background_process",
                    function_name="_process_smoke_probe",
                    args=(Path("input.xlsx"), Path("output")),
                    kwargs={},
                    cancel_event=threading.Event(),
                )

        self.assertIn("已开始", str(caught.exception))

if __name__ == "__main__":
    unittest.main()
