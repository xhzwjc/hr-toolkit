"""Single-worker process isolation for CPU-heavy desktop tool calls.

The project intentionally uses at most one business worker process.  This keeps
CPU, memory and disk concurrency bounded on low-spec machines while preventing
Excel/OCR Python work from starving the GUI interpreter thread.
"""

from __future__ import annotations

import importlib
import inspect
import multiprocessing
import queue
import time
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable

from .project_run import serializable


PROCESS_FILE_THRESHOLD_BYTES = 8 * 1024 * 1024
PROCESS_INPUT_COUNT_THRESHOLD = 5
PROCESS_CANCEL_GRACE_SECONDS = 2.0
PROCESS_POLL_SECONDS = 0.025
ALWAYS_ISOLATED_TOOL_IDS = frozenset(
    {
        "material_collector",
        "archive_import",
        "archive_export",
    }
)


@dataclass(frozen=True)
class ProcessCallResult:
    payload: dict[str, Any]
    elapsed_seconds: float


class BusinessProcessError(RuntimeError):
    def __init__(self, message: str, *, remote_traceback: str = "") -> None:
        super().__init__(message)
        self.remote_traceback = remote_traceback


class BusinessProcessCancelled(RuntimeError):
    pass


def _iter_paths(value: Any) -> Iterable[Path]:
    if isinstance(value, Path):
        yield value
        return
    if isinstance(value, (list, tuple)):
        for item in value:
            yield from _iter_paths(item)


def should_use_process(
    tool_id: str,
    args: Iterable[Any],
    kwargs: dict[str, Any],
) -> bool:
    """Choose isolation without recursively scanning user directories."""

    if tool_id in ALWAYS_ISOLATED_TOOL_IDS:
        return True
    paths: list[Path] = []
    for value in (*tuple(args), *tuple(kwargs.values())):
        paths.extend(_iter_paths(value))
    # Output paths may not exist yet; only existing inputs influence the choice.
    existing = [path for path in paths if path.exists()]
    if len(existing) >= PROCESS_INPUT_COUNT_THRESHOLD:
        return True
    for path in existing:
        try:
            if path.is_dir():
                return True
            if path.is_file() and path.stat().st_size >= PROCESS_FILE_THRESHOLD_BYTES:
                return True
        except OSError:
            # An unreadable input will be diagnosed by the original business
            # function; isolating that diagnosis is safer than blocking the UI.
            return True
    return False


def _result_payload(result: Any) -> dict[str, Any]:
    if hasattr(result, "to_dict") and callable(result.to_dict):
        payload = result.to_dict()
    elif isinstance(result, dict):
        payload = result
    else:
        payload = {"result": serializable(result)}
    return serializable(payload)


def _child_entry(
    connection,
    progress_queue,
    process_cancel_event,
    module_name: str,
    function_name: str,
    args: tuple[Any, ...],
    kwargs: dict[str, Any],
) -> None:
    started = time.monotonic()
    terminal_message: tuple[Any, ...]

    def progress(current: int, total: int, message: str) -> None:
        record = (int(current), int(total), str(message))
        try:
            progress_queue.put_nowait(record)
        except queue.Full:
            # Progress is replaceable telemetry; business results are not.
            pass

    try:
        module = importlib.import_module(module_name)
        function = getattr(module, function_name)
        call_kwargs = dict(kwargs)
        parameters = inspect.signature(function).parameters
        if "cancelled" in parameters:
            call_kwargs["cancelled"] = process_cancel_event.is_set
        if "progress_callback" in parameters:
            call_kwargs["progress_callback"] = progress
        result = function(*args, **call_kwargs)
        terminal_message = (
            "success",
            _result_payload(result),
            time.monotonic() - started,
        )
    except BaseException as exc:
        terminal_message = (
            "error",
            f"{type(exc).__name__}: {exc}",
            traceback.format_exc(),
        )
    finally:
        try:
            # multiprocessing.Queue.put_nowait() hands records to a feeder
            # thread.  On Windows the result pipe can otherwise overtake that
            # feeder, making the parent return before the final progress event
            # is visible.  Flush the child writer before publishing the
            # terminal result so both streams have a deterministic boundary.
            progress_queue.close()
            progress_queue.join_thread()
        except Exception:
            pass
        try:
            connection.send(terminal_message)
        finally:
            try:
                connection.close()
            except Exception:
                pass


def run_business_process(
    *,
    module_name: str,
    function_name: str,
    args: tuple[Any, ...],
    kwargs: dict[str, Any],
    cancel_event: Any,
    on_progress: Callable[[int, int, str], None] | None = None,
) -> ProcessCallResult:
    """Run exactly one business call in a spawned child and wait off the UI thread."""

    context = multiprocessing.get_context("spawn")
    process_cancel_event = context.Event()
    receive_connection, send_connection = context.Pipe(duplex=False)
    progress_queue = context.Queue(maxsize=64)
    process = context.Process(
        target=_child_entry,
        args=(
            send_connection,
            progress_queue,
            process_cancel_event,
            module_name,
            function_name,
            args,
            kwargs,
        ),
        daemon=False,
        name=f"HRToolkit-{function_name}",
    )
    process.start()
    send_connection.close()
    cancellation_started: float | None = None
    message: tuple[Any, ...] | None = None

    def drain_progress() -> None:
        while True:
            try:
                current, total, text = progress_queue.get_nowait()
            except queue.Empty:
                return
            if on_progress is not None:
                on_progress(current, total, text)

    try:
        while True:
            drain_progress()
            if receive_connection.poll(0):
                message = receive_connection.recv()
                # The child flushes its queue writer before sending this
                # terminal message, so this final drain cannot race its feeder.
                drain_progress()
                break
            if cancel_event.is_set():
                process_cancel_event.set()
                if cancellation_started is None:
                    cancellation_started = time.monotonic()
                elif time.monotonic() - cancellation_started >= PROCESS_CANCEL_GRACE_SECONDS:
                    if process.is_alive():
                        process.terminate()
                    raise BusinessProcessCancelled("本次处理已停止。")
            if not process.is_alive():
                if receive_connection.poll(0.1):
                    message = receive_connection.recv()
                    break
                if cancel_event.is_set():
                    raise BusinessProcessCancelled("本次处理已停止。")
                raise BusinessProcessError(
                    f"后台处理进程异常退出（退出码 {process.exitcode}）。"
                )
            cancel_event.wait(PROCESS_POLL_SECONDS)
        process.join(timeout=1.0)
        if not message:
            raise BusinessProcessError("后台处理没有返回结果。")
        if message[0] == "success":
            return ProcessCallResult(
                payload=dict(message[1]),
                elapsed_seconds=float(message[2]),
            )
        raise BusinessProcessError(str(message[1]), remote_traceback=str(message[2]))
    finally:
        if process.is_alive():
            process.terminate()
        process.join(timeout=1.0)
        receive_connection.close()
        progress_queue.close()
        progress_queue.join_thread()


def _process_smoke_probe(
    input_path: str | Path,
    output_dir: str | Path,
    *,
    progress_callback=None,
    cancelled=None,
) -> dict[str, Any]:
    """Tiny importable target used only by process/packaging smoke tests."""

    if cancelled is not None and cancelled():
        raise RuntimeError("cancelled")
    if progress_callback is not None:
        progress_callback(1, 1, "完成")
    return {
        "input_path": str(input_path),
        "output_dir": str(output_dir),
    }
