"""Unified asynchronous task runner for GUI."""

from __future__ import annotations

import queue
import threading
from typing import Any, Callable
from tkinter import Tk


class TaskToken:
    """Handle for a background task execution."""

    def __init__(self, task_id: str, cancel_event: threading.Event) -> None:
        self.task_id = task_id
        self._cancel_event = cancel_event

    def cancel(self) -> None:
        self._cancel_event.set()

    @property
    def is_cancelled(self) -> bool:
        return self._cancel_event.is_set()


class TaskRunner:
    """Manages background threads and routes completions/progress/errors to the Tk main thread."""

    def __init__(self, root: Tk, poll_interval_ms: int = 100) -> None:
        self._root = root
        self._poll_interval_ms = poll_interval_ms
        self._queue: queue.Queue[tuple[str, str, Any, Callable[..., Any] | None]] = queue.Queue()
        self._active_tasks: dict[str, tuple[threading.Thread, threading.Event]] = {}
        self._lock = threading.Lock()
        self._polling_active = False
        self.start_polling()

    def start_polling(self) -> None:
        if not self._polling_active:
            self._polling_active = True
            self._poll()

    def submit(
        self,
        task_id: str,
        worker_fn: Callable[[Callable[[], bool], Callable[[Any], None]], Any],
        on_success: Callable[[Any], None] | None = None,
        on_error: Callable[[Exception], None] | None = None,
        on_progress: Callable[[Any], None] | None = None,
        on_cancel: Callable[[], None] | None = None,
    ) -> TaskToken:
        with self._lock:
            if task_id in self._active_tasks:
                _prev_thread, prev_cancel = self._active_tasks[task_id]
                prev_cancel.set()

            cancel_event = threading.Event()

            def report_progress(data: Any) -> None:
                if on_progress is not None:
                    self._queue.put(("progress", task_id, data, on_progress))

            def worker() -> None:
                try:
                    result = worker_fn(cancel_event.is_set, report_progress)
                    if cancel_event.is_set():
                        if on_cancel is not None:
                            self._queue.put(("cancel", task_id, None, on_cancel))
                    else:
                        if on_success is not None:
                            self._queue.put(("success", task_id, result, on_success))
                except Exception as exc:
                    if cancel_event.is_set() and on_cancel is not None:
                        self._queue.put(("cancel", task_id, None, on_cancel))
                    elif on_error is not None:
                        self._queue.put(("error", task_id, exc, on_error))
                finally:
                    with self._lock:
                        if self._active_tasks.get(task_id) == (thread, cancel_event):
                            self._active_tasks.pop(task_id, None)

            thread = threading.Thread(target=worker, daemon=True)
            self._active_tasks[task_id] = (thread, cancel_event)
            thread.start()
            return TaskToken(task_id, cancel_event)

    def cancel(self, task_id: str) -> bool:
        with self._lock:
            if task_id in self._active_tasks:
                _thread, cancel_event = self._active_tasks[task_id]
                cancel_event.set()
                return True
            return False

    def is_running(self, task_id: str) -> bool:
        with self._lock:
            return task_id in self._active_tasks

    def stop(self) -> None:
        """Stop polling and signal cancellation to all active background tasks."""
        self._polling_active = False
        with self._lock:
            for _thread, cancel_event in list(self._active_tasks.values()):
                cancel_event.set()
            self._active_tasks.clear()

    def _poll(self) -> None:
        try:
            while True:
                msg_type, _task_id, payload, callback = self._queue.get_nowait()
                if callback is not None:
                    try:
                        if msg_type == "cancel":
                            callback()
                        else:
                            callback(payload)
                    except Exception:
                        pass
        except queue.Empty:
            pass
        if self._polling_active:
            try:
                self._root.after(self._poll_interval_ms, self._poll)
            except Exception:
                self._polling_active = False
