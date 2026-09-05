"""Framework-neutral project run coordinator used by desktop front ends."""

from __future__ import annotations

import inspect
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from . import runlog
from .background_process import (
    BusinessProcessCancelled,
    BusinessProcessError,
    BusinessProcessStartError,
    run_business_process,
    should_use_process,
)
from .project_run import (
    call_with_project_inputs,
    context_from_call,
    import_project_run_sources,
    project_batch_is_closed,
    rebase_project_replacements,
    serializable,
)


Callback = Callable[..., None]
PROGRESS_UI_INTERVAL_SECONDS = 0.1


@dataclass(frozen=True)
class RunRequest:
    tool_id: str
    tool_name: str
    group_name: str
    description: str
    function: Callable[..., Any]
    args: tuple[Any, ...]
    kwargs: dict[str, Any]


@dataclass(frozen=True)
class RunCallbacks:
    log: Callback = lambda *_args: None
    progress: Callback = lambda *_args: None
    success: Callback = lambda *_args: None
    error: Callback = lambda *_args: None
    stopped: Callback = lambda *_args: None
    finished: Callback = lambda *_args: None


class ProjectRunCoordinator:
    """Own at most one project run and keep its lifecycle off the UI thread."""

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._thread: threading.Thread | None = None
        self._cancel_event: threading.Event | None = None
        self._active_batch_id: str | None = None
        # One controller owns one coordinator for the whole desktop session.
        # Once this machine proves that the frozen worker cannot be launched,
        # remember it instead of paying the same failing CreateProcess cost on
        # every subsequent large batch.
        self._process_isolation_available = True
        self._process_isolation_failure = ""

    @property
    def running(self) -> bool:
        with self._lock:
            return self._thread is not None and self._thread.is_alive()

    @property
    def active_batch_id(self) -> str | None:
        with self._lock:
            return self._active_batch_id

    @property
    def process_isolation_available(self) -> bool:
        with self._lock:
            return self._process_isolation_available

    @property
    def process_isolation_failure(self) -> str:
        with self._lock:
            return self._process_isolation_failure

    def note_process_start_failure(self, error: BaseException) -> None:
        """Open the session circuit breaker after a proven launch failure."""

        with self._lock:
            self._process_isolation_available = False
            self._process_isolation_failure = str(error)

    def start(self, store: Any, request: RunRequest, callbacks: RunCallbacks) -> bool:
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return False
            cancel_event = threading.Event()
            worker = threading.Thread(
                target=self._run,
                args=(store, request, callbacks, cancel_event),
                daemon=True,
                name=f"HRToolkit-project-{request.tool_id}",
            )
            self._cancel_event = cancel_event
            self._thread = worker
            self._active_batch_id = None
            worker.start()
            return True

    def cancel(self) -> bool:
        with self._lock:
            event = self._cancel_event
            if event is None:
                return False
            event.set()
            return True

    def wait(self, timeout: float | None = None) -> bool:
        with self._lock:
            worker = self._thread
        if worker is None:
            return True
        worker.join(timeout=timeout)
        return not worker.is_alive()

    @staticmethod
    def _payload(result: Any) -> dict[str, Any]:
        if hasattr(result, "to_dict") and callable(result.to_dict):
            value = result.to_dict()
        elif isinstance(result, dict):
            value = result
        else:
            value = {"result": result}
        return serializable(value)

    def _business_call(
        self,
        request: RunRequest,
        args: tuple[Any, ...],
        kwargs: dict[str, Any],
        cancel_event: threading.Event,
        callbacks: RunCallbacks,
    ) -> tuple[dict[str, Any], bool]:
        call_kwargs = dict(kwargs)
        call_kwargs.pop("cancelled", None)
        call_kwargs.pop("progress_callback", None)
        progress_state: dict[str, Any] = {"last": 0.0, "pending": None, "phase": ""}

        def report_progress(current: int, total: int, message: str) -> None:
            payload = (int(current), int(total), str(message))
            progress_state["pending"] = payload
            now = time.monotonic()
            complete = payload[1] > 0 and payload[0] >= payload[1]
            phase = payload[2].partition("】")[0] if payload[2].startswith("【") else ""
            phase_changed = request.tool_id == "material_collector" and phase != progress_state["phase"]
            if (
                not complete
                and not phase_changed
                and now - float(progress_state["last"]) < PROGRESS_UI_INTERVAL_SECONDS
            ):
                return
            progress_state["last"] = now
            progress_state["phase"] = phase
            progress_state["pending"] = None
            callbacks.progress(*payload)

        def flush_progress() -> None:
            payload = progress_state["pending"]
            if payload is not None:
                progress_state["pending"] = None
                callbacks.progress(*payload)

        bound = inspect.signature(request.function).bind_partial(*args, **call_kwargs)
        load_arguments = {
            name: value
            for name, value in bound.arguments.items()
            if name not in {"output_dir", "cancelled", "progress_callback"}
        }
        use_process = (
            self.process_isolation_available
            and should_use_process(request.tool_id, (), load_arguments)
        )
        if use_process:
            try:
                result = run_business_process(
                    module_name=request.function.__module__,
                    function_name=request.function.__name__,
                    args=args,
                    kwargs=call_kwargs,
                    cancel_event=cancel_event,
                    on_progress=report_progress,
                )
            except BusinessProcessStartError as exc:
                self.note_process_start_failure(exc)
                fallback_message = (
                    "独立后台进程不可用，已自动切换兼容后台模式继续处理。"
                )
                callbacks.log(fallback_message)
                runlog.log_line(f"{fallback_message} 原因：{exc}")
            else:
                flush_progress()
                return result.payload, True

        thread_call_kwargs = dict(call_kwargs)
        parameters = inspect.signature(request.function).parameters
        if "cancelled" in parameters:
            thread_call_kwargs["cancelled"] = cancel_event.is_set
        if "progress_callback" in parameters:
            thread_call_kwargs["progress_callback"] = report_progress
        result = request.function(*args, **thread_call_kwargs)
        flush_progress()
        return self._payload(result), False

    def _run(
        self,
        store: Any,
        request: RunRequest,
        callbacks: RunCallbacks,
        cancel_event: threading.Event,
    ) -> None:
        started_at = time.monotonic()
        draft = None
        started = False
        batch_id: str | None = None
        try:
            sources, _parameters, _legacy_output = context_from_call(
                request.function,
                request.args,
                request.kwargs,
            )
            draft = store.create_draft(
                group_name=request.group_name,
                tool_id=request.tool_id,
                tool_name=request.tool_name,
                business_description=request.description,
                business_period=time.strftime("%Y-%m-%d"),
            )
            batch_id = draft.summary.id
            with self._lock:
                self._active_batch_id = batch_id
            callbacks.log(
                f"开始 {request.tool_name}（{len(sources)} 个资料来源，自动留存在当前项目）"
            )

            progress_state = {"last": 0.0, "phase": ""}

            def import_progress(event: Any) -> None:
                now = time.monotonic()
                phase = str(getattr(event, "phase", "") or "")
                scanned = int(getattr(event, "files_scanned", 0) or 0)
                completed = int(getattr(event, "files_completed", 0) or 0)
                total = getattr(event, "files_total", None)
                bytes_copied = int(getattr(event, "bytes_copied", 0) or 0)
                bytes_total = getattr(event, "bytes_total", None)
                force = phase != progress_state["phase"] or phase == "finalizing"
                if not force and now - float(progress_state["last"]) < 0.25:
                    return
                progress_state["last"] = now
                progress_state["phase"] = phase
                if phase == "copying":
                    if bytes_total:
                        percent = min(100, int(bytes_copied * 100 / int(bytes_total)))
                        text = f"正在安全保存项目资料：{percent}%"
                    else:
                        text = f"正在安全保存项目资料：已完成 {completed} 个文件"
                elif phase == "finalizing":
                    text = "项目资料已复制完成，正在核对并登记..."
                else:
                    text = f"正在检查项目资料：已发现 {scanned} 个文件"
                callbacks.progress(completed, int(total or 0), text)

            replacements = import_project_run_sources(
                store,
                batch_id,
                sources,
                cancel_event,
                on_progress=import_progress,
            )
            if cancel_event.is_set():
                raise BusinessProcessCancelled("本次处理已停止。")
            old_upload_root = draft.directories["uploads"]
            running = store.start_batch(batch_id)
            started = True
            replacements = rebase_project_replacements(
                replacements,
                old_upload_root,
                running.directories["uploads"],
            )
            result_dir = store.result_directory(batch_id)
            call_args, call_kwargs = call_with_project_inputs(
                request.function,
                request.args,
                request.kwargs,
                replacements,
                result_dir,
                store,
                batch_id,
            )
            payload, isolated = self._business_call(
                request,
                call_args,
                call_kwargs,
                cancel_event,
                callbacks,
            )
            if cancel_event.is_set():
                raise BusinessProcessCancelled("本次处理已停止。")
            material_progress = request.tool_id == "material_collector"
            if material_progress:
                callbacks.progress(0, 2, "【登记结果】已完成 0/2 项；正在登记结果文件")
            store.register_results(batch_id, result_dir)
            if material_progress:
                callbacks.progress(1, 2, "【登记结果】已完成 1/2 项；正在保存批次状态")
            if cancel_event.is_set():
                raise BusinessProcessCancelled("本次处理已停止。")
            store.mark_success(batch_id)
            if material_progress:
                callbacks.progress(2, 2, "【登记结果】已完成 2/2 项；结果文件和批次状态均已保存")
            try:
                upload_path = Path(store.root) / running.directories["uploads"]
                if upload_path.is_dir() and not any(upload_path.iterdir()):
                    upload_path.rmdir()
            except OSError:
                pass
            elapsed = time.monotonic() - started_at
            runlog.log_line(
                f"完成 {request.tool_name}，耗时 {elapsed:.1f} 秒"
                + ("（独立进程）" if isolated else "（后台线程）")
            )
            callbacks.success(payload, result_dir, elapsed, isolated)
        except BaseException as exc:
            stopped = cancel_event.is_set() or isinstance(exc, BusinessProcessCancelled)
            finalization_error: BaseException | None = None
            if batch_id is not None:
                try:
                    if started:
                        if stopped:
                            store.mark_stopped(batch_id)
                        else:
                            store.mark_failed(batch_id, str(exc))
                    elif draft is not None:
                        store.move_to_trash(batch_id)
                    if not project_batch_is_closed(store, batch_id):
                        raise RuntimeError("项目批次仍未进入安全结束状态。")
                except BaseException as project_exc:
                    finalization_error = project_exc
                    runlog.log_exception("保存项目批次状态失败", project_exc)
            if finalization_error is not None:
                callbacks.error(
                    RuntimeError(
                        f"处理失败，且项目未能安全结案：{exc}；{finalization_error}"
                    )
                )
            elif stopped:
                callbacks.stopped()
            else:
                if isinstance(exc, BusinessProcessError) and exc.remote_traceback:
                    runlog.log_line(exc.remote_traceback)
                runlog.log_exception(f"{request.tool_name} 失败", exc)
                callbacks.error(exc)
        finally:
            with self._lock:
                self._active_batch_id = None
                self._cancel_event = None
                self._thread = None
            callbacks.finished()
