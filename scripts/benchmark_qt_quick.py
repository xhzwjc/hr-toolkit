from __future__ import annotations

import argparse
import json
import multiprocessing
import os
import statistics
import sys
import tempfile
import time
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def _percentile(values: list[float], percentile: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    index = min(len(ordered) - 1, max(0, round((len(ordered) - 1) * percentile)))
    return ordered[index]


def _cpu_worker(stop_event) -> None:
    accumulator = 0
    while not stop_event.is_set():
        for value in range(40000):
            accumulator = (accumulator + value * 17) % 10000019


def run_benchmark(
    *,
    file_count: int,
    workspace_file_count: int,
    duration_seconds: float,
    request_interval_ms: int,
    mode: str,
    resize_axis: str,
    render_loop: str,
    processing_load: bool,
    cache_buffer: int | None = None,
    scroll_step_px: int | None = None,
    scroll_target: str = "upload",
    tool_id: str = "",
) -> dict[str, object]:
    os.environ.setdefault("HR_TOOLKIT_SKIP_UPDATE", "1")
    os.environ.setdefault("QT_QUICK_CONTROLS_STYLE", "Basic")
    if render_loop != "production":
        os.environ["QSG_RENDER_LOOP"] = render_loop
    from hr_toolkit.gui_qt.main import _prepare_environment

    _prepare_environment()

    from hr_toolkit.gui_qt.compat import (
        QApplication,
        QObject,
        QQmlApplicationEngine,
        QTimer,
        QUrl,
        QT_MAJOR,
        delete_qobject,
    )
    from hr_toolkit.gui_qt.controller import AppController

    style = "Default" if QT_MAJOR == 5 else "Basic"
    os.environ["QT_QUICK_CONTROLS_STYLE"] = style
    app = QApplication.instance() or QApplication([sys.argv[0]])

    with tempfile.TemporaryDirectory(prefix="hr_qt_benchmark_") as temporary:
        temp_root = Path(temporary)
        AppController._settings_path = staticmethod(lambda: temp_root / "workspace-ui.json")
        controller = AppController()
        if tool_id:
            controller.selectTool(tool_id)
        engine = QQmlApplicationEngine()
        engine.rootContext().setContextProperty("controller", controller)
        qml_path = REPO_ROOT / "hr_toolkit" / "gui_qt" / "qml" / "Main.qml"
        load_started = time.perf_counter()
        engine.load(QUrl.fromLocalFile(str(qml_path)))
        startup_ms = (time.perf_counter() - load_started) * 1000.0
        roots = engine.rootObjects()
        if not roots:
            raise RuntimeError("Qt Quick benchmark could not load Main.qml")
        window = roots[0]

        items = [
            {
                "name": f"异动表_{index + 1:05d}.xlsx",
                "path": str(temp_root / f"异动表_{index + 1:05d}.xlsx"),
                "kind": "file",
                "detail": "XLSX",
            }
            for index in range(file_count)
        ]
        model_started = time.perf_counter()
        controller._input_model.set_items(items)
        model_load_ms = (time.perf_counter() - model_started) * 1000.0

        workspace_model_load_ms = 0.0
        if workspace_file_count:
            workspace_items = [
                {
                    "name": f"项目资料_{index + 1:05d}.xlsx",
                    "path": str(temp_root / "项目文件" / f"项目资料_{index + 1:05d}.xlsx"),
                    "isDir": False,
                    "depth": index % 3,
                    "expanded": False,
                    "hasChildren": False,
                    "detail": "XLSX · 性能测试资料",
                }
                for index in range(workspace_file_count)
            ]
            workspace_started = time.perf_counter()
            controller._workspace_model.set_items(workspace_items)
            workspace_model_load_ms = (time.perf_counter() - workspace_started) * 1000.0

        main_scroll = window.findChild(QObject, "mainScroll")
        input_list = window.findChild(QObject, "inputList")
        workspace_drawer = window.findChild(QObject, "workspaceDrawer")
        workspace_list = window.findChild(QObject, "workspaceList")
        workspace_button = window.findChild(QObject, "workspaceButton")
        if main_scroll is None or input_list is None:
            raise RuntimeError("Qt Quick benchmark targets are missing")
        if workspace_drawer is None or workspace_list is None or workspace_button is None:
            raise RuntimeError("Qt Quick workspace benchmark targets are missing")
        if cache_buffer is not None:
            input_list.setProperty("cacheBuffer", max(0, cache_buffer))
        if workspace_file_count:
            workspace_drawer.open()
        scroll_view = {
            "main": main_scroll,
            "workspace": workspace_list,
        }.get(scroll_target, input_list)

        frame_times: list[float] = []
        heartbeat_lateness: list[float] = []
        request_gaps: list[float] = []
        resize_requests = 0
        scroll_requests = 0
        scroll_mutations = 0
        workspace_anchor_errors: list[float] = []
        workspace_height_errors: list[float] = []
        progress_messages = 0
        action_started = 0.0
        deadline = 0.0
        last_request = 0.0
        expected_heartbeat = 0.0
        peak_rss = 0

        try:
            import psutil

            process = psutil.Process()
        except Exception:
            process = None

        def sample_memory() -> None:
            nonlocal peak_rss
            if process is not None:
                try:
                    peak_rss = max(peak_rss, int(process.memory_info().rss))
                except Exception:
                    pass

        def frame_swapped() -> None:
            now = time.perf_counter()
            if action_started and now <= deadline + 0.30:
                frame_times.append(now)

        frame_signal = getattr(window, "frameSwapped", None)
        if frame_signal is not None:
            frame_signal.connect(frame_swapped)

        heartbeat_timer = QTimer()
        heartbeat_timer.setInterval(16)

        def heartbeat() -> None:
            nonlocal expected_heartbeat
            now = time.perf_counter()
            if action_started and now <= deadline:
                heartbeat_lateness.append(max(0.0, (now - expected_heartbeat) * 1000.0))
                expected_heartbeat = now + 0.016

        heartbeat_timer.timeout.connect(heartbeat)

        memory_timer = QTimer()
        memory_timer.setInterval(25)
        memory_timer.timeout.connect(sample_memory)
        memory_timer.start()

        action_timer = QTimer()
        action_timer.setInterval(max(1, request_interval_ms))
        phase = {"value": 0, "scrollDirection": 1}

        def sample_workspace_geometry() -> None:
            # Sample on the next event-loop turn after the preceding resize.
            # QML bindings are evaluated between turns, so this measures what
            # can actually be presented rather than transient values inside
            # the same Python setWidth()/setHeight() callback.
            if not bool(workspace_drawer.property("opened")):
                return
            drawer_width = float(workspace_drawer.property("width") or 0.0)
            drawer_x = float(workspace_drawer.property("x") or 0.0)
            drawer_height = float(workspace_drawer.property("height") or 0.0)
            workspace_anchor_errors.append(
                abs(float(window.width()) - drawer_width - drawer_x)
            )
            workspace_height_errors.append(
                abs(float(window.height()) - drawer_height)
            )

        def perform_action() -> None:
            nonlocal resize_requests, scroll_requests, scroll_mutations, last_request
            now = time.perf_counter()
            if now >= deadline:
                action_timer.stop()
                return
            sample_workspace_geometry()
            if last_request:
                request_gaps.append(max(0.0, (now - last_request) * 1000.0))
            last_request = now
            value = phase["value"]
            phase["value"] += 1
            if mode in {"resize", "both"}:
                cycle = value % 120
                progress = cycle / 60.0 if cycle <= 60 else (120 - cycle) / 60.0
                next_width = round(1180 - 360 * progress)
                next_height = round(760 - 140 * progress)
                if resize_axis == "width":
                    window.setWidth(next_width)
                elif resize_axis == "height":
                    window.setHeight(next_height)
                else:
                    window.setWidth(next_width)
                    window.setHeight(next_height)
                resize_requests += 1
            if mode in {"scroll", "both"}:
                content_height = float(scroll_view.property("contentHeight") or 0.0)
                height = float(scroll_view.property("height") or 0.0)
                maximum = max(0.0, content_height - height)
                if maximum:
                    scroll_mutations += 1
                    if scroll_step_px is not None:
                        current = float(scroll_view.property("contentY") or 0.0)
                        next_value = current + phase["scrollDirection"] * scroll_step_px
                        if next_value >= maximum:
                            next_value = maximum
                            phase["scrollDirection"] = -1
                        elif next_value <= 0.0:
                            next_value = 0.0
                            phase["scrollDirection"] = 1
                        scroll_view.setProperty("contentY", next_value)
                    else:
                        scroll_phase = (value % 200) / 199.0
                        if (value // 200) % 2:
                            scroll_phase = 1.0 - scroll_phase
                        scroll_view.setProperty("contentY", maximum * scroll_phase)
                scroll_requests += 1

        action_timer.timeout.connect(perform_action)

        progress_timer = QTimer()
        progress_timer.setInterval(10)

        def add_progress() -> None:
            nonlocal progress_messages
            if time.perf_counter() >= deadline:
                progress_timer.stop()
                return
            progress_messages += 1
            controller._append_log(f"压力测试进度 {progress_messages}", "info")

        progress_timer.timeout.connect(add_progress)

        context = multiprocessing.get_context("spawn")
        process_stop = context.Event() if processing_load else None
        load_process = (
            context.Process(target=_cpu_worker, args=(process_stop,), daemon=False)
            if process_stop is not None
            else None
        )

        def begin() -> None:
            nonlocal action_started, deadline, last_request, expected_heartbeat
            action_started = time.perf_counter()
            deadline = action_started + duration_seconds
            last_request = action_started
            expected_heartbeat = action_started + 0.016
            if load_process is not None:
                load_process.start()
                progress_timer.start()
            heartbeat_timer.start()
            action_timer.start()
            perform_action()
            QTimer.singleShot(round(duration_seconds * 1000 + 320), app.quit)

        QTimer.singleShot(350, begin)
        execute = getattr(app, "exec", None) or app.exec_
        execute()

        if process_stop is not None:
            process_stop.set()
        if load_process is not None:
            load_process.join(timeout=2.0)
            if load_process.is_alive():
                load_process.terminate()
                load_process.join(timeout=1.0)
        sample_memory()

        frame_intervals = [
            (current - previous) * 1000.0
            for previous, current in zip(frame_times, frame_times[1:])
            if current <= deadline
        ]
        delegate_count = int(input_list.property("activeDelegateCount") or 0)
        workspace_delegate_count = (
            int(workspace_list.property("activeDelegateCount") or 0)
        )
        result = {
            "renderer": "Qt Quick",
            "render_loop": os.environ.get("QSG_RENDER_LOOP", "auto"),
            "qt_major": QT_MAJOR,
            "mode": mode,
            "resize_axis": resize_axis,
            "file_count": file_count,
            "workspace_file_count": workspace_file_count,
            "duration_seconds": duration_seconds,
            "request_interval_ms": request_interval_ms,
            "processing_load": processing_load,
            "cache_buffer": int(input_list.property("cacheBuffer") or 0),
            "scroll_step_px": scroll_step_px or 0,
            "scroll_target": scroll_target,
            "tool_id": tool_id or "default",
            "qml_startup_ms": round(startup_ms, 2),
            "model_load_ms": round(model_load_ms, 2),
            "workspace_model_load_ms": round(workspace_model_load_ms, 2),
            "resize_requests_sent": resize_requests,
            "scroll_requests_sent": scroll_requests,
            "scroll_mutations_applied": scroll_mutations,
            "scroll_content_height": round(
                float(scroll_view.property("contentHeight") or 0.0), 2
            ),
            "scroll_view_height": round(
                float(scroll_view.property("height") or 0.0), 2
            ),
            "request_gap_median_ms": round(
                statistics.median(request_gaps) if request_gaps else 0.0,
                2,
            ),
            "request_gap_p95_ms": round(_percentile(request_gaps, 0.95), 2),
            "request_gap_max_ms": round(max(request_gaps, default=0.0), 2),
            "heartbeat_count": len(heartbeat_lateness),
            "heartbeat_lateness_p95_ms": round(_percentile(heartbeat_lateness, 0.95), 2),
            "heartbeat_lateness_max_ms": round(max(heartbeat_lateness, default=0.0), 2),
            "rendered_frames": len(frame_intervals),
            "frame_interval_median_ms": round(
                statistics.median(frame_intervals) if frame_intervals else 0.0,
                2,
            ),
            "frame_interval_p95_ms": round(_percentile(frame_intervals, 0.95), 2),
            "frame_interval_max_ms": round(max(frame_intervals, default=0.0), 2),
            "visible_delegate_objects": delegate_count,
            "visible_workspace_delegate_objects": workspace_delegate_count,
            "workspace_popup_opened": bool(workspace_drawer.property("opened")),
            "workspace_popup_visible": bool(workspace_drawer.property("visible")),
            "workspace_popup_x": round(float(workspace_drawer.property("x") or 0.0), 2),
            "workspace_popup_width": round(float(workspace_drawer.property("width") or 0.0), 2),
            "workspace_anchor_error_max_px": round(
                max(workspace_anchor_errors, default=0.0), 2
            ),
            "workspace_height_error_max_px": round(
                max(workspace_height_errors, default=0.0), 2
            ),
            "workspace_button_x": round(float(workspace_button.property("x") or 0.0), 2),
            "workspace_button_width": round(float(workspace_button.property("width") or 0.0), 2),
            "workspace_button_implicit_width": round(
                float(workspace_button.property("implicitWidth") or 0.0),
                2,
            ),
            "progress_messages": progress_messages,
            "log_rows_retained": len(controller._log_model),
            "peak_rss_mb": round(peak_rss / 1024 / 1024, 2) if peak_rss else 0.0,
        }
        controller.close()
        for root_object in engine.rootObjects():
            delete_qobject(root_object)
        return result


def main() -> None:
    parser = argparse.ArgumentParser(
        description="量化 Qt Quick 缩放、虚拟列表滚动和主线程响应。"
    )
    parser.add_argument("--files", type=int, default=10000)
    parser.add_argument("--workspace-files", type=int, default=0)
    parser.add_argument("--duration", type=float, default=2.0)
    parser.add_argument("--interval", type=int, default=8)
    parser.add_argument("--mode", choices=("resize", "scroll", "both"), default="both")
    parser.add_argument(
        "--resize-axis",
        choices=("both", "width", "height"),
        default="both",
    )
    parser.add_argument(
        "--render-loop",
        choices=("production", "basic", "threaded"),
        default="production",
    )
    parser.add_argument("--processing", action="store_true")
    parser.add_argument(
        "--cache-buffer",
        type=int,
        help="仅用于比较虚拟列表预取像素，不设置时使用生产值。",
    )
    parser.add_argument(
        "--scroll-step",
        type=int,
        help=(
            "按固定像素模拟滚轮；不设置时模拟跨越整个列表的滚动条拖动。"
        ),
    )
    parser.add_argument(
        "--scroll-target",
        choices=("upload", "main", "workspace"),
        default="upload",
        help="选择上传列表、主内容区或项目文件列表作为滚动目标。",
    )
    parser.add_argument(
        "--tool",
        default="",
        help="启动基准前切换到指定工具，例如 material_collector。",
    )
    args = parser.parse_args()
    payload = run_benchmark(
        file_count=max(1, args.files),
        workspace_file_count=max(0, args.workspace_files),
        duration_seconds=max(0.25, args.duration),
        request_interval_ms=max(1, args.interval),
        mode=args.mode,
        resize_axis=args.resize_axis,
        render_loop=args.render_loop,
        processing_load=args.processing,
        cache_buffer=args.cache_buffer,
        scroll_step_px=max(1, args.scroll_step) if args.scroll_step else None,
        scroll_target=args.scroll_target,
        tool_id=args.tool.strip(),
    )
    print(json.dumps(payload, ensure_ascii=False, sort_keys=True))


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
