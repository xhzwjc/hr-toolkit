from __future__ import annotations

import argparse
import json
import os
import statistics
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path


REPO_ROOT = Path(
    os.environ.get("HR_TOOLKIT_BENCHMARK_SOURCE", "")
    or Path(__file__).resolve().parent.parent
).resolve()
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from hr_toolkit.gui.app import HRToolkitApp


def _percentile(values: list[float], percentile: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    index = min(
        len(ordered) - 1,
        max(0, int(round((len(ordered) - 1) * percentile))),
    )
    return ordered[index]


def run_benchmark(
    *,
    file_count: int,
    duration_seconds: float,
    request_interval_ms: int,
    prewarm_ms: int = 0,
    resize_policy: str = "auto",
    status_message_count: int = 0,
) -> dict[str, object]:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        HRToolkitApp._workspace_settings_path = staticmethod(
            lambda: temp_path / "workspace-ui.json"
        )
        HRToolkitApp._check_updates_on_startup = lambda self: None

        root = tk.Tk()
        root.geometry("1180x760")
        app = HRToolkitApp(root)
        root.deiconify()
        root.update()
        app._dismiss_startup_loading_screen()
        app._select_tool("personnel_change_merge")
        root.update()
        app._flush_live_window_resize()
        root.update()

        pointer_state = {"down": False}
        if resize_policy == "live":
            app._window_resize_live_layout_enabled = True
            app._window_resize_pointer_state_reader = None
        elif resize_policy == "deferred":
            app._window_resize_live_layout_enabled = False
        elif resize_policy != "auto":
            raise ValueError(f"unsupported resize policy: {resize_policy}")
        if not app._window_resize_live_layout_enabled:
            # ``geometry()`` does not press a physical mouse button.  Model a
            # real native border gesture so auto/deferred measurements do not
            # publish layouts in artificial gaps between scheduled requests.
            app._window_resize_pointer_state_reader = lambda: pointer_state["down"]

        try:
            paths: list[Path] = []
            for index in range(file_count):
                path = temp_path / f"异动表_{index + 1:04d}.xlsx"
                path.touch()
                paths.append(path)
            if paths:
                app.change_input_paths = paths
                app._sync_input_path_text()
                app._refresh_upload_card()
                root.update()

            if prewarm_ms > 0:
                root.after(prewarm_ms, root.quit)
                root.mainloop()

            root_configures = 0
            root_configures_during_drag = 0
            native_tree_configures = 0
            native_tree_configures_during_drag = 0
            last_native_tree_configure_at = 0.0
            canvas_window_resizes = 0
            root_frame_path = str(app._root_frame)
            original_itemconfigure = app._right_canvas.itemconfig

            def counted_itemconfigure(item, *args, **kwargs):
                nonlocal canvas_window_resizes
                if item == app._right_canvas_window and (
                    "width" in kwargs or "height" in kwargs
                ):
                    canvas_window_resizes += 1
                return original_itemconfigure(item, *args, **kwargs)

            app._right_canvas.itemconfig = counted_itemconfigure

            def count_root_configure(event) -> None:
                nonlocal root_configures, root_configures_during_drag
                nonlocal native_tree_configures, native_tree_configures_during_drag
                nonlocal last_native_tree_configure_at
                during_drag = time.perf_counter() < deadline
                if event.widget is root:
                    root_configures += 1
                    if during_drag:
                        root_configures_during_drag += 1
                    return
                widget_path = str(event.widget)
                if widget_path == root_frame_path or widget_path.startswith(
                    root_frame_path + "."
                ):
                    native_tree_configures += 1
                    last_native_tree_configure_at = time.perf_counter()
                    if during_drag:
                        native_tree_configures_during_drag += 1

            root.bind("<Configure>", count_root_configure, add="+")
            started_at = time.perf_counter()
            deadline = started_at + duration_seconds
            sent_requests = 0
            request_gaps: list[float] = []
            heartbeat_lateness: list[float] = []
            last_request = started_at
            expected_heartbeat = started_at + 0.016
            hidden_tree_samples = 0
            settle_latency_ms = None
            layout_commit_latency_ms = None
            settle_candidate_at = None
            status_queue_at_release = None
            if not app._window_resize_live_layout_enabled:
                pointer_state["down"] = True
            for index in range(max(0, status_message_count)):
                app.status_queue.put(
                    ("progress", app._tool_run_token, f"性能测试进度 {index + 1}")
                )

            def send_resize() -> None:
                nonlocal sent_requests, last_request, hidden_tree_samples
                now = time.perf_counter()
                if now >= deadline:
                    pointer_state["down"] = False
                    return
                if app._root_frame.winfo_manager() != "place":
                    hidden_tree_samples += 1
                request_gaps.append(max(0.0, (now - last_request) * 1000.0))
                last_request = now
                phase = (now - started_at) % 1.0
                progress = phase * 2.0 if phase < 0.5 else (1.0 - phase) * 2.0
                width = round(1180 - 360 * progress)
                height = round(760 - 140 * progress)
                root.geometry(f"{width}x{height}")
                sent_requests += 1
                root.after(request_interval_ms, send_resize)

            def heartbeat() -> None:
                nonlocal expected_heartbeat
                now = time.perf_counter()
                if now >= deadline:
                    return
                heartbeat_lateness.append(
                    max(0.0, (now - expected_heartbeat) * 1000.0)
                )
                expected_heartbeat = now + 0.016
                root.after(16, heartbeat)

            def release_drag() -> None:
                nonlocal status_queue_at_release
                status_queue_at_release = app.status_queue.qsize()
                pointer_state["down"] = False

            def observe_settle() -> None:
                nonlocal settle_latency_ms, layout_commit_latency_ms, settle_candidate_at
                now = time.perf_counter()
                if now < deadline:
                    root.after(1, observe_settle)
                    return
                root_size = (root.winfo_width(), root.winfo_height())
                layout_committed = (
                    not app._window_resize_active
                    and app._live_root_frame_size == root_size
                )
                if layout_committed and layout_commit_latency_ms is None:
                    layout_commit_latency_ms = max(
                        0.0,
                        (now - deadline) * 1000.0,
                    )
                if (
                    layout_committed
                    and now - last_native_tree_configure_at >= 0.032
                ):
                    if settle_candidate_at is None:
                        settle_candidate_at = now
                    elif now - settle_candidate_at >= 0.016:
                        settle_latency_ms = max(0.0, (now - deadline) * 1000.0)
                        return
                else:
                    settle_candidate_at = None
                root.after(1, observe_settle)

            root.after(0, send_resize)
            root.after(16, heartbeat)
            root.after(max(1, int(duration_seconds * 1000)), release_drag)
            root.after(1, observe_settle)
            root.after(max(500, int(duration_seconds * 1000) + 500), root.quit)
            root.mainloop()

            canvas_size = (
                app._right_canvas.winfo_width(),
                app._right_canvas.winfo_height(),
            )
            frame_size = tuple(app._last_canvas_window_size)
            root_size = (root.winfo_width(), root.winfo_height())
            root_frame_size = (
                app._root_frame.winfo_width(),
                app._root_frame.winfo_height(),
            )
            measured_gaps = request_gaps[1:]
            return {
                "file_count": file_count,
                "duration_seconds": duration_seconds,
                "request_interval_ms": request_interval_ms,
                "prewarm_ms": prewarm_ms,
                "status_message_count": max(0, status_message_count),
                "status_queue_at_release": status_queue_at_release,
                "status_queue_after_settle": app.status_queue.qsize(),
                "pending_logs_after_settle": len(app._pending_log_entries),
                "resize_policy": (
                    "live" if app._window_resize_live_layout_enabled else "deferred"
                ),
                "resize_requests_sent": sent_requests,
                "root_configure_events": root_configures,
                "root_configure_events_during_drag": root_configures_during_drag,
                "native_tree_configure_events": native_tree_configures,
                "native_tree_configure_events_during_drag": (
                    native_tree_configures_during_drag
                ),
                "request_gap_median_ms": round(
                    statistics.median(measured_gaps) if measured_gaps else 0.0,
                    2,
                ),
                "request_gap_p95_ms": round(_percentile(measured_gaps, 0.95), 2),
                "request_gap_max_ms": round(max(measured_gaps, default=0.0), 2),
                "heartbeat_count": len(heartbeat_lateness),
                "heartbeat_lateness_p95_ms": round(
                    _percentile(heartbeat_lateness, 0.95),
                    2,
                ),
                "heartbeat_lateness_max_ms": round(
                    max(heartbeat_lateness, default=0.0),
                    2,
                ),
                "canvas_window_resizes": canvas_window_resizes,
                "live_layout_frames": getattr(
                    app, "_window_resize_layout_frames", 0
                ),
                "slow_layout_frames": getattr(
                    app, "_window_resize_slow_frames", 0
                ),
                "adaptive_frame_interval_ms": getattr(
                    app, "_window_resize_frame_interval_ms", 0
                ),
                "settle_latency_ms": (
                    round(settle_latency_ms, 2)
                    if settle_latency_ms is not None
                    else None
                ),
                "layout_commit_latency_ms": (
                    round(layout_commit_latency_ms, 2)
                    if layout_commit_latency_ms is not None
                    else None
                ),
                "real_tree_hidden_samples": hidden_tree_samples,
                "real_tree_visible": app._root_frame.winfo_manager() == "place",
                "root_frame_size_settled": root_frame_size == root_size,
                "canvas_width_fits_viewport": frame_size[0] <= canvas_size[0],
                "canvas_window_centered": app._last_canvas_window_x
                == max(0, (canvas_size[0] - frame_size[0]) // 2),
                "canvas_width": canvas_size[0],
                "canvas_window_width": frame_size[0],
            }
        finally:
            app.destroy()
            root.destroy()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="量化连续拖动主窗口边框时的 Tk 主线程响应。"
    )
    parser.add_argument("--files", type=int, default=40)
    parser.add_argument("--duration", type=float, default=2.0)
    parser.add_argument("--interval", type=int, default=8)
    parser.add_argument(
        "--prewarm",
        type=int,
        default=0,
        help="缩放前保持窗口静止的毫秒数。",
    )
    parser.add_argument(
        "--policy",
        choices=("auto", "live", "deferred"),
        default="auto",
        help="缩放布局策略；deferred 为 Windows/macOS 的保留式控件树路径。",
    )
    parser.add_argument(
        "--status-messages",
        type=int,
        default=0,
        help="缩放开始前注入的后台进度消息数。",
    )
    args = parser.parse_args()
    result = run_benchmark(
        file_count=max(0, args.files),
        duration_seconds=max(0.25, args.duration),
        request_interval_ms=max(1, args.interval),
        prewarm_ms=max(0, args.prewarm),
        resize_policy=args.policy,
        status_message_count=max(0, args.status_messages),
    )
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))


if __name__ == "__main__":
    main()
