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
            native_tree_configures = 0
            canvas_window_resizes = 0
            root_frame_widget = root.winfo_children()[0]
            root_frame_path = str(root_frame_widget)
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
                nonlocal root_configures, native_tree_configures
                if event.widget is root:
                    root_configures += 1
                    return
                widget_path = str(event.widget)
                if widget_path == root_frame_path or widget_path.startswith(
                    root_frame_path + "."
                ):
                    native_tree_configures += 1

            root.bind("<Configure>", count_root_configure, add="+")
            started_at = time.perf_counter()
            deadline = started_at + duration_seconds
            sent_requests = 0
            request_gaps: list[float] = []
            heartbeat_lateness: list[float] = []
            last_request = started_at
            expected_heartbeat = started_at + 0.016
            hidden_tree_samples = 0

            def send_resize() -> None:
                nonlocal sent_requests, last_request, hidden_tree_samples
                now = time.perf_counter()
                if now >= deadline:
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

            root.after(0, send_resize)
            root.after(16, heartbeat)
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
                "resize_requests_sent": sent_requests,
                "root_configure_events": root_configures,
                "native_tree_configure_events": native_tree_configures,
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
    args = parser.parse_args()
    result = run_benchmark(
        file_count=max(0, args.files),
        duration_seconds=max(0.25, args.duration),
        request_interval_ms=max(1, args.interval),
        prewarm_ms=max(0, args.prewarm),
    )
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))


if __name__ == "__main__":
    main()
