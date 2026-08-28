from __future__ import annotations

import argparse
import json
import statistics
import sys
import tempfile
import threading
import time
import tkinter as tk
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from hr_toolkit.gui.app import HRToolkitApp


def _percentile(values: list[float], percentile: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    index = min(len(ordered) - 1, max(0, int(round((len(ordered) - 1) * percentile))))
    return ordered[index]


def _walk_widgets(widget) -> list[object]:
    widgets: list[object] = [widget]
    for child in widget.winfo_children():
        widgets.extend(_walk_widgets(child))
    return widgets


def run_benchmark(
    *,
    file_count: int,
    duration_seconds: float,
    processing_load: bool = False,
    scroll_target: str = "outer",
) -> dict[str, object]:
    root = tk.Tk()
    root.geometry("1100x720")
    app = HRToolkitApp(root)
    root.deiconify()
    app._select_tool("personnel_change_merge")

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            paths: list[Path] = []
            for index in range(file_count):
                path = Path(temp_dir) / f"异动表_{index + 1:02d}.xlsx"
                path.write_bytes(b"benchmark")
                paths.append(path)
            app.change_input_paths = paths
            app._sync_input_path_text()

            refresh_samples: list[float] = []
            for _ in range(5):
                started = time.perf_counter()
                app._refresh_upload_card()
                root.update()
                refresh_samples.append((time.perf_counter() - started) * 1000.0)

            widgets = _walk_widgets(app._right_canvas)
            canvas_count = sum(isinstance(widget, tk.Canvas) for widget in widgets)
            upload_surface_count = len(app.upload_body.winfo_children())
            upload_canvas = getattr(app, "_upload_items_canvas", None)
            upload_canvas_item_count = len(upload_canvas.find_all()) if upload_canvas is not None else 0

            applied_outer_scrolls = 0
            original_scroll = app._right_canvas.yview_scroll

            def counted_scroll(units: int, mode: str):
                nonlocal applied_outer_scrolls
                applied_outer_scrolls += 1
                return original_scroll(units, mode)

            app._right_canvas.yview_scroll = counted_scroll
            applied_upload_scrolls = 0
            if upload_canvas is not None:
                original_upload_scroll = upload_canvas.yview_scroll

                def counted_upload_scroll(units: int, mode: str):
                    nonlocal applied_upload_scrolls
                    applied_upload_scrolls += 1
                    return original_upload_scroll(units, mode)

                upload_canvas.yview_scroll = counted_upload_scroll
            log_insert_calls = 0
            log_inserts_during_scroll = 0
            original_log_insert = app.log_text.insert

            def counted_log_insert(*args, **kwargs):
                nonlocal log_insert_calls, log_inserts_during_scroll
                log_insert_calls += 1
                if getattr(app, "_scroll_active", False):
                    log_inserts_during_scroll += 1
                return original_log_insert(*args, **kwargs)

            app.log_text.insert = counted_log_insert
            interval_ms = 4
            heartbeat_ms = 16
            duration_ms = max(250, int(duration_seconds * 1000))
            started_at = time.perf_counter()
            deadline = started_at + duration_ms / 1000.0
            sent_events = 0
            heartbeat_lateness: list[float] = []
            expected_heartbeat = started_at + heartbeat_ms / 1000.0
            stop_load = threading.Event()
            load_threads: list[threading.Thread] = []
            cpu_batches_completed = 0

            if processing_load:
                app._begin_tool_run()
                token = app._tool_run_token

                def cpu_worker() -> None:
                    nonlocal cpu_batches_completed
                    accumulator = 0
                    while not stop_load.is_set() and time.perf_counter() < deadline:
                        for value in range(25000):
                            accumulator = (accumulator + value * 17) % 10000019
                        cpu_batches_completed += 1

                def progress_worker() -> None:
                    progress_index = 0
                    while not stop_load.is_set() and time.perf_counter() < deadline:
                        app.status_queue.put(
                            ("progress", token, f"正在处理压力测试数据：{progress_index}")
                        )
                        progress_index += 1
                        stop_load.wait(0.01)

                for target in (cpu_worker, progress_worker):
                    worker = threading.Thread(target=target, daemon=True)
                    worker.start()
                    load_threads.append(worker)

            def send_wheel() -> None:
                nonlocal sent_events
                now = time.perf_counter()
                if now >= deadline:
                    return
                elapsed_ms = (now - started_at) * 1000.0
                direction = -120 if int(elapsed_ms // 250) % 2 == 0 else 120
                event_widget = (
                    upload_canvas
                    if scroll_target == "upload" and upload_canvas is not None
                    else app._right_canvas
                )
                event_widget.event_generate("<MouseWheel>", delta=direction)
                sent_events += 1
                root.after(interval_ms, send_wheel)

            def heartbeat() -> None:
                nonlocal expected_heartbeat
                now = time.perf_counter()
                if now >= deadline:
                    return
                heartbeat_lateness.append(max(0.0, (now - expected_heartbeat) * 1000.0))
                expected_heartbeat = now + heartbeat_ms / 1000.0
                root.after(heartbeat_ms, heartbeat)

            root.after(0, send_wheel)
            root.after(heartbeat_ms, heartbeat)
            root.after(duration_ms + 100, root.quit)
            root.mainloop()
            stop_load.set()
            for worker in load_threads:
                worker.join(timeout=1.0)
            if processing_load:
                app._finish_tool_run()
            controller = getattr(app, "_right_scroll_controller", None)
            if controller is not None:
                controller.flush_pending()
            upload_controller = getattr(app, "_upload_items_scroll_controller", None)
            if upload_controller is not None:
                upload_controller.flush_pending()
            # Allow the scroll-settle timer to release deferred log paints so
            # the benchmark also proves that prioritising input does not drop
            # progress messages.
            root.after(250, root.quit)
            root.mainloop()
            root.update_idletasks()

            applied_scrolls = applied_outer_scrolls + applied_upload_scrolls

            return {
                "file_count": file_count,
                "duration_seconds": duration_seconds,
                "upload_refresh_median_ms": round(statistics.median(refresh_samples), 2),
                "upload_refresh_max_ms": round(max(refresh_samples), 2),
                "right_panel_widget_count": len(widgets),
                "right_panel_canvas_count": canvas_count,
                "upload_surface_count": upload_surface_count,
                "upload_canvas_item_count": upload_canvas_item_count,
                "upload_body_height": app._upload_body_height,
                "upload_rendered_range": getattr(app, "_upload_items_rendered_range", None),
                "processing_load": processing_load,
                "scroll_target": scroll_target,
                "wheel_events_sent": sent_events,
                "canvas_scrolls_applied": applied_scrolls,
                "outer_scrolls_applied": applied_outer_scrolls,
                "upload_scrolls_applied": applied_upload_scrolls,
                "scroll_coalescing_ratio": round(applied_scrolls / max(sent_events, 1), 4),
                "heartbeat_lateness_p95_ms": round(_percentile(heartbeat_lateness, 0.95), 2),
                "heartbeat_lateness_max_ms": round(max(heartbeat_lateness, default=0.0), 2),
                "pending_log_entries": len(app._pending_log_entries),
                "cpu_batches_completed": cpu_batches_completed,
                "log_insert_calls": log_insert_calls,
                "log_inserts_during_scroll": log_inserts_during_scroll,
            }
    finally:
        app.destroy()
        root.destroy()


def main() -> None:
    parser = argparse.ArgumentParser(description="量化工具页多文件滚动与主线程响应性能。")
    parser.add_argument("--files", type=int, default=12, help="模拟上传文件数，默认 12。")
    parser.add_argument("--duration", type=float, default=2.0, help="连续滚动秒数，默认 2 秒。")
    parser.add_argument("--processing", action="store_true", help="叠加纯 Python 计算和高频进度消息。")
    parser.add_argument(
        "--target",
        choices=("outer", "upload"),
        default="outer",
        help="滚动外层工具页或大文件列表。",
    )
    args = parser.parse_args()
    result = run_benchmark(
        file_count=max(1, args.files),
        duration_seconds=max(0.25, args.duration),
        processing_load=args.processing,
        scroll_target=args.target,
    )
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))


if __name__ == "__main__":
    main()
