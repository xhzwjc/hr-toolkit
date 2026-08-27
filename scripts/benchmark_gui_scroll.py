from __future__ import annotations

import argparse
import json
import statistics
import sys
import tempfile
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


def run_benchmark(*, file_count: int, duration_seconds: float) -> dict[str, object]:
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

            applied_scrolls = 0
            original_scroll = app._right_canvas.yview_scroll

            def counted_scroll(units: int, mode: str):
                nonlocal applied_scrolls
                applied_scrolls += 1
                return original_scroll(units, mode)

            app._right_canvas.yview_scroll = counted_scroll
            interval_ms = 4
            heartbeat_ms = 16
            duration_ms = max(250, int(duration_seconds * 1000))
            started_at = time.perf_counter()
            deadline = started_at + duration_ms / 1000.0
            sent_events = 0
            heartbeat_lateness: list[float] = []
            expected_heartbeat = started_at + heartbeat_ms / 1000.0

            def send_wheel() -> None:
                nonlocal sent_events
                now = time.perf_counter()
                if now >= deadline:
                    return
                elapsed_ms = (now - started_at) * 1000.0
                direction = -120 if int(elapsed_ms // 250) % 2 == 0 else 120
                app._right_canvas.event_generate("<MouseWheel>", delta=direction)
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
            controller = getattr(app, "_right_scroll_controller", None)
            if controller is not None:
                controller.flush_pending()
            root.update_idletasks()

            return {
                "file_count": file_count,
                "duration_seconds": duration_seconds,
                "upload_refresh_median_ms": round(statistics.median(refresh_samples), 2),
                "upload_refresh_max_ms": round(max(refresh_samples), 2),
                "right_panel_widget_count": len(widgets),
                "right_panel_canvas_count": canvas_count,
                "upload_surface_count": upload_surface_count,
                "wheel_events_sent": sent_events,
                "canvas_scrolls_applied": applied_scrolls,
                "scroll_coalescing_ratio": round(applied_scrolls / max(sent_events, 1), 4),
                "heartbeat_lateness_p95_ms": round(_percentile(heartbeat_lateness, 0.95), 2),
                "heartbeat_lateness_max_ms": round(max(heartbeat_lateness, default=0.0), 2),
            }
    finally:
        app.destroy()
        root.destroy()


def main() -> None:
    parser = argparse.ArgumentParser(description="量化工具页多文件滚动与主线程响应性能。")
    parser.add_argument("--files", type=int, default=12, help="模拟上传文件数，默认 12。")
    parser.add_argument("--duration", type=float, default=2.0, help="连续滚动秒数，默认 2 秒。")
    args = parser.parse_args()
    result = run_benchmark(
        file_count=max(1, args.files),
        duration_seconds=max(0.25, args.duration),
    )
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))


if __name__ == "__main__":
    main()
