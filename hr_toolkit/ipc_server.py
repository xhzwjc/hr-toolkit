"""JSON-RPC 2.0 IPC server for HR Toolkit.

Exposes project management and tool execution to desktop frontends
(such as the native C# WPF client) via standard I/O streams.
"""

from __future__ import annotations

import json
import os
import sys
import threading
import time
from pathlib import Path
from typing import Any, Sequence, TextIO

from hr_toolkit import runlog
from hr_toolkit.desktop_contract import NAV_GROUPS, TOOL_NAV_ITEMS
from hr_toolkit.desktop_helpers import (
    default_workspace_project_name,
    workspace_project_creation_target,
)
from hr_toolkit.material_preferences import (
    BUILTIN_MATERIAL_PRESETS,
    BUILTIN_MATERIALS,
    MaterialPreferences,
)
from hr_toolkit.project_store import ProjectStore
from hr_toolkit.run_coordinator import (
    ProjectRunCoordinator,
    RunCallbacks,
    RunRequest,
)
from hr_toolkit.tools.registry import (
    ensure_default_tools_registered,
    get_tool_by_id,
)


class IpcServer:
    """Manages project store and tool execution via JSON-RPC 2.0."""

    def __init__(self, out_stream: TextIO | None = None) -> None:
        self._out_stream = out_stream or sys.stdout
        self._out_lock = threading.Lock()
        self._project_store: ProjectStore | None = None
        self._project_path: Path | None = None
        self._coordinator = ProjectRunCoordinator()
        self._material_prefs: MaterialPreferences = MaterialPreferences()
        self._running = True
        ensure_default_tools_registered()

    # -------------------------------------------------------------------------
    # JSON-RPC Output Protocol
    # -------------------------------------------------------------------------

    def send_response(self, request_id: Any, result: Any = None, error: dict[str, Any] | None = None) -> None:
        payload: dict[str, Any] = {"jsonrpc": "2.0", "id": request_id}
        if error is not None:
            payload["error"] = error
        else:
            payload["result"] = result
        self._write_line(payload)

    def send_event(self, event_name: str, data: Any = None) -> None:
        payload: dict[str, Any] = {
            "jsonrpc": "2.0",
            "method": "event",
            "params": {"name": event_name, "data": data or {}},
        }
        self._write_line(payload)

    def _write_line(self, obj: dict[str, Any]) -> None:
        line = json.dumps(obj, ensure_ascii=True)
        with self._out_lock:
            try:
                self._out_stream.write(line + "\n")
                self._out_stream.flush()
            except (BrokenPipeError, OSError):
                self._running = False

    # -------------------------------------------------------------------------
    # Method Dispatcher
    # -------------------------------------------------------------------------

    def handle_request(self, message: dict[str, Any]) -> None:
        req_id = message.get("id")
        method = message.get("method")
        params = message.get("params") or {}

        if not method:
            if req_id is not None:
                self.send_response(req_id, error={"code": -32600, "message": "Invalid Request: method is missing"})
            return

        handler = getattr(self, f"rpc_{method}", None)
        if handler is None or not callable(handler):
            if req_id is not None:
                self.send_response(req_id, error={"code": -32601, "message": f"Method not found: {method}"})
            return

        try:
            result = handler(params)
            if req_id is not None:
                self.send_response(req_id, result=result)
        except Exception as exc:
            runlog.log_exception(f"IPC method {method} execution error", exc)
            if req_id is not None:
                self.send_response(req_id, error={"code": -32000, "message": str(exc)})

    # -------------------------------------------------------------------------
    # RPC Handlers
    # -------------------------------------------------------------------------

    def rpc_ping(self, _params: dict[str, Any]) -> dict[str, Any]:
        return {"status": "pong", "time": time.time(), "version": "1.0.0"}

    def rpc_get_metadata(self, _params: dict[str, Any]) -> dict[str, Any]:
        """Return navigation, tool specs, and preferences for the WPF UI."""
        tools_list = []
        for tool_id, tool_name in TOOL_NAV_ITEMS:
            spec = get_tool_by_id(tool_id)
            tools_list.append({
                "id": tool_id,
                "name": tool_name,
                "group": spec.group if spec else "",
                "description": spec.help_text if spec else "",
                "cli_command": spec.cli_command if spec else "",
                "multi_input": spec.multi_input if spec else False,
            })

        default_docs = Path.home() / "Documents"
        default_parent = str(default_docs if default_docs.is_dir() else Path.home())

        return {
            "nav_groups": [{"name": g[0], "tool_ids": list(g[1])} for g in NAV_GROUPS],
            "tools": tools_list,
            "builtin_materials": list(BUILTIN_MATERIALS),
            "material_presets": {k: list(v) for k, v in BUILTIN_MATERIAL_PRESETS.items()},
            "custom_materials": list(self._material_prefs.custom_materials),
            "custom_presets": {k: list(v) for k, v in self._material_prefs.custom_presets.items()},
            "default_project_name": default_workspace_project_name(),
            "default_project_parent": default_parent,
        }

    def rpc_create_project(self, params: dict[str, Any]) -> dict[str, Any]:
        name = str(params.get("name", "")).strip()
        parent = str(params.get("parent", "")).strip()
        target, error = workspace_project_creation_target(parent, name)
        if error or target is None:
            raise ValueError(error or "无效的项目创建位置或名称。")

        store = ProjectStore.create(target, name)
        if self._project_store is not None:
            self._project_store.close()
        self._project_store = store
        self._project_path = target

        result = self._project_summary()
        self.send_event("project_changed", result)
        return result

    def rpc_open_project(self, params: dict[str, Any]) -> dict[str, Any]:
        path_str = str(params.get("path", "")).strip()
        path = Path(path_str).resolve()
        if not path.is_dir():
            raise FileNotFoundError(f"项目文件夹不存在：{path}")

        store = ProjectStore.open(path, writable=True, read_only_fallback=True)
        if self._project_store is not None:
            self._project_store.close()
        self._project_store = store
        self._project_path = path

        result = self._project_summary()
        self.send_event("project_changed", result)
        return result

    def rpc_close_project(self, _params: dict[str, Any]) -> dict[str, Any]:
        if self._project_store is not None:
            self._project_store.close()
            self._project_store = None
            self._project_path = None
        result = self._project_summary()
        self.send_event("project_changed", result)
        return result

    def rpc_get_project_status(self, _params: dict[str, Any]) -> dict[str, Any]:
        return self._project_summary()

    def _project_summary(self) -> dict[str, Any]:
        if self._project_store is None:
            return {"has_project": False, "name": "", "path": "", "writable": False, "read_only_reason": ""}
        return {
            "has_project": True,
            "name": self._project_store.workspace.name,
            "path": str(self._project_path or ""),
            "writable": bool(self._project_store.writable),
            "read_only_reason": self._project_store.workspace.read_only_reason or "",
        }

    def rpc_list_workspace_files(self, _params: dict[str, Any]) -> list[dict[str, Any]]:
        """List files and directories inside the current workspace."""
        if self._project_store is None:
            return []
        items = []
        try:
            # 1. Common directory files
            common_root = self._project_store.workspace.common_root
            if common_root.is_dir():
                for p in sorted(common_root.rglob("*")):
                    if p.is_file():
                        stat = p.stat()
                        items.append({
                            "name": p.name,
                            "relative_path": str(p.relative_to(self._project_store.root)),
                            "category": "common",
                            "batch_id": None,
                            "size": stat.st_size,
                            "mtime": stat.st_mtime,
                        })
            # 2. Batch directories
            for summary, directories in self._project_store.list_batch_locations():
                for cat, d_path in directories.items():
                    if d_path.is_dir():
                        for p in sorted(d_path.rglob("*")):
                            if p.is_file():
                                stat = p.stat()
                                items.append({
                                    "name": p.name,
                                    "relative_path": str(p.relative_to(self._project_store.root)),
                                    "category": cat,
                                    "batch_id": summary.id,
                                    "size": stat.st_size,
                                    "mtime": stat.st_mtime,
                                })
        except Exception as exc:
            runlog.log_exception("list_workspace_files error", exc)
        return items

    def rpc_import_workspace_files(self, params: dict[str, Any]) -> dict[str, Any]:
        """Safely import external files/directories into the project workspace."""
        if self._project_store is None:
            raise RuntimeError("未打开任何工作项目。")
        if not self._project_store.writable:
            raise PermissionError("当前项目为只读状态，无法导入文件。")

        source_strs = params.get("sources", [])
        sources = [Path(s) for s in source_strs if Path(s).exists()]
        if not sources:
            raise ValueError("没有找到有效的待导入源文件或目录。")

        common_dir = self._project_store.workspace.common_root
        self._project_store.import_to_directory(common_dir, sources)
        self.send_event("workspace_changed", {})
        return {"imported_count": len(sources)}

    def rpc_list_trash(self, _params: dict[str, Any]) -> list[dict[str, Any]]:
        if self._project_store is None:
            return []
        trash_items = []
        for detail in self._project_store.list_trash_details():
            trash_items.append({
                "batch_id": detail.batch_id,
                "tool_id": detail.tool_id,
                "tool_name": detail.tool_name,
                "directory_name": detail.directory_name,
                "business_description": detail.business_description,
                "deleted_at": detail.deleted_at,
                "file_count": detail.file_count,
                "total_size": detail.total_size,
            })
        return trash_items

    def rpc_restore_trash(self, params: dict[str, Any]) -> dict[str, Any]:
        if self._project_store is None:
            raise RuntimeError("未打开工作项目。")
        batch_id = str(params.get("batch_id", "")).strip()
        if not batch_id:
            raise ValueError("必须指定 batch_id。")
        self._project_store.restore_from_trash(batch_id)
        self.send_event("workspace_changed", {})
        return {"restored": True, "batch_id": batch_id}

    def rpc_move_to_trash(self, params: dict[str, Any]) -> dict[str, Any]:
        if self._project_store is None:
            raise RuntimeError("未打开工作项目。")
        batch_id = str(params.get("batch_id", "")).strip()
        if not batch_id:
            raise ValueError("必须指定 batch_id。")
        self._project_store.move_to_trash(batch_id)
        self.send_event("workspace_changed", {})
        return {"moved": True, "batch_id": batch_id}

    # -------------------------------------------------------------------------
    # Tool Execution
    # -------------------------------------------------------------------------

    def rpc_run_tool(self, params: dict[str, Any]) -> dict[str, Any]:
        if self._project_store is None:
            raise RuntimeError("请先新建或打开一个工作项目。")
        if not self._project_store.writable:
            raise PermissionError("当前项目只能查看，无法执行处理任务。")
        if self._coordinator.running:
            raise RuntimeError("已有任务正在执行，请等待完成或先停止当前任务。")

        tool_id = str(params.get("tool_id", "")).strip()
        spec = get_tool_by_id(tool_id)
        if spec is None:
            raise ValueError(f"未知的工具 ID: {tool_id}")

        input_paths = [Path(p) for p in params.get("inputs", [])]
        support_file = Path(params["support_file"]) if params.get("support_file") else None
        options = params.get("options") or {}

        args, kwargs = self._build_tool_arguments(tool_id, input_paths, support_file, options)

        callbacks = RunCallbacks(
            log=lambda msg, lvl="info": self.send_event("log", {"message": str(msg), "level": str(lvl)}),
            progress=lambda ratio, msg="": self.send_event("progress", {"ratio": float(ratio), "message": str(msg)}),
            success=lambda *payload: self.send_event("finished", {"success": True, "payload": payload}),
            error=lambda err: self.send_event("finished", {"success": False, "error": str(err)}),
            finished=lambda: self.send_event("workspace_changed", {}),
        )

        request = RunRequest(
            tool_id=spec.tool_id,
            tool_name=spec.name,
            group_name=spec.group,
            description=f"WPF 客户端调用：{spec.name}",
            function=spec.entry_point,
            args=args,
            kwargs=kwargs,
        )

        started = self._coordinator.start(self._project_store, request, callbacks)
        return {"started": started, "tool_id": tool_id}

    def rpc_cancel_tool(self, _params: dict[str, Any]) -> dict[str, Any]:
        if not self._coordinator.running:
            return {"cancelled": False, "message": "当前没有正在执行的任务。"}
        self._coordinator.cancel()
        return {"cancelled": True}

    def _build_tool_arguments(
        self,
        tool_id: str,
        input_paths: list[Path],
        support_file: Path | None,
        options: dict[str, Any],
    ) -> tuple[tuple[Any, ...], dict[str, Any]]:
        """Map tool parameters to Python function signature."""
        output_dir = self._project_path

        if tool_id == "social_security":
            roster = support_file or (input_paths[1] if len(input_paths) > 1 else None)
            inputs = input_paths[:1] if support_file else input_paths
            return (inputs, roster, output_dir), {"dry_run": bool(options.get("dry_run", False))}

        if tool_id == "insurance_ledger":
            return (input_paths, output_dir), {}

        if tool_id == "data_statistics":
            kwargs: dict[str, Any] = {
                "staff_file": support_file,
                "remark_unit": options.get("remark_unit", "day"),
                "include_business_trip": bool(options.get("include_business_trip", False)),
            }
            if options.get("week_start") and options.get("week_end"):
                kwargs["week_range"] = (str(options["week_start"]), str(options["week_end"]))
            if options.get("month_start") and options.get("month_end"):
                kwargs["month_range"] = (str(options["month_start"]), str(options["month_end"]))
            return (input_paths, output_dir), kwargs

        if tool_id == "salary_split":
            salary_file = input_paths[0] if input_paths else None
            return (salary_file, output_dir), {
                "output_format": options.get("output_format", "xlsx"),
            }

        if tool_id == "salary_merge":
            return (input_paths, output_dir), {}

        if tool_id == "personnel_change_merge":
            return (input_paths, output_dir), {"roster_file": support_file}

        if tool_id == "archive_import":
            return (input_paths, output_dir), {"mode": options.get("mode", "all")}

        if tool_id == "material_collector":
            return (input_paths, output_dir), {
                "material_types": options.get("material_types", []),
                "library_mode": options.get("library_mode", "person_folder"),
                "use_ocr_cache": bool(options.get("use_ocr_cache", True)),
            }

        if tool_id == "folder_rename":
            target = input_paths[0] if input_paths else None
            return (target,), {
                "excel_file": support_file,
                "mode": options.get("mode", "append"),
                "file_type": options.get("file_type", "folder"),
                "preview": bool(options.get("preview", False)),
            }

        # Generic fallback
        return (input_paths, output_dir), options

    # -------------------------------------------------------------------------
    # Main Event Loop
    # -------------------------------------------------------------------------

    def run_loop(self, in_stream: TextIO | None = None) -> int:
        stream = in_stream or sys.stdin
        while self._running:
            try:
                line = stream.readline()
                if not line:
                    break
                stripped = line.strip().lstrip("\ufeff")
                if not stripped:
                    continue
                message = json.loads(stripped)
                if isinstance(message, dict):
                    self.handle_request(message)
            except json.JSONDecodeError as exc:
                self.send_response(None, error={"code": -32700, "message": f"Parse error: {exc}"})
            except Exception as exc:
                runlog.log_exception("Unhandled error in IPC main loop", exc)

        if self._project_store is not None:
            try:
                self._project_store.close()
            except Exception:
                pass
        return 0


def main(argv: Sequence[str] | None = None) -> int:
    try:
        if hasattr(sys.stdin, "reconfigure"):
            sys.stdin.reconfigure(encoding="utf-8-sig", errors="replace")
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    server = IpcServer()
    return server.run_loop()


if __name__ == "__main__":
    sys.exit(main())
