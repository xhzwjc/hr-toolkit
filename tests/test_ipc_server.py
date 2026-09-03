"""Unit tests for the JSON-RPC 2.0 IPC server."""

from __future__ import annotations

import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from hr_toolkit.ipc_server import IpcServer
from hr_toolkit.project_store import ProjectStore


class IpcServerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.out_stream = io.StringIO()
        self.server = IpcServer(out_stream=self.out_stream)

    def _send_request(self, method: str, params: dict | None = None, req_id: int = 1) -> dict:
        request = {"jsonrpc": "2.0", "id": req_id, "method": method, "params": params or {}}
        self.server.handle_request(request)
        lines = [line for line in self.out_stream.getvalue().splitlines() if line.strip()]
        self.assertTrue(lines, f"Expected output for request {method}")
        return json.loads(lines[-1])

    def test_ping(self) -> None:
        resp = self._send_request("ping")
        self.assertEqual(resp.get("id"), 1)
        self.assertEqual(resp.get("result", {}).get("status"), "pong")

    def test_get_metadata(self) -> None:
        resp = self._send_request("get_metadata")
        result = resp.get("result", {})
        self.assertIn("nav_groups", result)
        self.assertIn("tools", result)
        self.assertIn("builtin_materials", result)
        self.assertIn("default_project_name", result)

        tool_ids = [t["id"] for t in result["tools"]]
        self.assertIn("social_security", tool_ids)
        self.assertIn("data_statistics", tool_ids)
        self.assertIn("material_collector", tool_ids)

    def test_unknown_method_returns_rpc_error(self) -> None:
        resp = self._send_request("non_existent_method", req_id=99)
        self.assertEqual(resp.get("id"), 99)
        self.assertIn("error", resp)
        self.assertEqual(resp["error"]["code"], -32601)

    def test_project_lifecycle(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_root = Path(tmp)
            # 1. Create project
            resp = self._send_request("create_project", {"name": "新项目", "parent": str(tmp_root)})
            result = resp.get("result", {})
            self.assertTrue(result.get("has_project"))
            self.assertEqual(result.get("name"), "新项目")
            self.assertTrue(result.get("writable"))
            self.assertTrue((tmp_root / "新项目").is_dir())

            # 2. Get status
            self.out_stream.seek(0)
            self.out_stream.truncate(0)
            status_resp = self._send_request("get_project_status")
            self.assertTrue(status_resp.get("result", {}).get("has_project"))

            # 3. Import file to workspace
            sample_file = tmp_root / "test_doc.txt"
            sample_file.write_text("hello world", encoding="utf-8")
            self.out_stream.seek(0)
            self.out_stream.truncate(0)
            import_resp = self._send_request("import_workspace_files", {"sources": [str(sample_file)]})
            self.assertEqual(import_resp.get("result", {}).get("imported_count"), 1)

            # 4. List files
            self.out_stream.seek(0)
            self.out_stream.truncate(0)
            list_resp = self._send_request("list_workspace_files")
            items = list_resp.get("result", [])
            self.assertTrue(len(items) >= 1)
            self.assertIn("test_doc.txt", [Path(i["relative_path"]).name for i in items])

            # 5. Close project
            self.out_stream.seek(0)
            self.out_stream.truncate(0)
            close_resp = self._send_request("close_project")
            self.assertFalse(close_resp.get("result", {}).get("has_project"))

            # 6. Reopen project
            self.out_stream.seek(0)
            self.out_stream.truncate(0)
            open_resp = self._send_request("open_project", {"path": str(tmp_root / "新项目")})
            self.assertTrue(open_resp.get("result", {}).get("has_project"))
            # 7. Test trash listing (initially empty)
            self.out_stream.seek(0)
            self.out_stream.truncate(0)
            trash_resp = self._send_request("list_trash")
            self.assertEqual(trash_resp.get("result"), [])

    def test_run_tool_without_project_errors(self) -> None:
        resp = self._send_request("run_tool", {"tool_id": "social_security"})
        self.assertIn("error", resp)
        self.assertIn("请先新建或打开一个工作项目", resp["error"]["message"])

    def test_run_loop_stream_processing(self) -> None:
        input_lines = (
            '{"jsonrpc": "2.0", "id": 10, "method": "ping"}\n'
            '{"jsonrpc": "2.0", "id": 11, "method": "get_project_status"}\n'
        )
        in_stream = io.StringIO(input_lines)
        out_stream = io.StringIO()
        server = IpcServer(out_stream=out_stream)
        code = server.run_loop(in_stream=in_stream)
        self.assertEqual(code, 0)

        lines = [json.loads(line) for line in out_stream.getvalue().splitlines() if line.strip()]
        self.assertEqual(len(lines), 2)
        self.assertEqual(lines[0]["id"], 10)
        self.assertEqual(lines[0]["result"]["status"], "pong")
        self.assertEqual(lines[1]["id"], 11)
        self.assertFalse(lines[1]["result"]["has_project"])


if __name__ == "__main__":
    unittest.main()
