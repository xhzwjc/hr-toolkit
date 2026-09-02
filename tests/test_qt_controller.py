from __future__ import annotations

import os
import unittest
from datetime import date
from types import SimpleNamespace


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("HR_TOOLKIT_SKIP_UPDATE", "1")

try:
    from hr_toolkit.gui_qt.compat import QCoreApplication
    from hr_toolkit.gui_qt.controller import AppController
except ImportError:
    QCoreApplication = None
    AppController = None


@unittest.skipUnless(AppController is not None, "PySide GUI runtime is not installed")
class QtControllerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.application = QCoreApplication.instance() or QCoreApplication([])

    def controller(self):
        value = AppController()
        value._save_workspace_preferences = lambda: None
        return value

    def test_date_presets_keep_legacy_calendar_ranges(self) -> None:
        controller = self.controller()
        controller.selectTool("data_statistics")
        controller.applyDatePreset("week", "this_week")
        state = controller._form_states[("data_statistics", "default")]
        today = date.today()
        start = date.fromisoformat(state["week_start"])
        end = date.fromisoformat(state["week_end"])
        self.assertEqual(start.weekday(), 0)
        self.assertEqual(end.weekday(), 6)
        self.assertLessEqual(start, today)
        self.assertGreaterEqual(end, today)

        controller.applyDatePreset("month", "this_month")
        self.assertEqual(date.fromisoformat(state["month_start"]).day, 1)
        self.assertEqual(date.fromisoformat(state["month_end"]).month, today.month)
        controller.applyDatePreset("month", "clear")
        self.assertEqual(state["month_start"], "")
        self.assertEqual(state["month_end"], "")
        controller.close()

    def test_legacy_history_reuse_accepts_every_current_navigation_tool(self) -> None:
        for tool_id in (
            "social_security",
            "insurance_ledger",
            "data_statistics",
            "salary_split",
            "salary_merge",
            "material_collector",
        ):
            with self.subTest(tool_id=tool_id):
                controller = self.controller()
                controller._history_selected = SimpleNamespace(
                    summary=SimpleNamespace(tool_id=tool_id, mode=None),
                    inputs=(),
                )
                controller.reuseHistory()
                self.assertEqual(controller.currentTool, tool_id)
                controller.close()

    def test_close_requires_confirmation_while_background_work_is_active(self) -> None:
        controller = self.controller()
        prompts = []
        controller.confirmationRequested.connect(lambda *args: prompts.append(args))
        controller._busy = True
        self.assertFalse(controller.requestClose())
        self.assertEqual(len(prompts), 1)
        self.assertFalse(controller.requestClose())
        self.assertEqual(len(prompts), 1)
        token = prompts[0][2]
        controller.confirmAction(token, False)
        controller._busy = False
        self.assertTrue(controller.requestClose())
        self.assertTrue(controller._closed)

    def test_close_waits_for_project_and_storage_workers(self) -> None:
        for attribute in ("_project_opening", "_history_busy", "_trash_busy"):
            with self.subTest(attribute=attribute):
                controller = self.controller()
                prompts = []
                controller.confirmationRequested.connect(lambda *args: prompts.append(args))
                setattr(controller, attribute, True)
                self.assertFalse(controller.requestClose())
                self.assertEqual(len(prompts), 1)
                controller.confirmAction(prompts[0][2], False)
                setattr(controller, attribute, False)
                controller.close()

    def test_cancelling_workspace_import_does_not_cancel_update_download(self) -> None:
        controller = self.controller()
        workspace_cancel = SimpleNamespace(called=False)
        update_cancel = SimpleNamespace(called=False)
        workspace_cancel.set = lambda: setattr(workspace_cancel, "called", True)
        update_cancel.set = lambda: setattr(update_cancel, "called", True)
        controller._workspace_cancel_event = workspace_cancel
        controller._update_cancel_event = update_cancel
        controller.cancelWorkspaceImport()
        self.assertTrue(workspace_cancel.called)
        self.assertFalse(update_cancel.called)
        controller._workspace_cancel_event = None
        controller._update_cancel_event = None
        controller.close()


if __name__ == "__main__":
    unittest.main()
