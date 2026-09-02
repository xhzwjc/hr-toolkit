"""Qt-facing controller; business and project rules remain in existing modules."""

from __future__ import annotations

import json
import os
import sys
import threading
import time
import calendar
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from hr_toolkit import __version__, runlog
from hr_toolkit.app_update import (
    UpdateCancelledError,
    UpdateInfo,
    check_for_update,
    cleanup_stale_update_files,
    download_update_package,
    launch_update_replacement,
    resolve_download_url,
    update_check_enabled,
)
from hr_toolkit.desktop_helpers import (
    default_workspace_project_name,
    open_path,
    workspace_project_create_error_message,
    workspace_project_creation_target,
)
from hr_toolkit.history_store import HISTORY_PAGE_SIZE, HistoryStore, TaskDetail
from hr_toolkit.material_preferences import MaterialPreferences
from hr_toolkit.project_store import ImportCancelled, ProjectStore
from hr_toolkit.run_coordinator import (
    ProjectRunCoordinator,
    RunCallbacks,
    RunRequest,
)
from hr_toolkit.tutorial_content import tutorial_groups

from .compat import (
    Property,
    QCoreApplication,
    QDesktopServices,
    QFileDialog,
    QObject,
    QTimer,
    QUrl,
    Signal,
    Slot,
    constant_property,
)
from .form_specs import (
    DEFAULT_VARIANTS,
    FormValidationError,
    ToolInvocation,
    build_invocation,
    default_values,
    spec_for,
    variants_for,
)
from .models import HistoryModel, InputFileModel, LogModel, TrashModel, WorkspaceModel


NAV_GROUPS = (
    ("社保与保险", (("social_security", "社保明细与汇总"), ("insurance_ledger", "保险台账与预警"))),
    ("考勤与统计", (("data_statistics", "考勤与周月报"),)),
    ("薪酬管理", (("salary_split", "工资表拆分"), ("salary_merge", "多月工资合并"))),
    (
        "人员与档案",
        (
            ("personnel_change_merge", "异动汇总"),
            ("archive_import", "档案入库"),
            ("material_collector", "员工资料打包"),
            ("folder_rename", "资料文件夹改名"),
        ),
    ),
)

WORKSPACE_HIDDEN_NAMES = frozenset({".hrtoolkit", ".DS_Store", "Thumbs.db", "desktop.ini"})
WORKSPACE_HIDDEN_SUFFIXES = (".partial", ".tmp", ".temp", ".lock")
WORKSPACE_SEARCH_LIMIT = 500
MAX_LOG_ROWS = 1000
HISTORY_STATUS_LABELS = {
    "draft": "未开始",
    "running": "处理中",
    "success": "已完成",
    "failed": "未完成",
    "stopped": "已停止",
}


class AppController(QObject):
    specChanged = Signal()
    projectChanged = Signal()
    busyChanged = Signal()
    runButtonTextChanged = Signal()
    workspaceBusyChanged = Signal()
    workspaceChanged = Signal()
    workspaceSelectionChanged = Signal()
    supportChanged = Signal()
    formRevisionChanged = Signal()
    lastResultChanged = Signal()
    historyChanged = Signal()
    trashChanged = Signal()
    materialChanged = Signal()
    updateChanged = Signal()
    notificationRequested = Signal(str, str, str)
    confirmationRequested = Signal(str, str, str)
    projectCreationRequested = Signal(str, str)
    textInputRequested = Signal(str, str, str, str)
    resizeSample = Signal(float)

    _projectOpened = Signal(int, object, str)
    _projectOpenFailed = Signal(int, str)
    _workspaceItemsReady = Signal(int, object)
    _workspaceChildrenReady = Signal(int, int, str, int, object)
    _logIncoming = Signal(str, str)
    _runProgress = Signal(int, int, str)
    _runSuccess = Signal(object, str, float, bool)
    _runError = Signal(str)
    _runStopped = Signal()
    _runFinished = Signal()
    _previewReady = Signal(object)
    _previewFailed = Signal(str)
    _workspaceImportFinished = Signal(bool, str)
    _historyListReady = Signal(int, object, int, str)
    _historyDetailReady = Signal(int, object, str)
    _historyActionFinished = Signal(str, bool, str)
    _trashReady = Signal(int, object, str)
    _trashActionFinished = Signal(bool, str)
    _updateResult = Signal(str, object)
    _updateProgressIncoming = Signal(int, int)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._nav_id = "social_security"
        self._variants: dict[str, str] = dict(DEFAULT_VARIANTS)
        self._spec = spec_for(self._nav_id)
        self._form_states: dict[tuple[str, str], dict[str, Any]] = {}
        self._input_states: dict[tuple[str, str], list[Path]] = {}
        self._support_states: dict[tuple[str, str], str] = {}
        self._form_revision = 0
        self._input_model = InputFileModel(self)
        self._log_model = LogModel(self)
        self._workspace_model = WorkspaceModel(self)
        self._history_model = HistoryModel(self)
        self._trash_model = TrashModel(self)
        self._workspace_items: list[dict[str, Any]] = []
        self._workspace_child_loads: set[str] = set()
        self._workspace_generation = 0
        self._workspace_scope = "all"
        self._workspace_search = ""
        self._workspace_selected_path: Path | None = None
        self._workspace_selected_item: dict[str, Any] | None = None
        self._workspace_expanded = False
        self._workspace_busy = False
        self._workspace_cancel_event: threading.Event | None = None
        self._project_store: ProjectStore | None = None
        self._project_path: Path | None = None
        self._project_opening = False
        self._project_generation = 0
        self._recent_projects: list[Path] = []
        self._busy = False
        self._run_coordinator = ProjectRunCoordinator()
        self._preview_cancel_event: threading.Event | None = None
        self._pending_preview: dict[str, Any] | None = None
        self._pending_confirmation: str | None = None
        self._pending_confirmation_action: tuple[str, Any] | None = None
        self._pending_text_action: str | None = None
        self._pending_text_payload: Any = None
        self._last_result_dir: Path | None = None
        self._last_run_by_key: dict[tuple[str, str], tuple[str, bool]] = {}
        self._original_switch_interval: float | None = None
        self._closed = False
        self._material_preferences = MaterialPreferences()
        self._material_preset_name = ""
        self._history_store: HistoryStore | None = None
        self._history_init_attempted = False
        self._history_generation = 0
        self._history_page = 0
        self._history_total = 0
        self._history_search = ""
        self._history_tool_id = ""
        self._history_date_filter = "全部时间"
        self._history_busy = False
        self._history_message = ""
        self._history_selected: TaskDetail | None = None
        self._history_detail: dict[str, Any] = {}
        self._trash_generation = 0
        self._trash_busy = False
        self._trash_search = ""
        self._trash_items: list[Any] = []
        self._trash_selected_id = ""
        self._update_busy = False
        self._update_status = ""
        self._update_progress = -1.0
        self._pending_update: UpdateInfo | None = None
        self._update_cancel_event: threading.Event | None = None
        self._workspace_scan_cancel_event: threading.Event | None = None
        self._shutdown_requested = False
        self._shutdown_wait_started = 0.0

        self._ensure_state(self._spec)
        self._sync_input_model()
        self._append_log(self._spec.log_text, "info")

        self._projectOpened.connect(self._apply_project_open)
        self._projectOpenFailed.connect(self._apply_project_error)
        self._workspaceItemsReady.connect(self._apply_workspace_items)
        self._workspaceChildrenReady.connect(self._apply_workspace_children)
        self._logIncoming.connect(self._append_log)
        self._runProgress.connect(self._apply_run_progress)
        self._runSuccess.connect(self._apply_run_success)
        self._runError.connect(self._apply_run_error)
        self._runStopped.connect(self._apply_run_stopped)
        self._runFinished.connect(self._apply_run_finished)
        self._previewReady.connect(self._apply_preview)
        self._previewFailed.connect(self._apply_preview_error)
        self._workspaceImportFinished.connect(self._apply_workspace_import_result)
        self._historyListReady.connect(self._apply_history_list)
        self._historyDetailReady.connect(self._apply_history_detail)
        self._historyActionFinished.connect(self._apply_history_action)
        self._trashReady.connect(self._apply_trash_list)
        self._trashActionFinished.connect(self._apply_trash_action)
        self._updateResult.connect(self._apply_update_result)
        self._updateProgressIncoming.connect(self._apply_update_progress)

        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(300)
        self._search_timer.timeout.connect(self.refreshWorkspace)
        self.specChanged.connect(self.runButtonTextChanged.emit)
        self.busyChanged.connect(self.runButtonTextChanged.emit)

    def _state_key(self) -> tuple[str, str]:
        return self._spec.nav_id, self._spec.variant

    def _ensure_state(self, spec) -> None:
        key = (spec.nav_id, spec.variant)
        if key not in self._form_states:
            values = default_values(spec)
            if spec.tool_id == "material_collector":
                values["material_types"] = list(self._material_preferences.available_materials)
            self._form_states[key] = values
        self._input_states.setdefault(key, [])
        self._support_states.setdefault(key, "")

    @constant_property(str)
    def appVersion(self) -> str:
        return __version__

    @constant_property("QVariantList")
    def navGroups(self):
        return [
            {
                "name": group,
                "items": [{"id": tool_id, "label": label} for tool_id, label in items],
            }
            for group, items in NAV_GROUPS
        ]

    @constant_property("QVariantList")
    def tutorialGroups(self):
        return tutorial_groups()

    @Property(str, notify=specChanged)
    def currentTool(self) -> str:
        return self._spec.nav_id

    @Property(str, notify=specChanged)
    def currentVariant(self) -> str:
        return self._spec.variant

    @Property(str, notify=specChanged)
    def toolTitle(self) -> str:
        return self._spec.title

    @Property(str, notify=specChanged)
    def toolDescription(self) -> str:
        return self._spec.description

    @Property(str, notify=specChanged)
    def toolGroup(self) -> str:
        return self._spec.group

    @Property(str, notify=specChanged)
    def inputLabel(self) -> str:
        return self._spec.input_label

    @Property(str, notify=specChanged)
    def inputHint(self) -> str:
        return self._spec.input_hint

    @Property(str, notify=specChanged)
    def inputDropTitle(self) -> str:
        return self._spec.input_drop_title

    @Property(bool, notify=specChanged)
    def inputAllowsFiles(self) -> bool:
        return self._spec.input_mode != "directory_single"

    @Property(bool, notify=specChanged)
    def inputAllowsFolder(self) -> bool:
        return self._spec.input_mode != "excel_single"

    @Property(bool, notify=specChanged)
    def inputAllowsMultiple(self) -> bool:
        return self._spec.input_mode == "excel_archive_multi"

    @Property(str, notify=specChanged)
    def supportLabel(self) -> str:
        return self._spec.support_label

    @Property(str, notify=specChanged)
    def supportButtonText(self) -> str:
        return self._spec.support_button

    @Property(bool, notify=specChanged)
    def hasSupportField(self) -> bool:
        if not self._spec.support_id:
            return False
        if self._spec.tool_id == "folder_rename":
            return self._form_states[self._state_key()].get("rename_mode") == "excel"
        return True

    @Property(bool, notify=specChanged)
    def supportAllowsFolder(self) -> bool:
        return self._spec.support_mode in {"excel_or_folder", "excel_archive_or_folder"}

    @Property(str, notify=supportChanged)
    def supportPath(self) -> str:
        return self._support_states[self._state_key()]

    @Property("QVariantList", notify=specChanged)
    def variants(self):
        specs = variants_for(self._spec.nav_id)
        if len(specs) < 2:
            return []
        labels = {
            ("personnel_change_merge", "merge"): "异动汇总",
            ("personnel_change_merge", "roster"): "花名册更新",
            ("archive_import", "import"): "档案入库",
            ("archive_import", "export"): "档案表生成",
        }
        return [
            {"id": item.variant, "label": labels[(item.nav_id, item.variant)]}
            for item in specs
        ]

    @Property("QVariantList", notify=formRevisionChanged)
    def formFields(self):
        values = self._form_states[self._state_key()]
        payload = []
        for source in self._spec.fields:
            field = dict(source)
            if field.get("kind") == "date_range":
                field["startValue"] = values.get(field["startId"], "")
                field["endValue"] = values.get(field["endId"], "")
            else:
                field["value"] = values.get(field["id"], field.get("default", ""))
            field["visible"] = self._field_visible(field["id"], values)
            if field["kind"] == "materials":
                selected = set(values.get("material_types") or [])
                field["options"] = [
                    {"label": name, "value": name, "selected": name in selected}
                    for name in self._material_preferences.available_materials
                ]
            payload.append(field)
        return payload

    def _field_visible(self, field_id: str, values: dict[str, Any]) -> bool:
        if self._spec.tool_id == "folder_rename":
            mode = values.get("rename_mode") or "append"
            if field_id == "target_name":
                return mode in {"append", "remove", "replace"}
            if field_id == "rename_text":
                return mode in {"append", "remove"}
            if field_id == "replacement_name":
                return mode == "replace"
        if self._spec.tool_id == "material_collector" and field_id == "material_types":
            return not bool(values.get("collect_all", True))
        return True

    @Property(int, notify=formRevisionChanged)
    def formRevision(self) -> int:
        return self._form_revision

    @constant_property(QObject)
    def inputModel(self):
        return self._input_model

    @constant_property(QObject)
    def logModel(self):
        return self._log_model

    @constant_property(QObject)
    def workspaceModel(self):
        return self._workspace_model

    @constant_property(QObject)
    def historyModel(self):
        return self._history_model

    @constant_property(QObject)
    def trashModel(self):
        return self._trash_model

    @Property(bool, notify=historyChanged)
    def historyBusy(self) -> bool:
        return self._history_busy

    @Property(str, notify=historyChanged)
    def historyMessage(self) -> str:
        return self._history_message

    @Property(str, notify=historyChanged)
    def historyPageText(self) -> str:
        pages = max(1, (self._history_total + HISTORY_PAGE_SIZE - 1) // HISTORY_PAGE_SIZE)
        return f"第 {min(self._history_page + 1, pages)} / {pages} 页"

    @Property(bool, notify=historyChanged)
    def historyHasPrevious(self) -> bool:
        return self._history_page > 0

    @Property(bool, notify=historyChanged)
    def historyHasNext(self) -> bool:
        return (self._history_page + 1) * HISTORY_PAGE_SIZE < self._history_total

    @Property("QVariantMap", notify=historyChanged)
    def historyDetail(self):
        return dict(self._history_detail)

    @constant_property("QVariantList")
    def historyToolOptions(self):
        values = [{"label": "全部功能", "value": ""}]
        seen: set[str] = set()
        for _group, items in NAV_GROUPS:
            for tool_id, label in items:
                if tool_id not in seen:
                    values.append({"label": label, "value": tool_id})
                    seen.add(tool_id)
        return values

    @constant_property("QVariantList")
    def historyDateOptions(self):
        return ["全部时间", "今天", "最近7天", "最近30天", "今年"]

    @Property(bool, notify=trashChanged)
    def trashBusy(self) -> bool:
        return self._trash_busy

    @Property(str, notify=trashChanged)
    def trashSelectedId(self) -> str:
        return self._trash_selected_id

    @Property(bool, notify=materialChanged)
    def materialEditorAvailable(self) -> bool:
        return self._spec.tool_id == "material_collector"

    @Property("QVariantList", notify=materialChanged)
    def materialPresets(self):
        return list(self._material_preferences.preset_names)

    @Property("QVariantList", notify=materialChanged)
    def customMaterials(self):
        return list(self._material_preferences.custom_materials)

    @Property(str, notify=materialChanged)
    def materialPresetName(self) -> str:
        return self._material_preset_name

    @Property(bool, notify=updateChanged)
    def updateBusy(self) -> bool:
        return self._update_busy

    @Property(str, notify=updateChanged)
    def updateStatus(self) -> str:
        return self._update_status

    @Property(float, notify=updateChanged)
    def updateProgress(self) -> float:
        return self._update_progress

    @Property(bool, notify=busyChanged)
    def busy(self) -> bool:
        return self._busy

    @Property(str, notify=runButtonTextChanged)
    def runButtonText(self) -> str:
        return "停止" if self._busy else self._spec.run_text

    @Property(str, notify=projectChanged)
    def projectName(self) -> str:
        if self._project_store is None:
            return "未打开工作项目"
        return str(self._project_store.workspace.name)

    @Property(str, notify=projectChanged)
    def projectPath(self) -> str:
        return "" if self._project_path is None else str(self._project_path)

    @Property(bool, notify=projectChanged)
    def hasProject(self) -> bool:
        return self._project_store is not None

    @Property(bool, notify=projectChanged)
    def projectWritable(self) -> bool:
        return bool(self._project_store is not None and self._project_store.writable)

    @constant_property(str)
    def defaultProjectName(self) -> str:
        return default_workspace_project_name()

    @Property(str, notify=projectChanged)
    def defaultProjectParent(self) -> str:
        if self._recent_projects:
            return str(self._recent_projects[0].parent)
        documents = Path.home() / "Documents"
        return str(documents if documents.is_dir() else Path.home())

    @Property("QVariantList", notify=projectChanged)
    def recentProjects(self):
        return [
            {"name": path.name, "path": str(path)}
            for path in self._recent_projects[:8]
        ]

    @Property(bool, notify=workspaceChanged)
    def workspaceExpanded(self) -> bool:
        return self._workspace_expanded

    @Property(str, notify=workspaceChanged)
    def workspaceScope(self) -> str:
        return self._workspace_scope

    @Property(bool, notify=workspaceBusyChanged)
    def workspaceBusy(self) -> bool:
        return self._workspace_busy

    @Property(bool, notify=workspaceSelectionChanged)
    def workspaceSelectionAvailable(self) -> bool:
        return self._workspace_selected_item is not None

    @Property(str, notify=workspaceSelectionChanged)
    def workspaceSelectedName(self) -> str:
        item = self._workspace_selected_item
        return "" if item is None else str(item.get("name") or "")

    @Property(str, notify=workspaceSelectionChanged)
    def workspaceSelectedDetail(self) -> str:
        item = self._workspace_selected_item
        if item is None:
            return "双击可以打开文件；文件夹可展开查看。"
        detail = str(item.get("detail") or "").strip()
        if detail:
            return detail
        return "文件夹" if item.get("isDir") else "文件"

    @Property(bool, notify=lastResultChanged)
    def canOpenLastResult(self) -> bool:
        return bool(self._last_result_dir is not None and self._last_result_dir.exists())

    @Property(str, notify=specChanged)
    def lastRunText(self) -> str:
        record = self._last_run_by_key.get(self._state_key())
        if record is None:
            return ""
        stamp, success = record
        return f"上次运行 {stamp} · {'成功' if success else '失败'}"

    @Slot(str)
    def selectTool(self, nav_id: str) -> None:
        if self._busy or nav_id == self._spec.nav_id:
            return
        variant = self._variants.get(nav_id, DEFAULT_VARIANTS.get(nav_id, "default"))
        try:
            spec = spec_for(nav_id, variant)
        except KeyError:
            return
        self._spec = spec
        self._ensure_state(spec)
        self._sync_input_model()
        self._log_model.clear()
        self._append_log(spec.log_text, "info")
        self.specChanged.emit()
        self.materialChanged.emit()
        self.supportChanged.emit()
        self._bump_form_revision()
        if self._workspace_scope == "tool":
            self.refreshWorkspace()

    @Slot(str)
    def selectVariant(self, variant: str) -> None:
        if self._busy or variant == self._spec.variant:
            return
        try:
            spec = spec_for(self._spec.nav_id, variant)
        except KeyError:
            return
        self._variants[self._spec.nav_id] = variant
        self._spec = spec
        self._ensure_state(spec)
        self._sync_input_model()
        self._log_model.clear()
        self._append_log(spec.log_text, "info")
        self.specChanged.emit()
        self.materialChanged.emit()
        self.supportChanged.emit()
        self._bump_form_revision()
        if self._workspace_scope == "tool":
            self.refreshWorkspace()

    @Slot(str, "QVariant")
    def setFieldValue(self, field_id: str, value: Any) -> None:
        state = self._form_states[self._state_key()]
        state[field_id] = value
        if field_id in {"rename_mode", "collect_all", "library_mode"}:
            if field_id == "library_mode" and value == "flat_ocr":
                state["use_ocr_cache"] = True
            self.specChanged.emit()
            self._bump_form_revision()

    @Slot(str, str)
    def applyDatePreset(self, group: str, preset: str) -> None:
        """Apply the same calendar ranges as the legacy UI."""

        if self._spec.tool_id != "data_statistics" or group not in {"week", "month"}:
            return
        state = self._form_states[self._state_key()]
        start_id = "week_start" if group == "week" else "month_start"
        end_id = "week_end" if group == "week" else "month_end"
        if preset == "clear":
            state[start_id] = ""
            state[end_id] = ""
            self._bump_form_revision()
            return
        today = date.today()
        if preset == "this_month":
            start = today.replace(day=1)
            end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
        elif preset == "last_month":
            end = today.replace(day=1) - timedelta(days=1)
            start = end.replace(day=1)
        elif group == "week" and preset in {"this_week", "last_week"}:
            monday = today - timedelta(days=today.weekday())
            if preset == "this_week":
                start = monday
                end = monday + timedelta(days=6)
            else:
                start = monday - timedelta(days=7)
                end = monday - timedelta(days=1)
        else:
            return
        state[start_id] = start.isoformat()
        state[end_id] = end.isoformat()
        self._bump_form_revision()

    @Slot(str, bool)
    def toggleMaterial(self, name: str, selected: bool) -> None:
        state = self._form_states[self._state_key()]
        values = list(state.get("material_types") or [])
        if selected and name not in values:
            values.append(name)
        elif not selected and name in values:
            values.remove(name)
        state["material_types"] = values

    @Slot()
    def selectAllMaterials(self) -> None:
        self._form_states[self._state_key()]["material_types"] = list(
            self._material_preferences.available_materials
        )
        self._bump_form_revision()

    @Slot()
    def clearMaterials(self) -> None:
        self._form_states[self._state_key()]["material_types"] = []
        self._bump_form_revision()

    def _selected_materials(self) -> list[str]:
        values = self._form_states.get(("material_collector", "default"), {})
        selected = set(values.get("material_types") or [])
        return [
            name
            for name in self._material_preferences.available_materials
            if name in selected
        ]

    def _material_mutated(self) -> None:
        state = self._form_states.get(("material_collector", "default"))
        if state is not None:
            available = set(self._material_preferences.available_materials)
            state["material_types"] = [
                name for name in state.get("material_types", []) if name in available
            ]
        self._save_workspace_preferences()
        self._bump_form_revision()
        self.materialChanged.emit()

    @Slot(str)
    def setMaterialPresetName(self, name: str) -> None:
        if name not in self._material_preferences.preset_names:
            return
        self._material_preset_name = name
        self.materialChanged.emit()

    @Slot(str)
    def applyMaterialPreset(self, name: str) -> None:
        materials = self._material_preferences.get_preset(name)
        if materials is None:
            self.notificationRequested.emit("预设不可用", "这个预设不存在或已经被删除。", "warning")
            return
        self._material_preset_name = name
        state = self._form_states[("material_collector", "default")]
        state["material_types"] = list(materials)
        state["collect_all"] = False
        self._bump_form_revision()
        self.materialChanged.emit()

    @Slot()
    def requestAddCustomMaterial(self) -> None:
        token = f"material-add:{time.monotonic_ns()}"
        self._pending_text_action = token
        self.textInputRequested.emit(
            "添加自定义材料",
            "输入材料名称（例如：户口本、体检报告）",
            "",
            token,
        )

    @Slot(str)
    def requestDeleteCustomMaterial(self, name: str) -> None:
        if name not in self._material_preferences.custom_materials:
            self.notificationRequested.emit("请选择材料", "请先选择要删除的自定义材料。", "warning")
            return
        referenced_by = [
            preset
            for preset, materials in self._material_preferences.custom_presets.items()
            if name in materials
        ]
        suffix = ""
        if referenced_by:
            suffix = "\n\n删除后会自动清理这些预设中的引用：" + "、".join(referenced_by)
        token = f"material-delete:{time.monotonic_ns()}"
        self._pending_confirmation = token
        self._pending_confirmation_action = ("material-delete", name)
        self.confirmationRequested.emit("确认删除材料", f"确定删除自定义材料“{name}”吗？{suffix}", token)

    @Slot()
    def requestCreateMaterialPreset(self) -> None:
        if not self._selected_materials():
            self.notificationRequested.emit("没有选择材料", "请先勾选至少一种材料，再保存为预设。", "warning")
            return
        token = f"preset-create:{time.monotonic_ns()}"
        self._pending_text_action = token
        self.textInputRequested.emit("保存自定义预设", "输入预设名称", "", token)

    @Slot(str)
    def updateMaterialPreset(self, name: str) -> None:
        try:
            saved = self._material_preferences.save_preset(
                name,
                self._selected_materials(),
                replacing=name,
            )
        except ValueError as exc:
            self.notificationRequested.emit("无法更新预设", str(exc), "warning")
            return
        self._material_preset_name = saved
        self._material_mutated()
        self.notificationRequested.emit("预设已更新", f"“{saved}”已按当前勾选更新。", "success")

    @Slot(str)
    def requestRenameMaterialPreset(self, name: str) -> None:
        if self._material_preferences.is_builtin_preset(name):
            self.notificationRequested.emit("内置预设不能重命名", "内置预设会一直保留；自定义预设可以重命名。", "info")
            return
        token = f"preset-rename:{time.monotonic_ns()}"
        self._pending_text_action = token
        self._pending_text_payload = name
        self.textInputRequested.emit("重命名预设", "输入新的预设名称", name, token)

    @Slot(str)
    def requestDeleteMaterialPreset(self, name: str) -> None:
        if self._material_preferences.is_builtin_preset(name):
            self.notificationRequested.emit("内置预设不能删除", "内置预设会一直保留；自定义预设可以删除。", "info")
            return
        token = f"preset-delete:{time.monotonic_ns()}"
        self._pending_confirmation = token
        self._pending_confirmation_action = ("preset-delete", name)
        self.confirmationRequested.emit("确认删除预设", f"确定删除自定义预设“{name}”吗？\n\n材料本身不会被删除。", token)

    @Slot(str, str)
    def submitTextAction(self, token: str, value: str) -> None:
        if token != self._pending_text_action:
            return
        self._pending_text_action = None
        try:
            if token.startswith("material-add:"):
                material = self._material_preferences.add_material(value)
                state = self._form_states[("material_collector", "default")]
                state.setdefault("material_types", []).append(material)
                self._material_mutated()
                self.notificationRequested.emit("材料已添加", f"已添加自定义材料“{material}”。", "success")
            elif token.startswith("preset-create:"):
                saved = self._material_preferences.save_preset(value, self._selected_materials())
                self._material_preset_name = saved
                self._material_mutated()
                self.notificationRequested.emit("预设已保存", f"已保存“{saved}”。", "success")
            elif token.startswith("preset-rename:"):
                current = str(getattr(self, "_pending_text_payload", ""))
                saved = self._material_preferences.rename_preset(current, value)
                self._material_preset_name = saved
                self._material_mutated()
        except ValueError as exc:
            self.notificationRequested.emit("无法保存", str(exc), "warning")

    def _bump_form_revision(self) -> None:
        self._form_revision += 1
        self.formRevisionChanged.emit()

    def _dialog_parent(self):
        from .compat import QApplication

        return QApplication.activeWindow()

    @Slot()
    def chooseInputFiles(self) -> None:
        if self._busy or not self.inputAllowsFiles:
            return
        parent = self._dialog_parent()
        file_filter = "Excel 或压缩包 (*.xlsx *.xls *.zip *.rar *.7z *.tar *.gz *.tgz);;所有文件 (*)"
        if self._spec.input_mode == "excel_single":
            filename, _selected = QFileDialog.getOpenFileName(parent, self._spec.input_drop_title, "", "Excel 工作簿 (*.xlsx *.xls);;所有文件 (*)")
            paths = [Path(filename)] if filename else []
        else:
            filenames, _selected = QFileDialog.getOpenFileNames(parent, self._spec.input_drop_title, "", file_filter)
            paths = [Path(filename) for filename in filenames]
        if paths:
            self._set_inputs(paths, replace=self._spec.tool_id == "data_statistics" or not self.inputAllowsMultiple)

    @Slot()
    def chooseInputFolder(self) -> None:
        if self._busy or not self.inputAllowsFolder:
            return
        selected = QFileDialog.getExistingDirectory(
            self._dialog_parent(), self._spec.input_drop_title, ""
        )
        if selected:
            self._set_inputs(
                [Path(selected)],
                replace=self._spec.tool_id == "data_statistics" or not self.inputAllowsMultiple,
            )

    def _set_inputs(self, paths: list[Path], *, replace: bool) -> None:
        key = self._state_key()
        current = [] if replace else list(self._input_states[key])
        for path in paths:
            if path not in current:
                current.append(path)
        self._input_states[key] = current
        self._sync_input_model()

    def _sync_input_model(self) -> None:
        items = []
        for path in self._input_states[self._state_key()]:
            is_dir = path.is_dir()
            detail = "文件夹" if is_dir else path.suffix.lower().lstrip(".").upper() or "文件"
            items.append(
                {
                    "name": path.name or str(path),
                    "path": str(path),
                    "kind": "folder" if is_dir else "file",
                    "detail": detail,
                }
            )
        self._input_model.set_items(items)

    @Slot(int)
    def removeInput(self, index: int) -> None:
        if self._busy:
            return
        values = self._input_states[self._state_key()]
        if 0 <= index < len(values):
            del values[index]
            self._sync_input_model()

    @Slot()
    def clearInputs(self) -> None:
        if self._busy:
            return
        self._input_states[self._state_key()].clear()
        self._sync_input_model()

    @Slot()
    def chooseSupportFile(self) -> None:
        if self._busy or not self._spec.support_id:
            return
        file_filter = "Excel 工作簿 (*.xlsx *.xls);;所有文件 (*)"
        if self._spec.support_mode == "excel_archive_or_folder":
            file_filter = "Excel 或压缩包 (*.xlsx *.xls *.zip *.rar *.7z *.tar *.gz *.tgz);;所有文件 (*)"
        filename, _selected = QFileDialog.getOpenFileName(
            self._dialog_parent(), self._spec.support_label, "", file_filter
        )
        if filename:
            self._support_states[self._state_key()] = filename
            self.supportChanged.emit()

    @Slot()
    def chooseSupportFolder(self) -> None:
        if self._busy or not self.supportAllowsFolder:
            return
        selected = QFileDialog.getExistingDirectory(
            self._dialog_parent(), self._spec.support_label, ""
        )
        if selected:
            self._support_states[self._state_key()] = selected
            self.supportChanged.emit()

    @Slot()
    def clearSupport(self) -> None:
        if self._busy:
            return
        self._support_states[self._state_key()] = ""
        self.supportChanged.emit()

    @Slot(result=str)
    def chooseProjectParent(self) -> str:
        selected = QFileDialog.getExistingDirectory(
            self._dialog_parent(), "选择项目保存位置", self.defaultProjectParent
        )
        return str(selected or "")

    @Slot()
    def requestCreateProject(self) -> None:
        if self._busy or self._workspace_busy:
            return
        self.projectCreationRequested.emit(self.defaultProjectName, self.defaultProjectParent)

    @Slot(str, str)
    def createProject(self, name: str, parent: str) -> None:
        if self._busy or self._workspace_busy or self._project_opening:
            return
        target, error = workspace_project_creation_target(parent, name)
        if error or target is None:
            self.notificationRequested.emit("无法创建项目", error or "项目位置无效。", "error")
            return
        self._project_generation += 1
        generation = self._project_generation
        self._project_opening = True

        def worker() -> None:
            try:
                store = ProjectStore.create(target, str(name).strip())
            except Exception as exc:
                self._projectOpenFailed.emit(generation, workspace_project_create_error_message(exc))
                return
            self._projectOpened.emit(generation, store, str(target))

        threading.Thread(target=worker, daemon=True, name="HRToolkit-create-project").start()

    @Slot()
    def openProjectDialog(self) -> None:
        if self._busy or self._workspace_busy or self._project_opening:
            return
        selected = QFileDialog.getExistingDirectory(
            self._dialog_parent(), "打开工作项目", self.defaultProjectParent
        )
        if selected:
            self.openProject(selected)

    @Slot(str)
    def openProject(self, path: str) -> None:
        if self._busy or self._workspace_busy or self._project_opening:
            return
        self._project_generation += 1
        generation = self._project_generation
        self._project_opening = True

        def worker() -> None:
            try:
                store = ProjectStore.open(Path(path), writable=True, read_only_fallback=True)
            except Exception as exc:
                self._projectOpenFailed.emit(generation, str(exc))
                return
            self._projectOpened.emit(generation, store, str(path))

        threading.Thread(target=worker, daemon=True, name="HRToolkit-open-project").start()

    @Slot(int, object, str)
    def _apply_project_open(self, generation: int, store: ProjectStore, path: str) -> None:
        if generation != self._project_generation or self._closed:
            store.close()
            return
        previous = self._project_store
        self._project_store = store
        self._project_path = Path(path)
        self._project_opening = False
        if previous is not None:
            try:
                previous.close()
            except Exception as exc:
                runlog.log_exception("关闭旧工作项目失败", exc)
        self._remember_project(self._project_path)
        self._save_workspace_preferences()
        self.projectChanged.emit()
        self.refreshWorkspace()
        if not store.writable:
            reason = store.workspace.read_only_reason or "项目当前为只读状态。"
            self.notificationRequested.emit("项目以只读方式打开", reason, "warning")

    @Slot(int, str)
    def _apply_project_error(self, generation: int, message: str) -> None:
        if generation != self._project_generation:
            return
        self._project_opening = False
        self.notificationRequested.emit("无法打开项目", message, "error")

    def _remember_project(self, path: Path) -> None:
        self._recent_projects = [path, *(item for item in self._recent_projects if item != path)][:8]

    @staticmethod
    def _settings_path() -> Path:
        if sys.platform.startswith("win"):
            base = Path(os.environ.get("LOCALAPPDATA", "").strip() or (Path.home() / "AppData" / "Local"))
        elif sys.platform == "darwin":
            base = Path.home() / "Library" / "Application Support"
        else:
            base = Path(os.environ.get("XDG_CONFIG_HOME", "").strip() or (Path.home() / ".config"))
        return base / "HRToolkit" / "workspace-ui.json"

    @Slot()
    def start(self) -> None:
        state: dict[str, Any] = {}
        path = self._settings_path()
        try:
            if path.is_file():
                payload = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(payload, dict):
                    state = payload
        except Exception as exc:
            runlog.log_exception("读取项目界面设置失败", exc)
        recent = []
        for value in state.get("recent_projects", []):
            candidate = Path(str(value)).expanduser()
            if candidate.is_dir() and candidate not in recent:
                recent.append(candidate)
        self._recent_projects = recent[:8]
        self._material_preferences = MaterialPreferences.from_payload(
            state.get("material_preferences")
        )
        preset_names = self._material_preferences.preset_names
        self._material_preset_name = preset_names[0] if preset_names else ""
        material_key = ("material_collector", "default")
        if material_key in self._form_states:
            selected = list(self._form_states[material_key].get("material_types") or [])
            if not selected or selected == list(MaterialPreferences().available_materials):
                self._form_states[material_key]["material_types"] = list(
                    self._material_preferences.available_materials
                )
        self._bump_form_revision()
        self.materialChanged.emit()
        self.projectChanged.emit()
        current = state.get("current_project")
        if current:
            self.openProject(str(current))
        threading.Thread(
            target=cleanup_stale_update_files,
            daemon=True,
            name="HRToolkit-update-cleanup",
        ).start()
        if update_check_enabled():
            QTimer.singleShot(600, self.requestStartupUpdateCheck)

    def _save_workspace_preferences(self) -> None:
        path = self._settings_path()
        payload: dict[str, Any] = {}
        try:
            if path.is_file():
                existing = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(existing, dict):
                    payload.update(existing)
        except Exception:
            pass
        payload.update(
            {
                "version": max(2, int(payload.get("version", 0) or 0)),
                "current_project": None if self._project_path is None else str(self._project_path),
                "recent_projects": [str(item) for item in self._recent_projects[:8]],
                "material_preferences": self._material_preferences.to_payload(),
            }
        )
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            temp = path.with_suffix(".tmp")
            temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(temp, path)
        except Exception as exc:
            runlog.log_exception("保存项目界面设置失败", exc)

    @Slot()
    def openProjectFolder(self) -> None:
        if self._project_path is not None:
            open_path(self._project_path)

    @Slot(bool)
    def setWorkspaceExpanded(self, expanded: bool) -> None:
        if self._workspace_expanded == bool(expanded):
            return
        self._workspace_expanded = bool(expanded)
        self.workspaceChanged.emit()
        if expanded:
            self.refreshWorkspace()

    @Slot(str)
    def setWorkspaceScope(self, scope: str) -> None:
        if scope not in {"all", "tool"} or scope == self._workspace_scope:
            return
        self._workspace_scope = scope
        self.workspaceChanged.emit()
        self.refreshWorkspace()

    @Slot(str)
    def setWorkspaceSearch(self, query: str) -> None:
        self._workspace_search = str(query or "").strip()
        self._search_timer.start()

    def _workspace_root(self) -> Path | None:
        if self._project_path is None:
            return None
        if self._workspace_scope == "all":
            return self._project_path
        return self._project_path / self._spec.group / self._project_tool_name()

    def _project_tool_name(self) -> str:
        names = {
            "social_security": "社保明细与汇总",
            "insurance_ledger": "保险台账与预警",
            "data_statistics": "考勤与周月报",
            "salary_split": "工资表拆分",
            "salary_merge": "多月工资合并",
            "personnel_change_merge": "异动汇总",
            "roster_update": "花名册更新",
            "archive_import": "档案入库",
            "archive_export": "档案表生成",
            "material_collector": "员工资料打包",
            "folder_rename": "资料文件夹改名",
        }
        return names[self._spec.tool_id]

    @staticmethod
    def _hide_workspace_path(path: Path) -> bool:
        name = path.name
        lower = name.casefold()
        return (
            name in WORKSPACE_HIDDEN_NAMES
            or name.startswith(".")
            or name.startswith("~$")
            or any(lower.endswith(suffix) for suffix in WORKSPACE_HIDDEN_SUFFIXES)
            or path.is_symlink()
        )

    @classmethod
    def _scan_directory(cls, root: Path, *, depth: int = 0) -> list[dict[str, Any]]:
        records = []
        try:
            with os.scandir(root) as entries:
                for entry in entries:
                    path = Path(entry.path)
                    if cls._hide_workspace_path(path):
                        continue
                    try:
                        is_dir = entry.is_dir(follow_symlinks=False)
                    except OSError:
                        is_dir = False
                    records.append(
                        {
                            "name": entry.name,
                            "path": str(path),
                            "isDir": is_dir,
                            "depth": depth,
                            "expanded": False,
                            "hasChildren": is_dir,
                            "detail": "文件夹" if is_dir else (path.suffix.lower().lstrip(".").upper() or "文件"),
                        }
                    )
        except OSError:
            return []
        records.sort(key=lambda item: (not bool(item["isDir"]), str(item["name"]).casefold()))
        return records

    @classmethod
    def _search_workspace(
        cls,
        root: Path,
        query: str,
        *,
        cancelled=None,
    ) -> list[dict[str, Any]]:
        results = []
        try:
            for current_root, dir_names, file_names in os.walk(root, followlinks=False):
                if cancelled is not None and cancelled():
                    return results
                current = Path(current_root)
                dir_names[:] = [
                    name for name in dir_names if not cls._hide_workspace_path(current / name)
                ]
                for name in (*dir_names, *file_names):
                    if cancelled is not None and cancelled():
                        return results
                    path = current / name
                    if cls._hide_workspace_path(path) or query not in name.casefold():
                        continue
                    try:
                        is_dir = path.is_dir()
                    except OSError:
                        is_dir = False
                    try:
                        relative_parent = path.parent.relative_to(root).as_posix()
                    except ValueError:
                        relative_parent = path.parent.name
                    results.append(
                        {
                            "name": name,
                            "path": str(path),
                            "isDir": is_dir,
                            "depth": 0,
                            "expanded": False,
                            "hasChildren": is_dir,
                            "detail": relative_parent if relative_parent != "." else ("文件夹" if is_dir else "文件"),
                        }
                    )
                    if len(results) >= WORKSPACE_SEARCH_LIMIT:
                        return results
        except OSError:
            return results
        return results

    @Slot()
    def refreshWorkspace(self) -> None:
        previous_cancel = self._workspace_scan_cancel_event
        if previous_cancel is not None:
            previous_cancel.set()
        cancel_event = threading.Event()
        self._workspace_scan_cancel_event = cancel_event
        root = self._workspace_root()
        self._workspace_generation += 1
        self._workspace_child_loads.clear()
        generation = self._workspace_generation
        if root is None or not root.is_dir():
            self._workspace_items = []
            self._workspace_selected_path = None
            self._workspace_selected_item = None
            self._workspace_model.clear()
            self.workspaceSelectionChanged.emit()
            return
        query = self._workspace_search.casefold()

        def worker() -> None:
            items = (
                self._search_workspace(root, query, cancelled=cancel_event.is_set)
                if query
                else self._scan_directory(root)
            )
            if cancel_event.is_set():
                return
            self._workspaceItemsReady.emit(generation, items)

        threading.Thread(target=worker, daemon=True, name="HRToolkit-workspace-scan").start()

    @Slot(int, object)
    def _apply_workspace_items(self, generation: int, items: list[dict[str, Any]]) -> None:
        if generation != self._workspace_generation or self._closed:
            return
        self._workspace_items = list(items)
        self._workspace_model.set_items(self._workspace_items)
        selected_text = str(self._workspace_selected_path or "")
        self._workspace_selected_item = next(
            (
                dict(item)
                for item in self._workspace_items
                if str(item.get("path") or "") == selected_text
            ),
            None,
        )
        self.workspaceSelectionChanged.emit()

    @Slot(int)
    def toggleWorkspaceRow(self, row: int) -> None:
        if self._workspace_search or row < 0 or row >= len(self._workspace_items):
            return
        item = self._workspace_items[row]
        if not item.get("isDir"):
            return
        depth = int(item.get("depth", 0))
        if item.get("expanded"):
            end = row + 1
            while end < len(self._workspace_items) and int(self._workspace_items[end].get("depth", 0)) > depth:
                end += 1
            item["expanded"] = False
            self._workspace_model.update_at(row, item)
            remove_count = end - row - 1
            del self._workspace_items[row + 1 : end]
            self._workspace_model.splice(row + 1, remove_count)
            return
        item["expanded"] = True
        self._workspace_model.update_at(row, item)
        generation = self._workspace_generation
        path = Path(str(item["path"]))
        path_text = str(path)
        if path_text in self._workspace_child_loads:
            return
        self._workspace_child_loads.add(path_text)

        def worker() -> None:
            children = self._scan_directory(path, depth=depth + 1)
            self._workspaceChildrenReady.emit(generation, row, path_text, depth, children)

        threading.Thread(target=worker, daemon=True, name="HRToolkit-workspace-children").start()

    @Slot(int, int, str, int, object)
    def _apply_workspace_children(self, generation: int, row: int, path: str, depth: int, children) -> None:
        self._workspace_child_loads.discard(path)
        if generation != self._workspace_generation or row >= len(self._workspace_items):
            return
        item = self._workspace_items[row]
        if str(item.get("path")) != path or not item.get("expanded") or int(item.get("depth", 0)) != depth:
            return
        inserted = list(children)
        self._workspace_items[row + 1 : row + 1] = inserted
        self._workspace_model.splice(row + 1, items=inserted)

    @Slot(int)
    def selectWorkspaceRow(self, row: int) -> None:
        if 0 <= row < len(self._workspace_items):
            item = self._workspace_items[row]
            self._workspace_selected_path = Path(str(item["path"]))
            self._workspace_selected_item = dict(item)
        else:
            self._workspace_selected_path = None
            self._workspace_selected_item = None
        self.workspaceSelectionChanged.emit()

    @Slot(int)
    def openWorkspaceRow(self, row: int) -> None:
        if 0 <= row < len(self._workspace_items):
            path = Path(str(self._workspace_items[row]["path"]))
            if path.is_dir():
                self.toggleWorkspaceRow(row)
            elif path.exists():
                open_path(path)

    @Slot()
    def launchWorkspaceSelection(self) -> None:
        path = self._workspace_selected_path
        if path is not None and path.exists():
            open_path(path)

    @Slot()
    def revealWorkspaceSelection(self) -> None:
        path = self._workspace_selected_path
        if path is None:
            return
        target = path if path.is_dir() else path.parent
        if target.exists():
            open_path(target)

    @Slot(int)
    def revealWorkspaceRow(self, row: int) -> None:
        if 0 <= row < len(self._workspace_items):
            path = Path(str(self._workspace_items[row]["path"]))
            target = path if path.is_dir() else path.parent
            if target.exists():
                open_path(target)

    @Slot()
    def importWorkspaceFiles(self) -> None:
        if not self.projectWritable or self._busy or self._workspace_busy:
            return
        names, _selected = QFileDialog.getOpenFileNames(
            self._dialog_parent(), "选择要导入项目的文件", "", "所有文件 (*)"
        )
        if names:
            self._start_workspace_import([Path(name) for name in names])

    @Slot()
    def importWorkspaceFolder(self) -> None:
        if not self.projectWritable or self._busy or self._workspace_busy:
            return
        selected = QFileDialog.getExistingDirectory(
            self._dialog_parent(), "选择要导入项目的文件夹", ""
        )
        if selected:
            self._start_workspace_import([Path(selected)])

    def _workspace_import_target(self) -> tuple[Path, tuple[str, str] | None] | None:
        store = self._project_store
        if store is None:
            return None
        common = store.workspace.common_root
        selected = self._workspace_selected_path
        target = selected if selected is not None and selected.is_dir() else common
        try:
            locations = store.list_batch_locations()
        except Exception:
            locations = ()
        for summary, directories in locations:
            for category in ("uploads", "supplements"):
                directory = directories.get(category)
                if directory is None:
                    continue
                try:
                    target.relative_to(directory)
                except ValueError:
                    continue
                if category == "uploads" and summary.status != "draft":
                    target = directories["supplements"]
                    category = "supplements"
                return target, (summary.id, category)
            result_dir = directories.get("results")
            if result_dir is not None:
                try:
                    target.relative_to(result_dir)
                except ValueError:
                    pass
                else:
                    return directories["supplements"], (summary.id, "supplements")
        try:
            target.relative_to(common)
        except ValueError:
            target = common
        return target, None

    def _start_workspace_import(self, sources: list[Path]) -> None:
        resolved = self._workspace_import_target()
        store = self._project_store
        if resolved is None or store is None:
            return
        target, batch = resolved
        self._workspace_busy = True
        self._workspace_cancel_event = threading.Event()
        self.workspaceBusyChanged.emit()
        self._append_log(f"正在把 {len(sources)} 项资料安全保存到项目…", "info")

        def worker() -> None:
            try:
                if batch is None:
                    store.import_to_directory(
                        target,
                        sources,
                        cancelled=self._workspace_cancel_event.is_set,
                    )
                else:
                    batch_id, category = batch
                    store.import_sources(
                        batch_id,
                        sources,
                        category=category,
                        role="workspace",
                        cancelled=self._workspace_cancel_event.is_set,
                    )
            except ImportCancelled:
                self._workspaceImportFinished.emit(False, "资料导入已取消。")
            except Exception as exc:
                self._workspaceImportFinished.emit(False, f"资料没有导入：{exc}")
            else:
                self._workspaceImportFinished.emit(True, "资料已安全保存到当前项目。")

        threading.Thread(target=worker, daemon=True, name="HRToolkit-workspace-import").start()

    @Slot()
    def cancelWorkspaceImport(self) -> None:
        if self._workspace_cancel_event is not None:
            self._workspace_cancel_event.set()

    @Slot(bool, str)
    def _apply_workspace_import_result(self, success: bool, message: str) -> None:
        self._workspace_busy = False
        self._workspace_cancel_event = None
        self.workspaceBusyChanged.emit()
        self._append_log(message, "success" if success else "warning")
        self.notificationRequested.emit("项目文件" if success else "导入未完成", message, "info" if success else "warning")
        self.refreshWorkspace()

    @staticmethod
    def _format_size(value: int) -> str:
        size = max(0, int(value))
        if size >= 1024**3:
            return f"{size / 1024**3:.1f} GB"
        if size >= 1024**2:
            return f"{size / 1024**2:.1f} MB"
        if size >= 1024:
            return f"{size / 1024:.1f} KB"
        return f"{size} B"

    @staticmethod
    def _format_time(value: str, short: bool = False) -> str:
        try:
            pattern = "%m-%d %H:%M" if short else "%Y-%m-%d %H:%M"
            return datetime.fromisoformat(value).astimezone().strftime(pattern)
        except (TypeError, ValueError):
            return str(value or "")

    @staticmethod
    def _names_text(names, empty: str, *, full: bool = False) -> str:
        values = tuple(str(item) for item in names if str(item))
        if not values:
            return empty
        if full:
            visible = "、".join(values[:5])
            return visible if len(values) <= 5 else f"{visible} 等 {len(values)} 个文件"
        return values[0] if len(values) == 1 else f"{values[0]} 等 {len(values)} 个"

    def _history_started_after(self) -> str | None:
        today = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
        if self._history_date_filter == "今天":
            start = today
        elif self._history_date_filter == "最近7天":
            start = today - timedelta(days=6)
        elif self._history_date_filter == "最近30天":
            start = today - timedelta(days=29)
        elif self._history_date_filter == "今年":
            start = today.replace(month=1, day=1)
        else:
            return None
        return start.astimezone(timezone.utc).isoformat(timespec="seconds")

    @Slot()
    def requestHistory(self) -> None:
        self.refreshHistory(self._history_search, self._history_tool_id, self._history_date_filter)

    @Slot(str, str, str)
    def refreshHistory(self, search: str, tool_id: str, date_filter: str) -> None:
        self._history_search = str(search or "").strip()
        self._history_tool_id = str(tool_id or "").strip()
        self._history_date_filter = str(date_filter or "全部时间")
        self._history_generation += 1
        generation = self._history_generation
        self._history_busy = True
        self._history_message = "正在读取旧版记录…"
        self.historyChanged.emit()

        def worker() -> None:
            try:
                store = self._history_store
                if store is None:
                    store = HistoryStore()
                    self._history_store = store
                    self._history_init_attempted = True
                tasks, total = store.list_tasks(
                    search=self._history_search,
                    tool_id=self._history_tool_id or None,
                    started_after=self._history_started_after(),
                    limit=HISTORY_PAGE_SIZE,
                    offset=self._history_page * HISTORY_PAGE_SIZE,
                )
                rows = [
                    {
                        "recordId": task.id,
                        "time": self._format_time(task.started_at, short=True),
                        "tool": task.tool_name,
                        "status": HISTORY_STATUS_LABELS.get(task.status, task.status),
                        "inputs": self._names_text(task.input_names, "未归档上传资料"),
                        "outputs": self._names_text(task.output_names, "暂无完整结果"),
                        "detail": task.error_message or "",
                    }
                    for task in tasks
                ]
                message = f"共找到 {total} 次处理记录"
                try:
                    stats = store.storage_stats()
                    message += f" · 已留存 {self._format_size(stats['total_bytes'])} · 磁盘可用 {self._format_size(stats['free_bytes'])}"
                    if stats["trash_bytes"]:
                        message += f"（回收站 {self._format_size(stats['trash_bytes'])}）"
                except Exception:
                    pass
                if not rows:
                    message = (
                        "没有找到相关记录，可以清除查找内容或更换筛选条件。"
                        if self._history_search or self._history_tool_id or self._history_date_filter != "全部时间"
                        else "没有找到旧版记录。新处理的资料和结果请在“项目文件”中查看。"
                    )
            except Exception as exc:
                self._historyListReady.emit(generation, [], 0, f"历史记录暂时无法读取：{exc}")
                return
            self._historyListReady.emit(generation, rows, total, message)

        threading.Thread(target=worker, daemon=True, name="HRToolkit-history-list").start()

    @Slot(int, object, int, str)
    def _apply_history_list(self, generation: int, rows, total: int, message: str) -> None:
        if generation != self._history_generation or self._closed:
            return
        self._history_busy = False
        self._history_total = max(0, int(total))
        pages = max(1, (self._history_total + HISTORY_PAGE_SIZE - 1) // HISTORY_PAGE_SIZE)
        if self._history_page >= pages:
            self._history_page = pages - 1
        self._history_message = message
        self._history_model.set_items(rows)
        self._history_selected = None
        self._history_detail = {}
        self.historyChanged.emit()
        if rows:
            self.selectHistoryRow(0)

    @Slot(int)
    def changeHistoryPage(self, delta: int) -> None:
        target = self._history_page + int(delta)
        if target < 0 or (target * HISTORY_PAGE_SIZE >= self._history_total and delta > 0):
            return
        self._history_page = target
        self.requestHistory()

    @Slot(int)
    def selectHistoryRow(self, row: int) -> None:
        item = self._history_model.item_at(row)
        store = self._history_store
        if item is None or store is None:
            return
        task_id = str(item["recordId"])
        generation = self._history_generation
        self._history_detail = {"id": task_id, "title": "正在读取详情…", "body": ""}
        self.historyChanged.emit()

        def worker() -> None:
            try:
                detail = store.get_task(task_id)
                if detail is None:
                    raise RuntimeError("历史记录不存在。")
                summary = detail.summary
                status = HISTORY_STATUS_LABELS.get(summary.status, summary.status)
                lines = [
                    f"处理时间：{self._format_time(summary.started_at)}",
                    f"上传资料：{self._names_text(summary.input_names, '未归档上传资料', full=True)}",
                    f"处理结果：{self._names_text(summary.output_names, '未生成完整结果', full=True)}",
                ]
                if summary.status == "failed":
                    lines.append("说明：上传资料已保存，但本次没有正常生成完整结果。")
                elif summary.status == "stopped":
                    lines.append("说明：这次处理没有正常完成，可以再次使用已保存的资料。")
                if summary.error_message and summary.status in {"failed", "stopped"}:
                    lines.append(f"原因：{summary.error_message}")
                payload = {
                    "batchId": summary.id,
                    "title": f"{summary.tool_name} · {status}",
                    "body": "\n".join(lines),
                    "canOpenOutput": bool(detail.outputs),
                    "canOpenInput": bool(detail.inputs),
                    "canReuse": summary.tool_id != "folder_rename" and bool(detail.inputs),
                    "canDelete": summary.status != "running",
                }
            except Exception as exc:
                self._historyDetailReady.emit(generation, None, str(exc))
                return
            self._historyDetailReady.emit(generation, (detail, payload), "")

        threading.Thread(target=worker, daemon=True, name="HRToolkit-history-detail").start()

    @Slot(int, object, str)
    def _apply_history_detail(self, generation: int, result, error: str) -> None:
        if generation != self._history_generation or self._closed:
            return
        if error or not result:
            self._history_selected = None
            self._history_detail = {"title": "这条记录暂时无法读取", "body": error}
        else:
            detail, payload = result
            self._history_selected = detail
            self._history_detail = dict(payload)
        self.historyChanged.emit()

    @Slot()
    def openHistoryInput(self) -> None:
        detail = self._history_selected
        if detail is not None and detail.inputs and detail.input_dir.is_dir():
            open_path(detail.input_dir)

    @Slot()
    def openHistoryOutput(self) -> None:
        detail = self._history_selected
        if detail is not None and detail.outputs and detail.output_dir.is_dir():
            open_path(detail.output_dir)

    @Slot()
    def openHistoryRoot(self) -> None:
        if self._history_store is not None:
            open_path(self._history_store.records_dir)

    @Slot()
    def openHistoryTrash(self) -> None:
        if self._history_store is not None:
            open_path(self._history_store.trash_dir)

    @Slot()
    def reuseHistory(self) -> None:
        detail = self._history_selected
        if detail is None or self._busy:
            return
        tool_id = detail.summary.tool_id
        nav_id = {"roster_update": "personnel_change_merge", "archive_export": "archive_import"}.get(tool_id, tool_id)
        try:
            spec_for(nav_id, DEFAULT_VARIANTS.get(nav_id, "default"))
        except KeyError:
            self.notificationRequested.emit("暂不支持", "这条旧记录暂时不能直接再次使用，可以先打开上传资料。", "warning")
            return
        if tool_id == "folder_rename":
            self.notificationRequested.emit("请先复制资料", "为了保护历史原件，文件夹改名记录不能直接再次处理。", "info")
            return
        variant = DEFAULT_VARIANTS.get(nav_id, "default")
        if nav_id == "personnel_change_merge" and (tool_id == "roster_update" or detail.summary.mode == "roster"):
            variant = "roster"
        elif nav_id == "archive_import" and (tool_id == "archive_export" or detail.summary.mode == "export"):
            variant = "export"
        self._variants[nav_id] = variant
        self._spec = spec_for(nav_id, variant)
        self._ensure_state(self._spec)
        main_inputs = [
            item.archived_path
            for item in detail.inputs
            if item.role in {"input_path", "input_paths"} and item.archived_path.exists()
        ]
        secondary: dict[str, list[Path]] = {}
        for item in detail.inputs:
            if item.role not in {"input_path", "input_paths"} and item.archived_path.exists():
                secondary.setdefault(item.role, []).append(item.archived_path)
        if self._spec.input_mode == "excel_single":
            main_inputs = main_inputs[:1]
        self._input_states[self._state_key()] = main_inputs
        self._support_states[self._state_key()] = ""
        if secondary:
            paths = next(iter(secondary.values()))
            try:
                support = paths[0] if len(paths) == 1 else Path(os.path.commonpath([str(path) for path in paths]))
            except ValueError:
                support = paths[0]
            self._support_states[self._state_key()] = str(support)
        self._sync_input_model()
        self._log_model.clear()
        self._append_log(self._spec.log_text, "info")
        self.specChanged.emit()
        self.supportChanged.emit()
        self.materialChanged.emit()
        self._bump_form_revision()
        self.notificationRequested.emit("资料已带入", "以前保存的资料已经放回当前功能，请确认后重新处理。", "success")

    @Slot()
    def requestMoveHistoryToTrash(self) -> None:
        detail = self._history_selected
        if detail is None or detail.summary.status == "running":
            return
        token = f"history-trash:{time.monotonic_ns()}"
        self._pending_confirmation = token
        self._pending_confirmation_action = ("history-trash", detail.summary.id)
        self.confirmationRequested.emit(
            "移到回收站",
            "这次处理的上传资料和结果会移到 HRToolkit 回收站，不会立即永久删除。是否继续？",
            token,
        )

    def _move_history_to_trash(self, task_id: str) -> None:
        store = self._history_store
        if store is None:
            return
        self._history_busy = True
        self.historyChanged.emit()

        def worker() -> None:
            try:
                store.move_to_trash(task_id)
            except Exception as exc:
                self._historyActionFinished.emit("trash", False, str(exc))
            else:
                self._historyActionFinished.emit("trash", True, "已移到旧版记录回收站。")

        threading.Thread(target=worker, daemon=True, name="HRToolkit-history-trash").start()

    @Slot()
    def rebuildHistoryIndex(self) -> None:
        store = self._history_store
        if store is None or self._history_busy:
            return
        self._history_busy = True
        self._history_message = "正在整理历史记录，请稍候…"
        self.historyChanged.emit()

        def worker() -> None:
            try:
                count = store.rebuild_index_from_manifests()
            except Exception as exc:
                self._historyActionFinished.emit("rebuild", False, str(exc))
            else:
                self._historyActionFinished.emit("rebuild", True, f"历史记录已经整理完成，恢复或修复了 {count} 条记录。")

        threading.Thread(target=worker, daemon=True, name="HRToolkit-history-rebuild").start()

    @Slot(str, bool, str)
    def _apply_history_action(self, kind: str, success: bool, message: str) -> None:
        self._history_busy = False
        self.historyChanged.emit()
        self.notificationRequested.emit("整理完成" if success else "操作未完成", message, "success" if success else "error")
        self.requestHistory()

    def _filtered_trash_rows(self) -> list[dict[str, Any]]:
        query = self._trash_search.casefold()
        rows: list[dict[str, Any]] = []
        for detail in self._trash_items:
            summary = detail.summary
            title = summary.business_description or summary.directory_name or summary.tool_name
            searchable = " ".join((title, summary.tool_name, summary.group_name, summary.business_period)).casefold()
            if query and query not in searchable:
                continue
            rows.append(
                {
                    "batchId": summary.id,
                    "title": title,
                    "tool": f"{summary.group_name} · {summary.tool_name}",
                    "status": HISTORY_STATUS_LABELS.get(summary.status, summary.status),
                    "deletedAt": self._format_time(summary.deleted_at or ""),
                    "counts": f"上传 {detail.upload_count} · 结果 {detail.result_count} · 补充 {detail.supplement_count}",
                    "restorePath": detail.original_relative_path,
                    "size": self._format_size(detail.total_size_bytes),
                }
            )
        return rows

    @Slot()
    def requestProjectTrash(self) -> None:
        store = self._project_store
        if store is None:
            self.notificationRequested.emit("请先打开项目", "打开工作项目后才能查看项目回收站。", "warning")
            return
        self._trash_generation += 1
        generation = self._trash_generation
        self._trash_busy = True
        self.trashChanged.emit()

        def worker() -> None:
            try:
                items = list(store.list_trash_details())
            except Exception as exc:
                self._trashReady.emit(generation, [], str(exc))
            else:
                self._trashReady.emit(generation, items, "")

        threading.Thread(target=worker, daemon=True, name="HRToolkit-project-trash-list").start()

    @Slot(int, object, str)
    def _apply_trash_list(self, generation: int, items, error: str) -> None:
        if generation != self._trash_generation or self._closed:
            return
        self._trash_busy = False
        if error:
            self._trash_items = []
            self._trash_model.clear()
            self._trash_selected_id = ""
            self.notificationRequested.emit("回收站暂时无法读取", error, "error")
        else:
            self._trash_items = list(items)
            rows = self._filtered_trash_rows()
            self._trash_model.set_items(rows)
            ids = {str(item["batchId"]) for item in rows}
            if self._trash_selected_id not in ids:
                self._trash_selected_id = str(rows[0]["batchId"]) if rows else ""
        self.trashChanged.emit()

    @Slot(str)
    def setTrashSearch(self, query: str) -> None:
        self._trash_search = str(query or "").strip()
        rows = self._filtered_trash_rows()
        self._trash_model.set_items(rows)
        ids = {str(item["batchId"]) for item in rows}
        if self._trash_selected_id not in ids:
            self._trash_selected_id = str(rows[0]["batchId"]) if rows else ""
        self.trashChanged.emit()

    @Slot(int)
    def selectTrashRow(self, row: int) -> None:
        item = self._trash_model.item_at(row)
        self._trash_selected_id = "" if item is None else str(item["batchId"])
        self.trashChanged.emit()

    @Slot()
    def restoreSelectedTrash(self) -> None:
        store = self._project_store
        batch_id = self._trash_selected_id
        if store is None or not batch_id or self._trash_busy or not store.writable or self._busy or self._workspace_busy:
            return
        self._trash_busy = True
        self.trashChanged.emit()

        def worker() -> None:
            try:
                store.restore_from_trash(batch_id)
            except Exception as exc:
                self._trashActionFinished.emit(False, str(exc))
            else:
                self._trashActionFinished.emit(True, "处理批次已恢复到当前项目，现有资料没有被覆盖。")

        threading.Thread(target=worker, daemon=True, name="HRToolkit-project-trash-restore").start()

    @Slot(bool, str)
    def _apply_trash_action(self, success: bool, message: str) -> None:
        self._trash_busy = False
        self.trashChanged.emit()
        self.notificationRequested.emit("已恢复到项目" if success else "恢复没有完成", message, "success" if success else "error")
        self.refreshWorkspace()
        self.requestProjectTrash()

    def _selected_workspace_batch(self):
        store = self._project_store
        selected = self._workspace_selected_path
        if store is None or selected is None:
            return None
        try:
            locations = store.list_batch_locations()
        except Exception:
            return None
        for summary, directories in locations:
            for directory in directories.values():
                try:
                    selected.relative_to(directory)
                except ValueError:
                    continue
                return summary
        return None

    @Slot()
    def requestMoveSelectedBatchToTrash(self) -> None:
        summary = self._selected_workspace_batch()
        if summary is None:
            self.notificationRequested.emit("请选择处理批次", "请先在项目文件中选择某个处理批次内的文件或文件夹。", "warning")
            return
        token = f"project-trash:{time.monotonic_ns()}"
        self._pending_confirmation = token
        self._pending_confirmation_action = ("project-trash", summary.id)
        title = summary.business_description or summary.directory_name or summary.tool_name
        self.confirmationRequested.emit(
            "移到项目回收站",
            f"“{title}”的上传资料、处理结果和补充资料会一起移到当前项目回收站，不会永久删除。是否继续？",
            token,
        )

    def _move_project_batch_to_trash(self, batch_id: str) -> None:
        store = self._project_store
        if store is None or not store.writable or self._busy or self._workspace_busy:
            return
        self._workspace_busy = True
        self.workspaceBusyChanged.emit()

        def worker() -> None:
            try:
                store.move_to_trash(batch_id)
            except Exception as exc:
                self._workspaceImportFinished.emit(False, f"无法移到项目回收站：{exc}")
            else:
                self._workspaceImportFinished.emit(True, "完整处理批次已移到当前项目回收站，之后可以恢复。")

        threading.Thread(target=worker, daemon=True, name="HRToolkit-project-trash-move").start()

    @Slot()
    def requestStartupUpdateCheck(self) -> None:
        if update_check_enabled():
            self._start_update_check(False)

    @Slot()
    def requestUpdateCheck(self) -> None:
        self._start_update_check(True)

    def _start_update_check(self, manual: bool) -> None:
        if self._update_busy:
            if manual:
                self.notificationRequested.emit("正在检查更新", "请稍候，检查完成后会自动提示。", "info")
            return
        self._update_busy = True
        self._update_manual = bool(manual)
        self._update_status = "正在检查更新…"
        self._update_progress = -1.0
        self.updateChanged.emit()

        def worker() -> None:
            try:
                update = check_for_update(__version__)
            except Exception as exc:
                self._updateResult.emit("check-error", str(exc))
            else:
                self._updateResult.emit("available" if update is not None else "none", update)

        threading.Thread(target=worker, daemon=True, name="HRToolkit-update-check").start()

    @Slot(str, object)
    def _apply_update_result(self, kind: str, payload) -> None:
        if kind == "check-error":
            self._update_busy = False
            self._update_status = "检查更新失败"
            self.updateChanged.emit()
            if getattr(self, "_update_manual", False):
                self.notificationRequested.emit("检查更新失败", str(payload), "error")
            return
        if kind == "none":
            self._update_busy = False
            self._update_status = "当前已经是最新版本"
            self.updateChanged.emit()
            if getattr(self, "_update_manual", False):
                self.notificationRequested.emit("当前已是最新版本", f"当前版本为 v{__version__}。", "success")
            return
        if kind == "available" and isinstance(payload, UpdateInfo):
            self._update_busy = False
            self._pending_update = payload
            self._update_status = f"发现新版本 v{payload.version}"
            self.updateChanged.emit()
            notes = "\n".join(f"• {item}" for item in (payload.notes or ("本次发布未填写更新说明。",)))
            if payload.update_mode == "manual":
                detail = "macOS 使用标准 DMG 手动更新。点击“是”后将打开下载地址。"
            elif payload.mandatory:
                detail = "这是必须安装的更新。选择“否”将退出程序。"
            else:
                detail = "建议尽快更新。选择“否”可以继续使用当前版本。"
            token = f"update:{time.monotonic_ns()}"
            self._pending_confirmation = token
            self._pending_confirmation_action = ("update", payload)
            self.confirmationRequested.emit(f"发现新版本 v{payload.version}", f"{detail}\n\n{notes}", token)
            return
        if kind == "manual-ready" and isinstance(payload, str):
            self._update_busy = False
            self._update_status = "已打开下载地址"
            self.updateChanged.emit()
            QDesktopServices.openUrl(QUrl(payload))
            return
        if kind == "download-error":
            self._update_busy = False
            self._update_status = "更新下载失败"
            self._update_progress = -1.0
            self.updateChanged.emit()
            self.notificationRequested.emit("更新失败", str(payload), "error")
            return
        if kind == "download-cancelled":
            self._update_busy = False
            self._update_status = "更新下载已取消"
            self._update_progress = -1.0
            self.updateChanged.emit()
            return
        if kind == "launch-error":
            self._update_busy = False
            self._update_status = "更新程序启动失败"
            self.updateChanged.emit()
            self.notificationRequested.emit("更新失败", str(payload), "error")
            return
        if kind == "launched":
            self._update_status = "安装程序已启动，正在关闭当前版本…"
            self._update_progress = 1.0
            self.updateChanged.emit()
            QTimer.singleShot(500, QCoreApplication.quit)

    def _accept_update(self, update: UpdateInfo) -> None:
        if self._busy or self._workspace_busy:
            self.notificationRequested.emit("请先完成当前处理", "请等待当前处理或资料保存安全结束后再更新。", "warning")
            return
        self._update_busy = True
        self._update_progress = -1.0
        self._update_status = f"正在准备 v{update.version}…"
        self.updateChanged.emit()
        if update.update_mode == "manual":
            def manual_worker() -> None:
                try:
                    url = resolve_download_url(update)
                except Exception as exc:
                    self._updateResult.emit("download-error", str(exc))
                else:
                    self._updateResult.emit("manual-ready", url)

            threading.Thread(target=manual_worker, daemon=True, name="HRToolkit-update-url").start()
            return
        self._update_cancel_event = threading.Event()

        def download_worker() -> None:
            last_emit = 0.0

            def progress(downloaded: int, total: int) -> None:
                nonlocal last_emit
                now = time.monotonic()
                if now - last_emit >= 0.1 or (total > 0 and downloaded >= total):
                    last_emit = now
                    self._updateProgressIncoming.emit(int(downloaded), int(total))

            try:
                package = download_update_package(
                    update,
                    progress_callback=progress,
                    cancel_event=self._update_cancel_event,
                )
                launch_update_replacement(package)
            except UpdateCancelledError:
                self._updateResult.emit("download-cancelled", None)
            except Exception as exc:
                self._updateResult.emit("launch-error" if 'package' in locals() else "download-error", str(exc))
            else:
                self._updateResult.emit("launched", None)

        threading.Thread(target=download_worker, daemon=True, name="HRToolkit-update-download").start()

    @Slot(int, int)
    def _apply_update_progress(self, downloaded: int, total: int) -> None:
        megabytes = downloaded / 1024 / 1024
        if total > 0:
            self._update_progress = min(1.0, downloaded / total)
            self._update_status = f"正在下载：{self._update_progress * 100:.0f}%（{megabytes:.1f}/{total / 1024 / 1024:.1f} MB）"
        else:
            self._update_progress = -1.0
            self._update_status = f"正在下载：{megabytes:.1f} MB"
        self.updateChanged.emit()

    @Slot()
    def cancelUpdate(self) -> None:
        if self._update_cancel_event is not None:
            self._update_cancel_event.set()

    @Slot()
    def runOrCancel(self) -> None:
        if self._busy:
            if self._preview_cancel_event is not None:
                self._preview_cancel_event.set()
            self._run_coordinator.cancel()
            self._append_log("已请求停止，正在安全结束…", "warning")
            return
        if self._workspace_busy:
            self.notificationRequested.emit("项目资料正在保存", "请等待资料保存完成后再开始处理。", "warning")
            return
        store = self._project_store
        if store is None:
            self.notificationRequested.emit("请先打开工作项目", "请先新建或打开一个工作项目。", "warning")
            return
        if not store.writable:
            self.notificationRequested.emit("当前项目只能查看", store.workspace.read_only_reason or "项目为只读状态。", "warning")
            return
        try:
            invocation = build_invocation(
                self._spec,
                input_paths=list(self._input_states[self._state_key()]),
                support_text=self._support_states[self._state_key()],
                values=dict(self._form_states[self._state_key()]),
                output_dir=self._project_path,
                preview=self._spec.tool_id == "folder_rename",
            )
        except FormValidationError as exc:
            self.notificationRequested.emit(exc.title, exc.message, "warning")
            return
        if invocation.preview:
            self._start_preview(invocation)
        else:
            self._start_project_run(invocation)

    def _set_busy(self, value: bool) -> None:
        if self._busy == value:
            return
        self._busy = value
        if value:
            try:
                original = float(sys.getswitchinterval())
                self._original_switch_interval = original
                if original > 0.001:
                    sys.setswitchinterval(0.001)
            except Exception:
                self._original_switch_interval = None
        elif self._original_switch_interval is not None:
            try:
                sys.setswitchinterval(self._original_switch_interval)
            except Exception:
                pass
            self._original_switch_interval = None
        self.busyChanged.emit()

    def _start_preview(self, invocation: ToolInvocation) -> None:
        from hr_toolkit.background_process import run_business_process

        self._set_busy(True)
        self._preview_cancel_event = threading.Event()
        self._log_model.clear()
        self._append_log("正在生成改名预览…", "info")

        def worker() -> None:
            try:
                result = run_business_process(
                    module_name=invocation.function_module,
                    function_name=invocation.function_name,
                    args=invocation.args,
                    kwargs=invocation.kwargs,
                    cancel_event=self._preview_cancel_event,
                )
            except Exception as exc:
                self._previewFailed.emit(str(exc))
                return
            self._previewReady.emit(result.payload)

        threading.Thread(target=worker, daemon=True, name="HRToolkit-rename-preview").start()

    @Slot(object)
    def _apply_preview(self, payload: dict[str, Any]) -> None:
        self._preview_cancel_event = None
        self._set_busy(False)
        operations = list(payload.get("operations", []))
        warnings = list(payload.get("warnings", []))
        self._append_log(f"预览完成：可改名 {len(operations)} 项。", "info")
        for item in operations[:120]:
            self._append_log(f"{Path(item['source']).name}  →  {Path(item['target']).name}", "muted")
        if len(operations) > 120:
            self._append_log(f"另有 {len(operations) - 120} 项，界面不再逐条显示。", "muted")
        for warning in warnings[:20]:
            self._append_log(str(warning), "warning")
        if not operations:
            self.notificationRequested.emit("没有可改名项目", "没有找到可以安全改名的项目，请检查目录、名单、文件类型和预览提醒。", "info")
            return
        self._pending_preview = dict(payload)
        token = f"rename:{time.monotonic_ns()}"
        self._pending_confirmation = token
        self._pending_confirmation_action = ("rename", None)
        message = f"即将改名 {len(operations)} 项。工具会再次核对预览，只有内容完全一致才执行。是否继续？"
        self.confirmationRequested.emit("确认改名", message, token)

    @Slot(str)
    def _apply_preview_error(self, message: str) -> None:
        self._preview_cancel_event = None
        self._set_busy(False)
        self._last_run_by_key[self._state_key()] = (datetime.now().strftime("%H:%M"), False)
        self.specChanged.emit()
        self.notificationRequested.emit("预览失败", message, "error")

    @Slot(str, bool)
    def confirmAction(self, token: str, accepted: bool) -> None:
        if token != self._pending_confirmation:
            return
        self._pending_confirmation = None
        action = self._pending_confirmation_action
        self._pending_confirmation_action = None
        if not accepted:
            if action and action[0] == "rename":
                self._pending_preview = None
                self._append_log("已取消执行。", "muted")
            elif action and action[0] == "update" and isinstance(action[1], UpdateInfo) and action[1].mandatory:
                QCoreApplication.quit()
            return
        if not action:
            return
        action_name, payload = action
        if action_name == "material-delete":
            try:
                result = self._material_preferences.remove_material(str(payload))
            except ValueError as exc:
                self.notificationRequested.emit("无法删除材料", str(exc), "warning")
                return
            self._material_mutated()
            message = f"已删除自定义材料“{payload}”。"
            if result.updated_presets:
                message += "\n已更新预设：" + "、".join(result.updated_presets)
            if result.removed_presets:
                message += "\n已删除空预设：" + "、".join(result.removed_presets)
            self.notificationRequested.emit("材料已删除", message, "success")
            return
        if action_name == "preset-delete":
            try:
                self._material_preferences.delete_preset(str(payload))
            except ValueError as exc:
                self.notificationRequested.emit("无法删除预设", str(exc), "warning")
                return
            names = self._material_preferences.preset_names
            self._material_preset_name = names[0] if names else ""
            self._material_mutated()
            return
        if action_name == "history-trash":
            self._move_history_to_trash(str(payload))
            return
        if action_name == "project-trash":
            self._move_project_batch_to_trash(str(payload))
            return
        if action_name == "update" and isinstance(payload, UpdateInfo):
            self._accept_update(payload)
            return
        if action_name == "close":
            self._begin_shutdown()
            return
        if action_name != "rename":
            return
        try:
            invocation = build_invocation(
                self._spec,
                input_paths=list(self._input_states[self._state_key()]),
                support_text=self._support_states[self._state_key()],
                values=dict(self._form_states[self._state_key()]),
                output_dir=self._project_path,
                preview=False,
                preview_result=self._pending_preview,
            )
        except FormValidationError as exc:
            self.notificationRequested.emit(exc.title, exc.message, "warning")
            return
        finally:
            self._pending_preview = None
        self._start_project_run(invocation)

    def _start_project_run(self, invocation: ToolInvocation) -> None:
        store = self._project_store
        if store is None:
            return
        request = RunRequest(
            tool_id=invocation.tool_id,
            tool_name=invocation.tool_name,
            group_name=invocation.group_name,
            description=invocation.description,
            function=invocation.resolve_function(),
            args=invocation.args,
            kwargs=invocation.kwargs,
        )
        self._set_busy(True)
        self._log_model.clear()
        self._append_log(f"开始{invocation.tool_name}，请稍候…", "info")
        callbacks = RunCallbacks(
            log=lambda message: self._logIncoming.emit(str(message), "info"),
            progress=lambda current, total, message: self._runProgress.emit(int(current), int(total), str(message)),
            success=lambda payload, result_dir, elapsed, isolated: self._runSuccess.emit(payload, str(result_dir), float(elapsed), bool(isolated)),
            error=lambda error: self._runError.emit(str(error)),
            stopped=self._runStopped.emit,
            finished=self._runFinished.emit,
        )
        if not self._run_coordinator.start(store, request, callbacks):
            self._set_busy(False)
            self.notificationRequested.emit("已有任务正在处理", "请等待当前任务结束。", "warning")

    @Slot(int, int, str)
    def _apply_run_progress(self, current: int, total: int, message: str) -> None:
        self._append_log(message, "info")

    @Slot(object, str, float, bool)
    def _apply_run_success(self, payload, result_dir: str, elapsed: float, isolated: bool) -> None:
        self._last_result_dir = Path(result_dir)
        self.lastResultChanged.emit()
        self._last_run_by_key[self._state_key()] = (datetime.now().strftime("%H:%M"), True)
        self.specChanged.emit()
        warnings = list(payload.get("warnings", [])) if isinstance(payload, dict) else []
        mode_text = "独立进程" if isolated else "后台线程"
        self._append_log(f"处理完成，用时 {elapsed:.1f} 秒（{mode_text}）。", "success")
        if warnings:
            self._append_log(f"共有 {len(warnings)} 条提醒。", "warning")
            for warning in warnings[:30]:
                self._append_log(str(warning), "warning")
        self.notificationRequested.emit("处理完成", "结果已安全保存到当前项目。", "success")
        self.refreshWorkspace()

    @Slot(str)
    def _apply_run_error(self, message: str) -> None:
        self._last_run_by_key[self._state_key()] = (datetime.now().strftime("%H:%M"), False)
        self.specChanged.emit()
        self._append_log(f"处理失败：{message}", "error")
        self.notificationRequested.emit("处理失败", message, "error")
        self.refreshWorkspace()

    @Slot()
    def _apply_run_stopped(self) -> None:
        self._append_log("本次处理已安全停止。", "warning")
        self.notificationRequested.emit("已停止", "本次处理已安全结束，未完成批次可在项目中追溯。", "info")
        self.refreshWorkspace()

    @Slot()
    def _apply_run_finished(self) -> None:
        self._set_busy(False)

    @Slot(str, str)
    def _append_log(self, text: str, level: str = "info") -> None:
        self._log_model.append(
            {"time": datetime.now().strftime("%H:%M:%S"), "text": str(text), "level": str(level)},
            maximum=MAX_LOG_ROWS,
        )

    @Slot()
    def openLastResult(self) -> None:
        if self._last_result_dir is not None and self._last_result_dir.exists():
            open_path(self._last_result_dir)

    @Slot()
    def openRunLog(self) -> None:
        try:
            path = runlog.run_log_path()
        except Exception:
            return
        if path.exists():
            open_path(path)

    @Slot(str)
    def openUrl(self, url: str) -> None:
        QDesktopServices.openUrl(QUrl(url))

    @Slot()
    def _begin_shutdown(self) -> None:
        if self._shutdown_requested or self._closed:
            return
        self._shutdown_requested = True
        self._shutdown_wait_started = time.monotonic()
        if self._preview_cancel_event is not None:
            self._preview_cancel_event.set()
        self._run_coordinator.cancel()
        if self._workspace_cancel_event is not None:
            self._workspace_cancel_event.set()
        if self._update_cancel_event is not None:
            self._update_cancel_event.set()
        self._append_log("正在安全结束后台任务并关闭…", "warning")
        QTimer.singleShot(50, self._poll_shutdown)

    def _shutdown_work_running(self) -> bool:
        return bool(
            self._busy
            or self._workspace_busy
            or self._project_opening
            or self._history_busy
            or self._trash_busy
            or self._run_coordinator.running
            or self._preview_cancel_event is not None
        )

    @Slot()
    def _poll_shutdown(self) -> None:
        if self._closed:
            return
        elapsed = time.monotonic() - self._shutdown_wait_started
        if self._shutdown_work_running() and elapsed < 10.0:
            QTimer.singleShot(100, self._poll_shutdown)
            return
        if self._shutdown_work_running():
            runlog.log_line("关闭等待超过 10 秒；保留项目写锁并交由下次启动恢复未完成批次。")
        self.close()
        QCoreApplication.quit()

    @Slot(result=bool)
    def requestClose(self) -> bool:
        if self._closed:
            return True
        if self._shutdown_requested:
            return False
        if self._shutdown_work_running():
            if self._pending_confirmation_action and self._pending_confirmation_action[0] == "close":
                return False
            token = f"close:{time.monotonic_ns()}"
            self._pending_confirmation = token
            self._pending_confirmation_action = ("close", None)
            self.confirmationRequested.emit(
                "处理尚未结束",
                "当前处理或资料保存还没有完全结束。现在退出会先请求安全停止，并把未完成批次留待下次打开时恢复。是否仍要退出？",
                token,
            )
            return False
        self.close()
        return True

    @Slot()
    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        if self._preview_cancel_event is not None:
            self._preview_cancel_event.set()
        self._run_coordinator.cancel()
        if self._workspace_cancel_event is not None:
            self._workspace_cancel_event.set()
        if self._workspace_scan_cancel_event is not None:
            self._workspace_scan_cancel_event.set()
        if self._update_cancel_event is not None:
            self._update_cancel_event.set()
        self._save_workspace_preferences()
        # Do not release the writer lock while a worker can still write.  The
        # process exits immediately after the Qt event loop and project recovery
        # will safely close any interrupted batch on next open.
        project_worker_running = bool(
            self._run_coordinator.running
            or self._workspace_busy
            or self._project_opening
            or self._trash_busy
        )
        if self._project_store is not None and not project_worker_running:
            try:
                self._project_store.close()
            except Exception as exc:
                runlog.log_exception("关闭工作项目失败", exc)
