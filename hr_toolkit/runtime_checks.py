from __future__ import annotations

import os
import sys
import tempfile
import zipfile
from pathlib import Path

from hr_toolkit import __version__
from hr_toolkit.common.resources import open_template_resource


CHECK_OUTPUT_ENV = "HR_TOOLKIT_CHECK_OUTPUT"
TEMPLATE_NAMES = (
    "archive_company_template.xlsx",
    "archive_summary_template.xlsx",
    "data_statistics_template.xlsx",
    "insurance_ledger_template.xlsx",
    "personnel_change_summary_template.xlsx",
    "social_security_detail_template.xlsx",
    "social_security_summary_template.xlsx",
)


def run_headless_command(argv: list[str]) -> int | None:
    """Handle packaged verification commands without creating a Tk window."""
    if argv == ["--version"]:
        _emit(__version__)
        return 0
    if argv == ["--smoke-test"]:
        smoke_test()
        _emit(f"HRToolkit {__version__} smoke-test OK")
        return 0
    if argv == ["--update-smoke-test"]:
        latest_version = update_smoke_test()
        _emit(f"HRToolkit {__version__} update-smoke-test OK; latest={latest_version}")
        return 0
    return None


def smoke_test() -> None:
    """Validate dependencies and packaged whitelist resources without a GUI."""
    import openpyxl  # noqa: F401
    import xlrd  # noqa: F401
    from hr_toolkit.app_update import create_https_context
    from hr_toolkit.project_store import ProjectStore

    # Loading the context proves that PyInstaller included certifi's CA bundle.
    create_https_context()

    for template_name in TEMPLATE_NAMES:
        with open_template_resource(template_name) as handle:
            if not zipfile.is_zipfile(handle):
                raise RuntimeError(f"模板资源不是有效的 xlsx：{template_name}")

    with tempfile.TemporaryDirectory(prefix="hr_toolkit_smoke_") as temp_root:
        # macOS 上 tempfile 可能返回经过 /var -> /private/var 的系统链接；
        # 运行检查使用真实路径，不降低项目对链接路径的安全限制。
        project_root = Path(temp_root).resolve() / "project"
        with ProjectStore.create(project_root, "运行检查项目") as project:
            draft = project.create_draft_batch(
                group_name="薪酬管理",
                tool_id="salary_split",
                tool_name="工资表拆分",
                business_description="运行检查",
                business_period="临时",
            )
            batch_id = draft.summary.id
            project.start_processing(batch_id)
            project.mark_success(batch_id)
            detail = project.get_batch(batch_id)
            if (
                detail is None
                or detail.summary.status != "success"
                or not (project_root / ".hrtoolkit").is_dir()
                or not project.verify_batch_files(batch_id)
            ):
                raise RuntimeError("本地项目工作区运行检查失败。")

    if getattr(sys, "frozen", False):
        bundle_root = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        if not (bundle_root / "README.md").is_file():
            raise RuntimeError("打包程序缺少 README.md。")


def update_smoke_test() -> str:
    """Verify secure Gitee-first metadata discovery with GitHub fallback."""
    from hr_toolkit.app_update import check_for_update

    update = check_for_update("0.0.0")
    if update is None or not update.version:
        raise RuntimeError("更新配置缺少当前平台版本。")
    return update.version


def _emit(text: str) -> None:
    """Write to an attached console and, optionally, a CI result file."""
    if sys.stdout is not None:
        print(text, flush=True)
    output_path = os.environ.get(CHECK_OUTPUT_ENV, "").strip()
    if output_path:
        Path(output_path).write_text(text + "\n", encoding="utf-8")
