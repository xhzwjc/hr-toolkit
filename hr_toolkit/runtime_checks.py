from __future__ import annotations

import os
import sys
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

    # Loading the context proves that PyInstaller included certifi's CA bundle.
    create_https_context()

    for template_name in TEMPLATE_NAMES:
        with open_template_resource(template_name) as handle:
            if not zipfile.is_zipfile(handle):
                raise RuntimeError(f"模板资源不是有效的 xlsx：{template_name}")

    if getattr(sys, "frozen", False):
        bundle_root = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        if not (bundle_root / "README.md").is_file():
            raise RuntimeError("打包程序缺少 README.md。")


def update_smoke_test() -> str:
    """Verify that the packaged runtime can securely read GitHub release metadata."""
    from hr_toolkit.app_update import (
        DEFAULT_UPDATE_MANIFEST_URL,
        fetch_update_manifest,
        manifest_version,
        platform_key,
    )

    manifest = fetch_update_manifest(DEFAULT_UPDATE_MANIFEST_URL, timeout=30)
    latest_version = manifest_version(manifest, platform=platform_key())
    if not latest_version:
        raise RuntimeError("GitHub 更新配置缺少当前平台版本。")
    return latest_version


def _emit(text: str) -> None:
    """Write to an attached console and, optionally, a CI result file."""
    if sys.stdout is not None:
        print(text, flush=True)
    output_path = os.environ.get(CHECK_OUTPUT_ENV, "").strip()
    if output_path:
        Path(output_path).write_text(text + "\n", encoding="utf-8")
