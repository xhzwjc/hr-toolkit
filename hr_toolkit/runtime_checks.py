from __future__ import annotations

import base64
import os
import struct
import sys
import tempfile
import zipfile
import zlib
from pathlib import Path

from hr_toolkit import __version__
from hr_toolkit.common.inputs import extract_archive_excel_files
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
_RAR_RUNTIME_FIXTURE = base64.b64decode(
    "UmFyIRoHAQAzkrXlCgEFBgAFAQGAgACVrTL6KgIDC58ABJ8ApIMCvJeqS4AAAQxydW50aW1lLnhsc3gKAxPxW4hqhkEQMUhSVG9vbGtpdCBhcmNoaXZlIHJ1bnRpbWUgc21va2Udd1ZRAwUEAA=="
)
_SEVEN_ZIP_RUNTIME_FIXTURE = base64.b64decode(
    "N3q8ryccAASuJxD0iAAAAAAAAAAUAAAAAAAAAGkm1z8BAB5IUlRvb2xraXQgYXJjaGl2ZSBydW50aW1lIHNtb2tlAOAAXgBdXQAAgTMHrg/QPBb8nzkQnG8VArnDFMcdhtRaWFWIWBRIxoCITwQTg7wuT9/dT/wnHngb888SGgIyduSDds/zn2STt2jHBoaB/DbbyKI+Dksb3jMZ88uaze/zjyAAAAAAFwYjAQllAAcLAQABISEBGAxfAAA="
)


def _png_chunk(chunk_type: bytes, payload: bytes) -> bytes:
    checksum = zlib.crc32(chunk_type + payload) & 0xFFFFFFFF
    return struct.pack(">I", len(payload)) + chunk_type + payload + struct.pack(">I", checksum)


def _write_ocr_smoke_image(path: Path) -> None:
    """Write a dependency-free RGB PNG that exercises RapidOCR inference."""
    width, height = 64, 32
    scanline = b"\x00" + (b"\xff\xff\xff" * width)
    payload = (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + _png_chunk(b"IDAT", zlib.compress(scanline * height))
        + _png_chunk(b"IEND", b"")
    )
    path.write_bytes(payload)


def ocr_runtime_smoke_test(root: Path | None = None) -> None:
    """Load bundled OCR models and execute one real, offline inference call."""
    if root is None:
        with tempfile.TemporaryDirectory(prefix="hr_toolkit_ocr_smoke_") as temp_root:
            ocr_runtime_smoke_test(Path(temp_root))
        return

    root.mkdir(parents=True, exist_ok=True)
    image_path = root / "blank.png"
    _write_ocr_smoke_image(image_path)
    try:
        from rapidocr_onnxruntime import RapidOCR

        result = RapidOCR()(str(image_path))
    except Exception as exc:
        raise RuntimeError(f"本地 OCR 引擎运行检查失败：{exc}") from exc
    if not isinstance(result, tuple) or len(result) != 2:
        raise RuntimeError(f"本地 OCR 引擎返回格式无效：{type(result).__name__}")


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
        resolved_temp_root = Path(temp_root).resolve()
        _smoke_test_archive_runtimes(resolved_temp_root / "archive-runtime")
        if getattr(sys, "frozen", False):
            ocr_runtime_smoke_test(resolved_temp_root / "ocr-runtime")
        project_root = resolved_temp_root / "project"
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


def _smoke_test_archive_runtimes(root: Path) -> None:
    root.mkdir(parents=True, exist_ok=True)
    payload = b"HRToolkit archive runtime smoke"
    archive_path = root / "runtime.7z"
    archive_path.write_bytes(_SEVEN_ZIP_RUNTIME_FIXTURE)
    warnings: list[str] = []
    files = extract_archive_excel_files(archive_path, root / "extract", warnings)
    if warnings or len(files) != 1 or files[0].read_bytes() != payload:
        raise RuntimeError(f"7Z 解压组件运行检查失败：warnings={warnings}，files={len(files)}")

    rar_path = root / "runtime.rar"
    rar_path.write_bytes(_RAR_RUNTIME_FIXTURE)
    rar_warnings: list[str] = []
    rar_files = extract_archive_excel_files(rar_path, root / "extract-rar", rar_warnings)
    if rar_warnings or len(rar_files) != 1 or rar_files[0].read_bytes() != payload:
        raise RuntimeError(f"RAR 解压组件运行检查失败：warnings={rar_warnings}，files={len(rar_files)}")


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
