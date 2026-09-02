from __future__ import annotations

import base64
import os
import struct
import sys
import tempfile
import threading
import traceback
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
_PDF_RUNTIME_FIXTURE = base64.b64decode(
    "JVBERi0xLjQKJeLjz9MKMSAwIG9iago8PCAvVHlwZSAvQ2F0YWxvZyAvUGFnZXMgMiAwIFIgPj4KZW5kb2JqCjIgMCBvYmoKPDwgL1R5cGUgL1BhZ2VzIC9LaWRzIFs1IDAgUl0gL0NvdW50IDEgPj4KZW5kb2JqCjMgMCBvYmoKPDwgL1R5cGUgL0ZvbnQgL1N1YnR5cGUgL1R5cGUxIC9CYXNlRm9udCAvSGVsdmV0aWNhID4+CmVuZG9iago0IDAgb2JqCjw8IC9MZW5ndGggNTkgPj4Kc3RyZWFtCkJUIC9GMSAxMiBUZiAyMCAxMDAgVGQgKEZVTExfVEVYVF9BRlRFUl9MSU1JVF9NQVJLRVIpIFRqIEVUCmVuZHN0cmVhbQplbmRvYmoKNSAwIG9iago8PCAvVHlwZSAvUGFnZSAvUGFyZW50IDIgMCBSIC9NZWRpYUJveCBbMCAwIDMwMCAzMDBdIC9SZXNvdXJjZXMgPDwgL0ZvbnQgPDwgL0YxIDMgMCBSID4+ID4+IC9Db250ZW50cyA0IDAgUiA+PgplbmRvYmoKeHJlZgowIDYKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDAwMDE1IDAwMDAwIG4gCjAwMDAwMDAwNjQgMDAwMDAgbiAKMDAwMDAwMDEyMSAwMDAwMCBuIAowMDAwMDAwMTkxIDAwMDAwIG4gCjAwMDAwMDAzMDAgMDAwMDAgbiAKdHJhaWxlcgo8PCAvU2l6ZSA2IC9Sb290IDEgMCBSID4+CnN0YXJ0eHJlZgo0MjYKJSVFT0YK"
)
_PDF_SCAN_RUNTIME_FIXTURE = base64.b64decode(
    "JVBERi0xLjQKJeLjz9MKMSAwIG9iago8PCAvVHlwZSAvQ2F0YWxvZyAvUGFnZXMgMiAwIFIgPj4KZW5kb2JqCjIgMCBvYmoKPDwg"
    "L1R5cGUgL1BhZ2VzIC9LaWRzIFs1IDAgUl0gL0NvdW50IDEgPj4KZW5kb2JqCjMgMCBvYmoKPDwgL1R5cGUgL1hPYmplY3QgL1N1"
    "YnR5cGUgL0ltYWdlIC9XaWR0aCAyIC9IZWlnaHQgMiAvQ29sb3JTcGFjZSAvRGV2aWNlUkdCIC9CaXRzUGVyQ29tcG9uZW50IDgg"
    "L0xlbmd0aCAxMiA+PgpzdHJlYW0K/wAAAP8AAAD/////CmVuZHN0cmVhbQplbmRvYmoKNCAwIG9iago8PCAvTGVuZ3RoIDMwID4+"
    "CnN0cmVhbQpxIDMwMCAwIDAgMzAwIDAgMCBjbSAvSW0wIERvIFEKZW5kc3RyZWFtCmVuZG9iago1IDAgb2JqCjw8IC9UeXBlIC9Q"
    "YWdlIC9QYXJlbnQgMiAwIFIgL01lZGlhQm94IFswIDAgMzAwIDMwMF0gL1Jlc291cmNlcyA8PCAvWE9iamVjdCA8PCAvSW0wIDMg"
    "MCBSID4+ID4+IC9Db250ZW50cyA0IDAgUiA+PgplbmRvYmoKeHJlZgowIDYKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDAwMDE1"
    "IDAwMDAwIG4gCjAwMDAwMDAwNjQgMDAwMDAgbiAKMDAwMDAwMDEyMSAwMDAwMCBuIAowMDAwMDAwMjc2IDAwMDAwIG4gCjAwMDAw"
    "MDAzNTYgMDAwMDAgbiAKdHJhaWxlcgo8PCAvU2l6ZSA2IC9Sb290IDEgMCBSID4+CnN0YXJ0eHJlZgo0ODYKJSVFT0YK"
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
        _mark_smoke_stage("ocr-import")
        from rapidocr_onnxruntime import RapidOCR

        _mark_smoke_stage("ocr-engine-init")
        engine = RapidOCR()
        _mark_smoke_stage("ocr-inference")
        result = engine(str(image_path))
    except Exception as exc:
        raise RuntimeError(f"本地 OCR 引擎运行检查失败：{exc}") from exc
    if not isinstance(result, tuple) or len(result) != 2:
        raise RuntimeError(f"本地 OCR 引擎返回格式无效：{type(result).__name__}")


def pdf_runtime_smoke_test(root: Path | None = None) -> None:
    """实际加载文字页和扫描页，验证打包后端的两条识别链路。"""
    if root is None:
        with tempfile.TemporaryDirectory(prefix="hr_toolkit_pdf_smoke_") as temp_root:
            pdf_runtime_smoke_test(Path(temp_root))
        return

    root.mkdir(parents=True, exist_ok=True)
    source = root / "runtime.pdf"
    scan_source = root / "runtime-scan.pdf"
    source.write_bytes(_PDF_RUNTIME_FIXTURE)
    scan_source.write_bytes(_PDF_SCAN_RUNTIME_FIXTURE)
    try:
        from hr_toolkit.tools.material_collector import (
            _extract_document_text,
            _iter_pdf_ocr_images,
        )

        text = _extract_document_text(source)
        scan_pages = 0
        for payload in _iter_pdf_ocr_images(scan_source):
            if not payload:
                raise RuntimeError("PDF 扫描页解码结果为空")
            scan_pages += 1
    except Exception as exc:
        raise RuntimeError(f"PDF 识别组件运行检查失败：{exc}") from exc
    if "FULL_TEXT_AFTER_LIMIT_MARKER" not in text:
        raise RuntimeError("PDF 识别组件未能提取完整文字层")
    if scan_pages != 1:
        raise RuntimeError(f"PDF 扫描页解码数量异常：{scan_pages} != 1")


def run_headless_command(argv: list[str]) -> int | None:
    """Handle packaged verification commands without creating a Tk window."""
    if argv == ["--version"]:
        _emit(__version__)
        return 0
    if argv == ["--smoke-test"]:
        try:
            smoke_test()
        except Exception:
            _emit(
                f"HRToolkit {__version__} smoke-test FAILED\n"
                f"{traceback.format_exc().rstrip()}"
            )
            return 1
        else:
            _emit(f"HRToolkit {__version__} smoke-test OK")
            return 0
    if argv == ["--update-smoke-test"]:
        try:
            latest_version = update_smoke_test()
        except Exception:
            _emit(
                f"HRToolkit {__version__} update-smoke-test FAILED\n"
                f"{traceback.format_exc().rstrip()}"
            )
            return 1
        else:
            _emit(f"HRToolkit {__version__} update-smoke-test OK; latest={latest_version}")
            return 0
    return None


def smoke_test() -> None:
    """Validate dependencies and packaged whitelist resources without a GUI."""
    _mark_smoke_stage("dependencies")
    import openpyxl  # noqa: F401
    try:
        import pypdf  # noqa: F401
    except ImportError:
        import pypdfium2  # noqa: F401
    import xlrd  # noqa: F401
    from hr_toolkit.app_update import create_https_context
    from hr_toolkit.project_store import ProjectStore

    # Loading the context proves that PyInstaller included certifi's CA bundle.
    _mark_smoke_stage("https-context")
    create_https_context()

    _mark_smoke_stage("templates")
    for template_name in TEMPLATE_NAMES:
        with open_template_resource(template_name) as handle:
            if not zipfile.is_zipfile(handle):
                raise RuntimeError(f"模板资源不是有效的 xlsx：{template_name}")

    with tempfile.TemporaryDirectory(prefix="hr_toolkit_smoke_") as temp_root:
        # macOS 上 tempfile 可能返回经过 /var -> /private/var 的系统链接；
        # 运行检查使用真实路径，不降低项目对链接路径的安全限制。
        resolved_temp_root = Path(temp_root).resolve()
        _mark_smoke_stage("pdf-runtime")
        pdf_runtime_smoke_test(resolved_temp_root / "pdf-runtime")
        _mark_smoke_stage("archive-7z")
        _smoke_test_archive_runtimes(resolved_temp_root / "archive-runtime")
        if getattr(sys, "frozen", False):
            _mark_smoke_stage("background-process")
            from hr_toolkit.background_process import run_business_process

            process_result = run_business_process(
                module_name="hr_toolkit.background_process",
                function_name="_process_smoke_probe",
                args=(resolved_temp_root / "probe-input.xlsx", resolved_temp_root / "probe-output"),
                kwargs={},
                cancel_event=threading.Event(),
            )
            if process_result.payload.get("output_dir") != str(resolved_temp_root / "probe-output"):
                raise RuntimeError("打包后台进程没有返回预期结果。")
            ocr_runtime_smoke_test(resolved_temp_root / "ocr-runtime")
        _mark_smoke_stage("project-store")
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
        _mark_smoke_stage("readme")
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
    _mark_smoke_stage("archive-rar")
    rar_warnings: list[str] = []
    rar_files = extract_archive_excel_files(rar_path, root / "extract-rar", rar_warnings)
    if rar_warnings or len(rar_files) != 1 or rar_files[0].read_bytes() != payload:
        raise RuntimeError(f"RAR 解压组件运行检查失败：warnings={rar_warnings}，files={len(rar_files)}")


def update_smoke_test() -> str:
    """Verify secure Gitee-only metadata discovery for domestic clients."""
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


def _mark_smoke_stage(stage: str) -> None:
    """Persist the active frozen-check stage so CI can diagnose hard hangs."""
    output_path = os.environ.get(CHECK_OUTPUT_ENV, "").strip()
    if output_path:
        Path(output_path).write_text(
            f"HRToolkit {__version__} smoke-test RUNNING: {stage}\n",
            encoding="utf-8",
        )
