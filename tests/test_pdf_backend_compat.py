from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path

from hr_toolkit.tools import material_collector as mc


def _write_minimal_pdf(
    path: Path,
    page_kinds: list[str],
    *,
    text_content: str = "FULL_TEXT_AFTER_LIMIT_MARKER",
) -> None:
    """生成不依赖第三方写入库的严格 PDF，供现代与 Win7 后端共用。"""
    objects: list[bytes] = [b""]

    def add_object(payload: bytes = b"") -> int:
        objects.append(payload)
        return len(objects) - 1

    catalog_id = add_object()
    pages_id = add_object()
    font_id = add_object(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    )
    page_ids: list[int] = []

    for index, kind in enumerate(page_kinds):
        if kind == "text":
            content = (
                b"BT /F1 12 Tf 20 100 Td "
                + b"(" + text_content.encode("ascii") + b") Tj ET"
            )
            content_id = add_object(
                b"<< /Length " + str(len(content)).encode("ascii") + b" >>\n"
                b"stream\n" + content + b"\nendstream"
            )
            resources = (
                b"<< /Font << /F1 "
                + str(font_id).encode("ascii")
                + b" 0 R >> >>"
            )
        elif kind == "scan":
            width = 16
            height = 16
            image_bytes = bytes([80 + index]) * width * height * 3
            image_id = add_object(
                b"<< /Type /XObject /Subtype /Image /Width 16 /Height 16 "
                b"/ColorSpace /DeviceRGB /BitsPerComponent 8 /Length "
                + str(len(image_bytes)).encode("ascii")
                + b" >>\nstream\n"
                + image_bytes
                + b"\nendstream"
            )
            content = b"q 300 0 0 300 0 0 cm /Im0 Do Q"
            content_id = add_object(
                b"<< /Length " + str(len(content)).encode("ascii") + b" >>\n"
                b"stream\n" + content + b"\nendstream"
            )
            resources = (
                b"<< /XObject << /Im0 "
                + str(image_id).encode("ascii")
                + b" 0 R >> >>"
            )
        else:
            raise ValueError(f"unknown test page kind: {kind}")

        page_id = add_object(
            b"<< /Type /Page /Parent "
            + str(pages_id).encode("ascii")
            + b" 0 R /MediaBox [0 0 300 300] /Resources "
            + resources
            + b" /Contents "
            + str(content_id).encode("ascii")
            + b" 0 R >>"
        )
        page_ids.append(page_id)

    objects[catalog_id] = (
        b"<< /Type /Catalog /Pages "
        + str(pages_id).encode("ascii")
        + b" 0 R >>"
    )
    kids = b" ".join(f"{page_id} 0 R".encode("ascii") for page_id in page_ids)
    objects[pages_id] = (
        b"<< /Type /Pages /Kids ["
        + kids
        + b"] /Count "
        + str(len(page_ids)).encode("ascii")
        + b" >>"
    )

    document = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for object_number, payload in enumerate(objects[1:], start=1):
        offsets.append(len(document))
        document.extend(f"{object_number} 0 obj\n".encode("ascii"))
        document.extend(payload)
        document.extend(b"\nendobj\n")
    xref_offset = len(document)
    document.extend(f"xref\n0 {len(objects)}\n".encode("ascii"))
    document.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        document.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    document.extend(
        b"trailer\n<< /Size "
        + str(len(objects)).encode("ascii")
        + b" /Root "
        + str(catalog_id).encode("ascii")
        + b" 0 R >>\nstartxref\n"
        + str(xref_offset).encode("ascii")
        + b"\n%%EOF\n"
    )
    path.write_bytes(document)


class PdfBackendCompatibilityTests(unittest.TestCase):
    def test_active_backend_extracts_full_text_without_mutating_source(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / "text.pdf"
            _write_minimal_pdf(source, ["text"])
            before = hashlib.sha256(source.read_bytes()).hexdigest()

            text = mc._extract_document_text(source)

            self.assertIn("FULL_TEXT_AFTER_LIMIT_MARKER", text)
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), before)

    def test_active_backend_ocr_processes_every_scan_page_and_releases_file(self) -> None:
        real_engine = mc._OCR_ENGINE
        real_attempted = mc._OCR_ATTEMPTED
        try:
            class CountingEngine:
                calls = 0

                def __call__(self, payload):
                    self.assert_payload(payload)
                    type(self).calls += 1
                    return ([[None, "劳动合同"]], None)

                @staticmethod
                def assert_payload(payload) -> None:
                    if not isinstance(payload, (bytes, bytearray)) or not payload:
                        raise AssertionError("PDF OCR payload is empty")

            mc._OCR_ENGINE = CountingEngine()
            mc._OCR_ATTEMPTED = True
            with tempfile.TemporaryDirectory() as td:
                source = Path(td) / "scan.pdf"
                renamed = Path(td) / "renamed.pdf"
                _write_minimal_pdf(source, ["scan", "scan"])
                before = hashlib.sha256(source.read_bytes()).hexdigest()
                progress: list[str] = []

                result = mc._analyze_folder_ocr_file(
                    source,
                    requested_types=["劳动合同"],
                    progress_callback=lambda _current, _total, message: progress.append(message),
                )
                source.replace(renamed)

                self.assertEqual(result[0], "劳动合同")
                self.assertEqual(CountingEngine.calls, 2)
                self.assertTrue(any("2/2" in message for message in progress))
                self.assertEqual(hashlib.sha256(renamed.read_bytes()).hexdigest(), before)
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted

    def test_active_backend_honors_cancel_and_rejects_corrupt_pdf(self) -> None:
        real_engine = mc._OCR_ENGINE
        real_attempted = mc._OCR_ATTEMPTED
        try:
            class CountingEngine:
                calls = 0

                def __call__(self, _payload):
                    type(self).calls += 1
                    return ([[None, "未分类"]], None)

            mc._OCR_ENGINE = CountingEngine()
            mc._OCR_ATTEMPTED = True
            with tempfile.TemporaryDirectory() as td:
                root = Path(td)
                scan = root / "scan.pdf"
                _write_minimal_pdf(scan, ["scan", "scan", "scan"])
                with self.assertRaises(mc.MaterialCollectionCancelled):
                    mc._analyze_folder_ocr_file(
                        scan,
                        cancelled=lambda: CountingEngine.calls >= 1,
                    )
                self.assertEqual(CountingEngine.calls, 1)

                corrupt = root / "corrupt.pdf"
                corrupt.write_bytes(b"%PDF-1.4\ntruncated")
                with self.assertRaisesRegex(mc.PDFRecognitionError, "损坏|异常"):
                    mc._extract_document_text(corrupt)
        finally:
            mc._OCR_ENGINE = real_engine
            mc._OCR_ATTEMPTED = real_attempted


if __name__ == "__main__":
    unittest.main()
