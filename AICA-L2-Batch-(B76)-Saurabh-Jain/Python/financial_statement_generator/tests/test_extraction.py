"""Tests for PDF upload validation and extraction."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from pypdf import PdfWriter

from app.extraction.exceptions import (
    EmptyPDFError,
    EncryptedPDFError,
    InvalidPDFError,
)
from app.extraction.ocr_extractor import (
    OCRExtractor,
    document_extraction_method,
    is_text_sufficient,
    resolve_extraction_method,
)
from app.extraction.pdf_extractor import NativeTextExtractor, PDFExtractor, extract_dates, extract_numbers
from app.extraction.table_extractor import TableExtractor
from app.main import app

TEXT_PDF = b"""%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj
4 0 obj
<< /Length 168 >>
stream
BT
/F1 12 Tf
72 720 Td
(BALANCE SHEET as at 31-03-2024) Tj
0 -20 Td
(Share Capital 1,00,000) Tj
0 -20 Td
(Revenue from operations 250000.50) Tj
ET
endstream
endobj
5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000266 00000 n 
0000000485 00000 n 
trailer
<< /Size 6 /Root 1 0 R >>
startxref
563
%%EOF
"""


def test_extraction_module_importable() -> None:
    assert PDFExtractor is not None
    assert TableExtractor is not None


def test_rejects_non_pdf_bytes() -> None:
    extractor = PDFExtractor()
    with pytest.raises(InvalidPDFError):
        extractor.validate_pdf_bytes(b"this is not a pdf", filename="notes.txt")


def test_rejects_wrong_extension() -> None:
    extractor = PDFExtractor()
    with pytest.raises(InvalidPDFError):
        extractor.validate_pdf_bytes(TEXT_PDF, filename="statement.xlsx")


def test_extracts_text_page_by_page(tmp_path: Path) -> None:
    pdf_path = tmp_path / "statement.pdf"
    pdf_path.write_bytes(TEXT_PDF)
    payload = PDFExtractor(upload_dir=tmp_path).extract(pdf_path, original_filename="statement.pdf")

    assert payload["document"]["filename"] == "statement.pdf"
    assert payload["document"]["pages"] == 1
    assert payload["pages"][0]["page_number"] == 1
    text = payload["pages"][0]["text"]
    assert "BALANCE SHEET" in text
    assert payload["pages"][0]["tables"] == [] or isinstance(payload["pages"][0]["tables"], list)
    assert any(section["name"] == "balance sheet" for section in payload["pages"][0]["sections"])
    assert any("31-03-2024" in item["raw"] for item in payload["pages"][0]["dates"])
    assert payload["pages"][0]["extraction_method"] == "text"
    assert payload["document"]["extraction_method"] == "text"


def test_empty_pdf_raises(tmp_path: Path) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    empty_path = tmp_path / "empty.pdf"
    writer.write(empty_path)
    writer.close()

    with pytest.raises(EmptyPDFError):
        PDFExtractor(upload_dir=tmp_path, enable_ocr=False).extract(empty_path)


def test_password_protected_pdf_raises(tmp_path: Path) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    writer.encrypt("secret-password")
    locked = tmp_path / "locked.pdf"
    writer.write(locked)
    writer.close()

    with pytest.raises(EncryptedPDFError):
        PDFExtractor(upload_dir=tmp_path).extract(locked)


def test_corrupted_pdf_is_rejected(tmp_path: Path) -> None:
    broken = tmp_path / "broken.pdf"
    broken.write_bytes(b"%PDF-1.4\nthis is not a complete pdf file")
    with pytest.raises(Exception):
        PDFExtractor(upload_dir=tmp_path).extract(broken)


def test_number_and_date_helpers() -> None:
    numbers = extract_numbers("Share Capital 1,00,000 and sales (25,000.50)")
    values = {item["value"] for item in numbers if item["value"] is not None}
    assert 100000.0 in values
    assert -25000.5 in values

    dates = extract_dates("Balance Sheet as at 31-03-2024")
    assert dates
    assert dates[0]["iso"] == "2024-03-31"


def test_upload_endpoint_extracts_pdf(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        "app.api.routes.PDFExtractor",
        lambda: PDFExtractor(upload_dir=tmp_path),
    )
    client = TestClient(app)
    response = client.post(
        "/extract",
        files={"file": ("statement.pdf", BytesIO(TEXT_PDF), "application/pdf")},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["document"]["filename"] == "statement.pdf"
    assert body["document"]["pages"] == 1
    assert "pages" in body
    assert body["pages"][0]["page_number"] == 1
    assert body["pages"][0]["extraction_method"] == "text"


def test_upload_endpoint_rejects_non_pdf() -> None:
    client = TestClient(app)
    response = client.post(
        "/extract",
        files={"file": ("notes.txt", BytesIO(b"hello world"), "text/plain")},
    )
    assert response.status_code == 400
    body = response.json()
    assert body["error"] == "invalid_pdf"


class _FakeOCR:
    def __init__(self, text: str = "OCR recovered BALANCE SHEET as at 31-03-2024") -> None:
        self.calls: list[int] = []
        self.text = text

    def extract_page_text(self, pdf_path: Path, page_number: int) -> str:
        self.calls.append(page_number)
        return f"{self.text} page {page_number}"


def test_native_and_ocr_share_page_text_interface() -> None:
    native = NativeTextExtractor()
    ocr = OCRExtractor()
    assert callable(native.extract_page_text)
    assert callable(ocr.extract_page_text)
    assert native.method == "text"
    assert ocr.method == "ocr"


def test_is_text_sufficient() -> None:
    assert is_text_sufficient("") is False
    assert is_text_sufficient("123") is False
    assert is_text_sufficient("BALANCE SHEET as at 31 March 2024 with share capital") is True


def test_text_extraction_does_not_call_ocr(tmp_path: Path) -> None:
    pdf_path = tmp_path / "digital.pdf"
    pdf_path.write_bytes(TEXT_PDF)
    fake = _FakeOCR()
    payload = PDFExtractor(upload_dir=tmp_path, ocr_engine=fake).extract(pdf_path)
    assert fake.calls == []
    assert payload["pages"][0]["extraction_method"] == "text"
    assert payload["document"]["extraction_method"] == "text"


def test_scanned_page_uses_ocr_fallback(tmp_path: Path) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    pdf_path = tmp_path / "scanned.pdf"
    writer.write(pdf_path)
    writer.close()

    fake = _FakeOCR()
    payload = PDFExtractor(upload_dir=tmp_path, ocr_engine=fake).extract(pdf_path)
    assert fake.calls == [1]
    page = payload["pages"][0]
    assert page["page_number"] == 1
    assert page["extraction_method"] == "ocr"
    assert "BALANCE SHEET" in page["text"]
    assert "tables" in page
    assert payload["document"]["extraction_method"] == "ocr"


def test_mixed_document_uses_text_and_ocr(tmp_path: Path) -> None:
    from pypdf import PdfReader

    writer = PdfWriter()
    writer.add_page(PdfReader(BytesIO(TEXT_PDF)).pages[0])
    writer.add_blank_page(width=612, height=792)
    pdf_path = tmp_path / "mixed.pdf"
    writer.write(pdf_path)
    writer.close()

    fake = _FakeOCR()
    payload = PDFExtractor(upload_dir=tmp_path, ocr_engine=fake).extract(pdf_path)
    assert fake.calls == [2]
    methods = [page["extraction_method"] for page in payload["pages"]]
    assert methods[0] == "text"
    assert methods[1] == "ocr"
    assert payload["document"]["extraction_method"] == "mixed"
    assert document_extraction_method(payload["pages"]) == "mixed"
    assert payload["pages"][0]["page_number"] == 1
    assert payload["pages"][1]["page_number"] == 2


def test_page_with_thin_native_text_is_mixed(tmp_path: Path) -> None:
    class _ThinNative:
        method = "text"

        def extract_page_text(self, pdf_path: Path, page_number: int) -> str:
            return "Header note"

    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    pdf_path = tmp_path / "thin.pdf"
    writer.write(pdf_path)
    writer.close()

    fake = _FakeOCR()
    payload = PDFExtractor(
        upload_dir=tmp_path,
        native_extractor=_ThinNative(),
        ocr_engine=fake,
    ).extract(pdf_path)
    page = payload["pages"][0]
    assert fake.calls == [1]
    assert page["extraction_method"] == "mixed"
    assert "Header note" in page["text"]
    assert "BALANCE SHEET" in page["text"]
    assert page["page_number"] == 1


def test_resolve_extraction_method() -> None:
    assert resolve_extraction_method(native_sufficient=True, used_ocr=False, had_native_signal=True) == "text"
    assert resolve_extraction_method(native_sufficient=False, used_ocr=True, had_native_signal=False) == "ocr"
    assert resolve_extraction_method(native_sufficient=False, used_ocr=True, had_native_signal=True) == "mixed"


def test_native_extractor_reads_page_text(tmp_path: Path) -> None:
    pdf_path = tmp_path / "native.pdf"
    pdf_path.write_bytes(TEXT_PDF)
    text = NativeTextExtractor().extract_page_text(pdf_path, 1)
    assert "BALANCE SHEET" in text


def test_ocr_extractor_unavailable_without_tesseract(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("app.extraction.ocr_extractor.shutil.which", lambda _: None)
    backend = OCRExtractor()
    assert backend.is_available() is False
    assert backend.extract_page_text(Path("missing.pdf"), 1) == ""
