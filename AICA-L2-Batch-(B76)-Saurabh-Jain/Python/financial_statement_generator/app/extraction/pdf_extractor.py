"""PDF upload validation and page-level text extraction.

This layer is independent of Schedule III mapping. It only turns an uploaded
PDF into structured text, tables, and lightweight document cues.
"""

from __future__ import annotations

import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

import pdfplumber
from pypdf import PdfReader
from pypdf.errors import FileNotDecryptedError, PdfReadError, PdfStreamError

from app.config import settings
from app.extraction.exceptions import (
    CorruptedPDFError,
    EmptyPDFError,
    EncryptedPDFError,
    InvalidPDFError,
    PDFTooLargeError,
)
from app.extraction.ocr_extractor import (
    OCRExtractor,
    document_extraction_method,
    is_text_sufficient,
    resolve_extraction_method,
)
from app.extraction.table_extractor import TableExtractor

PDF_MAGIC = b"%PDF-"

MONTHS = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Sept",
    "Oct",
    "Nov",
    "Dec",
)

DATE_PATTERNS = (
    re.compile(r"\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b"),
    re.compile(r"\b(\d{4}-\d{2}-\d{2})\b"),
    re.compile(
        r"\b(\d{1,2}\s+(?:" + "|".join(MONTHS) + r")[a-z]*\.?,?\s+\d{4})\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"\b((?:" + "|".join(MONTHS) + r")[a-z]*\.?\s+\d{1,2},?\s+\d{4})\b",
        re.IGNORECASE,
    ),
    re.compile(r"\b(as at\s+\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b", re.IGNORECASE),
    re.compile(r"\b(for the year ended\s+[^\n]{0,40}\d{4})\b", re.IGNORECASE),
)

NUMBER_PATTERN = re.compile(
    r"(?<![\w.])"
    r"(?:₹|Rs\.?|INR)?\s*"
    r"(\(?)(?=.)"
    r"("
    r"-?\d{1,3}(?:,\d{2})+(?:,\d{3})?(?:\.\d+)?"  # Indian grouping
    r"|-?\d{1,3}(?:,\d{3})+(?:\.\d+)?"  # Western grouping
    r"|-?\d+(?:\.\d+)?"
    r")"
    r"(\)?)"
    r"(?!\d)"
)

HEADING_CANDIDATE = re.compile(
    r"^(?:(?:note|part|schedule|annexure)\s+)?(?:\d+[.)]|[IVXLC]+\.?|[A-Z]\.)?\s*[A-Z].{2,90}$"
)

FINANCIAL_SECTIONS = (
    "balance sheet",
    "statement of profit and loss",
    "profit and loss",
    "statement of profit",
    "cash flow statement",
    "cash flow",
    "notes to accounts",
    "notes to the financial statements",
    "significant accounting policies",
    "share capital",
    "reserves and surplus",
    "shareholders' funds",
    "shareholders funds",
    "non-current liabilities",
    "current liabilities",
    "non-current assets",
    "current assets",
    "revenue from operations",
    "other income",
    "expenses",
    "tax expense",
    "earnings per share",
    "earnings per equity share",
    "property, plant and equipment",
    "inventories",
    "trade receivables",
    "trade payables",
    "borrowings",
    "contingent liabilities",
)


class NativeTextExtractor:
    """Digital PDF text extraction using the same page-text interface as OCR."""

    method = "text"

    def extract_page_text(self, pdf_path: Path, page_number: int) -> str:
        path = Path(pdf_path)
        text = ""
        try:
            with pdfplumber.open(path) as pdf:
                if 1 <= page_number <= len(pdf.pages):
                    text = pdf.pages[page_number - 1].extract_text() or ""
        except Exception:
            text = ""
        if text.strip():
            return text
        try:
            reader = PdfReader(str(path), strict=False)
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception:
                    return text
            if 1 <= page_number <= len(reader.pages):
                return reader.pages[page_number - 1].extract_text() or ""
        except Exception:
            return text
        return text


class PDFExtractor:
    """Validate an uploaded PDF and extract page-level content."""

    def __init__(
        self,
        *,
        table_extractor: TableExtractor | None = None,
        ocr_engine: Any | None = None,
        native_extractor: NativeTextExtractor | None = None,
        upload_dir: Path | None = None,
        max_upload_bytes: int | None = None,
        enable_ocr: bool | None = None,
    ) -> None:
        self.table_extractor = table_extractor or TableExtractor()
        self.native_extractor = native_extractor or NativeTextExtractor()
        self.upload_dir = Path(upload_dir) if upload_dir else settings.upload_dir
        self.max_upload_bytes = (
            max_upload_bytes if max_upload_bytes is not None else settings.max_upload_bytes
        )
        self.upload_dir.mkdir(parents=True, exist_ok=True)
        self.enable_ocr = settings.ocr_enabled if enable_ocr is None else enable_ocr
        self.ocr_engine = self._resolve_ocr_engine(ocr_engine)

    def _resolve_ocr_engine(self, ocr_engine: Any | None) -> Any | None:
        if ocr_engine is not None:
            return ocr_engine
        if not self.enable_ocr:
            return None
        backend = OCRExtractor(language=settings.ocr_language, dpi=settings.ocr_dpi)
        return backend if backend.is_available() else None

    def save_upload(self, filename: str, content: bytes) -> Path:
        """Validate PDF bytes and store them under the upload directory."""
        self.validate_pdf_bytes(content, filename=filename)
        safe_name = _safe_filename(filename)
        stored = self.upload_dir / f"{uuid.uuid4().hex}_{safe_name}"
        stored.write_bytes(content)
        return stored

    def validate_pdf_bytes(self, content: bytes, filename: str | None = None) -> None:
        if not content:
            raise EmptyPDFError("Uploaded file is empty.")
        if len(content) > self.max_upload_bytes:
            raise PDFTooLargeError(
                f"PDF exceeds the maximum upload size of {self.max_upload_bytes} bytes."
            )
        if filename and Path(filename).suffix.lower() not in {"", ".pdf"}:
            raise InvalidPDFError("Only PDF files are accepted.")
        if not content.lstrip().startswith(PDF_MAGIC):
            raise InvalidPDFError("Uploaded file is not a valid PDF.")

    def extract(self, file_path: str | Path, *, original_filename: str | None = None) -> dict[str, Any]:
        """Extract structured content from a PDF on disk."""
        path = Path(file_path)
        if not path.exists():
            raise InvalidPDFError(f"PDF not found: {path}")
        content = path.read_bytes()
        self.validate_pdf_bytes(content, filename=original_filename or path.name)
        reader = self._open_pypdf(path)
        page_count = len(reader.pages)
        if page_count == 0:
            raise EmptyPDFError("PDF contains no pages.")

        pages: list[dict[str, Any]] = []
        document_warnings: list[str] = []
        ocr_page_count = 0

        try:
            with pdfplumber.open(path) as plumber:
                plumber_pages = plumber.pages
                for page_number in range(1, page_count + 1):
                    plumber_page = plumber_pages[page_number - 1] if page_number <= len(plumber_pages) else None
                    page_payload = self._extract_page(
                        path=path,
                        reader=reader,
                        plumber_page=plumber_page,
                        page_number=page_number,
                    )
                    if page_payload["extraction_method"] in {"ocr", "mixed"}:
                        ocr_page_count += 1
                    pages.append(page_payload)
        except (PdfReadError, PdfStreamError, OSError) as exc:
            raise CorruptedPDFError(f"PDF could not be opened: {exc}") from exc

        extracted_chars = sum(len((page.get("text") or "").strip()) for page in pages)
        has_tables = any(page.get("tables") for page in pages)
        has_images = any(page.get("image_only") or page.get("image_count", 0) for page in pages)
        if extracted_chars == 0 and not has_tables and not has_images:
            raise EmptyPDFError("PDF has no extractable text, tables, or images.")
        if extracted_chars == 0 and has_images and ocr_page_count == 0:
            document_warnings.append(
                "PDF appears to be scanned or image-only and OCR is unavailable. "
                "Install Tesseract to enable OCR fallback."
            )
        elif any(page.get("needs_ocr") for page in pages):
            document_warnings.append(
                "Some pages had little native text; OCR fallback was used or is still needed."
            )

        filename = original_filename or path.name
        return {
            "document": {
                "filename": filename,
                "stored_path": str(path),
                "pages": page_count,
                "file_size_bytes": path.stat().st_size,
                "extraction_method": document_extraction_method(pages),
                "needs_ocr": any(page.get("needs_ocr") for page in pages),
                "ocr_engine_configured": self.ocr_engine is not None,
                "scanned_page_count": ocr_page_count,
                "warnings": document_warnings,
            },
            "pages": pages,
        }

    def extract_upload(self, filename: str, content: bytes) -> dict[str, Any]:
        """Save an upload, then extract structured content from the stored file."""
        stored = self.save_upload(filename, content)
        return self.extract(stored, original_filename=filename)

    def _open_pypdf(self, path: Path) -> PdfReader:
        try:
            reader = PdfReader(str(path), strict=False)
        except FileNotDecryptedError as exc:
            raise EncryptedPDFError("PDF is password protected.") from exc
        except (PdfReadError, PdfStreamError) as exc:
            raise CorruptedPDFError(f"PDF is corrupted or unreadable: {exc}") from exc
        except Exception as exc:
            message = str(exc).lower()
            if "password" in message or "encrypted" in message or "decrypt" in message:
                raise EncryptedPDFError("PDF is password protected.") from exc
            raise CorruptedPDFError(f"PDF could not be read: {exc}") from exc

        if reader.is_encrypted:
            unlocked = False
            try:
                result = reader.decrypt("")
                unlocked = bool(result)
            except Exception:
                unlocked = False
            if not unlocked:
                try:
                    _ = reader.pages[0]
                    unlocked = True
                except Exception as exc:
                    raise EncryptedPDFError("PDF is password protected.") from exc
            if not unlocked:
                raise EncryptedPDFError("PDF is password protected.")
        return reader

    def _extract_page(
        self,
        *,
        path: Path,
        reader: PdfReader,
        plumber_page: Any,
        page_number: int,
    ) -> dict[str, Any]:
        errors: list[str] = []
        text = ""
        image_count = 0
        tables: list[dict[str, Any]] = []

        if plumber_page is not None:
            try:
                text = plumber_page.extract_text() or ""
            except Exception as exc:
                errors.append(f"pdfplumber text extraction failed: {exc}")
            try:
                image_count = len(plumber_page.images or [])
            except Exception:
                image_count = 0
            try:
                tables = self.table_extractor.extract_page_tables(plumber_page, page_number)
            except Exception as exc:
                errors.append(f"table extraction failed: {exc}")

        if not text.strip():
            try:
                text = self.native_extractor.extract_page_text(path, page_number) or ""
            except Exception as exc:
                errors.append(f"native text extraction failed: {exc}")

        text = _normalize_text(text)
        native_text = text
        native_sufficient = is_text_sufficient(native_text) or bool(tables)
        used_ocr = False
        ocr_error = None

        if not native_sufficient and self.ocr_engine is not None:
            try:
                ocr_text = _read_ocr_page(self.ocr_engine, path, page_number)
            except Exception as exc:
                ocr_text = ""
                ocr_error = f"OCR backend failed: {exc}"
                errors.append(ocr_error)
            if ocr_text.strip():
                used_ocr = True
                normalized_ocr = _normalize_text(ocr_text)
                if native_text and normalized_ocr != native_text:
                    text = _normalize_text(f"{native_text}\n{normalized_ocr}")
                else:
                    text = normalized_ocr

        image_only = (not is_text_sufficient(native_text)) and image_count > 0
        extraction_method = resolve_extraction_method(
            native_sufficient=native_sufficient,
            used_ocr=used_ocr,
            had_native_signal=bool(native_text) or bool(tables),
        )
        needs_ocr = (not native_sufficient) and not is_text_sufficient(text)

        return {
            "page_number": page_number,
            "text": text,
            "tables": tables,
            "extraction_method": extraction_method,
            "numbers": extract_numbers(text),
            "headings": extract_headings(text),
            "dates": extract_dates(text),
            "sections": extract_financial_sections(text, page_number),
            "image_count": image_count,
            "image_only": image_only,
            "needs_ocr": needs_ocr,
            "unreadable": bool(errors) and not text.strip() and not tables,
            "errors": errors,
        }


def _read_ocr_page(engine: Any, path: Path, page_number: int) -> str:
    if hasattr(engine, "extract_page_text"):
        return engine.extract_page_text(path, page_number) or ""
    if hasattr(engine, "recognize_page"):
        return engine.recognize_page(path, page_number) or ""
    raise TypeError("OCR engine must implement extract_page_text or recognize_page.")


def _safe_filename(filename: str) -> str:
    name = Path(filename or "upload.pdf").name
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._")
    if not name.lower().endswith(".pdf"):
        name = f"{name or 'upload'}.pdf"
    return name[:180]


def _normalize_text(text: str) -> str:
    cleaned = text.replace("\xa0", " ")
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def extract_numbers(text: str) -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    seen: set[tuple[str, int]] = set()
    for match in NUMBER_PATTERN.finditer(text):
        raw = "".join(part for part in match.groups() if part)
        raw = raw.strip()
        if not raw or raw in {".", "-", "(", ")"}:
            continue
        key = (raw, match.start())
        if key in seen:
            continue
        seen.add(key)
        found.append(
            {
                "raw": match.group(0).strip(),
                "value": _parse_number(raw),
                "start": match.start(),
                "end": match.end(),
            }
        )
    return found


def extract_dates(text: str) -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    seen: set[str] = set()
    for pattern in DATE_PATTERNS:
        for match in pattern.finditer(text):
            raw = match.group(1).strip()
            key = raw.lower()
            if key in seen:
                continue
            seen.add(key)
            found.append(
                {
                    "raw": raw,
                    "iso": _to_iso_date(raw),
                    "start": match.start(),
                    "end": match.end(),
                }
            )
    return found


def extract_headings(text: str) -> list[dict[str, Any]]:
    headings: list[dict[str, Any]] = []
    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line or len(line) < 3 or len(line) > 100:
            continue
        letters = re.sub(r"[^A-Za-z]", "", line)
        if len(letters) < 3:
            continue
        uppercase_ratio = sum(1 for char in letters if char.isupper()) / len(letters)
        looks_like_heading = (
            bool(HEADING_CANDIDATE.match(line))
            or uppercase_ratio >= 0.7
            or (line.endswith(":") and len(line) <= 80)
        )
        if not looks_like_heading:
            continue
        headings.append({"text": line.rstrip(":"), "line_number": line_number})
    return headings


def extract_financial_sections(text: str, page_number: int) -> list[dict[str, Any]]:
    lowered = text.lower()
    sections: list[dict[str, Any]] = []
    for label in FINANCIAL_SECTIONS:
        idx = lowered.find(label)
        if idx < 0:
            continue
        snippet = text[max(0, idx - 40) : idx + len(label) + 80]
        sections.append(
            {
                "name": label,
                "page_number": page_number,
                "snippet": " ".join(snippet.split()),
            }
        )
    return sections


def _parse_number(raw: str) -> float | None:
    negative = raw.startswith("(") and raw.endswith(")")
    compact = raw.replace("(", "").replace(")", "").replace(",", "").replace("₹", "")
    compact = re.sub(r"^(Rs\.?|INR)", "", compact, flags=re.IGNORECASE).strip()
    try:
        value = float(compact)
    except ValueError:
        return None
    return -abs(value) if negative or compact.startswith("-") else value


def _to_iso_date(raw: str) -> str | None:
    cleaned = re.sub(r"^(as at|for the year ended)\s+", "", raw, flags=re.IGNORECASE).strip()
    cleaned = cleaned.replace(",", "")
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y", "%Y-%m-%d", "%d %B %Y", "%d %b %Y", "%B %d %Y", "%b %d %Y"):
        try:
            return datetime.strptime(cleaned, fmt).date().isoformat()
        except ValueError:
            continue
    return None
