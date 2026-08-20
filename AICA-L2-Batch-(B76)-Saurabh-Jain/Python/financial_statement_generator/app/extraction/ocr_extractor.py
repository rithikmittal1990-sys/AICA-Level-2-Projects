"""OCR fallback for scanned or image-only PDF pages.

Native digital-text extraction and OCR both implement ``PageTextExtractor``.
The pipeline tries native text first and only renders a page for OCR when that
text is not usable.
"""

from __future__ import annotations

import io
import shutil
from pathlib import Path
from typing import Any, Protocol

from PIL import Image

MIN_USABLE_TEXT_CHARS = 40
MIN_USABLE_LETTERS = 15


class PageTextExtractor(Protocol):
    """Shared interface for native PDF text extraction and OCR."""

    def extract_page_text(self, pdf_path: Path, page_number: int) -> str:
        """Return text for one 1-based page number."""


def is_text_sufficient(text: str) -> bool:
    """Return True when native extraction produced usable page text."""
    stripped = (text or "").strip()
    if len(stripped) < MIN_USABLE_TEXT_CHARS:
        return False
    letters = sum(1 for char in stripped if char.isalpha())
    return letters >= MIN_USABLE_LETTERS


class OCRExtractor:
    """Render a PDF page to an image and recognize text with Tesseract.

    Rendering prefers PyMuPDF and falls back to pdf2image. Tesseract is used
    only when native text is insufficient; this class is never required if
    digital text extraction already succeeded.
    """

    method = "ocr"

    def __init__(self, *, language: str = "eng", dpi: int = 200) -> None:
        self.language = language
        self.dpi = dpi

    def extract_page_text(self, pdf_path: Path, page_number: int) -> str:
        """Same interface as native text extraction."""
        return self.recognize_page(pdf_path, page_number)

    def recognize_page(self, pdf_path: Path, page_number: int) -> str:
        """OCR a single 1-based page and return recognized text."""
        if not self.is_available():
            return ""
        image = self.render_page(Path(pdf_path), page_number)
        return self._image_to_text(image)

    def render_page(self, pdf_path: Path, page_number: int) -> Image.Image:
        """Rasterize one PDF page. PyMuPDF is preferred; pdf2image is fallback."""
        try:
            return self._render_with_pymupdf(pdf_path, page_number)
        except Exception:
            return self._render_with_pdf2image(pdf_path, page_number)

    def is_available(self) -> bool:
        """True when Tesseract and at least one page renderer can be used."""
        if shutil.which("tesseract") is None:
            return False
        try:
            import pytesseract

            pytesseract.get_tesseract_version()
        except Exception:
            return False
        return self._renderer_available()

    def _image_to_text(self, image: Image.Image) -> str:
        import pytesseract

        if image.mode not in {"RGB", "L"}:
            image = image.convert("RGB")
        return pytesseract.image_to_string(image, lang=self.language, config="--psm 6") or ""

    def _renderer_available(self) -> bool:
        try:
            import fitz  # noqa: F401

            return True
        except ImportError:
            try:
                import pdf2image  # noqa: F401

                return True
            except ImportError:
                return False

    def _render_with_pymupdf(self, pdf_path: Path, page_number: int) -> Image.Image:
        import fitz

        document = fitz.open(pdf_path)
        try:
            if page_number < 1 or page_number > document.page_count:
                raise IndexError(f"Page {page_number} is out of range.")
            page = document.load_page(page_number - 1)
            scale = self.dpi / 72
            pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False)
            return Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
        finally:
            document.close()

    def _render_with_pdf2image(self, pdf_path: Path, page_number: int) -> Image.Image:
        from pdf2image import convert_from_path

        images = convert_from_path(
            str(pdf_path),
            dpi=self.dpi,
            first_page=page_number,
            last_page=page_number,
        )
        if not images:
            raise RuntimeError(f"pdf2image did not render page {page_number}.")
        return images[0].convert("RGB")


def resolve_extraction_method(*, native_sufficient: bool, used_ocr: bool, had_native_signal: bool) -> str:
    """Classify how a page was extracted."""
    if native_sufficient and not used_ocr:
        return "text"
    if used_ocr and had_native_signal:
        return "mixed"
    if used_ocr:
        return "ocr"
    return "text"


def document_extraction_method(pages: list[dict[str, Any]]) -> str:
    methods = {page.get("extraction_method") or "text" for page in pages}
    if methods == {"text"}:
        return "text"
    if methods == {"ocr"}:
        return "ocr"
    return "mixed"
