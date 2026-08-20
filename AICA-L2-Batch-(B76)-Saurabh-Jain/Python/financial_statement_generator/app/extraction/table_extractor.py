"""Table detection and extraction for uploaded financial statement PDFs.

This module is independent of Schedule III mapping. It only returns raw
tabular grids found on a page.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pdfplumber

from app.extraction.exceptions import CorruptedPDFError, InvalidPDFError, PDFExtractionError

_LINE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "intersection_tolerance": 5,
}

_TEXT_SETTINGS = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "min_words_vertical": 3,
    "min_words_horizontal": 1,
}


class TableExtractor:
    """Detect and extract tabular data from financial statement PDFs."""

    def extract_tables(self, file_path: str | Path) -> list[dict[str, Any]]:
        """Extract tables from every page of a PDF."""
        path = Path(file_path)
        if not path.exists():
            raise InvalidPDFError(f"PDF not found: {path}")

        tables: list[dict[str, Any]] = []
        try:
            with pdfplumber.open(path) as pdf:
                if not pdf.pages:
                    return []
                for index, page in enumerate(pdf.pages, start=1):
                    tables.extend(self.extract_page_tables(page, page_number=index))
        except PDFExtractionError:
            raise
        except Exception as exc:
            raise CorruptedPDFError(f"Could not read tables from PDF: {exc}") from exc
        return tables

    def extract_page_tables(self, page: Any, page_number: int) -> list[dict[str, Any]]:
        """Extract tables from a single pdfplumber page."""
        found = self._find_tables(page)
        tables: list[dict[str, Any]] = []
        for table_index, table in enumerate(found):
            try:
                rows = table.extract() or []
            except Exception:
                continue
            cleaned = [_clean_row(row) for row in rows]
            cleaned = [row for row in cleaned if any(cell for cell in row)]
            if not cleaned:
                continue
            width = max(len(row) for row in cleaned)
            normalized = [row + [""] * (width - len(row)) for row in cleaned]
            bbox = getattr(table, "bbox", None)
            tables.append(
                {
                    "page_number": page_number,
                    "table_index": table_index,
                    "bbox": list(bbox) if bbox else None,
                    "row_count": len(normalized),
                    "column_count": width,
                    "rows": normalized,
                }
            )
        return tables

    def _find_tables(self, page: Any) -> list[Any]:
        found = []
        try:
            found = page.find_tables(table_settings=_LINE_SETTINGS) or []
        except Exception:
            found = []
        if found:
            return found
        try:
            found = page.find_tables(table_settings=_TEXT_SETTINGS) or []
        except Exception:
            found = []
        return found


def _clean_row(row: list[Any] | None) -> list[str]:
    if not row:
        return []
    cleaned: list[str] = []
    for cell in row:
        if cell is None:
            cleaned.append("")
            continue
        text = " ".join(str(cell).split())
        cleaned.append(text)
    return cleaned
