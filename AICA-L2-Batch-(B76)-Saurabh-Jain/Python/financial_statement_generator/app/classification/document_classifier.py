"""Classify extracted PDF pages into financial-statement sections and values.

Native extraction runs first. This layer only reads that payload: it never
fills amounts that were not found, and low-confidence matches are stored as
null with a warning.
"""

from __future__ import annotations

import re
from difflib import SequenceMatcher
from typing import Any, Iterable

from app.extraction.pdf_extractor import extract_dates, extract_numbers
from app.mapping.field_mapping import lookup_mapping
from app.models.financial_models import (
    BalanceSheet,
    Borrowings,
    CashAndCashEquivalents,
    CashFlow,
    ClassifiedFinancialData,
    ClassificationWarning,
    CompanyInfo,
    CWIP,
    EPS,
    Expenses,
    FinancialBlock,
    FinancialNotes,
    IdentifiedDocumentSection,
    Inventory,
    Investments,
    LineItem,
    NoteExtract,
    OtherDisclosures,
    PPE,
    ProfitAndLoss,
    Ratios,
    ReservesAndSurplus,
    Revenue,
    ShareCapital,
    SourcedValue,
    TaxExpense,
    TradePayables,
    TradeReceivables,
    unextracted_value,
)

MIN_CONFIDENCE = 0.70

CIN_RE = re.compile(r"\b([UL]\d{5}[A-Z]{2}\d{4}[A-Z]{3}\d{6})\b")
COMPANY_NAME_LABEL_RE = re.compile(
    r"(?:name of (?:the )?company|company name)\s*[:\-]\s*(.+)$",
    re.IGNORECASE,
)
COMPANY_NAME_RE = re.compile(
    r"^(.+?\b(?:Private Limited|Pvt\.?\s*Ltd\.?|Limited|Ltd\.?))\b",
    re.IGNORECASE,
)
REGISTERED_OFFICE_RE = re.compile(
    r"(?:registered office|regd\.?\s*office)\s*[:\-]\s*(.+)$",
    re.IGNORECASE,
)
UNIT_RE = re.compile(
    r"(?:(?:figures|amounts?|rupees?|rs\.?|inr|₹)\s+)?in\s+"
    r"(crores?|lakhs?|lacs?|thousands?|millions?|absolute|units?)",
    re.IGNORECASE,
)
CURRENCY_RE = re.compile(r"\b(INR|Rs\.?|₹|Indian Rupees?)\b", re.IGNORECASE)
NATURE_RE = re.compile(
    r"\b(private limited company|public limited company|listed company|unlisted company)\b",
    re.IGNORECASE,
)
NOTE_HEADING_RE = re.compile(
    r"^(?:note\s*)?([A-Z]{1,2}|\d+[A-Z]?)\.?\s+([A-Za-z].{3,80})$",
    re.IGNORECASE,
)
NOTE_NO_RE = re.compile(r"^(?:\d{1,2}[A-Za-z]?|[A-Z]{1,2})$")
BLANK_AMOUNT_TOKENS = {"", "-", "—", "–", ".", "n.a.", "na", "n/a", "none"}

# Captions taken from ICAI Division I Schedule III (Annexure A / notes).
# Used when the reference mapper is unavailable; labels are not invented.
SECTION_CATALOG: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    (
        "company_information",
        "Company information",
        ("name of the company", "corporate identity number", "cin:", "registered office"),
    ),
    (
        "balance_sheet",
        "Balance Sheet",
        ("balance sheet", "equity and liabilities"),
    ),
    (
        "profit_and_loss",
        "Statement of Profit and Loss",
        ("statement of profit and loss", "profit and loss"),
    ),
    (
        "cash_flow",
        "Cash Flow Statement",
        ("cash flow statement", "statement of cash flows", "cash flow from operating"),
    ),
    (
        "notes_to_accounts",
        "Notes to Accounts",
        (
            "notes to accounts",
            "notes to the financial statements",
            "significant accounting policies",
        ),
    ),
    ("share_capital", "Share Capital", ("share capital",)),
    ("reserves_and_surplus", "Reserves and Surplus", ("reserves and surplus",)),
    (
        "borrowings",
        "Borrowings",
        ("long-term borrowings", "short-term borrowings", "borrowings"),
    ),
    (
        "ppe",
        "Property, Plant and Equipment",
        ("property, plant and equipment",),
    ),
    (
        "cwip",
        "Capital work-in-progress",
        ("capital work-in-progress", "capital work in progress"),
    ),
    (
        "investments",
        "Investments",
        ("non-current investments", "current investments"),
    ),
    ("trade_receivables", "Trade Receivables", ("trade receivables",)),
    ("trade_payables", "Trade Payables", ("trade payables",)),
    ("inventory", "Inventory", ("inventories",)),
    (
        "cash_and_cash_equivalents",
        "Cash and Cash Equivalents",
        ("cash and cash equivalents",),
    ),
    ("revenue", "Revenue", ("revenue from operations",)),
    (
        "expenses",
        "Expenses",
        (
            "employee benefits expense",
            "finance costs",
            "cost of materials consumed",
            "purchases of stock-in-trade",
            "other expenses",
        ),
    ),
    ("tax", "Tax", ("tax expense", "current tax", "deferred tax")),
    (
        "eps",
        "EPS",
        ("earnings per equity share", "earnings per share"),
    ),
    (
        "other_disclosures",
        "Other disclosures",
        (
            "additional regulatory information",
            "contingent liabilities and commitments",
            "other disclosures",
        ),
    ),
    (
        "ratios",
        "Ratios",
        (
            "analytical ratios",
            "current ratio",
            "debt-equity ratio",
            "following ratios to be disclosed",
        ),
    ),
)

SECTION_LABELS: dict[str, tuple[str, ...]] = {
    "balance_sheet": (
        "Share capital",
        "Reserves and surplus",
        "Money received against share warrants",
        "Share application money pending allotment",
        "Long-term borrowings",
        "Deferred tax liabilities (Net)",
        "Other Long term liabilities",
        "Long-term provisions",
        "Short-term borrowings",
        "Trade payables",
        "Other current liabilities",
        "Short-term provisions",
        "Property, Plant and Equipment",
        "Intangible assets",
        "Capital work-in-progress",
        "Intangible assets under development",
        "Non-current investments",
        "Deferred tax assets (net)",
        "Long-term loans and advances",
        "Other non-current assets",
        "Current investments",
        "Inventories",
        "Trade receivables",
        "Cash and cash equivalents",
        "Short-term loans and advances",
        "Other current assets",
    ),
    "profit_and_loss": (
        "Revenue from operations",
        "Other income",
        "Total Income",
        "Cost of materials consumed",
        "Purchases of Stock-in-Trade",
        "Changes in inventories of finished goods, work-in-progress and Stock-in-Trade",
        "Employee benefits expense",
        "Finance costs",
        "Depreciation and amortization expense",
        "Other expenses",
        "Total expenses",
        "Exceptional items",
        "Extraordinary Items",
        "Tax expense",
        "Current tax",
        "Deferred tax",
        "Profit before tax",
        "Profit/ (Loss)",
        "Earnings per equity share",
        "Basic",
        "Diluted",
    ),
    "cash_flow": (
        "Cash flow from operating activities",
        "Cash flow from investing activities",
        "Cash flow from financing activities",
        "Net increase in cash and cash equivalents",
        "Cash and cash equivalents at the beginning of the year",
        "Cash and cash equivalents at the end of the year",
    ),
    "share_capital": (
        "Share capital",
        "Authorised",
        "Issued",
        "Subscribed",
        "Subscribed and fully paid up",
        "Paid-up",
    ),
    "reserves_and_surplus": ("Reserves and surplus", "Surplus", "Securities Premium"),
    "borrowings": ("Long-term borrowings", "Short-term borrowings", "Borrowings"),
    "ppe": ("Property, Plant and Equipment",),
    "cwip": ("Capital work-in-progress",),
    "investments": ("Non-current investments", "Current investments"),
    "trade_receivables": ("Trade receivables",),
    "trade_payables": ("Trade payables",),
    "inventory": ("Inventories",),
    "cash_and_cash_equivalents": ("Cash and cash equivalents",),
    "revenue": ("Revenue from operations", "Other income"),
    "expenses": (
        "Cost of materials consumed",
        "Purchases of Stock-in-Trade",
        "Changes in inventories of finished goods, work-in-progress and Stock-in-Trade",
        "Employee benefits expense",
        "Finance costs",
        "Depreciation and amortization expense",
        "Other expenses",
        "Total expenses",
    ),
    "tax": ("Tax expense", "Current tax", "Deferred tax"),
    "eps": ("Earnings per equity share", "Basic", "Diluted"),
    "other_disclosures": (
        "Contingent liabilities and commitments (to the extent not provided for)",
        "Additional Regulatory Information",
    ),
    "ratios": (
        "Current Ratio",
        "Debt-Equity Ratio",
        "Debt Service Coverage Ratio",
        "Return on Equity Ratio",
        "Inventory turnover ratio",
        "Trade Receivables turnover ratio",
        "Trade payables turnover ratio",
        "Net capital turnover ratio",
        "Net profit ratio",
        "Return on Capital employed",
        "Return on investment",
    ),
}

KEY_FIELDS: dict[str, tuple[str, ...]] = {
    "share_capital": ("Share capital",),
    "reserves_and_surplus": ("Reserves and surplus",),
    "borrowings": ("Long-term borrowings", "Short-term borrowings"),
    "ppe": ("Property, Plant and Equipment",),
    "inventory": ("Inventories",),
    "trade_receivables": ("Trade receivables",),
    "trade_payables": ("Trade payables",),
    "cash_and_cash_equivalents": ("Cash and cash equivalents",),
    "revenue": ("Revenue from operations",),
    "tax": ("Tax expense",),
    "eps": ("Basic",),
}


class DocumentClassifier:
    """Identify statement sections and extract sourced financial values."""

    def __init__(
        self,
        *,
        mappings: dict[str, dict[str, Any]] | None = None,
        mapper: Any | None = None,
        min_confidence: float = MIN_CONFIDENCE,
    ) -> None:
        self._injected_mappings = mappings
        self._mapper = mapper
        self._mappings: dict[str, dict[str, Any]] | None = mappings
        self.min_confidence = min_confidence

    def classify(self, content: dict[str, Any] | list[dict[str, Any]]) -> ClassifiedFinancialData:
        """Classify an extraction payload into sourced financial models."""
        pages = _pages_from_content(content)
        result = ClassifiedFinancialData()
        if not pages:
            result.warnings.append(
                ClassificationWarning(
                    code="empty_extraction",
                    message="No extracted pages were provided, so no values were classified.",
                )
            )
            return result

        result.company = self._extract_company_info(pages, result.warnings)
        result.sections = self.identify_sections(pages)
        sections_by_id = {section.section_id: section for section in result.sections}

        result.balance_sheet = self._fill_block(BalanceSheet(), "balance_sheet", pages, sections_by_id, result.warnings)
        result.profit_and_loss = self._fill_block(
            ProfitAndLoss(), "profit_and_loss", pages, sections_by_id, result.warnings
        )
        result.cash_flow = self._fill_block(CashFlow(), "cash_flow", pages, sections_by_id, result.warnings)
        result.share_capital = self._fill_share_capital(pages, sections_by_id, result.warnings)
        result.reserves_and_surplus = self._fill_block(
            ReservesAndSurplus(), "reserves_and_surplus", pages, sections_by_id, result.warnings
        )
        result.borrowings = self._fill_borrowings(pages, sections_by_id, result.warnings)
        result.ppe = self._fill_block(PPE(), "ppe", pages, sections_by_id, result.warnings)
        result.cwip = self._fill_block(CWIP(), "cwip", pages, sections_by_id, result.warnings)
        result.investments = self._fill_block(Investments(), "investments", pages, sections_by_id, result.warnings)
        result.inventory = self._fill_block(Inventory(), "inventory", pages, sections_by_id, result.warnings)
        result.trade_receivables = self._fill_block(
            TradeReceivables(), "trade_receivables", pages, sections_by_id, result.warnings
        )
        result.trade_payables = self._fill_block(
            TradePayables(), "trade_payables", pages, sections_by_id, result.warnings
        )
        result.cash_and_cash_equivalents = self._fill_block(
            CashAndCashEquivalents(), "cash_and_cash_equivalents", pages, sections_by_id, result.warnings
        )
        result.revenue = self._fill_block(Revenue(), "revenue", pages, sections_by_id, result.warnings)
        result.expenses = self._fill_block(Expenses(), "expenses", pages, sections_by_id, result.warnings)
        result.tax_expense = self._fill_tax(pages, sections_by_id, result.warnings)
        result.eps = self._fill_eps(pages, sections_by_id, result.warnings)
        result.other_disclosures = self._fill_block(
            OtherDisclosures(), "other_disclosures", pages, sections_by_id, result.warnings
        )
        result.ratios = self._fill_block(Ratios(), "ratios", pages, sections_by_id, result.warnings)
        result.notes_to_accounts = self._fill_notes(pages, sections_by_id, result.warnings)
        return result

    def identify_sections(self, pages: list[dict[str, Any]]) -> list[IdentifiedDocumentSection]:
        """Return every catalog section that appears in the extracted pages."""
        hits: dict[str, list[tuple[int, float, str]]] = {section_id: [] for section_id, _, _ in SECTION_CATALOG}
        for page in pages:
            page_number = int(page.get("page_number") or 0)
            if page_number < 1:
                continue
            for section_id, _title, keywords in SECTION_CATALOG:
                score, excerpt = _score_page(page, keywords)
                if score >= 0.60:
                    hits[section_id].append((page_number, score, excerpt or ""))

        sections: list[IdentifiedDocumentSection] = []
        for section_id, title, _keywords in SECTION_CATALOG:
            matches = hits[section_id]
            if not matches:
                continue
            page_numbers = sorted({item[0] for item in matches})
            best = max(matches, key=lambda item: item[1])
            sections.append(
                IdentifiedDocumentSection(
                    section_id=section_id,
                    title=title,
                    start_page=page_numbers[0],
                    end_page=page_numbers[-1],
                    pages=page_numbers,
                    confidence=round(best[1], 4),
                    excerpt=_clip(best[2]),
                )
            )
        return sections

    def _fill_block(
        self,
        block: FinancialBlock,
        section_id: str,
        pages: list[dict[str, Any]],
        sections_by_id: dict[str, IdentifiedDocumentSection],
        warnings: list[ClassificationWarning],
    ) -> FinancialBlock:
        section = sections_by_id.get(section_id)
        block.apply_section(section)
        target_pages = _pages_for_section(pages, section)
        labels = self._labels_for(section_id)
        block.line_items = self._extract_line_items(target_pages, labels, section_id, warnings)
        if block.identified:
            self._warn_missing_keys(section_id, block.line_items, section, warnings)
        return block

    def _fill_share_capital(
        self,
        pages: list[dict[str, Any]],
        sections_by_id: dict[str, IdentifiedDocumentSection],
        warnings: list[ClassificationWarning],
    ) -> ShareCapital:
        block = self._fill_block(ShareCapital(), "share_capital", pages, sections_by_id, warnings)
        assert isinstance(block, ShareCapital)
        block.authorised = _named_amount(block.line_items, ("authorised", "authorized"))
        block.issued = _named_amount(block.line_items, ("issued",))
        block.subscribed = _named_amount(block.line_items, ("subscribed",))
        block.paid_up = _named_amount(block.line_items, ("paid-up", "paid up", "share capital"))
        return block

    def _fill_borrowings(
        self,
        pages: list[dict[str, Any]],
        sections_by_id: dict[str, IdentifiedDocumentSection],
        warnings: list[ClassificationWarning],
    ) -> Borrowings:
        block = self._fill_block(Borrowings(), "borrowings", pages, sections_by_id, warnings)
        assert isinstance(block, Borrowings)
        block.long_term = _named_amount(block.line_items, ("long-term borrowings", "long term borrowings"))
        block.short_term = _named_amount(block.line_items, ("short-term borrowings", "short term borrowings"))
        return block

    def _fill_tax(
        self,
        pages: list[dict[str, Any]],
        sections_by_id: dict[str, IdentifiedDocumentSection],
        warnings: list[ClassificationWarning],
    ) -> TaxExpense:
        block = self._fill_block(TaxExpense(), "tax", pages, sections_by_id, warnings)
        assert isinstance(block, TaxExpense)
        block.current_tax = _named_amount(block.line_items, ("current tax",))
        block.deferred_tax = _named_amount(block.line_items, ("deferred tax",))
        return block

    def _fill_eps(
        self,
        pages: list[dict[str, Any]],
        sections_by_id: dict[str, IdentifiedDocumentSection],
        warnings: list[ClassificationWarning],
    ) -> EPS:
        block = self._fill_block(EPS(), "eps", pages, sections_by_id, warnings)
        assert isinstance(block, EPS)
        block.basic = _named_amount(block.line_items, ("basic",))
        block.diluted = _named_amount(block.line_items, ("diluted",))
        return block

    def _fill_notes(
        self,
        pages: list[dict[str, Any]],
        sections_by_id: dict[str, IdentifiedDocumentSection],
        warnings: list[ClassificationWarning],
    ) -> FinancialNotes:
        block = self._fill_block(FinancialNotes(), "notes_to_accounts", pages, sections_by_id, warnings)
        assert isinstance(block, FinancialNotes)
        notes: list[NoteExtract] = []
        for page in _pages_for_section(pages, sections_by_id.get("notes_to_accounts")) or pages:
            page_number = int(page.get("page_number") or 0)
            text = page.get("text") or ""
            for raw_line in text.splitlines():
                line = raw_line.strip()
                match = NOTE_HEADING_RE.match(line)
                if not match:
                    continue
                heading = match.group(2).strip().rstrip(":")
                if _norm(heading) in {"particulars", "note no"}:
                    continue
                notes.append(
                    NoteExtract(
                        note_id=match.group(1).upper(),
                        title=heading,
                        start_page=page_number,
                        pages=[page_number],
                        excerpt=_clip(line),
                    )
                )
        block.notes = notes
        return block

    def _extract_company_info(
        self,
        pages: list[dict[str, Any]],
        warnings: list[ClassificationWarning],
    ) -> CompanyInfo:
        info = CompanyInfo()
        early_pages = pages[:3]
        for page in early_pages:
            page_number = int(page.get("page_number") or 0)
            text = page.get("text") or ""
            if not info.cin.is_extracted():
                cin_match = CIN_RE.search(text)
                if cin_match:
                    info.cin = SourcedValue.from_extraction(
                        cin_match.group(1),
                        source_page=page_number,
                        source_text=cin_match.group(0),
                        confidence=0.95,
                    )
            if not info.company_name.is_extracted():
                labeled = COMPANY_NAME_LABEL_RE.search(text)
                if labeled:
                    info.company_name = SourcedValue.from_extraction(
                        labeled.group(1).strip(" ."),
                        source_page=page_number,
                        source_text=labeled.group(0),
                        confidence=0.93,
                    )
                else:
                    for raw_line in text.splitlines():
                        name_match = COMPANY_NAME_RE.match(raw_line.strip())
                        if name_match and "name of" not in raw_line.lower():
                            info.company_name = SourcedValue.from_extraction(
                                name_match.group(1).strip(" ."),
                                source_page=page_number,
                                source_text=raw_line.strip(),
                                confidence=0.82,
                            )
                            break
            if not info.registered_office.is_extracted():
                office = REGISTERED_OFFICE_RE.search(text)
                if office:
                    info.registered_office = SourcedValue.from_extraction(
                        office.group(1).strip(" ."),
                        source_page=page_number,
                        source_text=office.group(0),
                        confidence=0.9,
                    )
            if not info.unit.is_extracted():
                unit = UNIT_RE.search(text)
                if unit:
                    info.unit = SourcedValue.from_extraction(
                        unit.group(0).strip(),
                        source_page=page_number,
                        source_text=unit.group(0),
                        confidence=0.9,
                    )
            if not info.currency.is_extracted():
                currency = CURRENCY_RE.search(text)
                if currency:
                    info.currency = SourcedValue.from_extraction(
                        currency.group(1),
                        source_page=page_number,
                        source_text=currency.group(0),
                        confidence=0.85,
                    )
            if not info.nature_of_company.is_extracted():
                nature = NATURE_RE.search(text)
                if nature:
                    info.nature_of_company = SourcedValue.from_extraction(
                        nature.group(1),
                        source_page=page_number,
                        source_text=nature.group(0),
                        confidence=0.85,
                    )
            if not info.period_end.is_extracted() or not info.reporting_period.is_extracted():
                self._fill_periods(info, page, page_number)

        if not info.company_name.is_extracted():
            warnings.append(
                ClassificationWarning(
                    code="missing_value",
                    message="Company name could not be confidently extracted.",
                    section="company_information",
                    field="company_name",
                )
            )
        if not info.period_end.is_extracted():
            warnings.append(
                ClassificationWarning(
                    code="missing_value",
                    message="Reporting period end date could not be confidently extracted.",
                    section="company_information",
                    field="period_end",
                )
            )
        return info

    def _fill_periods(self, info: CompanyInfo, page: dict[str, Any], page_number: int) -> None:
        text = page.get("text") or ""
        dates = page.get("dates") or extract_dates(text)
        as_at = re.search(r"as at\s+([^\n]{6,40})", text, flags=re.IGNORECASE)
        year_ended = re.search(r"for the year ended\s+([^\n]{6,40})", text, flags=re.IGNORECASE)
        if as_at and not info.period_end.is_extracted():
            snippet = as_at.group(0)
            iso = dates[0]["iso"] if dates else None
            info.period_end = SourcedValue.from_extraction(
                iso or as_at.group(1).strip(),
                source_page=page_number,
                source_text=snippet,
                confidence=0.92 if iso else 0.8,
            )
            info.reporting_period = SourcedValue.from_extraction(
                snippet.strip(),
                source_page=page_number,
                source_text=snippet,
                confidence=0.9,
            )
        elif year_ended and not info.reporting_period.is_extracted():
            snippet = year_ended.group(0)
            info.reporting_period = SourcedValue.from_extraction(
                snippet.strip(),
                source_page=page_number,
                source_text=snippet,
                confidence=0.9,
            )
            if not info.period_end.is_extracted():
                iso = dates[0]["iso"] if dates else year_ended.group(1).strip()
                info.period_end = SourcedValue.from_extraction(
                    iso,
                    source_page=page_number,
                    source_text=snippet,
                    confidence=0.85,
                )
        elif dates and not info.period_end.is_extracted():
            first = dates[0]
            info.period_end = SourcedValue.from_extraction(
                first.get("iso") or first.get("raw"),
                source_page=page_number,
                source_text=first.get("raw"),
                confidence=0.72,
            )

    def _extract_line_items(
        self,
        pages: list[dict[str, Any]],
        expected_labels: list[tuple[str, str | None]],
        section_id: str,
        warnings: list[ClassificationWarning],
    ) -> list[LineItem]:
        found: dict[str, LineItem] = {}
        for page in pages:
            page_number = int(page.get("page_number") or 0)
            for table in page.get("tables") or []:
                rows = table.get("rows") or []
                for row in rows:
                    item = self._line_item_from_row(row, expected_labels, page_number)
                    if item is not None:
                        _keep_better(found, item)
            for raw_line in (page.get("text") or "").splitlines():
                item = self._line_item_from_text(raw_line, expected_labels, page_number)
                if item is not None:
                    _keep_better(found, item)
        items = list(found.values())
        for item in items:
            self._apply_confidence_gate(item, section_id, warnings)
        return [item for item in items if _has_any_attempt(item)]

    def _line_item_from_row(
        self,
        row: list[Any],
        expected_labels: list[tuple[str, str | None]],
        page_number: int,
    ) -> LineItem | None:
        cells = [_cell_text(cell) for cell in row]
        if not any(cells):
            return None
        if _looks_like_header(cells):
            return None
        label_cell = cells[0]
        matched = self._match_label(label_cell, expected_labels)
        if matched is None:
            return None
        label, mapping_code, label_score = matched
        note_no, amounts = _split_note_and_amounts(cells[1:])
        return self._build_line_item(
            label=label,
            mapping_code=mapping_code,
            label_score=label_score,
            note_no=note_no,
            amounts=amounts,
            page_number=page_number,
            source_text=" | ".join(cell for cell in cells if cell),
        )

    def _line_item_from_text(
        self,
        raw_line: str,
        expected_labels: list[tuple[str, str | None]],
        page_number: int,
    ) -> LineItem | None:
        line = " ".join(raw_line.split()).strip()
        if not line or _looks_like_header([line]):
            return None
        matched = self._match_label(line, expected_labels, prefix_ok=True)
        if matched is None:
            return None
        label, mapping_code, label_score = matched
        remainder = re.sub(re.escape(label), "", line, count=1, flags=re.IGNORECASE).strip(" :-")
        tokens = remainder.split()
        note_no = None
        amount_text = remainder
        if tokens and NOTE_NO_RE.match(tokens[0]) and len(tokens) > 1:
            note_no = tokens[0]
            amount_text = " ".join(tokens[1:])
        numbers = extract_numbers(amount_text)
        amounts = [item.get("value") for item in numbers if item.get("value") is not None]
        if note_no is None and not amounts:
            return None
        return self._build_line_item(
            label=label,
            mapping_code=mapping_code,
            label_score=label_score,
            note_no=note_no,
            amounts=amounts,
            page_number=page_number,
            source_text=line,
        )

    def _build_line_item(
        self,
        *,
        label: str,
        mapping_code: str | None,
        label_score: float,
        note_no: str | None,
        amounts: list[float | None],
        page_number: int,
        source_text: str,
    ) -> LineItem:
        cleaned: list[float | None] = list(amounts)
        has_current = bool(cleaned) and cleaned[0] is not None
        has_previous = len(cleaned) >= 2 and cleaned[1] is not None
        amount_score = 0.0
        if has_current and has_previous:
            amount_score = 0.95
        elif has_current or has_previous:
            amount_score = 0.88
        confidence = round(min(label_score, 1.0) * (0.55 + 0.45 * amount_score), 4)
        if not has_current and not has_previous:
            confidence = round(label_score * 0.4, 4)

        current = unextracted_value()
        previous = unextracted_value()
        if cleaned:
            current = SourcedValue.from_extraction(
                cleaned[0],
                source_page=page_number,
                source_text=source_text,
                confidence=confidence if cleaned[0] is not None else None,
            )
        if len(cleaned) >= 2:
            previous = SourcedValue.from_extraction(
                cleaned[1],
                source_page=page_number,
                source_text=source_text,
                confidence=confidence if cleaned[1] is not None else None,
            )
        note_value = unextracted_value()
        if note_no:
            note_value = SourcedValue.from_extraction(
                note_no,
                source_page=page_number,
                source_text=source_text,
                confidence=min(0.9, confidence + 0.05),
            )
        return LineItem(
            label=label,
            mapping_code=mapping_code,
            note_no=note_value,
            current_period=current,
            previous_period=previous,
        )

    def _apply_confidence_gate(
        self,
        item: LineItem,
        section_id: str,
        warnings: list[ClassificationWarning],
    ) -> None:
        for field_name in ("current_period", "previous_period"):
            sourced: SourcedValue = getattr(item, field_name)
            if sourced.value is None:
                continue
            if (sourced.confidence or 0) >= self.min_confidence:
                continue
            warnings.append(
                ClassificationWarning(
                    code="low_confidence",
                    message=(
                        f"'{item.label}' {field_name.replace('_', ' ')} was below the "
                        f"confidence threshold and was stored as null."
                    ),
                    section=section_id,
                    field=item.label,
                    source_page=sourced.source_page,
                )
            )
            setattr(
                item,
                field_name,
                SourcedValue.from_extraction(
                    None,
                    source_page=sourced.source_page,
                    source_text=sourced.source_text,
                    confidence=sourced.confidence,
                ),
            )

    def _warn_missing_keys(
        self,
        section_id: str,
        items: list[LineItem],
        section: IdentifiedDocumentSection | None,
        warnings: list[ClassificationWarning],
    ) -> None:
        present = {_norm(item.label) for item in items if item.current_period.is_extracted()}
        for label in KEY_FIELDS.get(section_id, ()):
            if any(_norm(label) in found or found in _norm(label) for found in present):
                continue
            warnings.append(
                ClassificationWarning(
                    code="missing_value",
                    message=f"Could not confidently extract '{label}'.",
                    section=section_id,
                    field=label,
                    source_page=section.start_page if section else None,
                )
            )

    def _match_label(
        self,
        extracted: str,
        expected_labels: list[tuple[str, str | None]],
        *,
        prefix_ok: bool = False,
    ) -> tuple[str, str | None, float] | None:
        best: tuple[str, str | None, float] | None = None
        for label, code in expected_labels:
            score = _label_score(extracted, label, prefix_ok=prefix_ok)
            if score < 0.75:
                continue
            if best is None or score > best[2]:
                mapping_code = code or self._mapping_code_for(label)
                best = (label, mapping_code, score)
        return best

    def _mapping_code_for(self, label: str) -> str | None:
        mappings = self._load_mappings()
        if not mappings:
            return None
        found = lookup_mapping(label, mappings)
        if found:
            return found.get("code")
        return None

    def _labels_for(self, section_id: str) -> list[tuple[str, str | None]]:
        labels = list(SECTION_LABELS.get(section_id, ()))
        mappings = self._load_mappings()
        if mappings:
            needle = _norm(section_id.replace("_", " "))
            for item in mappings.values():
                path = " ".join(item.get("path") or [])
                blob = f"{item.get('label', '')} {path}"
                if needle and needle in _norm(blob):
                    labels.append(item.get("label") or "")
        unique: list[tuple[str, str | None]] = []
        seen: set[str] = set()
        for label in labels:
            key = _norm(label)
            if not key or key in seen:
                continue
            seen.add(key)
            unique.append((label, self._mapping_code_for(label)))
        return unique

    def _load_mappings(self) -> dict[str, dict[str, Any]]:
        if self._mappings is not None:
            return self._mappings
        if self._injected_mappings is not None:
            self._mappings = self._injected_mappings
            return self._mappings
        try:
            from app.mapping.schedule_iii_mapper import ScheduleIIIMapper

            mapper = self._mapper or ScheduleIIIMapper()
            self._mappings = mapper.get_field_mappings()
        except Exception:
            self._mappings = {}
        return self._mappings


def _pages_from_content(content: dict[str, Any] | list[dict[str, Any]]) -> list[dict[str, Any]]:
    if isinstance(content, list):
        return content
    pages = content.get("pages")
    if isinstance(pages, list):
        return pages
    return []


def _pages_for_section(
    pages: list[dict[str, Any]],
    section: IdentifiedDocumentSection | None,
) -> list[dict[str, Any]]:
    if section is None:
        return pages
    wanted = set(section.pages)
    matched = [page for page in pages if int(page.get("page_number") or 0) in wanted]
    return matched or pages


def _score_page(page: dict[str, Any], keywords: tuple[str, ...]) -> tuple[float, str | None]:
    text = _page_blob(page)
    lowered = _norm(text)
    headings = [_norm(item.get("text") if isinstance(item, dict) else str(item)) for item in page.get("headings") or []]
    cues = [_norm(item.get("name") if isinstance(item, dict) else str(item)) for item in page.get("sections") or []]
    best_score = 0.0
    excerpt = None
    for keyword in keywords:
        needle = _norm(keyword)
        if not needle:
            continue
        if any(needle in heading or heading in needle for heading in headings if heading):
            score = 0.96
        elif any(needle in cue or cue in needle for cue in cues if cue):
            score = 0.9
        elif needle in lowered:
            # Unique longer captions score higher than generic short words.
            score = 0.84 if len(needle) >= 16 else 0.68
        else:
            continue
        if score > best_score:
            best_score = score
            excerpt = _excerpt(text, keyword)
    return best_score, excerpt


def _page_blob(page: dict[str, Any]) -> str:
    parts = [page.get("text") or ""]
    for table in page.get("tables") or []:
        for row in table.get("rows") or []:
            parts.append(" ".join(_cell_text(cell) for cell in row))
    return "\n".join(part for part in parts if part)


def _label_score(extracted: str, expected: str, *, prefix_ok: bool) -> float:
    left = _norm(_strip_note_prefix(extracted))
    right = _norm(expected)
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    if prefix_ok and left.startswith(right):
        return 0.97 if len(right) >= 8 else 0.88
    if right in left or left in right:
        ratio = min(len(left), len(right)) / max(len(left), len(right))
        if ratio >= 0.7:
            return 0.92
        if min(len(left), len(right)) >= 12:
            return 0.86
        return 0.62
    return SequenceMatcher(None, left, right).ratio()


def _strip_note_prefix(text: str) -> str:
    return re.sub(r"^(?:note\s*)?(?:\d+[A-Za-z]?|[A-Z]{1,2})\.?\s+", "", text.strip(), flags=re.IGNORECASE)


def _looks_like_header(cells: list[str]) -> bool:
    blob = _norm(" ".join(cells))
    if not blob:
        return True
    return any(token in blob for token in ("particulars", "note no", "as at", "year ended", "description"))


def _split_note_and_amounts(cells: list[str]) -> tuple[str | None, list[float | None]]:
    note_no = None
    remaining = list(cells)
    while remaining and remaining[0] == "":
        remaining.pop(0)
    if remaining and _is_note_reference(remaining[0], remaining[1:]):
        note_no = remaining.pop(0)
    amounts: list[float | None] = []
    for cell in remaining:
        if _cell_is_blank_amount(cell):
            amounts.append(None)
            continue
        parsed = parse_amount(cell)
        if parsed is not None:
            amounts.append(parsed)
    return note_no, amounts


def _is_note_reference(cell: str, following: list[str]) -> bool:
    if not NOTE_NO_RE.match(cell.strip()):
        return False
    later_amounts = [parse_amount(item) for item in following if not _cell_is_blank_amount(item)]
    later_blanks = any(_cell_is_blank_amount(item) for item in following)
    return bool(later_amounts) or later_blanks


def parse_amount(raw: str) -> float | None:
    """Parse an Indian/Western grouped amount. Blank or unknown tokens are null."""
    text = (raw or "").strip()
    if _cell_is_blank_amount(text):
        return None
    if text.lower() in {"nil", "nill"}:
        return 0.0
    negative = text.startswith("(") and text.endswith(")")
    compact = text.replace("(", "").replace(")", "").replace(",", "").replace("₹", "")
    compact = re.sub(r"^(rs\.?|inr)", "", compact, flags=re.IGNORECASE).strip()
    try:
        value = float(compact)
    except ValueError:
        return None
    if negative or compact.startswith("-"):
        return -abs(value)
    return value


def _cell_is_blank_amount(cell: str) -> bool:
    return cell.strip().lower() in BLANK_AMOUNT_TOKENS


def _cell_text(cell: Any) -> str:
    if cell is None:
        return ""
    return str(cell).strip()


def _named_amount(items: list[LineItem], needles: Iterable[str]) -> SourcedValue:
    normalized = [_norm(item) for item in needles]
    best: LineItem | None = None
    best_score = 0.0
    for item in items:
        label = _norm(item.label)
        for needle in normalized:
            if needle == label:
                score = 1.0
            elif needle in label or label in needle:
                score = 0.9
            else:
                continue
            if score > best_score and item.current_period.is_extracted():
                best = item
                best_score = score
    if best is None:
        return unextracted_value()
    return best.current_period


def _keep_better(found: dict[str, LineItem], item: LineItem) -> None:
    key = _norm(item.label)
    existing = found.get(key)
    if existing is None:
        found[key] = item
        return
    if _item_confidence(item) > _item_confidence(existing):
        found[key] = item


def _item_confidence(item: LineItem) -> float:
    current = item.current_period.confidence or 0.0
    previous = item.previous_period.confidence or 0.0
    return max(current, previous)


def _has_any_attempt(item: LineItem) -> bool:
    return bool(item.label) and (
        item.current_period.source_text
        or item.previous_period.source_text
        or item.current_period.is_extracted()
        or item.previous_period.is_extracted()
    )


def _excerpt(text: str, needle: str, radius: int = 90) -> str:
    blob = " ".join(text.split())
    idx = blob.lower().find(needle.lower())
    if idx < 0:
        return blob[: radius * 2]
    start = max(0, idx - radius)
    end = min(len(blob), idx + len(needle) + radius)
    return blob[start:end]


def _clip(text: str | None, limit: int = 240) -> str | None:
    if not text:
        return None
    compact = " ".join(text.split())
    if len(compact) <= limit:
        return compact
    return compact[: limit - 1].rstrip() + "…"


def _norm(value: str) -> str:
    cleaned = (value or "").replace("’", "'").replace("‘", "'")
    cleaned = cleaned.replace("–", "-").replace("—", "-")
    cleaned = re.sub(r"[^a-z0-9']+", " ", cleaned.lower())
    return " ".join(cleaned.split())
