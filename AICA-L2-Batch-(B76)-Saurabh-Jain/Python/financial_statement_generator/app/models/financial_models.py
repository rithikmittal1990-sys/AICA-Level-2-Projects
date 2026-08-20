"""Pydantic models for classified financial-statement data.

Extracted amounts always keep provenance. Missing or low-confidence values
are stored as null; callers must not treat an absent number as zero.
"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field


def unextracted_value() -> SourcedValue:
    """Return an empty sourced value. Never fills in a numeric default."""
    return SourcedValue(value=None, source_page=None, source_text=None, confidence=None)


class HealthResponse(BaseModel):
    """Response body for the health endpoint."""

    status: str
    service: str
    version: str


class SourcedValue(BaseModel):
    """One extracted scalar with the page and text it came from."""

    model_config = ConfigDict(extra="forbid")

    value: float | str | int | None = None
    source_page: int | None = None
    source_text: str | None = None
    confidence: float | None = None

    def is_extracted(self) -> bool:
        return self.value is not None

    @classmethod
    def from_extraction(
        cls,
        value: float | str | int | None,
        *,
        source_page: int | None,
        source_text: str | None,
        confidence: float | None,
    ) -> SourcedValue:
        return cls(
            value=value,
            source_page=source_page,
            source_text=_clip_source_text(source_text),
            confidence=confidence,
        )


class ClassificationWarning(BaseModel):
    """A field that could not be extracted confidently."""

    model_config = ConfigDict(extra="forbid")

    code: str
    message: str
    section: str | None = None
    field: str | None = None
    source_page: int | None = None


class IdentifiedDocumentSection(BaseModel):
    """A detected financial-statement section and the pages it occupies."""

    model_config = ConfigDict(extra="forbid")

    section_id: str
    title: str
    start_page: int | None = None
    end_page: int | None = None
    pages: list[int] = Field(default_factory=list)
    confidence: float = 0.0
    excerpt: str | None = None


class LineItem(BaseModel):
    """One caption matched in the source, with current and previous amounts."""

    model_config = ConfigDict(extra="forbid")

    label: str
    mapping_code: str | None = None
    note_no: SourcedValue = Field(default_factory=unextracted_value)
    current_period: SourcedValue = Field(default_factory=unextracted_value)
    previous_period: SourcedValue = Field(default_factory=unextracted_value)


class NoteExtract(BaseModel):
    """One note heading located in the uploaded statements."""

    model_config = ConfigDict(extra="forbid")

    note_id: str | None = None
    title: str | None = None
    start_page: int | None = None
    pages: list[int] = Field(default_factory=list)
    excerpt: str | None = None
    line_items: list[LineItem] = Field(default_factory=list)


class FinancialBlock(BaseModel):
    """Shared container for a classified statement or note schedule."""

    model_config = ConfigDict(extra="forbid")

    identified: bool = False
    start_page: int | None = None
    end_page: int | None = None
    pages: list[int] = Field(default_factory=list)
    confidence: float = 0.0
    excerpt: str | None = None
    line_items: list[LineItem] = Field(default_factory=list)

    def apply_section(self, section: IdentifiedDocumentSection | None) -> None:
        if section is None:
            return
        self.identified = True
        self.start_page = section.start_page
        self.end_page = section.end_page
        self.pages = list(section.pages)
        self.confidence = section.confidence
        self.excerpt = section.excerpt


class CompanyInfo(BaseModel):
    """Cover-page / general information. Unread fields stay null."""

    model_config = ConfigDict(extra="forbid")

    company_name: SourcedValue = Field(default_factory=unextracted_value)
    cin: SourcedValue = Field(default_factory=unextracted_value)
    registered_office: SourcedValue = Field(default_factory=unextracted_value)
    period_end: SourcedValue = Field(default_factory=unextracted_value)
    period_start: SourcedValue = Field(default_factory=unextracted_value)
    reporting_period: SourcedValue = Field(default_factory=unextracted_value)
    currency: SourcedValue = Field(default_factory=unextracted_value)
    unit: SourcedValue = Field(default_factory=unextracted_value)
    nature_of_company: SourcedValue = Field(default_factory=unextracted_value)


class BalanceSheet(FinancialBlock):
    """Part I – Form of Balance Sheet."""


class ProfitAndLoss(FinancialBlock):
    """Part II – Statement of Profit and Loss."""


class CashFlow(FinancialBlock):
    """Cash Flow Statement, when present in the uploaded PDF."""


class ShareCapital(FinancialBlock):
    authorised: SourcedValue = Field(default_factory=unextracted_value)
    issued: SourcedValue = Field(default_factory=unextracted_value)
    subscribed: SourcedValue = Field(default_factory=unextracted_value)
    paid_up: SourcedValue = Field(default_factory=unextracted_value)


class ReservesAndSurplus(FinancialBlock):
    """Note – Reserves and surplus."""


class Borrowings(FinancialBlock):
    long_term: SourcedValue = Field(default_factory=unextracted_value)
    short_term: SourcedValue = Field(default_factory=unextracted_value)


class PPE(FinancialBlock):
    """Property, Plant and Equipment."""


class CWIP(FinancialBlock):
    """Capital work-in-progress."""


class Investments(FinancialBlock):
    """Non-current and current investments."""


class Inventory(FinancialBlock):
    """Inventories."""


class TradeReceivables(FinancialBlock):
    """Trade receivables."""


class TradePayables(FinancialBlock):
    """Trade payables."""


class CashAndCashEquivalents(FinancialBlock):
    """Cash and cash equivalents."""


class Revenue(FinancialBlock):
    """Revenue from operations and other income."""


class Expenses(FinancialBlock):
    """Expense captions from the Statement of Profit and Loss."""


class TaxExpense(FinancialBlock):
    current_tax: SourcedValue = Field(default_factory=unextracted_value)
    deferred_tax: SourcedValue = Field(default_factory=unextracted_value)


class EPS(FinancialBlock):
    basic: SourcedValue = Field(default_factory=unextracted_value)
    diluted: SourcedValue = Field(default_factory=unextracted_value)


class OtherDisclosures(FinancialBlock):
    """Additional regulatory information and other disclosures."""


class Ratios(FinancialBlock):
    """Analytical ratios disclosed under additional regulatory information."""


class FinancialNotes(FinancialBlock):
    notes: list[NoteExtract] = Field(default_factory=list)


class ClassifiedFinancialData(BaseModel):
    """Full classification result for one uploaded financial statement PDF."""

    model_config = ConfigDict(extra="forbid")

    company: CompanyInfo = Field(default_factory=CompanyInfo)
    sections: list[IdentifiedDocumentSection] = Field(default_factory=list)
    balance_sheet: BalanceSheet = Field(default_factory=BalanceSheet)
    profit_and_loss: ProfitAndLoss = Field(default_factory=ProfitAndLoss)
    cash_flow: CashFlow = Field(default_factory=CashFlow)
    notes_to_accounts: FinancialNotes = Field(default_factory=FinancialNotes)
    share_capital: ShareCapital = Field(default_factory=ShareCapital)
    reserves_and_surplus: ReservesAndSurplus = Field(default_factory=ReservesAndSurplus)
    borrowings: Borrowings = Field(default_factory=Borrowings)
    ppe: PPE = Field(default_factory=PPE)
    cwip: CWIP = Field(default_factory=CWIP)
    investments: Investments = Field(default_factory=Investments)
    inventory: Inventory = Field(default_factory=Inventory)
    trade_receivables: TradeReceivables = Field(default_factory=TradeReceivables)
    trade_payables: TradePayables = Field(default_factory=TradePayables)
    cash_and_cash_equivalents: CashAndCashEquivalents = Field(default_factory=CashAndCashEquivalents)
    revenue: Revenue = Field(default_factory=Revenue)
    expenses: Expenses = Field(default_factory=Expenses)
    tax_expense: TaxExpense = Field(default_factory=TaxExpense)
    eps: EPS = Field(default_factory=EPS)
    other_disclosures: OtherDisclosures = Field(default_factory=OtherDisclosures)
    ratios: Ratios = Field(default_factory=Ratios)
    warnings: list[ClassificationWarning] = Field(default_factory=list)

    def model_dump_classified(self) -> dict[str, Any]:
        return self.model_dump(mode="json")


def _clip_source_text(text: str | None, limit: int = 240) -> str | None:
    if text is None:
        return None
    compact = " ".join(text.split())
    if len(compact) <= limit:
        return compact
    return compact[: limit - 1].rstrip() + "…"
