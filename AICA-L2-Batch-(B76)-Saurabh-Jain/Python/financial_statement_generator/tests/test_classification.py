"""Tests for financial data classification and sourced models."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from fastapi.testclient import TestClient

from app.classification.document_classifier import DocumentClassifier, parse_amount
from app.extraction.pdf_extractor import PDFExtractor
from app.main import app
from app.models.financial_models import (
    ClassifiedFinancialData,
    CompanyInfo,
    HealthResponse,
    SourcedValue,
    unextracted_value,
)

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


def _classifier() -> DocumentClassifier:
    return DocumentClassifier(mappings={})


def _page(
    page_number: int,
    text: str,
    *,
    tables: list[dict] | None = None,
    sections: list[dict] | None = None,
    headings: list[dict] | None = None,
) -> dict:
    return {
        "page_number": page_number,
        "text": text,
        "tables": tables or [],
        "sections": sections or [],
        "headings": headings or [],
        "numbers": [],
        "dates": [],
    }


SAMPLE_PAGES = [
    _page(
        1,
        "ABC Private Limited\nCIN: U12345MH2010PTC123456\n"
        "BALANCE SHEET as at 31-03-2024\n(Rs. in lakhs)\n"
        "Share capital 1,00,000 80,000\n"
        "Reserves and surplus 50,000 40,000\n"
        "Property, Plant and Equipment 2,00,000 1,80,000",
        tables=[
            {
                "page_number": 1,
                "rows": [
                    ["Particulars", "Note", "31-03-2024", "31-03-2023"],
                    ["Share capital", "3", "1,00,000", "80,000"],
                    ["Reserves and surplus", "4", "50,000", "40,000"],
                    ["Property, Plant and Equipment", "12", "2,00,000", "1,80,000"],
                    ["Inventories", "8", "25,000", "20,000"],
                    ["Trade receivables", "9", "15,000", "12,000"],
                    ["Cash and cash equivalents", "10", "5,000", "4,000"],
                    ["Long-term borrowings", "5", "75,000", "90,000"],
                ],
            }
        ],
        sections=[{"name": "balance sheet", "page_number": 1, "snippet": "BALANCE SHEET"}],
        headings=[{"text": "BALANCE SHEET as at 31-03-2024", "line_number": 1}],
    ),
    _page(
        2,
        "Statement of Profit and Loss for the year ended 31-03-2024\n"
        "Revenue from operations 12,50,000 11,00,000\n"
        "Employee benefits expense 2,00,000 1,80,000\n"
        "Tax expense 40,000 35,000\n"
        "Current tax 38,000 33,000\n"
        "Deferred tax 2,000 2,000\n"
        "Earnings per equity share\nBasic 12.50 11.00\nDiluted 12.40 10.90",
        tables=[
            {
                "page_number": 2,
                "rows": [
                    ["Particulars", "Note", "31-03-2024", "31-03-2023"],
                    ["Revenue from operations", "20", "12,50,000", "11,00,000"],
                    ["Employee benefits expense", "22", "2,00,000", "1,80,000"],
                    ["Tax expense", "23", "40,000", "35,000"],
                    ["Current tax", "", "38,000", "33,000"],
                    ["Earnings per equity share", "", "", ""],
                    ["Basic", "", "12.50", "11.00"],
                    ["Diluted", "", "12.40", "10.90"],
                ],
            }
        ],
        sections=[{"name": "statement of profit and loss", "page_number": 2, "snippet": "P&L"}],
    ),
    _page(
        3,
        "Cash Flow Statement\n"
        "Cash flow from operating activities 30,000 22,000\n"
        "Notes to Accounts\nNote 3 Share Capital\nAuthorised 5,00,000 5,00,000\n"
        "Additional Regulatory Information\nCurrent Ratio 1.50 1.40",
        tables=[
            {
                "page_number": 3,
                "rows": [
                    ["Particulars", "Current", "Previous"],
                    ["Cash flow from operating activities", "30,000", "22,000"],
                    ["Current Ratio", "1.50", "1.40"],
                ],
            }
        ],
    ),
]


def test_health_response_model() -> None:
    payload = HealthResponse(status="ok", service="Financial Statement Generator", version="0.1.0")
    assert payload.status == "ok"


def test_unextracted_value_is_null_not_zero() -> None:
    value = unextracted_value()
    assert value.value is None
    assert value.source_page is None
    assert value.confidence is None
    dumped = SourcedValue(value=None, source_page=None, source_text=None, confidence=None).model_dump()
    assert dumped["value"] is None
    assert dumped["value"] != 0


def test_company_info_defaults_are_null() -> None:
    info = CompanyInfo()
    assert info.company_name.value is None
    assert info.cin.value is None


def test_parse_amount_does_not_invent_blanks() -> None:
    assert parse_amount("1,00,000") == 100000.0
    assert parse_amount("(25,000)") == -25000.0
    assert parse_amount("Nil") == 0.0
    assert parse_amount("-") is None
    assert parse_amount("") is None
    assert parse_amount("n/a") is None


def test_classifies_required_sections() -> None:
    result = _classifier().classify({"pages": SAMPLE_PAGES})
    section_ids = {section.section_id for section in result.sections}
    for required in (
        "company_information",
        "balance_sheet",
        "profit_and_loss",
        "cash_flow",
        "notes_to_accounts",
        "share_capital",
        "reserves_and_surplus",
        "borrowings",
        "ppe",
        "inventory",
        "trade_receivables",
        "cash_and_cash_equivalents",
        "revenue",
        "expenses",
        "tax",
        "eps",
        "other_disclosures",
        "ratios",
    ):
        assert required in section_ids, required


def test_extracted_values_keep_provenance() -> None:
    result = _classifier().classify({"pages": SAMPLE_PAGES})
    share = next(item for item in result.balance_sheet.line_items if item.label.lower() == "share capital")
    assert share.current_period.value == 100000.0
    assert share.previous_period.value == 80000.0
    assert share.current_period.source_page == 1
    assert share.current_period.source_text
    assert share.current_period.confidence is not None
    assert share.current_period.confidence >= 0.7
    assert share.note_no.value == "3"

    dumped = share.current_period.model_dump()
    assert set(dumped) == {"value", "source_page", "source_text", "confidence"}


def test_company_information_is_extracted() -> None:
    result = _classifier().classify({"pages": SAMPLE_PAGES})
    assert result.company.company_name.value == "ABC Private Limited"
    assert result.company.cin.value == "U12345MH2010PTC123456"
    assert result.company.period_end.value == "2024-03-31"
    assert result.company.unit.value
    assert "lakh" in result.company.unit.value.lower()


def test_named_note_models_are_populated() -> None:
    result = _classifier().classify({"pages": SAMPLE_PAGES})
    assert result.share_capital.identified is True
    assert result.ppe.identified is True
    assert result.revenue.line_items
    assert result.tax_expense.current_tax.value == 38000.0
    assert result.eps.basic.value == 12.5
    assert result.borrowings.long_term.value == 75000.0
    assert result.cash_flow.identified is True
    assert result.notes_to_accounts.notes
    assert any(note.title and "share capital" in note.title.lower() for note in result.notes_to_accounts.notes)


def test_missing_values_are_null_with_warning() -> None:
    pages = [
        _page(
            1,
            "BALANCE SHEET as at 31-03-2024\nShare capital",
            headings=[{"text": "BALANCE SHEET"}],
            sections=[{"name": "balance sheet"}],
        )
    ]
    result = _classifier().classify({"pages": pages})
    assert result.company.company_name.value is None
    codes = {warning.code for warning in result.warnings}
    assert "missing_value" in codes
    assert any(warning.field == "company_name" for warning in result.warnings)
    assert any("Share capital" in (warning.field or "") or "Share capital" in warning.message for warning in result.warnings)


def test_does_not_invent_missing_line_items() -> None:
    pages = [
        _page(
            1,
            "BALANCE SHEET as at 31-03-2024\nShare capital 1,00,000 80,000",
            tables=[
                {
                    "page_number": 1,
                    "rows": [
                        ["Particulars", "31-03-2024", "31-03-2023"],
                        ["Share capital", "1,00,000", "80,000"],
                    ],
                }
            ],
        )
    ]
    result = _classifier().classify({"pages": pages})
    labels = {item.label.lower() for item in result.balance_sheet.line_items}
    assert "share capital" in labels
    assert "inventories" not in labels
    assert result.inventory.line_items == []
    assert result.cwip.line_items == []
    assert result.trade_payables.line_items == []
    assert result.inventory.identified is False


def test_low_confidence_values_are_nulled() -> None:
    pages = [
        _page(
            1,
            "BALANCE SHEET as at 31-03-2024\nShare capital 1,00,000",
            tables=[
                {
                    "page_number": 1,
                    "rows": [["Particulars", "Amount"], ["Share capital", "1,00,000"]],
                }
            ],
        )
    ]
    result = DocumentClassifier(mappings={}, min_confidence=0.99).classify({"pages": pages})
    share = next(item for item in result.balance_sheet.line_items if item.label.lower() == "share capital")
    assert share.current_period.value is None
    assert share.current_period.source_text
    assert any(warning.code == "low_confidence" for warning in result.warnings)


def test_blank_table_cells_stay_null() -> None:
    pages = [
        _page(
            1,
            "Statement of Profit and Loss\nRevenue from operations",
            tables=[
                {
                    "page_number": 1,
                    "rows": [
                        ["Particulars", "Current", "Previous"],
                        ["Revenue from operations", "-", "11,00,000"],
                    ],
                }
            ],
        )
    ]
    result = _classifier().classify({"pages": pages})
    revenue = next(item for item in result.revenue.line_items if "revenue" in item.label.lower())
    assert revenue.current_period.value is None
    assert revenue.previous_period.value == 1100000.0


def test_classified_payload_is_json_serializable() -> None:
    result = _classifier().classify({"pages": SAMPLE_PAGES})
    payload = result.model_dump_classified()
    assert isinstance(payload, dict)
    assert payload["company"]["cin"]["value"] == "U12345MH2010PTC123456"
    assert "warnings" in payload
    ClassifiedFinancialData.model_validate(payload)


def test_classify_extracted_text_pdf(tmp_path: Path) -> None:
    pdf_path = tmp_path / "statement.pdf"
    pdf_path.write_bytes(TEXT_PDF)
    extracted = PDFExtractor(upload_dir=tmp_path, enable_ocr=False).extract(pdf_path)
    result = _classifier().classify(extracted)
    section_ids = {section.section_id for section in result.sections}
    assert "balance_sheet" in section_ids
    share = next((item for item in result.share_capital.line_items if "share capital" in item.label.lower()), None)
    assert share is not None
    assert share.current_period.value == 100000.0
    assert share.current_period.source_page == 1
    revenue = next((item for item in result.revenue.line_items if "revenue" in item.label.lower()), None)
    assert revenue is not None
    assert revenue.current_period.value == 250000.5


def test_classify_endpoint_returns_sourced_values() -> None:
    client = TestClient(app)
    response = client.post(
        "/classify",
        files={"file": ("statement.pdf", BytesIO(TEXT_PDF), "application/pdf")},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert "classified" in body
    classified = body["classified"]
    assert classified["company"]["period_end"]["value"]
    share_items = classified["share_capital"]["line_items"]
    assert share_items
    assert share_items[0]["current_period"]["source_page"] == 1
