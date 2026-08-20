"""Tests for human review before Excel generation."""

from __future__ import annotations

from io import BytesIO

from fastapi.testclient import TestClient

from app.excel.workbook_generator import WorkbookGenerator
from app.review.review_service import (
    ReviewIncompleteError,
    approved_mapped,
    build_review,
    parse_extracted_value,
    previous_year_label,
)
from tests.test_api_jobs import TEXT_PDF, _client, _wait_for_job


def test_blank_extracted_value_is_not_coerced_to_zero() -> None:
    assert parse_extracted_value("") is None
    assert parse_extracted_value("  ") is None
    assert parse_extracted_value(None) is None
    assert parse_extracted_value("1,00,000") == 100000
    assert parse_extracted_value("3", period="note") == "3"


def test_previous_year_label_from_range() -> None:
    assert previous_year_label("2023-24") == "2022-23"
    assert previous_year_label("2024-03-31") == "2023-03-31"


def test_below_threshold_requires_review_and_is_not_auto_written() -> None:
    mapped = {
        "placements": [
            {
                "field_key": "share_capital",
                "source_label": "Share capital",
                "extracted_value": 100000,
                "source_page": 1,
                "schedule_iii_category": "Shareholders' funds",
                "excel_sheet": "BS PnL",
                "excel_cell": "F20",
                "excel_destination": "BS PnL!F20",
                "period": "current",
                "confidence": 0.4,
                "action": "write",
                "resolution": "fuzzy",
            },
            {
                "field_key": "revenue_from_operations",
                "source_label": "Revenue from operations",
                "extracted_value": 250000.5,
                "source_page": 1,
                "schedule_iii_category": "Revenue",
                "excel_sheet": "BS PnL",
                "excel_cell": "F40",
                "excel_destination": "BS PnL!F40",
                "period": "current",
                "confidence": 0.96,
                "action": "write",
                "resolution": "exact",
            },
        ]
    }
    review = build_review(mapped, {})
    by_key = {item["source_field"]: item for item in review["items"]}
    assert by_key["Share capital"]["status"] == "needs_review"
    assert by_key["Revenue from operations"]["status"] == "pending"

    review["items"][1]["status"] = "approved"
    try:
        approved_mapped(mapped, review)
        raise AssertionError("expected ReviewIncompleteError")
    except ReviewIncompleteError as exc:
        assert "confidence threshold" in exc.message

    review["items"][0]["status"] = "rejected"
    result = approved_mapped(mapped, review)
    assert [item["field_key"] for item in result["placements"]] == ["revenue_from_operations"]


def test_edited_approved_value_is_written_rejected_is_not() -> None:
    mapped = {
        "placements": [
            {
                "field_key": "share_capital",
                "source_label": "Share capital",
                "extracted_value": 100000,
                "source_page": 1,
                "schedule_iii_category": "Shareholders' funds",
                "excel_sheet": "BS PnL",
                "excel_cell": "F20",
                "period": "current",
                "confidence": 0.9,
                "action": "write",
                "resolution": "exact",
            }
        ]
    }
    review = build_review(mapped, {})
    review["items"][0]["extracted_value"] = 125000
    review["items"][0]["status"] = "approved"
    result = approved_mapped(mapped, review)
    assert result["placements"][0]["extracted_value"] == 125000
    assert result["placements"][0]["review_status"] == "approved"


def test_upload_pauses_for_review_then_approve_generates(tmp_path, monkeypatch) -> None:
    client = _client(tmp_path, monkeypatch)
    uploaded = client.post("/upload", files={"file": ("statement.pdf", BytesIO(TEXT_PDF), "application/pdf")})
    assert uploaded.status_code == 200, uploaded.text
    job_id = uploaded.json()["job_id"]
    waiting = _wait_for_job(client, job_id)
    assert waiting["status"] == "review_required"
    assert waiting["stage"] == "review"

    review = client.post(f"/review/{job_id}", json={})
    assert review.status_code == 200, review.text
    body = review.json()
    assert body["job_id"] == job_id
    assert "items" in body
    if body["items"]:
        row = body["items"][0]
        for key in (
            "source_field",
            "extracted_value",
            "source_page",
            "schedule_iii_category",
            "excel_destination",
            "confidence",
            "status",
        ):
            assert key in row

    blocked = client.post(f"/approve/{job_id}")
    if (body.get("summary") or {}).get("needs_review"):
        assert blocked.status_code == 409
        assert blocked.json()["error"] == "review_incomplete"
        assert "traceback" not in blocked.text.lower()

    items = [{"item_id": item["item_id"], "status": "approved"} for item in body["items"]]
    saved = client.post(f"/review/{job_id}", json={"items": items, "financial_year": {"current": "2023-24", "previous": "2022-23"}})
    assert saved.status_code == 200, saved.text
    assert saved.json()["financial_year"]["current"] == "2023-24"

    approved = client.post(f"/approve/{job_id}")
    assert approved.status_code == 200, approved.text
    payload = approved.json()
    assert payload["status"] == "completed"
    assert payload["output_file"]
    assert payload["validation_status"] in {"PASS", "WARNING", "ERROR"}
    download = client.get(f"/download/{job_id}")
    assert download.status_code == 200
    assert download.content[:2] == b"PK"


def test_workbook_skips_values_that_were_not_approved(tmp_path) -> None:
    generator = WorkbookGenerator(output_dir=tmp_path)
    result = generator.generate_detailed(
        {
            "placements": [
                {
                    "field_key": "share_capital",
                    "extracted_value": 100000,
                    "excel_sheet": "MissingSheet",
                    "excel_cell": "A1",
                    "action": "write",
                    "review_status": "rejected",
                }
            ]
        }
    )
    assert result.written == []
    assert result.skipped
    assert result.skipped[0]["reason"] == "not_approved"
