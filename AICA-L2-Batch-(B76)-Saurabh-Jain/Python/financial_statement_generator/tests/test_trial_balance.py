"""Tests for trial balance Excel ingestion."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.classification.trial_balance_classifier import TrialBalanceClassifier
from app.config import settings
from app.extraction.trial_balance_reader import TrialBalanceReader
from app.main import app
from app.mapping.schedule_iii_mapper import ScheduleIIIMapper
from app.mapping.trial_balance_placements import prepare_trial_balance_mapped


def _sample_trial_balance(path: Path) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Trial Balance"
    sheet["A1"] = "XYZ Private Limited"
    sheet["A2"] = "Trial Balance"
    sheet["A3"] = "1-Apr-25 to 31-Mar-26"
    sheet["A5"] = "Particulars"
    sheet["B6"] = "Closing Balance"
    sheet["B7"] = "Debit"
    sheet["C7"] = "Credit"
    sheet["A8"] = "Share Capital - Ankit"
    sheet["C8"] = 1000000
    sheet["A9"] = "Reserve and Surplus"
    sheet["C9"] = 252871.28
    sheet["A10"] = "Sundry Debtors"
    sheet["B10"] = 157680
    sheet["A11"] = "Cash-in-hand"
    sheet["B11"] = 572946.81
    sheet["A12"] = "Bank Accounts"
    sheet["B12"] = 145269.13
    sheet["A13"] = "Sundry Creditors"
    sheet["C13"] = 16300
    sheet["A14"] = "Sale of Service"
    sheet["C14"] = 618478
    sheet["A15"] = "Depreciation"
    sheet["B15"] = 34379.17
    sheet["A16"] = "Salary"
    sheet["B16"] = 75000
    workbook.save(path)
    workbook.close()
    return path


def test_trial_balance_reader_parses_accounts(tmp_path: Path) -> None:
    path = _sample_trial_balance(tmp_path / "TrialBal.xlsx")
    document = TrialBalanceReader(upload_dir=tmp_path).read_path(path)
    assert document.company_name == "XYZ Private Limited"
    assert document.period_label == "1-Apr-25 to 31-Mar-26"
    labels = {row.label for row in document.rows}
    assert "Share Capital - Ankit" in labels
    assert "Sale of Service" in labels


def test_trial_balance_classifier_maps_to_schedule_iii(tmp_path: Path) -> None:
    path = _sample_trial_balance(tmp_path / "TrialBal.xlsx")
    document = TrialBalanceReader(upload_dir=tmp_path).read_path(path)
    classified = TrialBalanceClassifier().classify(document)
    bs_labels = {item.label for item in classified.balance_sheet.line_items}
    pl_labels = {item.label for item in classified.profit_and_loss.line_items}
    assert "Share capital" in bs_labels
    assert "Reserves and surplus" in bs_labels
    assert "Cash and cash equivalents" in bs_labels
    assert "Revenue from operations" in pl_labels


def test_trial_balance_pipeline_maps_into_excel_template(tmp_path: Path) -> None:
    path = _sample_trial_balance(tmp_path / "TrialBal.xlsx")
    document = TrialBalanceReader(upload_dir=tmp_path).read_path(path)
    classified = TrialBalanceClassifier().classify(document)
    classified_dict = classified.model_dump_classified()
    mapped = ScheduleIIIMapper().map_classified(classified)
    mapped = prepare_trial_balance_mapped(mapped, classified_dict)
    share = next(
        item
        for item in mapped.get("placements") or []
        if item.get("field_key") == "share_capital" and item.get("period") == "current"
    )
    assert share.get("extracted_value") == 1000000
    assert share.get("excel_sheet") == "BS PnL"
    assert share.get("action") == "write"
    company = next(item for item in mapped.get("placements") or [] if item.get("field_key") == "company_name")
    assert company.get("extracted_value") == "XYZ Private Limited"
    assert company.get("action") == "write"
    assert mapped.get("generation_mode") == "trial_balance"
    assert "Cash Flow" in (mapped.get("exclude_sheets") or [])


@pytest.mark.skipif(
    not Path("/Users/shruti/Downloads/TrialBal.xlsx").exists(),
    reason="User sample trial balance not available",
)
def test_user_trial_balance_sample_parses() -> None:
    document = TrialBalanceReader().read_path(Path("/Users/shruti/Downloads/TrialBal.xlsx"))
    assert len(document.rows) >= 20
    classified = TrialBalanceClassifier().classify(document)
    assert classified.balance_sheet.line_items
    assert classified.profit_and_loss.line_items


def test_upload_accepts_trial_balance_xlsx(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(settings, "output_dir", tmp_path)
    monkeypatch.setattr(settings, "input_dir", tmp_path / "input")
    settings.ensure_directories()
    path = _sample_trial_balance(tmp_path / "TrialBal.xlsx")
    content = path.read_bytes()
    client = TestClient(app)
    response = client.post(
        "/upload",
        files={
            "file": (
                "TrialBal.xlsx",
                BytesIO(content),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert response.status_code == 200, response.text
    assert "job_id" in response.json()
