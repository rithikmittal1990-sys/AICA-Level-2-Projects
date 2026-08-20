"""Tests for trial-balance-specific Excel placement adjustments."""

from __future__ import annotations

from app.mapping.trial_balance_placements import prepare_trial_balance_mapped


def test_prepare_trial_balance_mapped_replaces_template_formulas() -> None:
    mapped = {
        "placements": [
            {
                "field_key": "share_capital",
                "extracted_value": 1000000,
                "action": "skip_formula",
                "excel_sheet": "BS PnL",
                "excel_cell": "F13",
            },
            {
                "field_key": "company_name",
                "extracted_value": "XYZ Private Limited",
                "action": "unmapped_destination",
                "excel_cell": None,
            },
        ]
    }
    updated = prepare_trial_balance_mapped(
        mapped,
        {
            "company": {
                "company_name": {"value": "XYZ Private Limited"},
                "reporting_period": {"value": "1-Apr-25 to 31-Mar-26"},
            }
        },
    )
    share = next(item for item in updated["placements"] if item["field_key"] == "share_capital" and item["excel_cell"] == "F13")
    company = next(item for item in updated["placements"] if item["field_key"] == "company_name")
    assert share["action"] == "write"
    assert share["overwrite_formula"] is True
    assert company["action"] == "write"
    assert company["excel_cell"] == "A2"
    assert updated["generation_mode"] == "trial_balance"
    note = next(
        item
        for item in updated["placements"]
        if item["field_key"] == "share_capital" and item["excel_sheet"] == "Note 3 (Share Capital)"
    )
    assert note["extracted_value"] == 1000000
