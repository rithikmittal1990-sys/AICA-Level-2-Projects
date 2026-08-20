"""Tests for read-only financial statement validation."""

from __future__ import annotations

from copy import deepcopy

from app.validation.financial_validator import FinancialValidator


def _sourced(value, *, page: int = 1, confidence: float = 0.95) -> dict:
    return {"value": value, "source_page": page, "source_text": str(value), "confidence": confidence}


def _line(label: str, current, previous=None) -> dict:
    return {
        "label": label,
        "mapping_code": None,
        "note_no": _sourced(None, confidence=None),
        "current_period": _sourced(current),
        "previous_period": _sourced(previous) if previous is not None else _sourced(None, confidence=None),
    }


def _classified(*, bs=None, pnl=None, notes=None, identified_bs=True, identified_pnl=True) -> dict:
    return {
        "balance_sheet": {"identified": identified_bs, "line_items": bs or []},
        "profit_and_loss": {"identified": identified_pnl, "line_items": pnl or []},
        "share_capital": {"identified": False, "line_items": (notes or {}).get("share_capital", [])},
        "ppe": {"identified": False, "line_items": (notes or {}).get("ppe", [])},
        "warnings": [],
    }


BALANCED_BS = [
    _line("Share capital", 100, 80),
    _line("Reserves and surplus", 50, 40),
    _line("Long-term borrowings", 30, 20),
    _line("Trade payables", 20, 10),
    _line("Property, Plant and Equipment", 140, 100),
    _line("Inventories", 40, 30),
    _line("Cash and cash equivalents", 20, 20),
    _line("Total Assets", 200, 150),
    _line("Total Equity and Liabilities", 200, 150),
]

BALANCED_PNL = [
    _line("Revenue from operations", 300, 250),
    _line("Other income", 20, 10),
    _line("Total Income", 320, 260),
    _line("Employee benefits expense", 100, 90),
    _line("Other expenses", 80, 70),
    _line("Total expenses", 180, 160),
    _line("Profit before tax", 140, 100),
    _line("Tax expense", 40, 30),
    _line("Profit/ (Loss)", 100, 70),
    _line("Basic", 10, 7),
    _line("Diluted", 9.5, 6.8),
]


def test_validation_module_importable() -> None:
    assert FinancialValidator is not None


def test_balanced_statements_pass() -> None:
    result = FinancialValidator().validate(_classified(bs=BALANCED_BS, pnl=BALANCED_PNL))
    assert result["status"] == "PASS"
    assert result["errors"] == []
    identities = {
        "bs_assets_equal_equity_and_liabilities",
        "pnl_income_minus_expenses_equals_pbt",
        "pnl_pbt_minus_tax_equals_pat",
        "bs_current_plus_non_current_assets",
        "bs_equity_plus_liabilities",
    }
    for check in result["checks"]:
        if check["id"] in identities and check["period"] in {"current", "previous"}:
            assert check["status"] == "PASS", check["message"]


def test_unbalanced_balance_sheet_is_error_with_exact_difference() -> None:
    bs = [_line("Share capital", 100), _line("Total Assets", 200), _line("Total Equity and Liabilities", 180)]
    result = FinancialValidator().validate(_classified(bs=bs, pnl=[], identified_pnl=False))
    assert result["status"] == "ERROR"
    mismatch = next(
        item
        for item in result["errors"]
        if item["code"] == "bs_assets_equal_equity_and_liabilities" and item["period"] == "current"
    )
    assert "200" in mismatch["message"]
    assert "180" in mismatch["message"]
    assert "20" in mismatch["message"]
    assert "current" in mismatch["message"]
    assert "did not change" in mismatch["message"].lower() or "not change" in mismatch["message"].lower()


def test_periods_are_validated_separately() -> None:
    bs = [
        _line("Total Assets", 200, 150),
        _line("Total Equity and Liabilities", 200, 140),
        _line("Share capital", 100, 80),
    ]
    result = FinancialValidator().validate(_classified(bs=bs, pnl=[], identified_pnl=False))
    current = next(
        item
        for item in result["checks"]
        if item["id"] == "bs_assets_equal_equity_and_liabilities" and item["period"] == "current"
    )
    previous = next(
        item
        for item in result["checks"]
        if item["id"] == "bs_assets_equal_equity_and_liabilities" and item["period"] == "previous"
    )
    assert current["status"] == "PASS"
    assert previous["status"] == "ERROR"
    assert "previous" in previous["message"]
    assert "150" in previous["message"] and "140" in previous["message"]


def test_profit_identities() -> None:
    pnl = [
        _line("Revenue from operations", 100),
        _line("Total Income", 100),
        _line("Total expenses", 40),
        _line("Profit before tax", 50),
        _line("Tax expense", 10),
        _line("Profit/ (Loss)", 50),
    ]
    result = FinancialValidator().validate(_classified(bs=[], pnl=pnl, identified_bs=False))
    income = next(item for item in result["checks"] if item["id"] == "pnl_income_minus_expenses_equals_pbt" and item["period"] == "current")
    pat = next(item for item in result["checks"] if item["id"] == "pnl_pbt_minus_tax_equals_pat" and item["period"] == "current")
    assert income["status"] == "ERROR"
    assert "60" in income["message"] and "50" in income["message"]
    assert pat["status"] == "ERROR"
    assert "40" in pat["message"]


def test_missing_values_are_warnings_not_filled() -> None:
    payload = _classified(bs=[_line("Inventories", 10)], pnl=[], identified_pnl=False)
    before = deepcopy(payload)
    result = FinancialValidator().validate(payload)
    assert payload == before
    assert result["status"] in {"WARNING", "ERROR"}
    missing = [item for item in result["warnings"] if item["code"] == "missing_share_capital"]
    assert missing
    incomplete = next(
        item
        for item in result["checks"]
        if item["id"] == "bs_assets_equal_equity_and_liabilities" and item["period"] == "current"
    )
    assert incomplete["status"] == "WARNING"
    assert "not extracted" in incomplete["message"]


def test_note_total_reconciles_and_detects_mismatch() -> None:
    ok = FinancialValidator().validate(
        _classified(
            bs=[_line("Share capital", 100, 80)],
            pnl=[],
            identified_pnl=False,
            notes={"share_capital": [_line("Share capital", 100, 80)]},
        )
    )
    note_ok = next(item for item in ok["checks"] if item["id"] == "note_reconcile_share_capital" and item["period"] == "current")
    assert note_ok["status"] == "PASS"

    bad = FinancialValidator().validate(
        _classified(
            bs=[_line("Share capital", 100)],
            pnl=[],
            identified_pnl=False,
            notes={"share_capital": [_line("Share capital", 90)]},
        )
    )
    note_bad = next(item for item in bad["errors"] if item["code"] == "note_reconcile_share_capital")
    assert "100" in note_bad["message"]
    assert "90" in note_bad["message"]


def test_duplicate_mappings_are_reported() -> None:
    classified = _classified(bs=[_line("Share capital", 100)], pnl=[], identified_pnl=False)
    result = FinancialValidator().validate(
        {
            **classified,
            "placements": [
                {
                    "field_key": "share_capital",
                    "period": "current",
                    "extracted_value": 100,
                    "excel_sheet": "BS PnL",
                    "excel_cell": "F13",
                },
                {
                    "field_key": "reserves_and_surplus",
                    "period": "current",
                    "extracted_value": 50,
                    "excel_sheet": "BS PnL",
                    "excel_cell": "F13",
                },
            ],
        }
    )
    assert any(item["code"] == "duplicate_cell_mapping" for item in result["errors"])
    assert "F13" in result["errors"][0]["message"] or any("F13" in item["message"] for item in result["errors"])


def test_negative_share_capital_is_error() -> None:
    result = FinancialValidator().validate(
        _classified(bs=[_line("Share capital", -10), _line("Total Assets", 1), _line("Total Equity and Liabilities", 1)], pnl=[], identified_pnl=False)
    )
    assert any(item["code"] == "sign_share_capital" for item in result["errors"])
    assert any("-10" in item["message"] for item in result["errors"])


def test_eps_consistency_and_skip_when_insufficient() -> None:
    with_eps = FinancialValidator().validate(_classified(bs=[], pnl=BALANCED_PNL, identified_bs=False))
    eps = [item for item in with_eps["checks"] if "eps" in item["id"] and item["period"] == "current"]
    assert eps
    assert all(item["status"] in {"PASS", "SKIPPED"} for item in eps)

    opposite = FinancialValidator().validate(
        _classified(
            bs=[],
            identified_bs=False,
            pnl=[_line("Profit/ (Loss)", 100), _line("Basic", -2)],
        )
    )
    assert any(item["code"] == "eps_sign_vs_pat" for item in opposite["errors"])


def test_does_not_alter_source_values() -> None:
    payload = _classified(bs=[_line("Total Assets", 10), _line("Total Equity and Liabilities", 8)], pnl=[])
    snapshot = deepcopy(payload)
    FinancialValidator().validate(payload)
    assert payload == snapshot
    assert payload["balance_sheet"]["line_items"][0]["current_period"]["value"] == 10
