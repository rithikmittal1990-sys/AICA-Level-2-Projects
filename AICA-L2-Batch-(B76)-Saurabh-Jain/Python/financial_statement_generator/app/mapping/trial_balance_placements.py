"""Adjust Excel placements for trial balance uploads."""

from __future__ import annotations

from typing import Any

from app.mapping.trial_balance_notes import build_note_placements

EXCLUDED_SHEETS = ("Cash Flow",)


def prepare_trial_balance_mapped(
    mapped: dict[str, Any],
    classified: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Use only uploaded trial balance values and notes derived from them."""
    placements = mapped.get("placements")
    if not isinstance(placements, list):
        mapped["placements"] = []
        placements = mapped["placements"]

    for placement in placements:
        if not isinstance(placement, dict):
            continue
        _fix_company_name_placement(placement)
        _allow_formula_overwrite_for_values(placement)

    note_placements = build_note_placements(mapped, classified)
    placements.extend(note_placements)

    mapped["generation_mode"] = "trial_balance"
    mapped["exclude_sheets"] = list(EXCLUDED_SHEETS)
    if classified:
        company = classified.get("company") or {}
        if isinstance(company, dict):
            name = (company.get("company_name") or {}).get("value")
            period = (company.get("reporting_period") or {}).get("value")
            if name:
                mapped["company_name"] = name
            if period:
                mapped["reporting_period"] = period
    return mapped


def _fix_company_name_placement(placement: dict[str, Any]) -> None:
    if placement.get("field_key") != "company_name":
        return
    if not placement.get("extracted_value"):
        return
    placement["excel_sheet"] = placement.get("excel_sheet") or "BS PnL"
    placement["excel_cell"] = placement.get("excel_cell") or "A2"
    placement["value_role"] = "text"
    placement["overwrite_formula"] = True
    placement["action"] = "write"
    placement["resolution"] = placement.get("resolution") or "configured_cell"
    warnings = list(placement.get("warnings") or [])
    warnings.append("Company name from the trial balance replaces the template sample name.")
    placement["warnings"] = warnings


def _allow_formula_overwrite_for_values(placement: dict[str, Any]) -> None:
    if placement.get("field_key") == "company_name":
        return
    if placement.get("extracted_value") is None:
        return
    if placement.get("action") != "skip_formula":
        return
    placement["overwrite_formula"] = True
    placement["action"] = "write"
    warnings = list(placement.get("warnings") or [])
    warnings.append("Trial balance amount replaces the template formula on the face sheet.")
    placement["warnings"] = warnings
