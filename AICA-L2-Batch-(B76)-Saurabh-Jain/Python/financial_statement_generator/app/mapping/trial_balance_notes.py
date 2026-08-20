"""Build note-sheet placements from trial balance classified data."""

from __future__ import annotations

from typing import Any

# Face sheet and supporting note totals populated only from trial balance values.
NOTE_CELL_TARGETS: dict[str, list[tuple[str, str]]] = {
    "share_capital": [("Note 3 (Share Capital)", "H13")],
    "reserves_and_surplus": [("NOTE (4-12)", "E16")],
    "long_term_borrowings": [("NOTE (4-12)", "E30")],
    "short_term_borrowings": [("NOTE (4-12)", "E71")],
    "trade_payables": [("NOTE (4-12)", "G81")],
    "trade_receivables": [("Note (13-20)", "H46")],
    "cash_and_cash_equivalents": [("Note (13-20)", "F63")],
    "short_term_loans_and_advances": [("Note (13-20)", "F72")],
    "other_current_assets": [("Note (13-20)", "F88")],
    "property_plant_and_equipment": [("Note 12 (PPE)", "J20")],
    "intangible_assets": [("Note 12 (PPE)", "J28")],
    "revenue_from_operations": [("Note 20-31", "C10"), ("Note 20-31", "C11")],
    "other_income": [("Note 20-31", "C23")],
    "employee_benefits_expense": [("Note 20-31", "C71")],
    "depreciation_and_amortization_expense": [("Note 20-31", "C82"), ("BS PnL", "F95")],
    "other_expenses": [("Note 20-31", "C93")],
    "current_tax": [("Note 20-31", "C124")],
}

FACE_SHEET_TARGETS: dict[str, str] = {
    "revenue_from_operations": "F83",
    "other_income": "F84",
    "employee_benefits_expense": "F96",
    "depreciation_and_amortization_expense": "F95",
    "other_expenses": "F97",
    "current_tax": "F107",
}


def build_note_placements(mapped: dict[str, Any], classified: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    """Duplicate mapped trial balance amounts into note sheets and P&L note tables."""
    source_placements = [
        item for item in (mapped.get("placements") or []) if isinstance(item, dict)
    ]
    extra: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str]] = set()

    def add_placement(base: dict[str, Any], sheet: str, cell: str) -> None:
        key = (str(base.get("field_key") or ""), str(base.get("period") or "current"), sheet, cell.upper())
        if key in seen:
            return
        seen.add(key)
        extra.append(
            {
                **base,
                "excel_sheet": sheet,
                "excel_cell": cell,
                "excel_destination": f"{sheet}!{cell}",
                "action": "write",
                "overwrite_formula": True,
                "value_role": base.get("value_role") or "value",
                "resolution": "trial_balance_note",
                "warnings": [
                    *(base.get("warnings") or []),
                    "Note value taken from the uploaded trial balance.",
                ],
            }
        )

    for placement in source_placements:
        if placement.get("period") not in (None, "current"):
            continue
        if placement.get("extracted_value") is None:
            continue
        field_key = str(placement.get("field_key") or "")
        face_cell = FACE_SHEET_TARGETS.get(field_key)
        if face_cell:
            add_placement(placement, "BS PnL", face_cell)
        for sheet, cell in NOTE_CELL_TARGETS.get(field_key, []):
            add_placement(placement, sheet, cell)

    profit = _profit_for_period(source_placements, classified)
    if profit is not None:
        profit_base = {
            "field_key": "profit_for_period",
            "source_label": "Profit/(Loss) for the period",
            "extracted_value": round(profit, 2),
            "schedule_iii_category": "Statement of Profit and Loss",
            "schedule_iii_label": "Profit/(Loss) for the period",
            "period": "current",
            "confidence": 0.95,
            "action": "write",
            "overwrite_formula": True,
            "value_role": "value",
            "resolution": "trial_balance_note",
            "warnings": ["Profit computed from trial balance income and expense lines."],
        }
        add_placement(profit_base, "BS PnL", "F109")
        add_placement(profit_base, "NOTE (4-12)", "E11")

    return extra


def _profit_for_period(placements: list[dict[str, Any]], classified: dict[str, Any] | None) -> float | None:
    income_keys = {"revenue_from_operations", "other_income"}
    expense_keys = {
        "cost_of_materials_consumed",
        "purchases_of_stock_in_trade",
        "employee_benefits_expense",
        "finance_costs",
        "depreciation_and_amortization_expense",
        "other_expenses",
        "current_tax",
        "deferred_tax",
    }
    income = 0.0
    expense = 0.0
    found = False
    for placement in placements:
        if placement.get("period") not in (None, "current"):
            continue
        value = placement.get("extracted_value")
        if not isinstance(value, (int, float)):
            continue
        field_key = str(placement.get("field_key") or "")
        if field_key in income_keys:
            income += float(value)
            found = True
        elif field_key in expense_keys:
            expense += float(value)
            found = True
    if found:
        return income - expense
    if not classified:
        return None
    profit_loss = classified.get("profit_and_loss") or {}
    total = 0.0
    count = 0
    for line in profit_loss.get("line_items") or []:
        if not isinstance(line, dict):
            continue
        label = str(line.get("label") or "").lower()
        current = (line.get("current_period") or {}).get("value")
        if not isinstance(current, (int, float)):
            continue
        if "revenue" in label or "other income" in label:
            total += float(current)
            count += 1
        elif any(token in label for token in ("expense", "tax", "depreciation", "finance cost")):
            total -= float(current)
            count += 1
    return total if count else None
