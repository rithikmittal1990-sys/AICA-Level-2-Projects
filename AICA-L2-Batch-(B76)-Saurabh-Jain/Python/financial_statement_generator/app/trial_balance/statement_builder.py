"""Build Excel cell writes from classified trial balance data."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from app.trial_balance.models import TrialBalanceClassification
from app.trial_balance.parser import normalize_label

NOT_AVAILABLE = "Information not available from Trial Balance"
DEFAULT_FACE_VALUE = 10.0


@dataclass(slots=True)
class CellWrite:
    sheet: str
    cell: str
    value: float | str | None
    field_key: str
    overwrite_formula: bool = True


def build_cell_writes(classification: TrialBalanceClassification) -> list[CellWrite]:
    items = classification.line_items
    totals = classification.totals
    writes: list[CellWrite] = []

    def amount(key: str) -> float:
        return float(items.get(normalize_label(key), items.get(key, 0.0)) or 0.0)

    share_capital = amount("share_capital")
    reserves_tb = amount("reserves_and_surplus")
    current_year_profit = totals.current_year_profit
    reserves_face = round(reserves_tb + current_year_profit, 2)
    trade_payables = amount("trade_payables")
    other_cl = amount("other_current_liabilities") + amount("duties_and_taxes")
    short_term_provisions = amount("short_term_provisions")
    ppe = amount("property_plant_and_equipment")
    intangible = amount("intangible_assets")
    debtors = amount("trade_receivables")
    cash = amount("cash_and_cash_equivalents")
    advances = amount("short_term_loans_and_advances")
    other_ca = amount("other_current_assets") + amount("deferred_tax_asset")
    revenue = amount("revenue_from_operations")
    other_income = amount("other_income") + amount("deferred_tax_income")
    salary = amount("employee_benefits_expense")
    depreciation = amount("depreciation_and_amortization_expense")
    other_expenses = amount("other_expenses")
    current_tax = amount("current_tax")
    prior_period = amount("prior_period_items")

    profit_before_tax = totals.profit_before_tax
    profit_after_tax = totals.profit_after_tax

    face_map: dict[str, tuple[str, float | None]] = {
        "share_capital": ("F13", share_capital),
        "reserves_and_surplus": ("F14", reserves_face if reserves_face else None),
        "trade_payables": ("F26", trade_payables),
        "other_current_liabilities": ("F29", other_cl if other_cl else None),
        "short_term_provisions": ("F30", short_term_provisions if short_term_provisions else None),
        "property_plant_and_equipment": ("F38", ppe if ppe else None),
        "intangible_assets": ("F39", intangible if intangible else None),
        "trade_receivables": ("F46", debtors if debtors else None),
        "cash_and_cash_equivalents": ("F47", cash if cash else None),
        "short_term_loans_and_advances": ("F48", advances if advances else None),
        "other_current_assets": ("F49", other_ca if other_ca else None),
        "revenue_from_operations": ("F83", revenue if revenue else None),
        "other_income": ("F84", other_income if other_income else None),
        "depreciation_and_amortization_expense": ("F95", depreciation if depreciation else None),
        "employee_benefits_expense": ("F96", salary if salary else None),
        "other_expenses": ("F97", other_expenses if other_expenses else None),
        "current_tax": ("F107", current_tax if current_tax else None),
        "profit_for_period": ("F109", profit_after_tax),
    }

    for field_key, (cell, value) in face_map.items():
        if value is None:
            continue
        writes.append(CellWrite("BS PnL", cell, round(value, 2), field_key))

    # ---- BS sub-totals and grand totals ----------------------------------- #
    # Equity & Liabilities side
    total_shareholders_funds = round(share_capital + reserves_face, 2)
    total_non_current_liab   = 0.0          # no LT borrowings in this TB
    total_current_liab       = round(trade_payables + other_cl + short_term_provisions, 2)
    total_equity_liab        = round(total_shareholders_funds + total_non_current_liab + total_current_liab, 2)

    # Assets side
    total_non_current_assets = round(ppe + intangible, 2)
    total_current_assets     = round(debtors + cash + advances + other_ca, 2)
    total_assets             = round(total_non_current_assets + total_current_assets, 2)

    bs_totals: list[tuple[str, float, str]] = [
        # (cell, value, field_key)
        ("F15",  total_shareholders_funds,    "subtotal_shareholders_funds"),   # sub-total row after Reserves
        ("F17",  total_shareholders_funds,    "subtotal_shareholders_funds2"),  # blank row before non-current
        ("F23",  total_non_current_liab,      "subtotal_non_current_liab"),
        ("F31",  total_current_liab,          "subtotal_current_liab"),
        ("F32",  total_equity_liab,           "total_equity_liab"),             # TOTAL row
        ("F43",  total_non_current_assets,    "subtotal_non_current_assets"),
        ("F50",  total_current_assets,        "subtotal_current_assets"),
        ("F51",  total_assets,                "total_assets"),                  # TOTAL row
        # Trade payables split (MSME / non-MSME) — all goes to non-MSME
        ("F27",  0.0,                         "tp_msme"),
        ("F28",  trade_payables,              "tp_non_msme"),
    ]
    for cell, value, fk in bs_totals:
        writes.append(CellWrite("BS PnL", cell, round(value, 2), fk))

    # ---- P&L sub-totals ---------------------------------------------------- #
    total_income    = round(revenue + other_income, 2)
    # Expenses that flow into total (no stock / material for service company)
    fin_costs       = 0.0   # finance costs written separately in notes; not in face_map
    total_expenses  = round(depreciation + salary + other_expenses + fin_costs, 2)
    pbt             = round(total_income - total_expenses, 2)
    # Prior period net (negative in P&L expenses means income, add back)
    prior_net       = round(prior_period, 2)          # already signed correctly
    pbt_after_prior = round(pbt + prior_net, 2)       # prior period reduces PBT if positive expense
    deferred_tax    = round(amount("deferred_tax_income"), 2)

    pl_totals: list[tuple[str, float, str]] = [
        ("F86",  total_income,             "total_income"),
        ("F99",  total_expenses,           "total_expenses"),
        ("F100", round(pbt, 2),            "pbt_before_prior"),
        ("F101", round(-prior_net, 2),     "prior_period_items"),   # negative = expense sign
        ("F103", round(totals.profit_before_tax, 2), "pbt"),
        ("F104", 0.0,                      "exceptional_items"),
        ("F105", round(totals.profit_before_tax, 2), "pbt_after_exceptional"),
        ("F108", round(deferred_tax, 2),   "deferred_tax"),
        ("F109", round(profit_after_tax, 2), "profit_for_period"),
    ]
    for cell, value, fk in pl_totals:
        writes.append(CellWrite("BS PnL", cell, value, fk))

    if classification.company_name:
        writes.append(CellWrite("BS PnL", "A2", classification.company_name, "company_name"))
        writes.append(CellWrite("BS PnL", "E60", classification.company_name, "company_name"))

    if classification.period_label:
        writes.append(
            CellWrite(
                "BS PnL",
                "A6",
                f"Balance Sheet as at {classification.period_label.split('to')[-1].strip()}",
                "reporting_period",
            )
        )
        writes.append(
            CellWrite(
                "BS PnL",
                "A79",
                f"Statement of Profit and Loss for {classification.period_label}",
                "reporting_period",
            )
        )

    if share_capital:
        writes.append(CellWrite("Note 3 (Share Capital)", "H13", share_capital, "share_capital"))
        share_count = int(share_capital / DEFAULT_FACE_VALUE) if DEFAULT_FACE_VALUE else None
        if share_count:
            writes.append(CellWrite("Note 3 (Share Capital)", "F12", share_count, "share_count"))
            writes.append(CellWrite("Note 3 (Share Capital)", "H12", share_capital, "share_capital"))
        for cell in ("F20", "F21", "F22", "F23"):
            writes.append(CellWrite("Note 3 (Share Capital)", cell, NOT_AVAILABLE, "share_detail"))

    if reserves_tb:
        writes.append(CellWrite("NOTE (4-12)", "E16", reserves_tb, "reserves_and_surplus"))
    if current_year_profit:
        writes.append(CellWrite("NOTE (4-12)", "E11", round(current_year_profit, 2), "profit_for_period"))

    if revenue:
        writes.append(CellWrite("Note 20-31", "C10", revenue, "revenue_from_operations"))
        writes.append(CellWrite("Note 20-31", "C11", revenue, "revenue_from_operations"))
    if other_income:
        writes.append(CellWrite("Note 20-31", "C23", other_income, "other_income"))
    if salary:
        writes.append(CellWrite("Note 20-31", "C71", salary, "employee_benefits_expense"))
    if depreciation:
        writes.append(CellWrite("Note 20-31", "C82", depreciation, "depreciation_and_amortization_expense"))
    if other_expenses:
        writes.append(CellWrite("Note 20-31", "C93", other_expenses, "other_expenses"))
    if current_tax:
        writes.append(CellWrite("Note 20-31", "C124", current_tax, "current_tax"))

    if debtors:
        writes.append(CellWrite("Note (13-20)", "H46", debtors, "trade_receivables"))
    if cash:
        writes.append(CellWrite("Note (13-20)", "F63", cash, "cash_and_cash_equivalents"))
    if advances:
        writes.append(CellWrite("Note (13-20)", "F72", advances, "short_term_loans_and_advances"))
    if other_ca:
        writes.append(CellWrite("Note (13-20)", "F88", other_ca, "other_current_assets"))

    if ppe:
        writes.append(CellWrite("Note 12 (PPE)", "J20", ppe, "property_plant_and_equipment"))
    if intangible:
        writes.append(CellWrite("Note 12 (PPE)", "J28", intangible, "intangible_assets"))

    if classification.company_name:
        writes.append(CellWrite("Note 1-2", "A2", classification.company_name, "company_name"))
        writes.append(
            CellWrite(
                "Note 1-2",
                "B10",
                (
                    f"{classification.company_name} ('the Company') — financial statements prepared "
                    f"solely from the uploaded trial balance for {classification.period_label or 'the current period'}."
                ),
                "company_narrative",
            )
        )

    for sheet, cell in _comparative_cells():
        writes.append(CellWrite(sheet, cell, "N.A.", "comparative"))

    return writes


def _comparative_cells() -> list[tuple[str, str]]:
    cells: list[tuple[str, str]] = []
    for row in range(13, 32):
        cells.append(("BS PnL", f"G{row}"))
    for row in (83, 84, 95, 96, 97, 107, 109):
        cells.append(("BS PnL", f"G{row}"))
    for row in range(10, 17):
        cells.append(("NOTE (4-12)", f"F{row}"))
    return cells


_FIELD_KEY_LABELS: dict[str, tuple[str, str]] = {
    # field_key -> (human-readable source label, Schedule III category)
    "share_capital":                        ("Share Capital",                         "Shareholders' Funds — Share Capital"),
    "share_count":                          ("Share Capital — No. of Shares",         "Shareholders' Funds — Share Capital"),
    "share_detail":                         ("Share Capital — Disclosure Detail",     "Shareholders' Funds — Share Capital"),
    "reserves_and_surplus":                 ("Reserves and Surplus",                  "Shareholders' Funds — Reserves & Surplus"),
    "profit_for_period":                    ("Profit / (Loss) for the Period",        "Shareholders' Funds — Reserves & Surplus"),
    "trade_payables":                       ("Trade Payables",                        "Current Liabilities — Trade Payables"),
    "other_current_liabilities":            ("Other Current Liabilities",             "Current Liabilities — Other Current Liabilities"),
    "short_term_provisions":                ("Short-term Provisions",                 "Current Liabilities — Short-term Provisions"),
    "property_plant_and_equipment":         ("Property, Plant & Equipment",           "Non-Current Assets — PPE"),
    "intangible_assets":                    ("Intangible Assets",                     "Non-Current Assets — Intangible Assets"),
    "trade_receivables":                    ("Trade Receivables",                     "Current Assets — Trade Receivables"),
    "cash_and_cash_equivalents":            ("Cash and Cash Equivalents",             "Current Assets — Cash & Bank"),
    "short_term_loans_and_advances":        ("Short-term Loans and Advances",         "Current Assets — Loans & Advances"),
    "other_current_assets":                 ("Other Current Assets",                  "Current Assets — Other Current Assets"),
    "revenue_from_operations":              ("Revenue from Operations",               "Revenue — Operations"),
    "other_income":                         ("Other Income",                          "Revenue — Other Income"),
    "employee_benefits_expense":            ("Employee Benefits Expense",             "Expenses — Employee Benefits"),
    "depreciation_and_amortization_expense":("Depreciation & Amortisation",          "Expenses — Depreciation"),
    "other_expenses":                       ("Other Expenses",                        "Expenses — Other Expenses"),
    "current_tax":                          ("Current Tax",                           "Tax — Current Tax"),
    "deferred_tax":                         ("Deferred Tax",                          "Tax — Deferred Tax"),
    "company_name":                         ("Company Name",                          "Header — Company Identity"),
    "company_narrative":                    ("Company Description",                   "Notes — Note 1"),
    "reporting_period":                     ("Reporting Period",                      "Header — Reporting Period"),
    "comparative":                          ("Comparative (Prior Year)",              "N.A. — Prior Year"),
    # Sub-totals and grand totals
    "subtotal_shareholders_funds":          ("Total Shareholders' Funds",             "Shareholders' Funds — Sub-Total"),
    "subtotal_shareholders_funds2":         ("Total Shareholders' Funds",             "Shareholders' Funds — Sub-Total"),
    "subtotal_non_current_liab":            ("Total Non-Current Liabilities",         "Non-Current Liabilities — Sub-Total"),
    "subtotal_current_liab":               ("Total Current Liabilities",             "Current Liabilities — Sub-Total"),
    "total_equity_liab":                    ("Total Equity and Liabilities",          "Balance Sheet — Grand Total"),
    "subtotal_non_current_assets":          ("Total Non-Current Assets",             "Non-Current Assets — Sub-Total"),
    "subtotal_current_assets":             ("Total Current Assets",                  "Current Assets — Sub-Total"),
    "total_assets":                         ("Total Assets",                          "Balance Sheet — Grand Total"),
    "tp_msme":                              ("Trade Payables — MSME",                "Current Liabilities — Trade Payables"),
    "tp_non_msme":                          ("Trade Payables — Non-MSME",            "Current Liabilities — Trade Payables"),
    "total_income":                         ("Total Income",                          "P&L — Total Income"),
    "total_expenses":                       ("Total Expenses",                        "P&L — Total Expenses"),
    "pbt_before_prior":                     ("PBT before Prior Period Items",         "P&L — PBT"),
    "prior_period_items":                   ("Prior Period Items (Net)",              "P&L — Prior Period"),
    "pbt":                                  ("Profit Before Tax",                     "P&L — PBT"),
    "exceptional_items":                    ("Exceptional Items",                     "P&L — Exceptional"),
    "pbt_after_exceptional":                ("Profit Before Tax",                     "P&L — PBT After Exceptional"),
    "deferred_tax":                         ("Deferred Tax",                          "Tax — Deferred Tax"),
}


def writes_to_placements(writes: list[CellWrite]) -> list[dict[str, Any]]:
    return [
        {
            "field_key":            item.field_key,
            "excel_sheet":          item.sheet,
            "excel_cell":           item.cell,
            "extracted_value":      item.value,
            "action":               "write",
            "overwrite_formula":    item.overwrite_formula,
            "review_status":        "approved",
            "period":               "current",
            "value_role":           "text" if isinstance(item.value, str) else "value",
            # ---- Human-readable metadata for the review UI ----
            "source_label":         _FIELD_KEY_LABELS.get(item.field_key, (item.field_key, ""))[0],
            "schedule_iii_label":   _FIELD_KEY_LABELS.get(item.field_key, (item.field_key, ""))[0],
            "schedule_iii_category":_FIELD_KEY_LABELS.get(item.field_key, ("", item.field_key))[1],
            "source_page":          "Trial Balance",
            "confidence":           1.0,
        }
        for item in writes
    ]
