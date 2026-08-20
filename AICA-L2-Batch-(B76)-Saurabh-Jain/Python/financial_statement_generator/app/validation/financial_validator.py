"""Validate extracted financial statements without changing any figures.

Arithmetic identities are checked independently for the current year and the
previous year. Missing values stay missing; the validator never plugs zeros
or restates amounts to force a statement to balance.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any, Iterable

from app.mapping.field_mapping import normalize_label, slugify

PERIODS = ("current", "previous")
ABS_TOLERANCE = 1.0
REL_TOLERANCE = 1e-4

TOTAL_ASSET_LABELS = ("total assets", "total asset")
TOTAL_EQUITY_LIAB_LABELS = (
    "total equity and liabilities",
    "total equities and liabilities",
    "equity and liabilities",
)
CURRENT_ASSET_TOTAL_LABELS = ("current assets",)
NON_CURRENT_ASSET_TOTAL_LABELS = ("non current assets", "non-current assets", "non current asset")
CURRENT_LIAB_TOTAL_LABELS = ("current liabilities",)
NON_CURRENT_LIAB_TOTAL_LABELS = ("non-current liabilities", "non current liabilities")
EQUITY_TOTAL_LABELS = ("shareholders' funds", "shareholders funds", "equity", "total equity")
TOTAL_INCOME_LABELS = ("total income", "total income (i + ii)", "total revenue")
TOTAL_EXPENSE_LABELS = ("total expenses", "total expense")
PBT_LABELS = (
    "profit before tax",
    "profit before extraordinary items and tax",
    "profit before exceptional and extraordinary items and tax",
)
PAT_LABELS = (
    "profit/ (loss)",
    "profit (loss) for the period",
    "profit for the period",
    "profit/(loss)",
    "profit / (loss)",
)
TAX_LABELS = ("tax expense",)
CURRENT_TAX_LABELS = ("current tax",)
DEFERRED_TAX_LABELS = ("deferred tax",)
BASIC_EPS_LABELS = ("basic", "basic eps", "earnings per equity share")
DILUTED_EPS_LABELS = ("diluted", "diluted eps")
SHARE_COUNT_LABELS = (
    "weighted average number of equity shares",
    "number of equity shares",
    "weighted average shares",
)

CURRENT_ASSET_KEYS = (
    "current_investments",
    "inventories",
    "trade_receivables",
    "cash_and_cash_equivalents",
    "short_term_loans_and_advances",
    "other_current_assets",
)
NON_CURRENT_ASSET_KEYS = (
    "property_plant_and_equipment",
    "intangible_assets",
    "capital_work_in_progress",
    "intangible_assets_under_development",
    "non_current_investments",
    "deferred_tax_assets_net",
    "long_term_loans_and_advances",
    "other_non_current_assets",
)
CURRENT_LIAB_KEYS = (
    "short_term_borrowings",
    "trade_payables",
    "other_current_liabilities",
    "short_term_provisions",
)
NON_CURRENT_LIAB_KEYS = (
    "long_term_borrowings",
    "deferred_tax_liabilities_net",
    "other_long_term_liabilities",
    "long_term_provisions",
)
EQUITY_KEYS = (
    "share_capital",
    "reserves_and_surplus",
    "money_received_against_share_warrants",
    "share_application_money_pending_allotment",
)
INCOME_KEYS = ("revenue_from_operations", "other_income")
EXPENSE_KEYS = (
    "cost_of_materials_consumed",
    "purchases_of_stock_in_trade",
    "changes_in_inventories_of_finished_goods_work_in_progress_and_stock_in_trade",
    "employee_benefits_expense",
    "finance_costs",
    "depreciation_and_amortization_expense",
    "other_expenses",
)
FACE_ORIGINS = {"balance_sheet", "profit_and_loss"}
NOTE_ORIGINS = {
    "share_capital",
    "reserves_and_surplus",
    "borrowings",
    "ppe",
    "cwip",
    "investments",
    "inventory",
    "trade_receivables",
    "trade_payables",
    "cash_and_cash_equivalents",
    "revenue",
    "expenses",
    "tax_expense",
    "eps",
    "notes_to_accounts",
    "notes_to_accounts.notes",
}
NOTE_RECONCILE_KEYS = (
    "share_capital",
    "reserves_and_surplus",
    "property_plant_and_equipment",
    "capital_work_in_progress",
    "inventories",
    "trade_receivables",
    "trade_payables",
    "cash_and_cash_equivalents",
    "long_term_borrowings",
    "short_term_borrowings",
    "revenue_from_operations",
    "tax_expense",
)
MANDATORY_IF_BS = ("share_capital",)
MANDATORY_IF_PNL = ("revenue_from_operations",)
NON_NEGATIVE_ERROR_KEYS = ("share_capital", "total_assets")
NON_NEGATIVE_WARNING_KEYS = (
    "revenue_from_operations",
    "inventories",
    "trade_receivables",
    "property_plant_and_equipment",
    "cash_and_cash_equivalents",
    "employee_benefits_expense",
    "finance_costs",
    "other_expenses",
)

BLOCK_KEYS = (
    "balance_sheet",
    "profit_and_loss",
    "cash_flow",
    "notes_to_accounts",
    "share_capital",
    "reserves_and_surplus",
    "borrowings",
    "ppe",
    "cwip",
    "investments",
    "inventory",
    "trade_receivables",
    "trade_payables",
    "cash_and_cash_equivalents",
    "revenue",
    "expenses",
    "tax_expense",
    "eps",
    "other_disclosures",
    "ratios",
)


@dataclass(slots=True)
class Amount:
    key: str
    label: str
    value: float
    period: str
    origin: str
    source_text: str | None = None


@dataclass(slots=True)
class Ledger:
    """Numeric facts indexed by field key and period. Never fills gaps with zero."""

    items: list[Amount] = field(default_factory=list)

    def add(self, amount: Amount) -> None:
        self.items.append(amount)

    def get(self, key: str, period: str, *, origins: Iterable[str] | None = None) -> Amount | None:
        wanted = {normalize_label(key), slugify(key), key}
        allowed = set(origins) if origins is not None else None
        for item in self.items:
            if item.period != period:
                continue
            if allowed is not None and item.origin not in allowed:
                continue
            if item.key in wanted or normalize_label(item.label) in wanted or slugify(item.label) in wanted:
                return item
        return None

    def value(self, key: str, period: str, *, origins: Iterable[str] | None = None) -> float | None:
        found = self.get(key, period, origins=origins)
        return None if found is None else found.value

    def lookup(self, labels: Iterable[str], period: str, *, origins: Iterable[str] | None = None) -> Amount | None:
        for label in labels:
            found = self.get(label, period, origins=origins)
            if found is not None:
                return found
        return None

    def sum_keys(self, keys: Iterable[str], period: str, *, origins: Iterable[str] | None = None) -> tuple[float | None, list[str]]:
        total = 0.0
        used: list[str] = []
        seen: set[str] = set()
        for key in keys:
            found = self.get(key, period, origins=origins)
            if found is None or found.key in seen:
                continue
            seen.add(found.key)
            total += found.value
            used.append(found.label)
        if not used:
            return None, []
        return total, used

    def has_statement(self, origins: Iterable[str]) -> bool:
        allowed = set(origins)
        return any(item.origin in allowed for item in self.items)


class FinancialValidator:
    """Run Schedule III arithmetic and completeness checks. Read-only."""

    def validate(self, mapped_data: dict | Any, classified_data: dict | Any | None = None) -> dict:
        """Validate classified/mapped data and return PASS / WARNING / ERROR."""
        classified, placements = _split_inputs(mapped_data, classified_data)
        ledger = build_ledger(classified, placements)
        checks: list[dict[str, Any]] = []
        warnings: list[dict[str, Any]] = []
        errors: list[dict[str, Any]] = []

        for period in PERIODS:
            checks.extend(self._balance_sheet_checks(ledger, period))
            checks.extend(self._profit_and_loss_checks(ledger, period))
            checks.extend(self._eps_checks(ledger, period))
            checks.extend(self._note_reconciliations(ledger, period))
            checks.extend(self._sign_checks(ledger, period))

        checks.extend(self._missing_mandatory(ledger, classified))
        checks.extend(self._duplicate_mapping_checks(placements))
        checks.extend(self._period_separation_checks(ledger))

        for check in checks:
            issue = {
                "code": check["id"],
                "message": check["message"],
                "period": check.get("period"),
                "details": check.get("details") or {},
            }
            if check["status"] == "ERROR":
                errors.append(issue)
            elif check["status"] == "WARNING":
                warnings.append(issue)

        if errors:
            status = "ERROR"
        elif warnings:
            status = "WARNING"
        else:
            status = "PASS"
        return {
            "status": status,
            "checks": checks,
            "warnings": warnings,
            "errors": errors,
        }

    def _balance_sheet_checks(self, ledger: Ledger, period: str) -> list[dict[str, Any]]:
        checks: list[dict[str, Any]] = []
        total_assets = ledger.lookup(TOTAL_ASSET_LABELS, period, origins=FACE_ORIGINS) or ledger.lookup(
            TOTAL_ASSET_LABELS, period
        )
        total_equity_liab = ledger.lookup(TOTAL_EQUITY_LIAB_LABELS, period, origins=FACE_ORIGINS) or ledger.lookup(
            TOTAL_EQUITY_LIAB_LABELS, period
        )
        current_assets = _group_total(
            ledger,
            period,
            explicit=CURRENT_ASSET_TOTAL_LABELS,
            children=CURRENT_ASSET_KEYS,
            origins=FACE_ORIGINS,
        )
        non_current_assets = _group_total(
            ledger,
            period,
            explicit=NON_CURRENT_ASSET_TOTAL_LABELS,
            children=NON_CURRENT_ASSET_KEYS,
            origins=FACE_ORIGINS,
        )
        current_liab = _group_total(
            ledger,
            period,
            explicit=CURRENT_LIAB_TOTAL_LABELS,
            children=CURRENT_LIAB_KEYS,
            origins=FACE_ORIGINS,
        )
        non_current_liab = _group_total(
            ledger,
            period,
            explicit=NON_CURRENT_LIAB_TOTAL_LABELS,
            children=NON_CURRENT_LIAB_KEYS,
            origins=FACE_ORIGINS,
        )
        equity = _group_total(
            ledger,
            period,
            explicit=EQUITY_TOTAL_LABELS,
            children=EQUITY_KEYS,
            origins=FACE_ORIGINS,
        )

        checks.append(
            _identity_check(
                check_id="bs_assets_equal_equity_and_liabilities",
                period=period,
                left_name="Total Assets",
                left_value=None if total_assets is None else total_assets.value,
                right_name="Total Equity and Liabilities",
                right_value=None if total_equity_liab is None else total_equity_liab.value,
                formula="Total Assets = Total Equity and Liabilities",
            )
        )
        asset_sum, asset_parts = _sum_optional(current_assets, non_current_assets)
        checks.append(
            _identity_check(
                check_id="bs_current_plus_non_current_assets",
                period=period,
                left_name="Current Assets + Non-current Assets",
                left_value=asset_sum,
                right_name="Total Assets",
                right_value=None if total_assets is None else total_assets.value,
                formula="Current Assets + Non-current Assets = Total Assets",
                details={"components": asset_parts},
            )
        )
        equity_liab_sum, equity_liab_parts = _sum_optional(current_liab, non_current_liab, equity)
        checks.append(
            _identity_check(
                check_id="bs_equity_plus_liabilities",
                period=period,
                left_name="Current Liabilities + Non-current Liabilities + Equity",
                left_value=equity_liab_sum,
                right_name="Total Equity and Liabilities",
                right_value=None if total_equity_liab is None else total_equity_liab.value,
                formula=(
                    "Current Liabilities + Non-current Liabilities + Equity "
                    "= Total Equity and Liabilities"
                ),
                details={"components": equity_liab_parts},
            )
        )
        return checks

    def _profit_and_loss_checks(self, ledger: Ledger, period: str) -> list[dict[str, Any]]:
        income = _group_total(ledger, period, explicit=TOTAL_INCOME_LABELS, children=INCOME_KEYS, origins=("profit_and_loss", "revenue"))
        expenses = _group_total(
            ledger,
            period,
            explicit=TOTAL_EXPENSE_LABELS,
            children=EXPENSE_KEYS,
            origins=("profit_and_loss", "expenses"),
        )
        pbt = ledger.lookup(PBT_LABELS, period)
        tax = ledger.lookup(TAX_LABELS, period)
        if tax is None:
            tax_sum, tax_parts = ledger.sum_keys(("current_tax", "deferred_tax", *CURRENT_TAX_LABELS, *DEFERRED_TAX_LABELS), period)
            tax_value = tax_sum
            tax_detail = {"components": tax_parts}
        else:
            tax_value = tax.value
            tax_detail = {"components": [tax.label]}
        pat = ledger.lookup(PAT_LABELS, period)

        income_minus_expenses = None
        income_parts: dict[str, float] = {}
        if income is not None and expenses is not None:
            income_minus_expenses = income["value"] - expenses["value"]
            income_parts = {"Total Income": income["value"], "Total Expenses": expenses["value"]}
        checks = [
            _identity_check(
                check_id="pnl_income_minus_expenses_equals_pbt",
                period=period,
                left_name="Total Income - Total Expenses",
                left_value=income_minus_expenses,
                right_name="Profit Before Tax",
                right_value=None if pbt is None else pbt.value,
                formula="Total Income - Total Expenses = Profit Before Tax",
                details={"components": income_parts},
            )
        ]
        pbt_minus_tax = None
        if pbt is not None and tax_value is not None:
            pbt_minus_tax = pbt.value - tax_value
        checks.append(
            _identity_check(
                check_id="pnl_pbt_minus_tax_equals_pat",
                period=period,
                left_name="Profit Before Tax - Tax Expense",
                left_value=pbt_minus_tax,
                right_name="Profit/Loss for the Period",
                right_value=None if pat is None else pat.value,
                formula="Profit Before Tax - Tax Expense = Profit/Loss for the Period",
                details=tax_detail,
            )
        )
        return checks

    def _eps_checks(self, ledger: Ledger, period: str) -> list[dict[str, Any]]:
        basic = ledger.lookup(BASIC_EPS_LABELS, period, origins=("profit_and_loss", "eps"))
        diluted = ledger.lookup(DILUTED_EPS_LABELS, period, origins=("profit_and_loss", "eps"))
        pat = ledger.lookup(PAT_LABELS, period)
        shares = ledger.lookup(SHARE_COUNT_LABELS, period)
        if basic is None and diluted is None and shares is None:
            return [
                _check(
                    "eps_consistency",
                    period,
                    "SKIPPED",
                    "EPS consistency was not checked because basic EPS, diluted EPS, "
                    f"and share counts were not extracted for the {period} period.",
                )
            ]
        checks: list[dict[str, Any]] = []
        if basic is not None and diluted is not None:
            if (basic.value >= 0 and diluted.value > basic.value + ABS_TOLERANCE) or (
                basic.value < 0 and diluted.value < basic.value - ABS_TOLERANCE
            ):
                checks.append(
                    _check(
                        "eps_diluted_vs_basic",
                        period,
                        "ERROR",
                        (
                            f"Diluted EPS ({_fmt(diluted.value)}) is inconsistent with Basic EPS "
                            f"({_fmt(basic.value)}) for the {period} period. Diluted EPS should not "
                            "exceed Basic EPS when profit is positive, and should not be more "
                            "negative than Basic EPS when there is a loss."
                        ),
                        details={"basic": basic.value, "diluted": diluted.value},
                    )
                )
            elif _same_sign(basic.value, diluted.value) is False and not (
                nearly_equal(basic.value, 0) or nearly_equal(diluted.value, 0)
            ):
                checks.append(
                    _check(
                        "eps_diluted_vs_basic",
                        period,
                        "ERROR",
                        (
                            f"Basic EPS ({_fmt(basic.value)}) and Diluted EPS ({_fmt(diluted.value)}) "
                            f"have opposite signs for the {period} period."
                        ),
                        details={"basic": basic.value, "diluted": diluted.value},
                    )
                )
            else:
                checks.append(
                    _check(
                        "eps_diluted_vs_basic",
                        period,
                        "PASS",
                        (
                            f"Basic EPS ({_fmt(basic.value)}) and Diluted EPS ({_fmt(diluted.value)}) "
                            f"are consistent for the {period} period."
                        ),
                        details={"basic": basic.value, "diluted": diluted.value},
                    )
                )
        if basic is not None and pat is not None and shares is not None and shares.value:
            implied = pat.value / shares.value
            if nearly_equal(implied, basic.value):
                checks.append(
                    _check(
                        "eps_vs_pat",
                        period,
                        "PASS",
                        (
                            f"Basic EPS ({_fmt(basic.value)}) agrees with Profit/Loss "
                            f"({_fmt(pat.value)}) / shares ({_fmt(shares.value)}) = {_fmt(implied)} "
                            f"for the {period} period."
                        ),
                        details={"basic": basic.value, "implied": implied, "pat": pat.value, "shares": shares.value},
                    )
                )
            else:
                checks.append(
                    _check(
                        "eps_vs_pat",
                        period,
                        "ERROR",
                        (
                            f"Basic EPS is {_fmt(basic.value)} but Profit/Loss ({_fmt(pat.value)}) / "
                            f"shares ({_fmt(shares.value)}) is {_fmt(implied)} for the {period} period. "
                            f"Difference = {_fmt(basic.value - implied)}."
                        ),
                        details={"basic": basic.value, "implied": implied, "pat": pat.value, "shares": shares.value},
                    )
                )
        elif basic is not None and pat is not None:
            if _same_sign(basic.value, pat.value) is False and not (
                nearly_equal(basic.value, 0) or nearly_equal(pat.value, 0)
            ):
                checks.append(
                    _check(
                        "eps_sign_vs_pat",
                        period,
                        "ERROR",
                        (
                            f"Basic EPS ({_fmt(basic.value)}) and Profit/Loss for the Period "
                            f"({_fmt(pat.value)}) have opposite signs for the {period} period."
                        ),
                        details={"basic": basic.value, "pat": pat.value},
                    )
                )
            else:
                checks.append(
                    _check(
                        "eps_sign_vs_pat",
                        period,
                        "PASS",
                        (
                            f"Basic EPS ({_fmt(basic.value)}) has the same sign as Profit/Loss "
                            f"({_fmt(pat.value)}) for the {period} period. Per-share recomputation "
                            "was skipped because the share count was not extracted."
                        ),
                        details={"basic": basic.value, "pat": pat.value},
                    )
                )
        if not checks:
            checks.append(
                _check(
                    "eps_consistency",
                    period,
                    "SKIPPED",
                    f"Insufficient EPS inputs were extracted for the {period} period.",
                )
            )
        return checks

    def _note_reconciliations(self, ledger: Ledger, period: str) -> list[dict[str, Any]]:
        checks: list[dict[str, Any]] = []
        for key in NOTE_RECONCILE_KEYS:
            face = ledger.get(key, period, origins=FACE_ORIGINS)
            note = ledger.get(key, period, origins=NOTE_ORIGINS)
            if note is None:
                continue
            if face is None:
                checks.append(
                    _check(
                        f"note_reconcile_{key}",
                        period,
                        "WARNING",
                        (
                            f"Note amount for '{note.label}' is {_fmt(note.value)} for the {period} "
                            "period, but the corresponding Balance Sheet / P&L caption was not "
                            "extracted, so the note could not be reconciled."
                        ),
                        details={"field": key, "face": None, "note": note.value},
                    )
                )
                continue
            if nearly_equal(face.value, note.value):
                checks.append(
                    _check(
                        f"note_reconcile_{key}",
                        period,
                        "PASS",
                        (
                            f"Note total for '{face.label}' ({_fmt(note.value)}) reconciles with "
                            f"the face amount ({_fmt(face.value)}) for the {period} period."
                        ),
                        details={"field": key, "face": face.value, "note": note.value},
                    )
                )
            else:
                checks.append(
                    _check(
                        f"note_reconcile_{key}",
                        period,
                        "ERROR",
                        (
                            f"Note total for '{face.label}' is {_fmt(note.value)} but the "
                            f"Balance Sheet / P&L amount is {_fmt(face.value)} for the {period} "
                            f"period. Difference = {_fmt(face.value - note.value)}. "
                            "Values were not changed."
                        ),
                        details={"field": key, "face": face.value, "note": note.value},
                    )
                )
        return checks

    def _sign_checks(self, ledger: Ledger, period: str) -> list[dict[str, Any]]:
        checks: list[dict[str, Any]] = []
        for key in NON_NEGATIVE_ERROR_KEYS:
            found = ledger.get(key, period) or ledger.lookup(TOTAL_ASSET_LABELS if key == "total_assets" else (key,), period)
            if found is None or found.value >= -ABS_TOLERANCE:
                continue
            checks.append(
                _check(
                    f"sign_{key}",
                    period,
                    "ERROR",
                    (
                        f"'{found.label}' is {_fmt(found.value)} for the {period} period, which is "
                        "negative. This is inconsistent with the usual Schedule III presentation "
                        "for that caption. The value was not altered."
                    ),
                    details={"field": key, "value": found.value},
                )
            )
        for key in NON_NEGATIVE_WARNING_KEYS:
            found = ledger.get(key, period)
            if found is None or found.value >= -ABS_TOLERANCE:
                continue
            checks.append(
                _check(
                    f"sign_{key}",
                    period,
                    "WARNING",
                    (
                        f"'{found.label}' is {_fmt(found.value)} for the {period} period. A negative "
                        "amount may be a credit balance or a sign-mapping error. The value was not altered."
                    ),
                    details={"field": key, "value": found.value},
                )
            )
        return checks

    def _missing_mandatory(self, ledger: Ledger, classified: dict[str, Any]) -> list[dict[str, Any]]:
        checks: list[dict[str, Any]] = []
        bs_present = _block_identified(classified, "balance_sheet") or ledger.has_statement({"balance_sheet"})
        pnl_present = _block_identified(classified, "profit_and_loss") or ledger.has_statement({"profit_and_loss", "revenue", "expenses"})
        for period in PERIODS:
            if bs_present:
                for key in MANDATORY_IF_BS:
                    if ledger.get(key, period) is None:
                        checks.append(
                            _check(
                                f"missing_{key}",
                                period,
                                "WARNING",
                                (
                                    f"Mandatory Balance Sheet caption '{key.replace('_', ' ')}' was not "
                                    f"extracted for the {period} period."
                                ),
                                details={"field": key},
                            )
                        )
            if pnl_present:
                for key in MANDATORY_IF_PNL:
                    if ledger.get(key, period) is None and ledger.lookup(TOTAL_INCOME_LABELS, period) is None:
                        checks.append(
                            _check(
                                f"missing_{key}",
                                period,
                                "WARNING",
                                (
                                    f"Mandatory Profit and Loss caption '{key.replace('_', ' ')}' was not "
                                    f"extracted for the {period} period."
                                ),
                                details={"field": key},
                            )
                        )
        return checks

    def _duplicate_mapping_checks(self, placements: list[dict[str, Any]]) -> list[dict[str, Any]]:
        if not placements:
            return [
                _check(
                    "duplicate_mappings",
                    None,
                    "SKIPPED",
                    "Duplicate mapping detection was skipped because no Excel placements were provided.",
                )
            ]
        checks: list[dict[str, Any]] = []
        by_cell: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
        by_field: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
        for item in placements:
            sheet = item.get("excel_sheet") or ""
            cell = (item.get("excel_cell") or "").upper()
            field_key = item.get("field_key") or ""
            period = item.get("period") or ""
            if sheet and cell:
                by_cell[(sheet, cell)].append(item)
            if field_key and period:
                by_field[(field_key, period)].append(item)
        for (sheet, cell), group in by_cell.items():
            keys = {item.get("field_key") for item in group}
            values = {item.get("extracted_value") for item in group}
            if len(keys) > 1:
                checks.append(
                    _check(
                        "duplicate_cell_mapping",
                        None,
                        "ERROR",
                        (
                            f"{sheet}!{cell} is targeted by multiple fields: "
                            f"{sorted(k for k in keys if k)}. Extracted values = {sorted(_fmt(v) for v in values)}. "
                            "The validator did not choose a winner or change any value."
                        ),
                        details={"sheet": sheet, "cell": cell, "fields": sorted(k for k in keys if k)},
                    )
                )
        for (field_key, period), group in by_field.items():
            numeric = [item.get("extracted_value") for item in group if isinstance(item.get("extracted_value"), (int, float))]
            unique_values = {round(float(value), 6) for value in numeric}
            destinations = {(item.get("excel_sheet"), item.get("excel_cell")) for item in group if item.get("excel_cell")}
            if len(unique_values) > 1:
                checks.append(
                    _check(
                        "duplicate_field_values",
                        period,
                        "ERROR",
                        (
                            f"Field '{field_key}' has conflicting {period} amounts {sorted(unique_values)}. "
                            "Values were left unchanged."
                        ),
                        details={"field": field_key, "values": sorted(unique_values)},
                    )
                )
            elif len(destinations) > 1:
                checks.append(
                    _check(
                        "duplicate_field_destinations",
                        period,
                        "WARNING",
                        (
                            f"Field '{field_key}' ({period}) is mapped to multiple Excel cells: "
                            f"{sorted(f'{sheet}!{cell}' for sheet, cell in destinations if sheet and cell)}."
                        ),
                        details={"field": field_key, "destinations": sorted(destinations)},
                    )
                )
        if not any(check["id"].startswith("duplicate_") and check["status"] != "SKIPPED" for check in checks):
            checks.append(
                _check(
                    "duplicate_mappings",
                    None,
                    "PASS",
                    "No duplicate Excel destinations or conflicting field values were found.",
                )
            )
        return checks

    def _period_separation_checks(self, ledger: Ledger) -> list[dict[str, Any]]:
        mixed = [item for item in ledger.items if item.period not in PERIODS and item.period != "note"]
        if mixed:
            return [
                _check(
                    "period_separation",
                    None,
                    "ERROR",
                    (
                        "Amounts were found with unrecognised periods "
                        f"{sorted({item.period for item in mixed})}. Current year and previous year "
                        "must be kept separate."
                    ),
                )
            ]
        current_keys = {item.key for item in ledger.items if item.period == "current"}
        previous_keys = {item.key for item in ledger.items if item.period == "previous"}
        return [
            _check(
                "period_separation",
                None,
                "PASS",
                (
                    "Current year and previous year amounts were stored and checked separately "
                    f"({len(current_keys)} current caption(s), {len(previous_keys)} previous caption(s))."
                ),
                details={"current_captions": len(current_keys), "previous_captions": len(previous_keys)},
            )
        ]


def build_ledger(classified: dict[str, Any], placements: list[dict[str, Any]]) -> Ledger:
    ledger = Ledger()
    for block_key in BLOCK_KEYS:
        block = classified.get(block_key) or {}
        if not isinstance(block, dict):
            continue
        for line in block.get("line_items") or []:
            _add_line(ledger, line, origin=block_key)
        for field_name, sourced in block.items():
            if field_name in {"line_items", "notes", "identified", "start_page", "end_page", "pages", "confidence", "excerpt"}:
                continue
            _add_sourced(ledger, field_name, sourced, origin=block_key)
        for note in block.get("notes") or []:
            if not isinstance(note, dict):
                continue
            for line in note.get("line_items") or []:
                _add_line(ledger, line, origin=f"{block_key}.notes")
    for placement in placements:
        value = placement.get("extracted_value")
        period = placement.get("period")
        if period not in PERIODS or not isinstance(value, (int, float)):
            continue
        key = placement.get("field_key") or slugify(str(placement.get("source_label") or ""))
        if not key:
            continue
        if ledger.get(str(key), period) is not None:
            continue
        ledger.add(
            Amount(
                key=str(key),
                label=str(placement.get("schedule_iii_label") or placement.get("source_label") or key),
                value=float(value),
                period=period,
                origin="placement",
                source_text=str(placement.get("excel_destination") or ""),
            )
        )
    return ledger


def nearly_equal(left: float, right: float) -> bool:
    return abs(left - right) <= max(ABS_TOLERANCE, REL_TOLERANCE * max(abs(left), abs(right), 1.0))


def _split_inputs(mapped_data: Any, classified_data: Any) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    payload = _as_dict(mapped_data)
    classified = _as_dict(classified_data)
    placements = payload.get("placements") if isinstance(payload.get("placements"), list) else []
    if not classified:
        if any(key in payload for key in ("balance_sheet", "profit_and_loss", "company")):
            classified = payload
        else:
            nested = payload.get("classified") or payload.get("classified_data")
            classified = _as_dict(nested)
    return classified, [item for item in placements if isinstance(item, dict)]


def _as_dict(value: Any) -> dict[str, Any]:
    if value is None:
        return {}
    if hasattr(value, "model_dump"):
        dumped = value.model_dump(mode="json")
        return dumped if isinstance(dumped, dict) else {}
    return value if isinstance(value, dict) else {}


def _add_line(ledger: Ledger, line: Any, *, origin: str) -> None:
    if not isinstance(line, dict):
        return
    label = str(line.get("label") or "").strip()
    if not label:
        return
    key = slugify(label)
    current = _numeric(line.get("current_period"))
    previous = _numeric(line.get("previous_period"))
    if current is not None:
        ledger.add(Amount(key=key, label=label, value=current, period="current", origin=origin))
    if previous is not None:
        ledger.add(Amount(key=key, label=label, value=previous, period="previous", origin=origin))


def _add_sourced(ledger: Ledger, field_name: str, sourced: Any, *, origin: str) -> None:
    value = _numeric(sourced)
    if value is None:
        return
    ledger.add(
        Amount(
            key=slugify(field_name),
            label=field_name.replace("_", " "),
            value=value,
            period="current",
            origin=origin,
        )
    )


def _numeric(value: Any) -> float | None:
    if value is None:
        return None
    if hasattr(value, "model_dump"):
        value = value.model_dump(mode="json")
    if isinstance(value, dict):
        raw = value.get("value")
    else:
        raw = value
    if isinstance(raw, bool) or raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    return None


def _group_total(
    ledger: Ledger,
    period: str,
    *,
    explicit: Iterable[str],
    children: Iterable[str],
    origins: Iterable[str] | None,
) -> dict[str, Any] | None:
    found = ledger.lookup(explicit, period, origins=origins) or ledger.lookup(explicit, period)
    if found is not None:
        return {"value": found.value, "parts": [found.label], "source": "explicit"}
    total, used = ledger.sum_keys(children, period, origins=origins)
    if total is None:
        total, used = ledger.sum_keys(children, period)
    if total is None:
        return None
    return {"value": total, "parts": used, "source": "sum"}


def _sum_optional(*groups: dict[str, Any] | None) -> tuple[float | None, dict[str, float]]:
    if any(group is None for group in groups):
        parts: dict[str, float] = {}
        for group in groups:
            if group is None:
                continue
            for name in group.get("parts") or []:
                parts[str(name)] = group["value"]
        return None, parts
    total = 0.0
    parts = {}
    for group in groups:
        assert group is not None
        total += group["value"]
        label = " + ".join(group.get("parts") or []) or "group"
        parts[label] = group["value"]
    return total, parts


def _identity_check(
    *,
    check_id: str,
    period: str,
    left_name: str,
    left_value: float | None,
    right_name: str,
    right_value: float | None,
    formula: str,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload = details or {}
    if left_value is None or right_value is None:
        missing = []
        if left_value is None:
            missing.append(left_name)
        if right_value is None:
            missing.append(right_name)
        return _check(
            check_id,
            period,
            "WARNING",
            (
                f"{formula} could not be verified for the {period} period because "
                f"{' and '.join(missing)} was not extracted. No amounts were inferred."
            ),
            details={"formula": formula, "left": left_value, "right": right_value, **payload},
        )
    if nearly_equal(left_value, right_value):
        return _check(
            check_id,
            period,
            "PASS",
            (
                f"{formula} holds for the {period} period: {left_name} {_fmt(left_value)} "
                f"= {right_name} {_fmt(right_value)}."
            ),
            details={"formula": formula, "left": left_value, "right": right_value, **payload},
        )
    difference = left_value - right_value
    return _check(
        check_id,
        period,
        "ERROR",
        (
            f"{formula} failed for the {period} period: {left_name} is {_fmt(left_value)} but "
            f"{right_name} is {_fmt(right_value)}. Difference = {_fmt(difference)}. "
            "The validator did not change any values."
        ),
        details={"formula": formula, "left": left_value, "right": right_value, "difference": difference, **payload},
    )


def _check(
    check_id: str,
    period: str | None,
    status: str,
    message: str,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    return {
        "id": check_id,
        "period": period,
        "status": status,
        "message": message,
        "details": details or {},
    }


def _block_identified(classified: dict[str, Any], key: str) -> bool:
    block = classified.get(key) or {}
    if not isinstance(block, dict):
        return False
    return bool(block.get("identified") or block.get("line_items"))


def _same_sign(left: float, right: float) -> bool | None:
    if nearly_equal(left, 0) or nearly_equal(right, 0):
        return True
    return (left > 0) == (right > 0)


def _fmt(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.4f}".rstrip("0").rstrip(".")
    return str(value)
