"""Classify leaf trial balance accounts into Schedule III statement heads."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from app.trial_balance.models import MappedAccount, NormalizedAccount, StatementTotals, TrialBalanceClassification
from app.trial_balance.parser import TrialBalanceParseResult, financial_year_from_period, normalize_label

DEFAULT_ACCOUNT_MAP = Path(__file__).resolve().parent.parent / "mapping" / "trial_balance_account_map.json"

LOAN_LIABILITY_GROUPS = {"loans (liability)"}


@dataclass(slots=True)
class AccountRule:
    patterns: tuple[str, ...]
    schedule_iii_label: str
    statement_head: str
    statement: Literal["balance_sheet", "profit_and_loss"]
    nature: Literal["asset", "liability", "equity", "income", "expense"]
    aggregate_key: str | None = None
    allow_debit_on_liability: bool = False
    allow_credit_on_asset: bool = False


class TrialBalanceAccountClassifier:
    def __init__(self, config_path: Path | None = None) -> None:
        path = Path(config_path) if config_path else DEFAULT_ACCOUNT_MAP
        payload = json.loads(path.read_text(encoding="utf-8"))
        self.rules = [
            AccountRule(
                patterns=tuple(rule.get("patterns") or []),
                schedule_iii_label=str(rule.get("schedule_iii_label") or ""),
                statement_head=str(rule.get("statement_head") or rule.get("schedule_iii_label") or ""),
                statement=rule.get("statement") or "balance_sheet",
                nature=rule.get("nature") or "asset",
                aggregate_key=rule.get("aggregate_key"),
                allow_debit_on_liability=bool(rule.get("allow_debit_on_liability", False)),
                allow_credit_on_asset=bool(rule.get("allow_credit_on_asset", False)),
            )
            for rule in payload.get("accounts") or []
        ]
        self._asset_rules = {rule.aggregate_key: rule for rule in self.rules if rule.nature == "asset" and rule.aggregate_key}

    def classify(self, parsed: TrialBalanceParseResult) -> TrialBalanceClassification:
        mapped: list[MappedAccount] = []
        line_items: dict[str, float] = {}
        totals = StatementTotals()

        for account in parsed.leaf_accounts:
            mapped.append(self._map_account(account, line_items))

        self._compute_totals(mapped, line_items, totals, parsed)
        return TrialBalanceClassification(
            company_name=parsed.company_name,
            period_label=parsed.period_label,
            financial_year=financial_year_from_period(parsed.period_label),
            accounts=parsed.accounts,
            mapped=mapped,
            totals=totals,
            line_items=line_items,
        )

    def _map_account(self, account: NormalizedAccount, line_items: dict[str, float]) -> MappedAccount:
        rule = self._match_rule(account.account_name)
        if rule is None:
            return MappedAccount(
                account=account,
                schedule_iii_label="Unmapped",
                statement_head="Unmapped",
                statement="balance_sheet",
                amount=0.0,
                status="unmapped",
                reason="No mapping rule matched this ledger account.",
            )

        reclassification_reason: str | None = None
        if (
            rule.nature in {"liability", "equity"}
            and account.net_balance > 0
            and not rule.allow_debit_on_liability
            and self._is_loan_receivable(account)
        ):
            asset_rule = self._asset_rules.get("short_term_loans_and_advances") or self._asset_rules.get("loan_receivable")
            if asset_rule:
                rule = asset_rule
                reclassification_reason = (
                    "Debit balance on a loan/liability-group account classified as loan receivable (asset) "
                    "based on trial balance sign."
                )

        amount = self._statement_amount(account, rule)
        status: Literal["mapped", "review_required", "unmapped"] = "mapped"
        reason: str | None = None

        if (
            rule.nature in {"liability", "equity"}
            and account.net_balance > 0
            and not rule.allow_debit_on_liability
            and not reclassification_reason
        ):
            status = "review_required"
            reason = "Liability/equity account has a debit balance."
            amount = abs(account.net_balance)
        elif rule.nature in {"asset", "expense"} and account.net_balance < 0 and not rule.allow_credit_on_asset:
            if rule.nature == "expense" and "prior period" in normalize_label(account.account_name):
                amount = account.net_balance
            else:
                status = "review_required"
                reason = "Asset/expense account has a credit balance."
                amount = abs(account.net_balance)
        elif rule.nature == "expense" and account.net_balance < 0:
            amount = account.net_balance
        elif amount is None or abs(amount) < 0.005:
            return MappedAccount(
                account=account,
                schedule_iii_label=rule.schedule_iii_label,
                statement_head=rule.statement_head,
                statement=rule.statement,
                amount=0.0,
                status="mapped",
                reason="Zero balance ledger skipped.",
            )

        if status == "mapped":
            key = rule.aggregate_key or normalize_label(rule.schedule_iii_label)
            line_items[key] = round(line_items.get(key, 0.0) + float(amount), 2)

        return MappedAccount(
            account=account,
            schedule_iii_label=rule.schedule_iii_label,
            statement_head=rule.statement_head,
            statement=rule.statement,
            amount=round(float(amount), 2),
            status=status,
            reason=reason,
            reclassification_reason=reclassification_reason,
            note=rule.aggregate_key,  # reuse note field to carry aggregate_key
        )

    def _is_loan_receivable(self, account: NormalizedAccount) -> bool:
        group = normalize_label(account.parent_group or "")
        name = normalize_label(account.account_name)
        if group in LOAN_LIABILITY_GROUPS:
            return True
        return "loan" in name and account.net_balance > 0

    def _statement_amount(self, account: NormalizedAccount, rule: AccountRule) -> float | None:
        if rule.nature in {"asset", "expense"}:
            return account.net_balance
        return account.credit - account.debit

    def _match_rule(self, label: str) -> AccountRule | None:
        needle = normalize_label(label)
        best: tuple[int, AccountRule] | None = None
        for rule in self.rules:
            for pattern in rule.patterns:
                p = normalize_label(pattern)
                if p == needle:
                    score = 100 + len(p)
                elif p in needle:
                    score = 50 + len(p)
                elif needle in p:
                    score = 10 + len(p)
                else:
                    continue
                if best is None or score > best[0]:
                    best = (score, rule)
        return best[1] if best else None

    def _compute_totals(
        self,
        mapped: list[MappedAccount],
        line_items: dict[str, float],
        totals: StatementTotals,
        parsed: TrialBalanceParseResult,
    ) -> None:
        asset_keys = {
            "property_plant_and_equipment",
            "intangible_assets",
            "trade_receivables",
            "cash_and_cash_equivalents",
            "short_term_loans_and_advances",
            "loan_receivable",
            "other_current_assets",
            "deferred_tax_asset",
        }
        liability_keys = {
            "trade_payables",
            "other_current_liabilities",
            "short_term_provisions",
            "long_term_borrowings",
            "short_term_borrowings",
            "duties_and_taxes",
        }

        share_capital = line_items.get("share_capital", 0.0)
        reserves = line_items.get("reserves_and_surplus", 0.0)

        totals.total_income = round(
            sum(line_items.get(key, 0.0) for key in ("revenue_from_operations", "other_income", "deferred_tax_income")),
            2,
        )
        operating_expense_keys = (
            "employee_benefits_expense",
            "depreciation_and_amortization_expense",
            "other_expenses",
        )
        operating_expenses = sum(line_items.get(key, 0.0) for key in operating_expense_keys)
        prior = line_items.get("prior_period_items", 0.0)
        totals.total_expenses = round(operating_expenses + prior, 2)
        totals.profit_before_tax = round(totals.total_income - totals.total_expenses, 2)
        totals.current_tax = round(line_items.get("current_tax", 0.0), 2)
        totals.deferred_tax = round(line_items.get("deferred_tax_income", 0.0), 2)
        totals.profit_after_tax = round(totals.profit_before_tax - totals.current_tax, 2)
        totals.current_year_profit = totals.profit_after_tax

        totals.total_equity = round(share_capital + reserves + totals.current_year_profit, 2)

        for key, amount in line_items.items():
            if key in asset_keys:
                totals.total_assets += amount
            elif key in liability_keys:
                totals.total_liabilities += amount

        totals.total_assets = round(totals.total_assets, 2)
        totals.total_liabilities = round(totals.total_liabilities, 2)
        totals.balance_sheet_difference = round(
            totals.total_assets - (totals.total_equity + totals.total_liabilities),
            2,
        )

        tb_debit = parsed.grand_total_debit
        tb_credit = parsed.grand_total_credit
        if tb_debit is None or tb_credit is None:
            tb_debit = sum(a.debit for a in parsed.leaf_accounts)
            tb_credit = sum(a.credit for a in parsed.leaf_accounts)
        totals.tb_total_debit = round(float(tb_debit or 0), 2)
        totals.tb_total_credit = round(float(tb_credit or 0), 2)
