"""Balance sheet reconciliation and account-level debug reports."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

from app.trial_balance.models import TrialBalanceClassification

PL_PARENT_GROUPS = {"Sales Accounts", "Indirect Incomes", "Indirect Expenses"}


@dataclass(slots=True)
class AccountReconciliationRow:
    account_name: str
    tb_debit: float
    tb_credit: float
    net_balance: float
    mapped_statement: str
    statement_amount: float
    difference: float
    status: Literal["PASS", "WARNING", "REVIEW"]
    reason: str = ""


@dataclass(slots=True)
class BalanceSheetReconciliation:
    total_assets: float = 0.0
    total_equity: float = 0.0
    total_liabilities: float = 0.0
    current_year_profit: float = 0.0
    equity_plus_liabilities: float = 0.0
    difference: float = 0.0
    cause_summary: str = ""
    account_rows: list[AccountReconciliationRow] = field(default_factory=list)
    breakdown: dict[str, float] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "total_assets": self.total_assets,
            "total_equity": self.total_equity,
            "total_liabilities": self.total_liabilities,
            "current_year_profit": self.current_year_profit,
            "equity_plus_liabilities": self.equity_plus_liabilities,
            "difference": self.difference,
            "cause_summary": self.cause_summary,
            "breakdown": self.breakdown,
            "account_rows": [
                {
                    "account": row.account_name,
                    "tb_debit": row.tb_debit,
                    "tb_credit": row.tb_credit,
                    "net_balance": row.net_balance,
                    "mapped_statement": row.mapped_statement,
                    "statement_amount": row.statement_amount,
                    "difference": row.difference,
                    "status": row.status,
                    "reason": row.reason,
                }
                for row in self.account_rows
            ],
        }


def build_reconciliation(classification: TrialBalanceClassification) -> BalanceSheetReconciliation:
    totals = classification.totals
    recon = BalanceSheetReconciliation(
        total_assets=round(totals.total_assets, 2),
        total_equity=round(totals.total_equity, 2),
        total_liabilities=round(totals.total_liabilities, 2),
        current_year_profit=round(totals.current_year_profit, 2),
        equity_plus_liabilities=round(totals.total_equity + totals.total_liabilities, 2),
        difference=round(totals.balance_sheet_difference, 2),
        breakdown=dict(classification.line_items),
    )
    recon.cause_summary = _cause_summary(classification, recon.difference)

    for item in classification.mapped:
        status: Literal["PASS", "WARNING", "REVIEW"] = "PASS"
        if item.status == "review_required":
            status = "REVIEW"
        elif item.reclassification_reason:
            status = "WARNING"
        recon.account_rows.append(
            AccountReconciliationRow(
                account_name=item.account.account_name,
                tb_debit=item.account.debit,
                tb_credit=item.account.credit,
                net_balance=item.account.net_balance,
                mapped_statement=item.statement_head,
                statement_amount=item.amount,
                difference=round(item.account.net_balance - item.amount, 2),
                status=status,
                reason=item.reason or item.reclassification_reason or "",
            )
        )
    return recon


def _cause_summary(classification: TrialBalanceClassification, difference: float) -> str:
    if abs(difference) <= 0.01:
        return "Balance Sheet balances after including current-year profit in equity and correctly classifying debit-balance loan accounts as assets."

    review = classification.review_accounts()
    profit = classification.totals.current_year_profit
    review_debit_loans = [
        item for item in review if item.account.parent_group and "loan" in item.account.parent_group.lower()
    ]
    if review_debit_loans and abs(difference) > 0.01:
        loan = review_debit_loans[0]
        implied = round(loan.amount - profit, 2)
        if abs(implied - difference) < 0.05:
            return (
                f"Balance Sheet difference caused by: {loan.account.account_name} (₹{loan.amount:,.2f} debit) "
                f"excluded from assets while current-year profit (₹{profit:,.2f}) not yet added to equity — "
                f"net gap ₹{difference:,.2f} (= loan debit − current-year profit)."
            )
    return f"Unresolved Balance Sheet difference of ₹{difference:,.2f}. See account reconciliation table."


def tb_bs_pl_split(leaf_accounts: list) -> dict[str, float]:
    bs_debit = bs_credit = pl_debit = pl_credit = 0.0
    for account in leaf_accounts:
        if account.parent_group in PL_PARENT_GROUPS:
            pl_debit += account.debit
            pl_credit += account.credit
        else:
            bs_debit += account.debit
            bs_credit += account.credit
    return {
        "bs_debit": round(bs_debit, 2),
        "bs_credit": round(bs_credit, 2),
        "bs_net_debit": round(bs_debit - bs_credit, 2),
        "pl_debit": round(pl_debit, 2),
        "pl_credit": round(pl_credit, 2),
        "pl_net_profit": round(pl_credit - pl_debit, 2),
    }
