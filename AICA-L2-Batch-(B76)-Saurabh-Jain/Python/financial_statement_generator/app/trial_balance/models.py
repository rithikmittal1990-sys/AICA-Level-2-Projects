"""Normalized trial balance account structures."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

BalanceType = Literal["debit", "credit", "zero"]
AccountStatus = Literal["mapped", "review_required", "unmapped"]


@dataclass(slots=True)
class NormalizedAccount:
    account_name: str
    parent_group: str | None
    debit: float
    credit: float
    net_balance: float
    balance_type: BalanceType
    row: int
    is_group: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "account_name": self.account_name,
            "parent_group": self.parent_group,
            "debit": self.debit,
            "credit": self.credit,
            "net_balance": self.net_balance,
            "balance_type": self.balance_type,
            "row": self.row,
            "is_group": self.is_group,
        }


@dataclass(slots=True)
class MappedAccount:
    account: NormalizedAccount
    schedule_iii_label: str
    statement_head: str
    statement: Literal["balance_sheet", "profit_and_loss"]
    amount: float
    status: AccountStatus
    reason: str | None = None
    note: str | None = None
    reclassification_reason: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            **self.account.to_dict(),
            "schedule_iii_label": self.schedule_iii_label,
            "statement_head": self.statement_head,
            "statement": self.statement,
            "amount": self.amount,
            "status": self.status,
            "reason": self.reason,
            "note": self.note,
            "reclassification_reason": self.reclassification_reason,
        }


@dataclass(slots=True)
class StatementTotals:
    total_assets: float = 0.0
    total_equity: float = 0.0
    total_liabilities: float = 0.0
    total_income: float = 0.0
    total_expenses: float = 0.0
    profit_before_tax: float = 0.0
    current_tax: float = 0.0
    deferred_tax: float = 0.0
    profit_after_tax: float = 0.0
    current_year_profit: float = 0.0
    balance_sheet_difference: float = 0.0
    tb_total_debit: float = 0.0
    tb_total_credit: float = 0.0

    def to_dict(self) -> dict[str, float]:
        return {
            "total_assets": self.total_assets,
            "total_equity": self.total_equity,
            "total_liabilities": self.total_liabilities,
            "total_income": self.total_income,
            "total_expenses": self.total_expenses,
            "profit_before_tax": self.profit_before_tax,
            "current_tax": self.current_tax,
            "deferred_tax": self.deferred_tax,
            "profit_after_tax": self.profit_after_tax,
            "current_year_profit": self.current_year_profit,
            "balance_sheet_difference": self.balance_sheet_difference,
            "tb_total_debit": self.tb_total_debit,
            "tb_total_credit": self.tb_total_credit,
        }


@dataclass(slots=True)
class TrialBalanceClassification:
    company_name: str | None
    period_label: str | None
    financial_year: str | None
    accounts: list[NormalizedAccount] = field(default_factory=list)
    mapped: list[MappedAccount] = field(default_factory=list)
    totals: StatementTotals = field(default_factory=StatementTotals)
    line_items: dict[str, float] = field(default_factory=dict)

    def mapped_accounts(self) -> list[MappedAccount]:
        return [item for item in self.mapped if item.status == "mapped"]

    def review_accounts(self) -> list[MappedAccount]:
        return [item for item in self.mapped if item.status == "review_required"]

    def unmapped_accounts(self) -> list[MappedAccount]:
        return [item for item in self.mapped if item.status == "unmapped"]
