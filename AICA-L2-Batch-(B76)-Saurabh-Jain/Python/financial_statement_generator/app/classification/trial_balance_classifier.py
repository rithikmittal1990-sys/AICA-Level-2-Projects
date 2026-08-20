"""Classify trial balance account rows into Schedule III line items."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from app.extraction.trial_balance_reader import TrialBalanceDocument, closing_amount, normalize_account_label
from app.mapping.field_mapping import normalize_label
from app.models.financial_models import (
    BalanceSheet,
    ClassifiedFinancialData,
    ClassificationWarning,
    CompanyInfo,
    Expenses,
    IdentifiedDocumentSection,
    LineItem,
    ProfitAndLoss,
    Revenue,
    SourcedValue,
    unextracted_value,
)

DEFAULT_ACCOUNT_MAP = Path(__file__).resolve().parent.parent / "mapping" / "trial_balance_account_map.json"


@dataclass(slots=True)
class AccountRule:
    patterns: tuple[str, ...]
    schedule_iii_label: str
    statement: str
    nature: str
    aggregate_key: str | None = None


class TrialBalanceClassifier:
    """Map trial balance accounts to classified financial data."""

    def __init__(self, config_path: Path | None = None) -> None:
        path = Path(config_path) if config_path else DEFAULT_ACCOUNT_MAP
        payload = json.loads(path.read_text(encoding="utf-8"))
        self.group_labels = {normalize_label(item) for item in payload.get("group_labels") or []}
        self.rules = [
            AccountRule(
                patterns=tuple(rule.get("patterns") or []),
                schedule_iii_label=str(rule.get("schedule_iii_label") or ""),
                statement=str(rule.get("statement") or "balance_sheet"),
                nature=str(rule.get("nature") or "asset"),
                aggregate_key=rule.get("aggregate_key"),
            )
            for rule in payload.get("accounts") or []
        ]

    def classify(self, document: TrialBalanceDocument) -> ClassifiedFinancialData:
        balance_items: dict[str, dict[str, Any]] = {}
        profit_items: dict[str, dict[str, Any]] = {}
        unmapped: list[str] = []
        for row in document.rows:
            label_key = normalize_label(row.label)
            if label_key in self.group_labels:
                continue
            rule = self._match_rule(row.label)
            if rule is None:
                unmapped.append(row.label)
                continue
            amount = closing_amount(row.debit, row.credit, nature=rule.nature)
            if amount is None:
                continue
            bucket = balance_items if rule.statement == "balance_sheet" else profit_items
            key = rule.aggregate_key or normalize_label(rule.schedule_iii_label)
            entry = bucket.setdefault(
                key,
                {
                    "label": rule.schedule_iii_label,
                    "amount": 0.0,
                    "sources": [],
                    "statement": rule.statement,
                },
            )
            entry["amount"] += amount
            entry["sources"].append(row.label)

        balance_sheet = BalanceSheet(
            identified=True,
            confidence=0.95,
            line_items=[self._line_item(entry, document) for entry in balance_items.values()],
        )
        profit_and_loss = ProfitAndLoss(
            identified=True,
            confidence=0.95,
            line_items=[self._line_item(entry, document) for entry in profit_items.values()],
        )
        warnings = [
            ClassificationWarning(
                code="trial_balance_unmapped",
                message=f"Trial balance account {label!r} was not mapped to Schedule III.",
                section="trial_balance",
                field=label,
                source_page=1,
            )
            for label in unmapped
        ]
        for message in document.warnings:
            warnings.append(
                ClassificationWarning(code="trial_balance_warning", message=message, section="trial_balance")
            )
        company = CompanyInfo(
            company_name=self._sourced(document.company_name, source_text=document.company_name),
            reporting_period=self._sourced(document.period_label, source_text=document.period_label),
            period_end=self._sourced(_period_end(document.period_label), source_text=document.period_label),
        )
        return ClassifiedFinancialData(
            company=company,
            sections=[
                IdentifiedDocumentSection(
                    section_id="trial_balance",
                    title="Trial Balance",
                    start_page=1,
                    end_page=1,
                    pages=[1],
                    confidence=0.95,
                    excerpt=document.sheet_name,
                )
            ],
            balance_sheet=balance_sheet,
            profit_and_loss=profit_and_loss,
            revenue=Revenue(identified=bool(profit_items), line_items=profit_and_loss.line_items),
            expenses=Expenses(identified=bool(profit_items), line_items=profit_and_loss.line_items),
            warnings=warnings,
        )

    def _line_item(self, entry: dict[str, Any], document: TrialBalanceDocument) -> LineItem:
        source_text = "; ".join(entry["sources"])
        return LineItem(
            label=entry["label"],
            current_period=SourcedValue.from_extraction(
                round(float(entry["amount"]), 2),
                source_page=1,
                source_text=source_text,
                confidence=0.95,
            ),
            previous_period=unextracted_value(),
        )

    def _match_rule(self, label: str) -> AccountRule | None:
        needle = normalize_label(label)
        for rule in self.rules:
            for pattern in rule.patterns:
                if normalize_label(pattern) in needle or needle in normalize_label(pattern):
                    return rule
        return None

    def _sourced(self, value: str | None, *, source_text: str | None = None) -> SourcedValue:
        if not value:
            return unextracted_value()
        return SourcedValue.from_extraction(
            value,
            source_page=1,
            source_text=source_text or value,
            confidence=0.95,
        )


def _period_end(period_label: str | None) -> str | None:
    if not period_label:
        return None
    match = re.search(r"to\s+(.+)$", period_label, flags=re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return period_label.strip()
