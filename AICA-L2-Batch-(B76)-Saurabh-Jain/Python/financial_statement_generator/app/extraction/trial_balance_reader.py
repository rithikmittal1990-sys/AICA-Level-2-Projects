"""Read trial balance workbooks (Excel) into structured account rows."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from app.extraction.exceptions import InvalidTrialBalanceError
from app.trial_balance.parser import (
    TrialBalanceParser,
    normalize_account_label,
    is_trial_balance_upload,
    closing_amount,
)

EXCEL_SUFFIXES = {".xlsx", ".xlsm"}


@dataclass(slots=True)
class TrialBalanceRow:
    row: int
    label: str
    debit: float | None
    credit: float | None
    parent_group: str | None = None
    is_group: bool = False


@dataclass(slots=True)
class TrialBalanceDocument:
    company_name: str | None
    period_label: str | None
    sheet_name: str
    stored_path: str
    rows: list[TrialBalanceRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "document": {
                "kind": "trial_balance",
                "company_name": self.company_name,
                "period_label": self.period_label,
                "sheet_name": self.sheet_name,
                "stored_path": self.stored_path,
                "accounts": len(self.rows),
            },
            "rows": [
                {
                    "row": item.row,
                    "label": item.label,
                    "debit": item.debit,
                    "credit": item.credit,
                    "parent_group": item.parent_group,
                    "is_group": item.is_group,
                }
                for item in self.rows
            ],
            "warnings": list(self.warnings),
        }


class TrialBalanceReader:
    """Validate and parse uploaded trial balance Excel files."""

    def __init__(self, upload_dir: Path | None = None, max_upload_bytes: int | None = None) -> None:
        self._parser = TrialBalanceParser(upload_dir=upload_dir, max_upload_bytes=max_upload_bytes)

    def save_upload(self, filename: str, content: bytes) -> Path:
        return self._parser.save_upload(filename, content)

    def read_upload(self, filename: str, content: bytes) -> TrialBalanceDocument:
        parsed = self._parser.read_upload(filename, content)
        return self._to_document(parsed)

    def read_path(self, path: Path) -> TrialBalanceDocument:
        parsed = self._parser.read_path(path)
        return self._to_document(parsed)

    def _to_document(self, parsed) -> TrialBalanceDocument:
        rows = [
            TrialBalanceRow(
                row=item.row,
                label=item.account_name,
                debit=item.debit or None,
                credit=item.credit or None,
                parent_group=item.parent_group,
                is_group=item.is_group,
            )
            for item in parsed.accounts
            if not item.is_group
        ]
        if not rows:
            raise InvalidTrialBalanceError("No leaf account rows were found in the trial balance.")
        return TrialBalanceDocument(
            company_name=parsed.company_name,
            period_label=parsed.period_label,
            sheet_name=parsed.sheet_name,
            stored_path=parsed.stored_path,
            rows=rows,
            warnings=parsed.warnings,
        )


__all__ = [
    "TrialBalanceReader",
    "TrialBalanceDocument",
    "TrialBalanceRow",
    "normalize_account_label",
    "closing_amount",
    "is_trial_balance_upload",
]
