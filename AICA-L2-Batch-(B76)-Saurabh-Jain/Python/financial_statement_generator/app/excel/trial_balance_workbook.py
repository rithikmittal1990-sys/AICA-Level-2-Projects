"""Post-process generated workbooks for trial-balance-only output."""

from __future__ import annotations

import re
from typing import Iterable

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.excel.template_manager import canonical_sheet_name

EXCLUDED_SHEETS = ("Cash Flow",)
EXTERNAL_WORKBOOK_RE = re.compile(r"\[\d+\]")
SAMPLE_SUBSIDIARY_RE = re.compile(r"\bABCDE?\s*(Private\s+)?Limited\b", re.IGNORECASE)
SAMPLE_AUDITOR_RE = re.compile(r"\bFor\s+ABC\s*&\s*Company\b", re.IGNORECASE)


def remove_excluded_sheets(workbook: Workbook, excluded: Iterable[str] = EXCLUDED_SHEETS) -> list[str]:
    removed: list[str] = []
    excluded_norm = {canonical_sheet_name(name) for name in excluded}
    for name in list(workbook.sheetnames):
        if canonical_sheet_name(name) not in excluded_norm:
            continue
        del workbook[name]
        removed.append(name)
    return removed


def sanitize_trial_balance_workbook(
    workbook: Workbook,
    *,
    company_name: str | None,
    period_label: str | None = None,
    template_names: Iterable[str] = (),
) -> list[dict[str, str]]:
    """Remove sample-company content and external workbook references."""
    written: list[dict[str, str]] = []
    templates = list(template_names)
    for sheet in workbook.worksheets:
        written.extend(
            _sanitize_sheet(
                sheet,
                company_name=company_name,
                period_label=period_label,
                templates=templates,
            )
        )
    return written


def _sanitize_sheet(
    sheet: Worksheet,
    *,
    company_name: str | None,
    period_label: str | None,
    templates: list[str],
) -> list[dict[str, str]]:
    written: list[dict[str, str]] = []
    for row in sheet.iter_rows():
        for cell in row:
            value = cell.value
            if isinstance(value, str) and value.startswith("=") and EXTERNAL_WORKBOOK_RE.search(value):
                cell.value = 0
                written.append({"sheet": sheet.title, "cell": cell.coordinate, "value": "0"})
                continue
            if not isinstance(value, str) or not value.strip():
                continue
            updated = _sanitize_text(
                value,
                sheet_title=sheet.title,
                company_name=company_name,
                period_label=period_label,
                templates=templates,
            )
            if updated != value:
                cell.value = updated
                written.append({"sheet": sheet.title, "cell": cell.coordinate, "value": updated})
    return written


def _sanitize_text(
    value: str,
    *,
    sheet_title: str,
    company_name: str | None,
    period_label: str | None,
    templates: list[str],
) -> str:
    updated = value
    if SAMPLE_AUDITOR_RE.search(updated):
        updated = SAMPLE_AUDITOR_RE.sub("For ...........................", updated)
    if SAMPLE_SUBSIDIARY_RE.search(updated) and company_name:
        if normalize_key(company_name) not in normalize_key(updated):
            updated = ""
    for template in sorted(templates, key=len, reverse=True):
        pattern = re.compile(re.escape(template), flags=re.IGNORECASE)
        if pattern.search(updated) and company_name:
            updated = pattern.sub(company_name, updated)
    if updated.strip().startswith("REGD. OFFICE-"):
        return "-"
    if updated.strip().startswith("Email-") and "CIN-" in updated:
        return "-"
    if updated.strip().startswith("Factory Office:-"):
        return "-"
    if "was incorporated in India on January 01, 2000" in updated and company_name:
        period = period_label or "the reporting period"
        return (
            f"{company_name} ('the Company') is the entity for which these financial statements "
            f"have been prepared from the uploaded trial balance for {period}."
        )
    if updated.strip() == "Notes forming part of the Financial Statements for" and period_label:
        return f"Notes forming part of the Financial Statements for {period_label}"
    if sheet_title.strip() in {"Note 1-2", "BS PnL"} and updated.strip().startswith("Statement of Profit and Loss"):
        if period_label:
            return f"Statement of Profit and Loss for {period_label}"
    if sheet_title.strip() in {"Note 1-2", "BS PnL"} and "Balance Sheet as at" in updated and period_label:
        return f"Balance Sheet for {period_label}"
    return updated


def normalize_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
