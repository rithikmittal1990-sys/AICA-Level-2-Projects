"""Human review of mapped values before Excel is generated.

Uncertain amounts are never auto-filled or auto-approved. Only rows the
reviewer marks approved are written into the workbook.
"""

from __future__ import annotations

import logging
import re
from copy import deepcopy
from typing import Any

from app.config import settings
from app.mapping.field_mapping import _format_destination

logger = logging.getLogger(__name__)

ALLOWED_PERIODS = ("current", "previous", "note")
ALLOWED_STATUSES = ("pending", "needs_review", "approved", "rejected")
YEAR_RANGE_RE = re.compile(r"^(\d{4})\s*[-–/]\s*(\d{2}|\d{4})$")
ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")
DMY_RE = re.compile(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$")


class ReviewIncompleteError(Exception):
    """Raised when below-threshold rows still need a decision."""

    def __init__(self, message: str, *, needs_review: int) -> None:
        super().__init__(message)
        self.message = message
        self.needs_review = needs_review


def confidence_threshold() -> float:
    value = getattr(settings, "review_confidence_threshold", 0.85)
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.85


def build_review(mapped: dict[str, Any], classified: dict[str, Any] | None = None) -> dict[str, Any]:
    """Create the reviewer table from mapping placements."""
    threshold = confidence_threshold()
    placements = [item for item in (mapped.get("placements") or []) if isinstance(item, dict)]
    destinations = _destinations_by_field(placements)
    notes = _note_numbers_by_field(placements)
    years = extract_financial_years(classified or {})
    items: list[dict[str, Any]] = []
    for index, placement in enumerate(placements):
        field_key = str(placement.get("field_key") or "")
        period = _coerce_period(placement.get("period"))
        confidence = _as_float(placement.get("confidence"))
        status = initial_status(placement, threshold)
        year = years.get("current") if period == "current" else years.get("previous") if period == "previous" else None
        note_number = placement.get("extracted_value") if period == "note" else notes.get(field_key)
        items.append(
            {
                "item_id": f"p{index}",
                "placement_index": index,
                "source_field": placement.get("source_label") or placement.get("schedule_iii_label"),
                "extracted_value": placement.get("extracted_value"),
                "original_value": placement.get("extracted_value"),
                "source_page": placement.get("source_page"),
                "schedule_iii_category": placement.get("schedule_iii_category"),
                "excel_destination": placement.get("excel_destination")
                or _format_destination(placement.get("excel_sheet"), placement.get("excel_cell")),
                "excel_sheet": placement.get("excel_sheet"),
                "excel_cell": placement.get("excel_cell"),
                "confidence": confidence,
                "status": status,
                "requires_review": status == "needs_review",
                "period": period,
                "original_period": period,
                "financial_year": year,
                "note_number": note_number,
                "field_key": field_key,
                "action": placement.get("action"),
                "warnings": list(placement.get("warnings") or []),
                "destinations": destinations.get(field_key) or {},
            }
        )
    payload = {
        "threshold": threshold,
        "financial_year": years,
        "items": items,
        "unmapped_sources": list(mapped.get("unmapped_sources") or []),
        "summary": summarize_review(items, threshold),
    }
    return payload


def initial_status(placement: dict[str, Any], threshold: float) -> str:
    confidence = _as_float(placement.get("confidence"))
    action = str(placement.get("action") or "")
    resolution = str(placement.get("resolution") or "")
    value = placement.get("extracted_value")
    if action in {"missing_value", "unmapped_destination"}:
        return "needs_review"
    if value is None:
        return "needs_review"
    if confidence is None or confidence < threshold:
        return "needs_review"
    if "fuzzy" in resolution:
        return "needs_review"
    return "pending"


def summarize_review(items: list[dict[str, Any]], threshold: float | None = None) -> dict[str, Any]:
    counts = {"pending": 0, "needs_review": 0, "approved": 0, "rejected": 0}
    for item in items:
        status = item.get("status")
        if status in counts:
            counts[status] += 1
    return {
        "total": len(items),
        "threshold": threshold if threshold is not None else confidence_threshold(),
        **counts,
        "ready_to_generate": counts["needs_review"] == 0,
    }


def apply_review_updates(review: dict[str, Any], payload: dict[str, Any] | None) -> dict[str, Any]:
    """Merge reviewer edits. Never fills a blank amount with zero."""
    updated = deepcopy(review)
    body = payload or {}
    if isinstance(body.get("financial_year"), dict):
        years = updated.setdefault("financial_year", {})
        for key in ("current", "previous"):
            if key in body["financial_year"]:
                years[key] = _optional_text(body["financial_year"].get(key))
        _refresh_item_years(updated)

    by_id = {item.get("item_id"): item for item in updated.get("items") or [] if isinstance(item, dict)}
    for raw in body.get("items") or []:
        if not isinstance(raw, dict):
            continue
        item = by_id.get(str(raw.get("item_id") or ""))
        if item is None:
            continue
        if "extracted_value" in raw:
            item["extracted_value"] = parse_extracted_value(raw.get("extracted_value"), period=item.get("period"))
        if "period" in raw and raw["period"] is not None:
            _apply_period_change(item, str(raw["period"]), updated.get("financial_year") or {})
        if "financial_year" in raw:
            item["financial_year"] = _optional_text(raw.get("financial_year"))
        if "note_number" in raw:
            item["note_number"] = parse_extracted_value(raw.get("note_number"), period="note")
            _sync_note_number(updated["items"], item)
        if "status" in raw and raw["status"] is not None:
            item["status"] = _coerce_status(raw.get("status"), item, confidence_threshold())
        item["requires_review"] = item.get("status") == "needs_review"

    updated["summary"] = summarize_review(updated.get("items") or [], updated.get("threshold"))
    return updated


def approved_mapped(mapped: dict[str, Any], review: dict[str, Any]) -> dict[str, Any]:
    """Return mapping data containing only reviewer-approved writes."""
    incomplete = [item for item in review.get("items") or [] if item.get("status") == "needs_review"]
    if incomplete:
        raise ReviewIncompleteError(
            "Mappings below the confidence threshold must be approved or rejected before generating Excel.",
            needs_review=len(incomplete),
        )
    result = deepcopy(mapped)
    placements = [item for item in (result.get("placements") or []) if isinstance(item, dict)]
    approved: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    for item in review.get("items") or []:
        if not isinstance(item, dict):
            continue
        index = item.get("placement_index")
        if not isinstance(index, int) or index < 0 or index >= len(placements):
            continue
        placement = deepcopy(placements[index])
        placement["extracted_value"] = item.get("extracted_value")
        placement["period"] = _coerce_period(item.get("period"))
        placement["excel_sheet"] = item.get("excel_sheet")
        placement["excel_cell"] = item.get("excel_cell")
        placement["excel_destination"] = item.get("excel_destination") or _format_destination(
            item.get("excel_sheet"), item.get("excel_cell")
        )
        placement["review_status"] = item.get("status")
        placement["financial_year"] = item.get("financial_year")
        if item.get("period") == "note":
            placement["extracted_value"] = item.get("extracted_value")
        elif item.get("note_number") is not None and item.get("period") != "note":
            placement["note_number"] = item.get("note_number")

        if item.get("status") != "approved":
            skipped.append(placement)
            continue
        if placement.get("extracted_value") is None:
            placement["action"] = "missing_value"
            skipped.append(placement)
            continue
        if placement.get("action") == "skip_formula" and not placement.get("overwrite_formula"):
            skipped.append(placement)
            continue
        if not placement.get("excel_sheet") or not placement.get("excel_cell"):
            placement["action"] = "unmapped_destination"
            skipped.append(placement)
            continue
        placement["action"] = "write"
        approved.append(placement)

    note_writes = _approved_note_writes(review, placements)
    existing = {(row.get("field_key"), row.get("period")) for row in approved}
    for note in note_writes:
        key = (note.get("field_key"), "note")
        if key not in existing:
            approved.append(note)
            existing.add(key)

    result["placements"] = approved
    result["review_skipped"] = skipped
    result["financial_year"] = review.get("financial_year")
    return result


def extract_financial_years(classified: dict[str, Any]) -> dict[str, str | None]:
    company = classified.get("company") if isinstance(classified, dict) else {}
    company = company if isinstance(company, dict) else {}
    current = _optional_text(_sourced_value(company.get("reporting_period"))) or _optional_text(
        _sourced_value(company.get("period_end"))
    )
    previous = previous_year_label(current) if current else None
    return {"current": current, "previous": previous}


def previous_year_label(current: str) -> str | None:
    text = (current or "").strip()
    match = YEAR_RANGE_RE.match(text)
    if match:
        start = int(match.group(1))
        end = match.group(2)
        end_num = int(end) if len(end) == 4 else int(end) + ((start // 100) * 100)
        if len(end) == 2 and end_num < start:
            end_num += 100
        prev_start = start - 1
        if len(end) == 2:
            return f"{prev_start}-{str(end_num - 1)[-2:]}"
        return f"{prev_start}-{end_num - 1}"
    iso = ISO_DATE_RE.match(text)
    if iso:
        year = int(iso.group(1)) - 1
        return f"{year}-{iso.group(2)}-{iso.group(3)}"
    dmy = DMY_RE.match(text)
    if dmy:
        year = int(dmy.group(3)) - 1
        return f"{dmy.group(1)}-{dmy.group(2)}-{year}"
    year_only = re.search(r"(20\d{2}|19\d{2})", text)
    if year_only:
        return str(int(year_only.group(1)) - 1)
    return None


def parse_extracted_value(value: Any, *, period: str | None = None) -> float | str | int | None:
    """Parse a reviewer edit. Blank stays blank; it is never coerced to zero."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    if not text or text.lower() in {"none", "null", "n/a", "-"}:
        return None
    if period == "note":
        return text
    compact = text.replace(",", "").replace(" ", "")
    if re.fullmatch(r"-?\d+", compact):
        return int(compact)
    if re.fullmatch(r"-?\d+\.\d+", compact):
        return float(compact)
    return text


def _apply_period_change(item: dict[str, Any], period: str, years: dict[str, Any]) -> None:
    new_period = _coerce_period(period)
    if new_period == item.get("period"):
        return
    item["period"] = new_period
    destination = (item.get("destinations") or {}).get(new_period) or {}
    if destination.get("excel_sheet") and destination.get("excel_cell"):
        item["excel_sheet"] = destination.get("excel_sheet")
        item["excel_cell"] = destination.get("excel_cell")
        item["excel_destination"] = destination.get("excel_destination") or _format_destination(
            destination.get("excel_sheet"), destination.get("excel_cell")
        )
    if new_period == "current":
        item["financial_year"] = years.get("current")
    elif new_period == "previous":
        item["financial_year"] = years.get("previous")
    else:
        item["financial_year"] = None


def _sync_note_number(items: list[dict[str, Any]], source: dict[str, Any]) -> None:
    field_key = source.get("field_key")
    note_number = source.get("note_number")
    for item in items:
        if item.get("field_key") != field_key:
            continue
        item["note_number"] = note_number
        if item.get("period") == "note":
            item["extracted_value"] = note_number


def _refresh_item_years(review: dict[str, Any]) -> None:
    years = review.get("financial_year") or {}
    for item in review.get("items") or []:
        if item.get("period") == "current":
            item["financial_year"] = years.get("current")
        elif item.get("period") == "previous":
            item["financial_year"] = years.get("previous")


def _approved_note_writes(review: dict[str, Any], placements: list[dict[str, Any]]) -> list[dict[str, Any]]:
    writes: list[dict[str, Any]] = []
    grouped: dict[str, list[dict[str, Any]]] = {}
    for item in review.get("items") or []:
        grouped.setdefault(str(item.get("field_key") or ""), []).append(item)
    for field_key, rows in grouped.items():
        if not any(row.get("status") == "approved" for row in rows):
            continue
        note_row = next((row for row in rows if row.get("period") == "note"), None)
        approved_row = next((row for row in rows if row.get("status") == "approved"), None)
        note_number = (note_row or approved_row or {}).get("note_number")
        if note_number is None:
            continue
        if note_row and note_row.get("status") == "approved":
            continue
        template = next(
            (item for item in placements if item.get("field_key") == field_key and item.get("period") == "note"),
            None,
        )
        destination = ((approved_row or {}).get("destinations") or {}).get("note") or {}
        sheet = destination.get("excel_sheet") or (template or {}).get("excel_sheet")
        cell = destination.get("excel_cell") or (template or {}).get("excel_cell")
        if not sheet or not cell:
            continue
        writes.append(
            {
                **(template or {}),
                "field_key": field_key,
                "period": "note",
                "extracted_value": note_number,
                "excel_sheet": sheet,
                "excel_cell": cell,
                "excel_destination": _format_destination(sheet, cell),
                "action": "write",
                "review_status": "approved",
            }
        )
    return writes


def _destinations_by_field(placements: list[dict[str, Any]]) -> dict[str, dict[str, dict[str, Any]]]:
    result: dict[str, dict[str, dict[str, Any]]] = {}
    for item in placements:
        field_key = str(item.get("field_key") or "")
        period = _coerce_period(item.get("period"))
        result.setdefault(field_key, {})[period] = {
            "excel_sheet": item.get("excel_sheet"),
            "excel_cell": item.get("excel_cell"),
            "excel_destination": item.get("excel_destination")
            or _format_destination(item.get("excel_sheet"), item.get("excel_cell")),
        }
    return result


def _note_numbers_by_field(placements: list[dict[str, Any]]) -> dict[str, Any]:
    notes: dict[str, Any] = {}
    for item in placements:
        if item.get("period") == "note" and item.get("extracted_value") is not None:
            notes[str(item.get("field_key") or "")] = item.get("extracted_value")
    return notes


def _coerce_period(value: Any) -> str:
    text = str(value or "current").strip().lower()
    if text in {"prev", "py", "previous_year", "prior"}:
        return "previous"
    if text in {"cy", "current_year"}:
        return "current"
    if text in ALLOWED_PERIODS:
        return text
    return "current"


def _coerce_status(value: Any, item: dict[str, Any], threshold: float) -> str:
    status = str(value or "").strip().lower().replace(" ", "_")
    if status in {"approve", "ok"}:
        status = "approved"
    if status in {"reject", "ignored"}:
        status = "rejected"
    if status not in ALLOWED_STATUSES:
        return str(item.get("status") or "pending")
    if status == "pending" and item.get("requires_review"):
        return "needs_review"
    if status == "pending":
        confidence = _as_float(item.get("confidence"))
        if confidence is None or confidence < threshold or item.get("extracted_value") is None:
            return "needs_review"
    return status


def _as_float(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _optional_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _sourced_value(payload: Any) -> Any:
    if isinstance(payload, dict):
        return payload.get("value")
    return payload
