"""Tests for ICAI Division I Schedule III reference analysis."""

from __future__ import annotations

import json

from app.mapping.field_mapping import FIELD_MAPPINGS, lookup_mapping
from app.mapping.schedule_iii_mapper import ScheduleIIIMapper


def _mapper() -> ScheduleIIIMapper:
    return ScheduleIIIMapper()


def test_mapping_module_importable() -> None:
    from app.mapping.field_mapping import FIELD_MAPPINGS
    from app.mapping.schedule_iii_mapper import ScheduleIIIMapper

    assert isinstance(FIELD_MAPPINGS, dict)
    assert ScheduleIIIMapper is not None


def test_reference_pdf_is_the_icai_guidance_note() -> None:
    mapper = _mapper()
    reference = mapper.load_reference()
    assert reference.page_count > 50
    assert "schedule iii" in reference.source_title.lower() or "division i" in reference.source_title.lower()


def test_required_sections_are_identified() -> None:
    section_ids = {section.section_id for section in _mapper().load_reference().sections}
    required = {
        "general_instructions",
        "balance_sheet",
        "profit_and_loss",
        "notes_to_accounts",
        "current_non_current",
        "shareholders_funds",
        "non_current_liabilities",
        "current_liabilities",
        "non_current_assets",
        "current_assets",
        "revenue_from_operations",
        "other_income",
        "expenses",
        "tax_expense",
        "earnings_per_share",
        "other_disclosures",
        "ratios",
    }
    missing = required - section_ids
    assert not missing, f"Missing sections: {sorted(missing)}"


def test_balance_sheet_hierarchy_follows_annexure_a() -> None:
    labels = _flatten_labels(_mapper().load_reference().balance_sheet)
    for required in (
        "Balance Sheet",
        "EQUITY AND LIABILITIES",
        "Shareholders' funds",
        "Share capital",
        "Reserves and surplus",
        "Money received against share warrants",
        "Share application money pending allotment",
        "Non-current liabilities",
        "Current liabilities",
        "ASSETS",
        "Non Current Assets",
        "Current assets",
    ):
        assert any(_norm(required) == _norm(label) or _norm(required) in _norm(label) for label in labels), required


def test_profit_and_loss_hierarchy_follows_annexure_a() -> None:
    labels = _flatten_labels(_mapper().load_reference().profit_and_loss)
    for required in (
        "Statement of Profit and Loss",
        "Revenue from operations",
        "Other income",
        "Total Income",
        "Expenses",
        "Finance costs",
        "Tax expense",
        "Earnings per equity share",
    ):
        assert any(_norm(required) in _norm(label) for label in labels), required


def test_every_hierarchy_node_has_a_pdf_source() -> None:
    reference = _mapper().load_reference()
    for node in (
        [reference.balance_sheet, reference.profit_and_loss, reference.notes_to_accounts]
        + reference.ratios
    ):
        _assert_sourced(node)


def test_reference_is_json_serializable() -> None:
    payload = json.loads(_mapper().to_json())
    assert "hierarchy" in payload
    assert "field_mappings" in payload
    assert payload["hierarchy"]["balance_sheet"]["children"]
    assert payload["hierarchy"]["statement_of_profit_and_loss"]["children"]


def test_field_mappings_are_built_from_hierarchy() -> None:
    mappings = _mapper().get_field_mappings()
    assert mappings
    assert FIELD_MAPPINGS
    share_capital = lookup_mapping("Share capital", mappings)
    assert share_capital is not None
    assert share_capital["statement"] == "balance_sheet"


def _flatten_labels(node) -> list[str]:
    labels = [node.label]
    for child in node.children:
        labels.extend(_flatten_labels(child))
    return labels


def _assert_sourced(node) -> None:
    assert node.source_pages, f"Unsourced node: {node.label}"
    assert node.label.strip(), "Empty label is not allowed"
    for child in node.children:
        _assert_sourced(child)


def _norm(value: str) -> str:
    return " ".join(value.lower().replace("’", "'").split())
