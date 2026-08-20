"""Strict visual and structural comparison of the sample template vs generated workbook.

The sample workbook is the output template. Generation may change financial
values and intentionally dynamic fields. Layout, formatting, formulas, and
static labels must stay the same.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from app.config import settings
from app.excel.template_manager import TemplateManager, file_sha256

DEFAULT_GENERATED_FILENAME = "Financial_Statements_Generated.xlsx"

PLACEHOLDER_RE = re.compile(r"^(xxx+|-+|_+|n/?a|nil|none|null|\.|0)$", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
NUMERIC_TEXT_RE = re.compile(r"^-?[0-9]+(?:,[0-9]+)*(?:\.[0-9]+)?$")

LAYOUT_CATEGORIES = (
    "sheet_name",
    "sheet_order",
    "merged_cells",
    "row_heights",
    "column_widths",
    "font",
    "font_size",
    "bold",
    "italic",
    "borders",
    "alignment",
    "number_format",
    "formula",
    "page_orientation",
    "print_area",
    "header_footer",
    "freeze_panes",
    "hidden_rows",
    "hidden_columns",
    "sheet_state",
    "unexpected_value",
)


@dataclass(slots=True)
class Difference:
    """One comparison finding."""

    severity: str
    category: str
    message: str
    sheet: str | None = None
    cell: str | None = None
    template: Any = None
    generated: Any = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "severity": self.severity,
            "category": self.category,
            "message": self.message,
            "sheet": self.sheet,
            "cell": self.cell,
            "template": self.template,
            "generated": self.generated,
        }


@dataclass(slots=True)
class ComparisonReport:
    """Result of comparing the sample template with a generated workbook."""

    ok: bool
    template_path: str
    generated_path: str
    differences: list[Difference] = field(default_factory=list)
    template_hash: str | None = None
    generated_hash: str | None = None

    @property
    def errors(self) -> list[Difference]:
        return [item for item in self.differences if item.severity == "error"]

    @property
    def warnings(self) -> list[Difference]:
        return [item for item in self.differences if item.severity == "warning"]

    @property
    def ignored(self) -> list[Difference]:
        return [item for item in self.differences if item.severity == "ignored"]

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "template_path": self.template_path,
            "generated_path": self.generated_path,
            "template_hash": self.template_hash,
            "generated_hash": self.generated_hash,
            "error_count": len(self.errors),
            "warning_count": len(self.warnings),
            "ignored_value_changes": len(self.ignored),
            "differences": [item.to_dict() for item in self.differences],
            "errors": [item.message for item in self.errors],
            "warnings": [item.message for item in self.warnings],
        }


def compare_workbooks(
    template_path: str | Path | None = None,
    generated_path: str | Path | None = None,
    *,
    allowed_formula_overwrites: Iterable[tuple[str, str]] | None = None,
    allowed_value_cells: Iterable[tuple[str, str]] | None = None,
    allowed_missing_sheets: Iterable[str] | None = None,
) -> ComparisonReport:
    """Compare the sample template with a generated copy without modifying either file."""
    template = Path(template_path) if template_path else settings.template_path
    generated = Path(generated_path) if generated_path else settings.output_dir / DEFAULT_GENERATED_FILENAME
    differences: list[Difference] = []

    if not template.exists():
        differences.append(
            Difference("error", "sheet_name", f"Template was not found: {template}")
        )
        return ComparisonReport(False, str(template), str(generated), differences)
    if not generated.exists():
        differences.append(
            Difference("error", "sheet_name", f"Generated workbook was not found: {generated}")
        )
        return ComparisonReport(
            False,
            str(template),
            str(generated),
            differences,
            template_hash=file_sha256(template),
        )
    if generated.resolve() == template.resolve():
        differences.append(
            Difference(
                "error",
                "sheet_name",
                "Comparison target is the original template; generation must use a copy.",
            )
        )
        return ComparisonReport(False, str(template), str(generated), differences)

    original = TemplateManager(template).inspect().to_dict()
    produced = TemplateManager(generated).inspect().to_dict()
    report = compare_metadata(
        original,
        produced,
        allowed_formula_overwrites=allowed_formula_overwrites,
        allowed_value_cells=allowed_value_cells,
        allowed_missing_sheets=allowed_missing_sheets,
        template_path=str(template),
        generated_path=str(generated),
    )
    report.template_hash = original.get("source_sha256")
    report.generated_hash = produced.get("source_sha256")
    return report


def compare_metadata(
    original: dict[str, Any],
    produced: dict[str, Any],
    *,
    allowed_formula_overwrites: Iterable[tuple[str, str]] | None = None,
    allowed_value_cells: Iterable[tuple[str, str]] | None = None,
    allowed_missing_sheets: Iterable[str] | None = None,
    template_path: str = "",
    generated_path: str = "",
) -> ComparisonReport:
    """Compare two TemplateManager inspect payloads."""
    from app.excel.template_manager import canonical_sheet_name

    allowed_formulas = {
        (sheet, str(cell).upper()) for sheet, cell in (allowed_formula_overwrites or []) if sheet and cell
    }
    allowed_values = {
        (sheet, str(cell).upper()) for sheet, cell in (allowed_value_cells or []) if sheet and cell
    }
    missing_allowed = {canonical_sheet_name(name) for name in (allowed_missing_sheets or []) if name}
    found: list[Difference] = []

    original_order = list(original.get("sheet_order") or original.get("sheet_names") or [])
    produced_order = list(produced.get("sheet_order") or produced.get("sheet_names") or [])
    expected_order = [
        name for name in original_order if canonical_sheet_name(name) not in missing_allowed
    ]
    if produced_order != expected_order:
        found.append(
            Difference(
                "error",
                "sheet_order",
                f"Sheet order changed: {expected_order} -> {produced_order}",
                template=expected_order,
                generated=produced_order,
            )
        )
    if produced_order != list(produced.get("sheet_names") or produced_order):
        found.append(
            Difference(
                "error",
                "sheet_name",
                "Generated sheet names do not match sheet order.",
                template=original.get("sheet_names"),
                generated=produced.get("sheet_names"),
            )
        )

    original_sheets = {sheet["name"]: sheet for sheet in original.get("sheets") or []}
    produced_sheets = {sheet["name"]: sheet for sheet in produced.get("sheets") or []}
    for name in expected_order:
        if name not in produced_sheets:
            found.append(
                Difference("error", "sheet_name", f"Missing sheet {name!r}.", sheet=name, template=name)
            )
    for name in produced_order:
        if name not in original_sheets:
            found.append(
                Difference(
                    "error",
                    "sheet_name",
                    f"Unexpected sheet {name!r} was added.",
                    sheet=name,
                    generated=name,
                )
            )

    for name, source in original_sheets.items():
        target = produced_sheets.get(name)
        if target is None:
            continue
        found.extend(
            _compare_sheet(
                source,
                target,
                allowed_formulas=allowed_formulas,
                allowed_values=allowed_values,
            )
        )

    original_names = {item.get("name") for item in original.get("defined_names") or []}
    produced_names = {item.get("name") for item in produced.get("defined_names") or []}
    missing_names = original_names - produced_names
    if missing_names:
        found.append(
            Difference(
                "error",
                "formula",
                f"Defined names were dropped: {sorted(name for name in missing_names if name)}",
                template=sorted(original_names),
                generated=sorted(produced_names),
            )
        )

    return ComparisonReport(
        ok=not any(item.severity == "error" for item in found),
        template_path=template_path,
        generated_path=generated_path,
        differences=found,
    )


def _compare_sheet(
    source: dict[str, Any],
    target: dict[str, Any],
    *,
    allowed_formulas: set[tuple[str, str]],
    allowed_values: set[tuple[str, str]],
) -> list[Difference]:
    name = source["name"]
    found: list[Difference] = []

    if source.get("name") != target.get("name"):
        found.append(
            Difference(
                "error",
                "sheet_name",
                f"Sheet name changed: {source.get('name')!r} -> {target.get('name')!r}.",
                sheet=name,
                template=source.get("name"),
                generated=target.get("name"),
            )
        )

    source_merged = _normalized_set(source.get("merged_cells"))
    target_merged = _normalized_set(target.get("merged_cells"))
    if source_merged != target_merged:
        found.append(
            Difference(
                "error",
                "merged_cells",
                f"{name}: merged cells changed.",
                sheet=name,
                template=sorted(source_merged),
                generated=sorted(target_merged),
            )
        )

    if source.get("freeze_panes") != target.get("freeze_panes"):
        found.append(
            Difference(
                "error",
                "freeze_panes",
                f"{name}: freeze panes changed.",
                sheet=name,
                template=source.get("freeze_panes"),
                generated=target.get("freeze_panes"),
            )
        )

    source_print = _normalize_print_area(source.get("print_area"))
    target_print = _normalize_print_area(target.get("print_area"))
    if source_print != target_print:
        found.append(
            Difference(
                "error",
                "print_area",
                f"{name}: print area changed.",
                sheet=name,
                template=source_print,
                generated=target_print,
            )
        )

    if source.get("sheet_state") != target.get("sheet_state"):
        found.append(
            Difference(
                "error",
                "sheet_state",
                f"{name}: sheet state changed.",
                sheet=name,
                template=source.get("sheet_state"),
                generated=target.get("sheet_state"),
            )
        )

    found.extend(_compare_hidden(name, "hidden_rows", source.get("hidden_rows"), target.get("hidden_rows")))
    found.extend(
        _compare_hidden(name, "hidden_columns", source.get("hidden_columns"), target.get("hidden_columns"))
    )
    found.extend(_compare_dimensions(name, "row_heights", source.get("row_heights"), target.get("row_heights"), "row"))
    found.extend(
        _compare_dimensions(name, "column_widths", source.get("column_widths"), target.get("column_widths"), "column")
    )
    found.extend(_compare_page_setup(name, source.get("page_setup") or {}, target.get("page_setup") or {}))
    found.extend(_compare_formulas(name, source.get("formulas") or {}, target.get("formulas") or {}, allowed_formulas))
    found.extend(
        _compare_cells(
            name,
            source.get("cells") or {},
            target.get("cells") or {},
            allowed_formulas=allowed_formulas,
            allowed_values=allowed_values,
        )
    )
    return found


def _compare_hidden(sheet: str, category: str, source: Any, target: Any) -> list[Difference]:
    left = _normalized_set(source)
    right = _normalized_set(target)
    if left == right:
        return []
    label = "hidden rows" if category == "hidden_rows" else "hidden columns"
    return [
        Difference(
            "error",
            category,
            f"{sheet}: {label} changed.",
            sheet=sheet,
            template=sorted(left, key=str),
            generated=sorted(right, key=str),
        )
    ]


def _compare_dimensions(
    sheet: str,
    category: str,
    source: dict[str, Any] | None,
    target: dict[str, Any] | None,
    kind: str,
) -> list[Difference]:
    source_map = source or {}
    target_map = target or {}
    found: list[Difference] = []
    for key in sorted(set(source_map) | set(target_map), key=str):
        left = source_map.get(key)
        right = target_map.get(key)
        if _same_dimension(left, right):
            continue
        field_name = "row height" if kind == "row" else "column width"
        found.append(
            Difference(
                "error",
                category,
                f"{sheet}: {field_name} for {kind} {key} changed.",
                sheet=sheet,
                cell=str(key),
                template=left,
                generated=right,
            )
        )
    return found


def _compare_page_setup(sheet: str, source: dict[str, Any], target: dict[str, Any]) -> list[Difference]:
    found: list[Difference] = []
    if source.get("orientation") != target.get("orientation"):
        found.append(
            Difference(
                "error",
                "page_orientation",
                f"{sheet}: page orientation changed.",
                sheet=sheet,
                template=source.get("orientation"),
                generated=target.get("orientation"),
            )
        )
    for key in ("header", "footer", "even_header", "even_footer", "first_header", "first_footer"):
        if _normalized_header(source.get(key)) != _normalized_header(target.get(key)):
            label = key.replace("_", " ")
            found.append(
                Difference(
                    "error",
                    "header_footer",
                    f"{sheet}: {label} changed.",
                    sheet=sheet,
                    template=source.get(key),
                    generated=target.get(key),
                )
            )
    for attr in ("paperSize", "fitToPage", "fitToWidth", "fitToHeight", "pageOrder", "margins"):
        if not _same_page_setup_value(source.get(attr), target.get(attr)):
            found.append(
                Difference(
                    "error",
                    "page_orientation",
                    f"{sheet}: page setup {attr} changed.",
                    sheet=sheet,
                    template=source.get(attr),
                    generated=target.get(attr),
                )
            )
    return found


def _compare_formulas(
    sheet: str,
    source: dict[str, str],
    target: dict[str, str],
    allowed_formulas: set[tuple[str, str]],
) -> list[Difference]:
    found: list[Difference] = []
    for coord, formula in source.items():
        if (sheet, coord.upper()) in allowed_formulas:
            continue
        other = target.get(coord)
        if other != formula:
            found.append(
                Difference(
                    "error",
                    "formula",
                    f"{sheet}!{coord}: formula was not preserved.",
                    sheet=sheet,
                    cell=coord,
                    template=formula,
                    generated=other,
                )
            )
    for coord, formula in target.items():
        if coord in source or (sheet, coord.upper()) in allowed_formulas:
            continue
        found.append(
            Difference(
                "error",
                "formula",
                f"{sheet}!{coord}: unexpected formula was added.",
                sheet=sheet,
                cell=coord,
                generated=formula,
            )
        )
    return found


def _compare_cells(
    sheet: str,
    source: dict[str, dict[str, Any]],
    target: dict[str, dict[str, Any]],
    *,
    allowed_formulas: set[tuple[str, str]],
    allowed_values: set[tuple[str, str]],
) -> list[Difference]:
    found: list[Difference] = []
    for coord, cell in source.items():
        other = target.get(coord)
        if other is None:
            found.append(
                Difference(
                    "error",
                    "font",
                    f"{sheet}!{coord}: cell metadata missing after generation.",
                    sheet=sheet,
                    cell=coord,
                    template=cell.get("coordinate"),
                )
            )
            continue
        found.extend(_compare_cell_layout(sheet, coord, cell, other))
        found.extend(
            _compare_cell_value(
                sheet,
                coord,
                cell,
                other,
                allowed_formulas=allowed_formulas,
                allowed_values=allowed_values,
            )
        )

    for coord, other in target.items():
        if coord in source:
            continue
        if other.get("formula") and (sheet, coord.upper()) not in allowed_formulas:
            continue
        if _has_unexpected_style(other) and not _is_expected_value_change(None, other, allowed=False):
            found.append(
                Difference(
                    "error",
                    "font",
                    f"{sheet}!{coord}: unexpected formatting was added.",
                    sheet=sheet,
                    cell=coord,
                    generated=other.get("font"),
                )
            )
    return found


def _compare_cell_layout(
    sheet: str,
    coord: str,
    cell: dict[str, Any],
    other: dict[str, Any],
) -> list[Difference]:
    found: list[Difference] = []
    source_font = cell.get("font") or {}
    target_font = other.get("font") or {}
    if source_font.get("name") != target_font.get("name"):
        found.append(
            Difference(
                "error",
                "font",
                f"{sheet}!{coord}: font changed.",
                sheet=sheet,
                cell=coord,
                template=source_font.get("name"),
                generated=target_font.get("name"),
            )
        )
    if not _same_number(
        cell.get("font_size") if cell.get("font_size") is not None else source_font.get("size"),
        other.get("font_size") if other.get("font_size") is not None else target_font.get("size"),
    ):
        found.append(
            Difference(
                "error",
                "font_size",
                f"{sheet}!{coord}: font size changed.",
                sheet=sheet,
                cell=coord,
                template=cell.get("font_size"),
                generated=other.get("font_size"),
            )
        )
    if bool(cell.get("bold")) != bool(other.get("bold")):
        found.append(
            Difference(
                "error",
                "bold",
                f"{sheet}!{coord}: bold flag changed.",
                sheet=sheet,
                cell=coord,
                template=bool(cell.get("bold")),
                generated=bool(other.get("bold")),
            )
        )
    if bool(cell.get("italic")) != bool(other.get("italic")):
        found.append(
            Difference(
                "error",
                "italic",
                f"{sheet}!{coord}: italic flag changed.",
                sheet=sheet,
                cell=coord,
                template=bool(cell.get("italic")),
                generated=bool(other.get("italic")),
            )
        )
    if cell.get("number_format") != other.get("number_format"):
        found.append(
            Difference(
                "error",
                "number_format",
                f"{sheet}!{coord}: number format changed.",
                sheet=sheet,
                cell=coord,
                template=cell.get("number_format"),
                generated=other.get("number_format"),
            )
        )
    if (cell.get("border") or {}) != (other.get("border") or {}):
        found.append(
            Difference(
                "error",
                "borders",
                f"{sheet}!{coord}: border changed.",
                sheet=sheet,
                cell=coord,
                template=cell.get("border"),
                generated=other.get("border"),
            )
        )
    if (cell.get("alignment") or {}) != (other.get("alignment") or {}):
        found.append(
            Difference(
                "error",
                "alignment",
                f"{sheet}!{coord}: alignment changed.",
                sheet=sheet,
                cell=coord,
                template=cell.get("alignment"),
                generated=other.get("alignment"),
            )
        )
    return found


def _compare_cell_value(
    sheet: str,
    coord: str,
    cell: dict[str, Any],
    other: dict[str, Any],
    *,
    allowed_formulas: set[tuple[str, str]],
    allowed_values: set[tuple[str, str]],
) -> list[Difference]:
    source_formula = cell.get("formula")
    target_formula = other.get("formula")
    allowed_formula = (sheet, coord.upper()) in allowed_formulas
    if source_formula and not allowed_formula and source_formula != target_formula:
        return []
    if cell.get("value") == other.get("value") and source_formula == target_formula:
        return []
    if source_formula and source_formula == target_formula:
        return []
    allowed = (sheet, coord.upper()) in allowed_values
    if _is_expected_value_change(cell, other, allowed=allowed):
        return [
            Difference(
                "ignored",
                "unexpected_value",
                f"{sheet}!{coord}: expected value change.",
                sheet=sheet,
                cell=coord,
                template=cell.get("value"),
                generated=other.get("value"),
            )
        ]
    return [
        Difference(
            "error",
            "unexpected_value",
            f"{sheet}!{coord}: unexpected value change.",
            sheet=sheet,
            cell=coord,
            template=cell.get("value"),
            generated=other.get("value"),
        )
    ]


def _is_expected_value_change(
    template_cell: dict[str, Any] | None,
    generated_cell: dict[str, Any],
    *,
    allowed: bool,
) -> bool:
    if allowed:
        return True
    if template_cell is None:
        return generated_cell.get("formula") is None and not _has_unexpected_style(generated_cell)
    if template_cell.get("formula"):
        return template_cell.get("formula") == generated_cell.get("formula")
    original = template_cell.get("value")
    if original is None or template_cell.get("blank"):
        return True
    if isinstance(original, (int, float)) and not isinstance(original, bool):
        return True
    if isinstance(original, str) and _is_dynamic_text(original):
        return True
    return False


def _is_dynamic_text(value: str) -> bool:
    text = value.strip()
    if not text:
        return True
    if PLACEHOLDER_RE.fullmatch(text):
        return True
    compact = text.replace(",", "").replace(" ", "")
    if NUMERIC_TEXT_RE.fullmatch(compact):
        return True
    if YEAR_RE.search(text) and len(text) <= 40:
        return True
    return False


def _has_unexpected_style(cell: dict[str, Any]) -> bool:
    font = cell.get("font") or {}
    if font.get("bold") or font.get("italic"):
        return True
    if font.get("name") not in {None, "Calibri"}:
        return True
    if cell.get("font_size") not in {None, 11}:
        return True
    if cell.get("number_format") not in {None, "General"}:
        return True
    if cell.get("border"):
        return True
    alignment = cell.get("alignment") or {}
    if any(alignment.get(key) for key in ("horizontal", "wrap_text", "indent", "text_rotation")):
        return True
    return False


def _same_dimension(left: Any, right: Any) -> bool:
    if left == right:
        return True
    if not isinstance(left, dict) or not isinstance(right, dict):
        return left is None and right is None
    keys = set(left) | set(right)
    for key in keys:
        if key in {"height", "width"}:
            if not _same_number(left.get(key), right.get(key)):
                return False
        elif left.get(key) != right.get(key):
            return False
    return True


def _same_number(left: Any, right: Any, places: int = 4) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    try:
        return round(float(left), places) == round(float(right), places)
    except (TypeError, ValueError):
        return left == right


def _same_page_setup_value(left: Any, right: Any) -> bool:
    if left == right:
        return True
    if isinstance(left, dict) and isinstance(right, dict):
        keys = set(left) | set(right)
        return all(_same_page_setup_value(left.get(key), right.get(key)) for key in keys)
    return _same_number(left, right, places=8)


def _normalized_set(values: Any) -> set[str]:
    if not values:
        return set()
    if isinstance(values, dict):
        values = values.keys()
    return {str(item).replace("$", "") for item in values}


def _normalize_print_area(value: Any) -> str | None:
    if not value:
        return None
    return str(value).replace("$", "")


def _normalized_header(value: Any) -> dict[str, str | None]:
    if not isinstance(value, dict):
        return {"left": None, "center": None, "right": None}
    return {
        "left": value.get("left") or None,
        "center": value.get("center") or None,
        "right": value.get("right") or None,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Compare the ICAI sample workbook with a generated copy."
    )
    parser.add_argument("--template", default=str(settings.template_path))
    parser.add_argument("--generated", default=str(settings.output_dir / DEFAULT_GENERATED_FILENAME))
    parser.add_argument("--json", action="store_true", help="Print the full JSON report.")
    args = parser.parse_args(argv)
    report = compare_workbooks(args.template, args.generated)
    if args.json:
        print(json.dumps(report.to_dict(), indent=2, ensure_ascii=False, default=str))
        return 0 if report.ok else 1
    print(f"Template: {report.template_path}")
    print(f"Generated: {report.generated_path}")
    print(f"Status: {'PASS' if report.ok else 'FAIL'}")
    print(f"Layout errors: {len(report.errors)}")
    print(f"Expected value changes ignored: {len(report.ignored)}")
    for item in report.errors:
        print(f"  [error] {item.message}")
    for item in report.warnings:
        print(f"  [warning] {item.message}")
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
