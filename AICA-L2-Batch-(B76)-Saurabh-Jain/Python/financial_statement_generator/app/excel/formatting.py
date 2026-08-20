"""Safe Excel write helpers that keep template formatting intact.

These utilities change only a cell's value or formula. Fonts, fills, borders,
alignment, number formats, protection, merged ranges, and surrounding layout
are restored after the write.
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from typing import Any

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True, slots=True)
class CellWriteResult:
    """Outcome of a single safe write."""

    coordinate: str
    written: bool
    skipped: bool
    reason: str | None = None


@dataclass(frozen=True, slots=True)
class CellStyleSnapshot:
    font: Any
    fill: Any
    border: Any
    alignment: Any
    number_format: str
    protection: Any


def snapshot_cell_style(cell: Cell) -> CellStyleSnapshot:
    """Capture style attributes so they can be restored after a value write."""
    return CellStyleSnapshot(
        font=copy(cell.font),
        fill=copy(cell.fill),
        border=copy(cell.border),
        alignment=copy(cell.alignment),
        number_format=cell.number_format,
        protection=copy(cell.protection),
    )


def restore_cell_style(cell: Cell, snapshot: CellStyleSnapshot) -> None:
    """Re-apply a previously captured style. Does not change the cell value."""
    cell.font = snapshot.font
    cell.fill = snapshot.fill
    cell.border = snapshot.border
    cell.alignment = snapshot.alignment
    cell.number_format = snapshot.number_format
    cell.protection = snapshot.protection


def cell_has_formula(cell: Cell) -> bool:
    value = cell.value
    if isinstance(value, ArrayFormula):
        return True
    if isinstance(value, DataTableFormula):
        return True
    if isinstance(value, str) and value.startswith("="):
        return True
    return cell.data_type == "f" and value is not None


def writable_coordinate(worksheet: Worksheet, coordinate: str) -> str:
    """Return the top-left cell of a merge, otherwise the original coordinate."""
    coord = coordinate.upper()
    for merged in worksheet.merged_cells.ranges:
        if coord in merged:
            return worksheet.cell(merged.min_row, merged.min_col).coordinate
    return coord


def write_value(
    worksheet: Worksheet,
    coordinate: str,
    value: float | int | None,
    *,
    overwrite_formula: bool = False,
) -> CellWriteResult:
    """Write a numeric value without clearing formatting or formulas."""
    return _write(worksheet, coordinate, value, overwrite_formula=overwrite_formula, kind="value")


def write_text(
    worksheet: Worksheet,
    coordinate: str,
    value: str | None,
    *,
    overwrite_formula: bool = False,
) -> CellWriteResult:
    """Write text without clearing formatting or formulas."""
    return _write(worksheet, coordinate, value, overwrite_formula=overwrite_formula, kind="text")


def write_formula(
    worksheet: Worksheet,
    coordinate: str,
    formula: str,
    *,
    overwrite_formula: bool = False,
) -> CellWriteResult:
    """Write a formula only when explicitly allowed."""
    text = formula if str(formula).startswith("=") else f"={formula}"
    return _write(worksheet, coordinate, text, overwrite_formula=overwrite_formula, kind="formula")


def apply_standard_formatting(worksheet: Worksheet) -> None:
    """Leave template formatting untouched.

    Schedule III presentation comes from ``Financial Statements_Sample.xlsx``.
    This function is intentionally a no-op so generation cannot restyle sheets.
    """
    return None


def _write(
    worksheet: Worksheet,
    coordinate: str,
    value: Any,
    *,
    overwrite_formula: bool,
    kind: str,
) -> CellWriteResult:
    target = writable_coordinate(worksheet, coordinate)
    cell = worksheet[target]
    if cell_has_formula(cell) and not overwrite_formula:
        return CellWriteResult(
            coordinate=target,
            written=False,
            skipped=True,
            reason="formula",
        )
    if kind != "formula" and value is None:
        return CellWriteResult(coordinate=target, written=False, skipped=True, reason="empty")
    snapshot = snapshot_cell_style(cell)
    cell.value = value
    restore_cell_style(cell, snapshot)
    return CellWriteResult(coordinate=target, written=True, skipped=False)
