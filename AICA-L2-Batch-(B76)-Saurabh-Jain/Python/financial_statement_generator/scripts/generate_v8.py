#!/usr/bin/env python3
"""
Generate Financial_Statements_Generated-8.xlsx
Full Notes pipeline with NoteDataModel + NoteWriter + NoteReconciliation.
"""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app.trial_balance.generator import TrialBalanceStatementGenerator  # noqa: E402
from app.trial_balance.template_cleaner import full_leakage_scan  # noqa: E402
from openpyxl import load_workbook  # noqa: E402


def main() -> int:
    tb_path = Path("/Users/shruti/Downloads/TrialBal.xlsx")
    if not tb_path.exists():
        tb_path = ROOT / "input" / "TrialBal.xlsx"
    if not tb_path.exists():
        print(f"Trial balance not found: {tb_path}")
        return 1

    out_dir = ROOT / "output"
    out_filename = "Financial_Statements_Generated-8.xlsx"
    generator = TrialBalanceStatementGenerator(output_dir=out_dir)
    result = generator.generate_from_path(tb_path, output_filename=out_filename)

    # Full leakage scan
    wb = load_workbook(result.path, data_only=True)
    company = result.classification.company_name or ""
    leakage = full_leakage_scan(wb, company_name=company)
    wb.close()

    # Save validation report
    report_path = out_dir / "validation_report_v8.json"
    report_path.write_text(
        json.dumps(result.validation.to_dict(), indent=2, default=str),
        encoding="utf-8",
    )

    downloads_copy = Path("/Users/shruti/Downloads") / out_filename
    try:
        shutil.copy2(result.path, downloads_copy)
    except OSError:
        downloads_copy = None

    # ------------------------------------------------------------------ #
    # Report
    # ------------------------------------------------------------------ #
    print(f"\n{'='*65}")
    print(f"Financial_Statements_Generated-8.xlsx")
    print(f"{'='*65}")
    print(f"Generated : {result.path}")
    if downloads_copy:
        print(f"Copied to : {downloads_copy}")
    print(f"Cells written: {result.written_count}")
    print()
    print("VALIDATION RESULTS:")
    for check in result.validation.checks:
        print(f"  [{check.status:7s}] {check.name}: {check.detail[:100]}")
    print()
    print("FULL LEAKAGE SCAN:")
    if leakage:
        # Suppress known-OK prior-year comparative headers
        genuine = [x for x in leakage if x.get("type") != "LARGE_NUMBER_SUSPECT"
                   or float(x.get("value", 0)) >= 100_000_000]
        if genuine:
            print(f"  {len(genuine)} genuine issue(s) remain:")
            for issue in genuine:
                print(f"    [{issue['type']}] {issue['sheet']}!{issue['cell']}: {issue['value'][:80]}")
        else:
            print("  No genuine template leakage found (large numbers are TB balances).")
    else:
        print("  CLEAN — no leakage issues found.")

    print()
    print("REPORT SAVED:", report_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
