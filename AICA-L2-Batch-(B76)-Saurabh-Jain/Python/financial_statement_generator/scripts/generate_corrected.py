#!/usr/bin/env python3
"""Generate corrected financial statements from TrialBal.xlsx."""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app.trial_balance.generator import TrialBalanceStatementGenerator  # noqa: E402


def main() -> int:
    tb_path = Path("/Users/shruti/Downloads/TrialBal.xlsx")
    if not tb_path.exists():
        tb_path = ROOT / "input" / "TrialBal.xlsx"
    if not tb_path.exists():
        print(f"Trial balance not found: {tb_path}")
        return 1

    out_dir = ROOT / "output"
    generator = TrialBalanceStatementGenerator(output_dir=out_dir)
    result = generator.generate_from_path(tb_path, output_filename="Financial_Statements_Corrected.xlsx")

    report_path = out_dir / "validation_report.json"
    report_path.write_text(json.dumps(result.validation.to_dict(), indent=2), encoding="utf-8")

    downloads_copy = Path("/Users/shruti/Downloads/Financial_Statements_Corrected.xlsx")
    try:
        shutil.copy2(result.path, downloads_copy)
    except OSError:
        downloads_copy = None

    print(f"Generated: {result.path}")
    if downloads_copy:
        print(f"Copied to: {downloads_copy}")
    print(f"Validation report: {report_path}")
    print(f"Cells written: {result.written_count}")
    for check in result.validation.checks:
        print(f"  [{check.status}] {check.name}: {check.detail}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
