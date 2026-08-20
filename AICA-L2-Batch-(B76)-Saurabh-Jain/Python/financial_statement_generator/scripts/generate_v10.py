"""Generate Financial_Statements_Generated-10.xlsx."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app.trial_balance.generator import TrialBalanceStatementGenerator

TB_PATH  = Path("/Users/shruti/Downloads/TrialBal.xlsx")
OUT_FILE = "Financial_Statements_Generated-10.xlsx"


def main() -> None:
    gen    = TrialBalanceStatementGenerator()
    result = gen.generate_from_path(TB_PATH, output_filename=OUT_FILE)

    print(f"\n{'='*70}")
    print(f"OUTPUT: {result.path}")
    print(f"CELLS WRITTEN: {result.written_count}")
    print(f"{'='*70}\n")

    print("=== VALIDATION RESULTS ===\n")
    for check in result.validation.checks:
        icon = {"PASS": "✓", "FAIL": "✗", "WARNING": "⚠", "N.A.": "—", "INFO": "ℹ"}.get(check.status, "?")
        print(f"[{check.status:7s}] {icon} {check.name}")
        if check.detail:
            print(f"          {check.detail[:120]}")
    print()

    print("=== KEY METRICS ===")
    m = result.validation.metrics
    for key in [
        "tb_total_debit", "tb_total_credit",
        "total_assets", "total_equity", "total_liabilities",
        "profit_before_tax", "current_tax", "deferred_tax", "profit_after_tax",
        "mapped_accounts", "unmapped_accounts", "review_items",
    ]:
        val = m.get(key)
        if isinstance(val, float):
            print(f"  {key:<35s}: ₹{val:>15,.2f}")
        elif val is not None:
            print(f"  {key:<35s}: {val}")


if __name__ == "__main__":
    main()
