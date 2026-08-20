"""Trial balance → financial statements pipeline (TB is the only numeric source)."""

from app.trial_balance.generator import TrialBalanceStatementGenerator

__all__ = ["TrialBalanceStatementGenerator"]
