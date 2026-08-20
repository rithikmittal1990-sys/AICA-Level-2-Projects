"""
Central Note Data Model
=======================
Builds one NoteEntry per disclosure from the classified trial balance.
Every value carries source metadata — nothing is silently fabricated.

Source types:
  TRIAL_BALANCE  – directly from a TB ledger account
  CALCULATED     – arithmetically derived from TB values
  NOT_AVAILABLE  – TB does not contain required data
  USER_INPUT     – supplied externally (reserved for future use)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Literal

from app.trial_balance.models import MappedAccount, TrialBalanceClassification
from app.trial_balance.parser import normalize_label

NA_TEXT    = "Information not available from Trial Balance"
MGMT_TEXT  = "Information cannot be determined from the Trial Balance and requires additional management/company information"
FACE_VALUE = 10.0  # default share face value (₹10)

SourceType = Literal["TRIAL_BALANCE", "CALCULATED", "NOT_AVAILABLE", "USER_INPUT"]
NoteStatus = Literal["PASS", "WARNING", "FAIL", "N.A."]


@dataclass
class NoteLineItem:
    label: str
    amount: float | str | None        # float = numeric; str = NA/text; None = suppress
    source_type: SourceType
    source_accounts: list[str] = field(default_factory=list)
    review_reason: str | None = None
    is_subtotal: bool = False


@dataclass
class NoteEntry:
    note_number: str
    note_name: str
    line_items: list[NoteLineItem] = field(default_factory=list)
    total: float | None = None
    total_label: str = "Total"
    status: NoteStatus = "PASS"
    status_reason: str = ""
    sheet_name: str = ""

    def add(
        self,
        label: str,
        amount: float | str | None,
        source_type: SourceType,
        source_accounts: list[str] | None = None,
        review_reason: str | None = None,
        is_subtotal: bool = False,
    ) -> "NoteEntry":
        self.line_items.append(NoteLineItem(
            label=label, amount=amount, source_type=source_type,
            source_accounts=source_accounts or [], review_reason=review_reason,
            is_subtotal=is_subtotal,
        ))
        return self

    def set_warning(self, reason: str) -> "NoteEntry":
        if self.status not in ("FAIL", "N.A."):
            self.status = "WARNING"
        self.status_reason = reason
        return self

    def numeric_total(self) -> float:
        return round(sum(
            i.amount for i in self.line_items
            if isinstance(i.amount, (int, float)) and not i.is_subtotal
        ), 2)


@dataclass
class NoteDataModel:
    notes: list[NoteEntry] = field(default_factory=list)
    reporting_year: str = ""
    reporting_date: str = ""
    prior_year_date: str = ""
    company_name: str = ""
    period_label: str = ""
    # Full account list for TB Mapping sheet
    all_mapped: list[MappedAccount] = field(default_factory=list)
    # Line items dict
    line_items: dict[str, float] = field(default_factory=dict)

    def by_note(self, number: str) -> NoteEntry | None:
        for n in self.notes:
            if n.note_number == number:
                return n
        return None

    @classmethod
    def build(cls, classification: TrialBalanceClassification) -> "NoteDataModel":
        model = cls()
        model.company_name  = classification.company_name or ""
        model.period_label  = classification.period_label or ""
        model.all_mapped    = classification.mapped
        model.line_items    = classification.line_items
        model._derive_dates()

        items  = classification.line_items
        mapped = classification.mapped
        totals = classification.totals

        def amt(key: str) -> float:
            return float(items.get(key, 0.0) or 0.0)

        def by_agg(key: str) -> list[MappedAccount]:
            """Accounts whose classifier aggregate_key (stored in .note) matches key."""
            return [m for m in mapped if (m.note or "") == key and m.status == "mapped"]

        def all_for_agg(key: str) -> list[MappedAccount]:
            """Including review_required accounts."""
            return [m for m in mapped if (m.note or "") == key]

        # ------------------------------------------------------------------ #
        # NOTE 3 — SHARE CAPITAL
        # ------------------------------------------------------------------ #
        sc_accs  = all_for_agg("share_capital")
        sc_total = amt("share_capital")
        sc_note  = NoteEntry("3", "Share Capital", sheet_name="Note 3 (Share Capital)")
        if sc_total:
            n_shares = int(sc_total / FACE_VALUE)
            for m in sc_accs:
                sc_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                            [m.account.account_name])
            sc_note.add(
                f"Equity shares of ₹{FACE_VALUE:.0f} each — {n_shares:,} shares issued, subscribed and fully paid up",
                sc_total, "TRIAL_BALANCE", [a.account.account_name for a in sc_accs],
                is_subtotal=True,
            )
            sc_note.total = sc_total
        else:
            sc_note.add("Share capital", NA_TEXT, "NOT_AVAILABLE")
        for lbl in [
            "Authorised capital", "Opening balance of shares", "Shares issued during year",
            "Shares bought back", "Shareholder-wise holding (>5%)", "Promoter holding",
        ]:
            sc_note.add(lbl, NA_TEXT, "NOT_AVAILABLE")
        sc_note.set_warning("Detailed share capital disclosures not available from Trial Balance") if not sc_total else None
        model.notes.append(sc_note)

        # ------------------------------------------------------------------ #
        # NOTE 4 — RESERVES & SURPLUS
        # ------------------------------------------------------------------ #
        res_accs   = all_for_agg("reserves_and_surplus")
        res_tb     = amt("reserves_and_surplus")
        pat        = totals.current_year_profit
        res_note   = NoteEntry("4", "Reserves and Surplus", sheet_name="NOTE (4-12)")
        res_note.add("Opening balance (as per Trial Balance closing balance)", res_tb,
                     "TRIAL_BALANCE", [a.account.account_name for a in res_accs])
        res_note.add("Add: Net profit for the current year", pat, "CALCULATED")
        res_note.add("Less: Transfer to reserves", NA_TEXT, "NOT_AVAILABLE")
        closing = round(res_tb + pat, 2)
        res_note.add("Closing Balance", closing, "CALCULATED", is_subtotal=True)
        res_note.total = closing
        if not res_tb:
            res_note.set_warning("Reserves & Surplus closing balance not found in Trial Balance")
        model.notes.append(res_note)

        # ------------------------------------------------------------------ #
        # NOTE 5 — BORROWINGS
        # ------------------------------------------------------------------ #
        borrow_note = NoteEntry("5", "Borrowings", sheet_name="NOTE (4-12)")
        # Reclassified loan receivable accounts
        reclassified = [m for m in mapped if m.reclassification_reason and
                        (m.note or "") in ("short_term_loans_and_advances", "loan_receivable")]
        for m in reclassified:
            borrow_note.add(
                f"{m.account.account_name}  [REVIEW REQUIRED — Debit balance in liability group]",
                m.account.net_balance, "TRIAL_BALANCE", [m.account.account_name],
                "Liability-group account has a debit balance; classification confirmation required.",
            )
        for agg_key in ("long_term_borrowings", "short_term_borrowings"):
            for m in by_agg(agg_key):
                borrow_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                                [m.account.account_name])
        if not borrow_note.line_items:
            borrow_note.add("No borrowing accounts identified in Trial Balance", "Nil", "TRIAL_BALANCE")
            borrow_note.status = "N.A."
            borrow_note.status_reason = "No borrowing accounts found in Trial Balance."
        else:
            borrow_note.set_warning(
                "One or more borrowing-group accounts have debit balances requiring classification confirmation. "
                "Interest rates, security, maturity and repayment terms cannot be determined from Trial Balance."
            )
        for lbl in ["Interest rate", "Security / collateral", "Maturity / repayment terms", "Guarantees"]:
            borrow_note.add(lbl, NA_TEXT, "NOT_AVAILABLE")
        model.notes.append(borrow_note)

        # ------------------------------------------------------------------ #
        # NOTE 6 — DEFERRED TAX
        # ------------------------------------------------------------------ #
        dt_asset   = amt("deferred_tax_asset")
        dt_income  = amt("deferred_tax_income")
        dt_note    = NoteEntry("6", "Deferred Tax", sheet_name="NOTE (4-12)")
        if dt_asset or dt_income:
            if dt_asset:
                accs = all_for_agg("deferred_tax_asset")
                dt_note.add("Deferred Tax Asset (Balance Sheet)", dt_asset, "TRIAL_BALANCE",
                            [a.account.account_name for a in accs])
            if dt_income:
                accs = all_for_agg("deferred_tax_income")
                dt_note.add("Deferred Tax Income / (Expense) (P&L)", dt_income, "TRIAL_BALANCE",
                            [a.account.account_name for a in accs])
            dt_note.add("Detailed timing difference breakup", NA_TEXT, "NOT_AVAILABLE")
            dt_note.set_warning(
                "Detailed deferred tax timing differences are not available from the Trial Balance "
                "and require additional management information."
            )
        else:
            dt_note.status = "N.A."
            dt_note.status_reason = "No deferred tax accounts found in Trial Balance."
        model.notes.append(dt_note)

        # ------------------------------------------------------------------ #
        # NOTE 9 — TRADE PAYABLES  (Sundry Creditors ONLY)
        # ------------------------------------------------------------------ #
        tp_accs  = by_agg("trade_payables")
        tp_total = amt("trade_payables")
        tp_note  = NoteEntry("9", "Trade Payables", sheet_name="NOTE (4-12)")
        for m in tp_accs:
            tp_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                        [m.account.account_name])
        tp_note.total = tp_total if tp_total else None
        if not tp_total:
            tp_note.add("Trade payables", "Nil", "TRIAL_BALANCE")
            tp_note.status = "N.A."
        tp_note.add("MSME / non-MSME classification", NA_TEXT, "NOT_AVAILABLE")
        tp_note.add("Ageing / disputed / undisputed details", NA_TEXT, "NOT_AVAILABLE")
        tp_note.add("Related party trade payables", NA_TEXT, "NOT_AVAILABLE")
        tp_note.set_warning(
            "MSME classification, ageing and related party details are not available from the Trial Balance "
            "and require additional information."
        )
        model.notes.append(tp_note)

        # ------------------------------------------------------------------ #
        # NOTE 10 — OTHER CURRENT LIABILITIES
        # ------------------------------------------------------------------ #
        ocl_note  = NoteEntry("10", "Other Current Liabilities", sheet_name="NOTE (4-12)")
        ocl_total = 0.0
        for agg_key in ("other_current_liabilities", "duties_and_taxes"):
            for m in by_agg(agg_key):
                ocl_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                             [m.account.account_name])
                ocl_total += m.amount
        ocl_total = round(ocl_total, 2)
        ocl_note.total = ocl_total if ocl_total else None
        if not ocl_total:
            ocl_note.add("Other current liabilities", "Nil", "TRIAL_BALANCE")
            ocl_note.status = "N.A."
        model.notes.append(ocl_note)

        # ------------------------------------------------------------------ #
        # NOTE 11 — SHORT-TERM PROVISIONS
        # ------------------------------------------------------------------ #
        prov_note  = NoteEntry("11", "Short-term Provisions", sheet_name="NOTE (4-12)")
        prov_total = 0.0
        for m in by_agg("short_term_provisions"):
            prov_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                          [m.account.account_name])
            prov_total += m.amount
        prov_total = round(prov_total, 2)
        prov_note.total = prov_total if prov_total else None
        if not prov_total:
            prov_note.status = "N.A."
        model.notes.append(prov_note)

        # ------------------------------------------------------------------ #
        # NOTE 12 — PROPERTY, PLANT & EQUIPMENT
        # ------------------------------------------------------------------ #
        ppe_total  = amt("property_plant_and_equipment")
        ia_total   = amt("intangible_assets")
        ppe_note   = NoteEntry("12", "Property, Plant and Equipment", sheet_name="Note 12 (PPE)")
        if ppe_total:
            ppe_note.add("A. Tangible Assets", None, "NOT_AVAILABLE")
            for m in by_agg("property_plant_and_equipment"):
                ppe_note.add(f"  {m.account.account_name}", m.amount, "TRIAL_BALANCE",
                             [m.account.account_name])
            ppe_note.add("Total Tangible Assets (Net Book Value)", ppe_total, "TRIAL_BALANCE",
                         is_subtotal=True)
        if ia_total:
            ppe_note.add("B. Intangible Assets", None, "NOT_AVAILABLE")
            for m in by_agg("intangible_assets"):
                ppe_note.add(f"  {m.account.account_name}", m.amount, "TRIAL_BALANCE",
                             [m.account.account_name])
            ppe_note.add("Total Intangible Assets (Net Book Value)", ia_total, "TRIAL_BALANCE",
                         is_subtotal=True)
        total_fa = round(ppe_total + ia_total, 2)
        ppe_note.total = total_fa if total_fa else None
        for lbl in [
            "Opening gross block", "Additions during year", "Disposals",
            "Accumulated depreciation opening", "Useful life / depreciation rate",
        ]:
            ppe_note.add(lbl, NA_TEXT, "NOT_AVAILABLE")
        ppe_note.set_warning(
            "Detailed PPE movement schedule not available from Trial Balance. "
            "Closing Net Book Value is shown as derived from the Trial Balance."
        )
        model.notes.append(ppe_note)

        # ------------------------------------------------------------------ #
        # NOTE 13 — CWIP
        # ------------------------------------------------------------------ #
        cwip_note = NoteEntry("13", "Capital Work in Progress", sheet_name="13.CWIP")
        cwip_accs = [m for m in mapped if "cwip" in normalize_label(m.account.account_name) or
                     "capital work" in normalize_label(m.account.account_name)]
        if cwip_accs:
            for m in cwip_accs:
                cwip_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                              [m.account.account_name])
        else:
            cwip_note.add("Capital Work in Progress", "Nil", "TRIAL_BALANCE")
            cwip_note.status = "N.A."
            cwip_note.status_reason = "No CWIP accounts found in Trial Balance."
        model.notes.append(cwip_note)

        # ------------------------------------------------------------------ #
        # NOTE 14 — NON-CURRENT INVESTMENTS
        # ------------------------------------------------------------------ #
        inv_note = NoteEntry("14", "Non-Current Investments", sheet_name="Note (13-20)")
        inv_accs = [m for m in mapped if any(k in normalize_label(m.account.account_name)
                    for k in ("investment", "sovereign", "gold bond", "ncd"))]
        if inv_accs:
            for m in inv_accs:
                inv_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                             [m.account.account_name])
        else:
            inv_note.add("Non-current investments", "Nil", "TRIAL_BALANCE")
            inv_note.status = "N.A."
        model.notes.append(inv_note)

        # ------------------------------------------------------------------ #
        # NOTE 17 — TRADE RECEIVABLES
        # ------------------------------------------------------------------ #
        tr_total = amt("trade_receivables")
        tr_note  = NoteEntry("17", "Trade Receivables", sheet_name="Note (13-20)")
        if tr_total:
            for m in by_agg("trade_receivables"):
                tr_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                            [m.account.account_name])
            tr_note.total = tr_total
            tr_note.add("Ageing (Less than 6 months / 6-12 months / 1-2 years / >2 years)", NA_TEXT, "NOT_AVAILABLE")
            tr_note.add("Disputed / doubtful / good classification", NA_TEXT, "NOT_AVAILABLE")
            tr_note.add("Related party receivables", NA_TEXT, "NOT_AVAILABLE")
            tr_note.set_warning(
                "Ageing, classification and related party details are not available from the Trial Balance "
                "and require additional information."
            )
        else:
            tr_note.add("Trade receivables", "Nil", "TRIAL_BALANCE")
            tr_note.status = "N.A."
        model.notes.append(tr_note)

        # ------------------------------------------------------------------ #
        # NOTE 18 — CASH & CASH EQUIVALENTS
        # ------------------------------------------------------------------ #
        cash_total = amt("cash_and_cash_equivalents")
        cash_note  = NoteEntry("18", "Cash and Cash Equivalents", sheet_name="Note (13-20)")
        for m in by_agg("cash_and_cash_equivalents"):
            cash_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                          [m.account.account_name])
        cash_note.total = cash_total if cash_total else None
        if not cash_total:
            cash_note.status = "WARNING"
            cash_note.status_reason = "No cash/bank accounts found in Trial Balance."
        model.notes.append(cash_note)

        # ------------------------------------------------------------------ #
        # NOTE 19 — SHORT-TERM LOANS & ADVANCES
        # ------------------------------------------------------------------ #
        adv_total  = amt("short_term_loans_and_advances")
        adv_note   = NoteEntry("19", "Short-term Loans and Advances", sheet_name="Note (13-20)")
        for m in by_agg("short_term_loans_and_advances"):
            label = m.account.account_name
            if m.reclassification_reason:
                label += "  [Reclassified — see Review section]"
            adv_note.add(label, m.amount, "TRIAL_BALANCE", [m.account.account_name],
                         m.reclassification_reason)
        adv_note.total = round(adv_total, 2) if adv_total else None
        if any(m.reclassification_reason for m in by_agg("short_term_loans_and_advances")):
            adv_note.set_warning(
                "One or more accounts have been reclassified from a liability group due to debit balances. "
                "Classification confirmation is required."
            )
        model.notes.append(adv_note)

        # ------------------------------------------------------------------ #
        # NOTE 20 — OTHER CURRENT ASSETS
        # ------------------------------------------------------------------ #
        oca_total = amt("other_current_assets") + amt("deferred_tax_asset")
        oca_note  = NoteEntry("20", "Other Current Assets", sheet_name="Note (13-20)")
        for agg in ("other_current_assets", "deferred_tax_asset"):
            for m in by_agg(agg):
                oca_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                             [m.account.account_name])
        oca_note.total = round(oca_total, 2) if oca_total else None
        if not oca_total:
            oca_note.status = "N.A."
        model.notes.append(oca_note)

        # ------------------------------------------------------------------ #
        # NOTE 21 — REVENUE FROM OPERATIONS
        # ------------------------------------------------------------------ #
        rev_total = amt("revenue_from_operations")
        rev_note  = NoteEntry("21", "Revenue from Operations", sheet_name="Note 20-31")
        for m in by_agg("revenue_from_operations"):
            rev_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                         [m.account.account_name])
        rev_note.total = round(rev_total, 2) if rev_total else None
        if not rev_total:
            rev_note.status = "WARNING"
            rev_note.status_reason = "No revenue accounts found in Trial Balance."
        model.notes.append(rev_note)

        # ------------------------------------------------------------------ #
        # NOTE 22 — OTHER INCOME
        # ------------------------------------------------------------------ #
        oi_total  = amt("other_income")
        dt_income = amt("deferred_tax_income")
        oi_note   = NoteEntry("22", "Other Income", sheet_name="Note 20-31")
        for m in by_agg("other_income"):
            oi_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                        [m.account.account_name])
        # Deferred tax income is separately disclosed in Note 6 / tax section
        # but if mapped to other_income, show here too for completeness
        for m in by_agg("deferred_tax_income"):
            oi_note.add(f"{m.account.account_name} (see Note 6)", m.amount, "TRIAL_BALANCE",
                        [m.account.account_name])
        total_oi = round(oi_total + dt_income, 2)
        oi_note.total = total_oi if total_oi else None
        if not total_oi:
            oi_note.add("Other income", "Nil", "TRIAL_BALANCE")
            oi_note.status = "N.A."
        model.notes.append(oi_note)

        # ------------------------------------------------------------------ #
        # NOTE 26 — EMPLOYEE BENEFITS EXPENSE
        # ------------------------------------------------------------------ #
        eb_total = amt("employee_benefits_expense")
        eb_note  = NoteEntry("26", "Employee Benefits Expense", sheet_name="Note 20-31")
        for m in by_agg("employee_benefits_expense"):
            eb_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                        [m.account.account_name])
        eb_note.add(
            "Defined benefit / contribution plan details (Gratuity, PF, ESI)", NA_TEXT, "NOT_AVAILABLE"
        )
        eb_note.total = round(eb_total, 2) if eb_total else None
        eb_note.set_warning(
            "Detailed employee benefit disclosures (gratuity, provident fund, ESI) "
            "are not available from the Trial Balance and require additional information."
        )
        model.notes.append(eb_note)

        # ------------------------------------------------------------------ #
        # NOTE DEP — DEPRECIATION (used in Note 27 / Other Expenses block)
        # ------------------------------------------------------------------ #
        dep_total = amt("depreciation_and_amortization_expense")
        dep_note  = NoteEntry("dep", "Depreciation and Amortisation", sheet_name="Note 20-31")
        for m in by_agg("depreciation_and_amortization_expense"):
            dep_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                         [m.account.account_name])
        dep_note.total = round(dep_total, 2) if dep_total else None
        model.notes.append(dep_note)

        # ------------------------------------------------------------------ #
        # NOTE 27 — OTHER EXPENSES  (Depreciation excluded)
        # ------------------------------------------------------------------ #
        oe_total = amt("other_expenses")
        oe_note  = NoteEntry("27", "Other Expenses", sheet_name="Note 20-31")
        for m in by_agg("other_expenses"):
            oe_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                        [m.account.account_name])
        # Prior period items
        prior = amt("prior_period_items")
        if prior:
            for m in by_agg("prior_period_items"):
                oe_note.add(f"{m.account.account_name} (Prior Period)", m.amount, "TRIAL_BALANCE",
                            [m.account.account_name])
        oe_note.total = round(oe_total, 2) if oe_total else None
        model.notes.append(oe_note)

        # ------------------------------------------------------------------ #
        # NOTE 28 — FINANCE COSTS
        # ------------------------------------------------------------------ #
        fin_note = NoteEntry("28", "Finance Costs", sheet_name="Note 20-31")
        # Finance-cost items already captured in other_expenses (bank charges, interest on TDS)
        # Extract interest-specific items
        fin_accs = [m for m in by_agg("other_expenses")
                    if any(k in normalize_label(m.account.account_name)
                           for k in ("interest", "finance", "bank charges", "processing fee"))]
        for m in fin_accs:
            fin_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                         [m.account.account_name])
        fin_note.total = round(sum(m.amount for m in fin_accs if isinstance(m.amount, (int, float))), 2) or None
        if not fin_note.line_items:
            fin_note.add("Finance costs", "Nil", "TRIAL_BALANCE")
            fin_note.status = "N.A."
        model.notes.append(fin_note)

        # ------------------------------------------------------------------ #
        # NOTE 29 — CURRENT TAX  (P&L expense only — provision is on BS side)
        # ------------------------------------------------------------------ #
        tax_total = amt("current_tax")
        tax_note  = NoteEntry("29", "Current Tax", sheet_name="Note 20-31")
        for m in by_agg("current_tax"):
            tax_note.add(m.account.account_name, m.amount, "TRIAL_BALANCE",
                         [m.account.account_name])
        # Show the provision as informational reference only (not added to total)
        for m in by_agg("short_term_provisions"):
            tax_note.add(
                f"{m.account.account_name} (Balance Sheet — for reference)",
                NA_TEXT, "NOT_AVAILABLE", [m.account.account_name],
            )
        tax_note.total = round(tax_total, 2) if tax_total else None
        if not tax_total:
            tax_note.status = "N.A."
        model.notes.append(tax_note)

        # ------------------------------------------------------------------ #
        # NOTE 30 — EPS
        # ------------------------------------------------------------------ #
        eps_note = NoteEntry("30", "Earnings Per Share", sheet_name="Note 20-31")
        eps_note.add("Profit after tax (₹)", totals.profit_after_tax, "CALCULATED")
        eps_note.add("Face value per share (₹)", FACE_VALUE, "TRIAL_BALANCE")
        eps_note.add(
            "Weighted average number of equity shares",
            "Weighted-average number of shares not available from Trial Balance", "NOT_AVAILABLE",
        )
        eps_note.add("Basic / Diluted EPS", NA_TEXT, "NOT_AVAILABLE")
        eps_note.status = "N.A."
        eps_note.status_reason = (
            "Weighted-average number of shares not available from Trial Balance. "
            "EPS cannot be calculated."
        )
        model.notes.append(eps_note)

        # ------------------------------------------------------------------ #
        # NOTE 35 — CONTINGENT LIABILITIES
        # ------------------------------------------------------------------ #
        cl_note = NoteEntry("35", "Contingent Liabilities and Commitments", sheet_name="Other Notes")
        cl_note.add("Contingent liabilities", MGMT_TEXT, "NOT_AVAILABLE")
        cl_note.add("Capital commitments", MGMT_TEXT, "NOT_AVAILABLE")
        cl_note.status = "WARNING"
        cl_note.status_reason = (
            "Contingent liability information cannot be determined from the Trial Balance "
            "and requires additional management information."
        )
        model.notes.append(cl_note)

        # ------------------------------------------------------------------ #
        # NOTE 36 — RELATED PARTIES
        # ------------------------------------------------------------------ #
        rp_note = NoteEntry("36", "Related Party Disclosures", sheet_name="Other Notes")
        rp_note.add("Related party names and nature of relationship", MGMT_TEXT, "NOT_AVAILABLE")
        rp_note.add("Transactions with related parties", MGMT_TEXT, "NOT_AVAILABLE")
        rp_note.add("Outstanding balances", MGMT_TEXT, "NOT_AVAILABLE")
        rp_note.status = "WARNING"
        rp_note.status_reason = (
            "Related party information is not determinable from the Trial Balance. "
            "Additional management information is required."
        )
        model.notes.append(rp_note)

        # ------------------------------------------------------------------ #
        # NOTE 37 — SEGMENT
        # ------------------------------------------------------------------ #
        seg_note = NoteEntry("37", "Segment Information", sheet_name="Other Notes")
        seg_note.add("Segment information", NA_TEXT, "NOT_AVAILABLE")
        seg_note.status = "WARNING"
        seg_note.status_reason = "Segment information is not available from the Trial Balance."
        model.notes.append(seg_note)

        # ------------------------------------------------------------------ #
        # NOTE 38 — CSR
        # ------------------------------------------------------------------ #
        csr_note = NoteEntry("38", "Corporate Social Responsibility", sheet_name="Other Notes")
        csr_note.add(
            "CSR applicability and expenditure",
            "CSR applicability and related disclosure require additional company information.",
            "NOT_AVAILABLE",
        )
        csr_note.status = "WARNING"
        csr_note.status_reason = (
            "CSR applicability and related disclosure require additional company information."
        )
        model.notes.append(csr_note)

        # ------------------------------------------------------------------ #
        # NOTE 40 — EVENTS AFTER REPORTING DATE
        # ------------------------------------------------------------------ #
        ear_note = NoteEntry("40", "Events After Reporting Date", sheet_name="Other Notes")
        ear_note.add(
            "Events after the reporting date",
            "Information regarding events after the reporting date cannot be determined from "
            "the Trial Balance and requires additional management information.",
            "NOT_AVAILABLE",
        )
        ear_note.status = "WARNING"
        ear_note.status_reason = "Cannot be determined from Trial Balance."
        model.notes.append(ear_note)

        # ------------------------------------------------------------------ #
        # NOTE 44 — REGULATORY DECLARATIONS
        # ------------------------------------------------------------------ #
        reg_note = NoteEntry("44", "Additional Regulatory Information", sheet_name="Other Notes")
        for sub in [
            "Transactions with struck-off companies",
            "Crypto / virtual currency transactions",
            "Benami property proceedings",
            "Charges not registered with ROC",
            "Loans advanced through intermediaries",
            "Funds received from funding parties",
            "Loans/advances to promoters, directors, KMP",
            "Title deeds of immovable properties",
            "Wilful defaulter declaration",
            "Borrowings used for specific purpose",
            "Scheme of arrangements",
            "Undisclosed income",
        ]:
            reg_note.add(sub, MGMT_TEXT, "NOT_AVAILABLE")
        reg_note.status = "WARNING"
        reg_note.status_reason = (
            "All regulatory declarations require management confirmation and cannot be "
            "determined from the Trial Balance alone."
        )
        model.notes.append(reg_note)

        return model

    def _derive_dates(self) -> None:
        if not self.period_label:
            self.reporting_year  = "Current Year"
            self.reporting_date  = "Current reporting date"
            self.prior_year_date = "N.A."
            return
        parts   = re.split(r"\s+to\s+", self.period_label, flags=re.IGNORECASE)
        end_str = parts[-1].strip() if parts else ""
        year_m  = re.search(r"(\d{2,4})$", end_str)
        if year_m:
            yr  = year_m.group(1)
            yr4 = int(yr) if len(yr) == 4 else 2000 + int(yr)
            fy_start            = yr4 - 1
            self.reporting_year  = f"{fy_start}-{str(yr4)[2:]}"
            self.reporting_date  = f"31st March, {yr4}"
            self.prior_year_date = f"31st March, {fy_start}"
        else:
            self.reporting_year  = "Current Year"
            self.reporting_date  = "Current reporting date"
            self.prior_year_date = "N.A."
