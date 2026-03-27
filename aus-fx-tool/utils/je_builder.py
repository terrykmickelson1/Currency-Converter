"""
Journal Entry builder.

Produces the list of JE lines from translated BS and IS data, then
formats them for QuickBooks Online import (CSV).

Structure of the JE:
  1. BS section   – net USD movement for each non-equity BS account
                    memo: "Adjusting JE to reflect Australia BS Account Balances at month-end"
  2. IS/Equity    – USD movement for equity BS accounts (je_section='IS')
                    + IS activity for income/expense accounts
                    memo: "Adjusting JE to reflect Australia IS activity"
  3. Net income   – offsetting entry to Equity in Subsidiary
                    memo: "Australia P&L-Net Income / (Loss) for the month"
  4. Exch adj     – two-line entry (Equity in Sub + Currency Exchange Gain/Loss)
                    memo: "Australia Currency Exchange Gains/(Losses)"
"""
import io
import csv
from datetime import date
from typing import Any

import pandas as pd


# ---------------------------------------------------------------------------
# Debit / Credit logic
# ---------------------------------------------------------------------------

def _bs_debit_credit(xero_type: str, movement: float) -> tuple[float | None, float | None]:
    """
    For a BS-section account, determine which side receives the movement.

    Assets have a debit normal balance → increase = debit, decrease = credit.
    Liabilities & Equity have a credit normal balance → increase = credit, decrease = debit.
    """
    if abs(movement) < 0.005:
        return None, None
    amt = abs(movement)
    is_asset = xero_type in ("Asset",)
    if is_asset:
        return (amt, None) if movement > 0 else (None, amt)
    else:  # Liability, Equity
        return (None, amt) if movement > 0 else (amt, None)


def _is_debit_credit(qb_type: str, usd_amount: float) -> tuple[float | None, float | None]:
    """
    For an IS account, determine debit / credit.

    Revenue accounts have a credit normal balance.
    Expense accounts have a debit normal balance.
    """
    if abs(usd_amount) < 0.005:
        return None, None
    amt = abs(usd_amount)
    is_revenue = any(
        kw in qb_type.lower()
        for kw in ("income", "revenue", "grant revenue", "interest")
    )
    if is_revenue:
        return (None, amt) if usd_amount > 0 else (amt, None)
    else:  # Expense
        return (amt, None) if usd_amount > 0 else (None, amt)


# ---------------------------------------------------------------------------
# Line builder helpers
# ---------------------------------------------------------------------------

def _line(
    section: str,
    period_date: date,
    je_number: str,
    memo: str,
    qb_number: str,
    qb_account: str,
    debit: float | None,
    credit: float | None,
    xero_account: str = "",
    aud_amount: float | None = None,
    rate: float | None = None,
    usd_amount: float | None = None,
    is_mapped: bool = True,
) -> dict[str, Any]:
    return {
        "section": section,
        "date": period_date.strftime("%m/%d/%Y"),
        "type": "Journal Entry",
        "num": je_number,
        "name": "",
        "memo": memo,
        "account_number": qb_number,
        "account_name": qb_account,
        "debit": round(debit, 2) if debit is not None else None,
        "credit": round(credit, 2) if credit is not None else None,
        "xero_account": xero_account,
        "aud_amount": aud_amount,
        "rate": rate,
        "usd_amount": usd_amount,
        "is_mapped": is_mapped,
    }


# ---------------------------------------------------------------------------
# Main builder
# ---------------------------------------------------------------------------

BS_MEMO = "Adjusting JE to reflect Australia BS Account Balances at month-end"
IS_MEMO = "Adjusting JE to reflect Australia IS activity"
NI_MEMO = "Australia P&L-Net Income / (Loss) for the month"
EX_MEMO = "Australia Currency Exchange Gains/(Losses)"

EQUITY_IN_SUB = "Equity in Subsidiary"
EXCH_ACCT     = "Currency Exchange Gain (Loss)"


def build_je(
    translated_bs: pd.DataFrame,
    translated_is: pd.DataFrame,
    period_date: date,
    je_number: str,
    equity_sub_number: str = "3300",
    exch_acct_number: str  = "4800",
) -> list[dict]:
    """
    Build the full list of JE line dicts.

    Parameters
    ----------
    translated_bs       Output of translation.translate_balance_sheet()
    translated_is       Output of translation.translate_income_statement()
    period_date         The last day of the reporting month
    je_number           JE reference string, e.g. 'AUS-2026-01'
    equity_sub_number   QB account number for Equity in Subsidiary
    exch_acct_number    QB account number for Currency Exchange Gain (Loss)
    """
    lines: list[dict] = []

    # -----------------------------------------------------------------------
    # 1. BS section – accounts whose je_section == 'BS'
    # -----------------------------------------------------------------------
    bs_rows = translated_bs[translated_bs["je_section"] == "BS"]
    for _, r in bs_rows.iterrows():
        d, c = _bs_debit_credit(r["xero_type"], r["movement_usd"])
        lines.append(_line(
            section="BS",
            period_date=period_date,
            je_number=je_number,
            memo=BS_MEMO,
            qb_number=r["qb_number"],
            qb_account=r["qb_account"],
            debit=d,
            credit=c,
            xero_account=r["xero_account"],
            aud_amount=r["current_aud"] - r["prior_aud"],
            rate=r["current_rate"],
            usd_amount=r["movement_usd"],
            is_mapped=r["is_mapped"],
        ))

    # -----------------------------------------------------------------------
    # 2a. IS section – equity BS accounts (je_section == 'IS' in BS table)
    # -----------------------------------------------------------------------
    eq_rows = translated_bs[translated_bs["je_section"] == "IS"]
    for _, r in eq_rows.iterrows():
        d, c = _bs_debit_credit(r["xero_type"], r["movement_usd"])
        lines.append(_line(
            section="IS",
            period_date=period_date,
            je_number=je_number,
            memo=IS_MEMO,
            qb_number=r["qb_number"],
            qb_account=r["qb_account"],
            debit=d,
            credit=c,
            xero_account=r["xero_account"],
            aud_amount=r["current_aud"] - r["prior_aud"],
            rate=r["current_rate"],
            usd_amount=r["movement_usd"],
            is_mapped=r["is_mapped"],
        ))

    # -----------------------------------------------------------------------
    # 2b. IS section – income / expense accounts
    # -----------------------------------------------------------------------
    for _, r in translated_is.iterrows():
        d, c = _is_debit_credit(r["qb_type"], r["usd_amount"])
        lines.append(_line(
            section="IS",
            period_date=period_date,
            je_number=je_number,
            memo=IS_MEMO,
            qb_number=r["qb_number"],
            qb_account=r["qb_account"],
            debit=d,
            credit=c,
            xero_account=r["xero_account"],
            aud_amount=r["aud_amount"],
            rate=r["avg_rate"],
            usd_amount=r["usd_amount"],
            is_mapped=r["is_mapped"],
        ))

    # -----------------------------------------------------------------------
    # 3. Net income plug – offset the IS section into Equity in Subsidiary
    # -----------------------------------------------------------------------
    is_debits  = sum(l["debit"]  or 0 for l in lines if l["section"] == "IS")
    is_credits = sum(l["credit"] or 0 for l in lines if l["section"] == "IS")
    net_income = is_credits - is_debits   # positive = net income

    if abs(net_income) >= 0.005:
        ni_d = None if net_income > 0 else abs(net_income)
        ni_c = abs(net_income) if net_income > 0 else None
        lines.append(_line(
            section="IS",
            period_date=period_date,
            je_number=je_number,
            memo=NI_MEMO,
            qb_number=equity_sub_number,
            qb_account=EQUITY_IN_SUB,
            debit=ni_d,
            credit=ni_c,
            xero_account="(Net Income / Loss)",
            aud_amount=None,
            rate=None,
            usd_amount=net_income,
        ))

    # -----------------------------------------------------------------------
    # 4. Exchange adjustment – balancing plug
    # -----------------------------------------------------------------------
    total_d = sum(l["debit"]  or 0 for l in lines)
    total_c = sum(l["credit"] or 0 for l in lines)
    adj = round(total_c - total_d, 2)   # positive → need more debits

    if abs(adj) >= 0.005:
        if adj > 0:
            # Debit Equity in Sub, Credit Exchange Gain (Loss)
            eq_d, eq_c   = adj, None
            ex_d, ex_c   = None, adj
        else:
            # Credit Equity in Sub, Debit Exchange Gain (Loss)
            eq_d, eq_c   = None, abs(adj)
            ex_d, ex_c   = abs(adj), None

        lines.append(_line(
            section="EXCH",
            period_date=period_date,
            je_number=je_number,
            memo=EX_MEMO,
            qb_number=equity_sub_number,
            qb_account=EQUITY_IN_SUB,
            debit=eq_d,
            credit=eq_c,
            xero_account="(Exchange Adjustment)",
            usd_amount=adj,
        ))
        lines.append(_line(
            section="EXCH",
            period_date=period_date,
            je_number=je_number,
            memo=EX_MEMO,
            qb_number=exch_acct_number,
            qb_account=EXCH_ACCT,
            debit=ex_d,
            credit=ex_c,
            xero_account="(Exchange Adjustment)",
            usd_amount=-adj,
        ))

    return lines


# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------

def to_qbo_csv(je_lines: list[dict]) -> str:
    """
    Format JE lines as a QuickBooks Online-importable CSV.

    QBO JE import required columns:
        *JournalDate | *Memo | *AccountName | AccountNumber | *Debits | *Credits
        | Description | Name
    """
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow([
        "*JournalDate", "*Memo", "*AccountName", "AccountNumber",
        "*Debits", "*Credits", "Description", "Name",
    ])
    for ln in je_lines:
        writer.writerow([
            ln["date"],
            ln["memo"],
            ln["account_name"],
            ln.get("account_number", ""),
            f"{ln['debit']:.2f}"  if ln["debit"]  is not None else "",
            f"{ln['credit']:.2f}" if ln["credit"] is not None else "",
            ln.get("xero_account", ""),
            ln.get("name", ""),
        ])
    return buf.getvalue()


def je_summary(je_lines: list[dict]) -> dict:
    """Return total debits, total credits, and balance status."""
    total_d = round(sum(l["debit"]  or 0 for l in je_lines), 2)
    total_c = round(sum(l["credit"] or 0 for l in je_lines), 2)
    return {
        "total_debits":  total_d,
        "total_credits": total_c,
        "balanced":      abs(total_d - total_c) < 0.02,
        "difference":    round(total_d - total_c, 2),
    }
