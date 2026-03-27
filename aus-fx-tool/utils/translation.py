"""
Currency translation logic (ASC 830 simplified).

Balance Sheet  → translated at the period-end rate (current or prior).
Income Statement → translated at the average rate
                   (average of current and prior period-end rates).

The exchange adjustment is the plug that makes the combined JE balance.
It is posted to the Income Statement (Currency Exchange Gain / Loss).
"""
import pandas as pd
import numpy as np
from datetime import datetime


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _month_label_variants(year: int, month: int) -> list[str]:
    """
    Return several string formats for a month-end date so we can match
    against Xero's column headers regardless of exact formatting.
    e.g. for January 2026: ['Jan 2026', 'January 2026', '31 Jan 2026',
                             '31 January 2026', '1/31/2026', '2026-01-31']
    """
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    dt = datetime(year, month, last_day)
    return [
        dt.strftime("%b %Y"),          # Jan 2026
        dt.strftime("%B %Y"),          # January 2026
        dt.strftime("%-d %b %Y"),      # 31 Jan 2026  (Linux)
        dt.strftime("%#d %b %Y"),      # 31 Jan 2026  (Windows)
        dt.strftime("%-d %B %Y"),      # 31 January 2026 (Linux)
        dt.strftime("%#d %B %Y"),      # 31 January 2026 (Windows)
        dt.strftime("%d %b %Y"),       # 31 Jan 2026 (zero-padded)
        dt.strftime("%d %B %Y"),       # 31 January 2026 (zero-padded)
        dt.strftime("%m/%d/%Y"),       # 01/31/2026
        dt.strftime("%-m/%-d/%Y"),     # 1/31/2026
        dt.strftime("%Y-%m-%d"),       # 2026-01-31
    ]


def _find_col(df: pd.DataFrame, year: int, month: int,
              exclude: list[str] | None = None) -> str | None:
    """
    Search df.columns for a column whose name matches any label variant
    for the given year/month.  Returns the first match or None.
    """
    exclude = exclude or []
    variants = {v.lower() for v in _month_label_variants(year, month)}
    for col in df.columns:
        if col in exclude:
            continue
        if col.lower().strip() in variants:
            return col
    return None


def find_bs_columns(bs_df: pd.DataFrame, year: int, month: int) -> tuple[str | None, str | None]:
    """
    Locate current-period and prior-period columns in a parsed Balance Sheet df.
    Returns (current_col, prior_col).
    """
    current_col = _find_col(bs_df, year, month)

    # Prior period = one month earlier
    if month == 1:
        py, pm = year - 1, 12
    else:
        py, pm = year, month - 1
    prior_col = _find_col(bs_df, py, pm)

    return current_col, prior_col


def find_pl_column(pl_df: pd.DataFrame, year: int, month: int) -> str | None:
    """
    Locate the current-period column in a parsed P&L df.
    """
    return _find_col(pl_df, year, month)


# ---------------------------------------------------------------------------
# Translation
# ---------------------------------------------------------------------------

def translate_balance_sheet(
    bs_df: pd.DataFrame,
    mapping_df: pd.DataFrame,
    current_col: str,
    prior_col: str,
    current_rate: float,
    prior_rate: float,
) -> pd.DataFrame:
    """
    Translate each BS account from AUD to USD and compute the USD movement.

    Returns a DataFrame with one row per account, columns:
        xero_account, xero_type, qb_account, qb_type, qb_number, je_section,
        current_aud, prior_aud,
        current_rate, prior_rate,
        current_usd, prior_usd,
        movement_usd,
        is_mapped
    """
    # Build lookup: xero_account → mapping row
    mmap = mapping_df.set_index("xero_account").to_dict("index")

    rows = []
    for _, r in bs_df.iterrows():
        account = r["account"]
        xero_type = r.get("xero_type", "Unknown")

        current_aud = float(r.get(current_col, 0) or 0)
        prior_aud = float(r.get(prior_col, 0) or 0)

        current_usd = current_aud * current_rate
        prior_usd = prior_aud * prior_rate
        movement_usd = current_usd - prior_usd

        mapped = mmap.get(account, {})
        rows.append({
            "xero_account": account,
            "xero_type": xero_type,
            "qb_account": mapped.get("qb_account", "*** UNMAPPED ***"),
            "qb_type": mapped.get("qb_type", ""),
            "qb_number": str(mapped.get("qb_number", "")),
            "je_section": mapped.get("je_section", "BS"),
            "current_aud": current_aud,
            "prior_aud": prior_aud,
            "current_rate": current_rate,
            "prior_rate": prior_rate,
            "current_usd": current_usd,
            "prior_usd": prior_usd,
            "movement_usd": movement_usd,
            "is_mapped": account in mmap,
        })

    return pd.DataFrame(rows)


def translate_income_statement(
    pl_df: pd.DataFrame,
    mapping_df: pd.DataFrame,
    period_col: str,
    avg_rate: float,
) -> pd.DataFrame:
    """
    Translate each P&L account for the period from AUD to USD.

    Returns a DataFrame with one row per account, columns:
        xero_account, qb_account, qb_type, qb_number, je_section,
        aud_amount, avg_rate, usd_amount, is_mapped
    """
    mmap = mapping_df.set_index("xero_account").to_dict("index")

    rows = []
    for _, r in pl_df.iterrows():
        account = r["account"]
        aud_amount = float(r.get(period_col, 0) or 0)
        usd_amount = aud_amount * avg_rate

        mapped = mmap.get(account, {})
        rows.append({
            "xero_account": account,
            "qb_account": mapped.get("qb_account", "*** UNMAPPED ***"),
            "qb_type": mapped.get("qb_type", ""),
            "qb_number": str(mapped.get("qb_number", "")),
            "je_section": mapped.get("je_section", "IS"),
            "aud_amount": aud_amount,
            "avg_rate": avg_rate,
            "usd_amount": usd_amount,
            "is_mapped": account in mmap,
        })

    return pd.DataFrame(rows)
