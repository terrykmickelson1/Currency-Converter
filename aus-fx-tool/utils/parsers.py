"""
Parsers for Xero Excel report exports.
Handles both Profit & Loss and Balance Sheet in a single or separate file(s).
"""
import pandas as pd
import numpy as np
from datetime import datetime

# Row labels to skip when parsing (section headers, subtotals, calculated lines)
_PL_SKIP = {
    'Trading Income', 'Cost of Sales', 'Operating Expenses',
    'Gross Profit', 'Net Profit', 'Net Loss',
}
_BS_SECTION_A = {'Assets', 'Liabilities', 'Equity'}
_BS_SECTION_B_SKIP = {
    'Bank', 'Current Assets', 'Fixed Assets', 'Non-current Assets',
    'Current Liabilities', 'Non-current Liabilities', 'Long Term Liabilities',
    'Net Assets',
}
_BS_TYPE_MAP = {'Assets': 'Asset', 'Liabilities': 'Liability', 'Equity': 'Equity'}


def _to_float(val) -> float:
    """Convert a cell value (possibly a parenthetical string) to float."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(',', '')
    if s.startswith('(') and s.endswith(')'):
        try:
            return -float(s[1:-1])
        except ValueError:
            return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _find_sheet(xl: dict, keywords: list[str]) -> str | None:
    """Return the first sheet name whose name contains any of the keywords (case-insensitive)."""
    for name in xl:
        lower = name.lower()
        if any(kw in lower for kw in keywords):
            return name
    return None


def parse_profit_loss(file) -> pd.DataFrame:
    """
    Parse a Xero Profit & Loss Excel export.

    Returns a DataFrame with columns:
        account   – Xero account name (str)
        <month>   – one column per period present in the export (float AUD)

    Strips all section headers, subtotals, and blank rows.
    """
    xl = pd.read_excel(file, sheet_name=None, header=None)
    sheet = _find_sheet(xl, ['profit', 'loss', 'p&l', 'income statement'])
    if sheet is None:
        sheet = list(xl.keys())[0]

    raw = pd.read_excel(file, sheet_name=sheet, header=None)

    # Locate header row – the row whose first cell is exactly 'Account'
    header_idx = None
    for i, row in raw.iterrows():
        if str(row.iloc[0]).strip() == 'Account':
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(
            "Cannot find the 'Account' header row in the P&L sheet. "
            "Ensure the file is a Xero Profit & Loss export."
        )

    col_headers = [str(v).strip() for v in raw.iloc[header_idx].tolist()]
    data = raw.iloc[header_idx + 1:].reset_index(drop=True)

    rows = []
    for i in range(len(data)):
        account = str(data.iloc[i, 0]).strip()
        if not account or account in ('nan', 'NaN'):
            continue
        if account in _PL_SKIP:
            continue
        if account.startswith('Total '):
            continue

        row_dict = {'account': account}
        for j in range(1, len(col_headers)):
            col = col_headers[j]
            if not col or col in ('nan', 'NaN', 'Year to date', 'Year To Date', 'YTD'):
                continue
            row_dict[col] = _to_float(data.iloc[i, j])
        rows.append(row_dict)

    if not rows:
        raise ValueError("No account data found in P&L sheet after parsing.")

    return pd.DataFrame(rows).fillna(0.0)


def parse_balance_sheet(file) -> pd.DataFrame:
    """
    Parse a Xero Balance Sheet Excel export.

    Returns a DataFrame with columns:
        account    – Xero account name (str)
        xero_type  – 'Asset' | 'Liability' | 'Equity'
        <date>     – one column per period present (float AUD), e.g. '31 Jan 2026'

    Strips all section headers, subtotals, and blank rows.
    """
    xl = pd.read_excel(file, sheet_name=None, header=None)
    sheet = _find_sheet(xl, ['balance sheet', 'balance', 'bs'])
    if sheet is None:
        sheets = list(xl.keys())
        sheet = sheets[1] if len(sheets) > 1 else sheets[0]

    raw = pd.read_excel(file, sheet_name=sheet, header=None)

    # Header row: a cell whose value is exactly 'Account'
    header_idx = None
    account_col_idx = None
    for i, row in raw.iterrows():
        for j, val in enumerate(row):
            if str(val).strip() == 'Account':
                header_idx = i
                account_col_idx = j
                break
        if header_idx is not None:
            break
    if header_idx is None:
        raise ValueError(
            "Cannot find the 'Account' header row in the Balance Sheet. "
            "Ensure the file is a Xero Balance Sheet export."
        )

    # Date headers start in the column after 'Account'
    date_start_col = account_col_idx + 1
    date_headers = [
        str(raw.iloc[header_idx, c]).strip()
        for c in range(date_start_col, raw.shape[1])
    ]

    rows = []
    current_type = None

    for i in range(header_idx + 1, len(raw)):
        # Column A may contain section-level labels (Assets / Liabilities / Equity / Total …)
        col_a = str(raw.iloc[i, 0]).strip() if not pd.isna(raw.iloc[i, 0]) else ''
        # Column at account_col_idx is the account name
        col_b = str(raw.iloc[i, account_col_idx]).strip() \
            if not pd.isna(raw.iloc[i, account_col_idx]) else ''

        # Update section type
        if col_a in _BS_SECTION_A:
            current_type = _BS_TYPE_MAP[col_a]
            continue
        if col_a.startswith('Total') or col_a == 'Net Assets':
            continue
        if not col_b or col_b in ('nan', 'NaN'):
            continue
        if col_b in _BS_SECTION_B_SKIP:
            continue
        if col_b.startswith('Total') or col_b == 'Net Assets':
            continue

        row_dict = {'account': col_b, 'xero_type': current_type or 'Unknown'}
        for k, date_str in enumerate(date_headers):
            if not date_str or date_str in ('nan', 'NaN'):
                continue
            col_idx = date_start_col + k
            if col_idx < raw.shape[1]:
                row_dict[date_str] = _to_float(raw.iloc[i, col_idx])
        rows.append(row_dict)

    if not rows:
        raise ValueError("No account data found in Balance Sheet after parsing.")

    return pd.DataFrame(rows).fillna(0.0)
