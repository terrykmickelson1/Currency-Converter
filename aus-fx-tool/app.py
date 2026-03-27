"""
AUS → USD FX Conversion Tool
Converts Xero (AUD) financial statements to USD and generates a
QuickBooks Online Journal Entry + audit workpaper.
"""
import os
from calendar import month_name
from datetime import date, datetime

import pandas as pd
import streamlit as st

from utils.je_builder import build_je, je_summary, to_qbo_csv
from utils.parsers import parse_balance_sheet, parse_profit_loss
from utils.rates import fetch_both_rates, period_end_date
from utils.translation import (
    find_bs_columns,
    find_pl_column,
    translate_balance_sheet,
    translate_income_statement,
)
from utils.workpaper import generate_workpaper

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="AUS→USD FX Converter",
    page_icon="💱",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.title("🇦🇺 → 🇺🇸  Australia Subsidiary FX Conversion Tool")
st.caption(
    "Converts Xero (AUD) financials to USD using ASC 830 | "
    "Generates QuickBooks Journal Entry + Audit Workpaper"
)

# ---------------------------------------------------------------------------
# Mapping – load once into session state
# ---------------------------------------------------------------------------
MAPPING_PATH = os.path.join(os.path.dirname(__file__), "data", "mapping.csv")

def _load_mapping() -> pd.DataFrame:
    if os.path.exists(MAPPING_PATH):
        df = pd.read_csv(MAPPING_PATH, dtype=str).fillna("")
    else:
        df = pd.DataFrame(columns=[
            "xero_account", "xero_type",
            "qb_account", "qb_type", "qb_number", "je_section",
        ])
    return df

def _save_mapping(df: pd.DataFrame):
    df.to_csv(MAPPING_PATH, index=False)

if "mapping_df" not in st.session_state:
    st.session_state.mapping_df = _load_mapping()

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "1 · Period & Rates",
    "2 · Upload Statements",
    "3 · Account Mapping",
    "4 · Generate JE",
])

# ============================================================
# TAB 1 – Period & Rates
# ============================================================
with tab1:
    st.subheader("Reporting Period & Exchange Rates")
    st.markdown(
        "Select the period you are closing. Rates are fetched automatically from "
        "[Frankfurter](https://www.frankfurter.app/) (European Central Bank data). "
        "You may override any rate before proceeding."
    )

    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        sel_year  = st.selectbox("Year",  range(datetime.today().year, 2018, -1),
                                 index=0, key="sel_year")
    with c2:
        sel_month = st.selectbox("Month", range(1, 13), index=datetime.today().month - 1,
                                 format_func=lambda m: month_name[m], key="sel_month")
    with c3:
        entity_name = st.text_input("Entity / Subsidiary Name (for workpaper header)",
                                    value="Australian Subsidiary", key="entity_name")

    je_number = st.text_input(
        "Journal Entry Number  (e.g. AUS-2026-01)",
        value=f"AUS-{sel_year}-{sel_month:02d}",
        key="je_number",
    )

    fetch_col, _ = st.columns([2, 5])
    with fetch_col:
        fetch_btn = st.button("🔄  Fetch Rates from API", type="primary")

    if fetch_btn:
        with st.spinner("Fetching rates from Frankfurter API…"):
            try:
                rates = fetch_both_rates(sel_year, sel_month)
                st.session_state.rates = rates
                st.success("Rates fetched successfully.")
            except Exception as e:
                st.error(f"Could not fetch rates: {e}")
                st.info("Enter rates manually below.")

    # Show rate inputs (pre-filled from API or blank)
    stored = st.session_state.get("rates", {})
    st.markdown("---")
    st.markdown("**Confirm or manually enter rates** (AUD per 1 USD → shown as USD per 1 AUD):")

    rc1, rc2, rc3 = st.columns(3)
    curr_end = period_end_date(sel_year, sel_month)
    if sel_month == 1:
        prior_end = period_end_date(sel_year - 1, 12)
    else:
        prior_end = period_end_date(sel_year, sel_month - 1)

    with rc1:
        current_rate = st.number_input(
            f"Current period-end rate  ({curr_end})",
            min_value=0.0001, max_value=10.0,
            value=float(stored.get("current_rate", 0.6200)),
            format="%.6f", step=0.0001, key="current_rate",
        )
    with rc2:
        prior_rate = st.number_input(
            f"Prior period-end rate  ({prior_end})",
            min_value=0.0001, max_value=10.0,
            value=float(stored.get("prior_rate", 0.6200)),
            format="%.6f", step=0.0001, key="prior_rate",
        )
    with rc3:
        avg_rate = (current_rate + prior_rate) / 2
        st.metric("Average Rate (IS)", f"{avg_rate:.6f}")
        st.caption("Auto-calculated: (Current + Prior) ÷ 2")

    # Persist whatever is on screen into session state
    st.session_state.rates = {
        **stored,
        "current_rate":        current_rate,
        "prior_rate":          prior_rate,
        "avg_rate":            avg_rate,
        "current_date":        curr_end,
        "prior_date":          prior_end,
        "current_actual_date": stored.get("current_actual_date", curr_end),
        "prior_actual_date":   stored.get("prior_actual_date",   prior_end),
        "current_url":         stored.get("current_url",  "https://www.frankfurter.app (manual override)"),
        "prior_url":           stored.get("prior_url",    "https://www.frankfurter.app (manual override)"),
    }

    if stored.get("current_rate"):
        st.info(
            f"Current rate sourced from: {stored.get('current_url', '—')}  "
            f"(date returned: {stored.get('current_actual_date', '—')})\n\n"
            f"Prior rate sourced from: {stored.get('prior_url', '—')}  "
            f"(date returned: {stored.get('prior_actual_date', '—')})"
        )


# ============================================================
# TAB 2 – Upload Statements
# ============================================================
with tab2:
    st.subheader("Upload Xero Financial Statements")
    st.markdown(
        "Upload the Xero **Balance Sheet** and **Profit & Loss** Excel exports. "
        "These may be in a single workbook (multiple sheets) or separate files."
    )

    upload_mode = st.radio(
        "File format",
        ["Single workbook (BS + P&L in one file)", "Separate files (one each)"],
        horizontal=True,
    )

    bs_df = None
    pl_df = None

    if upload_mode == "Single workbook (BS + P&L in one file)":
        combined_file = st.file_uploader(
            "Upload combined Xero report (must contain 'Balance Sheet' and 'Profit & Loss' sheets)",
            type=["xlsx", "xls"],
            key="combined_upload",
        )
        if combined_file:
            with st.spinner("Parsing…"):
                try:
                    bs_df = parse_balance_sheet(combined_file)
                    combined_file.seek(0)
                    pl_df = parse_profit_loss(combined_file)
                    st.success(
                        f"Parsed **{len(bs_df)} BS accounts** and "
                        f"**{len(pl_df)} P&L accounts**."
                    )
                except Exception as e:
                    st.error(f"Parse error: {e}")
    else:
        f1, f2 = st.columns(2)
        with f1:
            bs_file = st.file_uploader("Balance Sheet export", type=["xlsx", "xls"], key="bs_upload")
        with f2:
            pl_file = st.file_uploader("Profit & Loss export", type=["xlsx", "xls"], key="pl_upload")

        if bs_file:
            try:
                bs_df = parse_balance_sheet(bs_file)
                st.success(f"Parsed **{len(bs_df)} BS accounts**.")
            except Exception as e:
                st.error(f"Balance Sheet parse error: {e}")
        if pl_file:
            try:
                pl_df = parse_profit_loss(pl_file)
                st.success(f"Parsed **{len(pl_df)} P&L accounts**.")
            except Exception as e:
                st.error(f"Profit & Loss parse error: {e}")

    if bs_df is not None:
        st.session_state.bs_df = bs_df
    if pl_df is not None:
        st.session_state.pl_df = pl_df

    # Preview
    if "bs_df" in st.session_state:
        with st.expander("Preview: Balance Sheet accounts", expanded=False):
            st.dataframe(st.session_state.bs_df, use_container_width=True)
    if "pl_df" in st.session_state:
        with st.expander("Preview: Profit & Loss accounts", expanded=False):
            st.dataframe(st.session_state.pl_df, use_container_width=True)

    # Detect unmapped accounts from uploaded files
    if "bs_df" in st.session_state and "pl_df" in st.session_state:
        all_xero = set(st.session_state.bs_df["account"]) | set(st.session_state.pl_df["account"])
        mapped = set(st.session_state.mapping_df["xero_account"])
        unmapped = all_xero - mapped
        if unmapped:
            st.warning(
                f"⚠️  **{len(unmapped)} account(s) not yet in your mapping table.** "
                "Go to Tab 3 to add them before generating the JE."
            )
            st.write(sorted(unmapped))
        else:
            st.success("✅  All accounts are mapped.")


# ============================================================
# TAB 3 – Account Mapping
# ============================================================
with tab3:
    st.subheader("Account Mapping  ·  Xero → QuickBooks")
    st.markdown(
        "This table defines how each Xero account maps to a QuickBooks account. "
        "Edit cells directly, then click **Save Mapping** to persist changes. "
        "You can also **Download** the mapping as a backup and **Upload** a previously saved version."
    )

    # Upload saved mapping
    uploaded_map = st.file_uploader(
        "Upload a previously saved mapping CSV (optional)",
        type=["csv"], key="mapping_upload",
    )
    if uploaded_map:
        try:
            new_map = pd.read_csv(uploaded_map, dtype=str).fillna("")
            st.session_state.mapping_df = new_map
            _save_mapping(new_map)
            st.success("Mapping loaded and saved.")
        except Exception as e:
            st.error(f"Could not read mapping CSV: {e}")

    st.markdown("---")

    # Add new row helper
    with st.expander("➕  Add a new account mapping", expanded=False):
        na1, na2 = st.columns(2)
        with na1:
            new_xero   = st.text_input("Xero Account Name (exact)", key="new_xero")
            new_xtype  = st.selectbox("Xero Type",
                                      ["Asset", "Liabilities", "Equity", "Income", "Expense"],
                                      key="new_xtype")
            new_je     = st.selectbox("JE Section", ["BS", "IS"], key="new_je")
        with na2:
            new_qb     = st.text_input("QuickBooks Account Name", key="new_qb")
            new_qtype  = st.selectbox("QB Type",
                                      ["Current Asset", "Other Current Liabilities",
                                       "Equity", "Income", "Expense"],
                                      key="new_qtype")
            new_qnum   = st.text_input("QB Account Number", key="new_qnum")

        if st.button("Add Row", key="add_row_btn"):
            if new_xero and new_qb:
                new_row = pd.DataFrame([{
                    "xero_account": new_xero,
                    "xero_type":    new_xtype,
                    "qb_account":   new_qb,
                    "qb_type":      new_qtype,
                    "qb_number":    new_qnum,
                    "je_section":   new_je,
                }])
                st.session_state.mapping_df = pd.concat(
                    [st.session_state.mapping_df, new_row], ignore_index=True
                )
                _save_mapping(st.session_state.mapping_df)
                st.success(f"Added mapping: {new_xero} → {new_qb}")
            else:
                st.warning("Please enter both a Xero account name and a QuickBooks account name.")

    # Editable table
    st.markdown("**Edit the mapping table below:**")
    edited_df = st.data_editor(
        st.session_state.mapping_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "xero_account": st.column_config.TextColumn("Xero Account Name", width="large"),
            "xero_type":    st.column_config.SelectboxColumn("Xero Type",
                            options=["Asset", "Liabilities", "Equity", "Income", "Expense"]),
            "qb_account":   st.column_config.TextColumn("QB Account Name", width="large"),
            "qb_type":      st.column_config.SelectboxColumn("QB Type",
                            options=["Current Asset", "Other Current Liabilities",
                                     "Equity", "Income", "Expense"]),
            "qb_number":    st.column_config.TextColumn("QB Acct #", width="small"),
            "je_section":   st.column_config.SelectboxColumn("JE Section",
                            options=["BS", "IS"]),
        },
        key="mapping_editor",
    )

    mc1, mc2, mc3 = st.columns([2, 2, 4])
    with mc1:
        if st.button("💾  Save Mapping", type="primary", key="save_map_btn"):
            st.session_state.mapping_df = edited_df
            _save_mapping(edited_df)
            st.success("Mapping saved.")
    with mc2:
        map_csv = edited_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️  Download Mapping CSV",
            data=map_csv,
            file_name="account_mapping.csv",
            mime="text/csv",
            key="dl_mapping",
        )

    # Summary stats
    st.markdown("---")
    n_total  = len(edited_df)
    n_multi  = edited_df.groupby("xb_account" if "xb_account" in edited_df.columns
                                 else "qb_account")["xero_account"].count()
    n_combined = (n_multi > 1).sum()
    st.caption(
        f"**{n_total}** Xero accounts mapped  ·  "
        f"**{n_combined}** QuickBooks accounts receive multiple Xero accounts (combined)"
    )


# ============================================================
# TAB 4 – Generate JE
# ============================================================
with tab4:
    st.subheader("Generate Journal Entry & Workpaper")

    # Pre-flight checks
    issues = []
    if "bs_df"  not in st.session_state: issues.append("Balance Sheet not uploaded (Tab 2)")
    if "pl_df"  not in st.session_state: issues.append("Profit & Loss not uploaded (Tab 2)")
    if "rates"  not in st.session_state or not st.session_state.rates.get("current_rate"):
        issues.append("Exchange rates not set (Tab 1)")

    if issues:
        for i in issues:
            st.warning(f"⚠️  {i}")
        st.stop()

    rates   = st.session_state.rates
    bs_df   = st.session_state.bs_df
    pl_df   = st.session_state.pl_df
    map_df  = st.session_state.mapping_df

    year   = st.session_state.sel_year
    month  = st.session_state.sel_month

    # ------------------------------------------------------------------
    # Find columns in uploaded data
    # ------------------------------------------------------------------
    current_col, prior_col = find_bs_columns(bs_df, year, month)
    pl_col = find_pl_column(pl_df, year, month)

    col_issues = []
    if not current_col:
        col_issues.append(
            f"Could not locate the current-period column in the Balance Sheet "
            f"for {month_name[month]} {year}.  "
            "Available columns: " + ", ".join(
                c for c in bs_df.columns if c not in ("account", "xero_type")
            )
        )
    if not prior_col:
        col_issues.append(
            f"Could not locate the prior-period column in the Balance Sheet."
        )
    if not pl_col:
        col_issues.append(
            f"Could not locate the current-period column in the P&L for "
            f"{month_name[month]} {year}.  "
            "Available columns: " + ", ".join(
                c for c in pl_df.columns if c != "account"
            )
        )
    if col_issues:
        for ci in col_issues:
            st.error(ci)
        st.markdown(
            "**Tip:** The period selected in Tab 1 must match the column dates in your "
            "uploaded Xero exports exactly.  Check that the month/year selection is correct."
        )
        st.stop()

    # ------------------------------------------------------------------
    # Advanced options
    # ------------------------------------------------------------------
    with st.expander("⚙️  Advanced options", expanded=False):
        ac1, ac2 = st.columns(2)
        with ac1:
            equity_sub_num = st.text_input(
                "QB account # for 'Equity in Subsidiary'", value="3300",
                key="equity_sub_num"
            )
        with ac2:
            exch_acct_num = st.text_input(
                "QB account # for 'Currency Exchange Gain (Loss)'", value="4800",
                key="exch_acct_num"
            )

    equity_sub_num = st.session_state.get("equity_sub_num", "3300")
    exch_acct_num  = st.session_state.get("exch_acct_num",  "4800")

    # ------------------------------------------------------------------
    # Generate button
    # ------------------------------------------------------------------
    if st.button("⚡  Generate Journal Entry", type="primary", key="generate_btn"):
        with st.spinner("Translating and building JE…"):
            try:
                # Translate
                t_bs = translate_balance_sheet(
                    bs_df, map_df, current_col, prior_col,
                    rates["current_rate"], rates["prior_rate"],
                )
                t_is = translate_income_statement(
                    pl_df, map_df, pl_col, rates["avg_rate"],
                )

                # Build JE
                period_dt = date(
                    year, month,
                    __import__("calendar").monthrange(year, month)[1]
                )
                je_lines = build_je(
                    t_bs, t_is, period_dt,
                    st.session_state.je_number,
                    equity_sub_number=equity_sub_num,
                    exch_acct_number=exch_acct_num,
                )
                summary = je_summary(je_lines)

                # Build workpaper
                wp_bytes = generate_workpaper(
                    period_date=period_dt,
                    entity_name=st.session_state.entity_name,
                    rates=rates,
                    translated_bs=t_bs,
                    translated_is=t_is,
                    je_lines=je_lines,
                    je_summary_data=summary,
                    current_col_label=current_col,
                    prior_col_label=prior_col,
                    period_label=pl_col,
                )

                st.session_state.je_lines   = je_lines
                st.session_state.summary    = summary
                st.session_state.wp_bytes   = wp_bytes
                st.session_state.t_bs       = t_bs
                st.session_state.t_is       = t_is
                st.session_state.period_dt  = period_dt
                st.success("Journal Entry generated successfully!")

            except Exception as e:
                st.exception(e)

    # ------------------------------------------------------------------
    # Display results (if generated)
    # ------------------------------------------------------------------
    if "je_lines" in st.session_state:
        summary  = st.session_state.summary
        je_lines = st.session_state.je_lines

        # Balance check banner
        if summary["balanced"]:
            st.success(
                f"✅  JE is **BALANCED**  —  "
                f"Debits: **${summary['total_debits']:,.2f}**  |  "
                f"Credits: **${summary['total_credits']:,.2f}**"
            )
        else:
            st.error(
                f"❌  JE is **OUT OF BALANCE** by ${abs(summary['difference']):,.2f}  —  "
                "Review unmapped accounts or contact support."
            )

        # Check for unmapped accounts
        unmapped_lines = [l for l in je_lines if not l.get("is_mapped", True)]
        if unmapped_lines:
            st.warning(
                f"⚠️  **{len(unmapped_lines)} line(s)** have unmapped accounts (shown in red).  "
                "Update the mapping in Tab 3 and regenerate."
            )

        # JE preview table
        with st.expander("📋  Journal Entry Preview", expanded=True):
            je_df = pd.DataFrame([{
                "Section":      l["section"],
                "Account #":    l["account_number"],
                "Account Name": l["account_name"],
                "Memo":         l["memo"],
                "Debit":        l["debit"],
                "Credit":       l["credit"],
                "Xero Account": l["xero_account"],
            } for l in je_lines])

            def _color_unmapped(row):
                if row["Account Name"] == "*** UNMAPPED ***":
                    return ["color: red"] * len(row)
                return [""] * len(row)

            st.dataframe(
                je_df.style.apply(_color_unmapped, axis=1)
                .format({"Debit": "${:,.2f}", "Credit": "${:,.2f}"}, na_rep=""),
                use_container_width=True,
            )

        # Exchange adjustment callout
        exch_lines = [l for l in je_lines if l["section"] == "EXCH"]
        if exch_lines:
            exch_amt = (exch_lines[0]["debit"] or 0) + (exch_lines[0]["credit"] or 0)
            is_gain = exch_lines[1]["credit"] is not None
            st.info(
                f"💱  **Exchange Adjustment: ${exch_amt:,.2f}**  "
                f"({'Gain' if is_gain else 'Loss'})  —  "
                "Posted to Currency Exchange Gain (Loss)"
            )

        # Download buttons
        st.markdown("---")
        st.markdown("### Download Outputs")
        dl1, dl2, dl3 = st.columns(3)

        qbo_csv = to_qbo_csv(je_lines).encode("utf-8")
        period_str = st.session_state.period_dt.strftime("%Y-%m")

        with dl1:
            st.download_button(
                "⬇️  QBO Import CSV",
                data=qbo_csv,
                file_name=f"JE_{period_str}_QBO.csv",
                mime="text/csv",
                key="dl_qbo",
            )
            st.caption("Import directly into QuickBooks Online")

        with dl2:
            st.download_button(
                "⬇️  Audit Workpaper (Excel)",
                data=st.session_state.wp_bytes,
                file_name=f"Workpaper_{period_str}_AUS_FX.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_wp",
            )
            st.caption("5-tab Excel workpaper for audit support")

        with dl3:
            # Full detail CSV for review
            detail_rows = []
            for l in je_lines:
                detail_rows.append({
                    "Date": l["date"],
                    "JE Num": l["num"],
                    "Type": l["type"],
                    "Memo": l["memo"],
                    "Acct #": l["account_number"],
                    "Account Name": l["account_name"],
                    "Debit": f"{l['debit']:.2f}" if l["debit"] is not None else "",
                    "Credit": f"{l['credit']:.2f}" if l["credit"] is not None else "",
                    "Xero Account": l["xero_account"],
                    "AUD Amount": f"{l['aud_amount']:.2f}" if l["aud_amount"] is not None else "",
                    "Rate Used": f"{l['rate']:.6f}" if l["rate"] is not None else "",
                    "USD Amount": f"{l['usd_amount']:.2f}" if l["usd_amount"] is not None else "",
                })
            detail_csv = pd.DataFrame(detail_rows).to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️  Detail Review CSV",
                data=detail_csv,
                file_name=f"JE_{period_str}_Detail.csv",
                mime="text/csv",
                key="dl_detail",
            )
            st.caption("Full detail with AUD amounts and rates")
