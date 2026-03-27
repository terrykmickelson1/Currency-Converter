# AUS → USD FX Conversion Tool

Converts the Australian subsidiary's Xero financials (AUD) to USD using ASC 830,
then produces a QuickBooks Online Journal Entry CSV and an audit-ready Excel workpaper.

## Deployment (Streamlit Community Cloud)

1. Create a free account at https://github.com and push this folder as a repository.
2. Create a free account at https://share.streamlit.io.
3. Click **New app** → connect your GitHub repo → set main file to `app.py`.
4. The app will be live at `https://your-app-name.streamlit.app`.

## Local development (optional)

Requires Python 3.10+.

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Monthly workflow

1. **Tab 1** – Select period, click "Fetch Rates", confirm or override.
2. **Tab 2** – Upload Xero Balance Sheet + P&L Excel exports.
3. **Tab 3** – Verify mapping; add any new accounts.
4. **Tab 4** – Click "Generate JE", review balance check, download outputs.

## Outputs

| File | Purpose |
|---|---|
| `JE_YYYY-MM_QBO.csv` | QuickBooks Online import CSV |
| `Workpaper_YYYY-MM_AUS_FX.xlsx` | 5-tab audit workpaper |
| `JE_YYYY-MM_Detail.csv` | Full detail with AUD amounts and rates |
