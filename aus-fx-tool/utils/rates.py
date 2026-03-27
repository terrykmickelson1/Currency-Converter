"""
Exchange rate fetching via the Frankfurter API (free, no key required).
https://www.frankfurter.app/

Fetches AUD/USD spot rates for any historical date.
Falls back to the nearest prior business day if the exact date is unavailable
(e.g. weekend, public holiday).
"""
import requests
from datetime import datetime, timedelta
from calendar import monthrange

_BASE = "https://api.frankfurter.app"
_TIMEOUT = 12  # seconds


def period_end_date(year: int, month: int) -> str:
    """Return the last calendar day of the given month as YYYY-MM-DD."""
    last = monthrange(year, month)[1]
    return f"{year}-{month:02d}-{last:02d}"


def fetch_aud_usd(date_str: str) -> tuple[float, str, str]:
    """
    Fetch the AUD → USD rate for *date_str* (YYYY-MM-DD).

    Returns
    -------
    rate        : float   – units of USD per 1 AUD
    source_url  : str     – the URL that returned data (for workpaper citation)
    actual_date : str     – the date the rate is sourced from (may differ if
                            the requested date was a non-trading day)

    Raises
    ------
    ConnectionError  if the API cannot be reached
    ValueError       if no rate can be obtained within 7 prior days
    """
    for delta in range(8):          # try requested date, then up to 7 days back
        try_date = (
            datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=delta)
        ).strftime("%Y-%m-%d")
        url = f"{_BASE}/{try_date}?from=AUD&to=USD"
        try:
            resp = requests.get(url, timeout=_TIMEOUT)
        except requests.RequestException as exc:
            raise ConnectionError(
                f"Could not reach the Frankfurter exchange rate API: {exc}"
            ) from exc

        if resp.status_code == 200:
            data = resp.json()
            return float(data["rates"]["USD"]), url, data["date"]
        if resp.status_code == 404:
            continue   # non-trading day → try previous day
        resp.raise_for_status()

    raise ValueError(
        f"No AUD/USD rate available within 7 days before {date_str}. "
        "Please enter the rate manually."
    )


def fetch_both_rates(
    current_year: int, current_month: int
) -> dict:
    """
    Convenience wrapper: fetches both the current and prior period end rates
    and returns a summary dict.

    Returns a dict with keys:
        current_date, current_rate, current_url, current_actual_date
        prior_date, prior_rate, prior_url, prior_actual_date
        avg_rate
    """
    # Current period
    curr_date_str = period_end_date(current_year, current_month)
    curr_rate, curr_url, curr_actual = fetch_aud_usd(curr_date_str)

    # Prior period (one month back)
    if current_month == 1:
        prior_year, prior_month = current_year - 1, 12
    else:
        prior_year, prior_month = current_year, current_month - 1
    prior_date_str = period_end_date(prior_year, prior_month)
    prior_rate, prior_url, prior_actual = fetch_aud_usd(prior_date_str)

    avg_rate = (curr_rate + prior_rate) / 2

    return {
        "current_date": curr_date_str,
        "current_rate": curr_rate,
        "current_url": curr_url,
        "current_actual_date": curr_actual,
        "prior_date": prior_date_str,
        "prior_rate": prior_rate,
        "prior_url": prior_url,
        "prior_actual_date": prior_actual,
        "avg_rate": avg_rate,
    }
