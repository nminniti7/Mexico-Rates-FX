"""
Mexico Rates & FX Desk — Cross-SEF Market Comparison Dashboard
=================================================================
Compares Tradition, LatAm SEF, GFI, BGC, and ICAP (tpSEF) daily activity
for MXN (and USD) TIIE swaps and FX products.

WHAT CHANGED IN THIS VERSION (desk feedback)
--------------------------------------------------------------------------
1. Mexico products ONLY — the Non-Mexico and All-Products tabs are gone.
2. Individual-trades drill-down now lives INSIDE each tab and is filtered
   to that tab's category: the IRS tab shows only IRS trades, the FX tab
   only FX trades.
3. IRS tenor ladder restricted to the desk's standard points: 1M, 2M, 3M,
   6M, 9M, 1Y-10Y, 12Y, 15Y, 20Y, 25Y, 30Y. Anything traded at another
   tenor rolls into one 'Other / non-standard tenor' row so totals still
   reconcile. Tenors with no trades that day are hidden (both tabs).
4. DV01 is the default comparison metric: DV01 (USD) ≈ USD notional ×
   tenor in years × 1bp — a flat-annuity approximation (no discounting)
   that normalizes volume across the curve. Rates products only; the FX
   tab automatically falls back to USD notional. Rows with no parseable
   single tenor (spread trades) get no DV01 and are flagged.
5. Month-to-date market share: pulls every business day of the month up
   to the selected date (cached 24h after first load) and shows an MTD
   share pie next to the daily one, converted with each day's ECB fixing.
6. ICAP daily snapshots: the tpSEF page has no date in its URL and only
   ever shows the latest business day, so each day the app saves that
   day's parsed ICAP table to disk (icap_snapshots/). Selecting a past
   date then loads the saved snapshot — accurate ICAP history from the
   first day the app ran. NOTE: on Streamlit Community Cloud the disk is
   wiped on redeploy/reboot; snapshots are fully durable on a local
   machine.

PREVIOUS VERSION (Tradition FX option notionals — still in effect)
--------------------------------------------------------------------------
Tradition's parser now reads the _NDA (Non-Delta-Adjusted, i.e. GROSS)
notional columns instead of the _DA (Delta-Adjusted) columns. Two reasons:
  1. Comparability — every other SEF here (LatAm, GFI, BGC, ICAP) reports
     gross notional; summing Tradition's delta-adjusted figure against
     those overstates Tradition's FX share.
  2. Tradition's DA column applies the delta as a WHOLE NUMBER (a 50-delta
     option on 10M shows 500M = 10M x 50, not 10M x 0.50), so DA is
     10-50x gross and 100x the true delta-adjusted figure. Verified on
     2026-07-09: strike-quoted options have DA == NDA exactly; every
     'delta of X' option has DA == NDA * X.
For swaps and NDFs, DA == NDA in every row, so this ONLY changes FX
option rows. The raw file is never modified — we just read the column
that means gross notional. (Mexico FX on 2026-07-09: 9.64B -> 594M.)

PREVIOUS VERSION (classification / 0M bucket / asset codes — still in effect)
--------------------------------------------------------------------------
1. FX options are now classified correctly. Descriptions like
       '1W USD vs. MXN Put at a delta of 50'
   were falling into "Other" because (a) Tradition was classified off the
   Sub_Prod column ALONE, and (b) the classifier didn't recognize option
   language. Fixes:
     - Tradition is now classified off Sub_Prod + description COMBINED.
     - The classifier recognizes a currency pair (USD vs. MXN, USDMXN,
       EUR/MXN...) combined with option markers (PUT, CALL, DELTA, RISK
       REVERSAL, BUTTERFLY, STRADDLE, STRANGLE, SEAGULL, DIGITAL, OPTION)
       and buckets those as FX.

2. 0M tenor bucket. ON/TN/SN, day tenors, and 1W/2W/3W now all roll up
   into a single "0M" row at the TOP of the comparison ladder — so
   sub-month FX trades are never dropped from the summary tables. The
   drill-down still shows the TRUE tenor (1W, 2W, ...) for each trade.

3. Nothing silently excluded. The comparison table now also carries a
   final "Other / unmatched" row, so its Total column reconciles EXACTLY
   with the market-share chart. A built-in reconciliation check warns on
   screen if the table total and the chart total ever disagree.

4. Asset-class codes. Every row now carries the raw regulatory asset
   class (IR / CU / CD) — taken directly from the source where published
   (GFI, BGC, ICAP) and derived from the product category otherwise.
   Shown ONLY in the individual-trades drill-down; the main tabs you
   switch between remain "IRS / Swap" and "FX".

PREVIOUS VERSION (accuracy / FX / layout rework — still in effect)
--------------------------------------------------------------------------
- Real FX conversion: Currency | Notional_Local | USD_Notional_Computed,
  using the ECB daily reference rate (via Frankfurter.app) for the report
  date. CLP/COP/PEN have no ECB fixing — manual sidebar entry, never
  silently estimated.
- Auditability: FX_Rate_Used / FX_Source per row, plus a cross-check
  against sources that publish their own USD figure (Tradition, LatAm),
  flagging any >1% disagreement.
- Mexico vs. Non-Mexico tabs off the MXN_Match flag (currency field OR
  TIIE/F-TIIE/UDI/MXN markers in the description).
- Exact numbers by default (toggle for 1.23M abbreviations).

WHY EACH SOURCE IS HANDLED DIFFERENTLY (parsing)
--------------------------------------------------------------------------
  - Tradition : pipe-delimited CSV, has trade count AND USD-converted notional
  - LatAm SEF : comma CSV, has USD-converted notional but NO trade count
  - GFI / BGC : identical Excel template ("SEF PERMITTED AND REQUIRED
                INSTRUMENTS..."), notional in LOCAL currency only, no
                trade count column
  - ICAP (tpSEF): a live HTML table at a fixed URL (no date in the link —
                it always shows the latest business day), WITH a real
                "Num of Trades" column
Because GFI/BGC/LatAm don't give a trade count, we fall back to counting
ROWS as a proxy and label it clearly as an estimate everywhere it's shown.

Run with:   streamlit run mexico_desk_dashboard.py
Install:    pip install streamlit pandas requests openpyxl beautifulsoup4 lxml plotly
"""

import streamlit as st
import pandas as pd
import requests
import re
import io
import plotly.express as px
from datetime import datetime, timedelta

st.set_page_config(page_title="Mexico Desk — Cross-SEF Dashboard", page_icon="🇲🇽", layout="wide")

# ═════════════════════════════════════════════════════════════════════════
# CONFIG — SOURCES
# ═════════════════════════════════════════════════════════════════════════
SOURCES = {
    "Tradition": {
        "kind": "tradition",
        "url_template": "https://www.traditionsef.com/dailyactivity/SEF16_MKTDATA_TFSU_{date}.csv",
        "date_fmt": "%Y%m%d",
    },
    "LatAm SEF": {
        "kind": "latam",
        "url_template": "https://www.latamsef.com/market-data/LatAmSEF_MarketActivityData_{date}.csv",
        "date_fmt": "%Y%m%d",
    },
    "GFI": {
        "kind": "gfi_bgc",
        "url_template": "https://www.gfigroup.com/doc/sef/marketdata/{date}_daily_trade_data.xlsx",
        "date_fmt": "%Y-%m-%d",
    },
    "BGC": {
        "kind": "gfi_bgc",
        "url_template": "https://www.bgcsef.com/TradingActivityReports/Daily/DailyAct_{date}.xlsx",
        "date_fmt": "%Y%m%d",
    },
    "ICAP": {
        "kind": "icap",
        "url_template": "https://www.tullettprebon.com/swap-execution-facility/daily-activity-summary.aspx",
        "date_fmt": None,       # no date in URL — page always shows latest business day
    },
}

SEF_COLORS = {
    "Tradition": "#1a56db",
    "LatAm SEF": "#0e9f6e",
    "GFI":       "#d97706",
    "BGC":       "#7e3af2",
    "ICAP":      "#e02424",
}

TENOR_ORDER = [
    "ON", "TN", "SN",
    "0M",                     # roll-up bucket for 1W/2W/3W (and odd day tenors)
    "1W", "2W", "3W",
    "1M", "2M", "3M", "4M", "5M", "6M", "7M", "8M", "9M", "10M", "11M",
    "1Y", "18M", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y",
    "10Y", "12Y", "15Y", "20Y", "25Y", "30Y", "Other"
]

# The FULL, fixed maturity ladder — Cash/ON, TN, and SN each on their OWN
# row, then 0M (the roll-up for 1W/2W/3W and odd day tenors), then every
# month a person can buy (1-11), then every year through 10Y. Past 10Y,
# only the standard market points (15Y, 20Y, 25Y, 30Y) are always shown;
# any other long-dated tenor (11Y, 12Y, 17Y, etc.) only appears if it
# actually traded that day — handled by build_comparison()'s "extra
# tenors" logic, which appends anything found in the data that isn't on
# this fixed ladder. A final "Other / unmatched" row (if present)
# guarantees the table total reconciles with the charts.
FULL_TENOR_LADDER = (
    ["ON", "TN", "SN", "0M"] +
    [f"{m}M" for m in range(1, 12)] +
    [f"{y}Y" for y in range(1, 11)] +
    ["15Y", "20Y", "25Y", "30Y"]
)
TENOR_DISPLAY_LABELS = {
    "0M": "0M (1W/2W/3W)",
    "ON": "Cash / ON",
    "TN": "TN",
    "SN": "SN",
    "Other": "Other / unmatched",
}

# The ONLY tenors displayed on the IRS ladder (desk request). Anything
# that traded at a tenor not on this list (e.g. 18M, 4M, ON) rolls into
# a single 'Other / non-standard tenor' row so the totals still
# reconcile — nothing is silently dropped, it's just not given its own row.
IRS_DISPLAY_TENORS = (
    ["1M", "2M", "3M", "6M", "9M"] +
    [f"{y}Y" for y in range(1, 11)] +
    ["12Y", "15Y", "20Y", "25Y", "30Y"]
)


def tenor_display(t: str) -> str:
    return TENOR_DISPLAY_LABELS.get(t, t)


# Only the week tenors (1W/2W/3W) roll into the "0M" reporting bucket for
# the comparison tables and charts — Cash/ON, TN, and SN each keep their
# OWN row on the ladder. Odd day tenors (e.g. '2D', '5D') also fall into
# 0M since they're sub-week oddballs with no ladder row of their own.
# The drill-down keeps the ORIGINAL tenor (1W, 2W, ...) so no information
# is lost — only the summary aggregation changes.
SUB_MONTH_TENORS = {"1W", "2W", "3W"}


def tenor_bucket(t: str) -> str:
    """Map a parsed tenor to its reporting bucket. Weeks (1W/2W/3W) and
    odd day tenors -> '0M'; ON/TN/SN and everything else unchanged."""
    if not isinstance(t, str):
        return "Other"
    original = t.strip()
    up = original.upper()
    if up in ("ON", "TN", "SN"):
        return up            # each keeps its own ladder row
    if up in SUB_MONTH_TENORS:
        return "0M"
    if re.fullmatch(r'(\d+)D', up):
        return "0M"          # odd day tenors (2D, 5D...) — no row of their own
    m = re.fullmatch(r'(\d+)W', up)
    if m and int(m.group(1)) <= 3:
        return "0M"
    return original  # pass through UNCHANGED ('Other', '1W vs 1M', '10Y', ...)


# Mexican TIIE swaps are quoted by NUMBER OF 28-DAY PERIODS, not calendar
# months/years — the market convention counts 13 periods (13*28=364 days)
# as "a year," 26 periods as "2 years," etc. This table converts those
# period-counts to the standard Y/M labels so every source's MXN swaps
# line up on the SAME row for comparison (verified by cross-checking
# prices: LatAm's 13x1 and Tradition's "1 Year" both quote 6.735 on the
# same day — they are the same instrument, just named differently).
LATAM_PERIOD_TENOR = {
    1: "1M", 3: "3M", 6: "6M", 9: "9M", 13: "1Y", 19: "18M", 26: "2Y",
    39: "3Y", 52: "4Y", 65: "5Y", 78: "6Y", 91: "7Y", 104: "8Y",
    117: "9Y", 130: "10Y", 195: "15Y", 260: "20Y", 390: "30Y",
}

# ═════════════════════════════════════════════════════════════════════════
# CONFIG — FX
# ═════════════════════════════════════════════════════════════════════════
# Frankfurter.app republishes the ECB's daily reference rates (fixed
# 16:00 CET) for free, with no API key, and supports historical dates —
# the closest thing to an "official daily fixing" available without a
# paid subscription (Banxico's own SIE API requires a registered token,
# so it isn't used here to keep this a zero-install dashboard).
FRANKFURTER_BASE = "https://api.frankfurter.app"
ECB_CURRENCIES = [
    "EUR", "GBP", "JPY", "CAD", "CHF", "SEK", "NOK", "TRY", "HUF", "MXN",
    "AUD", "NZD", "DKK", "PLN", "CZK", "ILS", "SGD", "ZAR",
]
# NOT published by the ECB — must be entered manually in the sidebar if
# any trades in these currencies show up. Never silently estimated.
EXOTIC_NO_ECB_RATE = ["CLP", "COP", "PEN"]

# ═════════════════════════════════════════════════════════════════════════
# SHARED HELPERS
# ═════════════════════════════════════════════════════════════════════════
def last_business_day() -> datetime:
    d = datetime.today() - timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


def fmt_m(n):
    """Format a raw number as millions/billions with 2-3 decimals."""
    if n is None or pd.isna(n):
        return "—"
    if abs(n) >= 1e9:
        return f"{n/1e9:,.3f}B"
    if abs(n) >= 1e6:
        return f"{n/1e6:,.2f}M"
    return f"{n:,.0f}"


def fmt_exact(n):
    """Format a raw number with full precision (no abbreviation)."""
    if n is None or pd.isna(n):
        return "—"
    return f"{n:,.0f}"


def tenor_sort_key(t):
    try:
        return TENOR_ORDER.index(t)
    except ValueError:
        return len(TENOR_ORDER)


def sort_tenors(series: pd.Series) -> pd.Categorical:
    vals = series.unique().tolist()
    known = [t for t in TENOR_ORDER if t in vals]
    other = sorted([t for t in vals if t not in known])
    return pd.Categorical(series, categories=known + other, ordered=True)


def tenor_to_years(t) -> float:
    """Convert a tenor label to years for the DV01 approximation.
    Returns None for spread tenors ('1W vs 1M') and unparseable rows —
    those get no DV01 and are excluded from DV01 comparisons."""
    if not isinstance(t, str):
        return None
    up = t.strip().upper()
    if up in ("ON", "TN", "SN"):
        return 1 / 365
    m = re.fullmatch(r'(\d+)Y', up)
    if m:
        return float(m.group(1))
    m = re.fullmatch(r'(\d+)M', up)
    if m:
        return int(m.group(1)) / 12
    m = re.fullmatch(r'(\d+)W', up)
    if m:
        return int(m.group(1)) / 52
    m = re.fullmatch(r'(\d+)D', up)
    if m:
        return int(m.group(1)) / 365
    return None


def add_dv01(df: pd.DataFrame) -> pd.DataFrame:
    """DV01 (USD) ≈ USD notional × tenor in years × 1bp. A flat-annuity
    desk approximation (no discounting) — good for cross-SEF volume
    comparison, not for risk. Only meaningful for rates products, so FX
    rows get NA and the FX section falls back to USD notional."""
    df = df.copy()
    df["Tenor_Years"] = df["Tenor"].astype(str).apply(tenor_to_years)
    df["DV01_USD"] = pd.NA
    is_irs = df["Category"] == "IRS / Swap"
    usd = pd.to_numeric(df["USD_Notional_Computed"], errors="coerce")
    yrs = pd.to_numeric(df["Tenor_Years"], errors="coerce")
    df.loc[is_irs, "DV01_USD"] = (usd * yrs * 1e-4)[is_irs]
    return df


# ── Product classification ────────────────────────────────────────────────
# Order matters, and mirrors how a Mexico rates & FX broker reads a ticket:
#   1. Rate-index / swap markers (TIIE, SOFR, IRS, BASIS, XCCY...) → IRS.
#      Checked FIRST so a "USDMXN basis swap" stays an IRS even though it
#      names a currency pair.
#   2. Currency pair + option language (Put/Call/delta/RR/fly/straddle...)
#      → FX. Catches tickets like '1W USD vs. MXN Put at a delta of 50'
#      that carry no NDF/FWD/FXO keyword at all.
#   3. Generic FX markers (NDF, FWD, SPOT, FXO...) → FX.
#   4. Otherwise → Other.
FX_OPTION_MARKERS = [
    "PUT", "CALL", "DELTA", "RISK REVERSAL", "BUTTERFLY",
    "STRADDLE", "STRANGLE", "SEAGULL", "DIGITAL", "OPTION",
]
# 'USD vs. MXN', 'USD v MXN', 'USD/MXN', 'USDMXN', 'USD-MXN' ...
CCY_PAIR_RE = re.compile(
    r'\b[A-Z]{3}\s*(?:VS\.?|V\.?|/|-)?\s*(?:MXN|USD|EUR|GBP|JPY|CAD|CLP|COP|PEN|BRL|CHF|AUD|NZD)\b'
)


def _has_ccy_pair(t: str) -> bool:
    if re.search(r'\b([A-Z]{3})\s*(?:VS\.?|V\.?|/|-)\s*([A-Z]{3})\b', t):
        return True
    # concatenated pairs like USDMXN / EURMXN / MXNUSD
    if re.search(r'\b(USD|EUR|GBP|JPY|CAD|CHF|AUD|NZD)(MXN|CLP|COP|PEN|BRL|USD|EUR|JPY)\b', t):
        return True
    return False


def classify_category(text: str) -> str:
    """Bucket a product into IRS/Swap, FX, or Other based on keywords."""
    if not isinstance(text, str):
        return "Other"
    t = text.upper()
    # 'FX Swap' is an FX product — must be caught BEFORE the generic 'SWAP'
    # marker below sends it to the rates bucket.
    if "FX SWAP" in t or "FXSWAP" in t:
        return "FX"
    ir_markers = ["IRS", "OIS", "SWAP", "BASIS", "XBS", "XCCY", "FRA", "SBS",
                  "ZCI", "TIIE", "SOFR", "SONIA", "ESTR", "CORRA", "TONAT", "TLREF"]
    if any(m in t for m in ir_markers):
        return "IRS / Swap"
    # FX options with no explicit FX keyword — e.g.
    # '1W USD vs. MXN Put at a delta of 50'
    if _has_ccy_pair(t) and any(m in t for m in FX_OPTION_MARKERS):
        return "FX"
    fx_markers = ["NDF", "FX", "FWD", "FXO", "SPOT", "CALL", "PUT"]
    if any(m in t for m in fx_markers):
        return "FX"
    return "Other"


# Regulatory asset-class codes, as printed on GFI/BGC/ICAP files:
#   IR = interest rates, CU = currency/FX, CD = credit.
# GFI/BGC/ICAP publish these directly — we keep theirs verbatim. For
# Tradition/LatAm (which don't publish a code) we derive it from the
# product category. Shown ONLY in the individual-trades drill-down; the
# main navigation stays IRS / FX.
CATEGORY_TO_ASSETCLASS = {"IRS / Swap": "IR", "FX": "CU"}


def derive_asset_class(category: str) -> str:
    return CATEGORY_TO_ASSETCLASS.get(category, "—")


# TIIE (Mexico's interbank rate), F-TIIE (its futures-referenced variant),
# and UDI (Mexico's inflation-linked unit) are Mexico-specific markers that
# can appear WITHOUT the literal string "MXN" — e.g. a USD-settled
# cross-currency basis swap referencing F-TIIE is still a Mexican trade.
# Checking for these catches those rows that a plain "MXN" search misses.
MEXICO_MARKERS = ["MXN", "TIIE", "F-TIIE", "UDI"]


def contains_mxn(row_text: str) -> bool:
    if not isinstance(row_text, str):
        return False
    t = row_text.upper()
    return any(marker in t for marker in MEXICO_MARKERS)


# ═════════════════════════════════════════════════════════════════════════
# FX — fetch official rates + convert every row to USD
# ═════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=1800, show_spinner=False)
def fetch_fx_rates(report_date_str: str):
    """
    Fetches ECB daily reference rates for report_date_str ('YYYY-MM-DD'),
    quoted as 'units of CCY per 1 USD' (so USD itself = 1.0, and the same
    formula — Local_Notional / rate — converts ANY currency to USD
    regardless of whether that currency is normally quoted as CCY-per-USD
    or USD-per-CCY in the market).

    Returns (rates_dict, source_label, note_or_None, error_or_None).
    """
    url = f"{FRANKFURTER_BASE}/{report_date_str}?from=USD&to={','.join(ECB_CURRENCIES)}"
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        return {}, "ECB daily reference rate — FETCH FAILED", None, f"FX fetch error: {e}"

    rates = dict(data.get("rates", {}))
    rates["USD"] = 1.0
    actual_date = data.get("date", report_date_str)
    note = None
    if actual_date != report_date_str:
        note = (f"ECB had no fixing dated {report_date_str} (weekend/holiday) — "
                f"used the most recent prior business day instead: {actual_date}.")
    return rates, f"ECB daily reference rate ({actual_date})", note, None


def apply_fx(df: pd.DataFrame, fx_rates: dict, manual_overrides: dict, fx_source_label: str) -> pd.DataFrame:
    """
    Adds, for every row:
      USD_Notional_Computed  — Original_Notional / (CCY per 1 USD), or NA if no rate
      FX_Rate_Used           — the rate applied (CCY per 1 USD)
      FX_Source              — where that rate came from
      FX_CrossCheck          — vs. the source's OWN published USD figure, if any
    Nothing is ever guessed: if a currency has no rate (official or manual),
    USD_Notional_Computed is NA and FX_Source says so explicitly.
    """
    df = df.copy()

    def convert(row):
        ccy = row["Currency"]
        local = row["Notional_Local"]
        if ccy == "USD":
            return pd.Series([local, 1.0, "USD (no conversion needed)"])
        if ccy in fx_rates and fx_rates.get(ccy):
            rate = fx_rates[ccy]
            usd = local / rate if rate else pd.NA
            return pd.Series([usd, rate, fx_source_label])
        if ccy in manual_overrides and manual_overrides.get(ccy):
            rate = manual_overrides[ccy]
            return pd.Series([local / rate, rate, "Manual entry (user-provided — NOT an official fixing)"])
        return pd.Series([pd.NA, pd.NA, "Missing — no FX rate available for this currency"])

    df[["USD_Notional_Computed", "FX_Rate_Used", "FX_Source"]] = df.apply(convert, axis=1)

    def check(row):
        pub = row.get("Notional_USD")
        calc = row["USD_Notional_Computed"]
        if pd.isna(pub) or pd.isna(calc) or pub == 0:
            return "N/A"
        diff_pct = abs(pub - calc) / abs(pub) * 100
        return "✓ Match" if diff_pct < 1.0 else f"⚠ {diff_pct:.1f}% diff"

    df["FX_CrossCheck"] = df.apply(check, axis=1)
    return df


# ═════════════════════════════════════════════════════════════════════════
# PARSER 1 — TRADITION (pipe-delimited CSV)
# ═════════════════════════════════════════════════════════════════════════
def extract_tenor_tradition(desc: str) -> str:
    """
    Tradition descriptions look like '3 Month Overnight Index Swap...' or
    '10 Year Overnight Index Swap...'. IMPORTANT: 'Overnight Index Swap' is
    the PRODUCT TYPE (OIS), not the tenor — the tenor is the leading number.
    """
    if not isinstance(desc, str):
        return "Other"
    d = desc.upper().strip()
    sp = re.search(
        r'(\d+\s*(?:Y|YR|YEAR|M|MO|MONTH|W|WK|WEEK))\s*(?:VS\.?|V\.?|/)\s*(\d+\s*(?:Y|YR|YEAR|M|MO|MONTH|W|WK|WEEK))', d)
    if sp:
        def fmt(t):
            m = re.match(r'(\d+)\s*(Y|YR|YEAR|M|MO|MONTH|W|WK|WEEK)', t.strip())
            if not m:
                return t.strip()
            n, u = m.group(1), m.group(2)
            return f"{n}Y" if u[0] == "Y" else (f"{n}M" if u[0] == "M" else f"{n}W")
        return f"{fmt(sp.group(1))} vs {fmt(sp.group(2))}"
    m = re.match(r'^(\d+)\s*(DAYS?|WKS?|WEEKS?|MOS?|MONTHS?|YRS?|YEARS?)\b', d)
    if m:
        num, unit = m.group(1), m.group(2)
        if unit.startswith("Y"):
            return f"{num}Y"
        if unit.startswith("MO") or unit.startswith("MONTH"):
            return f"{num}M"
        if unit.startswith("W"):
            return f"{num}W"
        if unit.startswith("D"):
            return "ON" if num == "1" else f"{num}D"
    if re.search(r'\bO/N\b', d):
        return "ON"
    if re.search(r'\bT/N\b', d):
        return "TN"
    if d.startswith("OVERNIGHT") and "INDEX SWAP" not in d:
        return "ON"
    m2 = re.search(r'(\d+)\s*(Y|YR|YEAR|M|MO|MONTH|W|WK|WEEK)', d)
    if m2:
        n, u = m2.group(1), m2.group(2)
        if u.startswith("Y"):
            return f"{n}Y"
        if u.startswith("M"):
            return f"{n}M"
        if u.startswith("W"):
            return f"{n}W"
    return "Other"


def parse_tradition(raw_text: str) -> pd.DataFrame:
    df = pd.read_csv(io.StringIO(raw_text), sep="|", dtype=str, on_bad_lines="skip")
    df.columns = [c.strip() for c in df.columns]
    # NOTE — _NDA (Non-Delta-Adjusted = GROSS) notionals, not _DA:
    # Tradition's _DA columns multiply option notional by the delta as a
    # WHOLE number (a 50-delta option on 10M shows 500M, i.e. 10M x 50,
    # not 10M x 0.50), so _DA is 10-50x gross for delta-quoted options and
    # not comparable to the gross notionals every other SEF publishes.
    # For swaps/NDFs, DA == NDA in every row, so only options are affected.
    for col in ["Total_Notional_USD_NDA", "Notional_Traded_Currency_NDA",
                "Total_Trade_Count", "First_Price", "Last_Price"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    desc_col = next((c for c in ["Internal_Prod_Des", "Internal_Prod_ID"] if c in df.columns), None)
    # Classify off Sub_Prod AND the description COMBINED. Classifying off
    # Sub_Prod alone (as before) sent FX options like '1W USD vs. MXN Put
    # at a delta of 50' into "Other" whenever the Sub_Prod field carried a
    # generic label with no FX keyword in it.
    sub_prod = df["Sub_Prod"].astype(str) if "Sub_Prod" in df.columns else pd.Series([""] * len(df))
    desc_ser = df[desc_col].astype(str) if desc_col else pd.Series([""] * len(df))
    cat_text = (sub_prod.fillna("") + " " + desc_ser.fillna(""))
    categories = cat_text.apply(classify_category)
    out = pd.DataFrame({
        "SEF": "Tradition",
        "Tenor": df[desc_col].apply(extract_tenor_tradition) if desc_col else "Other",
        "Category": categories,
        "AssetClass": categories.map(derive_asset_class),  # Tradition publishes no IR/CU/CD code — derived
        "Currency": df["Curr_Code"].str.strip().str.upper() if "Curr_Code" in df.columns else "N/A",
        "Notional_Local": df["Notional_Traded_Currency_NDA"] if "Notional_Traded_Currency_NDA" in df.columns else 0,
        "Notional_USD": df["Total_Notional_USD_NDA"] if "Total_Notional_USD_NDA" in df.columns else pd.NA,
        "Trades": df["Total_Trade_Count"] if "Total_Trade_Count" in df.columns else 1,
        "Trades_Estimated": False,
        "Last_Price": df["Last_Price"] if "Last_Price" in df.columns else pd.NA,
        "Description": df[desc_col] if desc_col else "",
        "MXN_Match": (df[desc_col].astype(str) if desc_col else pd.Series([""] * len(df))).apply(contains_mxn) |
                     (df["Curr_Code"].astype(str) == "MXN" if "Curr_Code" in df.columns else False),
    })
    return out


# ═════════════════════════════════════════════════════════════════════════
# PARSER 2 — LATAM SEF (comma CSV, has USD notional but no trade count)
# ═════════════════════════════════════════════════════════════════════════
def days_to_tenor_bucket(days: int) -> str:
    """Map a day-count to the nearest standard FX tenor bucket."""
    if days <= 1:
        return "ON"
    if days <= 3:
        return "TN"
    if days <= 10:
        return "1W"
    if days <= 18:
        return "2W"
    if days <= 25:
        return "3W"
    if days <= 45:
        return "1M"
    if days <= 75:
        return "2M"
    if days <= 105:
        return "3M"
    if days <= 195:
        return "6M"
    if days <= 285:
        return "9M"
    if days <= 400:
        return "1Y"
    if days <= 600:
        return "18M"
    if days <= 760:
        return "2Y"
    return "Other"


def extract_tenor_latam(row) -> str:
    """LatAm encodes tenor two different ways depending on currency:
    - MXN rows: trailing '- 130x1' (periods of 28 days)
    - CLP/COP rows: trailing '- 10-Year' or '- 12-Month' (plain English)
    - FX rows: 'NDF CCY - DDMonYY' (an explicit expiry date)
    """
    desc = row.get("Internal_Prod_Des", "")
    if not isinstance(desc, str):
        return "Other"
    d = desc.strip()
    m0 = re.search(r'-\s*(\d+)-(Year|Month)s?\s*$', d, re.IGNORECASE)
    if m0:
        num, unit = m0.group(1), m0.group(2)
        return f"{num}Y" if unit.lower() == "year" else f"{num}M"
    m = re.search(r'-\s*(\d+)x\d+\s*$', d)
    if m:
        periods = int(m.group(1))
        if periods in LATAM_PERIOD_TENOR:
            return LATAM_PERIOD_TENOR[periods]
        years = periods / 13.0
        if years >= 1:
            return f"{round(years)}Y"
        months = round(periods * 28 / 30)
        return f"{months}M"
    m2 = re.search(r'-\s*(\d{1,2}[A-Za-z]{3}\d{2,4})\s*$', d)
    if m2:
        try:
            expiry = datetime.strptime(m2.group(1), "%d%b%y")
        except ValueError:
            try:
                expiry = datetime.strptime(m2.group(1), "%d%b%Y")
            except ValueError:
                return "Other"
        trade_date_str = str(row.get("Trade_Date", ""))
        try:
            trade_date = datetime.strptime(trade_date_str, "%Y%m%d")
        except ValueError:
            return "Other"
        days = (expiry - trade_date).days
        return days_to_tenor_bucket(days)
    return "Other"


def parse_latam(raw_text: str) -> pd.DataFrame:
    df = pd.read_csv(io.StringIO(raw_text), dtype=str)
    df.columns = [c.strip() for c in df.columns]
    for col in ["Notional_USD", "Notional_Traded_Currency", "Last_Price"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    categories = (df["Internal_Prod_Des"].apply(classify_category)
                  if "Internal_Prod_Des" in df.columns
                  else pd.Series(["Other"] * len(df)))
    out = pd.DataFrame({
        "SEF": "LatAm SEF",
        "Tenor": df.apply(extract_tenor_latam, axis=1),
        "Category": categories,
        "AssetClass": categories.map(derive_asset_class),  # LatAm publishes no IR/CU/CD code — derived
        "Currency": df["Curr_Code"].str.strip().str.upper() if "Curr_Code" in df.columns else "N/A",
        "Notional_Local": df["Notional_Traded_Currency"] if "Notional_Traded_Currency" in df.columns else 0,
        "Notional_USD": df["Notional_USD"] if "Notional_USD" in df.columns else pd.NA,
        "Trades": 1,  # LatAm gives no trade count — one row = one print (proxy)
        "Trades_Estimated": True,
        "Last_Price": df["Last_Price"] if "Last_Price" in df.columns else pd.NA,
        "Description": df["Internal_Prod_Des"] if "Internal_Prod_Des" in df.columns else "",
        "MXN_Match": (df["Curr_Code"].astype(str) == "MXN" if "Curr_Code" in df.columns else False) |
                     (df["Internal_Prod_Des"].astype(str).apply(contains_mxn) if "Internal_Prod_Des" in df.columns else False),
    })
    return out


# ═════════════════════════════════════════════════════════════════════════
# PARSER 3 — GFI / BGC (identical Excel template, LOCAL currency notional)
# ═════════════════════════════════════════════════════════════════════════
def extract_tenor_ccy_gfi_bgc(desc: str):
    """
    GFI/BGC descriptions embed tenor+currency right together, e.g.
    '...10YMXN MXN 28D1 200...' or '...6MMXN MXN 28D1 200...'.
    Returns (tenor, currency) or (None, None) if no match.
    """
    if not isinstance(desc, str):
        return None, None
    m = re.search(r'\b(\d+)(Y|M)([A-Z]{3})\b', desc.upper())
    if m:
        num, unit, ccy = m.group(1), m.group(2), m.group(3)
        tenor = f"{num}{unit}"
        return tenor, ccy
    return None, None


def extract_fx_tenor_gfi_bgc(desc: str, report_date: datetime):
    """
    FX option rows look like 'EURMXN CALL 25D NEW YORK 06NOV2026 C1
    BILATERAL'. The trailing date is the expiry; tenor = expiry - report date.
    """
    if not isinstance(desc, str):
        return "Other"
    m = re.search(r'(\d{1,2}[A-Z]{3}\d{4})', desc.upper())
    if m:
        try:
            expiry = datetime.strptime(m.group(1), "%d%b%Y")
            days = (expiry - report_date).days
            return days_to_tenor_bucket(days)
        except ValueError:
            return "Other"
    return "Other"


def _col_letters_to_index(col_str: str) -> int:
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx - 1  # 0-based


def read_xlsx_stdlib(raw_bytes: bytes, sheet_name: str):
    """
    Minimal, dependency-free XLSX reader using ONLY Python's standard
    library (zipfile + xml.etree.ElementTree). Automatic fallback when
    openpyxl isn't installed — means GFI/BGC work with zero installs, same
    as Tradition/LatAm's plain CSVs.
    Returns a list of rows (each a list of string cell values).
    """
    import zipfile
    import xml.etree.ElementTree as ET
    NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    with zipfile.ZipFile(io.BytesIO(raw_bytes)) as z:
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.get("Id"): rel.get("Target") for rel in rels_xml}
        sheet_file = None
        for sheet in wb_xml.find("m:sheets", NS):
            if sheet.get("name") == sheet_name:
                rid = sheet.get(f"{REL_NS}id")
                target = rid_to_target.get(rid, "")
                sheet_file = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
                break
        if sheet_file is None:
            first = list(wb_xml.find("m:sheets", NS))[0]
            rid = first.get(f"{REL_NS}id")
            target = rid_to_target.get(rid, "")
            sheet_file = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_xml.findall("m:si", NS):
                texts = si.findall(".//m:t", NS)
                shared_strings.append("".join(t.text or "" for t in texts))
        sheet_xml = ET.fromstring(z.read(sheet_file))
        sheet_data = sheet_xml.find("m:sheetData", NS)
        rows = []
        for row_el in sheet_data.findall("m:row", NS):
            row_cells, max_col = {}, -1
            for c_el in row_el.findall("m:c", NS):
                ref = c_el.get("r", "A1")
                col_idx = _col_letters_to_index("".join(ch for ch in ref if ch.isalpha()))
                max_col = max(max_col, col_idx)
                cell_type = c_el.get("t")
                v_el = c_el.find("m:v", NS)
                if v_el is None or v_el.text is None:
                    value = ""
                elif cell_type == "s":
                    value = shared_strings[int(v_el.text)]
                else:
                    value = v_el.text
                row_cells[col_idx] = value
            rows.append([row_cells.get(i, "") for i in range(max_col + 1)])
    return rows


def parse_gfi_bgc(raw_bytes: bytes, sef_name: str, report_date: datetime) -> pd.DataFrame:
    try:
        df = pd.read_excel(io.BytesIO(raw_bytes), sheet_name="SEFTrades", header=2)
    except ImportError:
        # openpyxl not installed — fall back to the zero-dependency reader
        # above. header=2 in the pandas call means "skip the first 2 rows",
        # so we do the same thing manually: rows[2] is the header.
        rows = read_xlsx_stdlib(raw_bytes, "SEFTrades")
        header = rows[2]
        data_rows = [r[:len(header)] + [""] * (len(header) - len(r)) for r in rows[3:]]
        df = pd.DataFrame(data_rows, columns=header)

    df = df[df["AssetClass"].isin(["CD", "CU", "IR"])].copy()
    for col in ["Open", "Low", "High", "Close", "Volume"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    tenors, currencies, categories = [], [], []
    for _, row in df.iterrows():
        desc = row.get("InstrumentDescription", "")
        asset_class = row.get("AssetClass", "")
        if asset_class == "IR":
            tenor, ccy = extract_tenor_ccy_gfi_bgc(desc)
            tenors.append(tenor or "Other")
            currencies.append(ccy or row.get("Currency", "N/A"))
            categories.append("IRS / Swap")
        elif asset_class == "CU":
            tenor = extract_fx_tenor_gfi_bgc(desc, report_date)
            tenors.append(tenor)
            currencies.append(row.get("Currency", "N/A"))
            categories.append("FX")
        else:
            tenors.append("Other")
            currencies.append(row.get("Currency", "N/A"))
            categories.append("Other")

    out = pd.DataFrame({
        "SEF": sef_name,
        "Tenor": tenors,
        "Category": categories,
        "AssetClass": df["AssetClass"].values,  # real code straight from the file (IR / CU / CD)
        "Currency": [str(c).strip().upper() for c in currencies],
        "Notional_Local": df["Volume"].values if "Volume" in df.columns else 0,
        "Notional_USD": pd.NA,  # GFI/BGC don't give USD-converted notional
        "Trades": 1,            # no trade count column — one row = one print (proxy)
        "Trades_Estimated": True,
        "Last_Price": df["Close"].values if "Close" in df.columns else pd.NA,
        "Description": df["InstrumentDescription"].values,
        "MXN_Match": df["InstrumentDescription"].astype(str).apply(contains_mxn).values,
    })
    return out


# ═════════════════════════════════════════════════════════════════════════
# PARSER 4 — ICAP / tpSEF (live HTML table, no date in URL)
# ═════════════════════════════════════════════════════════════════════════
def extract_tenor_icap(instrument: str, description: str) -> str:
    """
    ICAP mixes several genuinely different naming conventions in the same
    feed — confirmed against a real full-day export:
      1. 'MXN.13*1.F-TIIE.OIS'   — dot/asterisk period notation, same math
         as LatAm's '13x1' (13 periods of 28 days ~ 1 year)
      2. 'IRS_MXN_..._0X2'       — 'AxB' forward-tenor notation where B is
         the tenor in MONTHS
      3. '1M_USD_MXN_A_FXO...'   — leading tenor, underscore-delimited
      4. '10Y CAD IRB Outright...' — leading tenor, space-delimited
      5. 'A-EUR 10Y'             — trailing tenor style
    Checked in this order because 1 and 2 are narrow/specific patterns that
    must be tried before the broader leading/trailing fallbacks.
    """
    text = instrument if isinstance(instrument, str) else ""
    up = text.upper()

    m1 = re.search(r'\.(\d+)\*\d+\.', up)
    if m1:
        periods = int(m1.group(1))
        if periods in LATAM_PERIOD_TENOR:
            return LATAM_PERIOD_TENOR[periods]
        years = periods / 13.0
        if years >= 1:
            return f"{round(years)}Y"
        return f"{round(periods * 28 / 30)}M"

    m2 = re.search(r'(?:^|_)(\d+)X(\d+)(?:_|$)', up)
    if m2:
        months = int(m2.group(2))
        if months >= 12 and months % 12 == 0:
            return f"{months // 12}Y"
        return f"{months}M"

    m3 = re.match(r'^(\d+)([DWMY])(?=[_\s]|$)', up)
    if m3:
        num, unit = m3.group(1), m3.group(2)
        return "ON" if (unit == "D" and int(num) <= 1) else f"{num}{unit}"

    m4 = re.match(r'^(\d+)\s*(DAY|WK|WEEK|MO|MONTH|YR|YEAR)S?\b', up)
    if m4:
        num, unit = m4.group(1), m4.group(2)
        if unit.startswith("Y"):
            return f"{num}Y"
        if unit.startswith("MO") or unit.startswith("MONTH"):
            return f"{num}M"
        if unit.startswith("W"):
            return f"{num}W"
        if unit.startswith("D"):
            return "ON" if num == "1" else f"{num}D"

    m5 = re.search(r'\b(\d+)([YM])\b', up)
    if m5:
        return f"{m5.group(1)}{m5.group(2)}"

    m6 = re.search(r'(\d+)[\s-]?(YEAR|MONTH)S?\b', up)
    if m6:
        num, unit = m6.group(1), m6.group(2)
        return f"{num}Y" if unit == "YEAR" else f"{num}M"

    return "Other"


def icap_category(asset_class: str) -> str:
    """ICAP already labels each row's Asset Class directly — use it rather
    than guessing from free-text, which misses rows like 'A-EUR 10Y' that
    carry no product-type keyword at all."""
    ac = str(asset_class).strip().upper()
    if ac == "IR":
        return "IRS / Swap"
    if ac == "CU":
        return "FX"
    return "Other"


# Known rate-index -> currency, for rows that name the index instead of
# a currency code (e.g. '10Y SOFR LCH' is a USD product; SOFR never
# appears with an explicit 'USD' token anywhere in the row).
ICAP_INDEX_CCY = {
    "SOFR": "USD", "SONIA": "GBP", "ESTR": "EUR", "CORRA": "CAD",
    "TONAT": "JPY", "TONA": "JPY", "TLREF": "TRY",
}
ICAP_CCY_CODES = ["MXN", "USD", "EUR", "GBP", "JPY", "CAD", "CLP", "COP",
                  "PEN", "NOK", "TRY", "AUD", "NZD", "CHF", "SEK", "HUF"]


def icap_infer_currency(text: str) -> str:
    if not isinstance(text, str):
        return "N/A"
    t = text.upper()
    # NOTE: use a letter-lookaround instead of \b — underscore counts as a
    # \w character, so \b does NOT match between '_' and 'USD' in strings
    # like '1M_USD_MXN_A_FXO'.
    for code in ICAP_CCY_CODES:
        if re.search(rf'(?<![A-Z]){code}(?![A-Z])', t):
            return code
    for idx, ccy in ICAP_INDEX_CCY.items():
        if idx in t:
            return ccy
    return "N/A"


class _StdlibTableExtractor:
    """
    Extracts every <table>...</table> on a page using ONLY Python's built-in
    html.parser — no bs4, no lxml. Automatic fallback when BeautifulSoup
    isn't installed, so ICAP works with zero installs too.
    """
    def __init__(self):
        from html.parser import HTMLParser

        class _Inner(HTMLParser):
            def __init__(inner_self):
                super().__init__()
                inner_self.tables = []
                inner_self.depth = 0
                inner_self.cur_table = None
                inner_self.cur_row = None
                inner_self.cur_cell = None
                inner_self.in_cell = False

            def handle_starttag(inner_self, tag, attrs):
                if tag == "table":
                    inner_self.depth += 1
                    if inner_self.depth == 1:
                        inner_self.cur_table = []
                elif tag == "tr" and inner_self.depth:
                    inner_self.cur_row = []
                elif tag in ("td", "th") and inner_self.depth:
                    inner_self.in_cell = True
                    inner_self.cur_cell = []

            def handle_endtag(inner_self, tag):
                if tag in ("td", "th") and inner_self.in_cell:
                    inner_self.cur_row.append("".join(inner_self.cur_cell).strip())
                    inner_self.in_cell = False
                elif tag == "tr" and inner_self.cur_row is not None:
                    inner_self.cur_table.append(inner_self.cur_row)
                    inner_self.cur_row = None
                elif tag == "table":
                    if inner_self.depth == 1 and inner_self.cur_table is not None:
                        inner_self.tables.append(inner_self.cur_table)
                    inner_self.depth = max(0, inner_self.depth - 1)

            def handle_data(inner_self, data):
                if inner_self.in_cell:
                    inner_self.cur_cell.append(data)

        self._parser = _Inner()

    def extract(self, html_text: str):
        self._parser.feed(html_text)
        return self._parser.tables


def extract_tables(html_text: str):
    """Try BeautifulSoup first (if installed); fall back to pure stdlib."""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_text, "html.parser")
        tables = []
        for t in soup.find_all("table"):
            rows = []
            for tr in t.find_all("tr"):
                cells = [c.get_text(strip=True) for c in tr.find_all(["td", "th"])]
                if cells:
                    rows.append(cells)
            tables.append(rows)
        return tables
    except ImportError:
        return _StdlibTableExtractor().extract(html_text)


def _standardize_icap_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Shared processing for ICAP data, regardless of whether it came from
    the live HTML page or a manually-exported .numbers file."""
    df = df.rename(columns={k: v for k, v in {
        "Asset Class": "AssetClass", "Tradeable Instrument": "Instrument",
        "Description": "Description", "Num of Trades": "Trades",
        "Total Notional Value": "Notional",
    }.items() if k in df.columns})
    for col in ["Trades", "Notional"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    combined_text = (df.get("Instrument", "").astype(str) + " " + df.get("Description", "").astype(str))
    return pd.DataFrame({
        "SEF": "ICAP",
        "Tenor": [extract_tenor_icap(i, d) for i, d in zip(df.get("Instrument", ""), df.get("Description", ""))],
        "Category": df["AssetClass"].apply(icap_category) if "AssetClass" in df.columns else "Other",
        "AssetClass": (df["AssetClass"].astype(str).str.strip().str.upper()
                       if "AssetClass" in df.columns else "—"),  # real code from the page (IR / CU / CD)
        "Currency": combined_text.apply(icap_infer_currency),
        "Notional_Local": df["Notional"] if "Notional" in df.columns else 0,
        "Notional_USD": pd.NA,
        "Trades": df["Trades"] if "Trades" in df.columns else 1,
        "Trades_Estimated": False,
        "Last_Price": df["Closing Price"] if "Closing Price" in df.columns else pd.NA,
        "Description": combined_text,
        "MXN_Match": combined_text.apply(contains_mxn),
    })


def parse_icap(html_text: str) -> pd.DataFrame:
    tables = extract_tables(html_text)
    tables = [t for t in tables if len(t) > 1]
    if not tables:
        raise ValueError("No data table found on the page")
    best = max(tables, key=len)
    header, data_rows = best[0], best[1:]
    data_rows = [r[:len(header)] + [""] * (len(header) - len(r)) for r in data_rows]
    df = pd.DataFrame(data_rows, columns=header)
    df.columns = [str(c).strip() for c in df.columns]
    return _standardize_icap_rows(df)


def parse_icap_numbers(raw_bytes: bytes) -> pd.DataFrame:
    """Parses a manually-exported Apple Numbers (.numbers) file of the same
    tpSEF daily activity table. Requires 'numbers-parser' (not stdlib)."""
    from numbers_parser import Document
    import tempfile
    import os
    with tempfile.NamedTemporaryFile(suffix=".numbers", delete=False) as tmp:
        tmp.write(raw_bytes)
        tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
        table = doc.sheets[0].tables[0]
        rows = table.rows(values_only=True)
    finally:
        os.unlink(tmp_path)
    header, data_rows = rows[0], rows[1:]
    df = pd.DataFrame(data_rows, columns=header)
    df.columns = [str(c).strip() for c in df.columns]
    return _standardize_icap_rows(df)


# ═════════════════════════════════════════════════════════════════════════
# FETCHERS (network layer — cached so we don't re-download on every click)
# ═════════════════════════════════════════════════════════════════════════
BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
                  "(KHTML, like Gecko) Version/17.4.1 Safari/605.1.15",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
}

# BGC and GFI need a warm session — visit the homepage first so the server
# sees a realistic referrer and any session cookies before requesting the file.
SITE_HOMEPAGES = {
    "BGC": "https://www.bgcsef.com/",
    "GFI": "https://www.gfigroup.com/",
}


def _make_session(homepage: str) -> requests.Session:
    s = requests.Session()
    s.headers.update(BROWSER_HEADERS)
    try:
        s.get(homepage, timeout=15)
        s.headers["Referer"] = homepage
    except Exception:
        pass
    return s


def _fetch_protected(url: str, homepage: str):
    """Fetch a file from a bot-protected site (GFI/BGC), trying the most
    capable method available and falling back gracefully:
      1. cloudscraper — solves Cloudflare-style browser checks (if installed)
      2. warmed requests.Session — homepage visit first for cookies+referer
    Returns the response of whichever attempt got a 200 first, else the
    last response so the caller can report the real status code."""
    last_r = None
    # Attempt 1: cloudscraper (optional dependency — add 'cloudscraper' to
    # requirements.txt to enable; skipped automatically if not installed)
    try:
        import cloudscraper
        scraper = cloudscraper.create_scraper(
            browser={"browser": "chrome", "platform": "darwin", "mobile": False}
        )
        try:
            scraper.get(homepage, timeout=15)
        except Exception:
            pass
        r = scraper.get(url, timeout=30)
        if r.status_code == 200:
            return r
        last_r = r
    except ImportError:
        pass
    except Exception:
        pass
    # Attempt 2: warmed plain session
    r = _make_session(homepage).get(url, timeout=25)
    if r.status_code == 200:
        return r
    return last_r if last_r is not None else r


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_source(sef_name: str, url: str, kind: str, report_date_str: str):
    try:
        if sef_name in SITE_HOMEPAGES:
            r = _fetch_protected(url, SITE_HOMEPAGES[sef_name])
        else:
            r = requests.get(url, headers=BROWSER_HEADERS, timeout=25)
    except Exception as e:
        return None, f"Network error: {e}"
    if r.status_code == 404:
        return None, "No file for this date (market may have been closed)."
    if r.status_code == 403:
        return None, "Site blocked the request (HTTP 403) — the server may require a login or block automated access entirely."
    if r.status_code != 200:
        return None, f"HTTP {r.status_code}"
    try:
        report_date = datetime.strptime(report_date_str, "%Y%m%d")
        if kind == "tradition":
            return parse_tradition(r.text), None
        elif kind == "latam":
            return parse_latam(r.text), None
        elif kind == "gfi_bgc":
            return parse_gfi_bgc(r.content, sef_name, report_date), None
        elif kind == "icap":
            return parse_icap(r.text), None
    except Exception as e:
        return None, f"Parse error: {e}"
    return None, "Unknown source kind"


def fetch_uploaded(sef_name: str, kind: str, uploaded_file, report_date: datetime):
    try:
        if kind == "gfi_bgc":
            return parse_gfi_bgc(uploaded_file.read(), sef_name, report_date), None
        elif kind == "latam":
            return parse_latam(uploaded_file.read().decode("utf-8")), None
        elif kind == "tradition":
            return parse_tradition(uploaded_file.read().decode("utf-8")), None
        elif kind == "icap":
            return parse_icap_numbers(uploaded_file.read()), None
    except ImportError:
        return None, "Missing package — run: pip install numbers-parser"
    except Exception as e:
        return None, f"Parse error: {e}"
    return None, "Unsupported upload kind"


# Historical fetches (for month-to-date) — past days' files never change,
# so cache them for a full day instead of 30 minutes.
@st.cache_data(ttl=86400, show_spinner=False, max_entries=400)
def fetch_source_hist(sef_name: str, url: str, kind: str, report_date_str: str):
    return fetch_source(sef_name, url, kind, report_date_str)


# ═════════════════════════════════════════════════════════════════════════
# ICAP DAILY SNAPSHOTS
# ═════════════════════════════════════════════════════════════════════════
# The tpSEF page has no date in the URL — it ONLY ever shows the latest
# business day. So each day the app successfully pulls ICAP, it saves that
# day's parsed table to disk. Selecting a past date then loads the saved
# snapshot, making ICAP history available from the first day the app ran.
# NOTE: on Streamlit Community Cloud the disk is wiped whenever the app is
# redeployed or rebooted — snapshots are durable on a local machine, best-
# effort on the cloud.
import os

ICAP_SNAPSHOT_DIR = "icap_snapshots"


def _icap_snapshot_path(d: datetime) -> str:
    return os.path.join(ICAP_SNAPSHOT_DIR, f"icap_{d.strftime('%Y%m%d')}.csv")


def save_icap_snapshot(df: pd.DataFrame, d: datetime):
    try:
        os.makedirs(ICAP_SNAPSHOT_DIR, exist_ok=True)
        df.to_csv(_icap_snapshot_path(d), index=False)
    except Exception:
        pass  # snapshot saving must never break the dashboard


def load_icap_snapshot(d: datetime):
    p = _icap_snapshot_path(d)
    if not os.path.exists(p):
        return None
    try:
        df = pd.read_csv(p, dtype=str)
        for c in ["Notional_Local", "Notional_USD", "Trades", "Last_Price"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        for c in ["MXN_Match", "Trades_Estimated"]:
            if c in df.columns:
                df[c] = df[c].astype(str).map({"True": True, "False": False}).fillna(False)
        return df
    except Exception:
        return None


def get_icap_for_date(cfg: dict, sel_date: datetime):
    """Live fetch (and snapshot) when the selected date is the latest
    business day; otherwise serve the saved snapshot for that date."""
    is_latest = sel_date.date() == last_business_day().date()
    if is_latest:
        df, err = fetch_source("ICAP", cfg["url_template"], "icap", sel_date.strftime("%Y%m%d"))
        if df is not None:
            save_icap_snapshot(df, sel_date)
            return df, None
        snap = load_icap_snapshot(sel_date)
        if snap is not None:
            return snap, None
        return None, err
    snap = load_icap_snapshot(sel_date)
    if snap is not None:
        return snap, None
    return None, ("No saved ICAP snapshot for this date — the tpSEF page only shows the "
                  "latest business day, so ICAP history is available from the first day "
                  "this app saved a snapshot onward.")


# ═════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("🇲🇽 Mexico Desk")
    st.caption("Cross-SEF market comparison")
    st.markdown("---")

    auto_date = last_business_day()
    override = st.checkbox("Select a specific date")
    if override:
        sel_date = datetime.combine(st.date_input("Trade date", value=auto_date), datetime.min.time())
    else:
        sel_date = datetime.combine(auto_date, datetime.min.time())
        st.info(f"Auto date: **{auto_date.strftime('%b %d, %Y')}**")

    st.markdown("---")
    st.caption("**Sources**")
    uploaded_files = {}
    for name, cfg in SOURCES.items():
        if cfg["url_template"] is None:
            st.caption(f"⬜ {name} — no public URL on file")
            f = st.file_uploader(f"Upload {name}'s file", key=f"upload_{name}", type=["csv", "xlsx"])
            if f:
                uploaded_files[name] = f
        elif cfg["kind"] == "icap":
            st.caption(f"🟢 {name} — auto-fetch")
            f = st.file_uploader(
                f"Optional: override with a {name} .numbers export", key=f"upload_{name}", type=["numbers"]
            )
            if f:
                uploaded_files[name] = f
        else:
            st.caption(f"🟢 {name} — auto-fetch")

    st.markdown("---")
    st.caption("**FX rates (official daily fixing)**")
    fx_rates, fx_source_label, fx_note, fx_error = fetch_fx_rates(sel_date.strftime("%Y-%m-%d"))
    if fx_error:
        st.error(f"⚠️ {fx_error}")
    else:
        st.success(f"✅ {fx_source_label}")
    if fx_note:
        st.caption(f"ℹ️ {fx_note}")
    manual_overrides = {}
    with st.expander("Manual FX — CLP / COP / PEN (not published by the ECB)"):
        st.caption("These currencies have no free official daily fixing source. "
                    "Enter units of local currency per 1 USD, or leave at 0 to exclude "
                    "those trades from USD totals (they'll be flagged as missing, not guessed).")
        for ccy in EXOTIC_NO_ECB_RATE:
            v = st.number_input(f"{ccy} per 1 USD", min_value=0.0, value=0.0, step=0.01, key=f"fx_{ccy}")
            if v > 0:
                manual_overrides[ccy] = v

    st.markdown("---")
    metric_choice = st.radio("Compare by", ["DV01 (USD)", "USD Notional (converted)", "Local Notional", "Trade count"])
    if metric_choice == "DV01 (USD)":
        st.caption("DV01 ≈ USD notional × tenor in years × 1bp (flat-annuity approximation, "
                   "no discounting) — normalizes short vs. long tenors for volume comparison. "
                   "Rates products only; the FX tab falls back to USD notional.")
    show_mtd = st.checkbox("Month-to-date market share", value=True)
    if show_mtd:
        st.caption("Pulls every prior business day of the month — first load each day is slower.")
    show_exact = st.checkbox("Show exact numbers (not abbreviated M/B)", value=True)
    st.caption("Cached 30 min. Refresh the page to force an update.")

numfmt = fmt_exact if show_exact else fmt_m

# ═════════════════════════════════════════════════════════════════════════
# FETCH ALL SOURCES
# ═════════════════════════════════════════════════════════════════════════
st.title("Mexico Rates & FX — Cross-SEF Comparison")
st.caption(f"Trade date: **{sel_date.strftime('%B %d, %Y')}**")

all_data = []
status_msgs = []
with st.spinner("Pulling data from all connected SEFs..."):
    for name, cfg in SOURCES.items():
        if name in uploaded_files:
            df, err = fetch_uploaded(name, cfg["kind"], uploaded_files[name], sel_date)
            if err:
                status_msgs.append((name, "error", err))
            else:
                if cfg["kind"] == "icap":
                    save_icap_snapshot(df, sel_date)  # uploads count as that day's snapshot too
                all_data.append(df)
                status_msgs.append((name, "ok", f"{len(df):,} rows (uploaded)"))
        elif cfg["kind"] == "icap":
            df, err = get_icap_for_date(cfg, sel_date)
            if err:
                status_msgs.append((name, "error", err))
            else:
                all_data.append(df)
                status_msgs.append((name, "ok", f"{len(df):,} rows"))
        elif cfg["url_template"]:
            date_str = sel_date.strftime(cfg["date_fmt"]) if cfg["date_fmt"] else ""
            url = cfg["url_template"].replace("{date}", date_str) if cfg["date_fmt"] else cfg["url_template"]
            df, err = fetch_source(name, url, cfg["kind"], sel_date.strftime("%Y%m%d"))
            if err:
                status_msgs.append((name, "error", err))
            else:
                all_data.append(df)
                status_msgs.append((name, "ok", f"{len(df):,} rows"))
        else:
            status_msgs.append((name, "skip", "not connected"))

cols = st.columns(len(SOURCES))
for i, (name, kind, msg) in enumerate(status_msgs):
    icon = "✅" if kind == "ok" else ("⚠️" if kind == "error" else "⬜")
    cols[i].caption(f"{icon} **{name}**\n{msg}")

if not all_data:
    st.error("No data loaded from any source. Check dates, URLs, or upload files in the sidebar.")
    st.stop()

combined = pd.concat(all_data, ignore_index=True)
# Reporting bucket: sub-month tenors (ON/TN/SN, day tenors, 1-3 weeks) roll
# into '0M' for the comparison tables. The original Tenor column is kept
# untouched so the drill-down always shows the true tenor of each trade.
combined["Tenor_Bucket"] = combined["Tenor"].astype(str).apply(tenor_bucket)
combined["Tenor"] = sort_tenors(combined["Tenor"])
combined = apply_fx(combined, fx_rates, manual_overrides, fx_source_label)
combined = add_dv01(combined)

if metric_choice == "DV01 (USD)":
    metric_col = "DV01_USD"
elif metric_choice == "USD Notional (converted)":
    metric_col = "USD_Notional_Computed"
elif metric_choice == "Local Notional":
    metric_col = "Notional_Local"
else:
    metric_col = "Trades"

missing_fx = combined[combined["USD_Notional_Computed"].isna() & (combined["Currency"] != "N/A")]
if not missing_fx.empty and metric_col in ("USD_Notional_Computed", "DV01_USD"):
    miss_ccys = sorted(missing_fx["Currency"].unique())
    st.warning(
        f"⚠️ No FX rate available for **{', '.join(miss_ccys)}** — "
        f"{len(missing_fx)} row(s) totaling {numfmt(missing_fx['Notional_Local'].sum())} (local ccy) "
        f"are excluded from USD totals below. Add a manual rate in the sidebar to include them."
    )

# ═════════════════════════════════════════════════════════════════════════
# COMPARISON TABLE BUILDER
# ═════════════════════════════════════════════════════════════════════════
def build_comparison(data: pd.DataFrame, sefs_present: list, mcol: str, category_label: str) -> pd.DataFrame:
    """
    One row per tenor with activity, in ladder order.
    - IRS: ONLY the desk's standard list (1M,2M,3M,6M,9M, 1Y-10Y, 12Y,15Y,
      20Y,25Y,30Y). Anything that traded at another tenor rolls into a
      single 'Other / non-standard tenor' row so totals still reconcile.
    - FX: the full ladder (0M roll-up, months, years) plus extras.
    Tenors with NO activity that day are hidden entirely.
    Every trade lands in exactly one row, so the Total column reconciles
    with the market-share chart.
    """
    data = data.copy()
    if category_label == "IRS / Swap":
        ladder = list(IRS_DISPLAY_TENORS)
        data["Disp_Bucket"] = data["Tenor_Bucket"].where(
            data["Tenor_Bucket"].isin(ladder), "Other")
        other_label = "Other / non-standard tenor"
    else:
        present = [t for t in data["Tenor_Bucket"].unique() if t != "Other"]
        extras = sorted([t for t in present if t not in FULL_TENOR_LADDER], key=tenor_sort_key)
        ladder = FULL_TENOR_LADDER + extras
        data["Disp_Bucket"] = data["Tenor_Bucket"]
        other_label = "Other / unmatched"
    if (data["Disp_Bucket"] == "Other").any():
        ladder = ladder + ["Other"]
    rows = []
    for tenor in ladder:
        t_df = data[data["Disp_Bucket"] == tenor]
        if t_df.empty:
            continue  # hide tenors with no trades that day
        label = other_label if tenor == "Other" else tenor_display(tenor)
        row = {"Tenor": label}
        total = 0
        for sef in sefs_present:
            v = t_df[t_df["SEF"] == sef][mcol].sum()
            row[sef] = v
            total += v if not pd.isna(v) else 0
        if total == 0:
            continue  # nothing measurable in this row for the chosen metric
        row["Total"] = total
        for sef in sefs_present:
            row[f"{sef} %"] = round(row[sef] / total * 100, 1) if total > 0 else 0
        rows.append(row)
    return pd.DataFrame(rows)


def render_market_share_chart(cat_data: pd.DataFrame, sefs_present: list, category_label: str, currency_label: str,
                              mcol: str, chart_key: str):
    totals = cat_data.groupby("SEF")[mcol].sum().reindex(sefs_present).fillna(0)
    grand_total = totals.sum()
    if grand_total <= 0:
        return
    shares = (totals / grand_total * 100).round(1)
    st.markdown(f"**Market share % — {category_label} ({currency_label})**")
    pie_df = pd.DataFrame({"SEF": totals.index, "Value": totals.values})
    color_map = {s: SEF_COLORS.get(s, "#999999") for s in sefs_present}
    fig = px.pie(
        pie_df, names="SEF", values="Value",
        color="SEF", color_discrete_map=color_map, hole=0.35,
    )
    fig.update_traces(textposition="inside", textinfo="percent+label", sort=False)
    fig.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        height=max(320, 55 * len(sefs_present)),
        showlegend=True,
    )
    st.plotly_chart(fig, use_container_width=True, key=chart_key)
    legend_line = "  ·  ".join(f"**{s}**: {shares[s]}% ({numfmt(totals[s])})" for s in sefs_present)
    st.caption(legend_line)


def render_currency_breakdown(cat_data: pd.DataFrame):
    st.markdown("**By original currency → USD converted**")
    g = cat_data.groupby("Currency").agg(
        Trades=("Trades", "sum"),
        Original_Notional=("Notional_Local", "sum"),
        USD_Notional=("USD_Notional_Computed", "sum"),
    ).reset_index()
    g = g.sort_values("USD_Notional", ascending=False, na_position="last")
    fx_lookup = cat_data.groupby("Currency")["FX_Rate_Used"].first()
    src_lookup = cat_data.groupby("Currency")["FX_Source"].first()
    g["FX Rate (CCY per USD)"] = g["Currency"].map(fx_lookup)
    g["FX Source"] = g["Currency"].map(src_lookup)
    st.dataframe(
        g.rename(columns={"Original_Notional": "Original Notional", "USD_Notional": "USD Notional"}).style.format({
            "Original Notional": numfmt, "USD Notional": numfmt,
            "FX Rate (CCY per USD)": lambda x: f"{x:,.4f}" if pd.notna(x) else "—",
            "Trades": "{:,.0f}",
        }),
        use_container_width=True, hide_index=True,
    )


def render_fx_crosscheck(cat_data: pd.DataFrame, sefs_present: list):
    rows = []
    for sef in sefs_present:
        sef_df = cat_data[cat_data["SEF"] == sef]
        published = sef_df["Notional_USD"].dropna()
        if published.empty:
            continue
        pub_total = published.sum()
        calc_total = sef_df.loc[published.index, "USD_Notional_Computed"].sum()
        diff_pct = abs(pub_total - calc_total) / pub_total * 100 if pub_total else 0
        status = "✓ Match" if diff_pct < 1.0 else f"⚠ {diff_pct:.1f}% diff"
        rows.append({"SEF": sef, "Published USD (by source)": pub_total,
                      "Calculated USD (our FX)": calc_total, "Status": status})
    if rows:
        st.markdown("**FX cross-check** — vs. sources that publish their own USD-converted figure")
        cc_df = pd.DataFrame(rows)
        st.dataframe(
            cc_df.style.format({"Published USD (by source)": numfmt, "Calculated USD (our FX)": numfmt}),
            use_container_width=True, hide_index=True,
        )


def render_section(data: pd.DataFrame, category_label: str, sefs_present: list, group_label: str,
                   mtd_data=None):
    cat_data = data[data["Category"] == category_label]
    if cat_data.empty:
        st.info(f"No {category_label} data for {group_label}.")
        return

    # DV01 is a rates concept — the FX tab falls back to USD notional.
    mcol, mlabel = metric_col, metric_choice
    if category_label == "FX" and metric_col == "DV01_USD":
        mcol, mlabel = "USD_Notional_Computed", "USD Notional (converted)"
        st.caption("ℹ️ DV01 doesn't apply to FX products — this tab compares by USD notional instead.")

    st.markdown(f"#### {category_label}")
    if category_label == "IRS / Swap":
        st.caption(
            "ℹ️ Mexican TIIE swaps are quoted by number of 28-day periods, not calendar time — "
            "the market counts 13 periods (364 days) as \"1 year,\" 26 as \"2 years,\" etc. "
            "Tenors below are converted to standard Y/M labels so every source lines up on the "
            "same row — verified by price match (e.g. LatAm's 13x1 and Tradition's \"1 Year\" "
            "both quoted 6.735 the same day)."
        )
        if mcol == "DV01_USD":
            no_dv01 = cat_data[cat_data["DV01_USD"].isna()]
            if not no_dv01.empty:
                st.caption(f"ℹ️ {len(no_dv01)} row(s) have no parseable single tenor (spread trades etc.) — "
                           f"no DV01 can be computed for them, so they're excluded from DV01 comparisons.")

    # KPI row — total per SEF, in the currently-selected metric
    kpi_cols = st.columns(len(sefs_present))
    totals = cat_data.groupby("SEF")[mcol].sum()
    grand_total = totals.sum()
    for i, sef in enumerate(sefs_present):
        v = totals.get(sef, 0)
        if mlabel == "Trade count":
            unit = ""
        elif mlabel.startswith("DV01"):
            unit = " USD DV01"
        elif mlabel.startswith("USD"):
            unit = " USD"
        else:
            unit = " (local ccy)"
        share = round(v / grand_total * 100, 1) if grand_total > 0 else 0
        kpi_cols[i].metric(sef, f"{share}%", f"{numfmt(v)}{unit}")

    render_market_share_chart(cat_data, sefs_present, category_label, f"{group_label} — today",
                              mcol, chart_key=f"pie_day_{category_label}")

    # Month-to-date market share
    if mtd_data is not None:
        mtd_cat = mtd_data[mtd_data["Category"] == category_label]
        if not mtd_cat.empty:
            mtd_mcol = mcol if mcol in mtd_cat.columns else "USD_Notional_Computed"
            days_covered = sorted(mtd_cat["Date"].unique())
            st.markdown("")
            render_market_share_chart(
                mtd_cat, sefs_present, category_label,
                f"Month-to-date · {days_covered[0]} → {days_covered[-1]} · {len(days_covered)} day(s)",
                mtd_mcol, chart_key=f"pie_mtd_{category_label}")
            if "ICAP" not in mtd_cat["SEF"].unique():
                st.caption("ℹ️ ICAP joins the month-to-date view as daily snapshots accumulate "
                           "(the tpSEF page only ever shows the latest day).")

    st.markdown("")
    render_currency_breakdown(cat_data)

    st.markdown("")
    render_fx_crosscheck(cat_data, sefs_present)

    # Tenor-ladder comparison table (traded tenors only)
    st.markdown("")
    comp = build_comparison(cat_data, sefs_present, mcol, category_label)
    if comp.empty:
        st.info("No rows with a measurable value for the chosen metric.")
        return
    display_cols = ["Tenor"] + sefs_present + ["Total"] + [f"{s} %" for s in sefs_present]
    if category_label == "IRS / Swap":
        ladder_note = ("IRS ladder: 1M, 2M, 3M, 6M, 9M, 1Y-10Y, 12Y, 15Y, 20Y, 25Y, 30Y — "
                       "only tenors that actually traded are shown; anything at a non-standard "
                       "tenor rolls into the 'Other / non-standard tenor' row so totals reconcile")
    else:
        ladder_note = ("FX ladder: 0M (1W/2W/3W), months, years — only tenors that actually "
                       "traded are shown; a final 'Other / unmatched' row catches anything "
                       "the parser couldn't bucket")
    st.caption(f"{ladder_note} — {len(comp)} row(s) for {category_label} in {group_label} — "
               f"values in **{mlabel}**")
    st.dataframe(
        comp[display_cols].style.format({**{s: numfmt for s in sefs_present}, "Total": numfmt,
                                          **{f"{s} %": "{:.1f}%" for s in sefs_present}}),
        use_container_width=True, hide_index=True,
        height=min(35 * (len(comp) + 1) + 3, 900),
    )

    # Reconciliation check — the ladder rows must account for every trade
    # with a measurable value for the chosen metric.
    table_total = comp["Total"].sum()
    chart_total = grand_total if not pd.isna(grand_total) else 0
    if chart_total > 0 and abs(table_total - chart_total) / chart_total > 0.001:
        st.error(
            f"⚠️ Reconciliation failure: the tenor table totals {numfmt(table_total)} but the "
            f"market-share chart totals {numfmt(chart_total)} — some trades are being dropped "
            f"between aggregation steps. This should never happen; check Tenor_Bucket assignment."
        )
    else:
        st.caption(f"✓ Reconciled: tenor table total ({numfmt(table_total)}) matches the market-share chart total.")

    proxy_sefs = [s for s in sefs_present if cat_data[cat_data["SEF"] == s]["Trades_Estimated"].any()]
    if proxy_sefs and mlabel == "Trade count":
        st.caption(f"⚠️ Trade counts for {', '.join(proxy_sefs)} are a proxy (row count) — "
                   f"these sources don't publish a real trade-count field.")

    # ── Individual trades — filtered to THIS category only ────────────────
    st.markdown("")
    with st.expander(f"🔍 Individual {category_label} trades"):
        dcol1, dcol2 = st.columns(2)
        with dcol1:
            drill_sef = st.selectbox("Source", sorted(cat_data["SEF"].unique()),
                                     key=f"drill_sef_{category_label}")
        sef_df = cat_data[cat_data["SEF"] == drill_sef]
        with dcol2:
            bucket_opts = sorted(sef_df["Tenor_Bucket"].unique(), key=tenor_sort_key)
            true_opts = sorted(sef_df["Tenor"].astype(str).unique(), key=tenor_sort_key)
            sef_tenors = ["All tenors"] + list(dict.fromkeys(bucket_opts + true_opts))
            drill_tenor = st.selectbox("Tenor", sef_tenors, key=f"drill_tenor_{category_label}",
                                       format_func=tenor_display)
        if drill_tenor == "All tenors":
            detail = sef_df
        else:
            detail = sef_df[(sef_df["Tenor_Bucket"] == drill_tenor) |
                            (sef_df["Tenor"].astype(str) == drill_tenor)]
        show_cols = ["Description", "AssetClass", "Currency", "Tenor", "Last_Price", "Notional_Local",
                     "USD_Notional_Computed", "DV01_USD", "FX_Rate_Used", "FX_Source", "FX_CrossCheck", "Trades"]
        if category_label == "FX":
            show_cols.remove("DV01_USD")  # DV01 not meaningful for FX
        show_cols = [c for c in show_cols if c in detail.columns]
        st.dataframe(
            detail[show_cols].rename(columns={
                "AssetClass": "Asset Class", "Tenor": "True Tenor",
                "Notional_Local": "Original Notional", "USD_Notional_Computed": "USD Notional",
                "DV01_USD": "DV01 (USD)",
                "FX_Rate_Used": "FX Rate (CCY/USD)", "FX_Source": "FX Source", "FX_CrossCheck": "Cross-Check",
            }),
            use_container_width=True, hide_index=True,
        )
        summary = (f"{len(detail)} row(s) · Total original notional: {numfmt(detail['Notional_Local'].sum())} "
                   f"· Total USD notional: {numfmt(detail['USD_Notional_Computed'].sum())} ")
        if category_label == "IRS / Swap" and "DV01_USD" in detail.columns:
            summary += f"· Total DV01: {numfmt(detail['DV01_USD'].sum())} USD "
        summary += f"· Total trades: {int(detail['Trades'].sum())}"
        st.caption(summary)
        st.caption("Asset Class codes: **IR** = interest rates · **CU** = currency/FX · **CD** = credit "
                   "(taken from the source file where published — GFI, BGC, ICAP — and derived from the "
                   "product category for Tradition and LatAm, which don't publish a code).")
        if detail["Trades_Estimated"].any():
            st.caption("⚠️ This source doesn't publish individual trade prints — each row above is already a daily aggregate.")


# ═════════════════════════════════════════════════════════════════════════
# MONTH-TO-DATE DATA (Mexico products only)
# ═════════════════════════════════════════════════════════════════════════
def _business_days_of_month(sel: datetime):
    d = sel.replace(day=1)
    days = []
    while d.date() <= sel.date():
        if d.weekday() < 5:
            days.append(d)
        d += timedelta(days=1)
    return days


@st.cache_data(ttl=1800, show_spinner=False)
def build_mtd_data(sel_date_str: str, manual_overrides: dict):
    """Fetch every business day of the month up to (and including) the
    selected date, keep only Mexico products, convert each day with THAT
    day's ECB fixing, and compute DV01. Days a source has no file for
    (holidays, blocked fetches, missing ICAP snapshots) are simply skipped
    for that source. Historical files are cached for 24h."""
    sel = datetime.strptime(sel_date_str, "%Y%m%d")
    frames = []
    for d in _business_days_of_month(sel):
        ds = d.strftime("%Y%m%d")
        day_frames = []
        for name, cfg in SOURCES.items():
            if cfg["kind"] == "icap":
                snap = load_icap_snapshot(d)
                if snap is not None:
                    day_frames.append(snap)
                continue
            date_str = d.strftime(cfg["date_fmt"])
            url = cfg["url_template"].replace("{date}", date_str)
            df, err = fetch_source_hist(name, url, cfg["kind"], ds)
            if df is not None:
                day_frames.append(df)
        if not day_frames:
            continue
        day = pd.concat(day_frames, ignore_index=True)
        day = day[day["MXN_Match"]]
        if day.empty:
            continue
        rates, lbl, _, _ = fetch_fx_rates(d.strftime("%Y-%m-%d"))
        day = apply_fx(day, rates, manual_overrides, lbl)
        day = add_dv01(day)
        day["Date"] = d.strftime("%Y-%m-%d")
        frames.append(day)
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


mtd_data = None
if show_mtd:
    with st.spinner("Building month-to-date view (pulls each prior business day — cached after first load)..."):
        try:
            mtd_data = build_mtd_data(sel_date.strftime("%Y%m%d"), manual_overrides)
        except Exception as e:
            st.warning(f"Month-to-date view unavailable: {e}")
            mtd_data = None


# ═════════════════════════════════════════════════════════════════════════
# MEXICO PRODUCTS — IRS / FX
# ═════════════════════════════════════════════════════════════════════════
st.markdown("---")


def get_mexico_slice(df):
    return df[df["MXN_Match"]]


mexico_df = get_mexico_slice(combined)
if mexico_df.empty:
    st.error("No Mexico products found in any source for this date.")
    st.stop()

sefs_present = sorted(mexico_df["SEF"].unique(), key=lambda s: list(SOURCES.keys()).index(s))
sub_irs, sub_fx = st.tabs(["📈 IRS / Swap", "💱 FX"])
with sub_irs:
    render_section(mexico_df, "IRS / Swap", sefs_present, "Mexico Products", mtd_data=mtd_data)
with sub_fx:
    render_section(mexico_df, "FX", sefs_present, "Mexico Products", mtd_data=mtd_data)

# ═════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.caption(
    "DV01 (USD) ≈ USD notional × tenor in years × 1bp — a flat-annuity approximation with no "
    "discounting, used to normalize volumes across the curve; not a risk number. "
    "USD Notional = Original Notional ÷ (official ECB daily reference rate for the currency, "
    "as of the trade date, via Frankfurter.app). USD figures a source publishes directly "
    "(Tradition, LatAm) are shown alongside our computed figure so any mismatch is visible "
    "immediately rather than hidden inside a single blended number. Tradition notionals are "
    "the GROSS (Non-Delta-Adjusted) figures from their file — their Delta-Adjusted columns "
    "apply option deltas as whole numbers (50x instead of 0.50x) and aren't comparable to "
    "the gross notionals other SEFs publish. CLP/COP/PEN require a "
    "manually-entered rate (sidebar) since the ECB doesn't publish fixings for them — those "
    "trades are excluded from USD totals, never estimated, until a rate is supplied. "
    "ICAP history: the tpSEF page only ever shows the latest business day, so the app saves a "
    "snapshot each day it runs — past dates load from those snapshots (on Streamlit Cloud, "
    "snapshots persist until the app is redeployed or rebooted)."
)
