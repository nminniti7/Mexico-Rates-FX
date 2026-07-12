"""
Mexico Rates & FX Desk — Cross-SEF Market Comparison Dashboard
=================================================================
Tradition · LatAm SEF · GFI · BGC · ICAP (tpSEF)

WHAT'S NEW IN THIS VERSION
-------------------------
1. GFI/BGC 403 FIX (RE-APPLIED). requests and cloudscraper both send a
   Python TLS fingerprint that the GFI/BGC WAF rejects with HTTP 403 no
   matter what headers you fake. The fix is curl_cffi with Chrome browser
   impersonation, which presents a real Chrome TLS fingerprint. Order is
   now: curl_cffi → cloudscraper → warmed requests session. Add BOTH
   `curl_cffi` and `cloudscraper` to requirements.txt or Streamlit Cloud
   won't install them and you're back to 403.
2. ICAP DATES NOW ACTUALLY SAVE. Two bugs, both fixed:
   a) The GitHub sync ran AFTER the sidebar rendered, so the sidebar
      counted snapshots before history was pulled down — a fresh container
      always said "1 day saved" even when the repo had more. Sync now runs
      FIRST, before anything renders.
   b) Snapshots were saved to local disk but only pushed to GitHub in one
      code path. Streamlit Cloud wipes local disk on every redeploy, so
      any snapshot that missed the push was gone forever. Every save point
      (daily auto-capture, today's fetch, sidebar upload, backfill) now
      pushes its file to the repo immediately.
   The sidebar also now says out loud whether GitHub sync is configured —
   no more silent data loss.
3. HISTORY IS REAL. Two things are persisted per day, one file per day,
   never combined:
       icap_snapshots/icap_YYYYMMDD.csv   — that day's ICAP table
       dv01_grids/dv01_YYYYMMDD.csv       — that day's DV01 grid
   Pick a past date and you get THAT day's ICAP data valued on THAT day's
   grid. If a grid was never edited for a given day, the last saved grid is
   carried forward (and if there's none, the built-in default) — always
   labelled on screen so you know which you're looking at.
4. GITHUB SYNC, SILENT. Streamlit Cloud wipes local disk on redeploy, so the
   app PULLS saved history from the repo at startup and PUSHES new days back.
   Config lives ONLY in .streamlit/secrets.toml under [github] — there's no
   panel and nothing to type. No secrets = local disk only, everything else
   still works (and the sidebar warns you).
5. BACKFILL. tpSEF only ever shows the latest business day, so past ICAP days
   can't be re-fetched — but you can upload an export and file it under
   any date (sidebar → ICAP history → Backfill).
6. NO FX CONVERSION ANYWHERE. Notional is whatever the source file says, in
   the currency it was traded in, so it always ties to the raw file. Metrics
   are DV01 and notional-as-published. One trades drill-down, not two.
7. DV01 grid can be PASTED in whole (sidebar) — two columns out of Excel or
   the desk email, '13x1' period notation and all.

Run:      streamlit run theNiche.py
Install:  pip install streamlit pandas requests openpyxl beautifulsoup4 lxml plotly curl_cffi cloudscraper
"""

import streamlit as st
import pandas as pd
import requests
import re
import io
import os
import base64
import plotly.express as px
from datetime import datetime, timedelta

st.set_page_config(page_title="Mexico Desk — Cross-SEF Dashboard", page_icon="🇲🇽", layout="wide")

# ═════════════════════════════════════════════════════════════════════════
# CONFIG — SOURCES
# ═════════════════════════════════════════════════════════════════════════
SOURCES = {
    "Tradition": {
        "kind": "tradition",
        "url_templates": [
            "https://www.traditionsef.com/dailyactivity/SEF16_MKTDATA_TFSU_{date}.csv",
        ],
        "date_fmt": "%Y%m%d",
    },
    "LatAm SEF": {
        "kind": "latam",
        "url_templates": [
            "http://latamsef.com/market-data/LatAmSEF_MarketActivityData_{date}.csv",     # direct link — works
            "https://www.latamsef.com/market-data/LatAmSEF_MarketActivityData_{date}.csv",
        ],
        "date_fmt": "%Y%m%d",
    },
    "GFI": {
        "kind": "gfi_bgc",
        "url_templates": [
            "http://www.gfigroup.com/doc/sef/marketdata/{date}_daily_trade_data.xlsx",    # direct link — works
            "https://www.gfigroup.com/doc/sef/marketdata/{date}_daily_trade_data.xlsx",
        ],
        "date_fmt": "%Y-%m-%d",
    },
    "BGC": {
        "kind": "gfi_bgc",
        "url_templates": [
            "https://www.bgcsef.com/TradingActivityReports/Daily/DailyAct_{date}.xlsx",   # direct link — works
            "http://www.bgcsef.com/TradingActivityReports/Daily/DailyAct_{date}.xlsx",
        ],
        "date_fmt": "%Y%m%d",
    },
    "ICAP": {
        "kind": "icap",
        "url_templates": [
            "https://www.tullettprebon.com/swap-execution-facility/daily-activity-summary.aspx",
        ],
        "date_fmt": None,       # no date in URL — page always shows latest business day
    },
}


def source_urls(cfg: dict, d: datetime) -> tuple:
    """All candidate URLs for a source on a date — direct link first, then
    scheme/www variants. Returned as a tuple so it's hashable for caching."""
    tpls = cfg["url_templates"]
    if cfg["date_fmt"]:
        ds = d.strftime(cfg["date_fmt"])
        return tuple(t.replace("{date}", ds) for t in tpls)
    return tuple(tpls)

SEF_COLORS = {
    "Tradition": "#1a56db",
    "LatAm SEF": "#0e9f6e",
    "GFI":       "#d97706",
    "BGC":       "#7e3af2",
    "ICAP":      "#e02424",
}

TENOR_ORDER = [
    "ON", "TN", "SN",
    "0M",
    "1W", "2W", "3W",
    "1M", "2M", "3M", "4M", "5M", "6M", "7M", "8M", "9M", "10M", "11M",
    "1Y", "18M", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y",
    "10Y", "12Y", "15Y", "20Y", "25Y", "30Y", "Other"
]

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

# The desk's standard IRS points. Anything traded elsewhere rolls into one
# 'Other / non-standard tenor' row so totals still reconcile.
IRS_DISPLAY_TENORS = (
    ["1M", "2M", "3M", "6M", "9M"] +
    [f"{y}Y" for y in range(1, 11)] +
    ["12Y", "15Y", "20Y", "25Y", "30Y"]
)

SUB_MONTH_TENORS = {"1W", "2W", "3W"}

# MXN TIIE swaps are quoted in 28-day periods (13 periods = "1 year").
LATAM_PERIOD_TENOR = {
    1: "1M", 3: "3M", 6: "6M", 9: "9M", 13: "1Y", 19: "18M", 26: "2Y",
    39: "3Y", 52: "4Y", 65: "5Y", 78: "6Y", 91: "7Y", 104: "8Y",
    117: "9Y", 130: "10Y", 195: "15Y", 260: "20Y", 390: "30Y",
}

# ═════════════════════════════════════════════════════════════════════════
# CONFIG — PERSISTENCE (one file per day, never combined)
# ═════════════════════════════════════════════════════════════════════════
ICAP_SNAPSHOT_DIR = "icap_snapshots"
SEF_SNAPSHOT_DIR = "sef_snapshots"      # Tradition/LatAm/GFI/BGC daily saves — MTD reads these
DV01_GRID_DIR = "dv01_grids"
LEGACY_DV01_FILE = "dv01_grid.csv"      # old single-grid file — still read as a fallback
HISTORY_START = datetime(2026, 7, 9)    # first day the desk wanted history from

DV01_GRID_BASE = 5000.0   # grid = MXN millions of notional per THIS many USD of DV01
DEFAULT_DV01_GRID = {
    "1M": 11293.90, "2M": 5661.80, "3M": 3784.42, "4M": 2845.76,
    "5M": 2282.58,  "6M": 1907.15, "7M": 1639.00, "8M": 1437.90,
    "9M": 1281.51,  "10M": 1156.40, "11M": 1054.05,
    "1Y": 896.60,   "18M": 623.41,
    "2Y": 464.53, "3Y": 321.45, "4Y": 250.47, "5Y": 208.21,
    "6Y": 180.31, "7Y": 160.64, "8Y": 146.06, "9Y": 134.96,
    "10Y": 126.22, "12Y": 113.51, "15Y": 101.47, "20Y": 90.64,
    "25Y": 84.89, "30Y": 81.50,
}



# ═════════════════════════════════════════════════════════════════════════
# SHARED HELPERS
# ═════════════════════════════════════════════════════════════════════════
def safe_sorted(values, key=None):
    """Sort anything without blowing up on mixed types. A blank cell in any
    source file comes back from pandas as float NaN — sorting that against
    real string values raises TypeError ('<' not supported between float and
    str), which is exactly what was killing the page. Coerce to str first."""
    vals = [("" if v is None else str(v)) for v in values]
    return sorted(vals, key=key) if key else sorted(vals)


def safe_unique(series) -> list:
    return safe_sorted(pd.Series(series).dropna().unique())


def last_business_day() -> datetime:
    d = datetime.today() - timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


def fmt_m(n):
    if n is None or pd.isna(n):
        return "—"
    if abs(n) >= 1e9:
        return f"{n/1e9:,.3f}B"
    if abs(n) >= 1e6:
        return f"{n/1e6:,.2f}M"
    return f"{n:,.0f}"


def fmt_exact(n):
    if n is None or pd.isna(n):
        return "—"
    return f"{n:,.0f}"


def tenor_sort_key(t):
    t = str(t)
    try:
        return TENOR_ORDER.index(t)
    except ValueError:
        return len(TENOR_ORDER)


def sort_tenors(series: pd.Series) -> pd.Categorical:
    vals = [str(v) for v in series.unique().tolist()]
    known = [t for t in TENOR_ORDER if t in vals]
    other = safe_sorted([t for t in vals if t not in known])
    return pd.Categorical(series.astype(str), categories=known + other, ordered=True)


def tenor_display(t: str) -> str:
    return TENOR_DISPLAY_LABELS.get(str(t), str(t))


def days_to_tenor_bucket(days: int) -> str:
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


def tenor_bucket(t) -> str:
    """Reporting bucket. Weeks (1W/2W/3W) and sub-month day counts -> '0M';
    longer day counts ('182D', '364D') -> their REAL month/year bucket;
    ON/TN/SN keep their own row."""
    if not isinstance(t, str):
        return "Other"
    original = t.strip()
    up = original.upper()
    if up in ("ON", "TN", "SN"):
        return up
    if up in SUB_MONTH_TENORS:
        return "0M"
    m = re.fullmatch(r'(\d+)D', up)
    if m:
        b = days_to_tenor_bucket(int(m.group(1)))
        return "0M" if b in ("ON", "TN", "1W", "2W", "3W") else b
    m = re.fullmatch(r'(\d+)W', up)
    if m and int(m.group(1)) <= 3:
        return "0M"
    return original


def tenor_to_years(t):
    if not isinstance(t, str):
        return None
    up = t.strip().upper()
    if up in ("ON", "TN", "SN"):
        return 1 / 365
    for pat, div in ((r'(\d+)Y', 1), (r'(\d+)M', 12), (r'(\d+)W', 52), (r'(\d+)D', 365)):
        m = re.fullmatch(pat, up)
        if m:
            return float(m.group(1)) / div
    return None


# ═════════════════════════════════════════════════════════════════════════
# DV01 — per-day grids
# ═════════════════════════════════════════════════════════════════════════
def _dv01_path(d: datetime) -> str:
    return os.path.join(DV01_GRID_DIR, f"dv01_{d.strftime('%Y%m%d')}.csv")


def _read_grid_csv(path: str):
    try:
        df = pd.read_csv(path)
        g = {}
        for _, r in df.iterrows():
            t = r.get("Tenor")
            v = pd.to_numeric(r.get("MXN_mm_per_5k_DV01"), errors="coerce")
            if isinstance(t, str) and pd.notna(v) and v > 0:
                g[t.strip().upper()] = float(v)
        return g or None
    except Exception:
        return None


def list_dv01_grid_dates():
    if not os.path.isdir(DV01_GRID_DIR):
        return []
    out = []
    for f in sorted(os.listdir(DV01_GRID_DIR)):
        m = re.fullmatch(r'dv01_(\d{8})\.csv', f)
        if m:
            out.append(datetime.strptime(m.group(1), "%Y%m%d"))
    return sorted(out)


def load_dv01_grid_for_date(d: datetime):
    """(grid, label). Exact day if saved → else last grid saved BEFORE that
    day (carried forward) → else the built-in default. Never guesses silently:
    the label is printed on screen."""
    p = _dv01_path(d)
    if os.path.exists(p):
        g = _read_grid_csv(p)
        if g:
            return g, f"saved for {d:%b %d}"
    prior = [x for x in list_dv01_grid_dates() if x.date() < d.date()]
    if prior:
        g = _read_grid_csv(_dv01_path(prior[-1]))
        if g:
            return g, f"carried forward from {prior[-1]:%b %d} (not edited on {d:%b %d})"
    if os.path.exists(LEGACY_DV01_FILE):
        g = _read_grid_csv(LEGACY_DV01_FILE)
        if g:
            return g, "default grid"
    return dict(DEFAULT_DV01_GRID), "default grid (never edited)"


def save_dv01_grid_for_date(g: dict, d: datetime) -> bool:
    try:
        os.makedirs(DV01_GRID_DIR, exist_ok=True)
        pd.DataFrame({"Tenor": list(g.keys()), "MXN_mm_per_5k_DV01": list(g.values())}) \
            .to_csv(_dv01_path(d), index=False)
        return True
    except Exception:
        return False


def ensure_todays_dv01_grid():
    """Every business day gets its own grid file, even if nobody touches it —
    the carried-forward/default values are written down so a past date always
    values on exactly what was in force that day. Pushed to GitHub on write."""
    d = last_business_day()
    if not os.path.exists(_dv01_path(d)):
        g, _ = load_dv01_grid_for_date(d)
        if save_dv01_grid_for_date(g, d):
            push_day_file(_dv01_path(d), DV01_GRID_DIR)


def _norm_tenor_label(tok: str):
    """Turn whatever the desk writes into a standard label.
    '13x1' -> 1Y (28-day periods) · '10 Year' -> 10Y · '6M' -> 6M · '18m' -> 18M"""
    t = str(tok).strip().upper().replace(" ", "")
    m = re.fullmatch(r'(\d+)X\d+', t)                 # period notation: 13x1, 26x1...
    if m:
        p = int(m.group(1))
        if p in LATAM_PERIOD_TENOR:
            return LATAM_PERIOD_TENOR[p]
        yrs = p / 13.0
        return f"{round(yrs)}Y" if yrs >= 1 else f"{round(p * 28 / 30)}M"
    m = re.fullmatch(r'(\d+)(Y|YR|YRS|YEAR|YEARS)', t)
    if m:
        return f"{m.group(1)}Y"
    m = re.fullmatch(r'(\d+)(M|MO|MOS|MONTH|MONTHS)', t)
    if m:
        return f"{m.group(1)}M"
    return None


def parse_pasted_grid(text: str, fallback_order: list):
    """Read a grid pasted straight out of Excel or the desk's email.

    Accepts, in any mix:
      · '10Y<tab>126.22'  ·  '13x1  896.60'  ·  '6 Month, 1,907.15'  ·  '30Y $81.50'
      · a bare column of numbers, in ladder order (mapped onto fallback_order)
    Header rows, blank lines, commas, $ and % are all tolerated.
    Returns (grid_dict, n_matched, unrecognised_lines)."""
    grid, loose_values, bad = {}, [], []
    # a tenor written any way the desk writes it: 10Y · 6M · 13x1 · 6 Month · 2 Years
    TENOR_RE = re.compile(r'\b(\d+)\s*(?:X\s*\d+|YEARS?|YRS?|Y|MONTHS?|MOS?|M)\b')

    for raw_line in str(text).strip().splitlines():
        line = raw_line.strip()
        if not line:
            continue
        # kill thousands separators FIRST — otherwise '11,293.90' splits into two numbers
        line = re.sub(r'(?<=\d),(?=\d{3}(?:\D|$))', '', line)
        line = line.replace("$", "").replace("%", "")

        up = line.upper()
        label, rest = None, up
        m = TENOR_RE.search(up)
        if m:
            label = _norm_tenor_label(m.group(0).replace(" ", ""))
            if label:
                rest = up[:m.start()] + " " + up[m.end():]   # so the tenor's digits aren't read as the value

        nums = re.findall(r'-?\d+(?:\.\d+)?', rest)
        if not nums:
            bad.append(raw_line.strip())      # header row or junk
            continue
        value = float(nums[-1])

        if label:
            grid[label] = value
        else:
            loose_values.append(value)

    # A bare column of numbers: map onto the existing ladder in order.
    if not grid and loose_values:
        for tenor, v in zip(fallback_order, loose_values):
            grid[tenor] = v
    grid = {k: v for k, v in grid.items() if v > 0}
    return grid, len(grid), bad


def _grid_year_points(grid: dict):
    pts = []
    for label, v in grid.items():
        y = tenor_to_years(str(label))
        if y is not None and v and v > 0:
            pts.append((y, DV01_GRID_BASE / float(v)))
    pts.sort()
    dedup = []
    for y, d in pts:
        if not dedup or y > dedup[-1][0]:
            dedup.append((y, d))
    return dedup


def dv01_per_million_mxn(bucket: str, grid: dict, pts: list):
    if bucket in grid and grid[bucket] and grid[bucket] > 0:
        return DV01_GRID_BASE / float(grid[bucket]), "desk grid"
    y = tenor_to_years(bucket)
    if y is None or not pts:
        return None, None
    if y <= pts[0][0]:
        return pts[0][1] * (y / pts[0][0]), "interpolated"
    if y >= pts[-1][0]:
        return pts[-1][1], "interpolated"
    for (y0, d0), (y1, d1) in zip(pts, pts[1:]):
        if y0 <= y <= y1:
            w = (y - y0) / (y1 - y0)
            return d0 + w * (d1 - d0), "interpolated"
    return None, None


def add_dv01(df: pd.DataFrame, grid: dict) -> pd.DataFrame:
    """DV01 (USD) = MXN notional (millions) / grid[tenor] × 5,000.

    MXN rates rows only. Nothing is converted anywhere in this app — the grid
    IS the MXN→USD-DV01 bridge and it's the desk's own number. A non-MXN rates
    row (SOFR etc.) or a spread trade with no single tenor simply gets no DV01
    rather than an invented one."""
    df = df.copy()
    pts = _grid_year_points(grid)
    buckets = df["Tenor"].astype(str).apply(tenor_bucket)
    df["DV01_USD"] = pd.NA
    df["DV01_Method"] = ""
    is_irs = df["Category"] == "IRS / Swap"
    local = pd.to_numeric(df["Notional_Local"], errors="coerce")
    for idx in df.index[is_irs]:
        if str(df.at[idx, "Currency"]).upper() != "MXN" or pd.isna(local.loc[idx]):
            continue
        per_mm, method = dv01_per_million_mxn(buckets.loc[idx], grid, pts)
        if per_mm is not None:
            df.at[idx, "DV01_USD"] = local.loc[idx] / 1e6 * per_mm
            df.at[idx, "DV01_Method"] = method
    return df


def tiie_pv01(notional_mxn: float, rate_pct: float, days: int):
    """PV01 the way the desk's Excel does it: annuity of 28-day periods,
    each discounted at (1 + r×28/360) compounded per period."""
    n = max(1, round(days / 28))
    p = (rate_pct / 100.0) * 28 / 360
    annuity = sum((1 + p) ** -i for i in range(1, n + 1)) * (28 / 360)
    return notional_mxn * annuity * 1e-4, n


# ═════════════════════════════════════════════════════════════════════════
# CLASSIFICATION
# ═════════════════════════════════════════════════════════════════════════
FX_OPTION_MARKERS = [
    "PUT", "CALL", "DELTA", "RISK REVERSAL", "BUTTERFLY",
    "STRADDLE", "STRANGLE", "SEAGULL", "DIGITAL", "OPTION",
]


def _has_ccy_pair(t: str) -> bool:
    if re.search(r'\b([A-Z]{3})\s*(?:VS\.?|V\.?|/|-)\s*([A-Z]{3})\b', t):
        return True
    if re.search(r'\b(USD|EUR|GBP|JPY|CAD|CHF|AUD|NZD)(MXN|CLP|COP|PEN|BRL|USD|EUR|JPY)\b', t):
        return True
    return False


def classify_category(text: str) -> str:
    if not isinstance(text, str):
        return "Other"
    t = text.upper()
    if "FX SWAP" in t or "FXSWAP" in t:
        return "FX"
    ir_markers = ["IRS", "OIS", "SWAP", "BASIS", "XBS", "XCCY", "FRA", "SBS",
                  "ZCI", "TIIE", "SOFR", "SONIA", "ESTR", "CORRA", "TONAT", "TLREF"]
    if any(m in t for m in ir_markers):
        return "IRS / Swap"
    if _has_ccy_pair(t) and any(m in t for m in FX_OPTION_MARKERS):
        return "FX"
    if any(m in t for m in ["NDF", "FX", "FWD", "FXO", "SPOT", "CALL", "PUT"]):
        return "FX"
    return "Other"


CATEGORY_TO_ASSETCLASS = {"IRS / Swap": "IR", "FX": "CU"}


def derive_asset_class(category: str) -> str:
    return CATEGORY_TO_ASSETCLASS.get(category, "—")


MEXICO_MARKERS = ["MXN", "TIIE", "F-TIIE", "UDI"]


def contains_mxn(row_text: str) -> bool:
    if not isinstance(row_text, str):
        return False
    t = row_text.upper()
    return any(marker in t for marker in MEXICO_MARKERS)


# ═════════════════════════════════════════════════════════════════════════
# PARSER 1 — TRADITION (pipe-delimited CSV)
# ═════════════════════════════════════════════════════════════════════════
def extract_tenor_tradition(desc: str) -> str:
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
    # _NDA = Non-Delta-Adjusted (GROSS). Tradition's _DA columns apply option
    # deltas as WHOLE numbers (50x, not 0.50x), so they aren't comparable to
    # the gross notionals every other SEF publishes. Swaps/NDFs: DA == NDA.
    for col in ["Total_Notional_USD_NDA", "Notional_Traded_Currency_NDA",
                "Total_Trade_Count", "First_Price", "Last_Price"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    desc_col = next((c for c in ["Internal_Prod_Des", "Internal_Prod_ID"] if c in df.columns), None)
    sub_prod = df["Sub_Prod"].astype(str) if "Sub_Prod" in df.columns else pd.Series([""] * len(df))
    desc_ser = df[desc_col].astype(str) if desc_col else pd.Series([""] * len(df))
    categories = (sub_prod.fillna("") + " " + desc_ser.fillna("")).apply(classify_category)
    return pd.DataFrame({
        "SEF": "Tradition",
        "Tenor": df[desc_col].apply(extract_tenor_tradition) if desc_col else "Other",
        "Category": categories,
        "AssetClass": categories.map(derive_asset_class),
        "Currency": df["Curr_Code"].str.strip().str.upper() if "Curr_Code" in df.columns else "N/A",
        "Notional_Local": df["Notional_Traded_Currency_NDA"] if "Notional_Traded_Currency_NDA" in df.columns else 0,
        "Notional_USD": df["Total_Notional_USD_NDA"] if "Total_Notional_USD_NDA" in df.columns else pd.NA,
        "Trades": df["Total_Trade_Count"] if "Total_Trade_Count" in df.columns else 1,
        "Trades_Estimated": False,
        "Last_Price": df["Last_Price"] if "Last_Price" in df.columns else pd.NA,
        "Description": df[desc_col] if desc_col else "",
        "MXN_Match": desc_ser.apply(contains_mxn) |
                     (df["Curr_Code"].astype(str) == "MXN" if "Curr_Code" in df.columns else False),
    })


# ═════════════════════════════════════════════════════════════════════════
# PARSER 2 — LATAM SEF
# ═════════════════════════════════════════════════════════════════════════
def extract_tenor_latam(row) -> str:
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
        return f"{round(periods * 28 / 30)}M"
    m2 = re.search(r'-\s*(\d{1,2}[A-Za-z]{3}\d{2,4})\s*$', d)
    if m2:
        expiry = None
        for fmt in ("%d%b%y", "%d%b%Y"):
            try:
                expiry = datetime.strptime(m2.group(1), fmt)
                break
            except ValueError:
                continue
        if expiry is None:
            return "Other"
        try:
            trade_date = datetime.strptime(str(row.get("Trade_Date", "")), "%Y%m%d")
        except ValueError:
            return "Other"
        return days_to_tenor_bucket((expiry - trade_date).days)
    return "Other"


def parse_latam(raw_text: str) -> pd.DataFrame:
    df = pd.read_csv(io.StringIO(raw_text), dtype=str)
    df.columns = [c.strip() for c in df.columns]
    for col in ["Notional_USD", "Notional_Traded_Currency", "Last_Price"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    categories = (df["Internal_Prod_Des"].apply(classify_category)
                  if "Internal_Prod_Des" in df.columns else pd.Series(["Other"] * len(df)))
    return pd.DataFrame({
        "SEF": "LatAm SEF",
        "Tenor": df.apply(extract_tenor_latam, axis=1),
        "Category": categories,
        "AssetClass": categories.map(derive_asset_class),
        "Currency": df["Curr_Code"].str.strip().str.upper() if "Curr_Code" in df.columns else "N/A",
        "Notional_Local": df["Notional_Traded_Currency"] if "Notional_Traded_Currency" in df.columns else 0,
        "Notional_USD": df["Notional_USD"] if "Notional_USD" in df.columns else pd.NA,
        "Trades": 1,   # LatAm publishes no trade count — one row = one print (proxy)
        "Trades_Estimated": True,
        "Last_Price": df["Last_Price"] if "Last_Price" in df.columns else pd.NA,
        "Description": df["Internal_Prod_Des"] if "Internal_Prod_Des" in df.columns else "",
        "MXN_Match": (df["Curr_Code"].astype(str) == "MXN" if "Curr_Code" in df.columns else False) |
                     (df["Internal_Prod_Des"].astype(str).apply(contains_mxn)
                      if "Internal_Prod_Des" in df.columns else False),
    })


# ═════════════════════════════════════════════════════════════════════════
# PARSER 3 — GFI / BGC
# ═════════════════════════════════════════════════════════════════════════
def extract_tenor_ccy_gfi_bgc(desc: str):
    if not isinstance(desc, str):
        return None, None
    m = re.search(r'\b(\d+)(Y|M)([A-Z]{3})\b', desc.upper())
    if m:
        return f"{m.group(1)}{m.group(2)}", m.group(3)
    return None, None


def extract_fx_tenor_gfi_bgc(desc: str, report_date: datetime):
    if not isinstance(desc, str):
        return "Other"
    m = re.search(r'(\d{1,2}[A-Z]{3}\d{4})', desc.upper())
    if m:
        try:
            expiry = datetime.strptime(m.group(1), "%d%b%Y")
            return days_to_tenor_bucket((expiry - report_date).days)
        except ValueError:
            return "Other"
    return "Other"


def _col_letters_to_index(col_str: str) -> int:
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx - 1


def read_xlsx_stdlib(raw_bytes: bytes, sheet_name: str):
    """Zero-dependency XLSX reader (zipfile + ElementTree) used when openpyxl
    isn't installed."""
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
                target = rid_to_target.get(sheet.get(f"{REL_NS}id"), "")
                sheet_file = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
                break
        if sheet_file is None:
            first = list(wb_xml.find("m:sheets", NS))[0]
            target = rid_to_target.get(first.get(f"{REL_NS}id"), "")
            sheet_file = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_xml.findall("m:si", NS):
                shared_strings.append("".join(t.text or "" for t in si.findall(".//m:t", NS)))
        sheet_xml = ET.fromstring(z.read(sheet_file))
        rows = []
        for row_el in sheet_xml.find("m:sheetData", NS).findall("m:row", NS):
            row_cells, max_col = {}, -1
            for c_el in row_el.findall("m:c", NS):
                ref = c_el.get("r", "A1")
                col_idx = _col_letters_to_index("".join(ch for ch in ref if ch.isalpha()))
                max_col = max(max_col, col_idx)
                v_el = c_el.find("m:v", NS)
                if v_el is None or v_el.text is None:
                    value = ""
                elif c_el.get("t") == "s":
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
            tenors.append(extract_fx_tenor_gfi_bgc(desc, report_date))
            currencies.append(row.get("Currency", "N/A"))
            categories.append("FX")
        else:
            tenors.append("Other")
            currencies.append(row.get("Currency", "N/A"))
            categories.append("Other")

    return pd.DataFrame({
        "SEF": sef_name,
        "Tenor": tenors,
        "Category": categories,
        "AssetClass": df["AssetClass"].values,
        "Currency": [str(c).strip().upper() for c in currencies],
        "Notional_Local": df["Volume"].values if "Volume" in df.columns else 0,
        "Notional_USD": pd.NA,
        "Trades": 1,   # no trade-count column — row count proxy
        "Trades_Estimated": True,
        "Last_Price": df["Close"].values if "Close" in df.columns else pd.NA,
        "Description": df["InstrumentDescription"].values,
        "MXN_Match": df["InstrumentDescription"].astype(str).apply(contains_mxn).values,
    })


# ═════════════════════════════════════════════════════════════════════════
# PARSER 4 — ICAP / tpSEF
# ═════════════════════════════════════════════════════════════════════════
def extract_tenor_icap(instrument: str, description: str) -> str:
    text = instrument if isinstance(instrument, str) else ""
    up = text.upper()

    m1 = re.search(r'\.(\d+)\*\d+\.', up)           # MXN.13*1.F-TIIE.OIS
    if m1:
        periods = int(m1.group(1))
        if periods in LATAM_PERIOD_TENOR:
            return LATAM_PERIOD_TENOR[periods]
        years = periods / 13.0
        return f"{round(years)}Y" if years >= 1 else f"{round(periods * 28 / 30)}M"

    m2 = re.search(r'(?:^|_)(\d+)X(\d+)(?:_|$)', up)  # IRS_MXN_..._0X2
    if m2:
        months = int(m2.group(2))
        return f"{months // 12}Y" if (months >= 12 and months % 12 == 0) else f"{months}M"

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
        return f"{m6.group(1)}Y" if m6.group(2) == "YEAR" else f"{m6.group(1)}M"

    return "Other"


def icap_category(asset_class: str) -> str:
    ac = str(asset_class).strip().upper()
    if ac == "IR":
        return "IRS / Swap"
    if ac == "CU":
        return "FX"
    return "Other"


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
    # letter-lookaround, not \b — underscore is a word char, so \b doesn't
    # match between '_' and 'USD' in '1M_USD_MXN_A_FXO'
    for code in ICAP_CCY_CODES:
        if re.search(rf'(?<![A-Z]){code}(?![A-Z])', t):
            return code
    for idx, ccy in ICAP_INDEX_CCY.items():
        if idx in t:
            return ccy
    return "N/A"


class _StdlibTableExtractor:
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
    df = df.rename(columns={k: v for k, v in {
        "Asset Class": "AssetClass", "Tradeable Instrument": "Instrument",
        "Num of Trades": "Trades", "Total Notional Value": "Notional",
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
                       if "AssetClass" in df.columns else "—"),
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
    tables = [t for t in extract_tables(html_text) if len(t) > 1]
    if not tables:
        raise ValueError("No data table found on the page")
    best = max(tables, key=len)
    header, data_rows = best[0], best[1:]
    data_rows = [r[:len(header)] + [""] * (len(header) - len(r)) for r in data_rows]
    df = pd.DataFrame(data_rows, columns=header)
    df.columns = [str(c).strip() for c in df.columns]
    return _standardize_icap_rows(df)


def parse_icap_numbers(raw_bytes: bytes) -> pd.DataFrame:
    from numbers_parser import Document
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".numbers", delete=False) as tmp:
        tmp.write(raw_bytes)
        tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
        rows = doc.sheets[0].tables[0].rows(values_only=True)
    finally:
        os.unlink(tmp_path)
    df = pd.DataFrame(rows[1:], columns=rows[0])
    df.columns = [str(c).strip() for c in df.columns]
    return _standardize_icap_rows(df)


def parse_icap_any(raw_bytes: bytes, filename: str) -> pd.DataFrame:
    """Accepts a .numbers export, a saved .html copy of the tpSEF page, or a
    .csv/.xlsx with the same columns — so a missed day can always be filed."""
    name = (filename or "").lower()
    if name.endswith(".numbers"):
        return parse_icap_numbers(raw_bytes)
    if name.endswith((".html", ".htm")):
        return parse_icap(raw_bytes.decode("utf-8", errors="ignore"))
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(raw_bytes))
        df.columns = [str(c).strip() for c in df.columns]
        return _standardize_icap_rows(df)
    df = pd.read_csv(io.BytesIO(raw_bytes))
    df.columns = [str(c).strip() for c in df.columns]
    # An already-standardized snapshot CSV round-trips as-is.
    if {"SEF", "Tenor", "Category", "Notional_Local"}.issubset(set(df.columns)):
        return df
    return _standardize_icap_rows(df)


# ═════════════════════════════════════════════════════════════════════════
# FETCHERS
# ═════════════════════════════════════════════════════════════════════════
BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
                  "(KHTML, like Gecko) Version/17.4.1 Safari/605.1.15",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}
SITE_HOMEPAGES = {"BGC": "https://www.bgcsef.com/", "GFI": "https://www.gfigroup.com/"}


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
    """GFI/BGC sit behind a WAF that fingerprints the TLS handshake itself —
    plain `requests` (and usually cloudscraper too) get HTTP 403 no matter
    what headers they send, because Python's TLS handshake doesn't look like
    a browser's. THE FIX: curl_cffi with Chrome impersonation presents a real
    Chrome TLS fingerprint and gets through. Order:
        1. curl_cffi (impersonate='chrome') — the one that actually works
        2. cloudscraper — solves plain-Cloudflare challenges
        3. warmed requests session — last resort
    curl_cffi MUST be in requirements.txt for this to work on Streamlit Cloud."""
    last_r = None

    # 1 — curl_cffi with a real Chrome TLS fingerprint
    try:
        from curl_cffi import requests as creq
        s = creq.Session(impersonate="chrome")
        try:
            s.get(homepage, timeout=15)
        except Exception:
            pass
        r = s.get(url, headers={"Referer": homepage}, timeout=30)
        if r.status_code == 200:
            return r
        last_r = r
    except Exception:
        pass

    # 2 — cloudscraper
    try:
        import cloudscraper
        scraper = cloudscraper.create_scraper(
            browser={"browser": "chrome", "platform": "darwin", "mobile": False})
        try:
            scraper.get(homepage, timeout=15)
        except Exception:
            pass
        r = scraper.get(url, timeout=30)
        if r.status_code == 200:
            return r
        if last_r is None:
            last_r = r
    except Exception:
        pass

    # 3 — warmed plain-requests session
    r = _make_session(homepage).get(url, timeout=25)
    return r if r.status_code == 200 else (last_r if last_r is not None else r)


def _looks_valid(kind: str, resp) -> bool:
    """A 200 isn't enough — a WAF can serve an HTML block page with status
    200. An xlsx must start with the ZIP magic bytes 'PK'."""
    try:
        if kind == "gfi_bgc":
            return resp.content[:2] == b"PK"
        return True
    except Exception:
        return False


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_source(sef_name: str, urls, kind: str, report_date_str: str):
    """Try each candidate URL in order. The direct links off the desk's sheet
    are FIRST and are tried with a plain request — they just work, no
    anti-bot dance needed (GFI in particular serves over plain http). The
    curl_cffi/cloudscraper chain is kept as a fallback per URL."""
    if isinstance(urls, str):
        urls = (urls,)
    r, last_err = None, None
    for url in urls:
        # 1 — plain direct request (this is all the desk's links need)
        try:
            resp = requests.get(url, headers=BROWSER_HEADERS, timeout=25, allow_redirects=True)
        except Exception as e:
            resp, last_err = None, f"Network error: {e}"
        if resp is not None and resp.status_code == 200 and _looks_valid(kind, resp):
            r = resp
            break
        # 2 — anti-bot chain, only if the direct request didn't get through
        if sef_name in SITE_HOMEPAGES:
            try:
                resp2 = _fetch_protected(url, SITE_HOMEPAGES[sef_name])
                if resp2.status_code == 200 and _looks_valid(kind, resp2):
                    r = resp2
                    break
                resp = resp2
            except Exception as e:
                last_err = f"Network error: {e}"
        if resp is not None:
            if resp.status_code == 404:
                last_err = "No file for this date (market may have been closed)."
            elif resp.status_code == 403:
                last_err = "Blocked (HTTP 403) — every URL variant was refused."
            elif resp.status_code == 200:
                last_err = "Got a page, but not the data file (WAF block page)."
            else:
                last_err = f"HTTP {resp.status_code}"
    if r is None:
        return None, last_err or "Fetch failed"
    try:
        report_date = datetime.strptime(report_date_str, "%Y%m%d")
        if kind == "tradition":
            return parse_tradition(r.text), None
        if kind == "latam":
            return parse_latam(r.text), None
        if kind == "gfi_bgc":
            return parse_gfi_bgc(r.content, sef_name, report_date), None
        if kind == "icap":
            return parse_icap(r.text), None
    except Exception as e:
        return None, f"Parse error: {e}"
    return None, "Unknown source kind"


def fetch_uploaded(sef_name: str, kind: str, uploaded_file, report_date: datetime):
    try:
        if kind == "gfi_bgc":
            return parse_gfi_bgc(uploaded_file.read(), sef_name, report_date), None
        if kind == "latam":
            return parse_latam(uploaded_file.read().decode("utf-8")), None
        if kind == "tradition":
            return parse_tradition(uploaded_file.read().decode("utf-8")), None
        if kind == "icap":
            return parse_icap_any(uploaded_file.read(), uploaded_file.name), None
    except ImportError:
        return None, "Missing package — run: pip install numbers-parser"
    except Exception as e:
        return None, f"Parse error: {e}"
    return None, "Unsupported upload kind"


@st.cache_data(ttl=86400, show_spinner=False, max_entries=400)
def fetch_source_hist(sef_name: str, urls, kind: str, report_date_str: str):
    """Cached ONLY on success. This function raises on failure, and
    st.cache_data doesn't cache exceptions — so a day that fails (WAF mood,
    transient timeout) is retried on the next run instead of the failure
    being frozen in cache for 24 hours. That freezing is exactly how a BGC
    day that failed ONCE kept showing as missing all day even after the
    fetch problem was fixed."""
    df, err = fetch_source(sef_name, urls, kind, report_date_str)
    if df is None:
        raise RuntimeError(err or "fetch failed")
    return df


# ═════════════════════════════════════════════════════════════════════════
# ICAP DAILY SNAPSHOTS — one file per day, never combined
# ═════════════════════════════════════════════════════════════════════════
def _icap_snapshot_path(d: datetime) -> str:
    return os.path.join(ICAP_SNAPSHOT_DIR, f"icap_{d.strftime('%Y%m%d')}.csv")


def save_icap_snapshot(df: pd.DataFrame, d: datetime, push: bool = True) -> bool:
    """Write the day-file AND immediately push it to GitHub. Local disk on
    Streamlit Cloud is wiped on every redeploy — a snapshot that isn't pushed
    the moment it's written can be lost forever (tpSEF can't be re-fetched)."""
    try:
        os.makedirs(ICAP_SNAPSHOT_DIR, exist_ok=True)
        df.to_csv(_icap_snapshot_path(d), index=False)
        if push:
            push_day_file(_icap_snapshot_path(d), ICAP_SNAPSHOT_DIR)
        return True
    except Exception:
        return False


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


def list_icap_snapshots():
    if not os.path.isdir(ICAP_SNAPSHOT_DIR):
        return []
    out = []
    for f in sorted(os.listdir(ICAP_SNAPSHOT_DIR)):
        m = re.fullmatch(r'icap_(\d{8})\.csv', f)
        if m:
            out.append(datetime.strptime(m.group(1), "%Y%m%d"))
    return sorted(out)


def ensure_todays_icap_snapshot() -> bool:
    """Capture the latest tpSEF page on EVERY app load, whatever date is
    selected, and push it straight to the repo. Returns True if a new file
    was written."""
    d = last_business_day()
    if os.path.exists(_icap_snapshot_path(d)):
        return False
    try:
        df, _ = fetch_source("ICAP", source_urls(SOURCES["ICAP"], d), "icap", d.strftime("%Y%m%d"))
        if df is not None and not df.empty:
            return save_icap_snapshot(df, d)   # save_icap_snapshot pushes too
    except Exception:
        pass
    return False


def get_icap_for_date(cfg: dict, sel_date: datetime):
    if sel_date.date() == last_business_day().date():
        df, err = fetch_source("ICAP", source_urls(cfg, sel_date), "icap", sel_date.strftime("%Y%m%d"))
        if df is not None:
            save_icap_snapshot(df, sel_date)   # saves AND pushes
            return df, None
        snap = load_icap_snapshot(sel_date)
        return (snap, None) if snap is not None else (None, err)
    snap = load_icap_snapshot(sel_date)
    if snap is not None:
        return snap, None
    return None, ("No ICAP snapshot saved for this date. tpSEF only ever shows the latest "
                  "business day, so this one can't be re-fetched — upload an export under "
                  "Sidebar → ICAP history → Backfill a past date.")


def _sef_snapshot_path(sef: str, d: datetime) -> str:
    safe = re.sub(r'\W+', '', sef)
    return os.path.join(SEF_SNAPSHOT_DIR, f"{safe}_{d.strftime('%Y%m%d')}.csv")


def save_sef_snapshot(df: pd.DataFrame, sef: str, d: datetime) -> bool:
    """Save any SEF's parsed day-table and push it to the repo — same idea as
    the ICAP snapshots. This is what makes month-to-date reliable on Streamlit
    Cloud: GFI/BGC block bursts of requests from datacenter IPs, so re-fetching
    8 past days on every MTD build silently loses them. Once a day is saved,
    it's never fetched again. Skips the write if the file already exists."""
    try:
        p = _sef_snapshot_path(sef, d)
        if os.path.exists(p):
            return True
        os.makedirs(SEF_SNAPSHOT_DIR, exist_ok=True)
        df.to_csv(p, index=False)
        push_day_file(p, SEF_SNAPSHOT_DIR)
        return True
    except Exception:
        return False


def load_sef_snapshot(sef: str, d: datetime):
    p = _sef_snapshot_path(sef, d)
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


def missing_history_days(since: datetime = HISTORY_START):
    """Business days from `since` to the last business day with no ICAP file."""
    have = {d.date() for d in list_icap_snapshots()}
    out, d, end = [], since, last_business_day()
    while d.date() <= end.date():
        if d.weekday() < 5 and d.date() not in have:
            out.append(d)
        d += timedelta(days=1)
    return out


# ═════════════════════════════════════════════════════════════════════════
# GITHUB SYNC — pull at startup, push new days (survives Cloud redeploys)
# ═════════════════════════════════════════════════════════════════════════
SYNC_DIRS = [ICAP_SNAPSHOT_DIR, SEF_SNAPSHOT_DIR, DV01_GRID_DIR]


def gh_cfg():
    """Config comes from .streamlit/secrets.toml, with environment variables
    as a fallback (GH_REPO / GH_BRANCH / GH_TOKEN).

    ON STREAMLIT CLOUD: paste the [github] block into the app's Secrets box.

    ON LOCALHOST: secrets don't follow you from the cloud — create the file
    yourself, in the SAME folder you run `streamlit run theNiche.py` from:

        mkdir -p .streamlit
        # then create .streamlit/secrets.toml containing:
        #   [github]
        #   repo   = "nminniti7/Mexico-Rates-FX"
        #   branch = "main"
        #   token  = "github_pat_..."

    and add `.streamlit/secrets.toml` to .gitignore so the token is never
    committed (GitHub auto-revokes any token it finds in a push, which kills
    sync everywhere). Restart streamlit after creating the file.

    If it's absent everywhere, the app just works off local disk. Best-effort:
    a failed push or pull never interrupts anything on screen — but the
    sidebar SAYS when sync is off, because on Streamlit Cloud that means
    history dies on the next redeploy."""
    try:
        s = st.secrets.get("github", {})
    except Exception:
        s = {}
    repo = s.get("repo", "") or os.environ.get("GH_REPO", "")
    branch = s.get("branch", "") or os.environ.get("GH_BRANCH", "") or "main"
    token = s.get("token", "") or os.environ.get("GH_TOKEN", "")
    return {"repo": repo, "branch": branch, "token": token}


def push_day_file(local_path: str, folder: str):
    """Best-effort push of one day-file. Silent no-op if GitHub isn't set up."""
    c = gh_cfg()
    if not (c["repo"] and c["token"] and os.path.exists(local_path)):
        return False
    ok, _ = github_push_file(local_path, c["repo"], c["branch"], c["token"],
                             f"{folder}/{os.path.basename(local_path)}")
    if ok:
        st.session_state.setdefault("gh_pushed", set()).add(
            f"{folder}/{os.path.basename(local_path)}")
    return ok


def _gh_headers(token):
    h = {"Accept": "application/vnd.github+json"}
    if token:
        h["Authorization"] = f"token {token}"
    return h


def github_push_file(local_path, repo, branch, token, remote_path):
    if not (repo and token and os.path.exists(local_path)):
        return False, "Missing repo, token, or file."
    api_url = f"https://api.github.com/repos/{repo}/contents/{remote_path}"
    try:
        with open(local_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()
        sha = None
        r = requests.get(api_url, headers=_gh_headers(token), params={"ref": branch}, timeout=15)
        if r.status_code == 200:
            sha = r.json().get("sha")
        payload = {"message": f"Update {os.path.basename(local_path)}",
                   "content": content_b64, "branch": branch}
        if sha:
            payload["sha"] = sha
        r = requests.put(api_url, headers=_gh_headers(token), json=payload, timeout=20)
        if r.status_code in (200, 201):
            return True, "OK"
        return False, f"HTTP {r.status_code}: {r.text[:160]}"
    except Exception as e:
        return False, str(e)


def github_push_dirs(repo, branch, token, only_new=False):
    ok = fail = 0
    for d in SYNC_DIRS:
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            if not fn.endswith(".csv"):
                continue
            key = f"{d}/{fn}"
            if only_new and key in st.session_state.get("gh_pushed", set()):
                continue
            success, _ = github_push_file(os.path.join(d, fn), repo, branch, token, key)
            if success:
                ok += 1
                st.session_state.setdefault("gh_pushed", set()).add(key)
            else:
                fail += 1
    return ok, fail


def github_pull_dirs(repo, branch, token):
    """Download any day-files in the repo that aren't on local disk. This is
    what makes the shared Streamlit link show the full history after a
    redeploy wipes the container's disk."""
    pulled = 0
    for d in SYNC_DIRS:
        api_url = f"https://api.github.com/repos/{repo}/contents/{d}"
        try:
            r = requests.get(api_url, headers=_gh_headers(token), params={"ref": branch}, timeout=20)
            if r.status_code != 200:
                continue
            os.makedirs(d, exist_ok=True)
            for item in r.json():
                if item.get("type") != "file" or not item.get("name", "").endswith(".csv"):
                    continue
                local = os.path.join(d, item["name"])
                if os.path.exists(local):
                    continue
                raw = requests.get(item["download_url"], headers=_gh_headers(token), timeout=20)
                if raw.status_code == 200:
                    with open(local, "wb") as f:
                        f.write(raw.content)
                    pulled += 1
        except Exception:
            continue
    return pulled


def sync_with_github_once():
    """Runs once per session: pull history down, capture today, push back up.
    MUST run before the sidebar renders — the sidebar counts snapshot files,
    and counting them before the pull made a fresh container always claim
    only one day was saved."""
    if st.session_state.get("gh_synced"):
        return
    st.session_state["gh_synced"] = True
    cfg = gh_cfg()
    if cfg["repo"]:
        st.session_state["gh_pulled"] = github_pull_dirs(cfg["repo"], cfg["branch"], cfg["token"])
    ensure_todays_icap_snapshot()
    ensure_todays_dv01_grid()
    if cfg["repo"] and cfg["token"]:
        github_push_dirs(cfg["repo"], cfg["branch"], cfg["token"], only_new=True)


# Pull history from GitHub, capture today's ICAP + grid, push back — BEFORE
# the sidebar renders, so the day counts and gap warnings are correct.
sync_with_github_once()

# ═════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("🇲🇽 Mexico Desk")
    st.caption("Cross-SEF market comparison")

    _gh = gh_cfg()
    if _gh["repo"] and _gh["token"]:
        _pulled = st.session_state.get("gh_pulled", 0)
        st.caption(f"☁️ GitHub sync **on** ({_gh['repo']})" +
                   (f" · pulled {_pulled} file(s)" if _pulled else ""))
    else:
        st.warning("GitHub sync is **not configured** — snapshots live on this "
                   "container's disk only and will be LOST on the next redeploy. "
                   "Add the [github] block to Secrets.", icon="⚠️")

    # ── 1. Date ─────────────────────────────────────────────────────────
    st.subheader("1 · Trade date")
    auto_date = last_business_day()
    override = st.checkbox("Pick a different date", help="Otherwise the last business day is used.")
    if override:
        sel_date = datetime.combine(st.date_input("Trade date", value=auto_date), datetime.min.time())
    else:
        sel_date = datetime.combine(auto_date, datetime.min.time())
        st.caption(f"Using **{auto_date:%b %d, %Y}** (last business day)")

    # ── 2. Metric ───────────────────────────────────────────────────────
    st.subheader("2 · Compare by")
    metric_choice = st.radio(
        "Metric", ["DV01 (USD)", "Notional (as published)", "Trade count"],
        label_visibility="collapsed",
        help=("DV01 = MXN notional (mm) ÷ desk grid × 5,000 — MXN rates rows only. "
              "Notional is whatever the source file says, in the currency it was traded "
              "in. Nothing is FX-converted anywhere in this app."),
    )
    show_exact = st.checkbox("Exact numbers (not 1.23M)", value=False)
    show_mtd = st.checkbox("Month-to-date share", value=True,
                           help="Pulls every prior business day this month. First load is slower.")

    # ── 3. DV01 grid (per day) ──────────────────────────────────────────
    st.subheader("3 · DV01 grid")
    grid_for_date, grid_label = load_dv01_grid_for_date(sel_date)
    st.caption(f"**{sel_date:%b %d}** — {grid_label}")

    with st.expander("📋 Paste the desk's grid", expanded=False):
        st.caption("Copy the tenor + number columns straight out of the email or Excel and "
                   "paste below. Two columns, one per line. `13x1` period notation, `10 Year`, "
                   "`10Y`, commas and $ signs all work. A bare column of numbers also works — "
                   "it maps onto the ladder in order.")
        pasted = st.text_area(
            "Paste here",
            height=140,
            placeholder="1x1\t11,293.90\n3x1\t3,784.42\n13x1\t896.60\n10Y\t126.22",
            key=f"paste_{sel_date:%Y%m%d}",
            label_visibility="collapsed",
        )
        if pasted.strip():
            ladder_order = sorted(grid_for_date.keys(), key=tenor_sort_key)
            pg, n, bad = parse_pasted_grid(pasted, ladder_order)
            if not pg:
                st.error("Couldn't read a single tenor/value pair out of that.")
            else:
                prev = pd.DataFrame({"Tenor": sorted(pg.keys(), key=tenor_sort_key)})
                prev["MXN mm per 5k DV01"] = prev["Tenor"].map(pg)
                st.caption(f"Read **{len(pg)} tenor(s)**" +
                           (f" · skipped {len(bad)} line(s) (header/junk)" if bad else ""))
                st.dataframe(prev, hide_index=True, use_container_width=True, height=200)
                merge = st.checkbox("Keep tenors I didn't paste", value=True,
                                    help="On: only the pasted tenors are overwritten. "
                                         "Off: the grid becomes exactly what you pasted.")
                if st.button(f"✅ Apply & save for {sel_date:%b %d}", type="primary",
                             use_container_width=True):
                    final = {**grid_for_date, **pg} if merge else dict(pg)
                    if save_dv01_grid_for_date(final, sel_date):
                        push_day_file(_dv01_path(sel_date), DV01_GRID_DIR)
                        st.success(f"Saved {len(final)} tenors for {sel_date:%b %d}.")
                        st.rerun()
                    else:
                        st.error("Couldn't save.")

    with st.expander("✏️ Edit one tenor at a time"):
        _gdf = pd.DataFrame({"Tenor": sorted(grid_for_date.keys(), key=tenor_sort_key)})
        _gdf["MXN mm per 5k DV01"] = _gdf["Tenor"].map(grid_for_date)
        edited = st.data_editor(_gdf, num_rows="dynamic", hide_index=True,
                                key=f"grid_{sel_date:%Y%m%d}", use_container_width=True)
        dv01_grid = {}
        for _, r in edited.iterrows():
            t = str(r.get("Tenor", "")).strip().upper()
            v = pd.to_numeric(r.get("MXN mm per 5k DV01"), errors="coerce")
            if t and pd.notna(v) and v > 0:
                dv01_grid[t] = float(v)
        if not dv01_grid:
            dv01_grid = dict(grid_for_date)
        c1, c2 = st.columns(2)
        with c1:
            if st.button(f"💾 Save for {sel_date:%b %d}", use_container_width=True):
                if save_dv01_grid_for_date(dv01_grid, sel_date):
                    push_day_file(_dv01_path(sel_date), DV01_GRID_DIR)
                    st.success("Saved.")
                    st.rerun()
                else:
                    st.error("Couldn't save.")
        with c2:
            if st.button("↩️ Reset to default", use_container_width=True):
                save_dv01_grid_for_date(DEFAULT_DV01_GRID, sel_date)
                push_day_file(_dv01_path(sel_date), DV01_GRID_DIR)
                st.rerun()

    saved_days = list_dv01_grid_dates()
    if saved_days:
        st.caption("Grids on file: " + " · ".join(f"{d:%b %d}" for d in saved_days))

    # ── 4. Sources & history ────────────────────────────────────────────
    st.subheader("4 · Sources")
    uploaded_files = {}
    for name, cfg in SOURCES.items():
        st.caption(f"🟢 {name} — auto-fetch")
    with st.expander("Upload / override a source file"):
        for name, cfg in SOURCES.items():
            types = ["numbers", "html", "csv", "xlsx"] if cfg["kind"] == "icap" else ["csv", "xlsx"]
            f = st.file_uploader(name, key=f"upload_{name}", type=types)
            if f:
                uploaded_files[name] = f

    with st.expander("📅 ICAP history"):
        snaps = list_icap_snapshots()
        st.caption(f"**{len(snaps)} day(s) saved** — one file per day, never combined.")
        if snaps:
            st.caption(" · ".join(f"{d:%b %d}" for d in snaps))
        gaps = missing_history_days()
        if gaps:
            st.warning("Missing since Jul 09: " + ", ".join(f"{d:%b %d}" for d in gaps))
            st.caption("tpSEF only shows the latest business day, so past days can't be "
                       "re-fetched — but you can file an export under any date below.")
        st.markdown("**Backfill a past date**")
        bf_date = st.date_input("Date to file it under", value=HISTORY_START, key="bf_date")
        bf_file = st.file_uploader("ICAP export (.numbers / .html / .csv / .xlsx)",
                                   key="bf_file", type=["numbers", "html", "csv", "xlsx"])
        if bf_file and st.button("Save as that day's snapshot"):
            try:
                bdf = parse_icap_any(bf_file.read(), bf_file.name)
                bdt = datetime.combine(bf_date, datetime.min.time())
                if save_icap_snapshot(bdf, bdt):   # saves AND pushes
                    st.success(f"Saved {len(bdf):,} rows for {bdt:%b %d}.")
                    st.rerun()
            except Exception as e:
                st.error(f"Couldn't parse that file: {e}")

    # ── 5. Tools ────────────────────────────────────────────────────────
    st.subheader("5 · Tools")
    with st.expander("🧮 DV01 / brokerage calculator"):
        calc_rate = st.number_input("Rate (%)", min_value=0.0, value=6.97, step=0.01, format="%.4f")
        calc_days = st.number_input("No. of days", min_value=1, value=364, step=28)
        calc_amt = st.number_input("Amount (MXN)", min_value=0.0, value=3_390_000_000.0,
                                   step=1_000_000.0, format="%.2f")
        calc_brk = st.number_input("Brokerage rate", min_value=0.0, value=0.100, step=0.005, format="%.3f")
        calc_spot = st.number_input("Spot (USDMXN)", min_value=0.0001, value=18.39, step=0.01, format="%.4f")
        pv01_mxn, n_per = tiie_pv01(calc_amt, calc_rate, int(calc_days))
        st.markdown(
            f"**{n_per} × 28-day periods**\n\n"
            f"PV01: **{pv01_mxn:,.2f} MXN** · **{pv01_mxn/calc_spot:,.2f} USD**\n\n"
            f"Brokerage: **{pv01_mxn*calc_brk:,.2f} MXN** · **{pv01_mxn*calc_brk/calc_spot:,.2f} USD**"
        )
        if calc_amt > 0 and pv01_mxn > 0:
            grid_equiv = DV01_GRID_BASE / ((pv01_mxn / calc_spot) / (calc_amt / 1e6))
            st.caption(f"Grid equivalent: **{grid_equiv:,.2f}** MXN mm per 5,000 USD DV01.")


numfmt = fmt_exact if show_exact else fmt_m

# ═════════════════════════════════════════════════════════════════════════
# HEADER
# ═════════════════════════════════════════════════════════════════════════
st.title("Mexico Rates & FX — Cross-SEF Comparison")
hc1, hc2, hc3 = st.columns([2, 2, 2])
hc1.caption(f"**Trade date:** {sel_date:%B %d, %Y}")
hc2.caption(f"**Metric:** {metric_choice}")
hc3.caption(f"**DV01 grid:** {grid_label}")

with st.expander("ℹ️ How this works (methodology, in one place)"):
    st.markdown(
        "**Nothing is FX-converted.** Every notional on screen is exactly what the source file "
        "publishes, in the currency it was traded in. Currencies are never added together — the "
        "currency picker shows them one at a time. That's why a bucket total can always be "
        "reconciled straight against the raw file.\n\n"
        "**DV01 (USD)** = MXN notional (millions) ÷ the desk grid value for that tenor × 5,000. "
        "The grid is the desk's own *notional per 5,000 USD DV01* table and moves daily with the "
        "TIIE curve and USDMXN spot, so it's saved **per day**: pick a past date and you value on "
        "that day's grid. If it was never edited that day, the last saved grid carries forward and "
        "says so at the top. Tenors between grid points are interpolated (flagged in the "
        "drill-down). MXN rates rows only — a non-MXN row or a spread trade with no single tenor "
        "gets no DV01 rather than an invented one.\n\n"
        "**Tenors.** MXN TIIE swaps are quoted in 28-day periods, not calendar time — 13 periods "
        "(364 days) is \"1 year.\" Everything is converted to standard Y/M labels so all five SEFs "
        "line up on the same row (verified by price match: LatAm's 13x1 and Tradition's \"1 Year\" "
        "quoted 6.735 the same day). Day-count prints ('182 Days') bucket to their real tenor (6M). "
        "1W/2W/3W roll into a 0M bucket on the FX ladder; the drill-down always shows the true "
        "tenor and the bucket side by side.\n\n"
        "**Tradition notionals** are the GROSS (Non-Delta-Adjusted) columns. Their DA columns apply "
        "option deltas as whole numbers (50x, not 0.50x), which isn't comparable to what the other "
        "SEFs publish.\n\n"
        "**Trade counts.** Tradition and ICAP publish a real count. LatAm, GFI and BGC don't — those "
        "use row count as a proxy and say so wherever it's shown.\n\n"
        "**History.** ICAP's page only ever shows the latest business day, so the app saves a "
        "snapshot every load (`icap_snapshots/icap_YYYYMMDD.csv`) and the DV01 grid the same way "
        "(`dv01_grids/dv01_YYYYMMDD.csv`) — one file per day, never combined or appended. Every "
        "save is pushed straight to the GitHub repo so Streamlit Cloud redeploys can't lose it."
    )
    _missing = st.session_state.get("mtd_missing", [])
    if _missing:
        st.markdown("---")
        st.markdown(
            "**Month-to-date gaps.** These source-days aren't in the MTD pies: **" +
            "** · **".join(_missing) + "**. They weren't saved and are no longer fetchable "
            "(BGC and tpSEF only keep the most recent days online, so a day that rolls off "
            "before it's snapshotted is gone from the site). The app retries these on every "
            "load and saves every day it sees going forward. To fill one in now: pick that "
            "date in the sidebar and upload the day's file under Sources → Upload — it saves "
            "to the repo permanently and the MTD pies update immediately.")

# ═════════════════════════════════════════════════════════════════════════
# FETCH ALL SOURCES
# ═════════════════════════════════════════════════════════════════════════
all_data, status_msgs = [], []
with st.spinner("Pulling data from all connected SEFs..."):
    for name, cfg in SOURCES.items():
        if name in uploaded_files:
            df, err = fetch_uploaded(name, cfg["kind"], uploaded_files[name], sel_date)
            if err:
                status_msgs.append((name, "error", err))
            else:
                if cfg["kind"] == "icap":
                    save_icap_snapshot(df, sel_date)   # saves AND pushes
                else:
                    save_sef_snapshot(df, name, sel_date)
                all_data.append(df)
                status_msgs.append((name, "ok", f"{len(df):,} rows (uploaded)"))
        elif cfg["kind"] == "icap":
            df, err = get_icap_for_date(cfg, sel_date)
            if err:
                status_msgs.append((name, "error", err))
            else:
                all_data.append(df)
                status_msgs.append((name, "ok", f"{len(df):,} rows"))
        else:
            df, err = fetch_source(name, source_urls(cfg, sel_date), cfg["kind"],
                                   sel_date.strftime("%Y%m%d"))
            if err:
                # A fetch can fail on Cloud (WAF) even when the day was saved
                # before — fall back to the saved snapshot instead of a blank.
                snap = load_sef_snapshot(name, sel_date)
                if snap is not None:
                    all_data.append(snap)
                    status_msgs.append((name, "ok", f"{len(snap):,} rows (saved snapshot)"))
                else:
                    status_msgs.append((name, "error", err))
            else:
                save_sef_snapshot(df, name, sel_date)
                all_data.append(df)
                status_msgs.append((name, "ok", f"{len(df):,} rows"))

cols = st.columns(len(SOURCES))
for i, (name, kind, msg) in enumerate(status_msgs):
    icon = "✅" if kind == "ok" else ("⚠️" if kind == "error" else "⬜")
    cols[i].caption(f"{icon} **{name}** — {msg}")

if not all_data:
    st.error("No data loaded from any source. Check the date, or upload files in the sidebar.")
    st.stop()

combined = pd.concat(all_data, ignore_index=True)

# Currency must always be a real string. A blank cell in any source file
# (seen on GFI/BGC) reads back as float NaN, and one stray float here breaks
# every later sorted()/groupby() that mixes it with real currency codes —
# that TypeError was what stopped the whole page from rendering.
combined["Currency"] = (
    combined["Currency"].astype(str).str.strip().str.upper()
    .replace({"NAN": "N/A", "NONE": "N/A", "": "N/A", "<NA>": "N/A"})
)
combined["Tenor_Bucket"] = combined["Tenor"].astype(str).apply(tenor_bucket)
combined["Tenor"] = sort_tenors(combined["Tenor"])
combined = add_dv01(combined, dv01_grid)

if metric_choice == "DV01 (USD)":
    metric_col = "DV01_USD"
elif metric_choice == "Notional (as published)":
    metric_col = "Notional_Local"
else:
    metric_col = "Trades"


# ═════════════════════════════════════════════════════════════════════════
# TABLES & CHARTS
# ═════════════════════════════════════════════════════════════════════════
def build_comparison(data: pd.DataFrame, sefs_present: list, mcol: str, category_label: str) -> pd.DataFrame:
    data = data.copy()
    if category_label == "IRS / Swap":
        ladder = list(IRS_DISPLAY_TENORS)
        data["Disp_Bucket"] = data["Tenor_Bucket"].where(data["Tenor_Bucket"].isin(ladder), "Other")
        other_label = "Other / non-standard tenor"
    else:
        present = [str(t) for t in data["Tenor_Bucket"].unique() if str(t) != "Other"]
        extras = safe_sorted([t for t in present if t not in FULL_TENOR_LADDER], key=tenor_sort_key)
        ladder = FULL_TENOR_LADDER + extras
        data["Disp_Bucket"] = data["Tenor_Bucket"]
        other_label = "Other / unmatched"
    if (data["Disp_Bucket"] == "Other").any():
        ladder = ladder + ["Other"]

    rows = []
    for tenor in ladder:
        t_df = data[data["Disp_Bucket"] == tenor]
        if t_df.empty:
            continue
        row = {"Tenor": other_label if tenor == "Other" else tenor_display(tenor)}
        total = 0
        for sef in sefs_present:
            v = t_df[t_df["SEF"] == sef][mcol].sum()
            row[sef] = v
            total += 0 if pd.isna(v) else v
        if total == 0:
            continue
        row["Total"] = total
        for sef in sefs_present:
            row[f"{sef} %"] = round(row[sef] / total * 100, 1) if total > 0 else 0
        rows.append(row)
    return pd.DataFrame(rows)


def render_market_share_chart(cat_data, sefs_present, category_label, currency_label, mcol, chart_key):
    totals = cat_data.groupby("SEF")[mcol].sum().reindex(sefs_present).fillna(0)
    grand_total = totals.sum()
    if grand_total <= 0:
        return
    shares = (totals / grand_total * 100).round(1)
    st.markdown(f"**Market share — {category_label} ({currency_label})**")
    pie_df = pd.DataFrame({"SEF": totals.index, "Value": totals.values})
    fig = px.pie(pie_df, names="SEF", values="Value", color="SEF",
                 color_discrete_map={s: SEF_COLORS.get(s, "#999999") for s in sefs_present}, hole=0.35)
    fig.update_traces(textposition="inside", textinfo="percent+label", sort=False)
    fig.update_layout(margin=dict(t=10, b=10, l=10, r=10),
                      height=max(320, 55 * len(sefs_present)), showlegend=True)
    st.plotly_chart(fig, use_container_width=True, key=chart_key)
    st.caption("  ·  ".join(f"**{s}**: {shares[s]}% ({numfmt(totals[s])})" for s in sefs_present))


def pick_currency(cat_data, category_label, key_suffix=""):
    """Notional is NEVER summed across currencies — pick one and show only it.
    MXN by default when present, otherwise the biggest by notional."""
    ccys = safe_unique(cat_data["Currency"])
    ccys = [c for c in ccys if c != "N/A"] or ccys
    if not ccys:
        return None
    if len(ccys) == 1:
        return ccys[0]
    if "MXN" in ccys:
        default = ccys.index("MXN")
    else:
        biggest = cat_data.groupby("Currency")["Notional_Local"].sum().idxmax()
        default = ccys.index(str(biggest)) if str(biggest) in ccys else 0
    return st.selectbox(
        "Notional currency", ccys, index=default,
        key=f"ccy_{category_label}{key_suffix}",
        help="Straight off the source file — no FX conversion anywhere, so currencies "
             "are shown one at a time rather than added together.",
    )


def render_section(data, category_label, sefs_present, group_label, mtd_data=None):
    cat_data = data[data["Category"] == category_label]
    if cat_data.empty:
        st.info(f"No {category_label} data for {group_label}.")
        return

    mcol, mlabel = metric_col, metric_choice

    # DV01 is a rates concept and the grid is MXN-only — the FX tab compares by
    # notional as published.
    if category_label == "FX" and mcol == "DV01_USD":
        mcol = "Notional_Local"
        st.caption("DV01 doesn't apply to FX — this tab compares by notional as published.")

    ccy = None
    if mcol == "Notional_Local":
        ccy = pick_currency(cat_data, category_label)
        if ccy:
            cat_data = cat_data[cat_data["Currency"] == ccy]
        mlabel = f"Notional ({ccy}, as published)"

    if cat_data.empty:
        st.info("Nothing traded in that currency on this date.")
        return

    if mcol == "DV01_USD":
        no_dv01 = cat_data[cat_data["DV01_USD"].isna()]
        if not no_dv01.empty:
            st.caption(f"{len(no_dv01)} row(s) get no DV01 — non-MXN rates rows and spread "
                       f"trades with no single tenor. They're excluded rather than approximated.")

    # KPI row
    totals = cat_data.groupby("SEF")[mcol].sum()
    grand_total = totals.sum()
    kpi_cols = st.columns(len(sefs_present))
    unit = ("" if mcol == "Trades" else (" USD DV01" if mcol == "DV01_USD" else f" {ccy}"))
    for i, sef in enumerate(sefs_present):
        v = totals.get(sef, 0)
        share = round(v / grand_total * 100, 1) if grand_total > 0 else 0
        kpi_cols[i].metric(sef, f"{share}%", f"{numfmt(v)}{unit}")

    render_market_share_chart(cat_data, sefs_present, category_label, "today", mcol,
                              chart_key=f"pie_day_{category_label}")

    if mtd_data is not None:
        mtd_cat = mtd_data[mtd_data["Category"] == category_label]
        if ccy:
            mtd_cat = mtd_cat[mtd_cat["Currency"] == ccy]
        if not mtd_cat.empty and mcol in mtd_cat.columns:
            days_covered = safe_sorted(mtd_cat["Date"].unique())
            render_market_share_chart(
                mtd_cat, sefs_present, category_label,
                f"month-to-date · {days_covered[0]} → {days_covered[-1]} · {len(days_covered)} day(s)",
                mcol, chart_key=f"pie_mtd_{category_label}")
            if "ICAP" not in mtd_cat["SEF"].unique():
                st.caption("ICAP joins the MTD view as daily snapshots accumulate.")

    # Tenor ladder
    comp = build_comparison(cat_data, sefs_present, mcol, category_label)
    if comp.empty:
        st.info("No rows with a measurable value for the chosen metric.")
        return
    display_cols = ["Tenor"] + sefs_present + ["Total"] + [f"{s} %" for s in sefs_present]
    st.markdown(f"**Tenor ladder — {len(comp)} row(s), values in {mlabel}**")
    st.dataframe(
        comp[display_cols].style.format({**{s: numfmt for s in sefs_present}, "Total": numfmt,
                                         **{f"{s} %": "{:.1f}%" for s in sefs_present}}),
        use_container_width=True, hide_index=True,
        height=min(35 * (len(comp) + 1) + 3, 900),
    )

    table_total = comp["Total"].sum()
    chart_total = 0 if pd.isna(grand_total) else grand_total
    if chart_total > 0 and abs(table_total - chart_total) / chart_total > 0.001:
        st.error(f"⚠️ Reconciliation failure: ladder totals {numfmt(table_total)} vs. chart "
                 f"{numfmt(chart_total)} — trades are being dropped between steps.")
    else:
        st.caption(f"✓ Reconciled — ladder total {numfmt(table_total)} matches the chart.")

    proxy_sefs = [s for s in sefs_present if cat_data[cat_data["SEF"] == s]["Trades_Estimated"].any()]
    if proxy_sefs and mcol == "Trades":
        st.caption(f"⚠️ Trade counts for {', '.join(proxy_sefs)} are a row-count proxy — "
                   f"these sources publish no trade-count field.")

    # ── Individual trades — the one drill-down. Straight from the file. ──
    with st.expander(f"🔍 Individual {category_label} trades"):
        dcol1, dcol2, dcol3 = st.columns(3)
        with dcol1:
            drill_sef = st.selectbox("Source", safe_unique(cat_data["SEF"]),
                                     key=f"drill_sef_{category_label}")
        sef_df = cat_data[cat_data["SEF"] == drill_sef]
        with dcol2:
            bucket_opts = safe_sorted(sef_df["Tenor_Bucket"].unique(), key=tenor_sort_key)
            true_opts = safe_sorted(sef_df["Tenor"].astype(str).unique(), key=tenor_sort_key)
            sef_tenors = ["All tenors"] + list(dict.fromkeys(bucket_opts + true_opts))
            drill_tenor = st.selectbox("Tenor", sef_tenors, key=f"drill_tenor_{category_label}",
                                       format_func=tenor_display)
        with dcol3:
            group_rows = st.selectbox("View", ["Every row (raw file)", "Grouped by instrument"],
                                      key=f"drill_view_{category_label}",
                                      help="Grouped collapses identical descriptions into one line "
                                           "with notional summed — how a bucket total is built.")

        if drill_tenor == "All tenors":
            detail = sef_df
        else:
            detail = sef_df[(sef_df["Tenor_Bucket"].astype(str) == drill_tenor) |
                            (sef_df["Tenor"].astype(str) == drill_tenor)]
        detail = detail.copy()
        detail["Tenor"] = detail["Tenor"].astype(str)

        if group_rows == "Grouped by instrument":
            agg = {"Rows": ("Description", "size"),
                   "Notional": ("Notional_Local", "sum"),
                   "Trades": ("Trades", "sum")}
            if category_label == "IRS / Swap":
                agg["DV01_USD"] = ("DV01_USD", "sum")
            shown = (detail.groupby(["Description", "Tenor", "Currency"], as_index=False)
                     .agg(**agg).sort_values("Notional", ascending=False))
            fmt = {"Notional": numfmt, "Rows": "{:,.0f}", "Trades": "{:,.0f}"}
            if "DV01_USD" in shown.columns:
                fmt["DV01_USD"] = numfmt
            st.dataframe(shown.rename(columns={"Tenor": "True Tenor", "DV01_USD": "DV01 (USD)"})
                         .style.format(fmt), use_container_width=True, hide_index=True)
        else:
            cols_ = ["Description", "AssetClass", "Currency", "Tenor", "Tenor_Bucket",
                     "Last_Price", "Notional_Local", "DV01_USD", "DV01_Method", "Trades"]
            if category_label == "FX":
                cols_ = [c for c in cols_ if c not in ("DV01_USD", "DV01_Method")]
            cols_ = [c for c in cols_ if c in detail.columns]
            st.dataframe(
                detail[cols_].rename(columns={
                    "AssetClass": "Asset Class", "Tenor": "True Tenor",
                    "Tenor_Bucket": "Bucket", "Notional_Local": "Notional (from file)",
                    "DV01_USD": "DV01 (USD)", "DV01_Method": "DV01 Basis"}),
                use_container_width=True, hide_index=True)

        by_ccy = detail.groupby("Currency")["Notional_Local"].sum()
        summary = (f"{len(detail)} row(s) · Notional " +
                   " · ".join(f"{numfmt(v)} {c}" for c, v in by_ccy.items()))
        if category_label == "IRS / Swap" and "DV01_USD" in detail.columns:
            summary += f" · DV01 {numfmt(detail['DV01_USD'].sum())} USD"
        summary += f" · Trades {int(detail['Trades'].sum())}"
        st.caption(summary)
        st.caption("Notional is exactly what the source file publishes — nothing converted. "
                   "Asset class: IR = rates · CU = FX · CD = credit.")
        if detail["Trades_Estimated"].any():
            st.caption("⚠️ This source publishes daily aggregates, not individual prints.")


# ═════════════════════════════════════════════════════════════════════════
# MONTH-TO-DATE (each day valued on THAT day's grid)
# ═════════════════════════════════════════════════════════════════════════
# Full US market holidays — no SEF publishes a file on these, so MTD skips
# them instead of flagging them missing. (Jul 4, 2026 falls on a Saturday, so
# it was observed Friday Jul 3.) Extend as needed.
MARKET_HOLIDAYS = {"20260101", "20260119", "20260216", "20260403", "20260525",
                   "20260619", "20260703", "20260907", "20261126", "20261225"}


def _business_days_of_month(sel: datetime):
    d, days = sel.replace(day=1), []
    while d.date() <= sel.date():
        if d.weekday() < 5 and d.strftime("%Y%m%d") not in MARKET_HOLIDAYS:
            days.append(d)
        d += timedelta(days=1)
    return days


def _snapshot_sig() -> str:
    """Fingerprint of every saved day-file. Part of the MTD cache key, so the
    moment a backfill or a new snapshot lands, the MTD view rebuilds instead
    of serving a stale 'missing' list for up to 30 minutes."""
    names = []
    for d in (ICAP_SNAPSHOT_DIR, SEF_SNAPSHOT_DIR):
        if os.path.isdir(d):
            names.extend(sorted(os.listdir(d)))
    return str(hash(tuple(names)))


@st.cache_data(ttl=1800, show_spinner=False)
def build_mtd_data(sel_date_str: str, grid_sig: str, snap_sig: str):
    """Returns (dataframe, missing) where missing lists every SEF-day that
    couldn't be loaded — nothing fails silently anymore.

    Snapshot-first: each day's parsed table is saved to sef_snapshots/ the
    first time it's seen and pushed to the repo, so MTD reads from disk and
    only fetches genuine gaps. Failed fetches are NOT cached (see
    fetch_source_hist) — they retry on every rebuild. But note: some sites
    only keep a rolling window of recent daily files online (BGC's daily
    activity directory has been 'most recent weeks only' for years), so a day
    that rolled off before it was ever snapshotted can ONLY come back via
    backfill upload."""
    sel = datetime.strptime(sel_date_str, "%Y%m%d")
    frames, missing = [], []
    for d in _business_days_of_month(sel):
        ds = d.strftime("%Y%m%d")
        day_frames = []
        for name, cfg in SOURCES.items():
            if cfg["kind"] == "icap":
                snap = load_icap_snapshot(d)
                if snap is not None:
                    day_frames.append(snap)
                else:
                    missing.append(f"ICAP {d:%b %d}")
                continue
            snap = load_sef_snapshot(name, d)
            if snap is not None:
                day_frames.append(snap)
                continue
            try:
                df = fetch_source_hist(name, source_urls(cfg, d), cfg["kind"], ds)
                save_sef_snapshot(df, name, d)
                day_frames.append(df)
            except Exception:
                missing.append(f"{name} {d:%b %d}")
        if not day_frames:
            continue
        day = pd.concat(day_frames, ignore_index=True)
        day["Currency"] = (day["Currency"].astype(str).str.strip().str.upper()
                           .replace({"NAN": "N/A", "NONE": "N/A", "": "N/A", "<NA>": "N/A"}))
        day = day[day["MXN_Match"]]
        if day.empty:
            continue
        day_grid, _ = load_dv01_grid_for_date(d)     # each day on its OWN grid
        day = add_dv01(day, day_grid)
        day["Date"] = d.strftime("%Y-%m-%d")
        frames.append(day)
    return (pd.concat(frames, ignore_index=True) if frames else None), missing


mtd_data = None
if show_mtd:
    with st.spinner("Building month-to-date view (cached after first load)..."):
        try:
            grid_sig = str(sorted(dv01_grid.items()))
            mtd_data, mtd_missing = build_mtd_data(sel_date.strftime("%Y%m%d"), grid_sig,
                                                   _snapshot_sig())
            # Surfaced only inside the "How this works" panel, not as a banner.
            st.session_state["mtd_missing"] = mtd_missing
        except Exception as e:
            st.caption(f"Month-to-date view unavailable: {e}")

# ═════════════════════════════════════════════════════════════════════════
# MEXICO PRODUCTS — IRS / FX
# ═════════════════════════════════════════════════════════════════════════
st.markdown("---")
mexico_df = combined[combined["MXN_Match"]]
if mexico_df.empty:
    st.error("No Mexico products found in any source for this date.")
    st.stop()

sefs_present = sorted(mexico_df["SEF"].astype(str).unique(),
                      key=lambda s: list(SOURCES.keys()).index(s) if s in SOURCES else 99)
tab_irs, tab_fx = st.tabs(["📈 IRS / Swap", "💱 FX"])
with tab_irs:
    render_section(mexico_df, "IRS / Swap", sefs_present, "Mexico Products", mtd_data=mtd_data)
with tab_fx:
    render_section(mexico_df, "FX", sefs_present, "Mexico Products", mtd_data=mtd_data)

st.markdown("---")
st.caption("Methodology and data-handling notes are in the **How this works** panel at the top.")
