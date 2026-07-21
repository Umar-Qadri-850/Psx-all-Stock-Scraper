#!/usr/bin/env python3
"""
============================================================================
 PSX (Pakistan Stock Exchange) Real-Time Scraper + Data Dictionary Mapper
============================================================================

Scrapes live company data from https://dps.psx.com.pk/company/<SYMBOL>,
the PSX indices page, and (for --all) the market-watch page that lists
EVERY currently-listed PSX company in one shot. Normalizes everything into
rows matching your DB schema (stock_price_daily / market_data_daily /
stock_financials_quarterly / stock_ratios / stock_payouts /
corporate_actions), and compares against your Data Dictionary to show
which columns were populated vs which are NOT obtainable from public
pages (and why, per your own dictionary's "Access Method" notes).

USAGE
  python3 main.py PSO                      # one company, full detail
  python3 main.py PSO APL WAFI HTL HASCOL  # a specific list, full detail
  python3 main.py --symbols-file list.txt  # a list from a file
  python3 main.py --all                    # EVERY listed company, fast
                                            #   (1 request, price/volume only)
  python3 main.py --all --detailed         # EVERY listed company, full
                                            #   detail (500+ requests, slow)
  python3 main.py --all --limit 20         # test on the first 20 first

  Extra modules (each adds more per-company requests - only meaningful
  together with --detailed or an explicit symbol list):
  python3 main.py PSO --ratios             # normalized ratio table
  python3 main.py PSO --payouts            # dividend / bonus / rights history
  python3 main.py PSO --reports            # annual/quarterly report PDF links
  python3 main.py PSO --announcements      # PSX notices / corporate announcements
  python3 main.py PSO --financial-pdfs     # full financial statement PDF links
  python3 main.py PSO --corporate-actions  # splits / bonus / rights / mergers
  python3 main.py PSO --extras             # turns ALL of the above on at once

------------------------------------------------------------------------
LEGAL / ToS NOTICE - READ THIS
------------------------------------------------------------------------
dps.psx.com.pk carries an explicit Terms of Use + "Legal Notice" that:
  - Prohibits bots/spiders/scrapers "except where you have prior written
    permission"
  - Prohibits "systematic retrieval to create collections, compilations,
    databases or directories"
  - Prohibits commercial dissemination of market data without a license
    from PSX (contact: marketdatarequest@psx.com.pk)

This script is intended for PERSONAL / EDUCATIONAL / RESEARCH use only.
Do not run it at high frequency, do not resell or redistribute the data,
and do not use it as the data source for a commercial product. If you
need production-grade / licensed PSX data, contact PSX's Market Data
Team or use a licensed vendor (e.g. Capital Stake's official API).
============================================================================
"""

import os
import re
import sys
import json
import time
import argparse
from datetime import datetime, date
from io import StringIO
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

try:
    import pandas as pd
except ImportError:
    pd = None


BASE = "https://dps.psx.com.pk"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/json,*/*",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


# ============================================================================
# 1. DATA DICTIONARY  (embedded copy of the relevant sheets from your xlsx)
#    Tables reachable from a PSX company/market page are included:
#    stock_price_daily, market_data_daily, stock_financials_quarterly,
#    stock_ratios, stock_payouts, corporate_actions
# ============================================================================

DATA_DICTIONARY = {
    "stock_price_daily": {
        "adjusted_open":     {"desc": "Adjusted opening price",  "unit": "PKR/share",
                               "access": "Downloadable authenticated web interface (NOT public scrape)"},
        "adjusted_high":     {"desc": "Adjusted high price",     "unit": "PKR/share",
                               "access": "Downloadable authenticated web interface (NOT public scrape)"},
        "adjusted_low":      {"desc": "Adjusted low price",      "unit": "PKR/share",
                               "access": "Downloadable authenticated web interface (NOT public scrape)"},
        "adjusted_price":    {"desc": "Adjusted closing price",  "unit": "PKR/share",
                               "access": "Downloadable authenticated web interface (NOT public scrape)"},
        "unadjusted_price":  {"desc": "Raw closing price",       "unit": "PKR/share",
                               "access": "Public company/quote page"},
        "mkt_cap_mn":        {"desc": "Market capitalization",   "unit": "PKR mn",
                               "access": "Public company/quote page (Equity Profile)"},
        "volume_000":        {"desc": "Trading volume",          "unit": "000 Shares",
                               "access": "Public company/quote page"},
        "eps_ttm":           {"desc": "Trailing EPS",            "unit": "PKR/share",
                               "access": "Web interface (Financial Statements) - browser automation"},
        "eps_ttm_cons":      {"desc": "Consensus EPS",           "unit": "PKR/share",
                               "access": "UNKNOWN - no confirmed source (flag, do not invent)"},
        "dps_ttm":           {"desc": "Trailing DPS",            "unit": "PKR/share",
                               "access": "Company Reports - browser automation"},
    },
    "market_data_daily": {
        "kse_100_close":     {"desc": "KSE-100 closing index",   "unit": "Index Points",
                               "access": "Public indices page"},
        "pe_ratio":          {"desc": "Market P/E ratio",        "unit": "Ratio",
                               "access": "KSE100 Market Data workbook (Zakheera) - not on public page"},
        "dividend_yield":    {"desc": "Market dividend yield",   "unit": "Fraction",
                               "access": "KSE100 Market Data workbook (Zakheera) - not on public page"},
        "pkrv_12m":          {"desc": "12M PKRV yield",          "unit": "Fraction",
                               "access": "KSE100 Market Data workbook (Zakheera) - not on public page"},
        "market_cap_mn":     {"desc": "Total market capitalization", "unit": "PKR mn",
                               "access": "KSE100 Market Data workbook (Zakheera) - not on public page"},
    },
    "stock_financials_quarterly": {
        "symbol":            {"desc": "Company symbol",              "unit": "Text"},
        "quarter_end":       {"desc": "Quarter end date",             "unit": "Date"},
        "statement":         {"desc": "Statement type",               "unit": "Text"},
        "line_item":         {"desc": "Financial statement line item","unit": "Various"},
        "value":             {"desc": "Line item value",              "unit": "Various"},
        "consolidated":      {"desc": "Consolidation flag",           "unit": "Boolean"},
    },
    # ------------------------------------------------------------------
    # New: normalized ratio table (was "Ratios - Basic only" in the gap
    # analysis). Column names are the normalized keys the scraper tries
    # to detect on the company page's Ratios widget.
    # ------------------------------------------------------------------
    "stock_ratios": {
        "gross_margin":      {"desc": "Gross margin",             "unit": "Fraction",
                               "access": "Public company page - Ratios widget"},
        "operating_margin":  {"desc": "Operating margin",         "unit": "Fraction",
                               "access": "Public company page - Ratios widget"},
        "net_margin":        {"desc": "Net margin",               "unit": "Fraction",
                               "access": "Public company page - Ratios widget"},
        "roe":                {"desc": "Return on equity",         "unit": "Fraction",
                               "access": "Public company page - Ratios widget"},
        "roa":                {"desc": "Return on assets",         "unit": "Fraction",
                               "access": "Public company page - Ratios widget"},
        "current_ratio":      {"desc": "Current ratio",            "unit": "Ratio",
                               "access": "Public company page - Ratios widget"},
        "quick_ratio":        {"desc": "Quick ratio",               "unit": "Ratio",
                               "access": "Public company page - Ratios widget"},
        "debt_ratio":         {"desc": "Debt ratio",                "unit": "Ratio",
                               "access": "Public company page - Ratios widget"},
        "peg_ratio":          {"desc": "PEG ratio",                 "unit": "Ratio",
                               "access": "Public company page - Ratios widget"},
        "price_to_book":      {"desc": "Price / Book value",        "unit": "Ratio",
                               "access": "Public company page - Ratios widget"},
    },
    # ------------------------------------------------------------------
    # New: payouts (dividend / bonus / rights history)
    # ------------------------------------------------------------------
    "stock_payouts": {
        "symbol":             {"desc": "Company symbol",           "unit": "Text"},
        "announcement_date":  {"desc": "Announcement date",        "unit": "Date"},
        "book_closure_start": {"desc": "Book closure start date",  "unit": "Date"},
        "book_closure_end":   {"desc": "Book closure end date",    "unit": "Date"},
        "dividend_pct":       {"desc": "Cash dividend",            "unit": "Percent"},
        "bonus_pct":          {"desc": "Bonus shares",             "unit": "Percent"},
        "right_pct":          {"desc": "Right shares",             "unit": "Percent"},
        "period":             {"desc": "Period the payout relates to", "unit": "Text"},
    },
    # ------------------------------------------------------------------
    # New: corporate actions (splits / bonus / rights / mergers)
    # ------------------------------------------------------------------
    "corporate_actions": {
        "symbol":       {"desc": "Company symbol",   "unit": "Text"},
        "action_date":  {"desc": "Action date",       "unit": "Date"},
        "action_type":  {"desc": "Split/Bonus/Right/Merger", "unit": "Text"},
        "detail":       {"desc": "Free-text detail of the action", "unit": "Text"},
    },
}


# ============================================================================
# 2. LOW LEVEL FETCHERS
# ============================================================================

def get_html(url: str, retries: int = 3, delay: float = 1.5) -> str:
    """GET a URL and return raw HTML text, with basic retry."""
    last_err = None
    for attempt in range(retries):
        try:
            resp = SESSION.get(url, timeout=15)
            resp.raise_for_status()
            return resp.text
        except requests.RequestException as e:
            last_err = e
            time.sleep(delay)
    raise RuntimeError(f"Failed to GET {url}: {last_err}")


def get_html_soft(url: str, retries: int = 2, delay: float = 1.0):
    """Like get_html but never raises - returns None on failure. Used for
    the newer, less-certain endpoints (payouts/announcements/reports) so a
    missing/renamed page doesn't kill the whole run.
    """
    try:
        return get_html(url, retries=retries, delay=delay)
    except Exception as e:
        print(f"  [warn] could not fetch {url}: {e}", file=sys.stderr)
        return None


def get_json(url: str, retries: int = 3, delay: float = 1.5):
    """GET a URL and return parsed JSON (used for timeseries endpoints)."""
    last_err = None
    for attempt in range(retries):
        try:
            resp = SESSION.get(url, timeout=15)
            resp.raise_for_status()
            return resp.json()
        except (requests.RequestException, ValueError) as e:
            last_err = e
            time.sleep(delay)
    print(f"  [warn] could not fetch JSON from {url}: {last_err}", file=sys.stderr)
    return None


def text_lines(soup: BeautifulSoup):
    """Flatten page text into a clean list of non-empty lines, in DOM order.
    This is deliberately layout-based (not CSS-class based) so it keeps
    working even if PSX changes their class names / JS framework.
    """
    raw = soup.get_text("\n")
    return [ln.strip() for ln in raw.split("\n") if ln.strip()]


def value_after(lines, label, occurrence=0, cast=str):
    """Find the n-th line that exactly matches `label` and return the very
    next non-empty line, cast to the given type. Returns None if not found.
    """
    hits = 0
    for i, ln in enumerate(lines):
        if ln == label:
            if hits == occurrence:
                if i + 1 < len(lines):
                    val = lines[i + 1]
                    try:
                        return cast(val)
                    except (ValueError, TypeError):
                        return val
                return None
            hits += 1
    return None


def to_float(s):
    if s is None:
        return None
    s = str(s).replace(",", "").replace("Rs.", "").strip()
    m = re.search(r"-?\d+(\.\d+)?", s)
    return float(m.group()) if m else None


def to_int(s):
    f = to_float(s)
    return int(f) if f is not None else None


def to_pct_fraction(s):
    """'12.5%' or '12.5' -> 0.125 (fraction). Returns None if unparsable."""
    f = to_float(s)
    if f is None:
        return None
    return round(f / 100.0, 6)


DATE_PATTERNS = [
    r"\d{4}-\d{2}-\d{2}",
    r"\d{2}-[A-Za-z]{3}-\d{4}",
    r"\d{2}/\d{2}/\d{4}",
    r"[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}",
]


def find_date(s):
    """Best-effort extraction of a date-looking substring; returns the raw
    matched text (not parsed to a date object) since PSX's own formats vary
    across widgets."""
    if not s:
        return None
    for pat in DATE_PATTERNS:
        m = re.search(pat, str(s))
        if m:
            return m.group()
    return None


# ============================================================================
# 3. SCRAPERS
# ============================================================================

def scrape_company_quote(symbol: str) -> dict:
    """Scrape the live quote block from a company page (REG / ready market)."""
    url = f"{BASE}/company/{symbol.upper()}"
    html = get_html(url)
    soup = BeautifulSoup(html, "html.parser")
    lines = text_lines(soup)

    data = {"symbol": symbol.upper(), "source_url": url,
            "fetched_at": datetime.now().isoformat(timespec="seconds")}

    # --- Company name / sector / current price -----------------------------
    price_idx = None
    for i, ln in enumerate(lines):
        if re.match(r"^Rs\.\d", ln):
            price_idx = i
            break

    if price_idx is not None:
        data["company_name"] = lines[price_idx - 2] if price_idx >= 2 else None
        data["sector"] = lines[price_idx - 1] if price_idx >= 1 else None
        data["current_price"] = to_float(lines[price_idx])
        # next two lines are typically: change, (change%)
        if price_idx + 1 < len(lines):
            data["change"] = to_float(lines[price_idx + 1])
        if price_idx + 2 < len(lines):
            data["change_pct"] = to_float(lines[price_idx + 2])

    # --- REG panel: Open / High / Low / Volume (first occurrence = REG) ----
    data["open"] = to_float(value_after(lines, "Open", occurrence=0))
    data["high"] = to_float(value_after(lines, "High", occurrence=0))
    data["low"] = to_float(value_after(lines, "Low", occurrence=0))
    data["volume"] = to_int(value_after(lines, "Volume", occurrence=0))
    data["ldcp"] = to_float(value_after(lines, "LDCP", occurrence=0))
    data["pe_ratio_ttm"] = to_float(
        next((lines[i + 1] for i, ln in enumerate(lines)
              if ln.startswith("P/E Ratio")), None)
    )

    # --- Day range / circuit breaker / 52 week range ------------------------
    for i, ln in enumerate(lines):
        if ln.startswith("CIRCUIT BREAKER") and "circuit_low" not in data:
            m = re.search(r"([\d,]+\.\d+)\s*—\s*([\d,]+\.\d+)", lines[i + 1] if i + 1 < len(lines) else "")
            if m:
                data["circuit_low"] = to_float(m.group(1))
                data["circuit_high"] = to_float(m.group(2))
        if ln.startswith("52-WEEK RANGE") and "week52_low" not in data:
            m = re.search(r"([\d,]+\.\d+)\s*—\s*([\d,]+\.\d+)", lines[i + 1] if i + 1 < len(lines) else "")
            if m:
                data["week52_low"] = to_float(m.group(1))
                data["week52_high"] = to_float(m.group(2))

    # --- Equity profile -------------------------------------------------
    data["market_cap_000"] = to_float(value_after(lines, "Market Cap (000's)"))
    data["shares"] = to_int(value_after(lines, "Shares"))
    data["free_float_shares"] = to_int(value_after(lines, "Free Float", occurrence=0))
    data["free_float_pct"] = to_float(value_after(lines, "Free Float", occurrence=1))

    return data


def scrape_company_financials(symbol: str) -> dict:
    """Scrape Annual + Quarterly financial summary tables (Sales / PAT / EPS)
    and the Ratios table shown on the company page, using pandas.read_html.
    Returns dict with 'annual', 'quarterly', 'ratios' DataFrames (as records).
    """
    if pd is None:
        print("  [warn] pandas not installed - skipping financials tables", file=sys.stderr)
        return {}

    url = f"{BASE}/company/{symbol.upper()}"
    html = get_html(url)

    try:
        tables = pd.read_html(StringIO(html))
    except ValueError:
        tables = []

    result = {}
    for t in tables:
        cols = [str(c) for c in t.columns]
        first_col_vals = t.iloc[:, 0].astype(str).tolist()

        looks_financial = any(
            "Sales" in v or "Profit after Taxation" in v or "EPS" in v
            for v in first_col_vals
        )
        looks_quarterly = any(re.match(r"^Q\d", c) for c in cols)
        looks_annual = any(re.match(r"^(19|20)\d{2}$", c) for c in cols)
        looks_ratios = any("Margin" in v or "PEG" in v for v in first_col_vals)

        if looks_financial and looks_quarterly:
            result["quarterly"] = t
        elif looks_financial and looks_annual:
            result["annual"] = t
        elif looks_ratios:
            result["ratios"] = t

    return result


def scrape_kse100() -> dict:
    """Scrape the market indices page for the KSE-100 index row."""
    url = f"{BASE}/indices"
    html = get_html(url)

    data = {"source_url": url, "fetched_at": datetime.now().isoformat(timespec="seconds")}

    if pd is not None:
        try:
            tables = pd.read_html(StringIO(html))
            for t in tables:
                if "Index" in [str(c) for c in t.columns]:
                    row = t[t["Index"].astype(str).str.contains("KSE100", na=False)
                             & ~t["Index"].astype(str).str.contains("PR", na=False)]
                    if not row.empty:
                        r = row.iloc[0]
                        data["kse_100_high"] = to_float(r.get("High"))
                        data["kse_100_low"] = to_float(r.get("Low"))
                        data["kse_100_close"] = to_float(r.get("Current"))
                        data["kse_100_change"] = to_float(r.get("Change"))
                        data["kse_100_change_pct"] = to_float(r.get("% Change"))
                    break
        except ValueError:
            pass

    # Fallback: text-based parse of the scrolling ticker bar
    if "kse_100_close" not in data:
        soup = BeautifulSoup(html, "html.parser")
        lines = text_lines(soup)
        data["kse_100_close"] = to_float(value_after(lines, "KSE100"))

    return data


def fetch_eod_timeseries(symbol: str, limit: int = 30):
    """Fetch end-of-day (close, volume) history from PSX's own JSON endpoint.
    Returns list of {date, close, volume} dicts, most recent last.
    NOTE: this is UNADJUSTED close - PSX does not expose adjusted history
    through this public endpoint (see DATA_DICTIONARY notes).
    """
    url = f"{BASE}/timeseries/eod/{symbol.upper()}"
    payload = get_json(url)
    if not payload or "data" not in payload:
        return []

    rows = []
    for point in payload["data"][:limit]:
        # PSX format is typically [unix_timestamp, close_price, volume]
        try:
            ts, close, volume = point[0], point[1], point[2]
            rows.append({
                "date": datetime.fromtimestamp(ts).date().isoformat(),
                "unadjusted_price": to_float(close),
                "volume_000": round(to_float(volume) / 1000, 3) if volume else None,
            })
        except (IndexError, TypeError, OSError):
            continue
    return rows


def scrape_market_watch_all() -> list:
    """Scrape https://dps.psx.com.pk/market-watch — PSX's own single-page
    table listing EVERY currently listed symbol (~500+) with live OHLCV,
    LDCP, change and volume, all in ONE request.

    This is the fast/light way to get "all companies" data: one page load
    instead of 500+. Returns a list of dicts:
        {symbol, sector_code, listed_in, ldcp, open, high, low,
         current, change, change_pct, volume}
    """
    url = f"{BASE}/market-watch"
    html = get_html(url)

    rows = []
    if pd is not None:
        try:
            tables = pd.read_html(StringIO(html))
            for t in tables:
                cols = [str(c).strip().upper() for c in t.columns]
                if "SYMBOL" in cols and "CURRENT" in cols:
                    t.columns = cols
                    for _, r in t.iterrows():
                        # symbol cell sometimes carries a trailing flag badge
                        # (e.g. "HIRAT NC", "MZNPETF XD") - keep the first token
                        sym_raw = str(r.get("SYMBOL", "")).strip()
                        symbol = sym_raw.split()[0].upper() if sym_raw else None
                        if not symbol or symbol in ("NAN", ""):
                            continue
                        rows.append({
                            "symbol": symbol,
                            "sector_code": str(r.get("SECTOR", "")).strip(),
                            "listed_in": str(r.get("LISTED IN", "")).strip(),
                            "ldcp": to_float(r.get("LDCP")),
                            "open": to_float(r.get("OPEN")),
                            "high": to_float(r.get("HIGH")),
                            "low": to_float(r.get("LOW")),
                            "current": to_float(r.get("CURRENT")),
                            "change": to_float(r.get("CHANGE")),
                            "change_pct": to_float(r.get("CHANGE (%)")),
                            "volume": to_int(r.get("VOLUME")),
                        })
                    break
        except ValueError:
            pass

    return rows


def fetch_intraday_timeseries(symbol: str):
    """Fetch intraday tick data from PSX's own JSON endpoint."""
    url = f"{BASE}/timeseries/int/{symbol.upper()}"
    payload = get_json(url)
    if not payload or "data" not in payload:
        return []
    ticks = []
    for point in payload["data"]:
        try:
            ts, price, volume = point[0], point[1], point[2]
            ticks.append({
                "time": datetime.fromtimestamp(ts).isoformat(timespec="seconds"),
                "price": to_float(price),
                "volume": to_int(volume),
            })
        except (IndexError, TypeError, OSError):
            continue
    return ticks


# ----------------------------------------------------------------------
# 3b. NEW SCRAPERS - Ratios / Payouts / Reports / Announcements /
#     Financial-statement PDFs / Corporate actions
#
# NOTE ON RELIABILITY: unlike the quote/market-watch/timeseries endpoints
# above (which are stable JSON/HTML PSX itself renders its price widgets
# from), these six modules read secondary tabs/widgets on the company page
# whose exact DOM layout is more likely to shift over time and could not
# be verified against a live fetch while writing this (this environment's
# network egress does not reach dps.psx.com.pk). Every function here is
# written defensively (try/except + graceful "not found" returns) and
# tries several matching strategies so it degrades to an empty list/dict
# rather than crashing if PSX's markup differs from what's assumed. Re-run
# with a sample symbol and inspect the raw HTML if a module comes back
# empty, and adjust the label/keyword lists near the top of each function.
# ----------------------------------------------------------------------

RATIO_LABEL_MAP = {
    "gross_margin":     ["Gross Margin"],
    "operating_margin": ["Operating Margin"],
    "net_margin":       ["Net Margin", "Net Profit Margin"],
    "roe":               ["ROE", "Return on Equity"],
    "roa":               ["ROA", "Return on Assets"],
    "current_ratio":     ["Current Ratio"],
    "quick_ratio":       ["Quick Ratio", "Acid Test Ratio"],
    "debt_ratio":        ["Debt Ratio", "Debt to Equity", "D/E Ratio"],
    "peg_ratio":         ["PEG", "PEG Ratio"],
    "price_to_book":     ["Price/Book", "Price to Book", "P/B Ratio", "P/BV"],
}


def scrape_company_ratios(symbol: str, financials: dict = None) -> dict:
    """Normalize the Ratios widget on the company page into flat, typed
    keys (gross_margin, operating_margin, roe, ...). Reuses the raw
    'ratios' DataFrame from scrape_company_financials() if the caller
    already fetched it, to avoid a duplicate request.
    """
    result = {"symbol": symbol.upper()}
    if financials is None:
        financials = scrape_company_financials(symbol)

    df = financials.get("ratios") if financials else None
    if df is None or pd is None:
        return result

    label_col = df.columns[0]
    value_col = df.columns[1] if len(df.columns) > 1 else None
    if value_col is None:
        return result

    for _, row in df.iterrows():
        label = str(row[label_col]).strip()
        raw_val = row[value_col]
        for norm_key, aliases in RATIO_LABEL_MAP.items():
            if any(alias.lower() in label.lower() for alias in aliases):
                is_pct = "margin" in norm_key or norm_key in ("roe", "roa")
                result[norm_key] = to_pct_fraction(raw_val) if is_pct else to_float(raw_val)
                break

    return result


def scrape_company_payouts(symbol: str) -> list:
    """Scrape the dividend / bonus / rights payout history from the
    company page's 'Payout' tab. PSX renders this as an HTML table with
    columns roughly: Announcement Date, Book Closure, Dividend/Bonus/
    Right(%), Period. Falls back gracefully to [] if the table isn't
    found (e.g. widget renamed or a company with no payout history).
    """
    url = f"{BASE}/company/{symbol.upper()}"
    html = get_html_soft(url)
    if not html or pd is None:
        return []

    rows_out = []
    try:
        tables = pd.read_html(StringIO(html))
    except ValueError:
        tables = []

    for t in tables:
        cols_upper = [str(c).strip().upper() for c in t.columns]
        looks_payout = any("BOOK CLOSURE" in c for c in cols_upper) or (
            any("DIVIDEND" in c for c in cols_upper)
            and any("BONUS" in c or "RIGHT" in c for c in cols_upper)
        )
        if not looks_payout:
            continue

        t.columns = cols_upper
        for _, r in t.iterrows():
            book_closure_raw = str(r.get("BOOK CLOSURE", r.get("BOOK CLOSURE DATE", "")))
            bc_dates = re.findall(r"[\d]{1,2}[-/][A-Za-z\d]{1,3}[-/]\d{2,4}", book_closure_raw)
            rows_out.append({
                "symbol": symbol.upper(),
                "announcement_date": find_date(r.get("ANNOUNCEMENT DATE") or r.get("DATE")),
                "book_closure_start": bc_dates[0] if len(bc_dates) >= 1 else None,
                "book_closure_end": bc_dates[1] if len(bc_dates) >= 2 else None,
                "dividend_pct": to_float(r.get("DIVIDEND") or r.get("DIVIDEND (%)") or r.get("CASH DIVIDEND")),
                "bonus_pct": to_float(r.get("BONUS") or r.get("BONUS (%)")),
                "right_pct": to_float(r.get("RIGHT") or r.get("RIGHT (%)")),
                "period": str(r.get("PERIOD", "")).strip() or None,
            })
        break  # first matching table only

    return rows_out


def _extract_pdf_links(html: str, keyword_filter=None) -> list:
    """Shared helper: pull every <a href=...pdf> from a page, resolve to an
    absolute URL, and optionally keep only links whose link text OR href
    contains one of `keyword_filter` (case-insensitive substrings).
    Returns list of {title, url}.
    """
    soup = BeautifulSoup(html, "html.parser")
    out = []
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if ".pdf" not in href.lower():
            continue
        abs_url = urljoin(BASE, href)
        if abs_url in seen:
            continue
        title = a.get_text(strip=True) or os.path.basename(href)
        if keyword_filter:
            hay = f"{title} {href}".lower()
            if not any(k.lower() in hay for k in keyword_filter):
                continue
        seen.add(abs_url)
        out.append({"title": title, "url": abs_url})
    return out


def scrape_company_reports(symbol: str) -> list:
    """Collect Annual/Quarterly Report PDF links from the company page's
    'Financial Reports' / 'Reports' tab. Returns list of {title, url}.
    """
    url = f"{BASE}/company/{symbol.upper()}"
    html = get_html_soft(url)
    if not html:
        return []
    return _extract_pdf_links(
        html, keyword_filter=["annual report", "quarterly report", "report"]
    )


def scrape_company_financial_statement_pdfs(symbol: str) -> list:
    """Collect financial-statement PDF links (balance sheet, P&L, cash flow,
    etc.) separately from generic annual/quarterly report PDFs.
    """
    url = f"{BASE}/company/{symbol.upper()}"
    html = get_html_soft(url)
    if not html:
        return []
    return _extract_pdf_links(
        html,
        keyword_filter=[
            "financial statement", "balance sheet", "profit", "cash flow",
            "income statement", "accounts",
        ],
    )


def scrape_company_announcements(symbol: str) -> list:
    """Scrape PSX corporate announcements/notices for a company. PSX
    exposes an 'Announcements' tab on the company page; when present it is
    read as an HTML table (Date, Subject, [PDF link]). If no such table is
    found, falls back to collecting any remaining un-filtered PDF links on
    the page as a best-effort proxy for "announcement documents".
    """
    url = f"{BASE}/company/{symbol.upper()}"
    html = get_html_soft(url)
    if not html:
        return []

    rows_out = []
    if pd is not None:
        try:
            tables = pd.read_html(StringIO(html))
        except ValueError:
            tables = []

        for t in tables:
            cols_upper = [str(c).strip().upper() for c in t.columns]
            looks_announcement = any("SUBJECT" in c for c in cols_upper) or any(
                "ANNOUNCEMENT" in c for c in cols_upper
            )
            if not looks_announcement:
                continue
            t.columns = cols_upper
            for _, r in t.iterrows():
                rows_out.append({
                    "symbol": symbol.upper(),
                    "date": find_date(r.get("DATE") or r.get("ANNOUNCEMENT DATE")),
                    "subject": str(r.get("SUBJECT", r.get("ANNOUNCEMENT", ""))).strip() or None,
                    "category": str(r.get("CATEGORY", "")).strip() or None,
                })
            break

    if not rows_out:
        # best-effort fallback: any PDF link that looks like a notice
        rows_out = [
            {"symbol": symbol.upper(), "date": None, "subject": link["title"],
             "category": "unclassified_pdf", "url": link["url"]}
            for link in _extract_pdf_links(html, keyword_filter=["notice", "announcement"])
        ]

    return rows_out


CORPORATE_ACTION_KEYWORDS = {
    "split":  ["split", "share split"],
    "bonus":  ["bonus"],
    "right":  ["right issue", "rights issue", "right shares"],
    "merger": ["merger", "amalgamation", "acquisition"],
}


def scrape_company_corporate_actions(symbol: str, payouts: list = None) -> list:
    """Corporate actions (splits/bonus/rights/mergers). PSX's public site
    doesn't expose a single dedicated corporate-actions feed, so this is
    built as a derived view: bonus/right entries come straight out of the
    payouts table (reusing it if already fetched to avoid a duplicate
    request), and splits/mergers are looked for as keyword matches inside
    any announcements found on the page.
    """
    actions = []

    if payouts is None:
        payouts = scrape_company_payouts(symbol)
    for p in payouts:
        if p.get("bonus_pct"):
            actions.append({
                "symbol": symbol.upper(), "action_date": p.get("announcement_date"),
                "action_type": "bonus", "detail": f"{p['bonus_pct']}% bonus, period {p.get('period')}",
            })
        if p.get("right_pct"):
            actions.append({
                "symbol": symbol.upper(), "action_date": p.get("announcement_date"),
                "action_type": "right", "detail": f"{p['right_pct']}% right, period {p.get('period')}",
            })

    announcements = scrape_company_announcements(symbol)
    for a in announcements:
        subject = (a.get("subject") or "").lower()
        for action_type, keywords in CORPORATE_ACTION_KEYWORDS.items():
            if action_type in ("bonus", "right"):
                continue  # already covered via payouts above
            if any(k in subject for k in keywords):
                actions.append({
                    "symbol": symbol.upper(), "action_date": a.get("date"),
                    "action_type": action_type, "detail": a.get("subject"),
                })

    return actions


# ============================================================================
# 4. MAP SCRAPED DATA -> YOUR DB SCHEMA
# ============================================================================

def build_stock_price_daily_row(quote: dict) -> dict:
    """Map scraped quote fields onto stock_price_daily columns.
    Fields the public page cannot provide are explicitly set to None.
    """
    return {
        "symbol": quote.get("symbol"),
        "date": date.today().isoformat(),
        "adjusted_open": None,          # not available via public scrape
        "adjusted_high": None,          # not available via public scrape
        "adjusted_low": None,           # not available via public scrape
        "adjusted_price": None,         # not available via public scrape
        "unadjusted_price": quote.get("current_price"),
        "mkt_cap_mn": round(quote["market_cap_000"] / 1000, 3)
                       if quote.get("market_cap_000") is not None else None,
        "volume_000": round(quote["volume"] / 1000, 3)
                       if quote.get("volume") is not None else None,
        "eps_ttm": None,                # needs financial statements page/automation
        "eps_ttm_cons": None,           # source unresolved (per dictionary)
        "dps_ttm": None,                # needs company reports
    }


def build_market_data_daily_row(kse: dict) -> dict:
    return {
        "date": date.today().isoformat(),
        "kse_100_close": kse.get("kse_100_close"),
        "pe_ratio": None,        # needs KSE100 Market Data workbook
        "dividend_yield": None,  # needs KSE100 Market Data workbook
        "pkrv_12m": None,        # needs KSE100 Market Data workbook
        "market_cap_mn": None,   # needs KSE100 Market Data workbook
    }


def build_financials_rows(symbol: str, financials: dict) -> list:
    """Flatten the Annual + Quarterly tables into long format matching
    stock_financials_quarterly (symbol, quarter_end, statement, line_item, value).
    """
    rows = []
    for period_type, df in financials.items():
        if period_type == "ratios" or df is None or pd is None:
            continue
        label_col = df.columns[0]
        for _, r in df.iterrows():
            line_item = str(r[label_col])
            for col in df.columns[1:]:
                rows.append({
                    "symbol": symbol.upper(),
                    "quarter_end": str(col),
                    "statement": f"summary_{period_type}",
                    "line_item": line_item,
                    "value": r[col],
                    "consolidated": None,  # not disclosed on this summary widget
                })
    return rows


def build_stock_ratios_row(ratios: dict) -> dict:
    """ratios dict already comes back with normalized keys from
    scrape_company_ratios(); this just ensures every dictionary column is
    present (None if not found) so compare_with_dictionary() works on it.
    """
    row = {"symbol": ratios.get("symbol")}
    for col in DATA_DICTIONARY["stock_ratios"]:
        row[col] = ratios.get(col)
    return row


def build_dps_ttm(payouts: list):
    """Sum the most recent 4 quarters/periods of cash dividend as a rough
    trailing-12-month DPS estimate, using the payouts table (this is what
    fills in the previously-empty dps_ttm column when --payouts is on).
    """
    if not payouts:
        return None
    divs = [p["dividend_pct"] for p in payouts[:4] if p.get("dividend_pct") is not None]
    return round(sum(divs), 4) if divs else None


# ============================================================================
# 5. COMPARE SCRAPED RESULT AGAINST THE DATA DICTIONARY
# ============================================================================

def compare_with_dictionary(table: str, row: dict) -> dict:
    """Return {matched: [...], missing: [...]} for a given DB table, based
    on whether the scraper actually produced a non-null value."""
    schema = DATA_DICTIONARY.get(table, {})
    matched, missing = [], []
    for col, meta in schema.items():
        val = row.get(col)
        entry = {"column": col, "description": meta["desc"], "unit": meta.get("unit"),
                  "value": val}
        if val is not None:
            matched.append(entry)
        else:
            entry["why_missing"] = meta.get("access", "not returned by scraper")
            missing.append(entry)
    return {"table": table, "matched": matched, "missing": missing}


def print_comparison(report: dict):
    table = report["table"]
    print(f"\n=== {table} — Data Dictionary Comparison ===")
    print(f"  Matched ({len(report['matched'])}):")
    for e in report["matched"]:
        print(f"    ✔ {e['column']:<20} = {e['value']!r:<20} [{e['unit']}]")
    print(f"  Missing / not available from live scrape ({len(report['missing'])}):")
    for e in report["missing"]:
        print(f"    ✘ {e['column']:<20} -> {e['why_missing']}")


# ============================================================================
# 6. MAIN / CLI
# ============================================================================

def quote_from_watch_row(row: dict) -> dict:
    """Adapt a market-watch table row into the same shape scrape_company_quote()
    returns, so it can feed straight into build_stock_price_daily_row(). Fields
    the market-watch table doesn't carry (market cap, shares, free float, P/E)
    are left unset - only obtainable via the per-company page (--detailed).
    """
    return {
        "symbol": row["symbol"],
        "source_url": f"{BASE}/market-watch",
        "fetched_at": datetime.now().isoformat(timespec="seconds"),
        "current_price": row.get("current"),
        "open": row.get("open"),
        "high": row.get("high"),
        "low": row.get("low"),
        "volume": row.get("volume"),
        "ldcp": row.get("ldcp"),
        "change": row.get("change"),
        "change_pct": row.get("change_pct"),
        "sector_code": row.get("sector_code"),
        "listed_in": row.get("listed_in"),
    }


def run_one(symbol: str, kse: dict, out_dir: str, history_days: int = 0,
            quote: dict = None, fetch_details: bool = True, extras: dict = None) -> dict:
    """Build + compare + save the snapshot for a single company.

    If `quote` is already provided (e.g. from the market-watch table), no
    quote-page network call is made. `fetch_details=False` also skips the
    financials-table and EOD-history network calls (used in fast --all mode).

    `extras` is a dict of booleans controlling the new optional modules:
        {"ratios": bool, "payouts": bool, "reports": bool,
         "announcements": bool, "financial_pdfs": bool,
         "corporate_actions": bool}
    Extras are only fetched when fetch_details is also True, since they all
    require a per-company page load.
    """
    extras = extras or {}
    print(f"--- {symbol.upper()} ---" + ("" if fetch_details else "  [fast/market-watch mode]"))

    if quote is None:
        quote = scrape_company_quote(symbol)
    print(f"  quote: price={quote.get('current_price')} "
          f"volume={quote.get('volume')} mkt_cap(000)={quote.get('market_cap_000')}")

    financials, history = {}, []
    ratios_row, payouts_rows, reports_rows = {}, [], []
    announcements_rows, financial_pdf_rows, corp_action_rows = [], [], []

    if fetch_details:
        financials = scrape_company_financials(symbol)
        print(f"  financials tables found: {list(financials.keys())}")
        if history_days:
            history = fetch_eod_timeseries(symbol, limit=history_days)
            print(f"  EOD history points fetched: {len(history)}")

        if extras.get("ratios"):
            ratios = scrape_company_ratios(symbol, financials=financials)
            ratios_row = build_stock_ratios_row(ratios)
            print(f"  ratios normalized: "
                  f"{sum(1 for v in ratios_row.values() if v is not None) - 1} fields")

        if extras.get("payouts") or extras.get("corporate_actions"):
            payouts_rows = scrape_company_payouts(symbol)
            print(f"  payout records found: {len(payouts_rows)}")

        if extras.get("reports"):
            reports_rows = scrape_company_reports(symbol)
            print(f"  report PDF links found: {len(reports_rows)}")

        if extras.get("announcements") or extras.get("corporate_actions"):
            announcements_rows = scrape_company_announcements(symbol)
            print(f"  announcements found: {len(announcements_rows)}")

        if extras.get("financial_pdfs"):
            financial_pdf_rows = scrape_company_financial_statement_pdfs(symbol)
            print(f"  financial statement PDF links found: {len(financial_pdf_rows)}")

        if extras.get("corporate_actions"):
            corp_action_rows = scrape_company_corporate_actions(symbol, payouts=payouts_rows)
            print(f"  corporate actions derived: {len(corp_action_rows)}")

    # ---- map onto DB schema -------------------------------------------------
    price_row = build_stock_price_daily_row(quote)
    if extras.get("payouts") and payouts_rows:
        price_row["dps_ttm"] = build_dps_ttm(payouts_rows)
    market_row = build_market_data_daily_row(kse)
    financial_rows = build_financials_rows(symbol, financials)

    # ---- compare vs data dictionary -----------------------------------------
    reports = [
        compare_with_dictionary("stock_price_daily", price_row),
        compare_with_dictionary("market_data_daily", market_row),
    ]
    if extras.get("ratios"):
        reports.append(compare_with_dictionary("stock_ratios", ratios_row))

    # ---- save per-company snapshot -------------------------------------------
    os.makedirs(out_dir, exist_ok=True)
    out_payload = {
        "quote_raw": quote,
        "kse100_raw": kse,
        "stock_price_daily_row": price_row,
        "market_data_daily_row": market_row,
        "stock_financials_quarterly_rows": financial_rows,
        "stock_ratios_row": ratios_row,
        "stock_payouts_rows": payouts_rows,
        "reports_pdf_links": reports_rows,
        "announcements_rows": announcements_rows,
        "financial_statement_pdf_links": financial_pdf_rows,
        "corporate_actions_rows": corp_action_rows,
        "eod_history": history,
        "dictionary_comparison": reports,
    }
    out_path = os.path.join(out_dir, f"{symbol.upper()}_psx_snapshot.json")
    with open(out_path, "w") as f:
        json.dump(out_payload, f, indent=2, default=str)

    return out_payload


def run(symbols: list, history_days: int, out_dir: str, delay: float,
        watch_rows_by_symbol: dict = None, fetch_details: bool = True,
        extras: dict = None):
    """Scrape every symbol in `symbols`. Fetches KSE-100 once (market-wide)
    and reuses it for every company's market_data_daily row. One company
    failing does not stop the rest of the batch.

    If `watch_rows_by_symbol` is given, per-company quote pages are skipped
    entirely (fast path) and quotes are built from that table instead.
    """
    os.makedirs(out_dir, exist_ok=True)

    print("Fetching market-wide KSE-100 index (shared across all companies) ...")
    try:
        kse = scrape_kse100()
        print(f"  KSE-100 close: {kse.get('kse_100_close')}")
    except Exception as e:
        print(f"  [warn] could not fetch KSE-100: {e}", file=sys.stderr)
        kse = {}

    all_results = {}
    failed = []
    n = len(symbols)

    for i, symbol in enumerate(symbols):
        su = symbol.upper()
        quote = None
        if watch_rows_by_symbol is not None:
            quote = quote_from_watch_row(watch_rows_by_symbol[su])
        try:
            payload = run_one(su, kse, out_dir, history_days,
                               quote=quote, fetch_details=fetch_details, extras=extras)
            all_results[su] = payload
        except Exception as e:
            print(f"  [error] {su} failed: {e}", file=sys.stderr)
            failed.append({"symbol": su, "error": str(e)})

        if (i + 1) % 25 == 0 or i == n - 1:
            print(f"  progress: {i + 1}/{n}")

        # only need to be polite when we're actually hitting a per-company
        # page (fast/watch-table mode makes no extra requests per symbol)
        if fetch_details and i < n - 1 and delay > 0:
            time.sleep(delay)

    combined_path = os.path.join(out_dir, "ALL_COMPANIES_psx_snapshot.json")
    with open(combined_path, "w") as f:
        json.dump({
            "fetched_at": datetime.now().isoformat(timespec="seconds"),
            "kse100_raw": kse,
            "companies": all_results,
            "failed": failed,
        }, f, indent=2, default=str)

    print(f"\n============================================================")
    print(f"Done. {len(all_results)} succeeded, {len(failed)} failed.")
    if failed:
        print(f"Failed symbols: {[f['symbol'] for f in failed]}")
    print(f"Combined snapshot: {combined_path}")
    print(f"Per-company snapshots: {out_dir}/<SYMBOL>_psx_snapshot.json")


def load_symbols(args) -> list:
    """Build the final list of symbols from --symbols-file and/or positional
    args, de-duplicated and order-preserved."""
    symbols = []

    if args.symbols_file:
        with open(args.symbols_file) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    symbols.extend(s.strip() for s in re.split(r"[,\s]+", line) if s.strip())

    if args.symbols:
        for s in args.symbols:
            symbols.extend(x.strip() for x in s.split(",") if x.strip())

    seen = set()
    ordered = []
    for s in symbols:
        su = s.upper()
        if su not in seen:
            seen.add(su)
            ordered.append(su)

    return ordered


def main():
    parser = argparse.ArgumentParser(
        description="Real-time PSX scraper + data-dictionary mapper - one company, "
                     "a list of companies, or the whole exchange."
    )
    parser.add_argument(
        "symbols", nargs="*",
        help="PSX ticker symbols, space or comma separated, e.g. PSO APL WAFI HTL HASCOL "
             "(or PSO,APL,WAFI). Ignored if --all is given."
    )
    parser.add_argument(
        "--symbols-file", default=None,
        help="path to a text file with one symbol per line (comma/space separated lines also OK)"
    )
    parser.add_argument(
        "--all", action="store_true",
        help="scrape EVERY company currently listed on PSX (discovered live from "
             "dps.psx.com.pk/market-watch, ~500+ symbols). Fast by default (1 request "
             "covers price/volume/OHLCV for all companies); add --detailed for full "
             "per-company financials/market-cap (slow, one page load per company)."
    )
    parser.add_argument(
        "--detailed", action="store_true",
        help="also visit each company's individual page for market cap, EPS/ratios, "
             "and financial statements. Only meaningful with --all or a symbol list; "
             "for --all this means 500+ requests to PSX - use responsibly."
    )
    parser.add_argument("--limit", type=int, default=0,
                         help="cap the number of companies processed (0 = no cap). "
                              "Useful for testing --all/--detailed on a small batch first.")
    parser.add_argument("--history", type=int, default=0,
                         help="also fetch N most recent EOD history points per company "
                              "(only used with --detailed, 0 = skip)")
    parser.add_argument("--out", default="./psx_output", help="output directory for JSON snapshots")
    parser.add_argument("--delay", type=float, default=1.0,
                         help="seconds to wait between companies when --detailed is on "
                              "(politeness / rate-limit safety)")
    parser.add_argument("--yes", action="store_true",
                         help="skip the confirmation prompt for large --detailed runs")

    # --- new optional modules (see gap-analysis doc: Ratios, Payouts,
    #     Reports, Announcements, Financial-statement PDFs, Corporate
    #     actions were previously 0-60% covered) --------------------------
    parser.add_argument("--ratios", action="store_true",
                         help="scrape + normalize the company Ratios widget "
                              "(gross/operating/net margin, ROE, ROA, current/quick "
                              "ratio, debt ratio, PEG, price/book)")
    parser.add_argument("--payouts", action="store_true",
                         help="scrape dividend/bonus/right payout history "
                              "(also used to estimate dps_ttm)")
    parser.add_argument("--reports", action="store_true",
                         help="collect Annual/Quarterly Report PDF links")
    parser.add_argument("--announcements", action="store_true",
                         help="scrape PSX notices / corporate announcements")
    parser.add_argument("--financial-pdfs", action="store_true", dest="financial_pdfs",
                         help="collect full financial-statement PDF links "
                              "(balance sheet, P&L, cash flow, etc.)")
    parser.add_argument("--corporate-actions", action="store_true", dest="corporate_actions",
                         help="derive splits/bonus/rights/mergers from payouts + announcements")
    parser.add_argument("--extras", action="store_true",
                         help="shorthand for turning ALL of --ratios --payouts --reports "
                              "--announcements --financial-pdfs --corporate-actions on at once")

    args = parser.parse_args()

    extras = {
        "ratios": args.ratios or args.extras,
        "payouts": args.payouts or args.extras,
        "reports": args.reports or args.extras,
        "announcements": args.announcements or args.extras,
        "financial_pdfs": args.financial_pdfs or args.extras,
        "corporate_actions": args.corporate_actions or args.extras,
    }
    any_extras = any(extras.values())

    watch_rows_by_symbol = None

    if args.all:
        print("Discovering all listed PSX companies from market-watch ...")
        watch_rows = scrape_market_watch_all()
        if not watch_rows:
            print("Could not read the market-watch table (site layout may have changed, "
                  "or it's outside trading hours with no cached table). Aborting.", file=sys.stderr)
            sys.exit(1)
        watch_rows_by_symbol = {r["symbol"]: r for r in watch_rows}
        symbols = list(watch_rows_by_symbol.keys())
        print(f"  found {len(symbols)} listed symbols")
    else:
        symbols = load_symbols(args)
        if not symbols:
            symbols = ["PSO"]  # sensible default

    if args.limit and args.limit > 0:
        symbols = symbols[:args.limit]

    # explicit symbol lists always go detailed; extras also force detailed
    # mode since every extra module needs a per-company page load
    fetch_details = args.detailed or not args.all or any_extras

    if fetch_details and len(symbols) > 25 and not args.yes:
        extra_requests_per_symbol = 1  # financials/ratios page is shared
        if extras["payouts"] or extras["corporate_actions"]:
            extra_requests_per_symbol += 1
        if extras["reports"]:
            extra_requests_per_symbol += 1
        if extras["announcements"]:
            extra_requests_per_symbol += 1
        if extras["financial_pdfs"]:
            extra_requests_per_symbol += 1
        est_minutes = round(len(symbols) * extra_requests_per_symbol * (args.delay + 1) / 60, 1)
        resp = input(
            f"\nThis will make a detailed, per-company scrape of {len(symbols)} PSX "
            f"companies (~{est_minutes} min at current --delay, modules: "
            f"{[k for k, v in extras.items() if v] or 'core only'}). PSX's Terms of Use "
            f"restrict automated/bulk data collection - this should stay personal/"
            f"research use only. Continue? [y/N]: "
        ).strip().lower()
        if resp != "y":
            print("Aborted.")
            sys.exit(0)

    print(f"Symbols to scrape ({len(symbols)}): "
          f"{', '.join(symbols[:15])}{' ...' if len(symbols) > 15 else ''}")
    print(f"Mode: {'DETAILED (per-company pages)' if fetch_details else 'FAST (market-watch table only)'}")
    if any_extras:
        print(f"Extra modules on: {[k for k, v in extras.items() if v]}")

    # in fast mode we already have every quote from market-watch - no need to
    # re-fetch per symbol
    watch_for_run = watch_rows_by_symbol if not fetch_details else None

    run(symbols, args.history, args.out, args.delay,
        watch_rows_by_symbol=watch_for_run, fetch_details=fetch_details,
        extras=extras)


if __name__ == "__main__":
    main()