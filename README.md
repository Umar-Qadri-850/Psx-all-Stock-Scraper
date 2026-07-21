# PSX Real-Time Scraper + Data Dictionary Mapper

A single-file Python tool that scrapes public data from the Pakistan Stock
Exchange's data portal (`dps.psx.com.pk`), normalizes it into rows matching
a predefined DB schema, and reports — column by column — what it could and
couldn't populate, against an embedded Data Dictionary.

It can run against one company, a custom list, or every company currently
listed on PSX.

---

## What it scrapes

| Module | Flag | What you get |
|---|---|---|
| Live Quote | *(always on)* | Current price, open/high/low, volume, LDCP, change %, circuit breaker range, 52-week range |
| Equity Profile | *(always on, `--detailed`)* | Market cap, shares outstanding, free float shares/% |
| Market Watch (all companies) | `--all` | One request → every listed symbol's OHLCV, LDCP, sector, change |
| KSE-100 Index | *(always on)* | Index close/high/low/change, shared across all companies in a run |
| EOD Price History | `--history N` | Last N daily closes + volume from PSX's own JSON timeseries endpoint |
| Intraday Ticks | *(available as a function, not yet wired to CLI)* | Tick-level price/volume |
| Financial Summary | `--detailed` | Annual + Quarterly Sales / PAT / EPS tables |
| **Ratios** *(new)* | `--ratios` | Normalized: gross/operating/net margin, ROE, ROA, current/quick ratio, debt ratio, PEG, price/book |
| **Payouts** *(new)* | `--payouts` | Dividend / bonus / right issue history, book closure dates, period — also back-fills trailing DPS |
| **Reports** *(new)* | `--reports` | Annual/Quarterly Report PDF links |
| **Announcements** *(new)* | `--announcements` | PSX notices / corporate announcements (date, subject, category) |
| **Financial Statement PDFs** *(new)* | `--financial-pdfs` | Balance sheet, P&L, cash flow statement PDF links |
| **Corporate Actions** *(new)* | `--corporate-actions` | Splits / bonus / rights / mergers, derived from payouts + announcements |

Turning on all six new modules at once:

```bash
python main.py PSO --extras
```

---

## Coverage vs. the Data Dictionary

The tool ships with an embedded copy of your Data Dictionary (`stock_price_daily`,
`market_data_daily`, `stock_financials_quarterly`, plus the two new tables
`stock_ratios` and `stock_payouts`, and the reference table `corporate_actions`).
After every scrape it prints/saves a report of which columns were filled in
and which were not — and *why* (per your dictionary's own "Access Method"
notes), e.g.:

```
=== stock_price_daily — Data Dictionary Comparison ===
  Matched (5):
    ✔ unadjusted_price     = 245.3               [PKR/share]
    ✔ mkt_cap_mn           = 18452.1             [PKR mn]
    ...
  Missing / not available from live scrape (5):
    ✘ adjusted_open        -> Downloadable authenticated web interface (NOT public scrape)
    ✘ eps_ttm_cons         -> UNKNOWN - no confirmed source (flag, do not invent)
    ...
```

This makes it easy to see, at a glance, which fields genuinely cannot be
obtained from public pages (and need an authenticated download, a licensed
data vendor, or a separate workbook) versus which ones the scraper just
missed and can be fixed.

### Rough coverage estimate (public site only)

| Area | Status |
|---|---|
| Live Quote / Market Watch / EOD History | ~95% |
| KSE-100 Index | ~80% (index-level P/E, dividend yield, PKRV live in an internal workbook, not the public site) |
| Financial Summary (Sales/PAT/EPS) | ~90% |
| Ratios | now normalized (was "basic only") |
| Payouts / Corporate Actions | now scraped (was 0%) |
| Reports / Financial Statement PDFs | now link-collected (was 0%) |
| Announcements | now scraped (was 0%) |
| Full financial statement *line items* (Assets, Liabilities, Cash Flow detail) | still 0% — these live inside the PDFs the tool now links to, but are not parsed into structured rows |
| Consensus EPS | still unresolved — no confirmed public source |
| Adjusted OHLC | still unavailable — PSX only exposes this via an authenticated downloadable interface |

---

## Installation

```bash
pip install requests beautifulsoup4 pandas lxml
```

(`pandas` is optional — the script degrades gracefully without it, but the
financials/ratios/payouts/announcements table-parsing all rely on
`pandas.read_html`, so install it if you want those modules.)

---

## Usage

```bash
# One company, full detail (quote + financials)
python main.py PSO

# A specific list
python main.py PSO APL WAFI HTL HASCOL

# A list from a file (one symbol per line, or comma/space separated)
python main.py --symbols-file list.txt

# Every listed company — fast path, 1 request, price/volume only
python main.py --all

# Every listed company — full detail (500+ requests, slow, prompts to confirm)
python main.py --all --detailed

# Test on a small batch first
python main.py --all --limit 20

# Also pull last 30 days of EOD history per company
python main.py PSO --history 30

# New modules
python main.py PSO --ratios
python main.py PSO --payouts
python main.py PSO --reports
python main.py PSO --announcements
python main.py PSO --financial-pdfs
python main.py PSO --corporate-actions
python main.py PSO --extras          # all six at once
```

Useful flags:

| Flag | Purpose |
|---|---|
| `--out DIR` | output directory for JSON snapshots (default `./psx_output`) |
| `--delay SEC` | pause between per-company requests (politeness/rate-limit safety, default 1.0s) |
| `--limit N` | cap number of companies processed |
| `--yes` | skip the confirmation prompt on large detailed runs |

---

## Output

For every symbol, a JSON snapshot is written to
`<out>/<SYMBOL>_psx_snapshot.json` containing:

- `quote_raw`, `kse100_raw` — raw scraped values
- `stock_price_daily_row`, `market_data_daily_row` — rows mapped to your schema
- `stock_financials_quarterly_rows` — flattened Annual/Quarterly financial line items
- `stock_ratios_row` — normalized ratios (if `--ratios`)
- `stock_payouts_rows` — payout history (if `--payouts`)
- `reports_pdf_links`, `financial_statement_pdf_links` — PDF link lists (if `--reports` / `--financial-pdfs`)
- `announcements_rows` — announcement/notice list (if `--announcements`)
- `corporate_actions_rows` — derived splits/bonus/rights/mergers (if `--corporate-actions`)
- `eod_history` — daily price/volume history (if `--history N`)
- `dictionary_comparison` — matched vs. missing columns, per table

A combined `ALL_COMPANIES_psx_snapshot.json` is also written for batch runs,
with a `failed` list of any symbols that errored out (one company failing
never stops the rest of the batch).

---

## Design notes / how the new modules work

- **Ratios** — reuses the same `Ratios` table already located by the
  financials scraper (no extra request) and maps PSX's row labels
  (`Gross Margin`, `ROE`, `P/B Ratio`, ...) onto normalized snake_case keys
  via an alias list (`RATIO_LABEL_MAP`), so it survives minor label wording
  changes.
- **Payouts** — looks for an HTML table with a `Book Closure` column (or
  `Dividend` + `Bonus`/`Right` columns) on the company page, and extracts
  both dates out of the "01-Feb-2026 to 05-Feb-2026" style cell PSX uses.
- **Reports / Financial Statement PDFs** — both walk every `<a href="...pdf">`
  on the company page and keep only the ones whose link text/URL contains
  report-like or statement-like keywords, so they can run off the same page
  fetch without hardcoding PSX's exact tab structure.
- **Announcements** — looks for a `Subject`/`Announcement` table first,
  falling back to any PDF links that look like notices if no table is found.
- **Corporate Actions** — deliberately *not* a separate PSX page. Bonus/right
  entries are pulled straight from the payouts table (no duplicate request),
  and splits/mergers are keyword-matched out of whatever announcements were
  found (`CORPORATE_ACTION_KEYWORDS`).
- All six new scrapers use `get_html_soft()` (a non-raising variant of the
  original `get_html()`) so a renamed/missing widget on one company's page
  degrades to an empty list/dict instead of crashing the whole batch run.

### A known limitation

This sandbox's outbound network allowlist does **not** include
`dps.psx.com.pk`, so the new modules were validated with unit tests against
mocked HTML (table structures matching PSX's documented layout) rather than
a live fetch. The original quote/market-watch/timeseries code was already
working and untouched. If a new module comes back empty against the real
site, PSX's markup for that widget differs from what's assumed — open the
page's HTML and adjust the label/keyword lists near the top of the relevant
function (`RATIO_LABEL_MAP`, `CORPORATE_ACTION_KEYWORDS`, or the
`keyword_filter` lists passed into `_extract_pdf_links`).

### Still not covered (per your gap analysis, unchanged)

- Full financial statement **line items** (Assets, Liabilities, Equity,
  Revenue, Expenses, Operating/Financing/Investing Cash Flow) — these live
  inside the PDFs the tool now *links* to, but parsing PDF financial
  statements into structured rows is a separate project (would need a PDF
  table extractor, e.g. `camelot` or `pdfplumber`).
- Consensus EPS — no confirmed public source; the dictionary flags this as
  unknown rather than guessing.
- Adjusted OHLC (open/high/low/close) — PSX only exposes this via an
  authenticated downloadable web interface, not any public page.
- Market-wide P/E ratio, dividend yield, PKRV 12M yield, total market cap —
  these live in the internal "KSE100 Market Data" (Zakheera) workbook, not
  on any public PSX page.
- Playwright/browser automation for genuinely JS-rendered widgets — the
  current implementation is 100% `requests` + `pandas.read_html` +
  `BeautifulSoup`, which covers everything PSX server-renders. If a future
  PSX redesign moves a widget to client-side JS rendering, that widget would
  need a headless-browser fetch instead.

---

## Legal / Terms of Use — read before running at scale

`dps.psx.com.pk` carries an explicit Terms of Use / Legal Notice that:

- Prohibits bots/spiders/scrapers "except where you have prior written permission"
- Prohibits "systematic retrieval to create collections, compilations, databases or directories"
- Prohibits commercial dissemination of market data without a license from PSX (contact `marketdatarequest@psx.com.pk`)

This tool is intended for **personal / educational / research use only**.
Don't run it at high frequency, don't resell or redistribute the data, and
don't use it as the data source for a commercial product. For
production-grade or licensed PSX data, contact PSX's Market Data team or use
a licensed vendor. The `--all --detailed` (and any `--extras` combined with
`--all`) paths issue 500+ to 2,000+ requests and will prompt for
confirmation before running unless `--yes` is passed — use `--delay` to stay
polite to PSX's servers.

---

## File

Everything lives in a single file: **`main.py`**. No other project files are
required beyond the pip dependencies above.
