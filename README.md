# Weekly Report Reconciler

Turns three messy per-system CSV exports (Shopify orders, warehouse stock,
returns) into one clean, deduplicated product summary — without hand-matching
SKUs across systems every week.

**Live tool:** enable GitHub Pages on this repo (see below) and it's a
drag-and-drop web page — no install, no server, runs entirely client-side.

## Two ways to run this

### 1. In the browser (for end users)
Open `index.html` (locally, or via GitHub Pages once enabled — see below).
Drop in this week's three CSVs, click **Generate report**, download
`weekly_report.xlsx`. Files never leave the browser; there's no backend.

### 2. From the command line (for anyone comfortable with Python)
```bash
python build_report.py orders.csv warehouse.csv returns.csv -o weekly_report.xlsx
```
Requires `pandas` and `openpyxl`. Produces a styled workbook with the same
Summary/Flags sheets, plus currency formatting and column widths — a nicer
finish than the browser version can produce with a dependency-free JS library.

**The two implementations share the same rules** (SKU normalization, column
matching, dedup, flagging) but are separate code paths — `app.js` for the
browser, `build_report.py` for the CLI. If you change a rule in one,
change it in the other. See `WRITEUP.md` for why they're not unified.

## How the reconciliation works

- **SKU matching**: each system formats the product identifier differently
  (`SMR-1001`, `1001`, `SMR_1001`). The join key is the trailing digit run,
  zero-padded — extracted with a regex, not a fixed prefix list.
- **Deduplication**: exact-duplicate rows (same SKU + same values) are
  dropped. Two different orders for the same product are *not* treated as
  duplicates — only truly identical rows are.
- **Missing data**: blanks are never silently coerced to zero. A blank
  stock count shows as `Unknown`, not `0` — those mean different things to
  whoever reads the report.
- **Column matching is fuzzy**: `Order_ID`, `order id`, and `orderid` all
  match the same field, so a renamed header next week doesn't break the
  script — see `ORDERS_COLUMNS` / `WAREHOUSE_COLUMNS` / `RETURNS_COLUMNS`
  at the top of `app.js` and `build_report.py` to add new aliases.
- Everything that needed a judgment call — a dropped duplicate, a missing
  SKU, an orphaned return — lands on the **Flags** sheet instead of being
  silently discarded.

## Enabling GitHub Pages

1. Push this repo to GitHub.
2. Repo **Settings → Pages → Source**: select `Deploy from a branch`,
   branch `main`, folder `/ (root)`. Save.
3. Wait ~1 minute; GitHub gives you a URL like
   `https://<username>.github.io/<repo-name>/`.

No build step — it's static HTML/CSS/JS, loaded straight from the repo.

## Repo layout

```
index.html         the web UI
style.css           its styling
app.js              reconciliation logic (browser + Node-testable)
ui.js               DOM wiring: file drop zones, results table, XLSX export
build_report.py     CLI/Python version (pandas + openpyxl), styled output
sample_data/        the three example CSVs used to build and test this
test/run_test.js    Node harness — runs app.js against sample_data
WRITEUP.md          project summary
```

## Testing

```bash
npm install        # installs papaparse, used only by the test harness
node test/run_test.js
```
Prints the reconciled summary and flags for `sample_data/` so they can be
eyeballed against `build_report.py`'s output on the same files.

## Known limitation

The join key is only the trailing digit run of the SKU. If two genuinely
different products ever shared the same digits under different prefixes,
they'd incorrectly merge into one row. Not a risk with the current
single-prefix (`SMR-`) SKU scheme, but worth knowing if that changes.
