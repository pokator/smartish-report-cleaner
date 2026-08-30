# Weekly Report Reconciler — Write-up

**What I built:** A tool that replaces the manual weekly process of
hand-combining three CSV exports (Shopify orders, warehouse stock, returns)
into one Excel summary. It comes in two forms: a `build_report.py` CLI
(pandas/openpyxl) for anyone comfortable with Python, and a static,
client-side web page (plain HTML/JS, hosted on GitHub Pages) so anyone on
the team can drag in the three files and download a report with zero setup.

**How it works:** The three systems format the shared product SKU
differently (`SMR-1001` / `1001` / `SMR_1001`). Rather than hardcoding a
prefix, I extract the trailing digit run from whatever string each system
gives and use that as the join key — so it doesn't care about case, dashes,
or underscores. Column names are matched fuzzily (case/space/underscore-
insensitive) so a renamed header doesn't break the pipeline outright — it
warns and continues. Every judgment call (dropped exact-duplicate rows,
blank SKUs, blank stock counts treated as "Unknown" rather than 0, orphaned
return SKUs with no matching order) lands on a separate **Flags** sheet
instead of being silently discarded. The web version is a direct JS port of
the Python logic, verified against it with a small Node test harness before
I trusted it.

**AI tools used:** Built this end-to-end with Claude (Sonnet) in an agentic
coding session — it inspected the raw CSVs directly, wrote both
implementations, and ran its own tests rather than me eyeballing output by
hand.

**What broke, and the fix:** The first version silently turned a genuinely
*unknown* stock count into `0` — `pandas.groupby().sum()` returns `0` for an
all-blank group by default, not `NaN`. That's a meaningful bug for this data
(unknown stock vs. confirmed zero stock are very different facts), caught by
manually inspecting the output row-by-row against the source CSV rather than
trusting a clean run. Fixed with `sum(min_count=1)` and a small `on_hand_known`
flag carried alongside the total.

**Next steps for production:** Replace the fuzzy-column-name guessing with
a lightweight schema/contract per source system (fail loudly, not just
warn, if a required field truly disappears); move the reconciliation rules
into one shared library instead of two parallel implementations (Python +
JS); add a scheduled run against the actual source systems instead of
manual CSV drops; and get one more real edge case reviewed by whoever owns
each source system — particularly whether the SKU-digit-only join key could
ever collide across product lines.
