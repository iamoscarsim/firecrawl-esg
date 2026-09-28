# Firecrawl ESG — Scope 1/2 Emissions Evidence Collection

## Purpose

This is a data-collection workspace, not a code repo. The recurring task: given a batch
of companies and their websites, determine whether each company reports **Scope 1 and
Scope 2 emissions data** on its website.

- Data visible on a webpage → take a **full-page screenshot** of that page
- Data not on a webpage → **download the ESG / sustainability / annual report PDFs**

After collection, **dig into the evidence and fill the actual Scope 1/2 values into
results.csv** (see "Scope 1/2 value extraction" below). Deliverable PDFs stay untouched
originals — extraction reads them, never modifies them. When the whole batch is done,
build the formatted workbook (see "Excel report") — a run isn't finished without it.

## Run-based folder layout

Every batch of companies is one **self-contained run**. Runs are independent: never
reuse, compare against, or write into a previous run's cache or output. (Resuming an
unfinished run is continuing the *same* run, not cross-run reuse — see "Batch state &
resumption".)

```
Firecrawl ESG/
  CLAUDE.md
  skills.md                       # Firecrawl onboarding doc (contains an API key — do not share/commit)
  tools/
    find-ghg-pages.py             # locates Scope 1/2 pages inside PDFs (shared across runs)
    render-pdf-page.py            # renders PDF pages to PNG for visual value confirmation
    build-report-example.py       # reference openpyxl script for the Excel report
  runs/
    <YYYY-MM-DD>_run-<nn>/        # one folder per run/batch, e.g. 2026-07-11_run-01
      cache/                      # firecrawl map/scrape JSON + helper scripts — scratch, regenerable
      companies/
        <Company Name>/           # deliverables: screenshots and/or PDFs
      queue.csv                   # batch state machine — see "Batch state & resumption"
      results.csv                 # tracker for this run only
      Scope-1-2-Emissions-Report.xlsx  # formatted end-of-run deliverable — see "Excel report"
```

File naming inside `companies/<Company Name>/`:

- Screenshots: `<Company>-scope-1-2-emissions-page.png`
- PDFs: `<Company>-<ReportName>-<Year>.pdf` — always rename to something descriptive;
  never keep server hash filenames (e.g. `d8c41071e0efa…pdf`)

## results.csv

One row per company, written into the run folder as each company completes:

| column            | contents                                                                                   |
| ----------------- | ------------------------------------------------------------------------------------------ |
| `company`         | Company name                                                                                |
| `website`         | The site given in the batch input                                                           |
| `outcome`         | `screenshot` \| `pdf` \| `none`                                                             |
| `report_year`     | Fiscal year(s) the data/report covers, e.g. `FY2025` or `2022`                              |
| `scope1`          | Scope 1 emissions in **tCO2e** (latest year)                                                |
| `scope2_market`   | Scope 2 market-based emissions in **tCO2e**                                                 |
| `scope2_location` | Scope 2 location-based emissions in **tCO2e**                                               |
| `scope1_2_combined` | Combined Scope 1+2 in **tCO2e** — ONLY when the company doesn't report the split          |
| `note`            | Provenance & caveats: source page number(s), unit conversions ("converted from million tons"), inferred values ("scope 2 basis not stated — inferred"), or why cells are empty |
| `source_urls`     | Screenshot rows: the exact live page URL captured. PDF rows: URL(s) each PDF came from     |
| `files`           | Relative link(s) to the deliverables, e.g. `companies/Acme Corp/Acme-Corp-GHG-Report-2025.pdf` |

All emission values in **tCO2e** — convert if published otherwise and say so in `note`.
If a company reports a single undifferentiated Scope 2 figure, put it in the column
matching the stated basis; if the basis is not stated, use `scope2_location` and mark it
inferred in `note`. If figures genuinely don't exist, leave cells empty and explain in
`note` — never guess a number.

`scope1_2_combined` is for companies that publish only a merged Scope 1+2 total: fill
it, leave `scope1`/`scope2_*` empty, and say "reported combined only" (plus the Scope 2
basis inside the total, if stated) in `note`. If a company publishes both the split and
a combined total, record the split only — the combined figure is derivable and two
sources of truth invite mismatches.

CSV hygiene: always quote fields containing commas or quotes (company names, `note`
text, multi-URL cells) — `note` fields routinely contain both.

No collection-date column — the run folder name carries the date.

## Excel report (end-of-run deliverable)

Every run ends with `Scope-1-2-Emissions-Report.xlsx` in the run folder. results.csv
stays the raw machine tracker (schema above, untouched); the workbook is the
human-facing view. It is built **once, after the last results.csv row lands** — by the
orchestrator in parallel mode — via an openpyxl script under `cache/`
(`pip install openpyxl` if missing). Start from the reference implementation
`tools/build-report-example.py` (a past run's script, styling and hyperlink handling
already correct) and swap in the current run's data.

One sheet ("Scope 1-2 Results"), Arial, gridlines off, title row + subtitle (run name,
company count, "all emission values in tCO2e"), navy header row (white bold, `1F3864`),
frozen header, autofilter, zebra banding, thin grey borders.

Columns (note the deltas vs results.csv):

- `Company · Website · Outcome · Report year · Scope 1 (tCO2e) · Scope 2 market-based
  (tCO2e) · Scope 2 location-based (tCO2e) · Scope 1+2 combined (tCO2e) · Confidence ·
  Note · Source (click to open at page)`
- **Website** cells are live hyperlinks to the given site.
- The **combined** column is always present and mirrors results.csv's
  `scope1_2_combined` exactly: filled only for combined-only reporters, empty otherwise.
- **Units live in the headers** — value cells are bare numbers, format `#,##0`.
- **Report year**: plain 4-digit year (latest year the data covers), never `FY2025`;
  fiscal-year nuance goes in the note. Rows with no data: `—`.
- **Note**: ONE line — the key caveat only (basis inferred, coarse rounding, needs-review
  reason). Full provenance stays in results.csv.
- **Source** merges results.csv's `source_urls` + `files` into a single hyperlink:
  screenshot rows → the live page URL; PDF rows → the company's **hosted** PDF URL with
  `#page=N` appended (the evidence page from the note) so the browser opens the report
  at the right page — label `"<Report name> · p.N"`; no-data rows → the company site,
  muted. Use **native cell hyperlinks** (`cell.hyperlink = url`), NOT `=HYPERLINK()`
  formulas — no LibreOffice on this machine to bake formula caches, and native links
  preserve the fragment. Verify after saving: unzip the xlsx and grep
  `xl/worksheets/_rels/` for `#page=`. (If the user ever reports Excel stripping the
  fragment on click, switch to `=HYPERLINK()` formulas.)
- **Confidence** (colored text; rubric repeated in a legend below the table). ONE
  signal per row — a row is EITHER rated for data quality OR marked for review, never
  both. **Data age is NOT a factor** (freshness is already visible in Report year):
  - `High` (green `1E7B34`) — exact figures lifted straight from the company's own
    report/page; no interpretation needed
  - `Medium` (amber `B45F06`) — interpretation was applied to get the numbers: Scope 2
    basis inferred, unit conversion/coarse rounding, or values read off a chart rather
    than a table
  - `Needs review` (red `C00000`) — the row's queue status is `needs_review`: a human
    ruling is pending and this OVERRIDES any data-quality rating. Triggers: values
    sourced from a parent/child company's report (ALWAYS review — parent/childco is
    the ONLY acceptable outside-the-company source; the aggregator ban is unchanged),
    identity unconfirmed (rebrand, wrong-brand domain), bot-blocked/non-original
    evidence, or a doubtful no-disclosure. The exact reason lives in the queue.csv
    note. (There is no `Low` tier — entity doubt IS a review reason.)
  - `—` (grey) — nothing published to extract, and no ruling pending
- Row highlighting: NO row fills beyond zebra banding — review status is carried by
  the Confidence column, never by a yellow/amber fill. Rows with **no emission values**
  (all four value columns empty) get muted grey text — keyed off the values, NOT off
  `outcome=none`: a company whose collected PDF contains no Scope 1/2 figures is just
  as "nothing published" to the reader as one with no report at all.
- After delivery the workbook is the user's to annotate — reviewed/accepted rows are
  edited by the user in Excel; queue.csv stays the machine record of what was flagged.
- Below the table: a small legend (confidence rubric incl. what Needs review means and
  where its reason lives, grey text = no figures published, what Source links do,
  pointer to `companies\<Name>\` and results.csv).

## Batch state & resumption (queue.csv)

Large batches span many sessions and may pause at the Firecrawl credit wall, so every
run carries a queue file as its single source of progress truth.

- At run start, generate `queue.csv` from the batch input:
  `company,website,status,note,completed_date`
- Statuses: `pending` | `in_progress` | `done` | `needs_review` | `deferred_credits`
- Every session starts by reading `queue.csv` and picking the first `pending` company.
- Every company completion flips its queue row (with `completed_date`) **and** appends
  its `results.csv` row — the two updates happen together.
- `needs_review` means a human ruling or manual retry is needed; the queue `note` must
  say exactly why (e.g. "PDF gated behind form", "domain redirects to parent Acme Corp").
- "continue the run" from the user is a complete resumption instruction: read
  `queue.csv` and keep going. Never re-derive progress by inspecting folder contents.
- **Locating the run to resume**: scan `runs/*/queue.csv` for unfinished rows
  (`pending` / `in_progress` / `deferred_credits`). Exactly one run matches → resume
  it. Multiple match → ask the user which. The user may also name the run directly
  ("continue run-02"). All work resumes **inside that run's folder** — its `queue.csv`,
  its `results.csv`, its `cache/`.

## Credit policy

A billing cycle's credit allowance cannot always cover a large batch (e.g. 500 names ≈
2,000–4,000 credits vs ~3,000/cycle on the Hobby tier). Pause proactively instead of
failing mid-company:

- Run `firecrawl --status` at session start and after every ~10 companies.
- When fewer than ~50 credits remain: finish the in-flight company, mark all remaining
  `pending` rows `deferred_credits` with a note of when the cycle resets, tell the user,
  and stop.
- On resume after the cycle resets, `deferred_credits` rows go back to the front of the
  queue as `pending`.

## Parallel mode (multi-worker runs)

Default for batches over ~20 names. Companies are independent, so collection fans out
across parallel workers; the per-company workflow, decision rule, edge-case rulings,
and verification steps apply **verbatim inside each worker** — parallel mode changes
who does the work, never what the work is.

**Roles:**

- **Orchestrator** (the main session) does no company work. It assigns slices, spawns
  workers, collects their results, runs the credit checkpoints, and is the **single
  writer of `queue.csv` and `results.csv`** — workers never touch these two files. It
  also builds the Excel report once every row has landed (see "Excel report").
- **Workers** are background subagents. Each gets a **disjoint, pre-assigned slice** of
  ~8–12 companies and runs the full loop for each: map → filter → scrape → judge →
  download → verify → extract. Workers DO write deliverables into their companies'
  own `companies/<Name>/` folders and cache files under `cache/` (slugs keep filenames
  disjoint) — only the two shared CSVs are orchestrator-only.
- Each worker's **final message returns structured results only**: per company, the
  results.csv row values, status (`done`/`needs_review`), note text, and relative file
  paths. The orchestrator writes the rows and flips the queue.

**Worker count:** ~2 workers per Firecrawl concurrency slot (a worker occupies a slot
only ~⅓ of its time — the rest is judgment, curl, and local PDF work). Concurrency 2 →
3–4 workers; concurrency 5 → 8–10. Check the plan's cap via `firecrawl --status`.
Excess simultaneous requests queue server-side — the cap self-throttles; workers never
coordinate timing.

**Worker model:** spawn workers **one model tier below the session model** — session on
Fable → workers on Opus; session on Opus → workers on Sonnet; session on Sonnet →
workers stay on Sonnet (never Haiku: extraction needs high-res vision and the
market/location + unit-conversion judgment calls are where correctness risk lives).
Pass the model explicitly on each Agent spawn. Extraction subagents follow the same
rule. Effort: **high** for session and workers alike — the doc pre-decides the
judgment calls, so `xhigh` overthinks and `medium` risks sloppy value verification.

**Extraction in parallel mode** runs inline within each worker, per company — the
worker level already provides the parallelism, so the serial-mode chunked waves don't
apply.

**Credit checkpoints:** before spawning a wave of workers, verify remaining credits ≥
(companies in the wave × 8). Too few → shrink the wave or mark the remainder
`deferred_credits` per the credit policy.

**Failure isolation:** a stuck company gets `needs_review` per the rulings and the
worker moves on — one worker's bot wall never stalls the others. If a worker dies or
returns nothing for some of its slice, the orchestrator reverts those rows to `pending`
for reassignment; deliverables already on disk for a reverted company may be reused
after re-verification (this is same-run resumption, not cross-run reuse).

## The decision rule

1. **Scope 1 AND Scope 2 absolute figures visible on a webpage** → full-page screenshot
   of that page. Visually verify the screenshot actually shows the figures before
   counting it done.
2. **Otherwise** → download, unparsed, the **latest year only** of each of:
   - the **annual report**
   - the **ESG report** (often called "sustainability report")

   One of each — no prior years. If the company also publishes a dedicated GHG/emissions
   or TCFD report for that same latest year, it may be included as well since it is the
   most direct carrier of Scope 1/2 data.
3. **Source restriction**: only PDFs from the company's own site/subdomains
   (e.g. `newsroom.<company>.com`) or files directly linked from the company's own
   pages (company CDNs count — branded asset hosts like `assets.<company>.com` or
   `<company>.widen.net`/Bynder tenants, and generic CMS/CDN links such as
   `cdn.sanity.io` or `data.maglr.com` when found on their pages). Third-party aggregators
   (responsibilityreports.com, annualreports.com, etc.) do NOT count.
4. Reduction **targets** ("cut Scope 1+2 by 75% by 2030") are NOT reported data — only
   absolute figures (tables/charts with tCO2e values) count for the screenshot path.

## Edge-case rulings

Pre-decided calls so a run never stalls on judgment. When one applies, cite it in the
queue/results `note`.

- **Scope 1 and Scope 2 on different webpages** → screenshot both pages; both URLs in
  `source_urls`, both PNGs in `files`.
- **Combined-only figure** → `scope1_2_combined` column (rule in the results.csv section).
- **XLSX/ESG data workbook** → valid deliverable, treated like a PDF: download with
  curl, verify it isn't an HTML error page (XLSX starts with `PK` magic bytes), extract
  locally.
- **Gated (form-wall) or bot-blocked (403) PDF** → one retry: alternate user-agent, or
  scrape the hosting page for a direct CDN link. Still blocked → `needs_review` with
  the reason. Never fall back to aggregators.
- **Aggregator links (annualreports.com etc.), even when linked from the company's own
  page** → still excluded; the aggregator ban is absolute.
- **Website redirects to a parent company** (or the company otherwise turns out to
  report under a parent) → collect the parent's group-level report as evidence; leave
  value cells empty unless the report breaks out the subsidiary; note "reports under
  parent <X>; group-level report collected"; status `needs_review`. Collecting the
  parent's report is NOT a side-grab: run the **full report-finding workflow on the
  parent's domain** — live-scrape the parent's sustainability/reports hub and apply
  the latest-report and staleness checks there exactly as for an assigned company.
  This has failed before when done lazily: a worker fetched Liberty Global's 2024
  report from a remembered URL while the parent's live sustainability page linked the
  2025 edition. Before writing "no subsidiary breakout", grep the parent PDF's FULL
  text for the subsidiary's name and inspect every hit (page-by-page, not the first
  few) — breakouts hide in appendices (observed in practice: Virgin Media Ireland's
  KPMG-assured Scope 1/2 figures sat in the Sustainability-Linked Loans appendix on
  p.35 of a 90+ page Liberty Global report and were missed by a truncated grep).
- **Intensity-only reporter (tCO2e per revenue/unit, no absolute figures)** → collect
  evidence per the normal rule, leave value cells empty, note "intensity-only; no
  absolute figures published".
- **Promising page but markdown grep finds no figures** (charts/images) → take the
  full-page screenshot anyway and decide visually before falling to the PDF path.
- **On-page figures too coarsely rounded** (rounding step ≥10% of the value itself,
  e.g. a page publishing "0.1 Mt" or "~100k" where the report carries exact tonnes) →
  the page alone is not sufficient evidence. Take the screenshot, then ALSO run the
  PDF path (annual/ESG report, same latest year) and extract the precise values from
  the report. results.csv gets the precise PDF values; keep both deliverables and
  both URLs in `source_urls`; `note` says "on-page figures too rounded (<what the
  page showed>); precise values from <report> p.N". Confidence follows the PDF
  values.
- **Scanned PDF (`NONE` verdict, no text layer)** → render the table-of-contents page
  to locate the GHG/ESG-data section; the ~5-page render cap still applies overall;
  otherwise note why and move on.
- **Screenshot cut off or lazy-loaded content missing** → one re-scrape with
  scroll/wait actions; still incomplete → fall to the PDF path.
- **Non-English site** → the URL filter and content greps also try localized terms
  (most non-English reports still print "Scope 1/2" verbatim; add e.g. "Alcance",
  "Emissionen", "排出", "スコープ" when the site language warrants).
- **Dead/parked/wrong domain** → one `firecrawl search` for the official site; none
  found → `outcome=none`, status `needs_review`.
- **Rebrand / successor domain** (input domain is stale and the company now lives on a
  new brand's site, found via redirect or search) → treat the successor domain as the
  company's own site and rerun the FULL per-company workflow on it from step 1 (map the
  new domain, filter, scrape) — do NOT settle for the first report a search surfaces,
  and confirm the collected report is the **latest year available on the new site**.
  Note "input domain stale; rebranded to <X>"; status `needs_review` (identity needs
  human confirmation). This has failed before when done lazily: a search-and-grab
  fetched one report without mapping the successor domain, which would have missed
  newer reports — always map.
- **Site down/unreachable** (timeout, connection refused, 5xx) → one quick retry, then
  skip and move on: `outcome=none`, note "site unreachable at collection time", status
  `needs_review`. Don't burn worker time on outages — the batch keeps moving.
- **"Nothing published" requires proof the site was ALIVE.** A successful
  `firecrawl map` is NOT that proof — map serves Firecrawl's *index* and can return
  thousands of URLs for a site that is down (observed in practice: ~2,900 mapped
  URLs while the site itself returned 504). Before writing a "no ESG/GHG disclosure"
  note, the company must have at least one successful live scrape in this run, or pass
  `curl -sSI --max-time 20 <homepage>` with a 2xx/3xx. Site dead → the
  site-down/unreachable ruling above (`needs_review`, "site unreachable"), never
  "no disclosure" — a downed site says nothing about what the company publishes.
- **Map output can't prove a report is the LATEST either.** Never download a report
  PDF whose URL was seen only in map output — the same stale index has missed a newer
  report before (observed in practice: map's newest for a site was the 2023 report
  while the live sustainability hub page linked the 2025 one). A report's download
  link must come from a **live-scraped company page** (the sustainability/reports hub)
  or a `firecrawl search` result. Use the map only to decide which pages to scrape.
  When a scraped page has no report links, the fallback is scraping the hub/archive
  pages found in the map — never grabbing a PDF URL straight from the map list.
- **A search hit that points DIRECTLY at a PDF can't prove latest either.** Search
  indexes lag just like the map: before downloading a PDF whose URL came straight from
  a search result, confirm it is the latest by live-scraping the hub/news page that
  hosts it (usually sitting in the same search results). Observed in practice: search
  surfaced VodafoneZiggo's 2024 annual-report PDF while the company's own live news
  page — the first result — linked the 2025 edition. Search hits that point at *pages*
  are fine: scrape the page and take the report link from there.
- **Report URLs from model memory are BANNED.** Every downloaded file's URL must be
  traceable to a cache artifact from this run (a scrape's `links`/markdown or a search
  output). If the agent "already knows" where a company hosts its report, it must
  still re-derive the URL live first — model knowledge is a stale index with no
  timestamp and no provenance trail. Observed in practice: a remembered Liberty Global
  report URL fetched the superseded 2024 edition without any live check ever touching
  the parent's site.
- **Latest report found is 2+ years old** (data year ≤ current year − 2 for a company
  that is clearly still operating) → treat as suspicious before finalizing: scrape the
  sustainability/reports hub page live and run one
  `firecrawl search "<company> sustainability report <current year> OR <current year − 1>"`.
  Newer report found → collect it per the normal rule. Genuinely nothing newer →
  proceed, and add "confirmed latest available (published <year>)" to `note`. Costs
  1–2 credits and only fires on old-looking results.
- **ABSENCE claims carry the same burden of proof as extracted values.** Positive
  values must be visually confirmed; negative conclusions ("no figures", "no
  subsidiary breakout", "no disclosure") must meet an equivalent standard:
  - The deciding search must be **exhaustive** — full text of the document, every hit
    inspected. Never pipe the deciding grep through `head -N`; truncation is for
    display, not for deciding (observed in practice: a breakout sat at hit ~11 of a
    grep read only to hit 10).
  - Never **inherit** a negative across documents: a conclusion about one report
    edition says nothing about another edition, and a worker's finding on last year's
    report does not transfer to this year's.
  - A **crashed or partially-run tool is not a completed scan**: if find-ghg-pages.py
    (or any scan) exits with a traceback, the pages after the crash point were never
    scanned — fix/rerun before concluding anything (observed in practice: a cp1252
    crash at p.17 of a 37-page report hid assured subsidiary figures on p.35).
  - The `note` for any absence claim must say what was searched and how (e.g.
    "full-text grep for 'Ireland'/'Virgin Media', all 14 hits inspected").

## Workflow per company

1. `firecrawl map "<site>" --limit 3000 -o cache/map-<slug>.json --json`
   — cheap URL discovery, ~1 credit. `--limit 3000` caps the returned URL list.
2. Filter the URL list locally (free) for
   `sustain|esg|environment|responsib|climate|emission|carbon|report|\.pdf`.
   Zero hits → before heading toward a "no disclosure" conclusion, prove the site is
   actually alive (see the "Nothing published requires proof" ruling — map output
   alone is not liveness).
3. Scrape the 1–3 most promising pages:
   `firecrawl scrape "<url>" --format markdown,links -o cache/scrape-<slug>.json --json`
   then grep the markdown for `scope\s*[12]` **with surrounding context** to distinguish
   real figures from targets/narrative.
4. Figures on page → `firecrawl scrape "<url>" --full-page-screenshot -o cache/shot-<slug>.json --json`,
   then `curl` the `.data.screenshot` URL to the company folder as PNG.
5. No figures on page → collect PDF links from the scrapes' `links` arrays. If none:
   scrape the news/report-archive pages found in the map. Last resort:
   `firecrawl search "<company> ESG sustainability report site:<domain>"` — this is how
   subdomain-hosted reports get found. Report URLs must come from these live scrapes
   or search results — NEVER download a report PDF straight off the map's URL list,
   from a remembered/model-knowledge URL, or from a direct-PDF search hit without
   confirming it on the live hosting page (see the "map can't prove latest",
   "search hit directly at a PDF", and "model memory" rulings).
6. Download PDFs with **curl, not Firecrawl** (zero credits):
   `curl -sSL --fail --retry 2 -C - -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" -o <dest> "<url>"`
7. **Verify every PDF** (mandatory — a truncated download has already happened once).
   Both checks are near-instant (they read only a few bytes, not the whole file):
   - starts with `%PDF-` magic bytes → `head -c 5 <file>` (catches HTML error pages saved as .pdf)
   - ends with a `%%EOF` trailer within the last 1 KB → `tail -c 1024 <file>` (catches truncated downloads)

   If either check fails, re-download with `-C -` (resume) and re-verify.
8. Visually verify screenshots (Read the PNG) show the Scope 1/2 figures.
9. Extract the Scope 1/2 values (see next section — extraction runs in chunked
   parallel waves, not necessarily inline). A company's `results.csv` row and its
   `queue.csv` flip happen together, once its extraction completes.

## Scope 1/2 value extraction

Every company's evidence is mined for the actual numbers. In **parallel mode** (the
default for large batches — see "Parallel mode"), extraction runs inline within each
worker, per company. In **serial mode** (small batches, single agent), the run is a
**chunked hybrid**: collect serially in chunks of ~10–15 companies, then fan out one
extraction subagent per company in that chunk, in parallel, while collection continues
on the next chunk. Extraction is local and free, so it trails just behind collection
at no cost. Either way, each subagent/worker gets the company's files and returns only
the structured values, keeping the main session lean — and each company's `results.csv`
row lands shortly after its collection, preserving resumability. Do NOT defer all
extraction to the end of the run: a run interrupted mid-way must leave completed rows,
not folders of unmined PDFs.

**For PDF companies:**

1. **Locate** — `python tools/find-ghg-pages.py <pdf>` — wraps `pdftotext -layout`
   (local, free, ~2 s per report; zero tokens except the snippet output). Prints the
   candidate pages plus a verdict: `DATA` / `KEYWORDS_ONLY` / `NONE`.
2. **Extract** — `pdftotext -f <page> -l <page> -layout <pdf> -` on the candidate pages
   only; pull the latest-year Scope 1, Scope 2 market-based, and Scope 2 location-based
   values.
3. **Visual fallback** — if the verdict is `KEYWORDS_ONLY`/`NONE`, or the text values
   are garbled/ambiguous: render the specific pages to PNG with
   `python tools/render-pdf-page.py <pdf> <page|first-last> [out_dir]` and Read the PNGs
   (direct PDF page-reading fails on this machine — pdftoppm missing; PyMuPDF installed).
   Never read the whole report — cap at ~5 pages; if still not found, record why in
   `note` and move on.
   ⚠ `pdftotext -layout` is known to silently shift rows in multi-column tables
   (repeatedly observed in practice: Scope 1's values landing on the Scope 2 line).
   ALWAYS visually confirm the final values against the rendered page before writing
   them to results.csv — text extraction locates, the rendered page decides.

**For screenshot companies:** pull the values from the already-scraped page markdown in
`cache/` (fall back to reading the screenshot PNG). Same columns, `source_urls` already
carries the page URL.

**Rules (both paths):**

- Convert everything to **tCO2e**; note the conversion in `note` (e.g. "converted from
  million tons CO2eq p.4")
- Record the source page number(s) in `note` for PDF values
- Capture both Scope 2 bases when published; inferred basis → say so in `note`
- Cross-check each value once against its surrounding context (right year? right unit?
  market vs location?) before writing it
- No figures found → empty cells + explanatory `note`; never guess

## Tooling notes / gotchas

- Firecrawl CLI (`firecrawl-cli`) is installed globally and authenticated via stored
  credentials. Run `firecrawl --status` first. Plan: **Hobby tier — concurrency 5**
  (~3,000 credits/cycle; confirm the live balance via `--status`, don't assume).
- Be credit-frugal: map once per site, scrape only shortlisted pages, and reuse this
  run's own `cache/` instead of re-fetching. PDF/screenshot-image downloads via curl
  cost no credits.
- `jq` is NOT available on this machine — use `node -e` for JSON work. Anything beyond
  a one-liner goes in a script file under `cache/` (e.g. `cache/ctx.js`); inline
  `node -e` code gets mangled by shell quoting.
- Firecrawl JSON shapes: map → `.data.links[].url`; scrape → `.data.markdown`,
  `.data.links`, `.data.screenshot` (some outputs are top-level, so use
  `const d = j.data || j`).
- Slow servers (10 MB+ PDFs) outlive the 2–5 min foreground Bash timeout — run long
  downloads in the background and use `-C -` so retries resume instead of restarting.
- Pipeline where possible: background curl downloads and local PDF work (pdftotext,
  page rendering) for company N can overlap with mapping/scraping company N+1 — only
  Firecrawl API calls are bound by the plan's concurrency cap; curl and local tools
  are not.
