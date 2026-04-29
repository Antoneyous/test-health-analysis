# Test Health Analysis

A single-page web app that turns a Codility assessment export into a branded
test-health report — engagement, scoring, time pressure, integrity flags,
similarity check — and a PowerPoint deck ready to share with a customer.

The app is intentionally a single static HTML file with three vendored JS
libraries. No build step, no server, no framework. Open `index.html` in a
browser and it runs.

This README is written for the engineer who will eventually port this into the
Codility platform so that data is pulled directly instead of uploaded by hand.
The current shape — XLS upload → in-memory parse → render — is designed to map
cleanly onto a future API-fed version.

---

## Repository layout

```
test-health-analysis/
├── index.html                 # The entire app: HTML, CSS, JS in one file
├── README.md                  # This file
├── .gitignore
└── assets/
    ├── codility-logo-dark.png   # Used on light backgrounds (dashboard, content slides)
    ├── codility-logo-light.png  # Used on dark backgrounds (cover slide, dividers)
    ├── logos.js                 # Same two PNGs as base64 — needed by PPTX export
    └── vendor/
        ├── xlsx.full.min.js     # SheetJS — parses .xls / .xlsx / .csv
        ├── pptxgen.bundle.js    # PptxGenJS — generates the .pptx
        └── chart.umd.min.js     # Chart.js — score distribution chart on the dashboard
```

`assets/logos.js` is generated from the two PNGs so the PPTX export works under
all three serving modes — `http://`, `file://`, and offline. PptxGenJS needs
data URIs, and `fetch()` of a sibling file is blocked under `file://`. Putting
the data URIs in a separate `<script>` sidesteps the issue while keeping
`index.html` itself logo-free.

That's the whole repo. Everything in `assets/vendor/` is third-party; everything
else is project code.

---

## Running it

Open `index.html` in any modern browser. That's it.

For local development with hot-reload-on-save you can run any static server:

```bash
python3 -m http.server 8000
# then open http://localhost:8000/
```

The app works on `file://` too, but a real HTTP server is closer to how it
will be served in production.

---

## What the app does

The user lands on an upload screen, picks an exported assessment file plus a
small set of configuration values (passing score, max time, etc.), and clicks
**Generate Report**. The app then:

1. Parses the XLS into rows of candidate attempt data.
2. Computes aggregate stats across All Time, by Year, and by Month.
3. Renders a single-page dashboard with KPI tiles, a score distribution
   chart, and a three-tab Integrity panel (Integrity Risk, Similarity Check,
   Behavioural Signals).
4. On click, generates a 16-slide PowerPoint deck and downloads it.

---

## Data flow

```
┌─────────────────┐     ┌────────────────┐     ┌─────────────────┐     ┌──────────────┐
│  XLS upload     │ ──▶ │  parseRows()   │ ──▶ │  computeStats() │ ──▶ │  renderHTML  │
│  + config form  │     │  (SheetJS)     │     │  (in-memory)    │     │  +  Chart.js │
└─────────────────┘     └────────────────┘     └─────────────────┘     └──────────────┘
                                                        │
                                                        ▼
                                                ┌──────────────────┐
                                                │  downloadPPTX()  │
                                                │   (PptxGenJS)    │
                                                └──────────────────┘
```

State lives in three module-level globals, all in `index.html`'s `<script>`:

| Global            | What it holds                                                       |
| ----------------- | ------------------------------------------------------------------- |
| `_allRows`        | Every parsed row from the XLS (one row per candidate attempt)        |
| `_lastCfg`        | The configuration object the user submitted on the upload screen     |
| `_integrityRows`  | Rows surfaced for the Integrity Risk tab (raw, not aggregated)       |
| `_activeYears`    | Set of years currently selected in the Year filter                   |
| `_activeMonths`   | Set of `YYYY-MM` strings currently selected in the Month filter      |
| `LOGO_DATA_URIS`  | `{dark, light}` data URIs, populated lazily before PPTX export       |

When the user changes a filter, the dashboard re-runs `computeStats()` over the
filtered subset of `_allRows` and re-renders. Nothing is persisted to disk or
to localStorage.

---

## Expected XLS shape

The app reads the first sheet and looks for these columns by header name. Order
doesn't matter; extra columns are ignored.

| Column                   | Type     | Used for                                                         |
| ------------------------ | -------- | ---------------------------------------------------------------- |
| `Test`                   | string   | The assessment's display name (shown in headers and PPTX title)  |
| `Create date`            | datetime | Year / month bucketing                                           |
| `Start date`             | datetime | Whether a candidate actually started the test (drives drop-off)  |
| `% total score`          | number   | Median score, distribution histogram, score buckets              |
| `Time Used (minutes)`    | number   | Median time, time pressure analysis                              |
| `Integrity Risk Status`  | string   | Integrity tab; counted as "high-risk" when value contains "high" |
| `Similarity Check`       | string   | Similarity Check tab                                             |

A row with no `Start date` is treated as **invited but not attempted** —
that's how drop-off is computed.

---

## Configuration object (`_lastCfg`)

The upload screen collects these values; everything downstream reads from
`_lastCfg`. When the platform port happens, this is the object you'll
populate from the test's existing platform configuration instead of from a
form.

| Key             | Type                | Source UI control     | Used by                                                  |
| --------------- | ------------------- | --------------------- | -------------------------------------------------------- |
| `creator`       | string              | `#cfg-creator`        | Subheader on dashboard + PPTX                            |
| `passingScore`  | number \| null      | `#cfg-passing`        | Passing-rate calculation (rows with `% total score >= passingScore` pass) |
| `maxTime`       | number \| null      | `#cfg-maxtime`        | Time-pressure analysis ("how much of the limit was used?") |
| `fairness`      | number \| null      | `#cfg-fairness`       | Fairness threshold for the Score Distribution tab        |
| `fairnessRel`   | `'above' \| 'below'`| `#fair-above/below`   | Whether the fairness threshold is a floor or a ceiling   |
| `rotation`      | boolean             | `#cfg-rotation`       | Surfaces "task rotation enabled" badge; also auto-detected from task slot uniqueness |
| `proctoring`    | boolean             | `#cfg-proctoring`     | Integrity tab framing                                    |
| `leaked`        | boolean             | `#cfg-leaked`         | Reserved — surfaces a "tasks may be leaked" warning       |
| `weighted`      | boolean             | `#cfg-weighted`       | If true, applies per-task-slot weights to the score calc |
| `slotWeights`   | `{[n: number]: number}` | `#cfg-weight-slot-N` | Weighted score formula                                   |

---

## Calculations

All metric definitions live alongside their use sites in `index.html`. Below is
a single source of truth for what each metric means.

### Engagement

```
invitations  = _allRows.length
attempts     = rows where Start date is present
non-zero     = rows where Start date is present AND % total score > 0
drop-off %   = round((1 − attempts / invitations) × 100)
non-zero drop-off % = round((1 − non-zero / attempts) × 100)
```

### Scoring

```
medScore     = median of % total score across all rows that started
passingRate  = pct of attempts where % total score >= cfg.passingScore
                (null when no passingScore is configured)
scoreSD      = standard deviation of % total score, rounded to whole percent
```

The 11-bucket histogram bins scores as `0%`, `1–10%`, `11–20%`, … `91–100%`.

### Time

```
medTime      = median of Time Used (minutes) across all rows that started
medTimeLeft  = cfg.maxTime − medTime
```

The Key Takeaway / Findings / Next Steps copy uses these thresholds:

- `medTimeLeft <= 5` minutes → time limit is "too tight"
- `medTimeLeft >= 10` minutes → time limit is "too loose"

### Integrity

```
highIntegrityPct = pct of attempts where Integrity Risk Status contains "high"
```

The Integrity tab also lists each individual flagged candidate.

### Difficulty label (chip on dashboard)

| Median score | Label           |
| ------------ | --------------- |
| `>= 80`      | High Difficulty |
| `60–79`      | Medium-High     |
| `40–59`      | Medium          |
| `20–39`      | Medium-Low      |
| `< 20`       | Low Difficulty  |

---

## Generating the PPTX

`downloadPPTX()` builds a 16-slide deck with PptxGenJS:

| #  | Slide                | What's on it                                              |
| -- | -------------------- | --------------------------------------------------------- |
| 1  | Cover                | Test name, generation date, dark ultramarine background   |
| 2  | Executive Summary    | 5 KPI tiles + up to 3 Key Findings                        |
| 3  | Table of Contents    | Section 01 / 02 / 03                                      |
| 4  | Section 01 divider   | "ALL TIME"                                                |
| 5  | All Time stats       | 3 stat cards, score distribution, **Key Takeaway** band   |
| 6  | Section 02 divider   | "YEAR BY YEAR"                                            |
| 7… | One slide per year   | Same layout as the All Time slide, no Key Takeaway        |
| …  | Section 03 divider   | "MONTH BY MONTH"                                          |
| …  | One slide per month  | Same layout                                               |
| 16 | Next Steps           | Up to 3 action recommendations                            |

Slide count varies with how many years / months the data covers; the table
above shows the structure, not exact slide numbers.

### Copy generation

Three helpers produce the deck's prose, all in `index.html`:

- `buildTakeaway(s, showFairness, cfg, maxTime)` — one-liner for the All Time
  slide. Returns `{stat, descriptor}`.
- `buildFindings(s)` — up to three `{stat, insight}` pairs for the Executive
  Summary slide.
- `buildRecommendations(s)` — up to three `{action, plan}` pairs for the
  closing Next Steps slide.

All three share a single voice: **"Given [data], [direct verdict]. [Specific
levers]."** No hedging language. Every line references one of the five KPI
tiles the deck shows (Invitations, Attempts, Median Score, Passing Rate,
Median Time) — they never introduce a metric the reader hasn't seen on a
slide already.

### Logos

The dashboard and welcome screen `<img>` tags load
`assets/codility-logo-{dark,light}.png` directly. The PPTX export uses
`assets/logos.js` — a small generated file with the same two images encoded
as base64 data URIs, exposed on `window.CODILITY_LOGO_{DARK,LIGHT}_URI`.

To regenerate `logos.js` after replacing the PNGs:

```bash
python3 -c "
import base64
for variant in ['dark', 'light']:
    with open(f'assets/codility-logo-{variant}.png', 'rb') as f:
        data = base64.b64encode(f.read()).decode()
    print(f'window.CODILITY_LOGO_{variant.upper()}_URI = \"data:image/png;base64,{data}\";')
" > assets/logos.js
```

---

## Brand

Codility Brand v2.0 colors are defined twice — once as CSS variables in the
`<style>` block, and once as a `C` color object inside `downloadPPTX()` for
the PPTX export.

| Role                     | Hex       | Usage                                          |
| ------------------------ | --------- | ---------------------------------------------- |
| Ultramarine (lead)       | `#4A64E9` | Primary buttons, big numerals, accents         |
| Raspberry (eyebrow)      | `#C82372` | Section labels (`OVERVIEW`, `ALL TIME`, etc.)  |
| Yellow (highlight)       | `#FEE64B` | Reserved for callouts                          |
| Navy (structural)        | `#253641` | Footer, body text                              |

Font: Inter on the web, Aptos in the PPTX (Aptos is the Office default in
recent versions and renders cleanly across machines without an embed).

---

## Where to plug in the platform integration

When porting this into the platform, replace the upload flow with an API fetch.
Two functions are the only seams you need to change:

**1. Replace the parser.** The current parser is invoked when a file is
selected; it reads the XLS via SheetJS and assigns to `_allRows` /
`_integrityRows` / `_lastCfg`. Replace it with an API call that returns the
same row shape (each row is a flat object keyed by the column names in the
"Expected XLS shape" table above) and the same `_lastCfg` object (every key
in the "Configuration object" table).

**2. Skip the upload screen.** `showUpload()` and `showReport()` are the two
view-toggling functions. Call `showReport()` directly with `_allRows` and
`_lastCfg` already populated.

Everything downstream — the dashboard, the filters, the PPTX export — only
reads from those globals. It doesn't care where the data came from.

A reasonable platform-side API shape:

```
GET /api/tests/{testId}/health-export
→ {
    rows: Row[],          // shape matches the XLS column table above
    config: {             // matches _lastCfg
      creator, passingScore, maxTime, fairness, fairnessRel,
      rotation, proctoring, leaked, weighted, slotWeights
    }
  }
```

---

## Development notes

- **No build step.** Edit `index.html` in any editor and refresh.
- **No tests yet.** The app has been hand-verified by rendering the deck and
  inspecting slides. A future port should add unit tests around
  `computeStats`, `buildTakeaway`, `buildFindings`, and `buildRecommendations`.
- **Vendored libraries.** All three live in `assets/vendor/`. Versions are
  intentionally pinned to whatever was checked in — bumping them needs a
  visual smoke test of the dashboard and the deck.
- **Async error handling.** All download buttons go through `safeRun(fn, btn)`
  which catches rejections and surfaces them as alerts. Without this wrapper,
  unhandled rejections in async onclick handlers die silently in the console.
