# Test Health Analysis

A single-page web app that turns a Codility assessment data set into a branded
test-health report covering engagement, scoring, time pressure, integrity flags,
and similarity check, plus a PowerPoint deck ready to share with a customer.

The app is intentionally a single static HTML file with three vendored JS
libraries. No build step, no server, no framework. Open `index.html` in a
browser and it runs.

This README is written for the engineer who will port this into the Codility
platform so that data is pulled directly from the test's existing records
instead of uploaded by hand. The current shape (XLS upload → in-memory parse
→ render) maps cleanly onto a future API-fed version. **All the dashboard
and PPTX logic only reads from two in-memory globals (`_allRows` and
`_lastCfg`); the upload screen exists only to populate them. Once those two
globals are populated, everything downstream works the same way.**

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
    ├── logos.js                 # Same two PNGs as base64, needed by PPTX export
    └── vendor/
        ├── xlsx.full.min.js     # SheetJS, parses .xls / .xlsx / .csv
        ├── pptxgen.bundle.js    # PptxGenJS, generates the .pptx
        └── chart.umd.min.js     # Chart.js, score distribution chart on the dashboard
```

`assets/logos.js` is generated from the two PNGs so the PPTX export works
under all three serving modes: `http://`, `file://`, and offline. PptxGenJS
needs data URIs, and `fetch()` of a sibling file is blocked under `file://`.
Putting the data URIs in a separate `<script>` sidesteps the issue while
keeping `index.html` itself logo-free.

That is the whole repo. Everything in `assets/vendor/` is third-party;
everything else is project code.

---

## Running it

Open `index.html` in any modern browser. That is it.

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
small set of configuration values, and clicks **Generate Report**. The app
then:

1. Parses the file into rows of candidate attempt data.
2. Computes aggregate stats across All Time, by Year, and by Month.
3. Renders a single-page dashboard with KPI tiles, a score distribution
   chart, a Task Analysis table, an Assessment Health Insights list, and a
   three-tab Integrity panel (Integrity Risk, Similarity Check, Behavioural
   Signals).
4. On click, generates a PowerPoint deck and downloads it. Slide count
   varies with the data, typically 14 to 25 slides depending on the time
   range, the number of task variants, and whether the export contains
   integrity data.

---

## ⭐ Platform integration contract

This is the section to read first if you are porting this into the platform.

Every downstream piece (dashboard render, filters, PPTX export) reads from
exactly two in-memory globals defined in `index.html`:

- `_allRows`: an array of row objects, one per candidate attempt.
- `_lastCfg`: a single object describing the test-level config.

The upload screen and file parser exist only to populate those two globals.
Once they are populated, the rest of the app is data-driven. To replace the
upload screen with platform data, fetch the data, populate both globals,
then call `buildReport(_allRows, _lastCfg)`. That is the whole integration
seam.

### Required row fields (per candidate attempt)

The app reads each row by header name. Order does not matter; extra fields
are ignored. The columns below match Codility's existing test session export
exactly, so a platform-side API can serialize the same column names without
translation.

| Field                                              | Type     | Required? | What the app uses it for                                                  |
| -------------------------------------------------- | -------- | --------- | ------------------------------------------------------------------------- |
| `Test`                                             | string   | Yes       | The assessment's display name, shown in dashboard headers and PPTX title  |
| `Create date`                                      | datetime | Yes       | Year and month bucketing for filters and Year-by-Year / Month-by-Month slides |
| `Start date`                                       | datetime | No        | Whether the candidate started the test. A blank value means **invited but did not attempt** and feeds drop-off |
| `% total score`                                    | number   | Yes       | Median score, distribution histogram, score buckets, passing rate, score SD |
| `Time Used (minutes)`                              | number   | Yes       | Median time, time pressure analysis                                        |
| `Integrity Risk Status`                            | string   | No        | Integrity Risk tab + Section 04 (Integrity Impact). Values expected: `high`, `moderate`, `low`, `no_risk`, blank. Section 04 is rendered only when at least one row has a non-blank value |
| `Similarity Check`                                 | string   | No        | Similarity Check tab. Values expected: `please resolve`, `acknowledged`, `skipped`, `in review`, blank |
| `Detected pasting code into Codility assessment`   | string   | No        | Behavioural Signals tab. Truthy values: `detected`, `yes`, `true`, `1` |
| `Detected attempt to copy task description`        | string   | No        | Behavioural Signals tab                                                    |
| `Detected leaving Codility assessment tab`         | string   | No        | Behavioural Signals tab                                                    |
| `Detected abnormally little time spent on task`    | string   | No        | Behavioural Signals tab                                                    |
| `Detected suspicious typing pattern`               | string   | No        | Behavioural Signals tab                                                    |
| `Detected use of AI assistant`                     | string   | No        | Behavioural Signals tab                                                    |
| `Task N name` (one per task slot, N starting at 1) | string   | Yes       | Task Analysis table and PPTX Task Analysis slide. Multiple distinct values per slot indicate task rotation |
| `Task N score %`                                   | number   | Yes       | Per-task average score and discrimination (Pearson r vs. total score)     |
| `Task N difficulty`                                | string   | No        | Color-coded difficulty chip on the dashboard and PPTX. Expected: `elementary`, `easy`, `medium`, `hard` |
| `Task N language used`                             | string   | No        | Most common language used per task, shown on the PPTX Task Analysis slide |

**The minimum useful set is `Test`, `Create date`, `% total score`, plus
`Task N name` / `Task N score %` for at least one task slot.** Everything
else degrades gracefully.

### Test-level config object (`_lastCfg`)

This is the object the upload form populates today. When porting, replace
the form with whatever your platform stores (test settings, customer-level
overrides, etc.).

| Key             | Type                       | Required? | Source today           | What it controls                                                              |
| --------------- | -------------------------- | --------- | ---------------------- | ----------------------------------------------------------------------------- |
| `creator`       | string                     | No        | Test Creator Name input  | Subheader on dashboard and on every PPTX slide ("Test created by ...")        |
| `passingScore`  | number \| null             | No        | Passing Score input    | Passing rate calculation. Rows with `% total score >= passingScore` pass. `null` skips the calculation |
| `maxTime`       | number \| null             | No        | Max Test Duration input | Time pressure analysis. `medTimeLeft = maxTime − medTime`. `null` hides the time-pressure badge |
| `fairness`      | number \| null             | No        | Fairness Rating input  | Fairness rating displayed in Assessment Experience. The platform should populate this from the customer-reported fairness score in the assessment dashboard. `null` hides the tile |
| `fairnessRel`   | `'above' \| 'below' \| null` | No      | Fairness vs Global Avg | Whether the fairness rating is above or below the global average. Drives the color of the Fairness Rating badge |
| `rotation`      | boolean                    | No        | Task rotation toggle   | Surfaces a "task rotation enabled" badge. Auto-detected from task slot uniqueness when this field is missing |
| `proctoring`    | boolean                    | No        | Proctoring toggle      | Configuration insight on the dashboard. The platform should populate this from the test's proctoring settings |
| `leaked`        | boolean                    | No        | Leaked item toggle     | Reserved. Surfaces a "tasks may be leaked" warning under Question Quality |
| `weighted`      | boolean                    | No        | Weighted scoring toggle | If `true`, applies per-task-slot weights to the score calculation             |
| `slotWeights`   | `{[n: number]: number}`    | No        | Per-slot Weight inputs | Map of task slot number to percentage weight. Should sum to 100 when `weighted` is `true` |

### Suggested API shape

```jsonc
GET /api/tests/{testId}/health-export
→ {
    "rows": [
      {
        "Test": "Spring Boot (Java) Senior (Nov,2025)",
        "Create date": "2025-11-04 09:23:00",
        "Start date":  "2025-11-04 09:30:00",
        "% total score": 78,
        "Time Used (minutes)": 165,
        "Integrity Risk Status": "low",
        "Similarity Check": "skipped",
        "Task 1 name": "SpringHealthcheck",
        "Task 1 score %": 90,
        "Task 1 difficulty": "easy",
        "Task 1 language used": "Java"
        // ... behavioural signal columns, more tasks, etc.
      }
      // ... one row per candidate attempt
    ],
    "config": {
      "creator": "Dylan Caldeira",
      "passingScore": 70,
      "maxTime": 180,
      "fairness": 79,
      "fairnessRel": "above",
      "rotation": false,
      "proctoring": true,
      "leaked": false,
      "weighted": false,
      "slotWeights": {}
    }
  }
```

### Wiring it up

The two seams to change:

**1. Skip the upload screen.** `showUpload()` and `showReport()` are the two
view-toggling functions in `index.html`. The platform port should hide
`#upload-screen` permanently and show `#report-screen` directly.

**2. Replace the parser.** The file-handling code (`handleFile()`,
`parseCSV()`, the SheetJS path) is invoked when a file is selected and it
populates `parsedRows` plus `_lastCfg` from the form. Replace that path
with a fetch from your API and assign the response directly:

```js
const res = await fetch(`/api/tests/${testId}/health-export`);
const { rows, config } = await res.json();
buildReport(rows, config);   // that is the whole integration
```

`buildReport()` does everything from there: detects task slots, computes
stats, renders the dashboard, and arms the PPTX download button.

---

## Data flow

```
┌─────────────────┐     ┌─────────────────┐     ┌──────────────┐
│  _allRows       │ ──▶ │  computeStats() │ ──▶ │  renderHTML  │
│  + _lastCfg     │     │  (in-memory)    │     │  +  Chart.js │
└─────────────────┘     └─────────────────┘     └──────────────┘
                                │
                                ▼
                        ┌──────────────────┐
                        │  downloadPPTX()  │
                        │   (PptxGenJS)    │
                        └──────────────────┘
```

State lives in module-level globals in `index.html`'s `<script>`:

| Global              | What it holds                                                      |
| ------------------- | ------------------------------------------------------------------ |
| `_allRows`          | Every row from the source data (one per candidate attempt)          |
| `_lastCfg`          | The test-level config object (see contract above)                  |
| `_integrityRows`    | Rows surfaced for the Integrity Risk tab (date-filtered, pre-exclusion) |
| `_activeYears`      | Set of years currently selected in the Year filter                  |
| `_activeMonths`     | Set of month numbers currently selected in the Month filter         |
| `_lastRenderData`   | The full data object passed to `render()`. Used by the PPTX builder for slot data |
| `LOGO_DATA_URIS`    | `{dark, light}` data URIs, populated at startup from `logos.js`     |

When the user changes a filter or toggles an integrity exclusion, the
dashboard re-runs `buildReport()` over the filtered subset of `_allRows` and
re-renders. Nothing is persisted to disk or to localStorage.

---

## Calculations

All metric definitions live alongside their use sites in `index.html`. This
section is a single source of truth for what each metric means and what
threshold drives each badge color.

### Engagement

```
invitations         = _allRows.length
attempts            = rows where Start date is present AND % total score is numeric
non-zero            = attempts where % total score > 0
drop-off %          = round((1 − attempts / invitations) × 100)
non-zero drop-off % = round((1 − non-zero / attempts) × 100)
completion rate %   = round((attempts / invitations) × 100)
```

Drop-off badge thresholds (dashboard, invitations → attempts):

| Drop-off %     | Badge color | Label        |
| -------------- | ----------- | ------------ |
| > 50           | red         | Very high    |
| > 25           | amber       | High         |
| <= 10          | green       | Healthy      |
| else           | gray        | (count only) |

Drop-off badge thresholds (attempts → non-zero):

| Drop-off %  | Badge color | Label        |
| ----------- | ----------- | ------------ |
| > 30        | red         | Very high    |
| > 15        | amber       | High         |
| <= 5        | green       | Healthy      |
| else        | gray        | (count only) |

### Scoring

```
medScore     = median of % total score across all attempts (including zeros)
passingRate  = pct of attempts where % total score >= cfg.passingScore
                (null when no passingScore is configured)
scoreSD      = standard deviation of % total score, rounded to whole percent
```

The score distribution histogram has 11 buckets: `0%`, `1-10%`, `11-20%`, …
`91-100%`. The `0%` bucket is exact-zero only; the remaining buckets are
ceil-based.

Median score difficulty badge:

| Median score | Badge color | Label              |
| ------------ | ----------- | ------------------ |
| < 20         | red         | High Difficulty    |
| > 85         | red         | Low Difficulty     |
| 40 to 75     | green       | Good Difficulty    |
| else         | gray        | Med Difficulty     |

Score SD spread badge:

| Score SD     | Badge color | Label              |
| ------------ | ----------- | ------------------ |
| < 10         | green       | Tight Cluster      |
| < 20         | green       | Healthy Spread     |
| < 35         | amber       | Wide Spread        |
| else         | red         | Bimodal / Uneven   |

Score distribution shape detector (chart-meta line). Checked in this order:

| Condition                                            | Badge color | Label                                          |
| ---------------------------------------------------- | ----------- | ---------------------------------------------- |
| 0% bucket >= 20% AND 91-100% bucket >= 20%           | red         | U-shaped (X% at 0, Y% at 91-100)               |
| 91-100% bucket >= 30% AND 0% bucket < 20%            | amber       | Top-heavy (X% at 91-100)                       |
| 0% bucket >= 30% AND 91-100% bucket < 20%            | amber       | Bottom-heavy (X% at 0)                         |
| Score SD < 10                                        | green       | Tight cluster                                  |
| Score SD < 20                                        | green       | Healthy spread                                 |
| Score SD < 35                                        | amber       | Wide spread                                    |
| else                                                 | red         | Uneven distribution                            |

Passing rate badge:

| Passing rate     | Badge color | Label       |
| ---------------- | ----------- | ----------- |
| < 20             | red         | Very Low    |
| > 90             | red         | Very High   |
| 40 to 70         | green       | Good Range  |

### Time

```
medTime      = median of Time Used (minutes) across all attempts that started
medTimeLeft  = cfg.maxTime − medTime
```

Time pressure badge (only shown when `cfg.maxTime` is present):

| medTimeLeft  | Badge color | Label              |
| ------------ | ----------- | ------------------ |
| <= 5         | red         | High Time Pressure |
| <= 10        | green       | Fine Time Pressure |
| else         | amber       | Too Much Time      |

### Per-task analytics

For each `Task N` slot:

```
For each unique Task N name in the slot:
  attempts_for_variant = rows with that Task N name and a numeric Task N score %
  avg_score            = mean(Task N score %) across attempts_for_variant
  discrimination       = pearson(Task N score %, % total score) across attempts_for_variant
  difficulty           = mode(Task N difficulty) across attempts_for_variant
  language             = mode(Task N language used) across attempts_for_variant
  N                    = attempts_for_variant.length
```

Discrimination color:

| Pearson r       | Color  |
| --------------- | ------ |
| >= 0.75         | green  |
| 0 to 0.75       | ink    |
| < 0             | red    |

Most assessments only have 3 to 5 task variants, so the green threshold is
set conservatively. Anything between 0 and 0.75 is normal but not
exceptional; the green color is reserved for tasks doing genuinely strong
discriminating work.

Difficulty chip color (Task / Difficulty column on the dashboard and PPTX):

| Task N difficulty | Color           |
| ----------------- | --------------- |
| `elementary`      | mint / teal     |
| `easy`            | green           |
| `medium`          | amber           |
| `hard`            | raspberry       |

Anything else (blank, unrecognised, typos) renders as a plain uncolored pill.

### Integrity

```
high-risk count   = number of attempts where Integrity Risk Status == 'high'
high-risk pct     = high-risk count ÷ (attempts with non-blank Integrity Risk Status)
```

Section 04 (Integrity Impact) computes a "high-risk excluded" cohort by
filtering out every row where `Integrity Risk Status == 'high'` and re-running
all the headline KPIs against the cleaned cohort. It then surfaces the deltas
on a side-by-side slide. **The section is gated on
`integrityWithData > 0`**: if no rows in the export carry any non-blank
Integrity Risk Status value, the section (divider + slide + TOC entry) is
silently skipped so the deck still works for tests without integrity signals.

### Completion Rate insight (dashboard)

```
completion %  = round((attempts / invitations) × 100)
```

| Completion % | Insight cls | Wording                      |
| ------------ | ----------- | ---------------------------- |
| >= 90        | ok          | "Engagement looks great."    |
| >= 80        | ok          | "Engagement looks good."     |
| >= 75        | warn        | "Engagement is borderline."  |
| else         | warn        | "Engagement is concerning."  |

---

## Generating the PPTX

`downloadPPTX()` builds a deck with PptxGenJS. Slide count varies with how
many years and months the data covers, the number of task variants, and
whether the export contains integrity data; the table below shows the
structure, not exact numbers.

| #  | Slide                | What is on it                                                    |
| -- | -------------------- | ---------------------------------------------------------------- |
| 1  | Cover                | Test name, creator, "First used on" + "Data as of" dates, dark ultramarine background |
| 2  | Executive Summary    | 5 KPI tiles plus up to 3 Key Findings                             |
| 3  | Table of Contents    | Section 01 / 02 / 03 (and 04 if integrity data is present)       |
| 4  | Section 01 divider   | "ALL TIME"                                                       |
| 5  | All Time stats       | 3 stat cards, score distribution, **Key Takeaway** band          |
| …  | Task Analysis        | Per-task table (name, difficulty, most common language, avg, discrimination, N). Paginates if there are 15+ variants |
| …  | Section 02 divider   | "YEAR BY YEAR"                                                   |
| 7… | One slide per year   | Same layout as the All Time slide, no Key Takeaway               |
| …  | Section 03 divider   | "MONTH BY MONTH"                                                 |
| …  | One slide per month  | Same layout                                                      |
| …  | Section 04 divider   | "INTEGRITY IMPACT" (only when the export has integrity data)     |
| …  | Integrity Impact     | Side-by-side 5-KPI comparison, all candidates vs. high-risk excluded |
| …  | Next Steps           | Up to 3 action recommendations                                   |

### Slide-header meta block

Every content slide has a right-side meta block in the header. Two variants:

- **Full meta** (Executive Summary, All Time stat slide, cover slide): test
  name → "Test created by X" → "First used on \<earliest attempt date>" →
  "Data as of \<today's date>" → yellow **Outdated** pill if the earliest
  attempt is more than one calendar year old.
- **Minimal meta** (TOC, Year, Month, Task Analysis, Integrity Impact, Next
  Steps): test name → "Data as of \<today's date>".

The full block is gated to high-signal slides on purpose — repeating the
creator name and first-used date on every page would clutter the deck. The
"Outdated" caution surfaces once on the Executive Summary and once on the
All Time stat slide, where it has the most impact. The dashboard mirrors
this with its own yellow Outdated pill in the meta badge.

"First used on" uses the earliest `Start date` in the export; "Data as of"
uses the current date (i.e. when the user clicked Download PPTX).

### Section 04, Integrity Impact (conditional)

A two-column slide that surfaces how the headline KPIs shift when candidates
flagged with **High** Integrity Risk Status are removed from the cohort.

Each KPI tile in the right column gets a delta chip:

- **Green** when Median Score or Passing Rate goes **up** after exclusion
  (high-risk attempts were dragging the headline numbers down).
- **Red** when those same metrics go **down** (the flagged candidates were
  posting competitive scores and were inflating the read).
- **Gray (neutral)** for population counts and Median Time.

A Key Takeaway band at the bottom interprets the most informative delta in
plain prose. When fewer than 3 points move on either Median Score or Passing
Rate, the takeaway falls back to "the headline KPIs barely move".

### Task Analysis slide (Section 01)

Mirrors the dashboard's Task Analysis card. Single-variant slots collapse to
one row (slot number shown in the `#` column); rotation slots get a divider
header row above their variants. The layout uses three density levels
(comfortable / medium / tight) chosen automatically based on how many rows
the data produces. If even tight density does not fit on one slide, the
section paginates with a "(2 of 3)" suffix on the slide title.

### Copy generation

Four helpers produce the deck's prose, all in `index.html`:

- `buildTakeaway(s, showFairness, cfg, maxTime)`: one-liner for the All Time
  slide. Returns `{stat, descriptor}`.
- `buildFindings(s)`: up to three `{stat, insight}` pairs for the Executive
  Summary slide.
- `buildRecommendations(s)`: up to three `{action, plan}` pairs for the
  closing Next Steps slide.
- `buildInsightsData(d)`: the six items rendered in the Assessment Health
  Insights card on the dashboard.

All four share a single voice: third person, non-contracted, same length
across cases. Every line references one of the five KPIs surfaced in the
deck or on the dashboard (Invitations, Attempts, Median Score, Passing
Rate, Median Time), so the copy never introduces a metric the reader has
not already seen.

### Logos

The dashboard and welcome screen `<img>` tags load
`assets/codility-logo-{dark,light}.png` directly. The PPTX export uses
`assets/logos.js`, a small generated file with the same two images encoded
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

Codility Brand v2.0 colors are defined twice: once as CSS variables in the
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

## Development notes

- **No build step.** Edit `index.html` in any editor and refresh.
- **No tests yet.** The app has been hand-verified against two real exports
  (Spring Boot Senior with 791 rows and C# Senior with 3,181 rows) using a
  Python harness that mirrors the JS calculation logic. A future port
  should add unit tests around `computeStats`, `buildTakeaway`,
  `buildFindings`, `buildRecommendations`, and `buildInsightsData`.
- **Vendored libraries.** All three live in `assets/vendor/`. Versions are
  intentionally pinned to whatever was checked in. Bumping them needs a
  visual smoke test of the dashboard and the deck.
- **Async error handling.** All download buttons go through
  `safeRun(fn, btn)` which catches rejections and surfaces them as alerts.
  Without this wrapper, unhandled rejections in async onclick handlers die
  silently in the console.
