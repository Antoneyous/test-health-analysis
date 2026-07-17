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
├── index.html                       # The entire app: HTML, CSS, JS in one file
├── README.md                        # This file
├── .gitignore
└── assets/
    ├── codility-logo-dark.png       # Rendered wordmark on light surfaces
    ├── codility-logo-light.png      # Rendered wordmark on dark surfaces
    ├── codility-wordmark-dark.svg   # Vector source of the dark wordmark
    ├── codility-wordmark-light.svg  # Vector source of the light wordmark
    ├── logos.js                     # Logo PNGs + banding SVG as base64 data URIs
    ├── fonts/
    │   ├── JetBrainsMono-Regular.ttf
    │   ├── JetBrainsMono-Medium.ttf
    │   ├── JetBrainsMono-SemiBold.ttf
    │   ├── JetBrainsMono-Bold.ttf
    │   └── JetBrainsMono-ExtraBold.ttf
    ├── patterns/
    │   ├── banding-element.svg        # Native pale variant (`#F1F4F7`)
    │   └── banding-element-paper.svg  # Slightly darker variant for white surfaces
    └── vendor/
        ├── xlsx.full.min.js         # SheetJS, parses .xls / .xlsx / .csv
        ├── pptxgen.bundle.js        # PptxGenJS, generates the .pptx
        └── chart.umd.min.js         # Chart.js: score distribution, cut score curve, per-task histograms
```

`assets/logos.js` is generated from the two PNGs + the banding SVG so the
PPTX export works under all three serving modes: `http://`, `file://`, and
offline. PptxGenJS needs data URIs, and `fetch()` of a sibling file is
blocked under `file://`. Putting the data URIs in a separate `<script>`
sidesteps the issue while keeping `index.html` itself asset-free.

Fonts are self-hosted (JetBrains Mono) so the brand display font renders
correctly without external requests. Inter is loaded via Google Fonts CDN
as a fallback in `<style>`.

The banding-element SVGs are no longer used by the current PPTX layout
(the brand-faithful redesign uses a sky-circle on the cover instead).
They are kept on disk for potential future use; `logos.js` still exposes
their data URIs but no slide function calls them.

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

The user lands on a configuration screen (operator-only), picks an exported
assessment file, fills in a small set of test-level configuration values,
and clicks **Generate report**. The app then:

1. Parses the file into rows of candidate attempt data.
2. Computes aggregate stats across All Time, by Year, and by Month.
3. Renders a single-page dashboard with KPI tiles, a score distribution
   chart with a passing-threshold marker, an interactive Cut Score
   Analysis panel (pass-rate curve + slider), a Task Analysis table whose
   rows expand into per-task score distribution histograms, an
   Assessment Health Insights list, and a three-tab Integrity panel
   (Integrity Risk, Similarity Check, Behavioural Signals). A date
   filter (preset pills + custom From/To month range) re-scopes the
   whole dashboard by invitation date.
4. On click, generates a PowerPoint deck and downloads it. Slide count
   varies with the data, typically 15 to 27 slides depending on the time
   range, the number of task variants, and whether the export contains
   integrity data. The deck mirrors the dashboard: stat slides per
   period, Task Analysis, a Candidate Funnel closing the All Time
   section, an Assessment Health Insights summary, and Next Steps.

The configuration screen is intentionally separate from the report screen:
on the production platform port, the configuration values come from the
test's existing settings, and the configuration UI can be dropped
entirely.

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

| Key                | Type                          | Required? | Source today              | What it controls                                                              |
| ------------------ | ----------------------------- | --------- | ------------------------- | ----------------------------------------------------------------------------- |
| `creator`          | string                        | No        | Test creator name input   | Subheader on dashboard and on the PPTX cover/Executive Summary ("Test created by ...") |
| `passingScore`     | number \| null                | No        | Passing score input       | Passing rate calculation. Rows with `% total score >= passingScore` pass. Drives the red dashed threshold line on the score-distribution chart. `null` skips the calculation |
| `maxTime`          | number \| null                | No        | Max test duration input   | Time pressure analysis. `medTimeLeft = maxTime − medTime`. `null` hides the time-pressure badge |
| `fairness`         | number \| null                | No        | Fairness rating input     | Fairness rating displayed in Assessment Experience. The platform should populate this from the post-assessment survey ("share of candidates who said the test fairly evaluated their coding skills"). `null` hides the tile |
| `fairnessRel`      | `'above' \| 'below' \| null`  | No        | Fairness vs global average | Whether the fairness rating is above or below the global average. Drives the color of the Fairness Rating badge |
| `rotation`         | boolean                       | No        | Task rotation toggle      | Whether task slots serve rotated variants. Auto-detected from task-slot uniqueness when this field is missing; the upload screen pre-fills this based on detection |
| `proctoringLevel`  | `0 \| 1 \| 2 \| 3`            | No        | Proctoring level picker   | 0=Off (warning insight). 1=Behavioural signals. 2=Level 1 + extended live multimedia (webcam snapshots and/or screen recording). 3=Level 2 + identity verification (full webcam/mic + ID check). Drives the Configuration insight wording and tone, including a "consider moving to Level N+1" suggestion for levels 1-2 |
| `leakedTasks`      | `string[]`                    | No        | Leaked tasks multi-select | Array of task names the operator has marked as leaked. Each name in the array is shown with a "Leaked" badge on the dashboard task table and triggers a "X should be replaced because it has been leaked" line on the Question Quality insight |
| `weighted`         | boolean                       | No        | Weighted scoring toggle   | When `true`, the Weight column on the task table is shown                     |
| `slotWeights`      | `{[n: number]: number}`       | No        | Per-slot weight inputs    | Map of task slot number to percentage weight. Should sum to 100 when `weighted` is `true` |

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
      "proctoringLevel": 2,
      "leakedTasks": [],
      "weighted": false,
      "slotWeights": {}
    }
  }
```

### Wiring it up

The two seams to change:

**1. Skip the configuration screen.** `showUpload()` toggles back to the
upload view; the report view is shown at the end of `render()` (which
`buildReport()` calls). The platform port should hide `#upload-screen`
permanently — calling `buildReport(rows, config)` already switches the
view to `#report-screen`.

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

`downloadPPTX()` is also exposed globally. The Download PPTX button on
the dashboard topbar wires straight to it, and the platform port can
call it from anywhere.

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

| Global                       | What it holds                                                      |
| ---------------------------- | ------------------------------------------------------------------ |
| `_allRows`                   | Every row from the source data (one per candidate attempt)          |
| `_lastCfg`                   | The test-level config object (see contract above)                  |
| `_integrityRows`             | Rows surfaced for the Integrity Risk tab (date-filtered, pre-exclusion) |
| `_dateRange`                 | `{from: Date\|null, to: Date\|null}` — the single active date filter range (inclusive) |
| `_activePreset`              | Id of the active date preset (`'all'`, `'30d'`, `'90d'`, `'6m'`, `'12m'`, `'ytd'`), or `null` when a custom From/To range is set |
| `_dataMinDate`/`_dataMaxDate`| Earliest / latest `Create date` in the file. Presets are anchored to `_dataMaxDate` |
| `_lastRenderData`            | The full data object passed to `render()`. Used by the PPTX builder and the per-task distribution rows for slot data |
| `_cutCurve`                  | Precomputed pass-rate curve: `_cutCurve[c] = {count, pct}` of attempts with `% total score >= c`, for c = 0..100 |
| `cutChartInstance`           | Chart.js instance of the Cut Score Analysis curve                   |
| `taskDistCharts`             | Map of canvas id → Chart.js instance for open per-task distribution rows. All destroyed whenever the task table re-renders |
| `LOGO_DATA_URIS`             | `{dark, light}` data URIs, populated at startup from `logos.js`     |

When the user changes a filter or toggles an integrity exclusion, the
dashboard re-runs `buildReport()` over the filtered subset of `_allRows` and
re-renders. Nothing is persisted to disk or to localStorage.

### Date filter

The date filter is a single `[from, to]` range over `Create date`
(invitation date), driven by two inputs that are mutually exclusive in
the UI:

- **Preset pills** — All time, Last 30 days, Last 90 days, Last 6
  months, Last 12 months, YTD (`DATE_PRESETS` in `index.html`). Presets
  are anchored to the **latest invitation date in the data**
  (`_dataMaxDate`), not today, so "Last 90 days" means the last 90 days
  of the data and never produces an empty report for an older export.
- **From/To month pickers** (`<input type="month">`) for custom ranges.
  From resolves to the first instant of its month, To to the last
  instant of its month, so both endpoint months are fully included.
  Selecting a custom range deselects the preset and vice versa.

`getFilteredRows(rows)` applies the range; `setDatePreset(id)`,
`onRangeInput()` and `clearDateFilter()` mutate it. Cross-year ranges
(e.g. Nov 2025 → Feb 2026) work correctly — this replaced an earlier
year-chips × month-chips design that could not express ranges spanning
a year boundary.

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

### Cut Score Analysis (dashboard)

The Cut Score Analysis card (side by side with the Live Score
Distribution) renders a pass-rate curve plus an interactive slider:

```
for c in 0..100:
  _cutCurve[c] = { count: attempts with % total score >= c,
                   pct:   count / attempts × 100 }
```

The pass rule is `>=`, matching the passing-rate KPI. The curve is a
Chart.js line chart over a linear 0–100 x-axis; a slider moves a red
dashed marker (custom `cutMarker` plugin reading `chart.$cut`, updated
via `chart.update('none')` so dragging is cheap) and drives a live
readout: "Cut score at 60% → 41.6% of candidates would pass (37 pass,
52 fail, of 89 scored attempts)". The slider initialises to
`cfg.passingScore` when set, else 60. The curve is recomputed inside
`renderCutScoreAnalysis()` on every `render()`, so it always reflects
the current date-filtered cohort. Relevant functions:
`renderCutScoreAnalysis`, `onCutSliderInput`, `updateCutReadout`.

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

**Task chips (platform style).** Task names render as white chips with a
grey hairline border, dark text, and a colored left stripe carrying the
category — matching the platform's "Tasks selected for the test" panel.
Same treatment on the dashboard (CSS `.task-name-chip` + `.chip-*`
classes) and on the PPTX Task Analysis slide (white rounded rect +
stripe rect + label).

| Category                     | Stripe color            | Hex       |
| ---------------------------- | ----------------------- | --------- |
| Custom content               | periwinkle              | `#AAB4F1` |
| `elementary`                 | aqua-400                | `#BEE5FA` |
| `easy`                       | mint-600                | `#A2EBD0` |
| `medium`                     | sunset-600              | `#FEE64B` |
| `hard`                       | strawberry-600          | `#FC9797` |
| blank / unrecognised         | coolgrey-300            | `#D8DFE4` |

**Custom content detection** (`isCustomTask(name)`): the export carries no
explicit custom-content flag, so the app uses a naming heuristic —
Codility library tasks are single CamelCase identifiers
(`MatchingPlates`, `WordGame`) while customer-authored tasks have
human-readable titles containing spaces ("Power Platform and Copilot
Studio Graduate Screening Test Set 1"). A name containing a space is
treated as custom. Custom wins over the difficulty stripe, matching the
platform, and the PPTX chip label reads "Custom" instead of the
difficulty. If the platform port has a real is-custom flag on the task
record, replace this heuristic with it.

### Per-task score distribution (expandable rows)

Every variant row in the Task Analysis table is clickable. Clicking
inserts an inline row beneath it with a histogram of that task's own
scores (same 11 buckets as the main chart: exact-zero bucket + ten
ceil-based deciles) and a meta line (attempts, median, avg, SD). Several
rows can be open at once for side-by-side comparison; clicking again
collapses. The raw per-attempt scores are already retained on each
variant (`taskScores` on the objects in `_lastRenderData.slots`), so no
recomputation is needed. Chart instances are tracked in `taskDistCharts`
and destroyed whenever the table re-renders (filter change, re-run).
Relevant functions: `toggleTaskDist`, `renderTaskDistChart`.

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
| 1  | Cover                | Codility wordmark top-left, "TEST HEALTH ANALYSIS" eyebrow, test name ending in sky-blue `_`, "Test created by" + "First used on" + "Data as of" meta at the bottom, paper-soft background with a large sky-blue circle bleeding off the top-right |
| 2  | Executive Summary    | 5 KPI tiles plus up to 3 Key Findings                             |
| 3  | Table of Contents    | Section 01 / 02 / 03 (and 04 if integrity data is present)       |
| 4  | Section 01 divider   | "ALL TIME"                                                       |
| 5  | All Time stats       | 3 stat cards, score distribution, **Key Takeaway** band          |
| …  | Task Analysis        | Per-task table (name, difficulty, most common language, avg, discrimination, N). Paginates if there are 15+ variants |
| …  | Candidate Funnel     | Closes Section 01. Platform-style wedge funnel on a shared bottom baseline: Candidates invited → Assessments taken → Non zero results → Results over Passing Score X% (last stage only when `cfg.passingScore` is set). Each stage's top edge slopes to the next stage's level. **Display scaling:** when every stage retains > 60% of invited, display heights are remapped linearly from `[minShare..1]` to `[0.5..1]` so a high-retention funnel (100% → 98% → 97%) doesn't render as identical blocks; the printed counts and percentages always carry the true numbers. Funnels with real drops keep true proportions |
| …  | Section 02 divider   | "YEAR BY YEAR"                                                   |
| 7… | One slide per year   | Same layout as the All Time slide, no Key Takeaway               |
| …  | Section 03 divider   | "MONTH BY MONTH"                                                 |
| …  | One slide per month  | Same layout                                                      |
| …  | Section 04 divider   | "INTEGRITY IMPACT" (only when the export has integrity data)     |
| …  | Integrity Impact     | Side-by-side 5-KPI comparison, all candidates vs. high-risk excluded |
| …  | Assessment Health Insights | 2×3 grid of the same six insights as the dashboard card, via the shared `buildInsightsData()` builder. Acts as the summary the Next Steps slide acts on |
| …  | Next Steps           | Up to 3 action recommendations                                   |

### Slide-header meta block

Every content slide has a right-side meta block in the header. Two
variants are used:

- **Full meta** (Executive Summary, All Time stat slide): test name →
  "Test created by X" → "First used on \<earliest attempt date>" → "Data
  as of \<today's date>" → yellow **Outdated** pill if the earliest
  attempt is more than one calendar year old.
- **Minimal meta** (TOC, Year, Month, Task Analysis, Integrity Impact,
  Next Steps): test name → "Data as of \<today's date>".

The full block is gated to high-signal slides on purpose: repeating the
creator name and first-used date on every page would clutter the deck.
The "Outdated" caution surfaces once on the Executive Summary and once
on the All Time stat slide, where it has the most impact. The dashboard
mirrors this with its own yellow Outdated pill in the meta badge.

The cover slide carries its own bottom-anchored meta strip (creator
left, dates + Outdated pill right). It does not use `addFullMeta` /
`addMinMeta` because its layout is bespoke.

Long testnames (longer than 40 characters) are truncated with an
ellipsis in the meta block to prevent wrapping. The cover still renders
the full untruncated name and sizes the title down based on character
count so it lays out cleanly on at most 2 lines.

"First used on" uses the earliest `Start date` in the export; "Data as
of" uses the current date (i.e. when the user clicked Download PPTX).

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

Task chips on this slide use the platform style described in
[Per-task analytics](#per-task-analytics): white rounded rect, hairline
border, colored left-stripe by category, with custom content detected
via `isCustomTask()` and labelled "Custom" with the periwinkle stripe.
The stripe hex values live in the `C` palette
(`C.stripeCustom`, `C.stripeEasy`, …).

The Score Distribution bar chart on the All Time / Year / Month slides
uses uniform periwinkle bars (`C.periwinkle`, `#7B88EA`) to match the
dashboard and the platform's overall score chart. The dashboard's Cut
Score Analysis panel is **not** yet mirrored in the deck.

### Candidate Funnel slide (Section 01)

Built by `funnelSlide()`. Wedge geometry: each stage is a bottom-anchored
rectangle at the *next* stage's level plus a right-triangle carrying the
sloped drop (the last stage is flat). Stage tones walk the Sunset
progression palest → strongest. See the slide table above for the
display-scaling rule that keeps high-retention funnels legible; the rule
lives in `funnelSlide()` next to the geometry.

### Assessment Health Insights slide (closing)

Built by `insightsSlide()`, rendered immediately before Next Steps. It
assembles a dashboard-shaped data object from `_overall` + `cfg` +
`_lastRenderData.slots` and feeds it to the shared `buildInsightsData()`
builder, so the slide copy is always identical to the dashboard's
Assessment Health Insights card — one source of truth, no drift. Layout
is a 2×3 grid of cards with the same ok / warn / neutral iconography.

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

### Logos and brand assets

The dashboard and configuration screen `<img>` tags load
`assets/codility-logo-{dark,light}.png` directly. The PPTX export uses
`assets/logos.js`, a generated file with the logo PNGs and the
banding-element SVGs encoded as base64 data URIs and exposed on:

- `window.CODILITY_LOGO_DARK_URI`
- `window.CODILITY_LOGO_LIGHT_URI`
- `window.CODILITY_BANDING_URI` (the native pale variant, currently
  unused by the deck layout but kept for future reuse)
- `window.CODILITY_BANDING_PAPER_URI` (the slightly darker variant)

To regenerate `logos.js` after replacing the PNGs or the banding SVG:

```bash
python3 -c "
import base64, json
def to_uri_png(path):
    with open(path, 'rb') as f:
        return 'data:image/png;base64,' + base64.b64encode(f.read()).decode()
def to_uri_svg(path):
    with open(path, 'rb') as f:
        return 'data:image/svg+xml;base64,' + base64.b64encode(f.read()).decode()
print('window.CODILITY_LOGO_DARK_URI    = ' + json.dumps(to_uri_png('assets/codility-logo-dark.png'))    + ';')
print('window.CODILITY_LOGO_LIGHT_URI   = ' + json.dumps(to_uri_png('assets/codility-logo-light.png'))   + ';')
print('window.CODILITY_BANDING_URI      = ' + json.dumps(to_uri_svg('assets/patterns/banding-element.svg')) + ';')
print('window.CODILITY_BANDING_PAPER_URI= ' + json.dumps(to_uri_svg('assets/patterns/banding-element-paper.svg')) + ';')
" > assets/logos.js
```

---

## Brand

Codility's 2026 brand colors are defined twice: once as CSS variables in
the `<style>` block, and once as a `C` color object inside `downloadPPTX()`
for the PPTX export. They are derived from the marketing system's
`colors_and_type.css`.

| Role                   | Hex       | Where it shows up                                                              |
| ---------------------- | --------- | ------------------------------------------------------------------------------ |
| Ultramarine (primary)  | `#4A64E9` | Primary buttons, big stat numbers on the PPTX, link colors, active filter pills |
| Ultramarine-800 (deep) | `#30418C` | The dashboard upload-screen hero surface                                       |
| Periwinkle (data-viz)  | `#7B88EA` | Overall score distribution bars (dashboard + PPTX) and the cut score curve — matches the platform's overall score chart |
| Sunset-600 (data-viz)  | `#FEE64B` | Per-task score distribution bars — matches the platform's "Score distribution by task" chart |
| Brand ink              | `#253641` | Section-divider slide background, 2pt hairline above stat numbers, body text   |
| Sky-blue (underscore)  | `#80D5FF` | The `_` mark after titles on light surfaces, section-divider eyebrows          |
| Sky-200 (cover circle) | `#C7E6FA` | Large circle decoration bleeding off the top-right of the PPTX cover           |
| Sunset yellow          | `#FFF196` | The `_` mark after titles on dark section-divider slides                       |
| Raspberry (secondary)  | `#C82372` | Reserved as secondary action color; used sparingly                             |
| Coolgrey-100 (paper)   | `#F5F7F9` | Cover slide background                                                         |
| Paper-soft             | `#F8F9FA` | Content slide background, dashboard background                                 |

**Data-viz convention** (mirrors the platform product UI): periwinkle for
charts derived from **overall** scores, yellow for **per-task** charts,
pastel category stripes on task chips (see the stripe table under
Per-task analytics). The report screen follows the product UI, not the
marketing site: plain sans-serif page title on the light page
background, blue primary buttons, white cards with `coolgrey-200`
hairline borders. The dark-ink + yellow banner treatment is reserved
for marketing surfaces and is not used on the report screen.

**Fonts.** Display: **JetBrains Mono** (Regular, Medium, SemiBold, Bold,
ExtraBold) self-hosted from `assets/fonts/`. Body: **Inter** loaded from
Google Fonts CDN. Both render in the PPTX and the dashboard. PptxGenJS
embeds the font name in the slide XML, so PowerPoint and Keynote use the
local copy if installed and fall back to a system mono / sans otherwise.

**Signature motifs.** Every chapter title in the PPTX ends with the
underscore mark (`_`) inline in the title text, sky-blue on light
surfaces, yellow on dark. The cover slide's only decoration is a large
sky-blue circle bleeding off the top-right corner. The brand explicitly
rejects: em dashes for emphasis, vertical pipe separators, left-rail
accent bars on cards, decorative drop shadows, and centered text. The
codebase follows all of these rules. (The colored left stripes on task
chips are not a violation of the left-rail rule: they replicate the
product UI's own task-chip component, which carries category color that
way.)

---

## Development notes

- **No build step.** Edit `index.html` in any editor and refresh.
- **No automated tests.** The app has been hand-verified against several
  real Codility exports and a synthetic harness that exercises every
  code path (every insight branch, every badge variant, the rotation /
  no-rotation paths, the integrity / no-integrity paths, long testnames,
  zero-attempt months, etc.). A future port should add unit tests around
  `computeStats`, `buildTakeaway`, `buildFindings`,
  `buildRecommendations`, and `buildInsightsData`.
- **Vendored libraries.** All three live in `assets/vendor/`. Versions
  are intentionally pinned to whatever was checked in. Bumping them
  requires a visual smoke test of the dashboard and the deck.
- **Async error handling.** Download buttons go through
  `safeRun(fn, btn)` which catches rejections and surfaces them as
  alerts. Without this wrapper, unhandled rejections in async onclick
  handlers die silently in the console.
- **Voice rules baked into copy generation.** The four copy builders
  (`buildTakeaway`, `buildFindings`, `buildRecommendations`,
  `buildInsightsData`) all follow the same rules: sentence case, third
  person, no contractions, no em dashes, no emoji, no exclamation
  marks. Every claim references one of the five surfaced KPIs
  (Invitations, Attempts, Median Score, Passing Rate, Median Time) so
  the copy never introduces a metric the reader hasn't already seen.
- **Recommendation narrative.** `buildRecommendations` only shows
  "Keep the assessment as it is" when there are zero actionable
  recommendations. It is never mixed with "Improve X" / "Set Y"
  recommendations on the same slide.
- **Pre-handoff checklist**:
  - JS parses cleanly (single inline script block, ~167k chars)
  - CSS braces balanced
  - All asset references resolve
  - No dead code (no `placeBanding`, no `bandIdx`, no orphan tokens)
  - No em dashes in user-facing strings
  - Integration contract (`_allRows`, `_lastCfg`, `buildReport`,
    `downloadPPTX`) verified accessible from `window.*`
