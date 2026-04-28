# Architecture Overview

This document describes the technical architecture of the **DEMaturityCalculator** Excel Add-in.

---

## Technology Stack

| Layer | Technology |
|---|---|
| Add-in platform | Microsoft Office Add-ins (Office.js) |
| Host application | Microsoft Excel (desktop & online) |
| Language | JavaScript (ES6+) |
| Bundler | webpack 4 |
| UI framework | Office UI Fabric (Fluent UI v1) |
| Dialog library | officejs.dialogs |
| Dev tooling | office-addin-* CLI tools |

---

## High-Level Architecture

```
Excel Workbook
│
│  Form1 / Table1  (survey responses)
│
└─► DEMaturityCalculator Add-in (task pane)
        │
        │  Office.js (ExcelApi 1.2+)
        │
        ├─► Read responses & questions from Table1
        ├─► Calculate Level 1 / 2 / 3 scores
        ├─► Create DEMaturitySummary sheet
        └─► Create per-project detail sheets
```

The add-in runs entirely inside Excel using the Office JavaScript API. There are **no external network calls** — all computation happens client-side inside the browser/Excel host runtime.

---

## Source Layout

```
src/
├── commands/
│   ├── commands.html   # Minimal HTML page loaded for ribbon commands
│   └── commands.js     # Ribbon command entry point (currently empty placeholder)
└── taskpane/
    ├── taskpane.html   # Task pane shell (ribbon button → shows this pane)
    ├── taskpane.css    # Styles for the task pane
    ├── taskpane.js     # ALL core business logic (see below)
    ├── dialogs.html    # Host page for Office dialog pop-ups
    └── dialogs.js      # officejs.dialogs wrapper (Wait / MessageBox helpers)
```

### taskpane.js — Core Logic

`taskpane.js` is the single source of truth for all maturity calculations. It is structured as follows:

```
Office.onReady()
└── run()                          ← triggered by "Calculate Maturity" button
    ├── Read Form1 / Table1
    ├── Delete previous result sheets
    ├── CalculateLevelScore()      ← scores one level for one project row
    ├── getQuestions()             ← extracts question text by index
    ├── getJsDateFromExcel()       ← converts Excel date serial to MM-DD-YYYY
    ├── Create DEMaturitySummary sheet
    ├── addLevelTable()            ← writes question/response table to a sheet
    ├── applyFilter()              ← highlights non-passing responses
    └── applyLevelQuestionTopRowProperties()
```

---

## Data Flow

```
1. User clicks "Calculate Maturity"
        │
2. run() fetches Form1 worksheet
        │
3. Header row  → question dictionaries (level1Questions, level2Questions, level3Questions)
   Data body rows → one object per project response
        │
4. For each project response:
   ┌─────────────────────────────────────────────────────────┐
   │  CalculateLevelScore(row, scores, levelIndexes, weight) │
   │  → levelFailures, unWeightedLevelScore, weightedScore   │
   └─────────────────────────────────────────────────────────┘
        │
5. finalScore = L1.weighted + L2.weighted + L3.weighted
   maturity   = M1 / M2 / M3 based on thresholds
        │
6. Append row to DEMaturitySummary table
   Add project detail sheet with three level tables
        │
7. Apply hyperlinks between summary and detail sheets
   Activate DEMaturitySummary sheet
```

---

## Maturity Level Indexes

The survey table has 91+ columns. Specific column indexes are pre-assigned to each maturity level:

| Level | Indexes (illustrative) | Weight |
|---|---|---|
| Level 1 | 7–14, 18–24, 27–38, 44–54, 59–68, 72–86 | 70% |
| Level 2 | 15–17, 25, 32, 39, 55–57, 65–66, 69, 77, 87–88 | 20% |
| Level 3 | 26, 40–43, 58, 70–71, 78, 89–91 | 10% |

The exact index arrays are defined in `taskpane.js` as `level1MaturityIndexes`, `level2MaturityIndexes`, and `level3MaturityIndexes`.

---

## Output Sheets

| Sheet | Contents |
|---|---|
| `DEMaturitySummary` | Summary table: ID, Project, Review Date, Email, Resource Count, L1/L2/L3 scores, Final Score, Maturity |
| `<ProjectName>_<ID>` | Per-project: summary data block + three level tables (Level 1, Level 2, Level 3) |

---

## Build Pipeline

```
webpack
├── entry: taskpane.js  →  taskpane bundle
├── entry: commands.js  →  commands bundle
├── HtmlWebpackPlugin   →  taskpane.html / commands.html
└── CopyWebpackPlugin   →  taskpane.css, dialogs.js, dialogs.html, assets/
```

Dev server runs at `https://localhost:3000` with HTTPS (required by Office Add-in platform).

---

## Security Notes

- The add-in requires **ReadWriteDocument** permission in `manifest.xml` because it creates new worksheets and modifies workbook content.
- No data is transmitted outside of Excel; all processing is local.
- The `officejs.dialogs` library is loaded from `node_modules` at build time (no CDN dependency at runtime).
