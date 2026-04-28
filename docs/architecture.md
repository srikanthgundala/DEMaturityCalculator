# Architecture Overview — DEMaturityCalculator

This document describes the technical architecture, component responsibilities, data flow, and key design decisions of the DEMaturityCalculator Excel add-in.

---

## Table of Contents

1. [High-Level Architecture](#1-high-level-architecture)
2. [Technology Stack](#2-technology-stack)
3. [Component Map](#3-component-map)
4. [Data Flow](#4-data-flow)
5. [Scoring Engine](#5-scoring-engine)
6. [Office.js Integration](#6-officejs-integration)
7. [Manifest & Ribbon Extension](#7-manifest--ribbon-extension)
8. [Build Pipeline](#8-build-pipeline)
9. [Dialog System](#9-dialog-system)
10. [Output Sheet Generation](#10-output-sheet-generation)
11. [Error Handling](#11-error-handling)
12. [Key Design Decisions](#12-key-design-decisions)

---

## 1. High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        Microsoft Excel                          │
│                                                                 │
│  ┌──────────────────┐          ┌──────────────────────────────┐ │
│  │  Form1 / Table1  │  reads   │     Task Pane WebView        │ │
│  │  (Survey Input)  │◄─────────│  taskpane.html / .js / .css  │ │
│  └──────────────────┘          │                              │ │
│                                │  ┌────────────────────────┐  │ │
│  ┌──────────────────┐  writes  │  │   Scoring Engine       │  │ │
│  │ DEMaturitySummary│◄─────────│  │ CalculateLevelScore()  │  │ │
│  │   (SummaryTable) │          │  └────────────────────────┘  │ │
│  └──────────────────┘          │                              │ │
│                                │  ┌────────────────────────┐  │ │
│  ┌──────────────────┐  writes  │  │  Sheet Writer          │  │ │
│  │ Per-Project Sheets│◄────────│  │ addLevelTable()        │  │ │
│  │  (Q&A Tables)    │          │  └────────────────────────┘  │ │
│  └──────────────────┘          └──────────────────────────────┘ │
│                                                                 │
│         ↑ Office JavaScript API (office.js v1.1)               │
└─────────────────────────────────────────────────────────────────┘
          ↑
    https://localhost:3000  (Webpack dev server)
    or hosted HTTPS server  (production)
```

The add-in runs entirely **client-side** inside a WebView embedded in Excel. There is no back-end server — all computation happens in the browser runtime and all data is read from and written to the active Excel workbook via the Office JavaScript API.

---

## 2. Technology Stack

| Layer | Technology | Version |
|---|---|---|
| Add-in host | Microsoft Excel (Microsoft 365) | — |
| Office API | Office.js | 1.1 (CDN) |
| UI framework | Office UI Fabric Core (CSS only) | 9.6.1 |
| Language | JavaScript (ES5 target, ES2015+ source) | — |
| Transpiler | Babel (`@babel/preset-env`) | 7.x |
| Bundler | Webpack | 4.x |
| Dev server | webpack-dev-server | 3.x |
| TLS certificates | office-addin-dev-certs | 1.x |
| Dialog library | officejs.dialogs | 1.0.9 |
| Sideloading | office-addin-debugging | 3.x |
| Linting | ESLint + office-addins config | — |
| Formatting | Prettier + office-addin-prettier-config | — |

---

## 3. Component Map

```
src/
├── taskpane/
│   ├── taskpane.js      ← Primary logic: read → calculate → write
│   ├── taskpane.html    ← Task pane shell (header, button, sideload warning)
│   ├── taskpane.css     ← Styles (Office UI Fabric layout conventions)
│   ├── dialogs.js       ← officejs.dialogs v1.0.9 (Wait, MessageBox, Alert, etc.)
│   └── dialogs.html     ← Hidden iframe rendered by Office.js dialog API
└── commands/
    ├── commands.js      ← Add-in commands function file (ribbon button handler)
    └── commands.html    ← Host page for commands.js (loaded by manifest FunctionFile)
```

### `taskpane.js` — Core Logic

Contains the single exported function `run()` plus five internal helper functions that are defined inside the `Excel.run()` callback scope:

| Function | Purpose |
|---|---|
| `run()` | Entry point — orchestrates the entire read → score → write pipeline |
| `getQuestions()` | Extracts question text from the header row by level index list |
| `CalculateLevelScore()` | Scores a single project row for one maturity level |
| `getJsDateFromExcel()` | Converts an Excel serial date number to `MM-DD-YYYY` string |
| `addLevelTable()` | Creates a Question/Response table on a project sheet |
| `applyFilter()` | Applies a column filter to hide perfect-score rows |
| `applyLevelQuestionTopRowProperties()` | Styles the section header row (bold, white text, blue fill) |

### `commands.js` — Ribbon Command Handler

Registers a no-op `action()` function that is wired to the ribbon button via the manifest. The real UI entry point is the task pane itself; `commands.js` fulfils the Office Add-in requirement for a `FunctionFile`.

### `dialogs.js` — Dialog Library

Third-party library (`officejs.dialogs` v1.0.9) providing:
- `Wait.Show()` / `Wait.CloseDialogAsync()` — indeterminate spinner overlay
- `MessageBox.Show()` / `MessageBox.CloseDialogAsync()` — modal message box with configurable buttons and icons
- `Alert`, `InputBox`, `Form`, `Progress`, `PrintPreview` — additional dialog helpers (not used by the current add-in)

---

## 4. Data Flow

### Step-by-step pipeline

```
1. User clicks "Calculate Maturity"
   └─▶ run() is called

2. Wait spinner displayed

3. Excel.run() context opened
   ├─▶ Load worksheet list (context.workbook.worksheets)
   ├─▶ Get Form1 worksheet
   ├─▶ Get Table1 from Form1
   ├─▶ Load header row (questions)
   └─▶ Load data body range (project responses)
          ↓  await context.sync()

4. Delete all non-Form1, non-system sheets
          ↓  await context.sync()

5. Parse header row → level question dictionaries
   ├─▶ level1Questions  { index → question text }  (58 items)
   ├─▶ level2Questions  { index → question text }  (15 items)
   └─▶ level3Questions  { index → question text }  (12 items)

6. Create DEMaturitySummary sheet + SummaryTable

7. For each project row i:
   ├─▶ CalculateLevelScore(row, scores, level1Indexes, weight=70) → level1Details
   ├─▶ CalculateLevelScore(row, scores, level2Indexes, weight=20) → level2Details
   ├─▶ CalculateLevelScore(row, scores, level3Indexes, weight=10) → level3Details
   ├─▶ finalScore = sum of three weightedLevelScores
   ├─▶ maturity = M1 / M2 / M3 based on finalScore
   ├─▶ Append row to SummaryTable
   ├─▶ Highlight Level 1 score cell red if < 70
   ├─▶ sheets.add(projectSheetName)
   └─▶ Push project data to projectSheetsData[]
          ↓  await context.sync()  (creates all sheets)

8. Auto-fit DEMaturitySummary columns/rows

9. For each entry in projectSheetsData[]:
   ├─▶ Write summary card (A1:B10)
   ├─▶ Style maturity cell (yellow fill, maroon bold, double border)
   ├─▶ Highlight Level 1 score red if < 70
   ├─▶ addLevelTable(level1Questions, ...) → Level1Table{ID}
   ├─▶ addLevelTable(level2Questions, ...) → Level2Table{ID}
   ├─▶ addLevelTable(level3Questions, ...) → Level3Table{ID}
   ├─▶ Auto-fit columns/rows
   └─▶ await context.sync()

10. For each project (second pass — requires sheets to exist):
    ├─▶ applyFilter(Level1Table response column)
    ├─▶ applyFilter(Level2Table response column)
    ├─▶ applyFilter(Level3Table response column)
    ├─▶ Add PROJECT hyperlink in SummaryTable (B row → project sheet A2)
    ├─▶ Add MATURITY hyperlink in SummaryTable (J row → project sheet B10)
    └─▶ Add "back to summary" hyperlink on project sheet (B11 → DEMaturitySummary!A1)

11. Activate DEMaturitySummary sheet
          ↓  await context.sync()

12. Close Wait spinner
```

### Why two passes over project data?

Hyperlinks and filters on named tables can only be set after all sheets and tables have been created and synced to the Excel context. The second pass (step 10) therefore operates on already-created sheets and tables, which allows the use of `getItem()` lookups without needing another `context.sync()`.

---

## 5. Scoring Engine

### Question index arrays

The column indexes that map to each maturity level are hard-coded constants at the top of the `Excel.run()` callback:

```js
var level1MaturityIndexes = [7,8,9,10,11,12,13,14,18,19,20,21,22,23,24,
  27,28,29,30,31,33,34,35,36,37,38,44,45,46,47,48,49,50,51,52,53,54,
  59,60,61,62,63,64,67,68,72,73,74,75,76,79,80,81,82,83,84,85,86];  // 58 questions

var level2MaturityIndexes = [15,16,25,17,32,39,55,56,57,65,66,69,77,87,88]; // 15 questions

var level3MaturityIndexes = [26,40,41,42,43,58,70,71,78,89,90,91];           // 12 questions
```

### `CalculateLevelScore()` algorithm

```
Input:
  projectRow      — array of cell values for one survey row
  responseScores  — { "NA": 10, "Always": 10, "Yes": 10,
                      "Frequently": 7, "Sometimes": 4,
                      "Never": 0, "No": 0, "Don't Know": 0 }
  levelIndexes    — array of column indexes for this level
  weightage       — 70, 20, or 10

Algorithm:
  levelScore = 0
  maxLevelScore = levelIndexes.length × 10
  levelFailures = {}   // column indexes of non-perfect answers

  for each index in levelIndexes:
    response = projectRow[index]
    if response exists and is a known key:
      score = responseScores[response]
      if score ≠ 10: levelFailures[index] = index
      levelScore += score
    else:
      levelScore += 0   // blank / unknown treated as 0

  levelPercentage   = (levelScore × 100) / maxLevelScore
  weightedLevelScore = ((levelPercentage × weightage) / 100).toFixed(2)

Output: { levelFailures, unWeightedLevelScore, unWeightedLevelPercentage, weightedLevelScore }
```

### Maturity decision

```js
finalScore = level1.weightedLevelScore + level2.weightedLevelScore + level3.weightedLevelScore;

if      (finalScore <= 70)                maturity = "M1";
else if (finalScore > 70 && finalScore <= 90)  maturity = "M2";
else if (finalScore > 90)                maturity = "M3";
```

### Worked example

Suppose a project has:
- Level 1: 40 out of 58 questions answered "Always" (×10) and 18 answered "Sometimes" (×4)
  - Raw = (40 × 10) + (18 × 4) = 472  
  - Max = 58 × 10 = 580  
  - Percentage = 472/580 × 100 ≈ 81.38%  
  - Weighted = 81.38 × 70 / 100 ≈ **56.97**
- Level 2: all 15 answered "Always"
  - Weighted = 100 × 20 / 100 = **20.00**
- Level 3: all 12 answered "Always"
  - Weighted = 100 × 10 / 100 = **10.00**
- Final = 56.97 + 20 + 10 = **86.97** → **M2**
- Level 1 weighted (56.97) < 70 → highlighted **red**

---

## 6. Office.js Integration

### Initialization

```js
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});
```

`Office.onReady()` is the modern (promise-based) initialization hook. The task pane UI is hidden until this callback fires with `HostType.Excel`, preventing interaction before the Office.js runtime is ready.

### `Excel.run()` context pattern

All worksheet operations happen inside an `Excel.run(async context => { ... })` callback. This pattern:
1. Opens a tracked request batch.
2. Queues operations (`.load()`, `.values =`, `.add()`, etc.) without immediately executing them.
3. Flushes the batch to the Excel host on each `await context.sync()`.

Multiple `context.sync()` calls are needed because later operations depend on data returned by earlier ones (e.g., you must sync after `sheets.load()` before iterating `sheets.items`).

### ExcelApi requirement checks

Before calling APIs that require ExcelApi 1.2 (autofit), the code checks:

```js
if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
  sheet.getUsedRange().format.autofitColumns();
}
```

This ensures graceful degradation on older Excel versions.

### Platform detection

```js
var platform = Office.context.platform;
// ...
if (platform == "OfficeOnline") {
  context.sync();          // fire-and-forget for Office Online
} else {
  await context.sync();    // await on desktop
}
```

Excel Online has different concurrency constraints; the add-in uses a non-awaited sync for that platform.

---

## 7. Manifest & Ribbon Extension

`manifest.xml` defines:

| Element | Value |
|---|---|
| Add-in type | `TaskPaneApp` |
| Add-in ID (GUID) | `1c9376c1-ee2d-4875-82ee-f1bf4eed4374` |
| Provider | Neudesic |
| Host | `Workbook` (Excel) |
| Permissions | `ReadWriteDocument` |
| Task pane URL | `https://localhost:3000/taskpane.html` |
| Commands file URL | `https://localhost:3000/commands.html` |

### Ribbon extension

A `PrimaryCommandSurface` extension point adds a button group to Excel's **Home** tab:

```
Home Tab
└── DE Group (CommandsGroup)
    └── [Button] Show DEMaturityCalculator
            └── Action: ShowTaskpane → taskpane.html
```

The button uses three icon sizes (`icon-16.png`, `icon-32.png`, `icon-80.png`) for different display densities.

---

## 8. Build Pipeline

Webpack bundles two entry points:

| Entry | Source | Output |
|---|---|---|
| `taskpane` | `src/taskpane/taskpane.js` | `dist/taskpane.js` |
| `commands` | `src/commands/commands.js` | `dist/commands.js` |
| `polyfill` | `@babel/polyfill` | `dist/polyfill.js` |

### Plugins

| Plugin | Role |
|---|---|
| `CleanWebpackPlugin` | Wipes `dist/` before each build |
| `HtmlWebpackPlugin` (×2) | Injects script tags into `taskpane.html` and `commands.html` |
| `CopyWebpackPlugin` | Copies `taskpane.css`, `dialogs.js`, `dialogs.html`, and `assets/` to `dist/` |

### Loaders

| Loader | File types | Role |
|---|---|---|
| `babel-loader` | `*.js` | Transpiles ES2015+ to ES5 (`@babel/preset-env`) |
| `html-loader` | `*.html` | Handles HTML imports |
| `file-loader` | `*.png`, `*.jpg`, `*.gif` | Copies image assets with content-hash filenames |

### Dev server

- Port: `3000` (configurable via `package.json → config.dev-server-port`)
- HTTPS: auto-configured via `office-addin-dev-certs.getHttpsServerOptions()`
- CORS: `Access-Control-Allow-Origin: *` header set for all responses

---

## 9. Dialog System

The add-in uses the `officejs.dialogs` library (`dialogs.js` + `dialogs.html`) for UI feedback. These dialogs are rendered inside a separate `Office.context.ui.displayDialogAsync()` window (an Office API feature) which appears as a centred modal overlay on top of Excel.

### Dialogs used in this add-in

| Dialog | When used |
|---|---|
| `Wait.Show(message, false, callback)` | Shown immediately when **Calculate Maturity** is clicked |
| `Wait.CloseDialogAsync(callback)` | Closed after processing completes or on error |
| `MessageBox.Show(message, title, buttons, icon, ...)` | Shown on caught errors (e.g. `ItemNotFound`) |

### Error message mapping

| `error.code` | User-facing message |
|---|---|
| `ItemNotFound` | "Please run add-in on DE Survey responses" |
| Any other | "There is an error in processing project responses. Please try again or later" |

---

## 10. Output Sheet Generation

### DEMaturitySummary sheet creation

```
sheets.add("DEMaturitySummary")
  └── tables.add("A1:J1", hasHeaders=true)  → name: "SummaryTable"
       └── getHeaderRowRange().values = [["ID","PROJECT","REVIEW DATE","EMAIL",
                                          "RESOURCE COUNT","LEVEL 1 SCORE","LEVEL 2 SCORE",
                                          "LEVEL 3 SCORE","FINAL SCORE","MATURITY"]]
```

### Project sheet naming

```js
var projectSheetName = projectName.substring(0, 25).replace(/[^a-zA-Z0-9]/g, '') + "_" + responseId;
```

- Maximum 25 characters from the project name (before stripping).
- All non-alphanumeric characters stripped (handles spaces, hyphens, special chars).
- Response ID appended after `_` to guarantee uniqueness.

### Sheet deletion policy

Every time **Calculate Maturity** is run, all existing sheets except `Form1` and system sheets (names starting with `_`) are deleted before new sheets are created. This ensures a clean, reproducible output on each run.

### Hyperlink objects

Hyperlinks are created using the `Excel.Range.hyperlink` property with a `documentReference` target (intra-workbook navigation), **not** an external URL:

```js
// DEMaturitySummary → Project sheet
{ textToDisplay: projectName, documentReference: "SheetName!A2" }

// Project sheet → DEMaturitySummary
{ textToDisplay: "Click here to go to DEMaturitySummary sheet", documentReference: "DEMaturitySummary!A1" }
```

---

## 11. Error Handling

The `Excel.run()` call is wrapped in a `.catch()` handler:

```js
}).catch(function(error) {
  console.error(error);
  var errormessage = "There is an error...";
  if (error.code == "ItemNotFound") {
    errormessage = "Please run add-in on DE Survey responses";
  }
  Wait.CloseDialogAsync(function() {
    MessageBox.Show(errormessage, "Error", MessageBoxButtons.OkOnly,
                    MessageBoxIcons.Error, false, null,
                    function(btn) { MessageBox.CloseDialogAsync(function(){}); },
                    false);
  });
});
```

The outer `run()` function also has a `try/catch` block that logs unexpected errors to the console.

Extended error logging is turned on globally:

```js
OfficeExtension.config.extendedErrorLogging = true;
```

---

## 12. Key Design Decisions

| Decision | Rationale |
|---|---|
| **All logic in one file** | Keeps the add-in self-contained and easy to deploy; complexity is manageable at current scale |
| **Hard-coded column indexes** | Survey form structure is fixed; avoids fragile column-name lookups that could break on minor form edits |
| **Two-pass sheet writing** | Hyperlinks and filters depend on completed table objects, which requires a `context.sync()` boundary |
| **Sheet deletion on each run** | Guarantees output is always fresh and consistent; avoids stale data from previous runs |
| **`toFixed(2)` on weighted scores** | Prevents floating-point noise in displayed scores while keeping them numeric for comparison |
| **`parseFloat()` after `toFixed()`** | `toFixed()` returns a string; `parseFloat()` converts back to number for the final sum |
| **Platform check for `context.sync`** | Office Online requires different async handling to avoid race conditions in the browser runtime |

---

*Next: [API / Function Reference →](api-reference.md)*
