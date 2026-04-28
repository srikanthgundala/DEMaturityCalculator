# DEMaturityCalculator — Technical Documentation

> **Provider:** Neudesic  
> **Add-in ID:** `1c9376c1-ee2d-4875-82ee-f1bf4eed4374`  
> **Version:** 1.0.0.0  
> **Host:** Microsoft Excel (Workbook)  
> **Permissions:** `ReadWriteDocument`

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Tech Stack](#2-tech-stack)
3. [Architecture](#3-architecture)
4. [Excel Workbook Structure](#4-excel-workbook-structure)
   - 4.1 [Input: Form1 / Table1](#41-input-form1--table1)
   - 4.2 [Output: DEMaturitySummary Sheet](#42-output-dematuritysummary-sheet)
   - 4.3 [Output: Per-Project Sheets](#43-output-per-project-sheets)
5. [Maturity Score Model](#5-maturity-score-model)
   - 5.1 [Response Scoring](#51-response-scoring)
   - 5.2 [Maturity Levels and Question Sets](#52-maturity-levels-and-question-sets)
   - 5.3 [Weighted Score Calculation](#53-weighted-score-calculation)
   - 5.4 [Maturity Thresholds](#54-maturity-thresholds)
6. [Data Flow](#6-data-flow)
7. [Key Functions Reference](#7-key-functions-reference)
   - 7.1 [getQuestions](#71-getquestions)
   - 7.2 [CalculateLevelScore](#72-calculatelevelscore)
   - 7.3 [addLevelTable](#73-addleveltable)
   - 7.4 [applyFilter](#74-applyfilter)
   - 7.5 [getJsDateFromExcel](#75-getjsdatefromexcel)
   - 7.6 [applyLevelQuestionTopRowProperties](#76-applylevelquestiontoprowproperties)
8. [UI & Dialog System](#8-ui--dialog-system)
9. [Setup & Build](#9-setup--build)
10. [Manifest Configuration](#10-manifest-configuration)
11. [Error Handling](#11-error-handling)
12. [How to Extend — Adding New Maturity Levels](#12-how-to-extend--adding-new-maturity-levels)

---

## 1. Project Overview

**DEMaturityCalculator** is a Microsoft Excel Task Pane Add-in that automates the assessment of a software project's *Data Engineering (DE) Maturity Level*. Users load a workbook containing DE survey responses (one row per project), then click **"Calculate Maturity"** in the add-in panel. The add-in reads every response, computes weighted scores across four practice levels, classifies each project into a maturity band (M1–M4), and writes richly formatted output sheets — a summary dashboard and one detailed drill-down sheet per project — directly back into the same workbook.

### Goals

| Goal | Description |
|------|-------------|
| **Automate scoring** | Eliminate manual score calculation from survey spreadsheets |
| **Standardise levels** | Enforce a consistent four-level maturity framework across all projects |
| **Enable drill-down** | Give each project its own sheet showing question-by-question results with failures highlighted |
| **Surface at-a-glance status** | Provide a single summary sheet with hyperlinks to every project detail sheet |

---

## 2. Tech Stack

| Layer | Technology | Version / Notes |
|-------|-----------|-----------------|
| **Add-in runtime** | Office.js (`office.js`) | v1.1, CDN-hosted from `appsforoffice.microsoft.com` |
| **Excel API** | Excel JavaScript API | Requires `ExcelApi 1.2`+ for auto-fit; `ReadWriteDocument` permission |
| **Language** | JavaScript (ES5 target via Babel) | `@babel/preset-env`, no TypeScript at runtime |
| **Type checking** | TypeScript (`allowJs: true`) | `tsconfig.json` — types only, no compiled output from `.js` files |
| **Bundler** | Webpack 4 | `webpack.config.js` — produces `taskpane.js`, `commands.js`, `taskpane.html` |
| **Transpilation** | Babel (`babel-loader`) | Converts modern JS to ES5 for broad Office client compatibility |
| **Dev server** | `webpack-dev-server` | Port 3000, HTTPS via `office-addin-dev-certs` |
| **Dialogs** | `officejs.dialogs` v1.0.9 | Provides `Wait`, `MessageBox`, `Alert` etc. via `dialogs.js` |
| **UI framework** | Office UI Fabric Core 9.6.1 | CSS only, CDN link in `taskpane.html` |
| **Dev toolchain** | `office-addin-debugging`, `office-addin-manifest`, `office-addin-lint` | Microsoft Office Add-in CLI tools |

---

## 3. Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Excel Desktop / Excel Online              │
│                                                             │
│  ┌──────────────┐    Office.js API     ┌──────────────────┐ │
│  │  Task Pane   │◄────────────────────►│  Excel Workbook  │ │
│  │  (Browser)   │                      │                  │ │
│  │              │   Excel.run(context) │  ┌────────────┐  │ │
│  │  taskpane.   │──────────────────────►  │   Form1    │  │ │
│  │  html / js   │                      │  │  Table1    │  │ │
│  │              │   context.sync()     │  └────────────┘  │ │
│  │  ┌────────┐  │◄─────────────────── │                  │ │
│  │  │ "Calc  │  │  Writes output:     │  ┌────────────┐  │ │
│  │  │Maturity│  │──────────────────── ►  │DEMaturity  │  │ │
│  │  │" button│  │                      │  │  Summary   │  │ │
│  │  └────────┘  │                      │  └────────────┘  │ │
│  └──────────────┘                      │  ┌────────────┐  │ │
│         ▲                              │  │ ProjectA_1 │  │ │
│  ┌──────┴──────┐                       │  │ ProjectB_2 │  │ │
│  │ manifest.   │                       │  │  …         │  │ │
│  │ xml         │                       │  └────────────┘  │ │
│  │ (sideloaded │                       └──────────────────┘ │
│  │  or deployed│                                            │
│  └─────────────┘                                            │
└─────────────────────────────────────────────────────────────┘
```

### Component Breakdown

| File | Role |
|------|------|
| `manifest.xml` | Declares the add-in to Excel: ID, provider, icons, ribbon button, task pane URL |
| `src/taskpane/taskpane.html` | Task pane HTML shell — Fabric CSS, logo, "Calculate Maturity" button |
| `src/taskpane/taskpane.js` | **Core logic** — all score calculation, sheet creation, formatting |
| `src/taskpane/taskpane.css` | Task pane styling (Fabric-compatible layout) |
| `src/taskpane/dialogs.js` | `officejs.dialogs` library — `Wait`, `MessageBox`, `Alert`, etc. |
| `src/taskpane/dialogs.html` | HTML template used by the dialogs library inside Office dialog iframes |
| `src/commands/commands.js` | Stub command handler — wires `action()` for add-in command events |
| `webpack.config.js` | Build configuration — entries, Babel, HTML plugin, CopyPlugin, dev-server |
| `assets/` | PNG icons (16×16, 32×32, 80×80) and the Neudesic logo |

### Execution Flow Summary

1. **Sideload / install**: `manifest.xml` registers the add-in. Excel injects a "Show DEMaturityCalculator" button in the **Home** tab ribbon under the *DE Group* group.
2. **Task pane opens**: `taskpane.html` loads, `Office.onReady()` fires, the task pane body is revealed, and the button's `onclick` is bound to `run()`.
3. **User clicks "Calculate Maturity"**: `run()` is called, a `Wait` spinner is displayed.
4. **Single `Excel.run` batch**: All reads, sheet manipulation, writes, and formatting happen inside one `Excel.run(async context => { … })` call, with strategic `await context.sync()` checkpoints to flush the command queue.
5. **Output is written**: `DEMaturitySummary` and per-project sheets are created. `DEMaturitySummary` is activated on completion.

---

## 4. Excel Workbook Structure

### 4.1 Input: Form1 / Table1

The add-in expects the active workbook to contain a worksheet named exactly **`Form1`** with an Excel Table named exactly **`Table1`**.

**Sheet name:** `Form1`  
**Table name:** `Table1`  
**Structure:** Header row + one data row per project response.

#### Key Column Indices (0-based)

These are the fixed column positions the add-in reads from each data row:

| Index | Field | Notes |
|-------|-------|-------|
| 0 | **Response ID** | Unique identifier for the survey submission |
| 2 | **Review Date** | Stored as an Excel serial date number; converted to `MM-DD-YYYY` by `getJsDateFromExcel` |
| 3 | **Respondent Email** | Used in the project summary |
| 5 | **Project Name** | Used to name the per-project sheet and as hyperlink text |
| 6 | **Resource Count** | Team/resource count for the project |
| 7–91 | **Survey Questions** | Maturity question responses (see §5.2 for level-to-index mapping) |

#### Valid Response Values

Each survey question cell must contain exactly one of the following text values:

| Response Text | Score |
|--------------|-------|
| `NA` | 10 |
| `Always` | 10 |
| `Yes` | 10 |
| `Frequently` | 7 |
| `Sometimes` | 4 |
| `Never` | 0 |
| `No` | 0 |
| `Don't Know` | 0 |

Any response value not in the table above is treated as **0** (the score is simply not added).

---

### 4.2 Output: DEMaturitySummary Sheet

**Sheet name:** `DEMaturitySummary`  
**Table name:** `SummaryTable`  
**Range:** Starting at `A1`, expanding with one row per project.

#### Columns

| Col | Header | Content |
|-----|--------|---------|
| A | ID | Response ID from Table1 |
| B | PROJECT | Project name (hyperlink → project detail sheet) |
| C | REVIEW DATE | Formatted date (`MM-DD-YYYY`) |
| D | EMAIL | Respondent email |
| E | RESOURCE COUNT | Team size |
| F | LEVEL 1 SCORE | Weighted L1 score (0–55) |
| G | LEVEL 2 SCORE | Weighted L2 score (0–20) |
| H | LEVEL 3 SCORE | Weighted L3 score (0–15) |
| I | LEVEL 4 SCORE | Weighted L4 score (0–10) |
| J | FINAL SCORE | Sum of all weighted scores (0–100) |
| K | MATURITY | M1 / M2 / M3 / M4 (hyperlink → project detail sheet cell B11) |

#### Formatting Rules

- **Level 1 Score cell** (col F): Red font (`#FF0000`) if the weighted L1 score is **< 40**.
- **Maturity cell** (col K): Yellow fill (`#FFFFE0`), bold, dark-red font (`#800000`), acts as a hyperlink.
- All columns and rows are auto-fitted after population.

---

### 4.3 Output: Per-Project Sheets

One sheet is created per project response.

**Sheet name format:** `{sanitisedProjectName}_{responseId}`  
- `sanitisedProjectName` = first 25 characters of the project name with all non-alphanumeric characters removed.
- Example: `"Acme Data Platform"` → `AcmeDataPlatform_42`

#### Layout

| Row(s) | Col | Content |
|--------|-----|---------|
| 1–11 | A:B | Summary table (label / value pairs) |
| 12 | B | Hyperlink back to `DEMaturitySummary!A1` |
| 13 | *(blank)* | — |
| 14 | B:C | **"Level 1 Questions"** header (merged, blue fill `#154CC5`, white bold text) |
| 15 … | B:C | Level 1 question/response table (`Level1Table{id}`) |
| *(gap)* | | 3 blank rows between level tables |
| next | B:C | **"Level 2 Questions"** header |
| … | B:C | Level 2 question/response table (`Level2Table{id}`) |
| next | B:C | **"Level 3 Questions"** header |
| … | B:C | Level 3 question/response table (`Level3Table{id}`) |
| next | B:C | **"Level 4 Questions"** header |
| … | B:C | Level 4 question/response table (`Level4Table{id}`) |

#### Summary Sub-Table (A1:B11)

| Row | A (label) | B (value) |
|-----|-----------|-----------|
| 1 | ID | `{responseId}` |
| 2 | PROJECT | `{projectName}` |
| 3 | REVIEW DATE | `MM-DD-YYYY` |
| 4 | EMAIL | `{email}` |
| 5 | RESOURCE COUNT | `{resourceCount}` |
| 6 | LEVEL 1 SCORE | `{weightedL1}` |
| 7 | LEVEL 2 SCORE | `{weightedL2}` |
| 8 | LEVEL 3 SCORE | `{weightedL3}` |
| 9 | LEVEL 4 SCORE | `{weightedL4}` |
| 10 | FINAL SCORE | `{finalScore}` |
| 11 | MATURITY | `M1`/`M2`/`M3`/`M4` |

- **B6 (Level 1 Score)**: Red font if < 40.
- **B11 (Maturity)**: Yellow fill, bold, dark-red font, double borders on all edges.
- **B12**: Hyperlink labelled *"Click here to go to DEMaturitySummary sheet"*, brown font (`#7C3606`), yellow fill (`#E1D70F`).

#### Level Question Tables

Each level has a named Excel Table with two columns:

| Column | Content |
|--------|---------|
| **Question** | Full question text from Table1 header row |
| **Response** | The respondent's answer |

- Table name pattern: `Level{N}Table{responseId}` (e.g., `Level1Table7`)
- **Row font**: Black (`#000000`) by default.
- **Failure rows**: Rows where the response is not a perfect-score answer are highlighted in red (`#FF0000`).
- **Filter applied**: Each table's `Response` column has an auto-filter pre-set to show only non-passing responses: `Frequently`, `Sometimes`, `Never`, `No`, `Don't Know`. This surfaces all weaknesses immediately on open.

---

## 5. Maturity Score Model

### 5.1 Response Scoring

Every survey question is answered with a text choice. The add-in maps each text choice to a numeric score:

```javascript
responseScores["NA"]         = 10;  // Not applicable — treated as full score
responseScores["Always"]     = 10;
responseScores["Yes"]        = 10;
responseScores["Frequently"] =  7;
responseScores["Sometimes"]  =  4;
responseScores["Never"]      =  0;
responseScores["No"]         =  0;
responseScores["Don't Know"] =  0;
```

A response is flagged as a **failure** (`levelFailures`) if it scores less than the maximum of **10**.

---

### 5.2 Maturity Levels and Question Sets

Questions are identified by their **0-based column index** in Table1's data rows.

#### Level 1 — Core / Foundational Practices (55 questions)

> **Weight: 55% of final score**

```
Column indices: 7, 8, 9, 10, 11, 12, 13, 14,
                18, 19, 20, 21, 22, 23, 24,
                27, 28, 29, 30, 31,
                33, 34, 35, 36, 37, 38,
                44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54,
                59, 60, 61, 62, 63, 64,
                67, 68,
                72, 73, 74, 75, 76,
                79, 80, 81, 82, 83
```

These questions cover the **foundational data engineering practices** every mature team should have in place.

---

#### Level 2 — Intermediate Practices (18 questions)

> **Weight: 20% of final score**

```
Column indices: 15, 16, 17, 25, 32, 39,
                55, 56, 57,
                65, 66, 69, 77,
                84, 85, 86, 87, 88
```

> **Note (re-ranking):** Columns 84, 85, and 86 were originally classified as Level 1 questions in an earlier version of the survey and have been re-ranked to Level 2 to reflect their intermediate nature.

---

#### Level 3 — Advanced Practices (9 questions)

> **Weight: 15% of final score**

```
Column indices: 26, 40, 41, 42, 43,
                58, 70, 71, 78
```

> **Note (re-ranking):** Columns 89, 90, and 91 were previously included in Level 3 but have been promoted to Level 4 to represent cutting-edge capabilities.

---

#### Level 4 — Cutting-Edge / Optimising Practices (3 questions)

> **Weight: 10% of final score**

```
Column indices: 89, 90, 91
```

These questions represent the most advanced, optimising practices. A project must perform strongly here (weighted score ≥ 8 out of 10) in addition to a high overall score to achieve M4 maturity.

---

### 5.3 Weighted Score Calculation

For each level, `CalculateLevelScore` computes:

```
maxLevelScore        = questionCount × 10
levelScore           = Σ responseScores[response] for each question in the level
levelPercentage      = (levelScore / maxLevelScore) × 100
weightedLevelScore   = (levelPercentage × weightage) / 100   [rounded to 2 decimal places]
```

#### Maximum Weighted Scores per Level

| Level | Questions | Max Raw Score | Weight | Max Weighted Score |
|-------|-----------|---------------|--------|--------------------|
| L1 | 55 | 550 | 55% | **55.00** |
| L2 | 18 | 180 | 20% | **20.00** |
| L3 | 9 | 90 | 15% | **15.00** |
| L4 | 3 | 30 | 10% | **10.00** |
| **Total** | **85** | — | **100%** | **100.00** |

#### Final Score

```
finalScore = weightedL1 + weightedL2 + weightedL3 + weightedL4
```

The `finalScore` ranges from **0** to **100**.

#### Return Value of `CalculateLevelScore`

```javascript
{
  levelFailures: { [columnIndex]: columnIndex, … },  // indices of non-perfect responses
  unWeightedLevelScore: number,                        // raw sum of all response scores
  unWeightedLevelPercentage: number,                   // percentage of max raw score
  weightedLevelScore: number                           // the contribution to finalScore
}
```

---

### 5.4 Maturity Thresholds

| Maturity Band | Condition | Description |
|--------------|-----------|-------------|
| **M1** | `finalScore ≤ 40` | Foundational practices are largely missing or inconsistent |
| **M2** | `40 < finalScore ≤ 65` | Core practices are established; intermediate practices are emerging |
| **M3** | `65 < finalScore ≤ 85` | Advanced practices adopted; cutting-edge practices not yet optimised |
| **M4** | `finalScore > 85` **AND** `weightedL4 ≥ 8` | Cutting-edge and optimising practices are actively in use |

> **M4 edge case:** If `finalScore > 85` but `weightedL4 < 8`, the project is classified as **M3** — i.e., a high overall score is insufficient for M4 without demonstrating strong performance on the Level 4 (cutting-edge) questions. This prevents teams from gaming M4 by scoring well only on foundational questions.

The L1 weighted score is also independently monitored: if it falls below **40** (≈ 73% of the maximum L1 contribution of 55), the L1 score cell is coloured red in both the summary and detail sheets as a warning that foundational practices are weak.

---

## 6. Data Flow

```
Excel Workbook
└── Form1!Table1
        │
        │  (1) Read header row → question text dictionary
        │  (2) Read data body  → one array per project response
        ▼
┌───────────────────────────────────────────────────────┐
│  run() — inside Excel.run(async context => { … })     │
│                                                       │
│  [Phase 0]  context.sync() — flush read queue         │
│                                                       │
│  [Phase 1]  Delete existing output sheets             │
│             (all except "Form1" and "_*" sheets)      │
│             context.sync()                            │
│                                                       │
│  [Phase 2]  Build question dictionaries               │
│             getQuestions(header, L1indexes) → L1Q     │
│             getQuestions(header, L2indexes) → L2Q     │
│             getQuestions(header, L3indexes) → L3Q     │
│             getQuestions(header, L4indexes) → L4Q     │
│                                                       │
│  [Phase 3]  Create DEMaturitySummary sheet + table    │
│                                                       │
│  [Phase 4]  For each project response row:            │
│    ├─ CalculateLevelScore(row, scores, L1Q, 55)       │
│    ├─ CalculateLevelScore(row, scores, L2Q, 20)       │
│    ├─ CalculateLevelScore(row, scores, L3Q, 15)       │
│    ├─ CalculateLevelScore(row, scores, L4Q, 10)       │
│    ├─ Compute finalScore & maturity band              │
│    ├─ Append row to SummaryTable                      │
│    ├─ Highlight L1 score cell if < 40                 │
│    └─ Queue new project sheet creation                │
│                                                       │
│  [Phase 5]  Auto-fit DEMaturitySummary                │
│             context.sync()                            │
│                                                       │
│  [Phase 6]  For each project sheet:                   │
│    ├─ Write A1:B11 summary block                      │
│    ├─ Format maturity cell (B11)                      │
│    ├─ addLevelTable(L1Q, …) → Level 1 table           │
│    ├─ addLevelTable(L2Q, …) → Level 2 table           │
│    ├─ addLevelTable(L3Q, …) → Level 3 table           │
│    ├─ addLevelTable(L4Q, …) → Level 4 table           │
│    ├─ Auto-fit sheet                                  │
│    └─ context.sync()                                  │
│                                                       │
│  [Phase 7]  For each project sheet:                   │
│    ├─ applyFilter(Level1Table.Response column)        │
│    ├─ applyFilter(Level2Table.Response column)        │
│    ├─ applyFilter(Level3Table.Response column)        │
│    ├─ applyFilter(Level4Table.Response column)        │
│    ├─ Add hyperlink: Summary PROJECT cell → sheet     │
│    ├─ Add hyperlink: Summary MATURITY cell → sheet    │
│    └─ Add hyperlink: Sheet B12 → DEMaturitySummary    │
│                                                       │
│  [Phase 8]  Activate DEMaturitySummary                │
│             await context.sync()                      │
└───────────────────────────────────────────────────────┘
        │
        ▼
Excel Workbook (updated)
├── Form1          (unchanged)
├── DEMaturitySummary  (new — summary dashboard)
├── ProjectA_1         (new — detail sheet)
├── ProjectB_2         (new — detail sheet)
└── …
```

> **Important:** Sheets are deleted and recreated on every run. Any existing non-`Form1`, non-`_`-prefixed sheets are removed before new output sheets are written. The `_` prefix convention can be used to protect utility sheets from deletion.

---

## 7. Key Functions Reference

All functions below are defined **inside** the `Excel.run(async context => { … })` callback in `taskpane.js` and therefore have closure access to `context`, `projectResponses`, `responseScores`, `RESPONSE_MAX_SCORE`, `sheets`, etc.

---

### 7.1 `getQuestions`

```javascript
function getQuestions(questionsRow, levelIndexes): Object
```

**Purpose:** Builds a dictionary mapping column index → question text for the given set of level indexes.

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `questionsRow` | `Array<Array<string>>` | The `.values` array from the header row range of Table1 |
| `levelIndexes` | `number[]` | Array of 0-based column indices for the target maturity level |

**Returns:** `Object` — `{ [columnIndex: number]: string }` mapping each column index to its question text.

**Example:**

```javascript
// questionsRow[0][7] = "Do you use version control for all data pipeline code?"
var level1Questions = getQuestions(questions.values, level1MaturityIndexes);
// → { 7: "Do you use version control for all data pipeline code?", 8: "…", … }
```

---

### 7.2 `CalculateLevelScore`

```javascript
function CalculateLevelScore(projectRow, responseScores, levelIndexes, weightage): Object
```

**Purpose:** Computes the raw score, percentage, and weighted contribution for one maturity level for one project.

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `projectRow` | `Array<any>` | One data row from Table1 (the full row array) |
| `responseScores` | `Object` | Map of response text → numeric score |
| `levelIndexes` | `number[]` | Column indices for this level |
| `weightage` | `number` | Level weight percentage (55, 20, 15, or 10) |

**Returns:**

```typescript
{
  levelFailures: { [columnIndex: number]: number },  // non-perfect responses
  unWeightedLevelScore: number,                        // raw score sum
  unWeightedLevelPercentage: number,                   // (rawScore / maxRawScore) × 100
  weightedLevelScore: number                           // contribution to final 0–100 score
}
```

**Algorithm:**

```
maxLevelScore = levelIndexes.length × 10
for each index in levelIndexes:
    response = projectRow[index]
    score    = responseScores[response] ?? 0
    if score < 10: levelFailures[index] = index
    levelScore += score

levelPercentage   = (levelScore / maxLevelScore) × 100
weightedLevelScore = (levelPercentage × weightage) / 100  [toFixed(2)]
```

---

### 7.3 `addLevelTable`

```javascript
function addLevelTable(questions, levelFailures, projectSheet, levelIndexes, level, projectResponse, rowIndex, context): number
```

**Purpose:** Creates a named Excel Table on a project detail sheet for one maturity level, populates it with question/response pairs, and highlights failures in red.

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `questions` | `Object` | `{ [index]: questionText }` from `getQuestions` |
| `levelFailures` | `Object` | `{ [index]: index }` from `CalculateLevelScore` |
| `projectSheet` | `Excel.Worksheet` | The target project detail worksheet |
| `levelIndexes` | `number[]` | Ordered column indices for this level |
| `level` | `string` | Level identifier string: `"Level1"`, `"Level2"`, `"Level3"`, or `"Level4"` |
| `projectResponse` | `Array<any>` | The full data row for this project |
| `rowIndex` | `number` | Starting row number (1-based) on the worksheet |
| `context` | `Excel.RequestContext` | The active Excel run context |

**Returns:** `number` — The updated `rowIndex` after all question rows have been added (incremented by one per question).

**Table naming convention:** `{level}Table{projectResponse[0]}`  
Example: `Level1Table42` (for response ID 42).

**Failure highlighting:** Any row whose column index is present in `levelFailures` has its font colour set to `#FF0000` (red).

---

### 7.4 `applyFilter`

```javascript
function applyFilter(filter): void
```

**Purpose:** Applies a pre-configured Excel column filter to a level table's `Response` column, showing only non-passing responses so developers immediately see where the gaps are.

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `filter` | `Excel.Filter` | The filter object from `tableColumn.filter` |

**Filter configuration:**

```javascript
filter.apply({
  filterOn: Excel.FilterOn.values,
  values: ["Frequently", "Sometimes", "Never", "No", "Don't Know"]
});
```

Only responses with these values are shown when the sheet opens. `Always`, `Yes`, and `NA` (perfect scores) are hidden.

---

### 7.5 `getJsDateFromExcel`

```javascript
function getJsDateFromExcel(dateValue): string
```

**Purpose:** Converts an Excel serial date number to a human-readable `MM-DD-YYYY` string.

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `dateValue` | `number` | Excel date serial number (days since 1 January 1900, with the Lotus 1-2-3 leap-year bug offset) |

**Returns:** `string` — Formatted date in `MM-DD-YYYY` format (zero-padded month and day).

**Algorithm:**

```javascript
// Excel epoch starts 1 Jan 1900; subtract 25569 (= 25567 + 2) days to get Unix epoch (1 Jan 1970)
// Multiply by 86400 seconds/day × 1000 ms/s to get JavaScript millisecond timestamp
var jsDate = new Date((dateValue - 25569) * 86400 * 1000);
```

> The constant `25567 + 2 = 25569` accounts for the well-known Excel epoch offset (Excel incorrectly treats 1900 as a leap year, so the adjustment is 25567 + 1 for the epoch difference + 1 for the phantom Feb 29 1900).

---

### 7.6 `applyLevelQuestionTopRowProperties`

```javascript
function applyLevelQuestionTopRowProperties(levelRange): void
```

**Purpose:** Applies consistent header-row formatting to the "Level N Questions" title cell above each level table.

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `levelRange` | `Excel.Range` | A two-cell merged range spanning columns B:C at the title row |

**Applied formatting:**

| Property | Value |
|----------|-------|
| `merge` | `true` (merges B:C) |
| `font.bold` | `true` |
| `font.color` | `#FFFFFF` (white) |
| `horizontalAlignment` | `"Center"` |
| `fill.color` | `#154CC5` (Neudesic blue) |

---

## 8. UI & Dialog System

The task pane UI is intentionally minimal:

- **Header**: Neudesic logo + "Welcome" heading.
- **Body**: A single hero button labelled **"Calculate Maturity"**, wired to `run()`.

The dialog system (`dialogs.js`, `officejs.dialogs` v1.0.9) provides modal popups rendered inside Office iframe dialogs:

| Global Object | Type | Used In |
|---------------|------|---------|
| `Wait` | Spinner dialog | Displayed while `run()` processes; dismissed on completion or error |
| `MessageBox` | OK/Error dialog | Shown when `Excel.run` throws (e.g., `ItemNotFound`) |
| `Alert` | Simple alert | Available but commented out in error path |

**Wait lifecycle:**

```javascript
Wait.Show("Processing Project Responses", false, callback);  // open
// … processing …
Wait.CloseDialogAsync(callback);                            // close
```

---

## 9. Setup & Build

### Prerequisites

| Requirement | Details |
|-------------|---------|
| **Node.js** | v12+ recommended (LTS) |
| **npm** | v6+ |
| **Microsoft 365** | Required to test/run the add-in (Excel desktop or Excel Online) |
| **HTTPS certificate** | Generated automatically by `office-addin-dev-certs` on first `npm start` |

### Install Dependencies

```bash
npm install
```

### Available Scripts

| Script | Command | Description |
|--------|---------|-------------|
| `start` | `office-addin-debugging start manifest.xml` | Start dev server + sideload in Excel (default: auto-detect desktop/web) |
| `start:desktop` | `office-addin-debugging start manifest.xml desktop` | Force Excel desktop |
| `start:web` | `office-addin-debugging start manifest.xml web` | Force Excel Online |
| `stop` | `office-addin-debugging stop manifest.xml` | Stop the dev server and remove sideloading |
| `build` | `webpack -p --mode production --https false` | Production bundle (minified) to `dist/` |
| `build:dev` | `webpack --mode development --https false` | Development bundle (source maps) to `dist/` |
| `watch` | `webpack --mode development --watch` | Incremental rebuild on file changes |
| `dev-server` | `webpack-dev-server --mode development` | Webpack dev server only (no sideloading) |
| `validate` | `office-addin-manifest validate manifest.xml` | Validate `manifest.xml` against the Office schema |
| `lint` | `office-addin-lint check` | Run ESLint with Office Add-in ruleset |
| `lint:fix` | `office-addin-lint fix` | Auto-fix lint issues |

### Build Output

Webpack emits to the `dist/` directory:

| Output File | Source |
|-------------|--------|
| `taskpane.html` | `src/taskpane/taskpane.html` (via HtmlWebpackPlugin) |
| `taskpane.js` | `src/taskpane/taskpane.js` + Babel |
| `taskpane.css` | `src/taskpane/taskpane.css` (via CopyWebpackPlugin) |
| `commands.html` | `src/commands/commands.html` (via HtmlWebpackPlugin) |
| `commands.js` | `src/commands/commands.js` + Babel |
| `dialogs.js` | `src/taskpane/dialogs.js` (via CopyWebpackPlugin) |
| `dialogs.html` | `src/taskpane/dialogs.html` (via CopyWebpackPlugin) |
| `assets/` | All files from `assets/` (via CopyWebpackPlugin) |

### Development Workflow

```bash
# 1. Install dependencies
npm install

# 2. Start add-in (launches Excel and sideloads automatically)
npm start

# 3. Open a workbook with "Form1" / "Table1" survey data
# 4. Click Home → DE Group → Show DEMaturityCalculator
# 5. Click "Calculate Maturity"
# 6. Edit code; webpack rebuilds automatically
# 7. Stop the dev server when done
npm stop
```

---

## 10. Manifest Configuration

The `manifest.xml` registers the add-in with Excel.

| Field | Value |
|-------|-------|
| Add-in ID | `1c9376c1-ee2d-4875-82ee-f1bf4eed4374` |
| Provider | Neudesic |
| Display Name | DEMaturityCalculator |
| Description | Add-in to calculate DE maturity level of project |
| Target Host | `Workbook` (Excel) |
| Permissions | `ReadWriteDocument` |
| Default Locale | `en-US` |
| Dev server port | `3000` (HTTPS) |
| Task pane URL | `https://localhost:3000/taskpane.html` |
| Commands URL | `https://localhost:3000/commands.html` |
| Ribbon location | Home tab → "DE Group" group → "Show DEMaturityCalculator" button |
| Support URL | `https://www.neudesic.com` |

> **Production deployment:** Replace all `https://localhost:3000/…` URLs in `manifest.xml` with the URL of the hosted static web server (e.g., Azure Static Web Apps, SharePoint, or a CDN endpoint).

---

## 11. Error Handling

The `Excel.run` promise chain has a `.catch` handler:

```javascript
.catch(function(error) {
  console.error(error);
  var errormessage = "There is an error in processing project responses. Please try again or later";
  if (error.code == "ItemNotFound") {
    errormessage = "Please run add-in on DE Survey responses";
  }
  Wait.CloseDialogAsync(function() {
    MessageBox.Show(errormessage, "Error", MessageBoxButtons.OkOnly, MessageBoxIcons.Error, …);
  });
});
```

| Error Condition | Detected By | User Message |
|-----------------|-------------|-------------|
| `Form1` sheet missing | `error.code === "ItemNotFound"` | "Please run add-in on DE Survey responses" |
| `Table1` not found | `error.code === "ItemNotFound"` | "Please run add-in on DE Survey responses" |
| Any other Excel API error | generic catch | "There is an error in processing project responses. Please try again or later" |

After the `Excel.run` completes (success or failure), if the `Wait` spinner is still displayed it is closed:

```javascript
if (Wait.Displayed()) {
  Wait.CloseDialogAsync(function () { });
}
```

Extended error logging is always enabled:

```javascript
OfficeExtension.config.extendedErrorLogging = true;
```

---

## 12. How to Extend — Adding New Maturity Levels

The scoring system is data-driven through index arrays and a weight parameter, making extensions straightforward.

### Step 1 — Define the New Level's Column Indices

Add a new array in `taskpane.js` inside `Excel.run`, following the same convention:

```javascript
// Level 5 (5% weight): Visionary / research practices
var level5MaturityIndexes = [92, 93, 94];
```

> Adjust the weight of other levels so all weights still sum to 100.

### Step 2 — Adjust Existing Weights (if needed)

```javascript
var level1CalculationDetails = CalculateLevelScore(row, responseScores, level1MaturityIndexes, 50); // was 55
var level2CalculationDetails = CalculateLevelScore(row, responseScores, level2MaturityIndexes, 20);
var level3CalculationDetails = CalculateLevelScore(row, responseScores, level3MaturityIndexes, 15);
var level4CalculationDetails = CalculateLevelScore(row, responseScores, level4MaturityIndexes,  8); // was 10
var level5CalculationDetails = CalculateLevelScore(row, responseScores, level5MaturityIndexes,  7); // new
```

### Step 3 — Extract Questions

```javascript
var level5Questions = getQuestions(headerValues, level5MaturityIndexes);
```

### Step 4 — Add to finalScore

```javascript
var finalScore = level1CalculationDetails.weightedLevelScore
               + level2CalculationDetails.weightedLevelScore
               + level3CalculationDetails.weightedLevelScore
               + level4CalculationDetails.weightedLevelScore
               + level5CalculationDetails.weightedLevelScore; // add new level
```

### Step 5 — Add to DEMaturitySummary Table

Update the `summaryTable.getHeaderRowRange().values` array and the `summaryTable.rows.add` call to include the new level score column.

### Step 6 — Add to Project Sheet Summary Block

Expand `summaryData` in the per-project sheet section and update the `A1:B{N}` range accordingly.

### Step 7 — Add the Level Table to Project Sheets

```javascript
projectSheet.getRange("B" + (rowIndex + 3)).values = [["Level 5 Questions"]];
var level5Range = projectSheet.getRange("B" + (rowIndex + 3) + ":" + "C" + (rowIndex + 3));
applyLevelQuestionTopRowProperties(level5Range);

rowIndex = addLevelTable(
  level5Questions,
  projectSheetData.level5CalculationDetails.levelFailures,
  projectSheet,
  level5MaturityIndexes,
  "Level5",
  projectResponses.values[i],
  rowIndex + 4,
  context
);
```

### Step 8 — Apply a Filter on the New Level Table

```javascript
var level5filter = projectSheet.tables
  .getItem("Level5Table" + projectResponses.values[i][0])
  .columns.getItem("Response").filter;
applyFilter(level5filter);
```

### Step 9 — Update Maturity Thresholds (if needed)

Revise the `if/else if` maturity decision block to account for the new level and any updated score ranges.

### Step 10 — Update Table1 Columns

Add the new survey questions to `Form1 / Table1` at columns 92–94 (or whatever indices you chose), ensuring every respondent answers them.

---

### Summary of Key Constants

| Constant | Value | Location |
|----------|-------|----------|
| `RESPONSE_MAX_SCORE` | `10` | `taskpane.js` line ~37 |
| `LEVEL_MIN_THRESHOLD` | `40` | `taskpane.js` line ~39 — L1 warning threshold |
| L4 M4 minimum | `8` | `taskpane.js` line ~119 — `level4CalculationDetails.weightedLevelScore >= 8` |
| Excel epoch offset | `25567 + 2` | `getJsDateFromExcel` |

---

*Documentation generated from source code at `src/taskpane/taskpane.js`, `manifest.xml`, `package.json`, `webpack.config.js`, and `tsconfig.json`.*
