# DE Maturity Calculator — Excel Add-in

> **Provider:** Neudesic &nbsp;|&nbsp; **Version:** 1.0.0.0 &nbsp;|&nbsp; **Platform:** Microsoft Excel (Office.js TaskPane)  
> **Manifest ID:** `1c9376c1-ee2d-4875-82ee-f1bf4eed4374`

---

## Table of Contents

1. [Overview](#1-overview)
2. [Architecture](#2-architecture)
3. [Repository Structure](#3-repository-structure)
4. [Maturity Scoring Algorithm](#4-maturity-scoring-algorithm)
   - 4.1 [Response Score Mapping](#41-response-score-mapping)
   - 4.2 [Level Score Calculation](#42-level-score-calculation)
   - 4.3 [Final Score Calculation](#43-final-score-calculation)
   - 4.4 [Maturity Classification (M1–M4)](#44-maturity-classification-m1m4)
5. [Maturity Levels & Question Indexes](#5-maturity-levels--question-indexes)
   - 5.1 [Level Weights](#51-level-weights)
   - 5.2 [Level 1 Question Indexes](#52-level-1-question-indexes)
   - 5.3 [Level 2 Question Indexes](#53-level-2-question-indexes)
   - 5.4 [Level 3 Question Indexes](#54-level-3-question-indexes)
   - 5.5 [Level 4 Question Indexes](#55-level-4-question-indexes)
6. [Excel Workbook Structure](#6-excel-workbook-structure)
   - 6.1 [Input Sheet — Form1 / Table1](#61-input-sheet--form1--table1)
   - 6.2 [Output Sheet — DEMaturitySummary](#62-output-sheet--dematuritysummary)
   - 6.3 [Output Sheets — Per-Project Detail Sheets](#63-output-sheets--per-project-detail-sheets)
7. [UI & Add-in Interaction Flow](#7-ui--add-in-interaction-flow)
8. [Key Functions Reference](#8-key-functions-reference)
9. [Error Handling](#9-error-handling)
10. [Developer Setup](#10-developer-setup)
    - 10.1 [Prerequisites](#101-prerequisites)
    - 10.2 [Install Dependencies](#102-install-dependencies)
    - 10.3 [Build & Run](#103-build--run)
    - 10.4 [Build Configuration](#104-build-configuration)
11. [Manifest Configuration](#11-manifest-configuration)
12. [Contributing](#12-contributing)
13. [License](#13-license)

---

## 1. Overview

The **DE Maturity Calculator** is a Microsoft Excel Office Add-in that automates the calculation and reporting of **Data Engineering (DE) maturity levels** for one or more projects. It reads survey responses collected in a structured Excel table (`Form1 / Table1`), evaluates each response against a four-level maturity model, computes weighted scores, classifies the overall maturity as **M1 through M4**, and generates formatted output worksheets — a single summary sheet and one detail sheet per project.

Key capabilities:
- Processes **85 survey questions** grouped across four progressive maturity levels.
- Assigns weighted scores to each level (L1 = 60 %, L2 = 20 %, L3 = 10 %, L4 = 10 %).
- Produces a **DEMaturitySummary** sheet with hyperlinked navigation to each project.
- Generates per-project detail sheets showing question/response tables, with non-passing responses highlighted in red.
- Auto-filters each question table to show only failing/partial responses.

---

## 2. Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Excel Workbook                           │
│                                                             │
│  ┌──────────────────┐         ┌───────────────────────┐    │
│  │  Form1 / Table1  │ ──────► │  Office.js Add-in     │    │
│  │  (Survey input)  │  reads  │  (taskpane.js)        │    │
│  └──────────────────┘         │                       │    │
│                                │  1. Parse responses   │    │
│                                │  2. Score each level  │    │
│                                │  3. Compute final     │    │
│                                │  4. Classify M1–M4    │    │
│                                │  5. Write output      │    │
│                                └────────┬──────────────┘    │
│                                         │ writes             │
│              ┌──────────────────────────┼──────────────┐    │
│              ▼                          ▼              ▼    │
│   ┌──────────────────┐    ┌─────────────────────────────┐  │
│   │ DEMaturitySummary│    │ {ProjectName}_{ID} sheets   │  │
│   │ (1 row / project)│    │ (1 sheet per project)       │  │
│   └──────────────────┘    └─────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘

Tech stack
──────────
  • JavaScript (ES6+) compiled with Babel (@babel/preset-env → ES5)
  • Office.js API  (ExcelApi 1.1 minimum, 1.2 for autofit)
  • officejs.dialogs  (Wait spinner, MessageBox)
  • Webpack 4  (bundler, dev-server on port 3000)
  • Office Add-in manifest (OfficeApp v1.1 TaskPaneApp)
```

The add-in runs entirely client-side inside Excel. There are **no external API calls** and **no server-side components**. All computation happens in the user's browser/Excel runtime via the Office JavaScript API.

---

## 3. Repository Structure

```
DEMaturityCalculator/
├── manifest.xml                    # Office Add-in manifest
├── package.json                    # npm scripts & dependencies
├── webpack.config.js               # Webpack build configuration
├── tsconfig.json                   # TypeScript/JS compiler options
├── .eslintrc.json                  # ESLint rules
├── assets/
│   ├── neudesilogo.png             # Neudesic logo displayed in task pane
│   ├── icon-16.png                 # Add-in ribbon icon (16 px)
│   ├── icon-32.png                 # Add-in ribbon icon (32 px)
│   └── icon-80.png                 # Add-in ribbon icon (80 px)
└── src/
    ├── taskpane/
    │   ├── taskpane.js             # ★ Core add-in logic (scoring engine)
    │   ├── taskpane.html           # Task pane UI shell
    │   ├── taskpane.css            # Task pane styles (Office Fabric)
    │   ├── dialogs.js              # officejs.dialogs library (Wait, MessageBox)
    │   └── dialogs.html            # Dialog host page
    └── commands/
        ├── commands.js             # Ribbon command handler
        └── commands.html           # Commands host page
```

> **Entry point:** `src/taskpane/taskpane.js` — contains the entire scoring engine and Excel output generation logic inside the exported `run()` async function.

---

## 4. Maturity Scoring Algorithm

### 4.1 Response Score Mapping

Survey respondents answer each question with one of the following options. The add-in maps each answer to a numeric score out of **10**:

| Answer | Score | Notes |
|--------|------:|-------|
| `Always` | **10** | Full credit |
| `Yes` | **10** | Full credit |
| `NA` | **10** | Not-Applicable treated as passing |
| `Frequently` | **7** | Partial credit — flagged as failure |
| `Sometimes` | **4** | Partial credit — flagged as failure |
| `No` | **0** | No credit — flagged as failure |
| `Never` | **0** | No credit — flagged as failure |
| `Don't Know` | **0** | No credit — flagged as failure |
| *(blank / unknown)* | **0** | Treated as no response |

> **Failure flag:** Any response with a score **below 10** is recorded in `levelFailures` and highlighted in **red** on the project detail sheet.

---

### 4.2 Level Score Calculation

For each maturity level the function `CalculateLevelScore()` computes:

```
1.  Raw level score   = Σ responseScore  for all questions in that level

2.  Max level score   = (number of questions in level) × 10

3.  Level percentage  = (rawLevelScore / maxLevelScore) × 100

4.  Weighted score    = (levelPercentage × levelWeightage) / 100
                        [rounded to 2 decimal places]
```

**Example — Level 1 (58 questions, 60 % weight):**

```
Raw score  = 520  (hypothetical)
Max score  = 58 × 10 = 580
Percentage = (520 / 580) × 100  ≈ 89.66 %
Weighted   = (89.66 × 60) / 100 ≈ 53.79
```

---

### 4.3 Final Score Calculation

```
finalScore = weightedL1 + weightedL2 + weightedL3 + weightedL4
```

The maximum achievable `finalScore` is **100** (all questions answered `Always`/`Yes`/`NA`).

| Level | Weight | Max Contribution |
|-------|-------:|-----------------:|
| Level 1 | 60 % | 60.00 pts |
| Level 2 | 20 % | 20.00 pts |
| Level 3 | 10 % | 10.00 pts |
| Level 4 | 10 % | 10.00 pts |
| **Total** | **100 %** | **100.00 pts** |

---

### 4.4 Maturity Classification (M1–M4)

After computing `finalScore`, the project is classified into one of four maturity grades:

```
finalScore ≤ 60            →  M1  (Foundational)
60 < finalScore ≤ 75       →  M2  (Developing)
75 < finalScore ≤ 90       →  M3  (Defined)
finalScore > 90            →  M4  (Optimized)
```

| Grade | Score Range | Description |
|-------|-------------|-------------|
| **M1** | 0 – 60 | Foundational — basic DE practices not yet consistently applied |
| **M2** | 61 – 75 | Developing — core practices in place; advanced practices emerging |
| **M3** | 76 – 90 | Defined — robust practices; near-complete adoption of advanced topics |
| **M4** | 91 – 100 | Optimized — full adoption including the most advanced DE disciplines |

> **Level 1 red-highlight behaviour:** The code defines `LEVEL_MIN_THRESHOLD = 70` and checks `weightedLevelScore < LEVEL_MIN_THRESHOLD`. Because the maximum possible Level 1 weighted score is `60.0` (100 % × 60 % weight), this condition is **always true** — Level 1 is therefore highlighted in **red** on every project in both the summary sheet and the project detail sheet. This appears to be an intended warning that Level 1 scores should be interpreted relative to the 60-point ceiling, not as a score-gated threshold. If the intent is to flag only projects that score below 70 % completion of Level 1, the effective score threshold would be `0.70 × 60 = 42.0`.

---

## 5. Maturity Levels & Question Indexes

Question indexes refer to **zero-based column positions** within the survey table header row (`Table1` in `Form1`). Columns 0–6 are metadata fields (ID, timestamps, project name, email, resource count, etc.). Survey questions start at index 7 and run through index 91 — **85 consecutive positions with no gaps** — giving **85 evaluated questions** in total. Each index appears in exactly one level (no overlapping indexes).

### 5.1 Level Weights

| Level | Weightage | # Questions | Max Raw Score | Max Weighted Score |
|-------|----------:|------------:|--------------:|-------------------:|
| Level 1 | **60 %** | 58 | 580 | 60.00 |
| Level 2 | **20 %** | 15 | 150 | 20.00 |
| Level 3 | **10 %** | 5 | 50 | 10.00 |
| Level 4 | **10 %** | 7 | 70 | 10.00 |
| **Totals** | **100 %** | **85** | **850** | **100.00** |

---

### 5.2 Level 1 Question Indexes

**58 questions** covering foundational DE practices.

```javascript
level1MaturityIndexes = [
  7, 8, 9, 10, 11, 12, 13, 14,
  18, 19, 20, 21, 22, 23, 24,
  27, 28, 29, 30, 31,
  33, 34, 35, 36, 37, 38,
  44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54,
  59, 60, 61, 62, 63, 64,
  67, 68,
  72, 73, 74, 75, 76,
  79, 80, 81, 82, 83, 84, 85, 86
]
```

---

### 5.3 Level 2 Question Indexes

**15 questions** covering intermediate DE practices.

```javascript
level2MaturityIndexes = [
  15, 16, 17, 25, 32, 39,
  55, 56, 57,
  65, 66, 69,
  77,
  87, 88
]
```

---

### 5.4 Level 3 Question Indexes

**5 questions** covering advanced DE practices.

```javascript
level3MaturityIndexes = [
  26, 40, 41, 42, 43
]
```

---

### 5.5 Level 4 Question Indexes

**7 questions** covering the most advanced / optimised DE practices. These questions were split out from the former Level 3 set when M4 maturity was introduced.

```javascript
level4MaturityIndexes = [
  58, 70, 71, 78, 89, 90, 91
]
```

> **History note:** Prior to the M4 addition, all 12 questions now in Levels 3 and 4 belonged to a single "Level 3" group and Level 1 carried a 70 % weight. The current model splits them to allow independent measurement of the most advanced capabilities, reduces Level 1 weight to 60 %, and allocates the freed 10 % to a new Level 4.

---

## 6. Excel Workbook Structure

### 6.1 Input Sheet — Form1 / Table1

The add-in reads data from a **pre-existing** sheet and table that must already be present in the workbook before the add-in is run.

| Property | Value |
|----------|-------|
| Sheet name | `Form1` |
| Table name | `Table1` |
| Header row | Row 1 — question text at each column index |
| Data rows | One row per project survey submission |

**Required metadata columns (0-based index):**

| Index | Field |
|------:|-------|
| 0 | Response ID |
| 2 | Review Date *(Excel serial date number)* |
| 3 | Reviewer Email |
| 5 | Project Name |
| 6 | Resource Count |

Survey question responses occupy columns **7 through 91**.

> If `Form1` or `Table1` is not found, the add-in surfaces an `ItemNotFound` error with the message *"Please run add-in on DE Survey responses"*.

---

### 6.2 Output Sheet — DEMaturitySummary

A single sheet named **`DEMaturitySummary`** is created (or re-created) each run. It contains a formatted Excel table named `SummaryTable` with **11 columns**:

| Col | Header | Description |
|----:|--------|-------------|
| A | **ID** | Response identifier from column 0 |
| B | **PROJECT** | Project name; hyperlinked to its detail sheet |
| C | **REVIEW DATE** | Formatted as `MM-DD-YYYY` |
| D | **EMAIL** | Reviewer email address |
| E | **RESOURCE COUNT** | Number of resources on the project |
| F | **LEVEL 1 SCORE** | Weighted score (max 60.00); shown in **red** if < 70 |
| G | **LEVEL 2 SCORE** | Weighted score (max 20.00) |
| H | **LEVEL 3 SCORE** | Weighted score (max 10.00) |
| I | **LEVEL 4 SCORE** | Weighted score (max 10.00) |
| J | **FINAL SCORE** | Sum of all weighted scores (max 100.00) |
| K | **MATURITY** | M1 / M2 / M3 / M4; yellow fill, dark-red bold font, hyperlinked to project detail sheet cell B11 |

After all rows are written, columns and rows are auto-fitted (`ExcelApi 1.2`).

---

### 6.3 Output Sheets — Per-Project Detail Sheets

One sheet is created per project response. The sheet name is derived as:

```
{first 25 characters of projectName, alphanumeric only}_{responseId}
```

*Example:* A project named `"E-Commerce Platform v2"` with response ID `42` produces the sheet name `ECommercePlatformv2_42`.

#### Summary Block (A1:B11)

Rows 1–11 of columns A–B contain a key/value summary:

| Row | Column A | Column B |
|----:|----------|----------|
| 1 | ID | *(response ID)* |
| 2 | PROJECT | *(project name)* |
| 3 | REVIEW DATE | *(MM-DD-YYYY)* |
| 4 | EMAIL | *(email)* |
| 5 | RESOURCE COUNT | *(count)* |
| 6 | LEVEL 1 SCORE | *(weighted, red if < 70)* |
| 7 | LEVEL 2 SCORE | *(weighted)* |
| 8 | LEVEL 3 SCORE | *(weighted)* |
| 9 | LEVEL 4 SCORE | *(weighted)* |
| 10 | FINAL SCORE | *(sum)* |
| 11 | MATURITY | **M1/M2/M3/M4** — yellow fill (#FFFFE0), dark-red bold font (#800000), double border |

**Row B12** contains a hyperlink back to the `DEMaturitySummary` sheet (`"Click here to go to DEMaturitySummary sheet"`), styled with bold dark-orange font on yellow background.

#### Level Question Tables

Starting at row 14, four consecutive Excel tables are written (one per level), each separated by three blank rows and preceded by a merged, blue-header label row:

| Table Name Pattern | Level Header Label | Columns |
|--------------------|--------------------|---------|
| `Level1Table{responseId}` | "Level 1 Questions" | Question, Response |
| `Level2Table{responseId}` | "Level 2 Questions" | Question, Response |
| `Level3Table{responseId}` | "Level 3 Questions" | Question, Response |
| `Level4Table{responseId}` | "Level 4 Questions" | Question, Response |

**Level header row styling:** merged B:C cells, white text (#FFFFFF), bold, centered, blue fill (#154CC5).

**Response column filter:** Each table's `Response` column is auto-filtered to show only non-passing values: `Frequently`, `Sometimes`, `Never`, `No`, `Don't Know`.

**Row highlighting:** Any question row where the response is not a full-credit answer is rendered in **red text** (#FF0000).

---

## 7. UI & Add-in Interaction Flow

```
User opens Excel workbook (Form1 with survey data)
           │
           ▼
Opens the DEMaturityCalculator task pane
  (Home tab → DE Group → Show DEMaturityCalculator)
           │
           ▼
Task pane loads (taskpane.html)
  • Neudesic logo + "Welcome" header
  • "Calculate Maturity" button
           │
  [Click "Calculate Maturity"]
           │
           ▼
Wait spinner appears ("Processing Project Responses")
           │
           ▼
Excel.run() context opens
  1. Load Form1 / Table1 header + data rows
  2. Delete all non-Form1, non-system sheets
  3. Build question dictionaries for L1–L4
  4. Create DEMaturitySummary sheet + SummaryTable
  5. For each project row:
       a. Calculate L1, L2, L3, L4 weighted scores
       b. Sum → finalScore
       c. Classify M1–M4
       d. Append row to SummaryTable
       e. Highlight L1 score red if < 70
       f. Create project sheet (name only at this stage)
  6. Sync → flush sheet creation
  7. For each project sheet:
       a. Write summary block A1:B11
       b. Style maturity cell (B11)
       c. Write Level 1–4 question tables
       d. Apply response filters
  8. Add cross-sheet hyperlinks:
       - SummaryTable PROJECT column → project sheet
       - SummaryTable MATURITY column → project sheet B11
       - Project sheet B12 → DEMaturitySummary
  9. Activate DEMaturitySummary
 10. Sync → final flush
           │
           ▼
Wait spinner closes
DEMaturitySummary sheet is active and visible
```

---

## 8. Key Functions Reference

All functions are defined **inside** the `Excel.run()` closure in `src/taskpane/taskpane.js`.

### `run()` — `async function` (exported)
Top-level entry point wired to the **Calculate Maturity** button click. Orchestrates the full scoring and output pipeline.

---

### `CalculateLevelScore(projectRow, responseScores, levelIndexes, weightage)`

Computes the weighted score for a single maturity level.

| Parameter | Type | Description |
|-----------|------|-------------|
| `projectRow` | `Array` | Full row of response values from Table1 |
| `responseScores` | `Object` | Map of answer string → numeric score |
| `levelIndexes` | `Array<number>` | Column indexes belonging to this level |
| `weightage` | `number` | Level weight as a percentage (60, 20, 10, or 10) |

**Returns:**
```javascript
{
  levelFailures: { [index]: index },   // indexes of non-perfect responses
  unWeightedLevelScore: number,        // raw sum of scores
  unWeightedLevelPercentage: number,   // (rawScore / maxScore) × 100
  weightedLevelScore: number           // (percentage × weightage) / 100
}
```

---

### `getQuestions(questionsRow, levelIndexes)`

Builds a dictionary mapping column index → question text for a given level.

| Parameter | Type | Description |
|-----------|------|-------------|
| `questionsRow` | `Array<Array>` | Header row values from `getHeaderRowRange()` |
| `levelIndexes` | `Array<number>` | Column indexes to extract |

**Returns:** `{ [index]: questionText }` (plain object dictionary)

---

### `addLevelTable(questions, levelFailures, projectSheet, levelIndexes, level, projectResponse, rowIndex, context)`

Creates a named Excel table for one maturity level on a project sheet, populates it with question/response pairs, and highlights failing rows in red.

| Parameter | Type | Description |
|-----------|------|-------------|
| `questions` | `Object` | Dictionary from `getQuestions()` |
| `levelFailures` | `Object` | Dictionary of failing indexes from `CalculateLevelScore()` |
| `projectSheet` | `WorksheetObject` | Target worksheet |
| `levelIndexes` | `Array<number>` | Ordered column indexes for this level |
| `level` | `string` | e.g. `"Level1"`, `"Level2"` — used in table name |
| `projectResponse` | `Array` | Full response row |
| `rowIndex` | `number` | Starting Excel row (1-based) |
| `context` | `RequestContext` | Office.js request context |

**Returns:** Updated `rowIndex` after the last written row.

---

### `applyFilter(filter)`

Applies an Excel column filter to show only non-passing responses.

```javascript
filter.apply({
  filterOn: Excel.FilterOn.values,
  values: ["Frequently", "Sometimes", "Never", "No", "Don't Know"]
});
```

---

### `applyLevelQuestionTopRowProperties(levelRange)`

Styles the merged header cell above each level table: blue fill (#154CC5), white bold centered text.

---

### `getJsDateFromExcel(dateValue)`

Converts an Excel serial date number to a human-readable `MM-DD-YYYY` string.

```javascript
// Formula: Excel dates start from 1900-01-00 (serial 1)
// Adjusted offset: 25567 + 2 days for epoch alignment
var reviewDate = new Date((dateValue - (25567 + 2)) * 86400 * 1000);
```

**Returns:** `"MM-DD-YYYY"` formatted string.

---

## 9. Error Handling

| Scenario | Behaviour |
|----------|-----------|
| `Form1` sheet or `Table1` not found | `ItemNotFound` error code caught → MessageBox: *"Please run add-in on DE Survey responses"* |
| Any other runtime error | Generic MessageBox: *"There is an error in processing project responses. Please try again or later"* |
| Wait spinner displayed during error | Automatically closed before the error MessageBox is shown via `Wait.CloseDialogAsync()` |

All errors are also logged to the browser/Excel developer console via `console.error(error)`.  
`OfficeExtension.config.extendedErrorLogging = true` is set on every run to include extended diagnostic information in Office.js error objects.

---

## 10. Developer Setup

### 10.1 Prerequisites

| Tool | Minimum Version | Purpose |
|------|-----------------|---------|
| Node.js | ≥ 14.x | JavaScript runtime |
| npm | ≥ 6.x | Package management |
| Microsoft Excel | Microsoft 365 / Office 2016+ | Hosting the add-in |
| office-addin-dev-certs | (installed via npm) | HTTPS for localhost dev server |

### 10.2 Install Dependencies

```bash
npm install
```

### 10.3 Build & Run

| Command | Description |
|---------|-------------|
| `npm start` | Start the add-in dev server and sideload into Excel (desktop) |
| `npm run start:desktop` | Explicitly start on Excel desktop |
| `npm run start:web` | Start on Excel Online |
| `npm stop` | Stop the development server and remove sideloaded add-in |
| `npm run build` | Production build (minified) to `dist/` |
| `npm run build:dev` | Development build (with source maps) to `dist/` |
| `npm run watch` | Webpack watch mode — rebuilds on file change |
| `npm run validate` | Validate `manifest.xml` against Office Add-in schema |
| `npm run lint` | Run ESLint checks |
| `npm run lint:fix` | Auto-fix ESLint issues |

> **Dev server port:** `3000` (configured in `package.json` → `config.dev-server-port`).  
> The dev server is configured for HTTPS using `office-addin-dev-certs`. On first run you may be prompted to trust a self-signed certificate.

### 10.4 Build Configuration

**Webpack entry points** (`webpack.config.js`):

| Bundle | Source | Output |
|--------|--------|--------|
| `polyfill` | `@babel/polyfill` | Included in `taskpane.html` |
| `taskpane` | `src/taskpane/taskpane.js` | Included in `taskpane.html` |
| `commands` | `src/commands/commands.js` | Included in `commands.html` |

**Copied assets** (unchanged to `dist/`):
- `taskpane.css` → `dist/taskpane.css`
- `dialogs.js` → `dist/dialogs.js`
- `dialogs.html` → `dist/dialogs.html`
- `assets/` → `dist/assets/`

**Babel preset:** `@babel/preset-env` — transpiles ES6+ to ES5 for broad Excel client compatibility.

**TypeScript compiler target** (`tsconfig.json`): `ES5`, with `allowJs: true` so `.js` files are accepted without strict TypeScript.

---

## 11. Manifest Configuration

File: `manifest.xml`

| Property | Value |
|----------|-------|
| Add-in ID | `1c9376c1-ee2d-4875-82ee-f1bf4eed4374` |
| Version | `1.0.0.0` |
| Provider | Nuedesic *(as declared in `manifest.xml`; trading as Neudesic)* |
| Display Name | DEMaturityCalculator |
| Description | Add-in to calculate DE maturity level of project |
| Supported Host | `Workbook` (Excel only) |
| Permissions | `ReadWriteDocument` |
| App Domain | `Neudesic.com` |
| Task Pane Source | `https://localhost:3000/taskpane.html` |
| Ribbon Location | Home tab → **DE Group** |
| Button Label | **Show DEMaturityCalculator** |
| Button Tooltip | Click to Show a DEMaturityCalculator |

**Required API sets:**
- `ExcelApi 1.1` — minimum for table read/write
- `ExcelApi 1.2` — optional for `autofitColumns` / `autofitRows` (checked at runtime with `isSetSupported`)

---

## 12. Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for full contribution guidelines including:
- How to provide better code comments
- How to fix open issues
- How to propose and add new features
- Pull-request checklist and review process

---

## 13. License

Copyright (c) Microsoft Corporation. All rights reserved.  
Licensed under the [MIT License](LICENSE).

---

*Documentation generated from source code in `src/taskpane/taskpane.js`, `manifest.xml`, `package.json`, and `webpack.config.js`.*
