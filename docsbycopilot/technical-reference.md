# DE Maturity Calculator — Technical Reference

> **Add-in Name:** DEMaturityCalculator  
> **Provider:** Neudesic  
> **Version:** 1.0.0.0  
> **Manifest ID:** `1c9376c1-ee2d-4875-82ee-f1bf4eed4374`  
> **Host Application:** Microsoft Excel (Workbook)  
> **Permission Level:** `ReadWriteDocument`  
> **Source:** `src/taskpane/taskpane.js`

---

## Table of Contents

1. [Overview](#1-overview)
2. [Architecture and Prerequisites](#2-architecture-and-prerequisites)
3. [Maturity Level Architecture](#3-maturity-level-architecture)
4. [Scoring Algorithm — `CalculateLevelScore`](#4-scoring-algorithm--calculatelevelscoreprojectrow-responsescores-levelindexes-weightage)
5. [Maturity Ranking Thresholds](#5-maturity-ranking-thresholds-m1--m4)
6. [Output Structure](#6-output-structure)
7. [Key Data Flow](#7-key-data-flow)
8. [Helper Functions Reference](#8-helper-functions-reference)
9. [Error Handling](#9-error-handling)
10. [Build and Development](#10-build-and-development)
11. [Changelog — M4 Addition](#11-changelog--m4-maturity-level-addition)

---

## 1. Overview

The **DE Maturity Calculator** is a Microsoft Excel Office Add-in built with the Office.js JavaScript API. It processes DE (presumably Digital Engineering or Developer Experience) survey responses stored in an Excel workbook and automatically calculates a **maturity score** and **maturity rank** for each project submission.

### What it does

1. Reads raw survey responses from a structured Excel table (`Table1` on the `Form1` sheet).
2. Classifies each survey question into one of four maturity levels (L1–L4).
3. Computes a weighted score for each level and aggregates them into a single **Final Score** (0–100).
4. Assigns a **Maturity Rank** of **M1**, **M2**, **M3**, or **M4** based on the Final Score.
5. Generates a **DEMaturitySummary** sheet with a consolidated results table and hyperlinks.
6. Generates an **individual project sheet** per respondent showing per-level question-and-response breakdowns, with non-optimal responses highlighted in red and auto-filtered for easy review.

### User Interaction

The add-in surfaces a task pane panel in Excel. Users click the **"Calculate Maturity"** button to trigger the full calculation pipeline. A modal wait dialog (`Wait.Show`) is displayed during processing.

---

## 2. Architecture and Prerequisites

### Required Workbook Structure

| Sheet Name | Required | Description |
|------------|----------|-------------|
| `Form1`    | ✅ Yes   | Contains the survey response data as an Excel table |

### Required Table Structure

| Table Name | Required | Location |
|------------|----------|----------|
| `Table1`   | ✅ Yes   | Must exist on the `Form1` sheet |

The table must have:
- A **header row** where each column corresponds to a survey question (used as the question text in the output).
- A **data body range** where each row is one project's complete survey response.

### Key Column Positions in `Table1`

Column indexes below are **zero-based** and refer to the raw column position in `Table1`.

| Column Index | Field Name     | Notes |
|-------------|----------------|-------|
| 0           | Response ID    | Unique identifier per submission |
| 2           | Review Date    | Stored as an Excel serial date number |
| 3           | Email          | Respondent's email address |
| 5           | Project Name   | Used as output sheet name and summary label |
| 6           | Resource Count | Number of project resources |
| 7–91        | Survey Responses | Question answers mapped to maturity levels |

> **Note:** Existing project sheets (any sheet that does not start with `_` and is not `Form1`) are **deleted and regenerated** on each run. System-generated sheets prefixed with `_` are preserved.

---

## 3. Maturity Level Architecture

Survey questions are classified into four maturity levels. Each level represents an increasing sophistication of practice. The column index of each question in `Table1` determines its level assignment.

### Level Definitions and Question Indexes

#### Level 1 — Foundational Practices (58 questions)

```
Indexes: [7, 8, 9, 10, 11, 12, 13, 14, 18, 19, 20, 21, 22, 23, 24, 27, 28,
          29, 30, 31, 33, 34, 35, 36, 37, 38, 44, 45, 46, 47, 48, 49, 50,
          51, 52, 53, 54, 59, 60, 61, 62, 63, 64, 67, 68, 72, 73, 74, 75,
          76, 79, 80, 81, 82, 83, 84, 85, 86]
```

- **Count:** 58 questions  
- **Weightage:** 60% of the Final Score  
- **Max raw score:** 580 points (58 × 10)  
- **Threshold alert:** Level 1 Weighted Score is highlighted **red** in output if it falls below 60 points.

#### Level 2 — Intermediate Practices (15 questions)

```
Indexes: [15, 16, 25, 17, 32, 39, 55, 56, 57, 65, 66, 69, 77, 87, 88]
```

> **Note:** Index `17` appears after `25` in the source array. The order of indexes within the array does not affect scoring (each is looked up individually), but the question iteration order in the output table follows this exact sequence.

- **Count:** 15 questions  
- **Weightage:** 20% of the Final Score  
- **Max raw score:** 150 points (15 × 10)

#### Level 3 — Advanced Practices (9 questions)

```
Indexes: [26, 40, 41, 42, 43, 58, 70, 71, 78]
```

- **Count:** 9 questions  
- **Weightage:** 10% of the Final Score  
- **Max raw score:** 90 points (9 × 10)

#### Level 4 — Expert Practices (3 questions)

```
Indexes: [89, 90, 91]
```

- **Count:** 3 questions  
- **Weightage:** 10% of the Final Score  
- **Max raw score:** 30 points (3 × 10)

### Weightage Summary Table

| Level   | Questions | Column Indexes       | Weightage | Max Weighted Score |
|---------|-----------|----------------------|-----------|--------------------|
| Level 1 | 58        | 7–86 (subset)        | 60%       | 60.00              |
| Level 2 | 15        | 15–88 (subset)       | 20%       | 20.00              |
| Level 3 | 9         | 26–78 (subset)       | 10%       | 10.00              |
| Level 4 | 3         | 89, 90, 91           | 10%       | 10.00              |
| **Total** | **85**  |                      | **100%**  | **100.00**         |

### Response Score Mapping

Each survey answer is mapped to a numeric score (0–10):

| Response Value | Score | Meaning |
|---------------|-------|---------|
| `Always`      | 10    | Best practice fully adopted |
| `Yes`         | 10    | Positive binary response |
| `NA`          | 10    | Not applicable; treated as passing |
| `Frequently`  | 7     | Mostly adopted |
| `Sometimes`   | 4     | Partially adopted |
| `Never`       | 0     | Not adopted |
| `No`          | 0     | Negative binary response |
| `Don't Know`  | 0     | Unknown / not assessed |
| *(missing/unrecognized)* | 0 | Defaults to 0 |

> A response is classified as a **failure** (highlighted red in project sheets) when its score is **less than the maximum of 10** — i.e., any value other than `Always`, `Yes`, or `NA`.

---

## 4. Scoring Algorithm — `CalculateLevelScore(projectRow, responseScores, levelIndexes, weightage)`

This inner function is called once per level per project response. It performs all score computation for a single maturity level.

### Parameters

| Parameter       | Type     | Description |
|----------------|----------|-------------|
| `projectRow`   | `Array`  | The full row of values for one project from `Table1` |
| `responseScores` | `Object` | The response-to-score mapping dictionary |
| `levelIndexes` | `Array`  | Column indexes belonging to this maturity level |
| `weightage`    | `number` | Weight percentage (60, 20, 10, or 10) |

### Algorithm Steps

```
1. Initialize:
   - levelScore       = 0.0
   - levelFailures    = {}   (dictionary of failing question indexes)
   - maxLevelScore    = levelIndexes.length × 10

2. For each columnIndex in levelIndexes:
   a. Read response = projectRow[columnIndex]
   b. If response exists and is a recognized key in responseScores:
      - score = responseScores[response]   (a float)
      - If score < 10 → add columnIndex to levelFailures
      - Add score to levelScore
   c. Else:
      - Add 0 to levelScore (response missing or unrecognized)

3. Compute:
   levelPercentage    = (levelScore / maxLevelScore) × 100
   weightedLevelScore = (levelPercentage × weightage) / 100
                        rounded to 2 decimal places

4. Return:
   {
     levelFailures,           // Object: { index: index } for each failing question
     unWeightedLevelScore,    // Raw sum of scores
     unWeightedLevelPercentage, // Percentage of raw score vs max
     weightedLevelScore       // Final contribution to total score
   }
```

### Example Calculation (Level 1)

Assume a project where all 58 Level 1 questions are answered `Always` (score 10):

```
levelScore          = 58 × 10 = 580
maxLevelScore       = 58 × 10 = 580
levelPercentage     = (580 / 580) × 100 = 100%
weightedLevelScore  = (100 × 60) / 100 = 60.00
```

Assume the same project answers all Level 1 questions as `Frequently` (score 7):

```
levelScore          = 58 × 7  = 406
maxLevelScore       = 580
levelPercentage     = (406 / 580) × 100 ≈ 70.0%
weightedLevelScore  = (70.0 × 60) / 100 = 42.00
```

### Final Score Aggregation

```
finalScore = level1CalculationDetails.weightedLevelScore
           + level2CalculationDetails.weightedLevelScore
           + level3CalculationDetails.weightedLevelScore
           + level4CalculationDetails.weightedLevelScore
```

The `finalScore` is a floating-point number in the range **0.00 – 100.00**.

---

## 5. Maturity Ranking Thresholds (M1 – M4)

The Final Score is mapped to a single letter-number maturity designation:

| Maturity Rank | Final Score Range        | Condition in Code |
|--------------|-------------------------|--------------------|
| **M1**       | 0.00 – 55.00 (inclusive) | `finalScore <= 55` |
| **M2**       | 55.01 – 75.00 (inclusive)| `finalScore > 55 && finalScore <= 75` |
| **M3**       | 75.01 – 90.00 (inclusive)| `finalScore > 75 && finalScore <= 90` |
| **M4**       | 90.01 – 100.00           | `finalScore > 90` |

> **Boundary note:** The boundaries at 55, 75, and 90 use `<=` (inclusive upper bound), meaning a score of exactly 55.00 is M1, exactly 75.00 is M2, and exactly 90.00 is M3. Scores must strictly exceed the threshold to advance to the next rank.

### Maturity Level Descriptions

| Rank | Description |
|------|-------------|
| M1   | **Initial** — Foundational practices are not consistently in place. Significant improvement opportunities exist. |
| M2   | **Developing** — Core practices are established but advanced and expert capabilities are still maturing. |
| M3   | **Defined** — Practices are well-established across foundational and intermediate levels with advanced capabilities emerging. |
| M4   | **Optimizing** — All four levels of practices are consistently applied, reflecting expert-level adoption. |

> **Visual indicator:** In all output sheets, the Maturity cell is formatted with a **yellow background** (`#FFFFE0`), **bold dark red font** (`#800000`), and **double borders** on all edges for immediate visual distinction.

### Level 1 Score Alert

Independently of the final maturity rank, if the **weighted Level 1 score is below 60 points** (the `LEVEL_MIN_THRESHOLD` constant), the Level 1 score cell is highlighted in **red** (`#FF0000`) in both the DEMaturitySummary sheet and the individual project sheet. This signals that foundational practices are insufficient regardless of the overall maturity rank.

---

## 6. Output Structure

### 6.1 `DEMaturitySummary` Sheet

A single consolidated results sheet named `DEMaturitySummary` is created. It contains one Excel table named `SummaryTable` with the following 11 columns:

| Column | Header | Data Source |
|--------|--------|-------------|
| A | `ID` | `projectResponses.values[i][0]` — Response ID |
| B | `PROJECT` | `projectResponses.values[i][5]` — Project name (hyperlink to project sheet) |
| C | `REVIEW DATE` | Excel serial date → `MM-DD-YYYY` string |
| D | `EMAIL` | `projectResponses.values[i][3]` |
| E | `RESOURCE COUNT` | `projectResponses.values[i][6]` |
| F | `LEVEL 1 SCORE` | `weightedLevelScore` (0–60.00); **red font** if < 60 |
| G | `LEVEL 2 SCORE` | `weightedLevelScore` (0–20.00) |
| H | `LEVEL 3 SCORE` | `weightedLevelScore` (0–10.00) |
| I | `LEVEL 4 SCORE` | `weightedLevelScore` (0–10.00) |
| J | `FINAL SCORE`  | Sum of all four weighted scores (0–100.00) |
| K | `MATURITY`     | M1 / M2 / M3 / M4 (hyperlink to project sheet; styled cell) |

**Hyperlinks:**
- Column B (`PROJECT`): navigates to cell `A2` of the corresponding project sheet.
- Column K (`MATURITY`): navigates to cell `B11` of the corresponding project sheet (the Maturity cell in the project summary).

### 6.2 Individual Project Sheets

One sheet is created per project response. The sheet name is derived as:

```
{first 25 characters of projectName, alphanumeric only}_{responseId}
```

*Example:* A project named `"My Project #1"` with ID `42` → sheet name `MyProject1_42`

#### Project Sheet Layout

**Rows A1:B11 — Summary Block**

| Row | Column A | Column B |
|-----|----------|----------|
| 1   | `ID`     | Response ID value |
| 2   | `PROJECT` | Project name |
| 3   | `REVIEW DATE` | Formatted date (`MM-DD-YYYY`) |
| 4   | `EMAIL`  | Email address |
| 5   | `RESOURCE COUNT` | Resource count |
| 6   | `LEVEL 1 SCORE` | Weighted L1 score (**red font** if < 60) |
| 7   | `LEVEL 2 SCORE` | Weighted L2 score |
| 8   | `LEVEL 3 SCORE` | Weighted L3 score |
| 9   | `LEVEL 4 SCORE` | Weighted L4 score |
| 10  | `FINAL SCORE` | Aggregated final score |
| 11  | `MATURITY` | M1/M2/M3/M4 (yellow background, bold dark red, double borders) |

**Row B12 — Navigation Hyperlink**

Cell `B12` contains a styled hyperlink:
> *"Click here to go to DEMaturitySummary sheet"*  
> Navigates to `DEMaturitySummary!A1`. Formatted in bold, brown-orange font (`#7C3606`), centered with yellow background (`#E1D70F`).

**Rows 13 onward — Maturity Level Question Tables**

Four sequential question-and-response tables are rendered in column B:C, each preceded by a merged, blue-header section label:

| Section Label | Table Name Pattern | Columns |
|--------------|-------------------|---------|
| `Level 1 Questions` | `Level1Table{responseId}` | Question, Response |
| `Level 2 Questions` | `Level2Table{responseId}` | Question, Response |
| `Level 3 Questions` | `Level3Table{responseId}` | Question, Response |
| `Level 4 Questions` | `Level4Table{responseId}` | Question, Response |

**Section header formatting:**
- Merged across B:C columns
- Bold white text (`#FFFFFF`)
- Blue background fill (`#154CC5`)
- Horizontally centered

**Row-level formatting in each table:**
- All row text defaults to black (`#000000`).
- Any row where the response scored **below 10** (i.e., not `Always`/`Yes`/`NA`) is highlighted in **red** (`#FF0000`).

**Auto-filter applied to all four tables:**  
The `Response` column of every level table is pre-filtered to show only the following non-optimal values:

```
["Frequently", "Sometimes", "Never", "No", "Don't Know"]
```

This means when a project sheet is first opened, only the **improvement-area responses** are visible, making it easy for reviewers to focus on gaps without manually filtering.

---

## 7. Key Data Flow

```
┌─────────────────────────────────────────────────────────────────┐
│                  EXCEL WORKBOOK (Input)                         │
│  Sheet: Form1                                                   │
│  Table: Table1                                                  │
│  ┌────────────────────────────────────────────────────────┐     │
│  │ Header Row: [col0..col91] → Question Text              │     │
│  │ Data Rows:  [col0..col91] → Survey Responses per row   │     │
│  └────────────────────────────────────────────────────────┘     │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                    User clicks "Calculate Maturity"
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                    INITIALIZATION                               │
│  1. Load all existing worksheets                               │
│  2. Delete all sheets except Form1 and _* sheets               │
│  3. Partition header columns into 4 level question dictionaries │
│     • level1Questions  (indexes: L1 array)                     │
│     • level2Questions  (indexes: L2 array)                     │
│     • level3Questions  (indexes: L3 array)                     │
│     • level4Questions  (indexes: L4 array)                     │
│  4. Create DEMaturitySummary sheet + SummaryTable (11 columns) │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│            PER-PROJECT SCORING LOOP (Pass 1)                   │
│  For each project response row:                                 │
│    ┌──────────────────────────────────────────────────────┐    │
│    │ CalculateLevelScore(row, scores, L1indexes, 60)       │    │
│    │   → { weightedLevelScore, levelFailures, ... }        │    │
│    │ CalculateLevelScore(row, scores, L2indexes, 20)       │    │
│    │ CalculateLevelScore(row, scores, L3indexes, 10)       │    │
│    │ CalculateLevelScore(row, scores, L4indexes, 10)       │    │
│    └──────────────────────────────────────────────────────┘    │
│    finalScore = L1.weighted + L2.weighted + L3.weighted         │
│                + L4.weighted                                    │
│    maturity   = M1 | M2 | M3 | M4  (threshold lookup)         │
│    → Append row to SummaryTable                                │
│    → Conditionally red-highlight L1 score if < 60             │
│    → Create blank project sheet                                │
│    → Push all data to projectSheetsData[]                      │
└──────────────────────────────┬──────────────────────────────────┘
                               │ context.sync()
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│          PER-PROJECT SHEET WRITING LOOP (Pass 2)               │
│  For each projectSheetData entry:                               │
│    1. Write A1:B11 summary block (metadata + scores)           │
│    2. Style Maturity cell (yellow bg, bold dark red, borders)  │
│    3. Red-highlight L1 score cell if below threshold           │
│    4. Render Level 1 question table  (addLevelTable)           │
│    5. Render Level 2 question table  (addLevelTable)           │
│    6. Render Level 3 question table  (addLevelTable)           │
│    7. Render Level 4 question table  (addLevelTable)           │
│    8. Auto-fit columns and rows                                │
│    9. context.sync() per sheet                                 │
└──────────────────────────────┬──────────────────────────────────┘
                               │ context.sync()
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│           HYPERLINK & FILTER PASS (Pass 3)                     │
│  For each project:                                              │
│    1. Apply response filter to L1/L2/L3/L4 tables             │
│       (show only: Frequently/Sometimes/Never/No/Don't Know)    │
│    2. Add PROJECT → project sheet hyperlink in SummaryTable    │
│    3. Add MATURITY → project sheet hyperlink in SummaryTable   │
│    4. Add "Go to DEMaturitySummary" hyperlink in project sheet │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                    FINALIZATION                                 │
│  1. Auto-fit DEMaturitySummary columns and rows                │
│  2. Activate DEMaturitySummary sheet                           │
│  3. Close Wait dialog                                          │
└─────────────────────────────────────────────────────────────────┘
```

---

## 8. Helper Functions Reference

All helper functions are defined as closures inside the `Excel.run` callback, giving them direct access to the `context` and all level data structures.

### `getQuestions(questionsRow, levelIndexes)`

Extracts the question text for a given level from the header row.

| Parameter | Type | Description |
|-----------|------|-------------|
| `questionsRow` | `Array` | The full header row values (2D array; uses `[0]`) |
| `levelIndexes` | `Array` | Column indexes for the target level |

**Returns:** `Object` — `{ [columnIndex]: questionText }` dictionary.

---

### `CalculateLevelScore(projectRow, responseScores, levelIndexes, weightage)`

See [Section 4](#4-scoring-algorithm--calculatelevelscoreprojectrow-responsescores-levelindexes-weightage) for full documentation.

**Returns:**
```js
{
  levelFailures: Object,             // { index: index } for each non-perfect response
  unWeightedLevelScore: number,      // Raw score sum
  unWeightedLevelPercentage: number, // Raw score as percentage of max possible
  weightedLevelScore: number         // Contribution to final score (2 decimal places)
}
```

---

### `getJsDateFromExcel(dateValue)`

Converts an Excel serial date number to a `MM-DD-YYYY` string.

| Parameter | Type | Description |
|-----------|------|-------------|
| `dateValue` | `number` | Excel serial date (days since 1900-01-00) |

**Formula used:**
```js
new Date((dateValue - (25567 + 2)) * 86400 * 1000)
```
The offset `25569` (`25567 + 2`) accounts for Excel's epoch (January 1, 1900) adjusted for the Excel 1900 leap year bug.

**Returns:** `string` in format `"MM-DD-YYYY"`.

---

### `addLevelTable(questions, levelFailures, projectSheet, levelIndexes, level, projectResponse, rowIndex, context)`

Renders a named Excel table for one maturity level onto the given project sheet.

| Parameter | Type | Description |
|-----------|------|-------------|
| `questions` | `Object` | Question dictionary for this level |
| `levelFailures` | `Object` | Failing question indexes from `CalculateLevelScore` |
| `projectSheet` | `WorksheetObject` | Target Excel worksheet |
| `levelIndexes` | `Array` | Column indexes for this level |
| `level` | `string` | Level prefix string: `"Level1"`, `"Level2"`, `"Level3"`, `"Level4"` |
| `projectResponse` | `Array` | Full response row for this project |
| `rowIndex` | `number` | Starting row index for table placement |
| `context` | `RequestContextObject` | Excel API context |

**Behavior:**
- Creates a table named `{level}Table{responseId}` (e.g., `Level1Table42`).
- Sets headers: `["Question", "Response"]`.
- Adds one row per question in `levelIndexes`.
- Colors rows red (`#FF0000`) where the question index appears in `levelFailures`.

**Returns:** `number` — Updated `rowIndex` after all rows have been added (for chaining next table placement).

---

### `applyFilter(filter)`

Applies a pre-configured value filter to the `Response` column of a level table.

**Filter values shown:** `["Frequently", "Sometimes", "Never", "No", "Don't Know"]`

This ensures only non-perfect responses are visible by default when the project sheet is opened.

---

### `applyLevelQuestionTopRowProperties(levelRange)`

Applies the blue section header style to a level's title row.

| Style Property | Value |
|---------------|-------|
| Merge cells | `true` (B:C merged) |
| Font bold | `true` |
| Font color | `#FFFFFF` (white) |
| Fill color | `#154CC5` (blue) |
| Horizontal alignment | `Center` |

---

## 9. Error Handling

The `Excel.run` promise chain includes a `.catch` handler. Two error conditions are explicitly handled:

| Error Code | User Message |
|-----------|--------------|
| `ItemNotFound` | `"Please run add-in on DE Survey responses"` — Indicates `Form1` or `Table1` was not found. |
| Any other error | `"There is an error in processing project responses. Please try again or later"` |

Errors are surfaced to the user through the `MessageBox.Show` dialog from the `officejs.dialogs` library, using `MessageBoxButtons.OkOnly` and `MessageBoxIcons.Error`. The wait dialog is closed before showing the error dialog.

All errors are also logged to `console.error` for debugging via browser developer tools or the Office Add-in debugger.

---

## 10. Build and Development

### Prerequisites

- **Node.js** (v12+)
- **npm**
- **Microsoft Excel** (Desktop or Excel Online)

### Install Dependencies

```bash
npm install
```

### Development Server

Starts the webpack dev server with HTTPS and auto-reloading on port **3000**:

```bash
npm run dev-server
```

### Build

**Production build** (minified):

```bash
npm run build
```

**Development build** (with source maps):

```bash
npm run build:dev
```

**Watch mode** (rebuilds on file change):

```bash
npm run watch
```

### Sideload and Debug in Excel

**Start (desktop):**

```bash
npm start
# or explicitly:
npm run start:desktop
```

**Start (Excel Online / web):**

```bash
npm run start:web
```

**Stop debugging:**

```bash
npm stop
```

### Validate Manifest

```bash
npm run validate
```

### Lint

```bash
npm run lint        # Check only
npm run lint:fix    # Auto-fix
```

### Key Source Files

| File | Purpose |
|------|---------|
| `src/taskpane/taskpane.js` | Core add-in logic — all scoring, sheet generation, and Excel API calls |
| `src/taskpane/taskpane.html` | Task pane UI — single "Calculate Maturity" button |
| `src/taskpane/taskpane.css` | Task pane styling (Office UI Fabric) |
| `src/taskpane/dialogs.js` | Dialog helpers (`Wait`, `MessageBox`, `Alert` from `officejs.dialogs`) |
| `src/commands/commands.js` | Commands module (entry point for ribbon button actions) |
| `manifest.xml` | Add-in manifest — registration, permissions, UI extension points |
| `webpack.config.js` | Build configuration with Babel transpilation, HTTPS dev server |

### Webpack Entry Points

| Entry | File | Output |
|-------|------|--------|
| `polyfill` | `@babel/polyfill` | Browser compatibility polyfills |
| `taskpane` | `src/taskpane/taskpane.js` | Main add-in logic bundle |
| `commands` | `src/commands/commands.js` | Ribbon command handlers |

---

## 11. Changelog — M4 Maturity Level Addition

This section documents the specific code changes introduced when the M4 maturity tier was added.

### Changes to `taskpane.js`

#### 1. Level 4 Index Array Added

```js
// NEW
var level4MaturityIndexes = [89, 90, 91];
```

#### 2. Level 3 Indexes Updated (indexes 89–91 moved to Level 4)

```js
// BEFORE
var level3MaturityIndexes = [26, 40, 41, 42, 43, 58, 70, 71, 78, 89, 90, 91];

// AFTER
var level3MaturityIndexes = [26, 40, 41, 42, 43, 58, 70, 71, 78];
```

#### 3. Level 4 Score Calculation Added

```js
// NEW call in the per-project scoring loop
var level4CalculationDetails = CalculateLevelScore(
  projectResponses.values[i], responseScores, level4MaturityIndexes, 10
);
```

#### 4. Final Score Updated to Include Level 4

```js
// BEFORE
var finalScore = level1CalculationDetails.weightedLevelScore
               + level2CalculationDetails.weightedLevelScore
               + level3CalculationDetails.weightedLevelScore;

// AFTER
var finalScore = level1CalculationDetails.weightedLevelScore
               + level2CalculationDetails.weightedLevelScore
               + level3CalculationDetails.weightedLevelScore
               + level4CalculationDetails.weightedLevelScore;
```

#### 5. M4 Maturity Threshold Added

```js
// BEFORE: max rank was M3 for finalScore > 75
else if (finalScore > 75) {
  maturity = "M3";
}

// AFTER: M3 capped at 90, M4 added
else if (finalScore > 75 && finalScore <= 90) {
  maturity = "M3";
}
else if (finalScore > 90) {
  maturity = "M4";
}
```

#### 6. Summary Table Header Expanded to 11 Columns (A1:K1)

```js
// BEFORE: A1:J1 (10 columns)
var summaryTable = maturitySheet.tables.add("A1:J1", true);
summaryTable.getHeaderRowRange().values = [[
  "ID", "PROJECT", "REVIEW DATE", "EMAIL", "RESOURCE COUNT",
  "LEVEL 1 SCORE", "LEVEL 2 SCORE", "LEVEL 3 SCORE",
  "FINAL SCORE", "MATURITY"
]];

// AFTER: A1:K1 (11 columns)
var summaryTable = maturitySheet.tables.add("A1:K1", true);
summaryTable.getHeaderRowRange().values = [[
  "ID", "PROJECT", "REVIEW DATE", "EMAIL", "RESOURCE COUNT",
  "LEVEL 1 SCORE", "LEVEL 2 SCORE", "LEVEL 3 SCORE", "LEVEL 4 SCORE",
  "FINAL SCORE", "MATURITY"
]];
```

#### 7. Level 4 Score Added to Summary Table Row

```js
summaryTable.rows.add(null, [[
  responseId, projectName, reviewDate, email, resourceCount,
  level1CalculationDetails.weightedLevelScore,
  level2CalculationDetails.weightedLevelScore,
  level3CalculationDetails.weightedLevelScore,
  level4CalculationDetails.weightedLevelScore,  // NEW
  finalScore, maturity
]]);
```

#### 8. Level 4 Question Table Added to Project Sheets

```js
// NEW section rendered after Level 3 table
projectSheet.getRange("B" + (rowIndex + 3)).values = [["Level 4 Questions"]];
var level4Range = projectSheet.getRange("B" + (rowIndex + 3) + ":" + "C" + (rowIndex + 3));
applyLevelQuestionTopRowProperties(level4Range);

rowIndex = addLevelTable(
  level4Questions,
  projectSheetData.level4CalculationDetails.levelFailures,
  projectSheet,
  level4MaturityIndexes,
  "Level4",
  projectResponses.values[i],
  rowIndex + 4,
  context
);
```

#### 9. Level 4 Filter Applied in Pass 3

```js
// NEW
var level4filter = projectSheet.tables
  .getItem("Level4Table" + projectResponses.values[i][0])
  .columns.getItem("Response").filter;
applyFilter(level4filter);
```

#### 10. Project Sheet Summary Expanded to A1:B11

```js
// BEFORE: A1:B10 (10 rows, no Level 4 score row)

// AFTER: A1:B11 (11 rows, includes Level 4 score)
var summaryRange = projectSheet.getRange("A1:B11");
summaryData = [
  ...,
  ["LEVEL 4 SCORE", projectSheetData.level4CalculationDetails.weightedLevelScore],  // NEW (row 9)
  ["FINAL SCORE",   ...],
  ["MATURITY",      ...]
];
// Maturity cell reference updated from getCell(9, 1) to getCell(10, 1)
// to account for the new Level 4 Score row pushing Maturity one position down
var maturityCellRange = summaryRange.getCell(10, 1);
// Level 1 score cell remains at getCell(5, 1) — unchanged
var level1ScoreRange = summaryRange.getCell(5, 1);
```

---

*Documentation generated from source code analysis of `src/taskpane/taskpane.js` and `manifest.xml`.*  
*Provider: Neudesic | License: MIT*
