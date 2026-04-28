# API Reference — DEMaturityCalculator Excel Add-in

> This document provides detailed reference documentation for every function, constant, data structure, and scoring rule in `src/taskpane/taskpane.js`.

---

## Table of Contents

1. [Module Overview](#1-module-overview)
2. [Constants](#2-constants)
3. [Configuration Objects](#3-configuration-objects)
   - 3.1 [responseScores](#31-responsescores)
   - 3.2 [Maturity Level Index Arrays](#32-maturity-level-index-arrays)
4. [Public Functions](#4-public-functions)
   - 4.1 [run()](#41-run)
5. [Internal Functions](#5-internal-functions)
   - 5.1 [getQuestions()](#51-getquestions)
   - 5.2 [CalculateLevelScore()](#52-calculatelevelscore)
   - 5.3 [getJsDateFromExcel()](#53-getjsdatefromexcel)
   - 5.4 [addLevelTable()](#54-addleveltable)
   - 5.5 [applyLevelQuestionTopRowProperties()](#55-applylevelquestiontoprowproperties)
   - 5.6 [applyFilter()](#56-applyfilter)
6. [Scoring Logic](#6-scoring-logic)
   - 6.1 [Per-Response Scoring](#61-per-response-scoring)
   - 6.2 [Level Score Calculation](#62-level-score-calculation)
   - 6.3 [Weighted Score Calculation](#63-weighted-score-calculation)
   - 6.4 [Final Score & Maturity Label](#64-final-score--maturity-label)
   - 6.5 [Worked Example](#65-worked-example)
7. [Maturity Level Definitions](#7-maturity-level-definitions)
8. [Output Data Structures](#8-output-data-structures)
   - 8.1 [LevelCalculationDetails Object](#81-levelcalculationdetails-object)
   - 8.2 [ProjectSheetData Object](#82-projectsheetdata-object)
   - 8.3 [DEMaturitySummary Row](#83-dematuritysummary-row)
   - 8.4 [Project Sheet Summary Block](#84-project-sheet-summary-block)
9. [Error Handling](#9-error-handling)
10. [Office JS API Dependencies](#10-office-js-api-dependencies)

---

## 1. Module Overview

**File:** `src/taskpane/taskpane.js`  
**Module system:** ES Modules (`export async function run`)  
**Global dependencies:** `console`, `document`, `Excel` (Office JS), `Office` (Office JS), `Wait`, `MessageBox`, `MessageBoxButtons`, `MessageBoxIcons` (from `dialogs.js`)

The module exports a single public function (`run`) and defines all other functions as closures **inside** the `Excel.run` callback. This scoping means the internal functions have direct access to the Excel `context` object and the shared index arrays without needing to pass them as parameters (where not already done so).

---

## 2. Constants

These constants are defined inside `run()` and are visible to all inner functions.

| Name | Value | Type | Description |
|---|---|---|---|
| `RESPONSE_MAX_SCORE` | `10` | `number` | The maximum score that any single response can contribute. Responses equal to this value are considered passing. |
| `LEVEL_MIN_THRESHOLD` | `60` | `number` | The minimum acceptable weighted score for Level 1. Cells are highlighted red when a project's Level 1 score falls below this value. |

---

## 3. Configuration Objects

### 3.1 `responseScores`

**Type:** `Object.<string, number>`

A lookup table mapping each valid survey response string to its numeric score.

```javascript
var responseScores = {};
responseScores["NA"]         = 10;  // Not Applicable — treated as full score
responseScores["Always"]     = 10;  // Full score
responseScores["Frequently"] = 7;   // Partial score
responseScores["Sometimes"]  = 4;   // Partial score
responseScores["Never"]      = 0;   // Zero score
responseScores["Yes"]        = 10;  // Full score (binary questions)
responseScores["No"]         = 0;   // Zero score (binary questions)
responseScores["Don't Know"] = 0;   // Zero score (treated as non-compliant)
```

**Scoring table:**

| Response | Score | Category |
|---|---|---|
| `Always` | 10 | ✅ Full |
| `NA` | 10 | ✅ Full (exemption) |
| `Yes` | 10 | ✅ Full |
| `Frequently` | 7 | ⚠️ Partial |
| `Sometimes` | 4 | ⚠️ Partial |
| `Never` | 0 | ❌ Zero |
| `No` | 0 | ❌ Zero |
| `Don't Know` | 0 | ❌ Zero |
| *(empty / unrecognised)* | 0 | ❌ Zero (implicit) |

> **Failure detection:** A response is recorded as a **failure** (highlighted red, included in filter) when its score is **not equal to** `RESPONSE_MAX_SCORE` (10). This means `Frequently` and `Sometimes` are treated as failures even though they contribute a non-zero score.

---

### 3.2 Maturity Level Index Arrays

These arrays define which **0-based column indexes** in the `Table1` data body belong to each maturity level. The indexes correspond to specific survey questions in the Microsoft Forms export.

#### `level1MaturityIndexes`
**Weight:** 60% of final score  
**Question count:** 58

```javascript
var level1MaturityIndexes = [
  7, 8, 9, 10, 11, 12, 13, 14,
  18, 19, 20, 21, 22, 23, 24,
  27, 28, 29, 30, 31,
  33, 34, 35, 36, 37, 38,
  44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54,
  59, 60, 61, 62, 63, 64,
  67, 68,
  72, 73, 74, 75, 76,
  79, 80, 81, 82, 83, 84, 85, 86
];
```

#### `level2MaturityIndexes`
**Weight:** 20% of final score  
**Question count:** 15

```javascript
var level2MaturityIndexes = [
  15, 16, 25, 17, 32, 39,
  55, 56, 57,
  65, 66, 69, 77,
  87, 88
];
```

#### `level3MaturityIndexes`
**Weight:** 10% of final score  
**Question count:** 6

```javascript
var level3MaturityIndexes = [26, 40, 41, 42, 43, 58];
```

#### `level4MaturityIndexes`
**Weight:** 10% of final score  
**Question count:** 6

```javascript
var level4MaturityIndexes = [70, 71, 78, 89, 90, 91];
```

**Summary table:**

| Level | Indexes | Question Count | Max Raw Score | Weight | Max Weighted Score |
|---|---|---|---|---|---|
| Level 1 | 7–86 (non-contiguous) | 58 | 580 | 60% | 60.00 |
| Level 2 | 15–88 (non-contiguous) | 15 | 150 | 20% | 20.00 |
| Level 3 | 26–58 (non-contiguous) | 6 | 60 | 10% | 10.00 |
| Level 4 | 70–91 (non-contiguous) | 6 | 60 | 10% | 10.00 |
| **Total** | — | **85** | **850** | **100%** | **100.00** |

---

## 4. Public Functions

### 4.1 `run()`

**Signature:**
```javascript
export async function run(): Promise<void>
```

**Description:**  
The sole entry point for the add-in's business logic. Invoked when the user clicks the **Calculate Maturity** button. Orchestrates the full lifecycle: reading input data, computing scores, and writing all output sheets.

**Returns:** `Promise<void>` — The function is async; it resolves when all Excel operations are complete and the DEMaturitySummary sheet is activated.

**Side effects:**

1. Displays a loading spinner (`Wait.Show`).
2. Deletes all sheets in the workbook whose names do not start with `_` and are not `Form1`.
3. Creates a new `DEMaturitySummary` sheet with a `SummaryTable`.
4. Creates one new sheet per project row in `Table1`.
5. Closes the loading spinner.
6. On error: closes the spinner and shows a `MessageBox` with a user-friendly message.

**Throws:** Catches all `Excel.run` errors internally. Outer `try/catch` logs unexpected errors to `console.error` without surfacing them to the user.

**Usage:**

```javascript
// Wired up automatically in Office.onReady:
document.getElementById("run").onclick = run;
```

---

## 5. Internal Functions

All internal functions are defined as `function` declarations inside the `Excel.run` callback. They have closure access to `context`, all index arrays, `responseScores`, `RESPONSE_MAX_SCORE`, and `LEVEL_MIN_THRESHOLD`.

---

### 5.1 `getQuestions()`

```javascript
function getQuestions(questionsRow, levelIndexes): Object.<number, string>
```

**Description:**  
Extracts question text for a specific maturity level from the table's header row.

**Parameters:**

| Parameter | Type | Description |
|---|---|---|
| `questionsRow` | `Array<Array<string>>` | The 2D array from `responsesTable.getHeaderRowRange().values` (i.e., `questions.values`). The outer array has one element (the header row); the inner array contains all column header strings. |
| `levelIndexes` | `Array<number>` | The index array for the desired maturity level (e.g., `level1MaturityIndexes`). |

**Returns:** `Object.<number, string>` — A dictionary where each key is a column index and each value is the question text at that index.

**Example:**

```javascript
// Input header row (simplified):
questionsRow = [["ID", "Date", ..., "Do you use CI/CD?", ..., "Do you automate tests?"]]
// levelIndexes = [7, 8]

// Output:
{ 7: "Do you use CI/CD?", 8: "Do you automate tests?" }
```

---

### 5.2 `CalculateLevelScore()`

```javascript
function CalculateLevelScore(
  projectRow,
  responseScores,
  levelIndexes,
  weightage
): LevelCalculationDetails
```

**Description:**  
Computes all scoring metrics for a single maturity level for a single project. This is the core scoring function.

**Parameters:**

| Parameter | Type | Description |
|---|---|---|
| `projectRow` | `Array<any>` | One row from `responsesTable.getDataBodyRange().values` — the full array of cell values for one project. |
| `responseScores` | `Object.<string, number>` | The response-to-score lookup table (see [§3.1](#31-responsescores)). |
| `levelIndexes` | `Array<number>` | Column indexes for the questions belonging to this level. |
| `weightage` | `number` | The percentage weight for this level (60, 20, 10, or 10). |

**Returns:** [`LevelCalculationDetails`](#81-levelcalculationdetails-object)

**Algorithm:**

```
maxLevelScore = levelIndexes.length × 10

for each index in levelIndexes:
  response = projectRow[index]
  if response exists AND responseScores[response] is defined:
    score = responseScores[response]
    if score ≠ RESPONSE_MAX_SCORE (10):
      record index in levelFailures
    accumulate score
  else:
    add 0

levelPercentage     = (levelScore / maxLevelScore) × 100
weightedLevelScore  = (levelPercentage × weightage) / 100
                      → rounded to 2 decimal places (toFixed(2))
```

**Example call:**

```javascript
var result = CalculateLevelScore(
  projectRow,
  responseScores,
  level1MaturityIndexes,  // 58 questions
  60                      // 60% weight
);
// result.weightedLevelScore is between 0.00 and 60.00
```

---

### 5.3 `getJsDateFromExcel()`

```javascript
function getJsDateFromExcel(dateValue): string
```

**Description:**  
Converts an Excel serial date number (days since 1900-01-00, with a known +2 offset correction) into a formatted date string.

**Parameters:**

| Parameter | Type | Description |
|---|---|---|
| `dateValue` | `number` | The numeric Excel date serial value stored in a cell (e.g., `44927` for 2023-01-01). |

**Returns:** `string` — Date formatted as `MM-DD-YYYY`.

**Implementation detail:**  
Excel's epoch is January 0, 1900. The formula subtracts `25567 + 2` (the Unix epoch offset plus Excel's known leap-year bug correction) and multiplies by `86400 * 1000` milliseconds per day:

```javascript
var reviewDate = new Date((dateValue - (25567 + 2)) * 86400 * 1000);
```

**Example:**

```javascript
getJsDateFromExcel(44927)  // → "01-01-2023"
getJsDateFromExcel(45292)  // → "12-31-2023"
```

---

### 5.4 `addLevelTable()`

```javascript
function addLevelTable(
  questions,
  levelFailures,
  projectSheet,
  levelIndexes,
  level,
  projectResponse,
  rowIndex,
  context
): number
```

**Description:**  
Creates and populates a two-column Excel table (`Question` | `Response`) for one maturity level on a project sheet. Highlights failing rows in red.

**Parameters:**

| Parameter | Type | Description |
|---|---|---|
| `questions` | `Object.<number, string>` | Dictionary of `{ columnIndex: questionText }` for this level. |
| `levelFailures` | `Object.<number, number>` | Dictionary of `{ columnIndex: columnIndex }` for questions that did not receive full score. |
| `projectSheet` | `Excel.Worksheet` | The worksheet object to write the table into. |
| `levelIndexes` | `Array<number>` | Column indexes for this level's questions. |
| `level` | `string` | Level identifier used to name the table: `"Level1"`, `"Level2"`, `"Level3"`, or `"Level4"`. |
| `projectResponse` | `Array<any>` | The full project response row (used to read answer values and derive the table name from `projectResponse[0]`, the response ID). |
| `rowIndex` | `number` | The 1-based Excel row number at which to start the table. |
| `context` | `Excel.RequestContext` | The current Excel run context (not used directly but kept for future use). |

**Returns:** `number` — The updated `rowIndex` after all rows have been added (i.e., the row immediately after the last question row).

**Table naming convention:**
```
"Level1Table" + responseId    (e.g., "Level1Table42")
"Level2Table" + responseId
"Level3Table" + responseId
"Level4Table" + responseId
```

**Table structure:**

| Column | Header | Content |
|---|---|---|
| B (col 1) | Question | Question text |
| C (col 2) | Response | Project's answer |

**Failure highlighting:**  
Any row where `levelFailures[levelIndex] !== undefined` has its entire row's font colour set to `#FF0000` (red).

---

### 5.5 `applyLevelQuestionTopRowProperties()`

```javascript
function applyLevelQuestionTopRowProperties(levelRange): void
```

**Description:**  
Applies consistent header formatting to the two-cell merged range above each level questions table.

**Parameters:**

| Parameter | Type | Description |
|---|---|---|
| `levelRange` | `Excel.Range` | A two-column range (e.g., `B13:C13`) that serves as the section heading for a level's question table. |

**Applied styles:**

| Property | Value |
|---|---|
| Merge | `true` (merges the two cells) |
| Font weight | Bold |
| Font colour | `#FFFFFF` (white) |
| Horizontal alignment | Center |
| Fill colour | `#154CC5` (Neudesic blue) |

---

### 5.6 `applyFilter()`

```javascript
function applyFilter(filter): void
```

**Description:**  
Applies a values-based auto-filter to the `Response` column of a level questions table so that only non-perfect (failing or partial) responses are shown by default.

**Parameters:**

| Parameter | Type | Description |
|---|---|---|
| `filter` | `Excel.Filter` | The filter object for the `Response` column of a level table. |

**Filter values applied:**

```javascript
["Frequently", "Sometimes", "Never", "No", "Don't Know"]
```

This means rows with `Always`, `NA`, or `Yes` responses are **hidden** after processing, making it easy to focus on areas needing improvement.

---

## 6. Scoring Logic

### 6.1 Per-Response Scoring

Each question's answer cell is looked up in `responseScores`. If the response string is not found (empty cell, unrecognised value), the score defaults to **0**.

```
score = responseScores[response] ?? 0
```

A response is a **failure** (flagged and highlighted) if `score ≠ 10`.

---

### 6.2 Level Score Calculation

For a level with *N* questions:

```
Raw score     = Σ score(response_i)   for i = 1..N
Max raw score = N × 10
Level %       = (Raw score / Max raw score) × 100
```

| Level | N (questions) | Max raw score |
|---|---|---|
| Level 1 | 58 | 580 |
| Level 2 | 15 | 150 |
| Level 3 | 6 | 60 |
| Level 4 | 6 | 60 |

---

### 6.3 Weighted Score Calculation

```
Weighted score = (Level % × Weight) / 100
```

| Level | Weight | Formula | Max weighted score |
|---|---|---|---|
| Level 1 | 60% | `(L1% × 60) / 100` | 60.00 |
| Level 2 | 20% | `(L2% × 20) / 100` | 20.00 |
| Level 3 | 10% | `(L3% × 10) / 100` | 10.00 |
| Level 4 | 10% | `(L4% × 10) / 100` | 10.00 |

Weighted scores are rounded to **2 decimal places** using `Number.prototype.toFixed(2)` and then converted back to a `float` with `parseFloat()`.

---

### 6.4 Final Score & Maturity Label

```
Final score = Weighted L1 + Weighted L2 + Weighted L3 + Weighted L4
            ∈ [0.00, 100.00]
```

**Maturity label assignment:**

| Condition | Label | Interpretation |
|---|---|---|
| `finalScore ≤ 60` | **M1** | Foundational — core practices not yet in place |
| `60 < finalScore ≤ 80` | **M2** | Developing — baseline practices established |
| `80 < finalScore ≤ 90` | **M3** | Advanced — strong practices with some gaps |
| `finalScore > 90` | **M4** | Optimised — near-complete adoption across all levels |

> **Boundary note:** A score of exactly 60 is classified as **M1**, not M2. A score of exactly 80 is **M2**, not M3. A score of exactly 90 is **M3**, not M4.

---

### 6.5 Worked Example

Suppose a project has the following responses for a simplified scenario:

**Level 1 (3 questions for illustration, weight = 60%):**

| Question index | Response | Score |
|---|---|---|
| 7 | Always | 10 |
| 8 | Frequently | 7 |
| 9 | Never | 0 |

```
Raw score     = 10 + 7 + 0 = 17
Max raw score = 3 × 10 = 30
Level %       = (17 / 30) × 100 = 56.67%
Weighted L1   = (56.67 × 60) / 100 = 34.00
```

**Level 2 (2 questions for illustration, weight = 20%):**

| Question index | Response | Score |
|---|---|---|
| 15 | Yes | 10 |
| 16 | No | 0 |

```
Raw score     = 10 + 0 = 10
Max raw score = 2 × 10 = 20
Level %       = (10 / 20) × 100 = 50.00%
Weighted L2   = (50.00 × 20) / 100 = 10.00
```

**Assuming Level 3 weighted = 8.00 and Level 4 weighted = 7.50:**

```
Final score = 34.00 + 10.00 + 8.00 + 7.50 = 59.50
Maturity    = M1  (59.50 ≤ 60)
```

---

## 7. Maturity Level Definitions

The add-in defines four Digital Engineering maturity levels. Each level's question set focuses on progressively more advanced practices.

| Level | Label | Score Range | Focus Area |
|---|---|---|---|
| **1** | M1 | 0 – 60 | Foundational DE practices (58 questions, 60% weight) |
| **2** | M2 | 60 – 80 | Intermediate / process-oriented practices (15 questions, 20% weight) |
| **3** | M3 | 80 – 90 | Advanced automation and integration practices (6 questions, 10% weight) |
| **4** | M4 | 90 – 100 | Optimisation and continuous improvement practices (6 questions, 10% weight) |

> The **score range** in the table above refers to the **final composite score** threshold that places a project at that maturity label, not the individual level score.

---

## 8. Output Data Structures

### 8.1 `LevelCalculationDetails` Object

Returned by `CalculateLevelScore()`.

| Property | Type | Description |
|---|---|---|
| `levelFailures` | `Object.<number, number>` | Dictionary of `{ index: index }` for all questions where the response score is less than 10. |
| `unWeightedLevelScore` | `number` | Raw sum of all response scores for this level (0 – N×10). |
| `unWeightedLevelPercentage` | `number` | Raw score expressed as a percentage of the maximum possible (0 – 100). |
| `weightedLevelScore` | `number` | Final weighted score contribution for this level (0.00 – weight). Rounded to 2 decimal places. |

**Example:**
```javascript
{
  levelFailures: { 8: 8, 9: 9 },
  unWeightedLevelScore: 17,
  unWeightedLevelPercentage: 56.666...,
  weightedLevelScore: 34.00
}
```

---

### 8.2 `ProjectSheetData` Object

Accumulated in `projectSheetsData[]` during Phase 2. Consumed in Phase 3 to populate project sheets.

| Property | Type | Description |
|---|---|---|
| `Id` | `any` | Response ID from column 0 of the project row. |
| `projectName` | `string` | Project name from column 5 of the project row. |
| `email` | `string` | Respondent email from column 3. |
| `reviewDate` | `string` | Formatted date string (`MM-DD-YYYY`) from column 2. |
| `resourceCount` | `any` | Resource count from column 6. |
| `level1CalculationDetails` | `LevelCalculationDetails` | Scoring details for Level 1. |
| `level2CalculationDetails` | `LevelCalculationDetails` | Scoring details for Level 2. |
| `level3CalculationDetails` | `LevelCalculationDetails` | Scoring details for Level 3. |
| `level4CalculationDetails` | `LevelCalculationDetails` | Scoring details for Level 4. |
| `finalScore` | `number` | Composite score (sum of all four weighted scores). |
| `maturity` | `string` | Maturity label: `"M1"`, `"M2"`, `"M3"`, or `"M4"`. |
| `projectSheetName` | `string` | Sanitised Excel sheet name for this project. |

---

### 8.3 `DEMaturitySummary` Row

Each row added to `SummaryTable` in the `DEMaturitySummary` sheet:

| Position | Column header | Source | Type |
|---|---|---|---|
| 0 | ID | `projectResponses.values[i][0]` | any |
| 1 | PROJECT | `projectResponses.values[i][5]` | string (hyperlink) |
| 2 | REVIEW DATE | `getJsDateFromExcel(column[2])` | string (MM-DD-YYYY) |
| 3 | EMAIL | `projectResponses.values[i][3]` | string |
| 4 | RESOURCE COUNT | `projectResponses.values[i][6]` | number |
| 5 | LEVEL 1 SCORE | `level1CalculationDetails.weightedLevelScore` | number (0–60) |
| 6 | LEVEL 2 SCORE | `level2CalculationDetails.weightedLevelScore` | number (0–20) |
| 7 | LEVEL 3 SCORE | `level3CalculationDetails.weightedLevelScore` | number (0–10) |
| 8 | LEVEL 4 SCORE | `level4CalculationDetails.weightedLevelScore` | number (0–10) |
| 9 | FINAL SCORE | `finalScore` | number (0–100) |
| 10 | MATURITY | `maturity` | string (hyperlink) |

**Conditional formatting on the summary sheet:**
- Column F (LEVEL 1 SCORE): font colour `#FF0000` (red) if value < 60.
- Column K (MATURITY): fill `#FFFFE0`, font colour `#800000` (dark red), bold.

---

### 8.4 Project Sheet Summary Block

Written to range `A1:B11` on each per-project sheet:

| Row | Col A (label) | Col B (value) |
|---|---|---|
| 1 | ID | Response ID |
| 2 | PROJECT | Project name |
| 3 | REVIEW DATE | `MM-DD-YYYY` formatted date |
| 4 | EMAIL | Respondent email |
| 5 | RESOURCE COUNT | Team size |
| 6 | LEVEL 1 SCORE | Weighted score (0–60); **red** font if < 60 |
| 7 | LEVEL 2 SCORE | Weighted score (0–20) |
| 8 | LEVEL 3 SCORE | Weighted score (0–10) |
| 9 | LEVEL 4 SCORE | Weighted score (0–10) |
| 10 | FINAL SCORE | Composite score (0–100) |
| 11 | MATURITY | M1 / M2 / M3 / M4; **light-yellow** fill, **dark-red bold** font, double-border |

**Navigation link (B12):**
- Text: `"Click here to go to DEMaturitySummary sheet"`
- Hyperlink target: `DEMaturitySummary!A1`
- Styling: bold, colour `#7C3606`, yellow fill (`#E1D70F`), centred.

---

## 9. Error Handling

The add-in uses a two-level error handling strategy:

### Level 1 — Excel.run catch

```javascript
Excel.run(async context => {
  // ... all logic
}).catch(function(error) {
  var errormessage = "There is an error in processing...";
  if (error.code == "ItemNotFound") {
    errormessage = "Please run add-in on DE Survey responses";
  }
  Wait.CloseDialogAsync(function() {
    MessageBox.Show(errormessage, "Error", MessageBoxButtons.OkOnly,
      MessageBoxIcons.Error, false, null,
      function(button) { MessageBox.CloseDialogAsync(function(){}) },
      false);
  });
});
```

| `error.code` | User message |
|---|---|
| `"ItemNotFound"` | "Please run add-in on DE Survey responses" |
| Any other code | "There is an error in processing project responses. Please try again or later" |

The spinner is always dismissed before showing the error dialog.

### Level 2 — Outer try/catch

```javascript
try {
  // Wait.Show + Excel.run
} catch (error) {
  console.error(error);
}
```

Any error thrown outside `Excel.run` (e.g., during `Wait.Show`) is silently logged to the console to prevent an unhandled promise rejection.

---

## 10. Office JS API Dependencies

| API | Requirement Set | Usage |
|---|---|---|
| `Office.onReady` | Office JS 1.1 | Add-in initialisation |
| `Excel.run` | ExcelApi 1.1 | All Excel operations |
| `context.workbook.worksheets` | ExcelApi 1.1 | Sheet enumeration, add, delete |
| `worksheet.tables.add` | ExcelApi 1.1 | Create tables |
| `table.getHeaderRowRange()` | ExcelApi 1.1 | Read/write headers |
| `table.getDataBodyRange()` | ExcelApi 1.1 | Read response data |
| `table.rows.add` | ExcelApi 1.1 | Append rows |
| `range.values` | ExcelApi 1.1 | Read/write cell values |
| `range.format.font.*` | ExcelApi 1.1 | Font colour, bold |
| `range.format.fill.color` | ExcelApi 1.1 | Cell background colour |
| `range.format.borders.*` | ExcelApi 1.1 | Cell border styles |
| `range.hyperlink` | ExcelApi 1.1 | Intra-workbook hyperlinks |
| `column.filter.apply` | ExcelApi 1.2 | Auto-filter on Response column |
| `range.format.autofitColumns()` | **ExcelApi 1.2** | Auto-size columns |
| `range.format.autofitRows()` | **ExcelApi 1.2** | Auto-size rows |
| `range.merge` | ExcelApi 1.1 | Merge heading cells |
| `worksheet.activate()` | ExcelApi 1.1 | Bring DEMaturitySummary to front |
| `Office.context.requirements.isSetSupported` | Office JS 1.1 | ExcelApi 1.2 capability check |
| `OfficeExtension.config.extendedErrorLogging` | Office JS | Enhanced error messages |
