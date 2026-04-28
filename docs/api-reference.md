# API Reference

This document describes the public functions exported from `src/taskpane/taskpane.js` and the key internal helper functions.

---

## Exported Functions

### `run()`

**File:** `src/taskpane/taskpane.js`

The main entry point triggered when the user clicks **Calculate Maturity** in the task pane.

```js
export async function run(): Promise<void>
```

**Behaviour:**

1. Shows a *Processing* wait dialog.
2. Opens an `Excel.run` context.
3. Reads questions from the header row of `Form1 > Table1`.
4. Reads all project response rows from the same table.
5. Deletes any previously generated result sheets (all sheets except `Form1` and sheets whose names start with `_`).
6. For each project response row:
   - Calls `CalculateLevelScore()` for each of the three levels.
   - Computes the final score and assigns a maturity (M1 / M2 / M3).
   - Appends a row to the `DEMaturitySummary` table.
   - Creates a project detail sheet.
   - Calls `addLevelTable()` for Level 1, Level 2, and Level 3.
7. Applies cross-sheet hyperlinks.
8. Activates the `DEMaturitySummary` sheet.
9. Closes the wait dialog.

**Error handling:** All errors are caught and displayed to the user via the `MessageBox` dialog. If the workbook does not contain `Form1 > Table1`, the user sees "Please run add-in on DE Survey responses".

---

## Internal Helper Functions

All helpers are defined inside the `Excel.run` callback in `run()` and are not exported.

---

### `getQuestions(questionsRow, levelIndexes)`

Returns a dictionary mapping column indexes to their question text.

| Parameter | Type | Description |
|---|---|---|
| `questionsRow` | `any[][]` | The 2-D values array from the header row range |
| `levelIndexes` | `number[]` | Array of column indexes for the desired maturity level |

**Returns:** `{ [index: number]: string }` — key is the column index, value is the question text.

---

### `CalculateLevelScore(projectRow, responseScores, levelIndexes, weightage)`

Calculates the weighted score for a single maturity level for one project response row.

| Parameter | Type | Description |
|---|---|---|
| `projectRow` | `any[]` | Single row of values from `Table1` |
| `responseScores` | `{ [response: string]: number }` | Map of response text → numeric score |
| `levelIndexes` | `number[]` | Column indexes belonging to this level |
| `weightage` | `number` | Level weight as a percentage (70, 20, or 10) |

**Returns:**

```js
{
  levelFailures: { [index: number]: number },   // indexes where score < 10
  unWeightedLevelScore: number,                  // raw sum of scores
  unWeightedLevelPercentage: number,             // percentage of maximum possible
  weightedLevelScore: number                     // weighted contribution to final score
}
```

**Score mapping used by callers:**

| Response text | Score |
|---|---|
| `"Always"` | 10 |
| `"Yes"` | 10 |
| `"NA"` | 10 |
| `"Frequently"` | 7 |
| `"Sometimes"` | 4 |
| `"Never"` | 0 |
| `"No"` | 0 |
| `"Don't Know"` | 0 |

---

### `addLevelTable(questions, levelFailures, projectSheet, levelIndexes, level, projectResponse, rowIndex, context)`

Writes a Question / Response table for one maturity level onto a project detail sheet.

| Parameter | Type | Description |
|---|---|---|
| `questions` | `{ [index: number]: string }` | Question dictionary from `getQuestions()` |
| `levelFailures` | `{ [index: number]: number }` | Failing column indexes from `CalculateLevelScore()` |
| `projectSheet` | `Excel.Worksheet` | Target worksheet object |
| `levelIndexes` | `number[]` | Column indexes for this level |
| `level` | `string` | Level name, e.g. `"Level1"` (used to name the table) |
| `projectResponse` | `any[]` | Single row of response values |
| `rowIndex` | `number` | Starting row (1-based) on the sheet |
| `context` | `Excel.RequestContext` | Current Excel request context |

**Returns:** `number` — the next available row index after the table.

The table is named `<level>Table<responseId>` (e.g. `Level1Table42`). Rows with failing responses are coloured red (`#FF0000`).

---

### `applyFilter(filter)`

Applies a value filter to a column so that only non-passing responses are visible.

| Parameter | Type | Description |
|---|---|---|
| `filter` | `Excel.Filter` | The column filter object |

Filtered values: `"Frequently"`, `"Sometimes"`, `"Never"`, `"No"`, `"Don't Know"`.

---

### `getJsDateFromExcel(dateValue)`

Converts an Excel date serial number to a `MM-DD-YYYY` string.

| Parameter | Type | Description |
|---|---|---|
| `dateValue` | `number` | Excel date serial number (days since 1900-01-01) |

**Returns:** `string` in the format `MM-DD-YYYY`.

---

### `applyLevelQuestionTopRowProperties(levelRange)`

Applies heading formatting to the merged header cell above a level table.

| Parameter | Type | Description |
|---|---|---|
| `levelRange` | `Excel.Range` | A merged two-column range |

Applied styles:
- Merged cells
- Bold white text
- Blue background (`#154CC5`)
- Centre-aligned

---

## Response Score Constants

Defined locally inside `run()`:

```js
const RESPONSE_MAX_SCORE = 10;   // maximum score per question
const LEVEL_MIN_THRESHOLD = 70;  // Level 1 score threshold; below this → red highlight
```

---

## Maturity Thresholds

| Condition | Maturity |
|---|---|
| `finalScore <= 70` | `"M1"` |
| `finalScore > 70 && finalScore <= 90` | `"M2"` |
| `finalScore > 90` | `"M3"` |
