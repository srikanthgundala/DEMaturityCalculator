# API / Function Reference — DEMaturityCalculator

This document describes every public and internal function in the DEMaturityCalculator add-in, with signatures, parameter descriptions, return values, and usage examples drawn directly from the source code.

---

## Table of Contents

- [Module: `taskpane.js`](#module-taskpanejs)
  - [`run()`](#run)
  - [`getQuestions()`](#getquestions)
  - [`CalculateLevelScore()`](#calculatelevelscore)
  - [`getJsDateFromExcel()`](#getjsdatefromexcel)
  - [`addLevelTable()`](#addleveltable)
  - [`applyFilter()`](#applyfilter)
  - [`applyLevelQuestionTopRowProperties()`](#applylevelquestiontoprowproperties)
- [Module: `commands.js`](#module-commandsjs)
  - [`action()`](#action)
  - [`getGlobal()`](#getglobal)
- [Data Structures](#data-structures)
  - [`LevelCalculationDetails`](#levelcalculationdetails)
  - [`ProjectSheetData`](#projectsheetdata)
  - [`responseScores`](#responsescores)
- [Constants](#constants)
- [Office.js APIs Used](#officejs-apis-used)

---

## Module: `taskpane.js`

**Source:** `src/taskpane/taskpane.js`

This is the primary module of the add-in. It exports one function (`run`) and defines helper functions in the inner scope of the `Excel.run()` callback.

---

### `run()`

**Type:** `async function`  
**Exported:** Yes  
**Registered as:** Click handler for `document.getElementById("run")` (the "Calculate Maturity" button)

#### Description

The top-level entry point for the entire maturity calculation pipeline. When invoked:

1. Displays a "Processing Project Responses" wait spinner.
2. Opens an `Excel.run()` context.
3. Reads the survey workbook (`Form1` / `Table1`).
4. Deletes all previously generated output sheets.
5. Builds question dictionaries for each maturity level.
6. Creates the `DEMaturitySummary` sheet.
7. Iterates over every project row and calculates Level 1, 2, and 3 weighted scores.
8. Writes per-project detail sheets with Q&A tables and formatting.
9. Applies column filters and adds bidirectional hyperlinks.
10. Activates the `DEMaturitySummary` sheet.
11. Closes the wait spinner.

On any Office.js error, closes the spinner and shows a `MessageBox` with a user-friendly error message.

#### Signature

```js
export async function run(): Promise<void>
```

#### Parameters

None. All inputs are read from the active Excel workbook via the Office.js API.

#### Returns

`Promise<void>`

#### Side Effects

| Resource | Action |
|---|---|
| All sheets except `Form1` and `_*` sheets | **Deleted** |
| `DEMaturitySummary` sheet | **Created** with `SummaryTable` |
| `{ProjectName}_{ID}` sheets | **Created** for each project row |
| Console | Errors logged via `console.error()` |

#### Error Handling

| Error Condition | User Message |
|---|---|
| `error.code === "ItemNotFound"` (no `Form1` or `Table1`) | "Please run add-in on DE Survey responses" |
| Any other `Excel.run()` error | "There is an error in processing project responses. Please try again or later" |

#### Example Usage

This function is wired up automatically in `Office.onReady()`:

```js
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
  }
});
```

---

### `getQuestions()`

**Type:** `function` (defined inside `Excel.run()` callback scope)  
**Visibility:** Internal

#### Description

Extracts question text from the header row of `Table1` for a specified set of column indexes. Returns a dictionary mapping each column index to its question text.

#### Signature

```js
function getQuestions(
  questionsRow: any[][],
  levelIndexes: number[]
): { [index: number]: string }
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| `questionsRow` | `any[][]` | The `.values` property of the header row range — a 2D array; the question text is in `questionsRow[0][index]` |
| `levelIndexes` | `number[]` | Array of column indexes to extract (e.g. `level1MaturityIndexes`) |

#### Returns

An object (dictionary) where:
- **Key** = column index (`number`)
- **Value** = question text (`string`) at that index in the header row

#### Example

```js
// Given a header row where index 7 = "Do you use version control?"
var questions = getQuestions(headerValues, [7, 8, 9]);
// questions → { 7: "Do you use version control?", 8: "...", 9: "..." }
```

---

### `CalculateLevelScore()`

**Type:** `function` (defined inside `Excel.run()` callback scope)  
**Visibility:** Internal

#### Description

Calculates the weighted maturity score for a single project row at one maturity level. Also identifies "failing" question indexes — those where the response was not a perfect score of 10.

#### Signature

```js
function CalculateLevelScore(
  projectRow: any[],
  responseScores: { [response: string]: number },
  levelIndexes: number[],
  weightage: number
): LevelCalculationDetails
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| `projectRow` | `any[]` | One row of values from `Table1`'s data body range |
| `responseScores` | `object` | Map of response string → numeric score (see [`responseScores`](#responsescores)) |
| `levelIndexes` | `number[]` | Column indexes for this level (e.g. `level1MaturityIndexes`) |
| `weightage` | `number` | The level weight: `70` for Level 1, `20` for Level 2, `10` for Level 3 |

#### Returns

A [`LevelCalculationDetails`](#levelcalculationdetails) object.

#### Scoring Algorithm

```
maxLevelScore      = levelIndexes.length × 10
rawScore           = Σ responseScores[projectRow[index]] for each index
                     (unknown/blank responses contribute 0)
levelPercentage    = (rawScore / maxLevelScore) × 100
weightedLevelScore = (levelPercentage × weightage) / 100  [toFixed(2), then parseFloat]
```

#### Example

```js
var level1Details = CalculateLevelScore(
  projectResponses.values[0],   // first project row
  responseScores,                // scoring map
  level1MaturityIndexes,         // 58 column indexes
  70                             // Level 1 weight = 70%
);

// level1Details.weightedLevelScore → e.g. 56.97
// level1Details.levelFailures     → { 12: 12, 18: 18, ... }  (imperfect columns)
```

---

### `getJsDateFromExcel()`

**Type:** `function` (defined inside `Excel.run()` callback scope)  
**Visibility:** Internal

#### Description

Converts an Excel serial date number (the internal numeric representation Excel uses for dates) into a formatted date string `MM-DD-YYYY`.

Excel serial dates count the number of days since January 0, 1900 (with a deliberate off-by-two correction for a historical Lotus 1-2-3 bug). The formula used is:

```
JS timestamp = (excelSerial - 25567 - 2) × 86400 × 1000
```

Where `25567` = days from 1900-01-01 to 1970-01-01 (Unix epoch), and `2` corrects for the Lotus leap-year bug.

#### Signature

```js
function getJsDateFromExcel(dateValue: number): string
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| `dateValue` | `number` | Excel serial date number (e.g. `44927` = 2023-01-01) |

#### Returns

`string` — Formatted date string in `MM-DD-YYYY` format with zero-padded month and day.

#### Example

```js
getJsDateFromExcel(44927)   // → "01-01-2023"
getJsDateFromExcel(45000)   // → "03-15-2023"
getJsDateFromExcel(44561)   // → "01-01-2022"
```

> **Note:** The function reads `projectResponses.values[i][2]` which is the "Review Date" column (0-based index 2 in the data body range).

---

### `addLevelTable()`

**Type:** `function` (defined inside `Excel.run()` callback scope)  
**Visibility:** Internal

#### Description

Creates a two-column Excel table ("Question" / "Response") on the given project worksheet for one maturity level. Rows with imperfect responses (present in `levelFailures`) are coloured red. Returns the updated row index so subsequent calls can position the next table correctly.

#### Signature

```js
function addLevelTable(
  questions: { [index: number]: string },
  levelFailures: { [index: number]: number },
  projectSheet: Excel.Worksheet,
  levelIndexes: number[],
  level: string,
  projectResponse: any[],
  rowIndex: number,
  context: Excel.RequestContext
): number
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| `questions` | `object` | Dictionary of `{ columnIndex → questionText }` for this level (output of `getQuestions()`) |
| `levelFailures` | `object` | Dictionary of `{ columnIndex → columnIndex }` for questions not scoring 10 (output of `CalculateLevelScore().levelFailures`) |
| `projectSheet` | `Excel.Worksheet` | The worksheet object on which to create the table |
| `levelIndexes` | `number[]` | Ordered array of column indexes for this level |
| `level` | `string` | Level identifier string: `"Level1"`, `"Level2"`, or `"Level3"` |
| `projectResponse` | `any[]` | The full project row array (to extract response values by index) |
| `rowIndex` | `number` | Starting Excel row number (1-based) where the table header should be placed |
| `context` | `Excel.RequestContext` | The active Excel.run context (not currently used in the function body but passed for future use) |

#### Returns

`number` — The next available row index after all table rows have been written. Used by the caller to position the next section.

#### Table naming convention

The table is named: `{level}Table{responseId}`  
Example: `Level1Table42`, `Level2Table42`, `Level3Table42`

This naming scheme is used later to look up the table for `applyFilter()`.

#### Formatting applied

- All data rows: black font (`#000000`)
- Rows where `levelFailures[index]` is defined: red font (`#FF0000`)
- Header row: styled by the caller using `applyLevelQuestionTopRowProperties()`

#### Example

```js
var rowIndex = 14;  // start writing Level 1 table at row 14

rowIndex = addLevelTable(
  level1Questions,
  level1CalculationDetails.levelFailures,
  projectSheet,
  level1MaturityIndexes,
  "Level1",
  projectResponses.values[i],
  rowIndex,
  context
);

// rowIndex is now 14 + 58 (number of Level 1 questions) = 72
```

---

### `applyFilter()`

**Type:** `function` (defined inside `Excel.run()` callback scope)  
**Visibility:** Internal

#### Description

Applies a value-based column filter to the "Response" column of a level Q&A table so that only imperfect responses are shown. Rows with perfect responses (`NA`, `Always`, `Yes`) are hidden by the filter.

#### Signature

```js
function applyFilter(filter: Excel.Filter): void
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| `filter` | `Excel.Filter` | The filter object obtained from `table.columns.getItem("Response").filter` |

#### Filter Configuration

```js
filter.apply({
  filterOn: Excel.FilterOn.values,
  values: ["Frequently", "Sometimes", "Never", "No", "Don't Know"]
});
```

Only rows where the Response column contains one of the five imperfect values are shown. Perfect responses (`NA`, `Always`, `Yes` = 10 points) are hidden.

#### Usage Pattern

```js
var level1filter = projectSheet
  .tables.getItem("Level1Table" + responseId)
  .columns.getItem("Response")
  .filter;

applyFilter(level1filter);
```

---

### `applyLevelQuestionTopRowProperties()`

**Type:** `function` (defined inside `Excel.run()` callback scope)  
**Visibility:** Internal

#### Description

Applies formatting to the section header row that labels each Q&A table block (e.g., "Level 1 Questions"). Merges two cells and applies bold, white, centred text on a blue background.

#### Signature

```js
function applyLevelQuestionTopRowProperties(levelRange: Excel.Range): void
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| `levelRange` | `Excel.Range` | A two-cell range spanning columns B and C on the project sheet (e.g. `"B13:C13"`) |

#### Formatting applied

| Property | Value |
|---|---|
| Merge | `true` (merges B and C) |
| Font weight | Bold |
| Font colour | `#FFFFFF` (white) |
| Horizontal alignment | Center |
| Fill colour | `#154CC5` (Neudesic blue) |

---

## Module: `commands.js`

**Source:** `src/commands/commands.js`

This module is the "function file" required by the Office Add-in manifest for ribbon command registration. It does not contain maturity calculation logic.

---

### `action()`

**Type:** `function`  
**Visibility:** Global (attached to `g.action` for manifest resolution)

#### Description

Handles the ribbon button command action event. Currently shows an informational notification message. This is a placeholder that could be extended to perform add-in command actions from the ribbon.

#### Signature

```js
function action(event: Office.AddinCommands.Event): void
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| `event` | `Office.AddinCommands.Event` | The command event object — **must** call `event.completed()` to signal completion |

#### Side Effects

- Calls `Office.context.mailbox.item.notificationMessages.replaceAsync("action", message)` to show an "Performed action." notification.
- Calls `event.completed()`.

> **Note:** The `mailbox` API is intended for Outlook add-ins. This function is a template remnant; it is not triggered during normal DEMaturityCalculator usage because the ribbon button uses a `ShowTaskpane` action, not a `ExecuteFunction` action.

---

### `getGlobal()`

**Type:** `function`  
**Visibility:** Module-private

#### Description

Returns a reference to the JavaScript global object in a cross-environment way (works in Service Workers, Web Workers, browser windows, and Node.js).

#### Signature

```js
function getGlobal(): Window | typeof globalThis | undefined
```

#### Returns

The global object: `self` (if defined), then `window`, then `global`, then `undefined`.

#### Usage

```js
const g = getGlobal();
g.action = action;  // makes action() accessible as a global for manifest lookup
```

---

## Data Structures

### `LevelCalculationDetails`

Returned by `CalculateLevelScore()`.

```ts
{
  levelFailures: {
    [columnIndex: number]: number   // map of failing column indexes (value === key)
  };
  unWeightedLevelScore: number;     // sum of raw response scores for this level
  unWeightedLevelPercentage: number;// raw score as a percentage of maximum possible
  weightedLevelScore: number;       // (levelPercentage × weightage) / 100, 2 decimal places
}
```

#### Example value

```js
{
  levelFailures: { 12: 12, 24: 24, 38: 38 },
  unWeightedLevelScore: 490,
  unWeightedLevelPercentage: 84.48,
  weightedLevelScore: 59.14          // 84.48 × 70 / 100
}
```

---

### `ProjectSheetData`

Collected during the first project loop pass and written in the second pass.

```ts
{
  Id: string | number;              // projectResponses.values[i][0]
  projectName: string;              // projectResponses.values[i][5]
  email: string;                    // projectResponses.values[i][3]
  reviewDate: string;               // MM-DD-YYYY string from getJsDateFromExcel()
  resourceCount: number | string;   // projectResponses.values[i][6]
  level1CalculationDetails: LevelCalculationDetails;
  level2CalculationDetails: LevelCalculationDetails;
  level3CalculationDetails: LevelCalculationDetails;
  finalScore: number;               // sum of three weightedLevelScores
  maturity: "M1" | "M2" | "M3";
  projectSheetName: string;         // generated sheet name (≤31 chars)
}
```

---

### `responseScores`

A lookup dictionary defined at the top of `Excel.run()` that maps survey response strings to numeric scores.

```js
var responseScores = {
  "NA":         10,
  "Always":     10,
  "Yes":        10,
  "Frequently":  7,
  "Sometimes":   4,
  "Never":       0,
  "No":          0,
  "Don't Know":  0
};
```

Any response string not present in this dictionary — including blank cells — is treated as **0** by `CalculateLevelScore()`.

---

## Constants

Defined inside `Excel.run()` in `taskpane.js`:

| Constant | Value | Meaning |
|---|---|---|
| `RESPONSE_MAX_SCORE` | `10` | The maximum score for a single question response |
| `LEVEL_MIN_THRESHOLD` | `70` | The minimum acceptable weighted score for Level 1; scores below this are highlighted red |

Maturity thresholds (used inline in `run()`):

| Threshold | Value | Maturity |
|---|---|---|
| `finalScore <= 70` | 70 | M1 |
| `finalScore > 70 && finalScore <= 90` | 70–90 | M2 |
| `finalScore > 90` | > 90 | M3 |

Level weight values (passed as `weightage` to `CalculateLevelScore()`):

| Level | Weight |
|---|---|
| Level 1 | `70` |
| Level 2 | `20` |
| Level 3 | `10` |

---

## Office.js APIs Used

The following Office JavaScript API objects and methods are called in `taskpane.js`:

| API Object / Method | Purpose |
|---|---|
| `Office.onReady(callback)` | Initialize add-in after Office.js is ready |
| `Office.HostType.Excel` | Guard: only activate UI when running in Excel |
| `Office.context.platform` | Detect Excel Online vs. desktop for `context.sync()` strategy |
| `Office.context.requirements.isSetSupported("ExcelApi", "1.2")` | Guard before calling autofit APIs |
| `OfficeExtension.config.extendedErrorLogging` | Enable verbose error detail in caught exceptions |
| `Excel.run(async context => {...})` | Open a tracked batch context for Excel operations |
| `context.workbook.worksheets` | Collection of all worksheets in the workbook |
| `worksheets.load("items/name")` | Queue a load of worksheet name properties |
| `worksheets.getItem(name)` | Get a specific worksheet by name |
| `worksheets.add(name)` | Create a new worksheet |
| `worksheet.delete()` | Delete a worksheet |
| `worksheet.activate()` | Make a worksheet the active/visible sheet |
| `worksheet.tables.add(address, hasHeaders)` | Create a named table on a worksheet |
| `worksheet.tables.getItem(name)` | Get a table by name |
| `worksheet.getRange(address)` | Get a range by A1 address |
| `worksheet.getUsedRange()` | Get the bounding range of all used cells |
| `table.getHeaderRowRange()` | Get the header row range of a table |
| `table.getDataBodyRange()` | Get the data rows range of a table |
| `table.rows.add(index, values)` | Append rows to a table |
| `table.rows.getItemAt(index)` | Get a table row by position |
| `table.columns.getItem(name)` | Get a table column by name |
| `column.filter` | The `Excel.Filter` object for that column |
| `filter.apply(criteria)` | Apply a filter with `FilterOn.values` criteria |
| `range.values` | Read or write cell values (2D array) |
| `range.format.font.color` | Set font colour (hex string) |
| `range.format.font.bold` | Set bold (boolean) |
| `range.format.fill.color` | Set fill/background colour (hex string) |
| `range.format.horizontalAlignment` | Set horizontal alignment |
| `range.format.borders.getItem(name)` | Get a border by name (EdgeTop, EdgeBottom, etc.) |
| `border.style` | Set border line style (e.g. `"Double"`) |
| `range.format.autofitColumns()` | Auto-fit column widths to content |
| `range.format.autofitRows()` | Auto-fit row heights to content |
| `range.getCell(row, col)` | Get a single cell within a range |
| `range.merge` | Merge cells in a range (set to `true`) |
| `range.hyperlink` | Set an intra-workbook or external hyperlink |
| `context.sync()` | Flush queued commands to the Excel host |

---

*Back to: [Architecture Overview](architecture.md) | [Getting Started](getting-started.md)*
