# DE Maturity Calculator – Excel Add-in

An **Excel Office Add-in** that reads DE (Data Engineering) survey responses
from a workbook and automatically calculates each project's maturity level,
producing both a high-level summary sheet and a per-project detail sheet.

---

## What It Does

When you click **Calculate Maturity** in the task pane the add-in:

1. Reads every survey-response row from `Form1 / Table1`.
2. Scores each answer against a predefined rubric (see [Scoring](#scoring)).
3. Computes a weighted score across three maturity levels (L1 / L2 / L3).
4. Assigns an overall maturity classification (M1, M2, or M3).
5. Creates a **DEMaturitySummary** sheet with one row per project.
6. Creates one **detail sheet** per project showing every question, the
   project's answer, and which responses need improvement (highlighted in red).
7. Adds navigation hyperlinks between the summary and detail sheets.

---

## Scoring

### Response Scores

| Response | Score |
|----------|-------|
| Always / Yes / NA | 10 |
| Frequently | 7 |
| Sometimes | 4 |
| Never / No / Don't Know | 0 |

### Maturity Level Weightings

Survey questions are divided into three groups corresponding to foundational,
intermediate, and advanced engineering practices:

| Level   | Weight | Description |
|---------|--------|-------------|
| Level 1 | 70 %   | Foundational practices – must be solid for any project |
| Level 2 | 20 %   | Intermediate practices |
| Level 3 | 10 %   | Advanced practices |

Each level's weighted contribution is calculated as:

```
weighted score = (sum of response scores / max possible score) × 100 × weight%
```

The **final score** is the sum of all three weighted contributions (max 100).

### Maturity Classification

| Final Score | Maturity |
|-------------|----------|
| ≤ 70        | **M1** – Foundational |
| 71 – 90     | **M2** – Intermediate |
| > 90        | **M3** – Advanced |

A project's Level 1 score is highlighted in **red** on both the summary and
detail sheets when it falls below 70 (out of 70 possible points), indicating
that the foundational practices are not yet in place.

---

## Source Layout

```
src/
  taskpane/
    taskpane.js   – Main add-in logic (scoring, sheet generation, formatting)
    taskpane.html – Task pane UI
    taskpane.css  – Task pane styles
    dialogs.js    – Office JS dialogs library (Wait spinner, MessageBox, etc.)
    dialogs.html  – Dialog host page
  commands/
    commands.js   – Add-in command handler
    commands.html – Command host page
```

### Key Functions in `taskpane.js`

| Function | Purpose |
|----------|---------|
| `run()` | Main entry point – orchestrates the entire calculation and sheet-generation workflow |
| `getQuestions(questionsRow, levelIndexes)` | Builds an index→question-text dictionary for a set of level question indexes |
| `CalculateLevelScore(projectRow, responseScores, levelIndexes, weightage)` | Computes the weighted score and identifies failing responses for one maturity level |
| `applyFilter(filter)` | Filters a table's Response column to show only non-passing answers |
| `getJsDateFromExcel(dateValue)` | Converts an Excel date serial number to a `MM-DD-YYYY` string |
| `addLevelTable(...)` | Creates a Question / Response table on a project detail sheet and highlights failures |
| `applyLevelQuestionTopRowProperties(levelRange)` | Applies the blue-header style to a level's label row |

---

## Prerequisites

- Microsoft Excel (desktop or Excel Online)
- Node.js ≥ 14 and npm

## Getting Started

```bash
npm install
npm start        # builds the add-in and launches Excel with the add-in sideloaded
```

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## Questions and Comments

Questions about Microsoft Office 365 development in general should be posted to
[Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API).
If your question is about the Office JavaScript APIs, tag it with `office-js`.

## Additional Resources

- [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [OfficeDev samples on GitHub](https://github.com/officedev)

This project has adopted the
[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the
[Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact
[opencode@microsoft.com](mailto:opencode@microsoft.com).

## Copyright

Copyright (c) 2019 Microsoft Corporation. All rights reserved.
