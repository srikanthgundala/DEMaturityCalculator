# DEMaturityCalculator

A Microsoft Excel Office Add-in that calculates the **Data Engineering (DE) Maturity Level** of a project based on survey responses.

## Overview

The DEMaturityCalculator add-in processes DE survey responses stored in an Excel workbook and produces a maturity assessment for each project. It evaluates responses across three weighted levels and assigns a final maturity rating of **M1**, **M2**, or **M3**.

## Maturity Levels

| Maturity | Final Score |
|----------|-------------|
| M1       | ≤ 70        |
| M2       | 71 - 90     |
| M3       | > 90        |

## Scoring Model

Responses are scored as follows:

| Response                    | Score |
|-----------------------------|-------|
| Always / Yes / NA           | 10    |
| Frequently                  | 7     |
| Sometimes                   | 4     |
| Never / No / Don't Know     | 0     |

The final score is a weighted sum of three level scores:

- **Level 1** - 70% weight (core DE practices)
- **Level 2** - 20% weight (advanced practices)
- **Level 3** - 10% weight (leading-edge practices)

> **Note:** A Level 1 score below 70 is highlighted in red as a critical threshold indicator.

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)
- Microsoft Excel (Desktop or Online)
- An Excel workbook containing a sheet named **Form1** with a table named **Table1** populated with DE survey responses

## Getting Started

### Install dependencies

```bash
npm install
```

### Start the add-in (development)

```bash
npm start
```

This command starts the local dev server and sideloads the add-in into Excel.

### Build for production

```bash
npm run build
```

### Stop the add-in

```bash
npm stop
```

## How to Use

1. Open the Excel workbook that contains the DE survey responses (sheet: **Form1**, table: **Table1**).
2. Click the **Show DEMaturityCalculator** button in the Excel ribbon (HOME tab → DE Group).
3. Click **Calculate Maturity** in the task pane.
4. The add-in will:
   - Create a **DEMaturitySummary** sheet with a summary table showing scores and maturity levels for all projects.
   - Create an individual sheet for each project showing level-by-level question responses and highlighted failures.
   - Apply filters to surface only non-perfect responses for easy review.

## Project Structure

```
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Task pane UI
│   │   ├── taskpane.js        # Core maturity calculation logic
│   │   ├── taskpane.css       # Task pane styles
│   │   ├── dialogs.html       # Dialogs UI
│   │   └── dialogs.js         # OfficeJS.dialogs integration
│   └── commands/
│       ├── commands.html      # Add-in commands page
│       └── commands.js        # Ribbon command handlers
├── assets/                    # Icons and images
├── manifest.xml               # Office Add-in manifest
├── webpack.config.js          # Webpack build configuration
└── package.json
```

## Debugging

The add-in supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Excel JavaScript API reference](https://docs.microsoft.com/javascript/api/excel)
- [Neudesic](https://www.neudesic.com) - Developed by Neudesic

## License

MIT License. See [LICENSE](LICENSE) for details.
