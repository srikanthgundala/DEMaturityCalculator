# DEMaturityCalculator

A Microsoft Excel Add-in that calculates the **Data Engineering (DE) Maturity Level** of projects. The add-in reads project survey responses from an Excel workbook, scores them across three maturity levels, and produces a detailed summary with per-project drill-down sheets.

## Overview

The DEMaturityCalculator add-in processes project survey data submitted via a Microsoft Form (exported to Excel) and computes a weighted maturity score for each project. Based on the final score, each project is assigned one of three maturity classifications:

| Maturity | Final Score |
|----------|-------------|
| **M1**   | ≤ 70        |
| **M2**   | 71 – 90     |
| **M3**   | > 90        |

### Scoring Methodology

Responses are mapped to numeric scores:

| Response                  | Score |
|---------------------------|-------|
| Always / Yes / NA         | 10    |
| Frequently                | 7     |
| Sometimes                 | 4     |
| Never / No / Don't Know   | 0     |

The questions are divided into three maturity levels with the following weights:

- **Level 1** — 70% weight (foundational practices)
- **Level 2** — 20% weight (intermediate practices)
- **Level 3** — 10% weight (advanced practices)

### Output

After clicking **Calculate Maturity**, the add-in generates:

- **DEMaturitySummary** sheet — a summary table with ID, project name, review date, email, resource count, level scores, final score, and maturity rating for every project.
- **Per-project sheets** — individual sheets for each project with detailed level-by-level breakdowns and filtered response tables (Level1Table, Level2Table, Level3Table).

## Prerequisites

- [Node.js](https://nodejs.org/) (v12 or later)
- Microsoft Excel (Desktop or Online)
- A Microsoft 365 subscription (for sideloading Office Add-ins)

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Start the development server

```bash
npm start
```

This command starts the webpack dev-server and sideloads the add-in into Excel automatically.

### 3. Build for production

```bash
npm run build
```

The production bundle is written to the `dist/` directory.

## Input Format

The add-in expects an Excel workbook with:

- A sheet named **`Form1`**
- A table named **`Table1`** on that sheet containing project survey responses exported from Microsoft Forms

Each row in `Table1` represents one project's survey submission.

## Project Structure

```
DEMaturityCalculator/
├── assets/               # Add-in icons and images
├── src/
│   ├── commands/
│   │   ├── commands.html
│   │   └── commands.js   # Office ribbon command handlers
│   └── taskpane/
│       ├── taskpane.html # Task pane UI
│       ├── taskpane.js   # Core maturity calculation logic
│       ├── taskpane.css  # Task pane styles
│       ├── dialogs.html  # Dialog UI
│       └── dialogs.js    # Dialog helpers
├── manifest.xml          # Office Add-in manifest
├── webpack.config.js     # Webpack build configuration
└── package.json
```

## Available Scripts

| Script               | Description                                   |
|----------------------|-----------------------------------------------|
| `npm start`          | Start dev server and sideload add-in in Excel |
| `npm run build`      | Build production bundle                       |
| `npm run build:dev`  | Build development bundle                      |
| `npm run watch`      | Watch for changes and rebuild                 |
| `npm run validate`   | Validate the Office Add-in manifest           |
| `npm run lint`       | Run the linter                                |
| `npm run lint:fix`   | Auto-fix lint issues                          |

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on the code of conduct and the process for submitting pull requests.

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

## Copyright

Copyright (c) Neudesic. All rights reserved.
