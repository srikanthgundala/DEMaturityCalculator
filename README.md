# DEMaturityCalculator

**DEMaturityCalculator** is a Microsoft Excel add-in that automates the calculation of Data Engineering (DE) maturity levels for projects. It processes survey responses stored in a workbook and generates a detailed maturity assessment report across three maturity levels.

## Overview

The add-in reads DE survey response data from the active Excel workbook, evaluates responses against weighted scoring criteria across three maturity levels, and produces:

- A **DEMaturitySummary** sheet with a summary table showing each project's scores and overall maturity level.
- Individual **project sheets** with detailed breakdowns of level-by-level question scores and failure analysis.

## Features

- Calculates Level 1, Level 2, and Level 3 maturity scores for each project.
- Applies configurable weightage per level (Level 1: 70%, Level 2: 20%, Level 3: 10%).
- Generates formatted Excel output with color-coded maturity indicators.
- Provides drill-down tables showing which questions caused level failures.
- Displays a hyperlink to a detailed assessment report for each project.

## Prerequisites

- Microsoft Excel (Desktop or Web)
- Node.js (v12 or later)
- npm

## Getting Started

### Install dependencies

```bash
npm install
```

### Run in development mode

```bash
npm start
```

This command starts the development server and sideloads the add-in into Excel.

### Build for production

```bash
npm run build
```

### Validate the manifest

```bash
npm run validate
```

## Usage

1. Open the Excel workbook containing DE survey responses.
2. Go to the **Home** tab and click **Show DEMaturityCalculator** in the ribbon.
3. In the task pane, click **Calculate Maturity**.
4. The add-in will process all project responses and create:
   - A **DEMaturitySummary** sheet with aggregated scores.
   - A dedicated sheet for each project with detailed level breakdowns.

## Project Structure

```
├── assets/              # Add-in icons and images
├── src/
│   ├── commands/        # Ribbon command handlers
│   └── taskpane/        # Task pane UI and core calculation logic
│       ├── taskpane.html
│       ├── taskpane.js  # Main maturity calculation logic
│       └── taskpane.css
├── manifest.xml         # Office Add-in manifest
├── webpack.config.js    # Build configuration
└── package.json
```

## Debugging

The add-in can be debugged using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
