# DE Maturity Calculator

A Microsoft Excel Add-in that calculates the Data Engineering (DE) maturity level of a project based on survey responses.

## Overview

The DE Maturity Calculator reads project survey responses from an Excel workbook and produces a maturity score (M1, M2, or M3) for each project. It automatically generates:

- A **DEMaturitySummary** sheet containing scores and maturity levels for all projects.
- Individual **project sheets** breaking down scores by Level 1, Level 2, and Level 3 questions.

## Maturity Levels

| Level | Final Score |
|-------|-------------|
| M1    | ≤ 70        |
| M2    | > 70 and ≤ 90 |
| M3    | > 90        |

The final score is a weighted sum of three level scores:

- **Level 1** — weighted at 70%
- **Level 2** — weighted at 20%
- **Level 3** — weighted at 10%

Response values are mapped to scores as follows:

| Response      | Score |
|---------------|-------|
| Always / Yes / NA | 10 |
| Frequently    | 7     |
| Sometimes     | 4     |
| Never / No / Don't Know | 0 |

## Prerequisites

- [Node.js](https://nodejs.org) (LTS version recommended)
- Microsoft Excel (desktop or online)
- An Excel workbook containing a sheet named **Form1** with a table named **Table1** populated with DE survey responses

## Getting Started

1. **Install dependencies**

   ```bash
   npm install
   ```

2. **Start the development server and sideload the add-in**

   ```bash
   npm start
   ```

   This command starts a local HTTPS development server and sideloads the add-in into Excel.

3. **Run the calculator**

   - Open the target Excel workbook that contains the survey responses.
   - In Excel, navigate to the **Home** tab and click **Show DEMaturityCalculator** in the DE Group.
   - In the task pane, click the **Calculate Maturity** button.

4. **Stop the add-in**

   ```bash
   npm stop
   ```

## Project Structure

```
DEMaturityCalculator/
├── assets/               # Icons and images used by the add-in
├── src/
│   ├── commands/         # Add-in command handlers
│   └── taskpane/
│       ├── taskpane.html # Task pane UI
│       ├── taskpane.js   # Core calculation logic
│       ├── taskpane.css  # Task pane styles
│       ├── dialogs.html  # Dialog UI
│       └── dialogs.js    # Dialog helpers (officejs.dialogs)
├── manifest.xml          # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Available Scripts

| Script | Description |
|--------|-------------|
| `npm start` | Start the dev server and sideload the add-in |
| `npm run build` | Build for production |
| `npm run build:dev` | Build for development |
| `npm run lint` | Check code style |
| `npm run lint:fix` | Auto-fix code style issues |
| `npm run validate` | Validate the add-in manifest |
| `npm run stop` | Stop the dev server and remove the sideloaded add-in |

## Debugging

The add-in supports the following debugging techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Provider

Developed by [Neudesic](https://www.neudesic.com).

## License

MIT
