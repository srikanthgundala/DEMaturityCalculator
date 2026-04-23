# DEMaturityCalculator

A Microsoft Excel Office Add-in that calculates the Data Engineering (DE) Maturity level of projects. The add-in processes project survey responses and generates a comprehensive maturity summary with scores across multiple levels.

## Overview

DEMaturityCalculator reads project responses from an Excel workbook and evaluates them against a defined set of Data Engineering maturity criteria organized into three levels:

- **Level 1** – Foundational practices (weighted at 70%)
- **Level 2** – Intermediate practices
- **Level 3** – Advanced practices

Each question is scored based on the response given (e.g., *Always*, *Frequently*, *Sometimes*, *Never*, *Yes*, *No*, *NA*), and a final maturity score is computed for every project.

## Features

- Reads project survey responses from a structured Excel table (`Form1` / `Table1`)
- Calculates Level 1, Level 2, and Level 3 maturity scores per project
- Generates a `DEMaturitySummary` sheet with a summary table covering:
  - Project ID, name, review date, email, resource count
  - Level 1, Level 2, and Level 3 scores
  - Final weighted maturity score and maturity label
- Creates individual project sheets with detailed breakdowns of question-level results
- Highlights failing questions to make remediation straightforward

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)
- Microsoft Excel (desktop or online)
- An Office 365 subscription or a standalone version of Excel that supports Office Add-ins

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Start the development server and sideload the add-in

```bash
npm start
```

This command starts the local development server and sideloads the add-in in Excel automatically.

### 3. Use the add-in

1. Open or create an Excel workbook that contains a sheet named **Form1** with a table named **Table1** holding the project survey responses.
2. Click the **Show DEMaturityCalculator** button in the **Home** tab ribbon.
3. In the task pane, click **Calculate Maturity**.
4. The add-in will process all project responses and generate the **DEMaturitySummary** sheet along with individual project sheets.

## Available Scripts

| Script | Description |
|---|---|
| `npm start` | Starts the dev server and sideloads the add-in in Excel |
| `npm run build` | Builds the add-in for production |
| `npm run build:dev` | Builds the add-in for development |
| `npm run stop` | Stops the debugging session |
| `npm run validate` | Validates the add-in manifest |
| `npm run lint` | Runs the linter |

## Project Structure

```
DEMaturityCalculator/
├── assets/                 # Add-in icons and images
├── src/
│   ├── commands/           # Ribbon command handlers
│   └── taskpane/
│       ├── taskpane.html   # Task pane UI
│       ├── taskpane.css    # Task pane styles
│       └── taskpane.js     # Core maturity calculation logic
├── manifest.xml            # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Debugging

The add-in supports the following debugging approaches:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute to this project.

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.
