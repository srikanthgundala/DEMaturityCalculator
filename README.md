# DEMaturityCalculator

A Microsoft Excel Office Add-in built by [Neudesic](https://www.neudesic.com) that automates the calculation of Data Engineering (DE) Maturity levels for projects based on survey responses collected in an Excel workbook.

## Overview

The DEMaturityCalculator add-in reads project survey responses from a structured Excel sheet and scores them across three maturity levels. It then generates a summary sheet and individual project detail sheets, giving teams a clear view of their DE maturity standing.

### What it does

- Reads project survey responses from the `Form1` worksheet (table name: `Table1`)
- Evaluates responses across **Level 1**, **Level 2**, and **Level 3** maturity criteria
- Calculates weighted scores for each level and a final overall maturity score
- Creates a **DEMaturitySummary** sheet with a consolidated table (ID, Project, Review Date, Email, Resource Count, Level scores, Final Score, and Maturity)
- Creates individual per-project detail sheets listing all questions and responses, with level breakdowns and visual formatting

## Prerequisites

- Microsoft Excel (desktop or web)
- Node.js and npm
- A valid SSL certificate for local development (handled by `office-addin-dev-certs`)

## Getting Started

### Install dependencies

```bash
npm install
```

### Start the add-in (development mode)

```bash
npm start
```

This will start a local development server, sideload the add-in into Excel, and open it automatically.

To start targeting a specific platform:

```bash
# Desktop (default)
npm run start:desktop

# Web (Excel Online)
npm run start:web
```

### Build for production

```bash
npm run build
```

### Stop the add-in

```bash
npm stop
```

## Usage

1. Open an Excel workbook that contains a sheet named **Form1** with a table named **Table1**.
2. The table header row should contain the survey questions; each subsequent row should represent a project's responses.
3. Click the **Show DEMaturityCalculator** button in the Excel Home tab ribbon.
4. In the task pane, click **Calculate Maturity**.
5. The add-in will process the responses and produce:
   - A **DEMaturitySummary** sheet with aggregated scores for all projects.
   - Individual project sheets with detailed level-by-level breakdowns.

## Project Structure

```
DEMaturityCalculator/
├── assets/               # Add-in icons and images
├── src/
│   └── taskpane/
│       ├── taskpane.html # Task pane UI
│       ├── taskpane.css  # Task pane styles
│       └── taskpane.js   # Core add-in logic
├── manifest.xml          # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on how to contribute to this project.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
