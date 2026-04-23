# DEMaturityCalculator

**DEMaturityCalculator** is a Microsoft Excel task pane add-in that calculates the **Data Engineering (DE) maturity level** of a project. It reads project survey responses stored in an Excel workbook, scores them across three maturity levels, and generates a formatted summary report directly inside Excel.

## Overview

Organizations use this add-in to evaluate how mature their data engineering practices are. After collecting survey responses via a Microsoft Form (exported to Excel), the add-in processes each project's answers and assigns one of three maturity ratings:

| Rating | Score Range | Description |
|--------|-------------|-------------|
| **M1** | ≤ 70        | Initial / foundational maturity |
| **M2** | 71 – 90     | Intermediate maturity           |
| **M3** | > 90        | Advanced / optimized maturity   |

## How It Works

1. Survey responses are stored in a sheet named **Form1** inside an Excel table called **Table1**.
2. Questions are grouped into three weighted levels:
   - **Level 1** – 70% weight (core DE practices)
   - **Level 2** – 20% weight (intermediate DE practices)
   - **Level 3** – 10% weight (advanced DE practices)
3. Each response is scored (`NA`/`Always` = 10, `Frequently` = 7, `Sometimes` = 4, `Never`/`No`/`Don't Know` = 0).
4. Clicking **Calculate Maturity** in the task pane triggers the calculation and produces:
   - A **DEMaturitySummary** sheet with a summary table (one row per project).
   - Individual **project sheets** detailing the question-level scores for each project.

## Prerequisites

- Microsoft Excel (desktop or online)
- Node.js ≥ 12
- npm ≥ 6
- A trusted SSL certificate for local development (handled by `office-addin-dev-certs`)

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Start the development server

```bash
npm start
```

This will start the webpack dev server, sideload the add-in into Excel, and open the workbook.

### 3. Use the add-in

1. Open your Excel workbook that contains the survey responses in **Form1 → Table1**.
2. In the **Home** ribbon, click **Show DEMaturityCalculator**.
3. In the task pane, click **Calculate Maturity**.
4. The add-in will generate the **DEMaturitySummary** sheet and individual project sheets.

## Available Scripts

| Script | Description |
|--------|-------------|
| `npm start` | Start the add-in in Excel (desktop) with a development server |
| `npm run start:web` | Start the add-in in Excel Online |
| `npm run build` | Build for production |
| `npm run build:dev` | Build for development |
| `npm run lint` | Run the linter |
| `npm run lint:fix` | Auto-fix linting issues |
| `npm stop` | Stop the development server |
| `npm run validate` | Validate the add-in manifest |

## Project Structure

```
DEMaturityCalculator/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Task pane UI
│   │   ├── taskpane.js        # Core calculation logic
│   │   ├── taskpane.css       # Task pane styles
│   │   ├── dialogs.html       # Dialog helper HTML
│   │   └── dialogs.js         # Dialog helper library
│   └── commands/
│       └── commands.html      # Ribbon command entry point
├── assets/                    # Add-in icons
├── manifest.xml               # Office add-in manifest
├── webpack.config.js          # Webpack configuration
└── package.json
```

## Debugging

The add-in supports the following debugging approaches:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute to this project.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
