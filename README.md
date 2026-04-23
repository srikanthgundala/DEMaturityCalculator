# DEMaturityCalculator

**DEMaturityCalculator** is a Microsoft Excel Add-in developed by [Neudesic](https://www.neudesic.com) that calculates the Data Engineering (DE) maturity level of projects based on survey responses collected via Microsoft Forms (or any compatible survey tool).

## Overview

The add-in reads project survey responses from an Excel workbook, scores each project across three maturity levels, and generates:

- A **DEMaturitySummary** sheet with an at-a-glance summary of all projects.
- Individual **project sheets** with detailed question-by-question breakdowns and scores for each maturity level.

## Maturity Levels

Projects are scored across three levels, each with a different weight:

| Level | Weight | Description |
|-------|--------|-------------|
| Level 1 | 70% | Foundational DE practices |
| Level 2 | 20% | Intermediate DE practices |
| Level 3 | 10% | Advanced DE practices |

The weighted scores are summed to produce a **Final Score** (0–100), which maps to a maturity rating:

| Final Score | Maturity Rating |
|-------------|-----------------|
| ≤ 70        | M1 (Basic)      |
| 71 – 90     | M2 (Intermediate) |
| > 90        | M3 (Advanced)   |

## Response Scoring

Survey responses are mapped to numeric scores as follows:

| Response     | Score |
|--------------|-------|
| Always / Yes / NA | 10 |
| Frequently   | 7     |
| Sometimes    | 4     |
| Never / No / Don't Know | 0 |

## Prerequisites

- [Node.js](https://nodejs.org) (v12 or later)
- Microsoft Excel (desktop or Excel Online)
- A survey response workbook containing a sheet named **Form1** with a table named **Table1**

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Start the development server

```bash
npm start
```

This will launch Excel with the add-in sideloaded automatically.

### 3. Build for production

```bash
npm run build
```

## Usage

1. Open the Excel workbook that contains your DE survey responses in the **Form1** sheet.
2. Click the **Show DEMaturityCalculator** button in the Excel ribbon (Home tab → DE Group).
3. In the task pane, click **Calculate Maturity**.
4. The add-in will:
   - Create a **DEMaturitySummary** sheet listing all projects with their scores and maturity ratings.
   - Create an individual sheet for each project showing Level 1, Level 2, and Level 3 question responses (failures highlighted in red).
   - Add hyperlinks between the summary sheet and individual project sheets for easy navigation.

## Project Structure

```
DEMaturityCalculator/
├── assets/              # Icons and images used by the add-in
├── src/
│   └── taskpane/
│       ├── taskpane.html    # Task pane UI
│       ├── taskpane.js      # Core add-in logic
│       └── taskpane.css     # Task pane styles
├── manifest.xml         # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Available Scripts

| Command | Description |
|---------|-------------|
| `npm start` | Start the add-in in Excel (desktop) |
| `npm run start:web` | Start the add-in in Excel Online |
| `npm run build` | Build for production |
| `npm run build:dev` | Build for development |
| `npm run lint` | Run linter |
| `npm run validate` | Validate the manifest |

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## License

This project is licensed under the [MIT License](LICENSE).

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines.

## Questions and Comments

For questions or feedback, please open an issue in the *Issues* section of this repository.
