# DEMaturityCalculator

A Microsoft Excel Office Add-in built by [Neudesic](https://www.neudesic.com) that calculates the Data Engineering (DE) Maturity level of a project based on survey responses collected in an Excel workbook.

## Overview

The DEMaturityCalculator add-in reads project survey responses from an Excel table, scores each response across three maturity levels, and generates a summary sheet alongside individual project detail sheets — giving stakeholders a clear, data-driven view of where each project stands on the DE Maturity scale.

## Features

- Reads survey responses from a structured Excel table (`Table1` on the `Form1` worksheet)
- Scores responses using a weighted, three-level maturity model
- Generates a **DEMaturitySummary** sheet with a consolidated view across all projects
- Creates individual project sheets with a detailed breakdown of scores per question
- Highlights Level 1 scores that fall below the minimum threshold (70) in red
- Assigns a final maturity rating — **M1**, **M2**, or **M3** — based on the weighted overall score

## Maturity Model

### Scoring

Responses are mapped to numeric scores:

| Response | Score |
|----------|-------|
| Always / Yes | 10 |
| NA (Not Applicable) | 10 |
| Frequently | 7 |
| Sometimes | 4 |
| Never / No / Don't Know | 0 |

> **Note:** Responses marked **NA** (Not Applicable) receive the maximum score of 10 because the question does not apply to the project. The project is not penalized for practices that are out of scope.

### Maturity Levels and Weights

| Level | Weight | Description |
|-------|--------|-------------|
| Level 1 | 70% | Core DE practices (foundational questions) |
| Level 2 | 20% | Advanced DE practices |
| Level 3 | 10% | Expert DE practices |

> **Note:** The **weight** (e.g., 70% for Level 1) determines how much each level contributes to the **Final Score**. This is separate from the **minimum threshold** of 70 points that a project's Level 1 score must reach before the score is highlighted in red on the summary sheet.

### Final Maturity Rating

| Final Score | Maturity Rating |
|-------------|-----------------|
| ≤ 70 | M1 |
| 71 – 90 | M2 |
| > 90 | M3 |

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)
- Microsoft Excel (Desktop or Excel on the web)
- A survey responses workbook with:
  - A worksheet named **`Form1`**
  - A table named **`Table1`** containing the project survey data

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
cd DEMaturityCalculator
```

### 2. Install dependencies

```bash
npm install
```

### 3. Start the development server

```bash
npm start
```

This command starts the webpack development server and sideloads the add-in into Excel automatically.

### 4. Build for production

```bash
npm run build
```

The production bundle is output to the `dist/` folder.

## Usage

1. Open the Excel workbook that contains your project survey responses on the **Form1** sheet (in a table named **Table1**).
2. In Excel, open the **DEMaturityCalculator** add-in from the **Home** tab → **DE Group** → **Show DEMaturityCalculator**.
3. In the task pane, click **Calculate Maturity**.
4. The add-in will:
   - Create (or refresh) a **DEMaturitySummary** sheet with scores for all projects.
   - Create individual project sheets with a detailed question-level breakdown.

## Project Structure

```
DEMaturityCalculator/
├── assets/                  # Icons and images
├── src/
│   ├── commands/            # Add-in command handlers
│   └── taskpane/
│       ├── taskpane.html    # Task pane UI
│       ├── taskpane.css     # Task pane styles
│       ├── taskpane.js      # Core calculation logic
│       ├── dialogs.html     # Dialog UI
│       └── dialogs.js       # Dialog helper library
├── manifest.xml             # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Available Scripts

| Script | Description |
|--------|-------------|
| `npm start` | Start development server and sideload the add-in in Excel |
| `npm run build` | Build production bundle |
| `npm run build:dev` | Build development bundle |
| `npm run lint` | Run linter checks |
| `npm run lint:fix` | Auto-fix linting issues |
| `npm run validate` | Validate the Office Add-in manifest |
| `npm run stop` | Stop the development server |

## Debugging

The add-in supports the following debugging approaches:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute to this project.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Questions and Comments

For questions or feedback, open an issue in the [Issues](../../issues) section of this repository.

For general questions about Microsoft Office 365 development, visit [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) and tag your question with `[office-js]`.
