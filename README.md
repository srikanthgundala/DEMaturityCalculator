# DEMaturityCalculator

A Microsoft Excel Office Add-in developed by [Neudesic](https://www.neudesic.com) to calculate the **Data Engineering (DE) Maturity Level** of a project based on survey responses.

## Overview

The DEMaturityCalculator add-in processes DE survey responses stored in an Excel workbook and automatically generates:

- A **DEMaturitySummary** sheet with an aggregated summary across all projects
- Individual **per-project sheets** with detailed question-level breakdowns and maturity scores

Projects are classified into one of three maturity levels:

| Maturity Level | Final Score |
|---|---|
| **M1** | ≤ 70 |
| **M2** | 71 – 90 |
| **M3** | > 90 |

## Scoring Model

Survey questions are organized into three levels, each with a different weighting:

| Level | Questions | Weight |
|---|---|---|
| Level 1 | Core foundational practices | 70% |
| Level 2 | Intermediate practices | 20% |
| Level 3 | Advanced practices | 10% |

Each response is scored as follows:

| Response | Score |
|---|---|
| Always / Yes / NA | 10 |
| Frequently | 7 |
| Sometimes | 4 |
| Never / No / Don't Know | 0 |

The final score is the sum of the three weighted level scores (out of 100).

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)
- Microsoft Excel (Desktop or Online)
- An Excel workbook containing a sheet named **Form1** with a table named **Table1** holding the DE survey responses

## Getting Started

### Install dependencies

```bash
npm install
```

### Start the add-in (development)

```bash
npm start
```

This launches a local HTTPS dev server at `https://localhost:3000` and sideloads the add-in into Excel.

### Build for production

```bash
npm run build
```

### Validate the manifest

```bash
npm run validate
```

## Usage

1. Open the Excel workbook that contains the DE survey responses on a sheet named **Form1** (table **Table1**).
2. In Excel, navigate to the **Home** tab and click **Show DEMaturityCalculator** in the DE Group.
3. In the task pane, click **Calculate Maturity**.
4. The add-in will generate:
   - A **DEMaturitySummary** sheet listing all projects with their level scores, final score, and maturity rating.
   - One sheet per project showing the per-question responses with failing responses highlighted in red.

## Project Structure

```
DEMaturityCalculator/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Task pane UI
│   │   ├── taskpane.js        # Core maturity calculation logic
│   │   ├── taskpane.css       # Task pane styles
│   │   └── dialogs.js         # Dialog helpers (Wait, MessageBox, etc.)
│   └── commands/
│       ├── commands.html
│       └── commands.js
├── assets/                    # Add-in icons and images
├── manifest.xml               # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on how to contribute to this project.

## Questions and comments

Open an issue in this repository for feedback, bug reports, or feature requests.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
