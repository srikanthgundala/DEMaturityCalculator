# DEMaturityCalculator

A Microsoft Excel Office Add-in built by [Neudesic](https://www.neudesic.com) that calculates the **DevOps Engineering (DE) Maturity Level** of a project based on survey responses.

## Overview

The DEMaturityCalculator add-in reads project survey responses from an Excel workbook and automatically computes a maturity score across three levels (Level 1, Level 2, and Level 3). It generates a summary sheet and individual project sheets with detailed breakdowns, highlighting areas that need improvement.

### Maturity Levels

| Level | Score Range | Description |
|-------|-------------|-------------|
| M1    | ≤ 70        | Initial / Foundational maturity |
| M2    | 71 – 90     | Managed / Intermediate maturity |
| M3    | > 90        | Optimized / Advanced maturity   |

### Scoring Criteria

Survey responses are scored as follows:

| Response     | Score |
|--------------|-------|
| Always / Yes / NA | 10 |
| Frequently   | 7     |
| Sometimes    | 4     |
| Never / No / Don't Know | 0 |

Level scores are weighted:

- **Level 1** – 70% weight
- **Level 2** – 20% weight
- **Level 3** – 10% weight

## Prerequisites

- Microsoft Excel (Desktop or Online)
- Node.js (v12 or later)
- npm

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

This command starts a local HTTPS development server and sideloads the add-in into Excel.

### 4. Using the add-in

1. Open the Excel workbook containing the DE Survey responses.
2. Ensure the responses are in a sheet named **`Form1`** with a table named **`Table1`**.
3. Click the **"Show DEMaturityCalculator"** button in the Excel ribbon (Home tab, DE Group).
4. Click **"Calculate Maturity"** in the task pane.
5. The add-in will generate:
   - A **`DEMaturitySummary`** sheet with an overview of all projects.
   - Individual project sheets with detailed Level 1, Level 2, and Level 3 question breakdowns.

## Build

To build the add-in for production:

```bash
npm run build
```

To build for development:

```bash
npm run build:dev
```

## Project Structure

```
DEMaturityCalculator/
├── assets/               # Icons and image assets
├── src/
│   ├── commands/         # Office ribbon command handlers
│   └── taskpane/
│       ├── taskpane.html # Task pane UI
│       ├── taskpane.js   # Core maturity calculation logic
│       ├── taskpane.css  # Task pane styles
│       ├── dialogs.html  # Dialog UI
│       └── dialogs.js    # Dialog helpers (officejs.dialogs)
├── manifest.xml          # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Debugging

This add-in supports several debugging approaches:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute to this project.

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.

## Support

For questions or issues, please open an issue in this repository or visit [Neudesic](https://www.neudesic.com).
