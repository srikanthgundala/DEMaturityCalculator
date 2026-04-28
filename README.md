# DEMaturityCalculator

**DEMaturityCalculator** is a Microsoft Excel Add-in built by Neudesic that calculates the Digital Engineering (DE) maturity level of a project. The add-in reads survey responses from an Excel workbook and produces a detailed maturity report across three progressive levels, resulting in a final maturity score (M1, M2, or M3).

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Maturity Scoring](#maturity-scoring)
- [Project Structure](#project-structure)
- [Development](#development)
- [Debugging](#debugging)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

The DEMaturityCalculator add-in automates the calculation of a project's DE maturity by:

1. Reading survey responses submitted via a Microsoft Forms-linked Excel table (`Form1` / `Table1`).
2. Scoring each response across three maturity levels (Level 1, Level 2, and Level 3).
3. Generating a **DEMaturitySummary** sheet summarising all projects with their scores and maturity level.
4. Creating a dedicated detail sheet for every project containing level-by-level question/response tables with failing answers highlighted in red.

---

## Prerequisites

| Requirement | Version |
|---|---|
| Node.js | ≥ 12 |
| npm | ≥ 6 |
| Microsoft Excel | 2016 / Microsoft 365 (desktop or online) |

---

## Getting Started

1. **Clone the repository**

   ```bash
   git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
   cd DEMaturityCalculator
   ```

2. **Install dependencies**

   ```bash
   npm install
   ```

3. **Start the development server**

   ```bash
   npm start
   ```

   This starts the local HTTPS development server on `https://localhost:3000` and sideloads the add-in into Excel.

4. **Build for production**

   ```bash
   npm run build
   ```

For more detailed setup instructions see [docs/getting-started.md](docs/getting-started.md).

---

## Usage

1. Open the Excel workbook that contains the DE survey responses (the workbook must have a sheet named **Form1** with a table named **Table1**).
2. In Excel, go to the **Home** tab and click **Show DEMaturityCalculator** in the **DE Group** ribbon section.
3. In the task pane that opens, click **Calculate Maturity**.
4. The add-in processes all project responses and creates:
   - A **DEMaturitySummary** sheet with an overview of every project.
   - Individual project detail sheets (e.g. `ProjectName_<ID>`) with level-by-level question/response tables.

---

## Maturity Scoring

Responses are mapped to numeric scores:

| Response | Score |
|---|---|
| Always / Yes / NA | 10 |
| Frequently | 7 |
| Sometimes | 4 |
| Never / No / Don't Know | 0 |

Scores are calculated across three weighted levels:

| Level | Weight | Description |
|---|---|---|
| Level 1 | 70% | Core DE practices |
| Level 2 | 20% | Intermediate DE practices |
| Level 3 | 10% | Advanced DE practices |

The **Final Score** is the sum of the three weighted level scores. A project is then assigned a maturity rating:

| Final Score | Maturity |
|---|---|
| ≤ 70 | **M1** — Initial |
| 71 – 90 | **M2** — Developing |
| > 90 | **M3** — Advanced |

A Level 1 weighted score below **70** is highlighted in red as a critical indicator.

---

## Project Structure

```
DEMaturityCalculator/
├── assets/                  # Icons and images used by the add-in
├── src/
│   ├── commands/            # Ribbon command handler (commands.js / commands.html)
│   └── taskpane/
│       ├── taskpane.html    # Task pane UI
│       ├── taskpane.css     # Task pane styles
│       ├── taskpane.js      # Core maturity calculation logic
│       ├── dialogs.html     # Office dialog host page
│       └── dialogs.js       # OfficeDev dialogs library wrapper
├── docs/                    # Project documentation
├── manifest.xml             # Office Add-in manifest
├── package.json
└── webpack.config.js
```

---

## Development

| Script | Description |
|---|---|
| `npm start` | Start the dev server and sideload in Excel |
| `npm run build` | Production build (minified) |
| `npm run build:dev` | Development build |
| `npm run watch` | Incremental development build with watch mode |
| `npm run lint` | Lint source files |
| `npm run lint:fix` | Auto-fix lint issues |
| `npm run validate` | Validate `manifest.xml` |
| `npm run stop` | Stop the debugging session |

---

## Debugging

The add-in can be debugged using the following methods:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

---

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on the contribution process.

---

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

© 2024 Neudesic. All rights reserved.
