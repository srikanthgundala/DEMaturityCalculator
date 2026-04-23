# DEMaturityCalculator

**DEMaturityCalculator** is a Microsoft Excel Office Add-in developed by [Neudesic](https://www.neudesic.com) that calculates the **Data Engineering (DE) Maturity Level** of a project based on survey responses.

## Overview

The add-in reads DE survey responses from an Excel workbook, scores them across three maturity levels, and produces a detailed summary along with per-project breakdown sheets. It helps teams quickly identify gaps and understand where a project stands on the DE maturity scale (M1 → M2 → M3).

## Features

- Reads survey responses from a structured Excel table (`Table1` on the `Form1` sheet)
- Calculates weighted scores across **Level 1**, **Level 2**, and **Level 3** questions
- Generates a **DEMaturitySummary** sheet with scores and maturity ratings for all projects
- Creates individual **project sheets** with detailed question/response tables, highlighting non-compliant answers in red
- Applies response filters to surface only failing responses (`Never`, `Sometimes`, `Frequently`, `No`, `Don't Know`)
- Adds hyperlinks between the summary sheet and each project sheet for easy navigation

## Maturity Levels

| Score Range | Maturity Level |
|-------------|---------------|
| ≤ 70        | **M1**        |
| 71 – 90     | **M2**        |
| > 90        | **M3**        |

### Score Weightings

| Level   | Weight | Description                        |
|---------|--------|------------------------------------|
| Level 1 | 70%    | Core / foundational DE practices   |
| Level 2 | 20%    | Intermediate DE practices          |
| Level 3 | 10%    | Advanced DE practices              |

### Response Scores

| Response                        | Score |
|---------------------------------|-------|
| Always / Yes / NA               | 10    |
| Frequently                      | 7     |
| Sometimes                       | 4     |
| Never / No / Don't Know         | 0     |

## Prerequisites

- [Node.js](https://nodejs.org/) (v12 or later)
- [npm](https://www.npmjs.com/)
- Microsoft Excel (desktop or online)
- Office Add-in developer tools: `office-addin-debugging`, `office-addin-dev-certs`

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

3. **Start the development server and sideload the add-in**

   ```bash
   npm start
   ```

   This starts a local HTTPS server on port 3000 and opens Excel with the add-in sideloaded.

4. **Open your DE survey workbook** – the workbook must have a sheet named `Form1` with a table named `Table1` containing the survey responses.

5. **Click "Calculate Maturity"** in the add-in task pane to generate the results.

## Available Scripts

| Script              | Description                                       |
|---------------------|---------------------------------------------------|
| `npm start`         | Start debugging (sideloads add-in in Excel)       |
| `npm run build`     | Production build                                  |
| `npm run build:dev` | Development build                                 |
| `npm run watch`     | Webpack watch mode (development)                  |
| `npm run lint`      | Check code style with office-addin-lint           |
| `npm run lint:fix`  | Auto-fix lint issues                              |
| `npm stop`          | Stop the debugging session                        |
| `npm run validate`  | Validate the add-in manifest                      |

## Project Structure

```
DEMaturityCalculator/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Add-in task pane UI
│   │   ├── taskpane.js        # Maturity calculation logic
│   │   ├── taskpane.css       # Task pane styles
│   │   └── dialogs.js         # Dialog utilities
│   └── commands/
│       ├── commands.html      # Add-in commands page
│       └── commands.js        # Ribbon command handlers
├── assets/                    # Icons and images
├── manifest.xml               # Office Add-in manifest
├── webpack.config.js          # Webpack configuration
└── package.json
```

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines.

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.
