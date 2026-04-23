# DEMaturityCalculator

DEMaturityCalculator is a Microsoft Excel Add-in built by [Neudesic](https://www.neudesic.com) that calculates the **Data Engineering (DE) Maturity level** of a project based on survey responses collected in an Excel workbook.

## Overview

The add-in reads project responses from a structured Excel table (`Form1 / Table1`), scores each response across three maturity levels, and generates a per-project summary sheet along with a consolidated `DEMaturitySummary` sheet. Each project is assigned a final maturity score and a maturity label:

| Final Score | Maturity Label |
|-------------|---------------|
| ≤ 70        | M1            |
| 70 – 90     | M2            |
| > 90        | M3            |

### Response Scoring

| Response      | Score |
|---------------|-------|
| Always / Yes / NA | 10 |
| Frequently    | 7     |
| Sometimes     | 4     |
| Never / No / Don't Know | 0 |

## Prerequisites

- [Node.js](https://nodejs.org/) (v12 or later)
- Microsoft Excel (desktop or online)
- [office-addin-debugging](https://www.npmjs.com/package/office-addin-debugging) (installed via `npm install`)

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

3. **Start the add-in** (launches Excel and sideloads the add-in automatically)

   ```bash
   npm start
   ```

   To start against the desktop app specifically:

   ```bash
   npm run start:desktop
   ```

   To start against Excel on the web:

   ```bash
   npm run start:web
   ```

4. **Stop the add-in**

   ```bash
   npm stop
   ```

## Build

| Command | Description |
|---------|-------------|
| `npm run build` | Production build |
| `npm run build:dev` | Development build |
| `npm run watch` | Incremental development build (watch mode) |

## Usage

1. Open the Excel workbook that contains the DE Maturity survey responses in a sheet named **Form1** with a table named **Table1**.
2. In the **Home** tab, click **Show DEMaturityCalculator** in the *DE Group* section of the ribbon.
3. In the task pane, click **Calculate Maturity**.
4. The add-in will process all project responses and:
   - Create individual project summary sheets.
   - Populate the **DEMaturitySummary** sheet with scores and hyperlinks for each project.

## Project Structure

```
DEMaturityCalculator/
├── assets/              # Icons and images
├── src/
│   ├── commands/        # Add-in commands (ribbon buttons)
│   └── taskpane/        # Task pane UI and main calculation logic
│       ├── taskpane.html
│       ├── taskpane.js  # Core maturity calculation logic
│       ├── taskpane.css
│       ├── dialogs.html
│       └── dialogs.js
├── manifest.xml         # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Linting

```bash
npm run lint        # Check for lint errors
npm run lint:fix    # Auto-fix lint errors
```

## Validate Manifest

```bash
npm run validate
```

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/overview/excel?view=excel-js-preview)
- [Neudesic](https://www.neudesic.com)

## License

MIT
