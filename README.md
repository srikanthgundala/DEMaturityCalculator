# DEMaturityCalculator

A Microsoft Excel Office Add-in that calculates the **Data Engineering (DE) Maturity Level** of projects based on survey responses collected in an Excel workbook.

## Overview

The DEMaturityCalculator add-in processes project survey responses stored in an Excel table and generates a detailed maturity assessment for each project. It produces:

- A **DEMaturitySummary** sheet with an overview of all projects and their maturity levels.
- Individual **project sheets** with detailed question-level breakdowns for each level, highlighting areas that need improvement.

## Maturity Levels

Maturity is determined by a **weighted final score** calculated from three levels of questions:

| Level | Weight | Description |
|-------|--------|-------------|
| Level 1 | 70% | Core DE practices |
| Level 2 | 20% | Advanced DE practices |
| Level 3 | 10% | Cutting-edge DE practices |

Final maturity scores map to the following grades:

| Score Range | Maturity Grade |
|-------------|---------------|
| ≤ 70 | M1 |
| 71 – 90 | M2 |
| > 90 | M3 |

### Response Scoring

Survey responses are scored as follows:

| Response | Score |
|----------|-------|
| NA / Always / Yes | 10 |
| Frequently | 7 |
| Sometimes | 4 |
| Never / No / Don't Know | 0 |

## Prerequisites

- [Node.js](https://nodejs.org/) (v12 or later)
- [npm](https://www.npmjs.com/) (comes with Node.js)
- Microsoft Excel (desktop or online)
- An Excel workbook with:
  - A sheet named **Form1**
  - A table named **Table1** containing the DE survey responses

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
   cd DEMaturityCalculator
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

## Usage

### Running the Add-in (Development)

Start the add-in in development mode (launches Excel and sideloads the add-in automatically):

```bash
npm start
```

To target a specific platform:

```bash
npm run start:desktop   # Excel desktop
npm run start:web       # Excel Online
```

### Stopping the Add-in

```bash
npm stop
```

### Using the Add-in in Excel

1. Open an Excel workbook containing a **Form1** sheet with a **Table1** table of DE survey responses.
2. From the **HOME** tab in Excel, click **Show DEMaturityCalculator** in the **DE Group**.
3. In the task pane that appears, click **Run** to process the responses.
4. The add-in will generate:
   - A **DEMaturitySummary** sheet listing all projects with their maturity levels and scores.
   - Individual project sheets with level-by-level question breakdowns (failing responses are highlighted in red).

## Building

Build for production:

```bash
npm run build
```

Build for development:

```bash
npm run build:dev
```

Watch mode (rebuilds automatically on file changes):

```bash
npm run watch
```

## Project Structure

```
DEMaturityCalculator/
├── assets/                  # Add-in icons and images
├── src/
│   ├── commands/            # Add-in ribbon command handlers
│   └── taskpane/            # Task pane UI and main logic
│       ├── taskpane.html    # Task pane HTML
│       ├── taskpane.js      # Core maturity calculation logic
│       ├── taskpane.css     # Task pane styles
│       ├── dialogs.html     # Dialog helper UI
│       └── dialogs.js       # Dialog helper logic
├── manifest.xml             # Office Add-in manifest
├── package.json             # Node.js project metadata and scripts
└── webpack.config.js        # Webpack build configuration
```

## Debugging

The add-in supports several debugging approaches:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Linting

Check code for lint issues:

```bash
npm run lint
```

Auto-fix lint issues:

```bash
npm run lint:fix
```

## Validating the Manifest

```bash
npm run validate
```

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on contributing to this project.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Office Add-in samples at OfficeDev on GitHub](https://github.com/officedev)
- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/overview/excel?view=excel-js-preview)
