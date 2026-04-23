# DEMaturityCalculator

**DEMaturityCalculator** is a Microsoft Excel Office Add-in that calculates the **Data Engineering (DE) Maturity** level for one or more projects based on survey responses collected in an Excel workbook.

## Overview

The add-in reads project survey responses from a structured Excel table, scores each response, and produces:

- A **DEMaturitySummary** sheet listing every project with its Level 1, Level 2, Level 3, and final maturity scores.
- Individual **per-project sheets** showing the detailed question/response breakdown for each maturity level, with failing responses highlighted in red.

## Maturity Levels

| Maturity | Final Score Range |
|----------|-------------------|
| **M1**   | ≤ 70              |
| **M2**   | 71 – 90           |
| **M3**   | > 90              |

### Scoring

Responses are mapped to numeric scores as follows:

| Response                | Score |
|-------------------------|-------|
| NA / Always             | 10    |
| Frequently              | 7     |
| Sometimes               | 4     |
| Never / No / Don't Know | 0     |

### Weighted Levels

| Level   | Weight |
|---------|--------|
| Level 1 | 70 %   |
| Level 2 | 20 %   |
| Level 3 | 10 %   |

The **final score** is the sum of the three weighted level scores.  
A Level 1 weighted score below **70** is flagged in red as it indicates a fundamental gap.

## Prerequisites

- **Microsoft Excel** (desktop or Excel on the web)
- **Node.js** ≥ 12 and **npm**

## Getting Started

### Install dependencies

```bash
npm install
```

### Start the add-in (development)

```bash
npm start
```

This sideloads the add-in into Excel and starts the webpack dev server.

### Build for production

```bash
npm run build
```

## Usage

1. Open the Excel workbook that contains the DE survey responses.  
   The responses must be in a sheet named **`Form1`** inside a table named **`Table1`**.
2. Open the **DEMaturityCalculator** task pane from the Excel ribbon.
3. Click **Calculate Maturity**.
4. The add-in will:
   - Create a **DEMaturitySummary** sheet with all projects and their maturity levels.
   - Create one sheet per project containing the full question/response detail for each maturity level.

## Project Structure

```
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html   # Task pane UI
│   │   ├── taskpane.js     # Main add-in logic
│   │   └── taskpane.css    # Styles
│   └── commands/           # Ribbon command handlers
├── assets/                 # Icons and images
├── manifest.xml            # Office Add-in manifest
└── webpack.config.js       # Webpack build configuration
```

## Debugging

The add-in supports the following debugging approaches:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please see [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines.

## License

MIT License — see [LICENSE](LICENSE) for details.
