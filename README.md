# DEMaturityCalculator

A Microsoft Excel Office Add-in that calculates the **Data Engineering (DE) Maturity Level** of projects based on survey responses collected in an Excel workbook.

## Overview

The DEMaturityCalculator add-in reads project survey responses from an Excel table (`Table1` on the `Form1` worksheet) and generates per-project maturity reports directly inside the workbook. Each project receives:

- A summary maturity score across three levels (Level 1, Level 2, Level 3).
- Individual worksheets with detailed responses, pass/fail status per question, and level scores.
- A consolidated **DEMaturitySummary** sheet with hyperlinks to each project's detail sheet.

### Maturity Levels

| Level | Minimum Threshold | Description |
|-------|:-----------------:|-------------|
| Level 1 | 70% | Foundational DE practices |
| Level 2 | 70% | Intermediate DE practices |
| Level 3 | 70% | Advanced DE practices |

Response scoring:

| Response | Score |
|----------|------:|
| Always / Yes / NA | 10 |
| Frequently | 7 |
| Sometimes | 4 |
| Never / No / Don't Know | 0 |

## Prerequisites

- [Node.js](https://nodejs.org/) (v12 or later)
- Microsoft Excel (Desktop or Online)
- A DE survey responses workbook with a sheet named **Form1** containing a table named **Table1**

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Generate development certificates

```bash
npx office-addin-dev-certs install
```

### 3. Start the development server

```bash
npm start
```

This opens Excel with the add-in sideloaded automatically. You can also sideload the `manifest.xml` manually via **Insert → Add-ins → Upload My Add-in**.

### 4. Calculate maturity

1. Open your DE survey responses workbook in Excel.
2. Click the **Show DEMaturityCalculator** button on the **Home** ribbon.
3. Click **Calculate Maturity** in the task pane.
4. The add-in processes all project responses and generates individual project sheets along with a **DEMaturitySummary** sheet.

## Available Scripts

| Script | Description |
|--------|-------------|
| `npm start` | Start debugging (sideloads add-in in Excel desktop) |
| `npm run start:web` | Start debugging in Excel Online |
| `npm run build` | Production build |
| `npm run build:dev` | Development build |
| `npm run watch` | Webpack watch mode |
| `npm run validate` | Validate `manifest.xml` |
| `npm run lint` | Run linter |
| `npm stop` | Stop the debugging session |

## Project Structure

```
├── assets/               # Add-in icons and images
├── src/
│   ├── commands/         # Ribbon command handlers
│   └── taskpane/         # Task pane UI (HTML, CSS, JS) and dialog helpers
├── manifest.xml          # Office Add-in manifest
├── webpack.config.js     # Webpack configuration
└── package.json
```

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on how to contribute to this project.

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/overview/excel)
- [Neudesic](https://www.neudesic.com)

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
