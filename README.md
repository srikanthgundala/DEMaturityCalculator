# DEMaturityCalculator

**DEMaturityCalculator** is a Microsoft Excel Office Add-in developed by [Neudesic](https://www.neudesic.com) that automates the calculation of Data Engineering (DE) Maturity levels for projects. It processes survey responses collected in an Excel workbook and produces a detailed maturity summary for each project.

## Features

- Reads project survey responses from an Excel table (`Table1` in the `Form1` sheet)
- Calculates maturity scores across three levels:
  - **Level 1** – Core DE practices (weighted at 70%)
  - **Level 2** – Intermediate DE practices (weighted at 20%)
  - **Level 3** – Advanced DE practices (weighted at 10%)
- Classifies each project into one of three maturity tiers:
  - **M1** – Final score ≤ 70
  - **M2** – Final score between 70 and 90
  - **M3** – Final score > 90
- Generates a **DEMaturitySummary** sheet with a summary table for all projects
- Creates individual project sheets with question-level detail and highlights failing responses in red
- Applies filters to surface non-compliant responses (Never, Sometimes, Frequently, No, Don't Know)
- Adds hyperlinks between the summary sheet and individual project sheets for easy navigation

## Prerequisites

- Microsoft Excel (desktop or online)
- [Node.js](https://nodejs.org/) and npm

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Start the development server

```bash
npm start
```

This command starts a local dev server and sideloads the add-in into Excel.

### 3. Build for production

```bash
npm run build
```

## Usage

1. Open an Excel workbook that contains the DE survey responses in a sheet named **Form1** with a table named **Table1**.
2. Open the add-in task pane via the **Home** tab → **DE Group** → **Show DEMaturityCalculator**.
3. Click the **Calculate Maturity** button.
4. The add-in will generate a **DEMaturitySummary** sheet and individual project sheets with maturity results.

## Project Structure

```
├── src/
│   ├── taskpane/        # Task pane UI and maturity calculation logic
│   └── commands/        # Add-in command handlers
├── assets/              # Icons and images
├── manifest.xml         # Office Add-in manifest
├── webpack.config.js    # Webpack build configuration
└── package.json
```

## Scoring

Responses are scored as follows:

| Response     | Score |
|--------------|-------|
| Always / Yes / NA | 10 |
| Frequently   | 7     |
| Sometimes    | 4     |
| Never / No / Don't Know | 0 |

A level score is computed as a weighted percentage of the total possible score for that level. The final maturity score is the sum of all three weighted level scores (out of 100).

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on how to contribute to this project.

## License

This project is licensed under the [MIT License](LICENSE).
