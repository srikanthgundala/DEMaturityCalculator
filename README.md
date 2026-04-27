# DEMaturityCalculator

**DEMaturityCalculator** is a Microsoft Excel Task Pane Add-in that automates the calculation of Data Engineering (DE) maturity levels for projects based on survey responses.

## Overview

The add-in reads DE survey response data from an Excel workbook, evaluates answers across three maturity levels, computes weighted scores, and generates a summary sheet along with individual project detail sheets. Each project is assigned a maturity rating of **M1**, **M2**, or **M3** based on its overall final score.

## Features

- Reads survey responses from the `Form1` worksheet (`Table1`)
- Calculates weighted scores across three DE maturity levels:
  - **Level 1** (weight: 70%) — foundational practices
  - **Level 2** (weight: 20%) — intermediate practices
  - **Level 3** (weight: 10%) — advanced practices
- Computes a final maturity score and assigns a maturity rating:
  - **M1**: Final score ≤ 70
  - **M2**: Final score > 70 and ≤ 90
  - **M3**: Final score > 90
- Generates a `DEMaturitySummary` sheet with a summary table for all projects
- Creates individual project sheets with detailed question/response breakdowns
- Highlights failing responses in red for easy identification
- Adds navigation hyperlinks between the summary and project sheets

## Prerequisites

- [Node.js](https://nodejs.org/) (v12 or later)
- Microsoft Excel (Desktop or Online)
- A DE Survey responses workbook with a `Form1` sheet containing a `Table1` response table

## Getting Started

### Install dependencies

```bash
npm install
```

### Build the add-in

For production:

```bash
npm run build
```

For development:

```bash
npm run build:dev
```

### Start the add-in (with sideloading)

```bash
npm start
```

This command starts the local dev server and sideloads the add-in into Excel.

### Stop the add-in

```bash
npm stop
```

## Usage

1. Open the DE Survey responses workbook in Excel.
2. Open the **DEMaturityCalculator** task pane from the **Home** tab ribbon.
3. Click the **Run** button in the task pane.
4. The add-in will process all project responses and generate:
   - A `DEMaturitySummary` sheet with scores and maturity ratings for every project.
   - Individual project sheets with question-level details and highlighted failures.

## Project Structure

```
DEMaturityCalculator/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Task pane UI
│   │   └── taskpane.js        # Core add-in logic
│   └── commands/
│       └── commands.html      # Ribbon command handlers
├── assets/                    # Icons and static assets
├── manifest.xml               # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Debugging

The add-in supports the following debugging approaches:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on how to contribute to this project.

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/office)
- [Neudesic](https://www.neudesic.com)

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
