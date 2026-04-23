# DE Maturity Calculator

**DEMaturityCalculator** is a Microsoft Excel Add-in (Office Add-in) built by [Neudesic](https://www.neudesic.com) that calculates the Data Engineering (DE) maturity level of a project based on survey responses.

## Overview

The add-in reads DE survey responses from an Excel workbook, evaluates answers across three maturity levels, computes weighted scores, and assigns an overall maturity rating (M1, M2, or M3) to each project. It then generates:

- A **DEMaturitySummary** sheet with a consolidated table of all project scores.
- Individual **project sheets** showing a breakdown of Level 1, Level 2, and Level 3 questions along with the responses and highlighted failures.

## Maturity Levels

| Rating | Final Score |
|--------|-------------|
| M1     | ≤ 70        |
| M2     | 71 – 90     |
| M3     | > 90        |

### Score Weights

| Level   | Weight |
|---------|--------|
| Level 1 | 70%    |
| Level 2 | 20%    |
| Level 3 | 10%    |

### Response Scoring

| Response                | Score |
|-------------------------|-------|
| Always / NA / Yes       | 10    |
| Frequently              | 7     |
| Sometimes               | 4     |
| Never / No / Don't Know | 0     |

## Prerequisites

- [Node.js](https://nodejs.org) (v12 or later)
- Microsoft Excel (desktop or Excel on the web)
- A DE survey responses workbook with a sheet named **Form1** containing a table named **Table1**

## Getting Started

1. **Install dependencies**

   ```bash
   npm install
   ```

2. **Start the development server and sideload the add-in**

   ```bash
   npm start
   ```

   This command starts the webpack dev server, generates dev certificates, and sideloads the add-in into Excel.

3. **Using the add-in**

   - Open your DE survey responses workbook in Excel.
   - Click **Show DEMaturityCalculator** in the **Home** tab ribbon.
   - In the task pane, click **Calculate Maturity**.
   - The add-in processes each project's responses and creates a **DEMaturitySummary** sheet plus an individual sheet per project.

## Build

```bash
npm run build
```

For a development build:

```bash
npm run build:dev
```

## Validate the Manifest

```bash
npm run validate
```

## Project Structure

```
├── assets/               # Add-in icons and images
├── src/
│   ├── commands/         # Ribbon command handlers
│   └── taskpane/
│       ├── taskpane.html # Task pane UI
│       ├── taskpane.js   # Core maturity calculation logic
│       └── taskpane.css  # Task pane styles
├── manifest.xml          # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/overview/office)
- [Neudesic](https://www.neudesic.com)

## Copyright

Copyright (c) Neudesic. All rights reserved.
