# DEMaturityCalculator

DEMaturityCalculator is a Microsoft Excel Office Add-in built by [Neudesic](https://www.neudesic.com) that calculates the Data Engineering (DE) maturity level of a project based on survey responses.

## Overview

The add-in reads survey responses from a Microsoft Forms export in Excel, scores each project across three maturity levels, and generates a summary sheet along with individual project detail sheets. Each project is assigned a final maturity rating of **M1**, **M2**, or **M3**.

### Maturity Levels

| Rating | Final Score |
|--------|-------------|
| M1     | ≤ 70        |
| M2     | 71 – 90     |
| M3     | > 90        |

The final score is a weighted combination of three level scores:

- **Level 1** – 70% weight (foundational practices)
- **Level 2** – 20% weight (intermediate practices)
- **Level 3** – 10% weight (advanced practices)

Responses are scored as follows:

| Response                  | Score |
|---------------------------|-------|
| Always / Yes / NA         | 10    |
| Frequently                | 7     |
| Sometimes                 | 4     |
| Never / No / Don't Know   | 0     |

## Prerequisites

- [Node.js](https://nodejs.org/) (v12 or later)
- Microsoft Excel (Desktop or Excel Online)
- A DE survey responses Excel workbook exported from Microsoft Forms (sheet named **Form1**, table named **Table1**)

## Getting Started

### 1. Install dependencies

```bash
npm install
```

### 2. Start the development server

```bash
npm start
```

This will start the webpack dev server and sideload the add-in into Excel Desktop.

### 3. Use the add-in in Excel Online

```bash
npm run start:web
```

### 4. Build for production

```bash
npm run build
```

## Usage

1. Open the DE survey responses Excel workbook (exported from Microsoft Forms).
2. Launch the **DEMaturityCalculator** add-in from the **Home** tab in Excel.
3. Click **Calculate Maturity** in the task pane.
4. The add-in will:
   - Generate a **DEMaturitySummary** sheet with scores for every project.
   - Create one detail sheet per project showing Level 1, Level 2, and Level 3 question responses (failures highlighted in red).
   - Add hyperlinks between the summary sheet and each project sheet for easy navigation.

## Project Structure

```
├── assets/                  # Icons and images
├── src/
│   ├── commands/            # Add-in ribbon command handlers
│   └── taskpane/
│       ├── taskpane.html    # Task pane UI
│       ├── taskpane.js      # Core maturity calculation logic
│       ├── taskpane.css     # Task pane styles
│       ├── dialogs.html     # Dialog UI
│       └── dialogs.js       # Dialog helper (officejs.dialogs)
├── manifest.xml             # Office Add-in manifest
├── package.json
└── webpack.config.js
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on how to contribute to this project.

## License

MIT License. See [LICENSE](LICENSE) for details.
