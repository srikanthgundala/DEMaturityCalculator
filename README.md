# DEMaturityCalculator

A Microsoft Excel task pane add-in built by [Neudesic](https://www.neudesic.com) that calculates the **Data Engineering (DE) Maturity Level** of projects based on survey responses collected in an Excel workbook.

## Overview

The DEMaturityCalculator add-in reads project survey responses from a structured Excel table and automatically:

- Calculates weighted maturity scores across three levels.
- Generates a **DEMaturitySummary** sheet with a consolidated view of all projects.
- Creates individual project sheets with detailed question-by-question breakdowns.
- Highlights failing areas (scores below threshold) in red for quick identification.
- Adds hyperlinks between the summary and individual project sheets for easy navigation.

## Maturity Levels

Each project is assigned one of the following maturity designations based on its final score:

| Designation | Final Score Range | Description     |
|-------------|-------------------|-----------------|
| **M1**      | ≤ 70              | Foundational    |
| **M2**      | 71 – 90           | Intermediate    |
| **M3**      | > 90              | Advanced        |

### Score Calculation

The final score is a weighted sum of three level scores:

| Level   | Weight | Questions Evaluated                           |
|---------|--------|-----------------------------------------------|
| Level 1 | 70%    | Core DE practices (foundational capabilities) |
| Level 2 | 20%    | Intermediate DE practices                    |
| Level 3 | 10%    | Advanced DE practices                         |

Each survey response is mapped to a numeric score:

| Response                | Score |
|-------------------------|-------|
| Always / NA             | 10    |
| Frequently              | 7     |
| Sometimes               | 4     |
| Never / No / Don't Know | 0     |

## Prerequisites

- [Node.js](https://nodejs.org/) (v12 or later)
- [npm](https://www.npmjs.com/) (v6 or later)
- Microsoft Excel (desktop or online)

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

   This command starts the webpack dev server and sideloads the add-in into Excel.

4. **Build for production**

   ```bash
   npm run build
   ```

## Usage

1. Open the Excel workbook that contains DE survey responses.  
   The workbook must have a sheet named **`Form1`** with a table named **`Table1`** whose columns follow the standard DE survey format.

2. In Excel, open the **DEMaturityCalculator** task pane from the **Home** tab → **DE Group** → **Show DEMaturityCalculator**.

3. Click **Calculate Maturity** in the task pane.

4. The add-in will:
   - Delete any previously generated project sheets (sheets not named `Form1` or starting with `_`).
   - Create a **DEMaturitySummary** sheet with scores and maturity designations for all projects.
   - Create individual project sheets (named `<ProjectName>_<ID>`) with per-level question tables, filtered to show only non-optimal responses.

## Project Structure

```
DEMaturityCalculator/
├── assets/                  # Add-in icons and images
├── src/
│   ├── commands/
│   │   ├── commands.html    # Commands page
│   │   └── commands.js      # Add-in command handlers
│   └── taskpane/
│       ├── taskpane.html    # Task pane UI
│       ├── taskpane.css     # Task pane styles
│       ├── taskpane.js      # Core maturity calculation logic
│       ├── dialogs.html     # Dialog helper page
│       └── dialogs.js       # Dialog helper library
├── manifest.xml             # Office add-in manifest
├── package.json
└── webpack.config.js
```

## Available Scripts

| Script              | Description                                   |
|---------------------|-----------------------------------------------|
| `npm start`         | Start dev server and sideload add-in in Excel |
| `npm run build`     | Build for production                          |
| `npm run build:dev` | Build for development                         |
| `npm run validate`  | Validate the add-in manifest                  |
| `npm run lint`      | Lint the source code                          |
| `npm run lint:fix`  | Auto-fix lint issues                          |
| `npm stop`          | Stop the development server                   |

## Debugging

The add-in supports several debugging approaches:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on the code of conduct and the process for submitting pull requests.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/overview/office)
- [Neudesic](https://www.neudesic.com)
