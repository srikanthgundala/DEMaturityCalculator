# DEMaturityCalculator

DEMaturityCalculator is a Microsoft Excel Office Add-in developed by [Neudesic](https://www.neudesic.com) that evaluates and scores the Data Engineering (DE) maturity level of a project. It reads survey responses from a Microsoft Form exported to an Excel workbook and produces a per-project maturity report with detailed question-level breakdowns.

## Features

- Processes survey responses stored in an Excel workbook (`Form1` sheet, `Table1` table).
- Calculates weighted scores across three maturity levels.
- Generates a **DEMaturitySummary** sheet with scores and maturity ratings for all projects.
- Creates individual project sheets with question-level detail and response highlighting.
- Adds navigation hyperlinks between the summary sheet and each project sheet.
- Automatically filters responses to highlight items that need improvement.

## Maturity Levels

| Level | Description | Weight |
|-------|-------------|--------|
| Level 1 | Foundation practices (58 questions) | 70% |
| Level 2 | Advanced practices (15 questions) | 20% |
| Level 3 | Expert practices (12 questions) | 10% |

### Final Maturity Rating

| Rating | Final Score |
|--------|-------------|
| M1 | ≤ 70 |
| M2 | > 70 and ≤ 90 |
| M3 | > 90 |

### Response Scoring

| Response | Score |
|----------|-------|
| Always / NA / Yes | 10 |
| Frequently | 7 |
| Sometimes | 4 |
| Never / No / Don't Know | 0 |

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)
- [npm](https://www.npmjs.com/)
- Microsoft Excel (desktop or Excel on the web)
- A Microsoft Form exported to Excel with the survey responses in a sheet named `Form1` and a table named `Table1`

## Getting Started

### 1. Clone the Repository

\`\`\`bash
git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
cd DEMaturityCalculator
\`\`\`

### 2. Install Dependencies

\`\`\`bash
npm install
\`\`\`

### 3. Start the Add-in (Desktop)

\`\`\`bash
npm start
\`\`\`

This command starts the local dev server on `https://localhost:3000` and sideloads the add-in into Excel on the desktop.

### 4. Start the Add-in (Excel on the Web)

\`\`\`bash
npm run start:web
\`\`\`

### 5. Build for Production

\`\`\`bash
npm run build
\`\`\`

## Usage

1. Open the Excel workbook that contains the DE survey responses exported from Microsoft Forms.
2. Ensure the responses are in a sheet named **Form1** and the table is named **Table1**.
3. Open the **DEMaturityCalculator** add-in task pane from the **Home** tab → **DE Group** → **Show DEMaturityCalculator**.
4. Click the **Calculate Maturity** button.
5. The add-in will:
   - Create a **DEMaturitySummary** sheet listing all projects with their Level 1, Level 2, Level 3, and Final scores, and the maturity rating.
   - Create individual project sheets with question-level detail. Responses that did not score the maximum are highlighted in red.
   - Add hyperlinks between the summary sheet and individual project sheets for easy navigation.

## Project Structure

\`\`\`
DEMaturityCalculator/
├── assets/                  # Icons and images
├── src/
│   ├── commands/            # Add-in command handlers
│   └── taskpane/
│       ├── taskpane.html    # Task pane UI
│       ├── taskpane.js      # Main calculation logic
│       ├── taskpane.css     # Task pane styles
│       ├── dialogs.html     # Dialog UI
│       └── dialogs.js       # Dialog helpers
├── manifest.xml             # Office Add-in manifest
├── package.json
├── tsconfig.json
└── webpack.config.js
\`\`\`

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on the code contribution process and guidelines.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/office)
- [Stack Overflow – office-js](http://stackoverflow.com/questions/tagged/office-js+API)
