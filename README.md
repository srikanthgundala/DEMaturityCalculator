# DEMaturityCalculator

A Microsoft Excel Add-in that calculates the **Digital Engineering (DE) Maturity Level** of projects based on survey responses. Developed by [Neudesic](https://www.neudesic.com).

## Overview

DEMaturityCalculator processes DE survey responses stored in an Excel workbook and automatically:

- Calculates weighted maturity scores across three levels (Level 1, Level 2, Level 3)
- Assigns a final maturity classification (**M1**, **M2**, or **M3**) to each project
- Generates a `DEMaturitySummary` sheet with scores and hyperlinks for all projects
- Creates individual project sheets (named `<ProjectName>_<ID>`) that break down question-level responses and highlight failing areas

## Maturity Levels

| Maturity | Final Score Range | Description |
|----------|------------------|-------------|
| M1 | ≤ 70 | Initial / foundational DE practices |
| M2 | 71 – 90 | Developing / intermediate DE practices |
| M3 | > 90 | Advanced / mature DE practices |

### Score Weighting

| Level | Weightage |
|-------|-----------|
| Level 1 | 70% |
| Level 2 | 20% |
| Level 3 | 10% |

### Response Scores

| Response | Score |
|----------|-------|
| Always / Yes / NA | 10 |
| Frequently | 7 |
| Sometimes | 4 |
| Never / No / Don't Know | 0 |

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)
- Microsoft Excel (desktop or Excel on the web)
- A DE survey responses workbook with a sheet named **Form1** containing a table named **Table1**

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
   cd DEMaturityCalculator
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Generate and trust development certificates:
   ```bash
   npx office-addin-dev-certs install
   ```

## Usage

### Start the development server

```bash
npm start
```

This command starts the webpack dev server and sideloads the add-in into Excel.

### Build for production

```bash
npm run build
```

### Validate the manifest

```bash
npm run validate
```

### Stop the add-in

```bash
npm stop
```

## Running the Add-in

1. Open your DE survey responses Excel workbook (the workbook must contain a sheet named **Form1** with a table named **Table1**).
2. In Excel, navigate to the **Home** tab and click **Show DEMaturityCalculator** in the DE Group.
3. In the task pane, click **Calculate Maturity**.
4. The add-in will process all project responses and generate:
   - A **DEMaturitySummary** sheet summarising all projects with scores and maturity classifications.
   - Individual project sheets listing Level 1, Level 2, and Level 3 questions with responses, with failing items highlighted in red.

## Debugging

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Project Structure

```
DEMaturityCalculator/
├── assets/                  # Icons and images
├── src/
│   ├── commands/            # Add-in command handlers
│   └── taskpane/            # Task pane UI and main logic
│       ├── taskpane.html
│       ├── taskpane.js      # Core maturity calculation logic
│       └── taskpane.css
├── manifest.xml             # Office Add-in manifest
├── webpack.config.js        # Webpack build configuration
└── package.json
```

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on the code of conduct and the process for submitting pull requests.

## Questions and Comments

Send feedback or questions via the [Issues](../../issues) section of this repository.

For general Office JavaScript API questions, visit [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) and tag with `[office-js]`.

## Additional Resources

- [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/overview/office)
- [Neudesic](https://www.neudesic.com)

## License

MIT — see [LICENSE](LICENSE) for details.
