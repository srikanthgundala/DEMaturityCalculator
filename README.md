# DEMaturityCalculator

> **A Microsoft Excel Add-in that evaluates the Data Engineering (DE) Maturity level of a project from survey responses — producing a weighted score, a classified maturity tier, and richly formatted summary and per-project worksheets, all without leaving Excel.**

Developed by **[Neudesic](https://www.neudesic.com)** · Office.js Task Pane Add-in · MIT License

---

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [Prerequisites](#prerequisites)
4. [Installation & Setup](#installation--setup)
5. [Getting Started — Sideloading the Add-in](#getting-started--sideloading-the-add-in)
6. [How to Use](#how-to-use)
7. [Maturity Scoring Details](#maturity-scoring-details)
8. [Output Sheets Reference](#output-sheets-reference)
9. [Project Structure](#project-structure)
10. [Development](#development)
11. [Debugging](#debugging)
12. [Contributing](#contributing)
13. [License](#license)

---

## Overview

DEMaturityCalculator is a Microsoft Excel Task Pane Add-in built with the [Office JavaScript API (Office.js)](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins). It reads structured Data Engineering survey responses from a table in the active workbook and computes a **Data Engineering Maturity Score** for every project in the dataset.

Survey questions are categorised across three maturity levels — foundational (Level 1), intermediate (Level 2), and advanced (Level 3) — each carrying a different weight in the final score. The add-in then classifies each project as **M1**, **M2**, or **M3** and writes:

* A **`DEMaturitySummary`** sheet that lists every project with its individual level scores, final score, and maturity classification.
* One **per-project sheet** containing the full breakdown of question responses for each maturity level, with non-optimal answers highlighted in red and low-scoring items pre-filtered for immediate attention.

All sheets include bidirectional hyperlinks so reviewers can navigate seamlessly between the summary and individual project detail views.

---

## Features

* **One-click analysis** — a single "Calculate Maturity" button processes every project row in the survey table.
* **Weighted three-level scoring** — Level 1 (70 %), Level 2 (20 %), Level 3 (10 %) contribute to a final score out of 100.
* **Automatic maturity classification** — M1 / M2 / M3 determined by configurable thresholds.
* **DEMaturitySummary sheet** — a named Excel Table (`SummaryTable`) with all projects, individual level scores, final score, and maturity rating.
* **Per-project detail sheets** — one sheet per survey response containing full question-response tables for all three levels.
* **Red highlighting** — Level 1 weighted scores below 70 are coloured red in the summary; individual failing responses are highlighted red on each project sheet.
* **Pre-applied column filters** — each per-project level table is filtered to show only non-perfect responses (`Frequently`, `Sometimes`, `Never`, `No`, `Don't Know`).
* **Bidirectional hyperlinks** — the summary PROJECT column links to the project sheet; the MATURITY cell links directly to the maturity cell on the project sheet; each project sheet has a link back to `DEMaturitySummary`.
* **Auto-fit columns and rows** — output sheets auto-size for readability (requires ExcelApi 1.2+).
* **Modal wait spinner** — a blocking "Processing…" dialog prevents accidental interaction during calculation.
* **User-friendly error handling** — displays a descriptive error dialog if the required sheet or table is missing.
* **Progressive enhancement** — gracefully detects `ExcelApi` requirement sets at runtime for broad compatibility.

---

## Prerequisites

| Requirement | Minimum Version | Notes |
|---|---|---|
| **Node.js** | 12 LTS or later | Required to install dependencies and run build scripts |
| **npm** | 6 or later | Bundled with Node.js |
| **Microsoft 365** | Current subscription | Required to run Office Add-ins in Excel |
| **Excel** | Desktop (Windows/macOS) **or** Excel on the Web | Both platforms are supported |
| **Git** | Any recent version | For cloning the repository |

> **Note:** A Microsoft 365 developer tenant (available free via the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)) is the recommended environment for testing and development.

---

## Installation & Setup

### 1. Clone the Repository

```bash
git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
cd DEMaturityCalculator
```

### 2. Install Dependencies

```bash
npm install
```

This installs all production and development dependencies, including:

* `officejs.dialogs` — modal dialog library (Wait spinner, MessageBox, Alert)
* Webpack 4 + Babel — bundling and ES5 transpilation
* `office-addin-debugging` — sideloading and debugging toolchain
* `office-addin-dev-certs` — local HTTPS certificate generation

### 3. Generate a Trusted Development Certificate

The add-in dev server runs on `https://localhost:3000`. Office requires HTTPS even for localhost. The toolchain handles certificate generation automatically when you run `npm start`, but you can also generate the certificate manually:

```bash
npx office-addin-dev-certs install
```

Follow the OS prompts to trust the self-signed certificate.

---

## Getting Started — Sideloading the Add-in

Sideloading installs the add-in from `manifest.xml` into your local Excel instance for testing without publishing to AppSource.

### Option A — Automated (Recommended)

Start the webpack dev server **and** sideload into Excel in one command:

```bash
# Auto-detect platform (Desktop or Web)
npm start

# Force Excel Desktop
npm run start:desktop

# Force Excel on the Web
npm run start:web
```

`npm start` will:
1. Start the webpack dev server at `https://localhost:3000`.
2. Register `manifest.xml` with your local Excel installation.
3. Launch Excel (Desktop) or open a browser (Web) with the add-in already loaded.

### Option B — Manual Sideloading

1. Run the dev server separately:

   ```bash
   npm run dev-server
   ```

2. Follow the [Microsoft sideloading guide](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) for your platform:
   * **Windows** — Insert → Add-ins → Manage My Add-ins → Upload My Add-in → browse to `manifest.xml`.
   * **macOS** — Insert → Add-ins → Upload My Add-in → browse to `manifest.xml`.
   * **Excel on the Web** — Insert → Add-ins → Upload My Add-in → browse to `manifest.xml`.

### Stopping the Add-in

```bash
npm stop
```

This stops the dev server and removes the sideloaded registration.

---

## How to Use

### Step 1 — Prepare the Survey Workbook

The add-in reads from a specific location in the active workbook. You must provide:

| Requirement | Value |
|---|---|
| **Sheet name** | `Form1` |
| **Table name** | `Table1` |
| **Table has a header row** | Yes |

The table columns must follow the layout below. Questions begin at **column index 7** (column H) and run through **column index 91** (column CP) for a total of **85 survey questions**:

| Column Index | Column Letter | Content |
|---|---|---|
| 0 | A | Response ID |
| 1 | B | *(reserved / additional metadata)* |
| 2 | C | Review Date *(Excel date serial number)* |
| 3 | D | Reviewer Email |
| 4 | E | *(reserved)* |
| 5 | F | Project Name |
| 6 | G | Resource Count |
| 7–91 | H–CP | Survey question responses (85 questions) |

Survey responses must be one of the recognised values listed in [Response Scoring](#response-scoring).

### Step 2 — Open the Add-in Task Pane

1. Open the Excel workbook that contains your `Form1 / Table1` data.
2. Go to the **Home** tab on the Excel ribbon.
3. Locate the **"DE Group"** ribbon group.
4. Click **"Show DEMaturityCalculator"** to open the task pane.

### Step 3 — Run the Calculation

1. In the task pane, click the **"Calculate Maturity"** button.
2. A "Processing Project Responses" wait dialog appears while the add-in:
   * Reads all question headers and response data from `Table1`.
   * Deletes any previously generated output sheets (all sheets except `Form1` and sheets whose names begin with `_`).
   * Creates a `DEMaturitySummary` sheet and one sheet per project.
   * Calculates maturity scores and writes formatted output.
   * Applies filters and hyperlinks.
3. When processing completes, the `DEMaturitySummary` sheet is activated automatically.

### Step 4 — Review Results

* Inspect **`DEMaturitySummary`** for an at-a-glance view of all projects.
* Click any hyperlinked **project name** in the `PROJECT` column to navigate to the detailed project sheet.
* On a project sheet, the three level tables are pre-filtered to show only responses that were not perfect (`Frequently`, `Sometimes`, `Never`, `No`, `Don't Know`) — the items most likely to require remediation.
* Click **"Click here to go to DEMaturitySummary sheet"** (cell B11 on any project sheet) to return to the summary.

---

## Maturity Scoring Details

### Overview

The final score is a **weighted composite** of three level scores. A perfect score equals **100**.

```
finalScore = weightedLevel1Score + weightedLevel2Score + weightedLevel3Score
```

### Response Scoring

Each survey response is mapped to a numeric score out of 10:

| Response Value | Score |
|---|---|
| `NA` | 10 |
| `Always` | 10 |
| `Yes` | 10 |
| `Frequently` | 7 |
| `Sometimes` | 4 |
| `Never` | 0 |
| `No` | 0 |
| `Don't Know` | 0 |

Unrecognised or blank responses contribute **0** to the level score.

### Level Score Formula

For each level:

```
levelRawScore      = sum of responseScore for every question in the level
levelMaxScore      = numberOfQuestionsInLevel × 10
levelPercentage    = (levelRawScore / levelMaxScore) × 100
weightedLevelScore = (levelPercentage × levelWeightage) / 100
```

### Level Definitions

| Level | Description | Question Count | Column Indexes | Weight | Max Weighted Score |
|---|---|---|---|---|---|
| **Level 1** | Foundational | 58 | 7–14, 18–24, 27–31, 33–38, 44–54, 59–64, 67–68, 72–76, 79–86 | **70 %** | 70.00 |
| **Level 2** | Intermediate | 15 | 15–17, 25, 32, 39, 55–57, 65–66, 69, 77, 87–88 | **20 %** | 20.00 |
| **Level 3** | Advanced | 12 | 26, 40–43, 58, 70–71, 78, 89–91 | **10 %** | 10.00 |
| **Total** | | **85** | | **100 %** | **100.00** |

> **Important:** Level 1 carries the majority of the weight (70 %). A project cannot achieve a high final score without a strong Level 1 performance. Level 1 weighted scores below **70** are flagged in red on the summary sheet.

### Maturity Classifications

| Classification | Condition | Interpretation |
|---|---|---|
| **M1** | `finalScore ≤ 70` | Foundational — significant gaps in core DE practices |
| **M2** | `70 < finalScore ≤ 90` | Intermediate — solid foundations with room to mature |
| **M3** | `finalScore > 90` | Advanced — consistently high DE practice adoption |

### Worked Example

Assume a project answers all Level 1 questions as `Frequently` (score 7), all Level 2 questions as `Always` (score 10), and all Level 3 questions as `Never` (score 0):

```
Level 1 raw   = 58 × 7  = 406   |  max = 580  |  pct = 70.0 %  |  weighted = 49.00
Level 2 raw   = 15 × 10 = 150   |  max = 150  |  pct = 100.0 % |  weighted = 20.00
Level 3 raw   = 12 × 0  = 0     |  max = 120  |  pct =   0.0 % |  weighted =  0.00
──────────────────────────────────────────────────────────────────────────────────────
Final Score   = 49.00 + 20.00 + 0.00 = 69.00  →  Maturity: M1
```

---

## Output Sheets Reference

### `DEMaturitySummary` Sheet

Contains a named Excel Table (`SummaryTable`) with one row per survey response:

| Column | Description |
|---|---|
| **ID** | Response ID (from column A of `Table1`) |
| **PROJECT** | Project name — hyperlinked to the project detail sheet |
| **REVIEW DATE** | Review date (converted from Excel date serial to `MM-DD-YYYY`) |
| **EMAIL** | Reviewer email address |
| **RESOURCE COUNT** | Number of resources on the project |
| **LEVEL 1 SCORE** | Weighted Level 1 score (out of 70); shown in **red** if < 70 |
| **LEVEL 2 SCORE** | Weighted Level 2 score (out of 20) |
| **LEVEL 3 SCORE** | Weighted Level 3 score (out of 10) |
| **FINAL SCORE** | Sum of all weighted level scores (out of 100) |
| **MATURITY** | M1 / M2 / M3 — hyperlinked to the maturity cell on the project sheet; yellow background, dark-red bold text |

### Per-Project Detail Sheets

Sheet naming convention: `{projectName_first25AlphanumChars}_{responseId}`

Each sheet contains:

| Range | Content |
|---|---|
| `A1:B10` | Project summary (ID, PROJECT, REVIEW DATE, EMAIL, RESOURCE COUNT, LEVEL 1–3 SCORES, FINAL SCORE, MATURITY) |
| `B11` | Hyperlink back to `DEMaturitySummary!A1` — yellow background, dark-red bold text |
| Starting at `B13` | **Level 1 Questions** table (`Level1Table{id}`) — columns: Question, Response |
| After Level 1 table | **Level 2 Questions** table (`Level2Table{id}`) — columns: Question, Response |
| After Level 2 table | **Level 3 Questions** table (`Level3Table{id}`) — columns: Question, Response |

**Formatting conventions on project sheets:**

* Section header rows (e.g., "Level 1 Questions") have a blue background (`#154CC5`) with white bold centred text.
* The Maturity cell (`B10`) has a yellow (`#FFFFE0`) background, dark-red (`#800000`) bold text, and double-line borders on all sides.
* Individual question rows with a non-perfect response (score < 10) are highlighted red (`#FF0000`).
* All three level tables have column filters pre-set to display only: `Frequently`, `Sometimes`, `Never`, `No`, `Don't Know`.

---

## Project Structure

```
DEMaturityCalculator/
├── assets/                        # Add-in icons and branding
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-80.png
│   ├── logo-filled.png
│   └── neudesilogo.png            # Neudesic logo shown in the task pane header
│
├── src/
│   ├── commands/
│   │   ├── commands.html          # Invisible HTML host for ribbon command functions
│   │   └── commands.js            # Office ribbon command handler (action function)
│   │
│   └── taskpane/
│       ├── taskpane.html          # Task pane UI (Welcome header + Calculate Maturity button)
│       ├── taskpane.js            # Core logic: reads survey data, calculates scores, writes output
│       ├── taskpane.css           # Office UI Fabric-based styles for the task pane
│       ├── dialogs.js             # officejs.dialogs library (Wait spinner, MessageBox, Alert)
│       └── dialogs.html           # HTML host for the officejs.dialogs overlay dialogs
│
├── manifest.xml                   # Office Add-in manifest (ID, host, ribbon config, URLs)
├── package.json                   # npm metadata, scripts, and dependency declarations
├── webpack.config.js              # Webpack 4 build configuration (entries, loaders, plugins)
├── tsconfig.json                  # TypeScript config (allowJs: true — used for type checking JS)
├── .eslintrc.json                 # ESLint rules (extends eslint-config-office-addins)
├── CONTRIBUTING.md                # Contribution guidelines
└── LICENSE                        # MIT License
```

### Key Source Files

| File | Purpose |
|---|---|
| `src/taskpane/taskpane.js` | Entry point for all business logic: reads `Form1/Table1`, calculates level scores with `CalculateLevelScore()`, creates `DEMaturitySummary` and per-project sheets, applies formatting, filters, and hyperlinks |
| `src/taskpane/taskpane.html` | Minimal task pane UI: Neudesic logo, welcome header, and the **Calculate Maturity** button |
| `src/taskpane/dialogs.js` | Bundled copy of the `officejs.dialogs` v1.0.9 library — provides `Wait`, `MessageBox`, `Alert`, `InputBox`, `Progress`, `Form` globals |
| `manifest.xml` | Declares the add-in to Office: host (`Workbook`), permissions (`ReadWriteDocument`), ribbon button placement (`Home → DE Group`), and dev-server URLs |
| `webpack.config.js` | Builds two entry points (`taskpane`, `commands`) and copies static assets; dev server uses `office-addin-dev-certs` for HTTPS on port 3000 |

---

## Development

### Available npm Scripts

| Script | Command | Description |
|---|---|---|
| `build` | `npm run build` | Production webpack build (minified, mode: production) |
| `build:dev` | `npm run build:dev` | Development webpack build (unminified, with source maps) |
| `dev-server` | `npm run dev-server` | Start webpack-dev-server (HTTPS, port 3000, hot reload) |
| `watch` | `npm run watch` | Webpack watch mode — rebuilds on file changes without starting the dev server |
| `start` | `npm start` | Start dev server **and** sideload the add-in (auto-detect platform) |
| `start:desktop` | `npm run start:desktop` | Start dev server **and** sideload into Excel Desktop |
| `start:web` | `npm run start:web` | Start dev server **and** sideload into Excel on the Web |
| `stop` | `npm stop` | Stop the dev server and remove the sideloaded add-in |
| `validate` | `npm run validate` | Validate `manifest.xml` against the Office Add-in schema |
| `lint` | `npm run lint` | Run ESLint checks (office-addin-lint) |
| `lint:fix` | `npm run lint:fix` | Auto-fix linting issues |
| `prettier` | `npm run prettier` | Format code with Prettier (office-addin-prettier-config) |

### Build Output

Webpack writes compiled artefacts to the `dist/` directory (cleaned on each build via `CleanWebpackPlugin`):

```
dist/
├── taskpane.html       # Task pane HTML (injected with hashed script tags)
├── taskpane.js         # Bundled and transpiled task pane code
├── taskpane.css        # Copied stylesheet
├── commands.html       # Commands HTML
├── commands.js         # Bundled commands code
├── dialogs.js          # Copied dialogs library
├── dialogs.html        # Copied dialogs host HTML
├── polyfill.js         # Babel polyfill bundle
└── assets/             # Copied icon and image assets
```

### Development Server

The webpack dev server runs at **`https://localhost:3000`** with HTTPS enabled via `office-addin-dev-certs`. The port is configurable via the `config.dev-server-port` field in `package.json`.

```bash
npm run dev-server
```

CORS is unrestricted (`Access-Control-Allow-Origin: *`) to allow Office clients on any domain to load resources from localhost.

### Linting & Formatting

The project extends `eslint-config-office-addins` for lint rules and uses `office-addin-prettier-config` for code style:

```bash
npm run lint          # Check for lint errors
npm run lint:fix      # Auto-fix lint errors
npm run prettier      # Format all files with Prettier
```

### Manifest Validation

Before deploying a new manifest version, validate it against the official Office Add-in schema:

```bash
npm run validate
```

---

## Debugging

The add-in supports several debugging approaches:

### Browser Developer Tools

When running in Excel on the Web or with `npm run start:web`, open the browser's DevTools (`F12`). The task pane renders in a web page context, so all standard browser debugging tools (Sources, Console, Network) are available.

### Desktop Debugger

On Excel Desktop (Windows), the task pane runs in an embedded browser. Attach a debugger using:

* **Visual Studio Code** — use the `.vscode/` configuration included in this repository.
* **F12 Developer Tools** — on Windows, press `F12` inside the task pane to open the Edge DevTools.

For more information, see the [Microsoft Office Add-in debugging documentation](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online).

### Extended Error Logging

The add-in enables `OfficeExtension.config.extendedErrorLogging = true` at runtime. Extended errors from Excel API calls are printed to the browser/DevTools console, including property paths and inner errors.

### Common Errors

| Error | Cause | Fix |
|---|---|---|
| `ItemNotFound` | `Form1` sheet or `Table1` table does not exist in the workbook | Ensure the workbook has a sheet named exactly `Form1` containing a table named `Table1` |
| Certificate warnings in browser | Dev cert not trusted by the OS | Run `npx office-addin-dev-certs install` and trust the certificate when prompted |
| Task pane shows "Please sideload your add-in" | Add-in not sideloaded, or not running in Excel | Ensure `npm start` completed successfully and the add-in was loaded via the ribbon |
| Blank or incorrect scores | Response values in the table don't match recognised strings | Check that responses are one of: `NA`, `Always`, `Yes`, `Frequently`, `Sometimes`, `Never`, `No`, `Don't Know` |

---

## Contributing

Contributions are welcome! Please read [`CONTRIBUTING.md`](CONTRIBUTING.md) for full guidelines. A summary:

1. **Fork** the repository and create a feature branch from `master`.
2. **Make your changes** — follow the existing code style (ESLint + Prettier).
3. **Test** your changes in both Excel Desktop and Excel on the Web where possible.
4. **Validate** the manifest if you modify `manifest.xml`:

   ```bash
   npm run validate
   ```

5. **Lint** your code before submitting:

   ```bash
   npm run lint
   ```

6. **Open a pull request** against the `master` branch with a clear description of the change.

Pull requests are typically reviewed within 10 business days. For new features, please open an issue first to discuss intent with the repository maintainers.

---

## License

This project is licensed under the **MIT License**. See the [`LICENSE`](LICENSE) file for full details.

```
MIT License — Copyright (c) Microsoft Corporation. All rights reserved.
```

---

## Additional Resources

* [Office Add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [Excel JavaScript API reference](https://docs.microsoft.com/javascript/api/excel)
* [Office UI Fabric Core](https://developer.microsoft.com/fabric)
* [officejs.dialogs library](http://theofficecontext.com)
* [Neudesic](https://www.neudesic.com)
* [Office Add-in sideloading guide](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
* [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
