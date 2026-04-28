# DEMaturityCalculator

A **Microsoft Excel Office Add-in** built by [Neudesic](https://www.neudesic.com) that reads Data Engineering (DE) survey responses from an Excel workbook and automatically calculates DE maturity levels (M1, M2, M3) for each project. Results are written to a summary sheet and individual per-project detail sheets — all without leaving Excel.

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation & Setup](#installation--setup)
- [Running the Add-in](#running-the-add-in)
- [Input Data Format](#input-data-format)
- [Output Description](#output-description)
- [Maturity Level Scoring](#maturity-level-scoring)
- [npm Scripts](#npm-scripts)
- [Project Structure](#project-structure)
- [Debugging](#debugging)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

DEMaturityCalculator processes DE assessment survey data stored in an Excel table and produces:

- A **`DEMaturitySummary`** worksheet with a score and maturity rating (M1/M2/M3) for every project.
- **Per-project worksheets** with a score breakdown and filtered Q&A tables showing only questions that need improvement.
- Bidirectional **hyperlinks** between the summary sheet and each project sheet for quick navigation.

For a full technical walkthrough, see [`docs/architecture.md`](docs/architecture.md).

---

## Prerequisites

| Requirement | Version / Notes |
|---|---|
| **Node.js** | v14 or later (LTS recommended) |
| **npm** | v6 or later (bundled with Node.js) |
| **Microsoft Excel** | Microsoft 365 desktop (Windows or macOS) or Excel on the web |
| **Git** | For cloning the repository |

---

## Installation & Setup

1. **Clone the repository**

   ```bash
   git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
   cd DEMaturityCalculator
   ```

2. **Install dependencies**

   ```bash
   npm install
   ```

3. **Trust the development HTTPS certificate** (first run only)

   The local dev server runs on HTTPS. The `office-addin-dev-certs` package generates a self-signed cert automatically when you run `npm start`, but you may be prompted to trust it in your OS certificate store.

   ```bash
   npx office-addin-dev-certs install
   ```

4. **Validate the manifest** (optional sanity check)

   ```bash
   npm run validate
   ```

---

## Running the Add-in

### Desktop (Excel for Windows / macOS)

```bash
npm start
```

This command:
1. Starts the Webpack dev server on `https://localhost:3000`.
2. Sideloads `manifest.xml` into Excel and launches the application.

Excel will open automatically with the add-in sideloaded. A **"DE Group"** section appears on the **Home** ribbon tab. Click **Show DEMaturityCalculator** to open the task pane.

### Stopping the Add-in

```bash
npm stop
```

### Excel on the Web

```bash
npm run start:web
```

Follow the prompts to sideload the add-in in the browser-based Excel session.

### Using the Add-in

1. Open (or have open) the Excel workbook that contains your DE survey responses — the workbook **must** include a sheet named **`Form1`** with a table named **`Table1`**.
2. Open the DEMaturityCalculator task pane via the **Home → DE Group → Show DEMaturityCalculator** ribbon button.
3. Click **Calculate Maturity** in the task pane.
4. A progress spinner appears while the add-in processes all project rows.
5. When complete, the **`DEMaturitySummary`** sheet is activated automatically.

---

## Input Data Format

The add-in reads data from a specific location in the active workbook:

| Setting | Value |
|---|---|
| **Sheet name** | `Form1` |
| **Table name** | `Table1` |
| **Header row** | Row 1 — question text in each column |
| **Data rows** | One row per project survey submission |

### Required Column Positions (0-based index)

| Index | Field |
|---|---|
| 0 | Response / Submission ID |
| 2 | Review Date (Excel serial date number) |
| 3 | Respondent Email |
| 5 | Project Name |
| 6 | Resource Count |
| 7–91 | Survey question responses |

### Valid Response Values

| Response | Score |
|---|---|
| `NA` | 10 |
| `Always` | 10 |
| `Yes` | 10 |
| `Frequently` | 7 |
| `Sometimes` | 4 |
| `Never` | 0 |
| `No` | 0 |
| `Don't Know` | 0 |

Any response that is blank or does not match a known value is treated as **0**.

---

## Output Description

### `DEMaturitySummary` Sheet

A table named **`SummaryTable`** is created with one row per project:

| Column | Description |
|---|---|
| **ID** | Submission ID from the source table |
| **PROJECT** | Project name (hyperlink to project detail sheet) |
| **REVIEW DATE** | Formatted date (`MM-DD-YYYY`) |
| **EMAIL** | Respondent email |
| **RESOURCE COUNT** | Team size / resource count |
| **LEVEL 1 SCORE** | Weighted Level 1 score (max 70). Shown in **red** if below 70 |
| **LEVEL 2 SCORE** | Weighted Level 2 score (max 20) |
| **LEVEL 3 SCORE** | Weighted Level 3 score (max 10) |
| **FINAL SCORE** | Sum of all three weighted scores (max 100) |
| **MATURITY** | `M1`, `M2`, or `M3` — hyperlink to project detail sheet |

### Per-Project Detail Sheets

Each project gets its own worksheet named `{ProjectName}_{ID}` (alphanumeric only, max 25 chars from project name).

Each sheet contains:

- **Rows A1:B10** — Summary card (ID, project name, review date, email, resource count, level scores, final score, maturity).
- **Cell B11** — Hyperlink back to `DEMaturitySummary`.
- **Level 1 Questions table** — All Level 1 Q&A rows; filtered to show only non-perfect responses (`Frequently`, `Sometimes`, `Never`, `No`, `Don't Know`). Failing rows are highlighted in **red**.
- **Level 2 Questions table** — Same structure as Level 1 but for Level 2 questions.
- **Level 3 Questions table** — Same structure for Level 3 questions.

---

## Maturity Level Scoring

### Question Levels

| Level | Column Indexes | Weight | Max Contribution |
|---|---|---|---|
| **Level 1** | 58 questions (indexes 7–86) | 70% | 70 points |
| **Level 2** | 15 questions (indexes 15–88) | 20% | 20 points |
| **Level 3** | 12 questions (indexes 26–91) | 10% | 10 points |

### Scoring Formula

For each level:

```
Raw Score         = sum of all individual question scores in the level
Max Possible      = number_of_questions × 10
Level Percentage  = (Raw Score / Max Possible) × 100
Weighted Score    = (Level Percentage × Weight) / 100
```

Final Score = Level 1 Weighted Score + Level 2 Weighted Score + Level 3 Weighted Score

### Maturity Thresholds

| Final Score | Maturity Level | Meaning |
|---|---|---|
| ≤ 70 | **M1** | Initial / foundational practices |
| > 70 and ≤ 90 | **M2** | Developing / intermediate practices |
| > 90 | **M3** | Advanced / mature practices |

> **Note:** The Level 1 weighted score also carries a standalone threshold of 70. If it is below 70, it is highlighted in red in both the summary and project sheets to flag a critical gap, regardless of the overall maturity level.

---

## npm Scripts

| Script | Command | Description |
|---|---|---|
| `start` | `office-addin-debugging start manifest.xml` | Start dev server + sideload into Excel desktop |
| `start:desktop` | `office-addin-debugging start manifest.xml desktop` | Explicitly target Excel desktop |
| `start:web` | `office-addin-debugging start manifest.xml web` | Sideload into Excel on the web |
| `stop` | `office-addin-debugging stop manifest.xml` | Stop dev server and remove sideloaded add-in |
| `build` | `webpack -p --mode production` | Production build (minified) |
| `build:dev` | `webpack --mode development` | Development build with source maps |
| `watch` | `webpack --mode development --watch` | Continuous rebuild on file changes |
| `dev-server` | `webpack-dev-server --mode development` | Dev server only (no auto-sideload) |
| `validate` | `office-addin-manifest validate manifest.xml` | Validate `manifest.xml` against the Office schema |
| `lint` | `office-addin-lint check` | Run ESLint checks |
| `lint:fix` | `office-addin-lint fix` | Auto-fix ESLint issues |
| `prettier` | `office-addin-lint prettier` | Format code with Prettier |

---

## Project Structure

```
DEMaturityCalculator/
├── assets/                     # Add-in icons (16×16, 32×32, 80×80) and Neudesic logo
├── src/
│   ├── taskpane/
│   │   ├── taskpane.js         # Core add-in logic: reads survey data, calculates maturity, writes output
│   │   ├── taskpane.html       # Task pane UI (Welcome header + "Calculate Maturity" button)
│   │   ├── taskpane.css        # Task pane styles (Office UI Fabric conventions)
│   │   ├── dialogs.js          # officejs.dialogs library — Wait spinner, MessageBox, Alert, etc.
│   │   └── dialogs.html        # Companion HTML page for the dialogs iframe
│   └── commands/
│       ├── commands.js         # Add-in command handler (ribbon button action)
│       └── commands.html       # Commands function file host page
├── docs/                       # Project documentation
│   ├── getting-started.md
│   ├── architecture.md
│   └── api-reference.md
├── manifest.xml                # Office Add-in manifest (ID, permissions, ribbon extension points)
├── package.json                # npm metadata, scripts, and dependencies
├── webpack.config.js           # Webpack bundler configuration
├── tsconfig.json               # TypeScript/Babel compiler options
├── .eslintrc.json              # ESLint configuration
├── CONTRIBUTING.md             # Contribution guidelines
└── LICENSE                     # MIT License
```

---

## Debugging

| Technique | When to use |
|---|---|
| **Browser DevTools** (`F12`) | Excel on the web or when the task pane opens in a browser WebView |
| **Attach debugger from task pane** | Excel desktop — right-click the task pane → "Inspect" (Windows) |
| **F12 Developer Tools** | Excel on Windows 10/11 with the legacy IE WebView engine |
| **VS Code** | Use the `.vscode/` launch configurations included in the repo |

Extended error logging is enabled in `taskpane.js` to capture full Office.js error details:

```js
OfficeExtension.config.extendedErrorLogging = true;
```

---

## Contributing

We welcome contributions! Please read [`CONTRIBUTING.md`](CONTRIBUTING.md) for:

- Code comment standards
- How to fix open issues
- How to propose new features
- Pull request guidelines and the CLA process

---

## License

MIT — see [`LICENSE`](LICENSE) for full terms.

---

*Provider: [Neudesic](https://www.neudesic.com) · Add-in ID: `1c9376c1-ee2d-4875-82ee-f1bf4eed4374`*
