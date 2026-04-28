# Getting Started — DEMaturityCalculator Excel Add-in

> **Provider:** Neudesic  
> **Add-in ID:** `1c9376c1-ee2d-4875-82ee-f1bf4eed4374`  
> **Version:** 1.0.0.0  
> **Host:** Microsoft Excel (Desktop & Online)  
> **License:** MIT

---

## Table of Contents

1. [What Is This Add-in?](#1-what-is-this-add-in)
2. [Prerequisites](#2-prerequisites)
3. [Clone & Install Dependencies](#3-clone--install-dependencies)
4. [Build the Add-in](#4-build-the-add-in)
5. [Run Locally (Development)](#5-run-locally-development)
6. [Sideload the Manifest into Excel](#6-sideload-the-manifest-into-excel)
7. [Preparing the Input Workbook](#7-preparing-the-input-workbook)
8. [Running a Maturity Calculation](#8-running-a-maturity-calculation)
9. [Understanding the Output](#9-understanding-the-output)
10. [Deploying to Production](#10-deploying-to-production)
11. [Debugging](#11-debugging)
12. [Common Errors](#12-common-errors)

---

## 1. What Is This Add-in?

The **DEMaturityCalculator** is a Microsoft Excel task-pane add-in built by Neudesic that automates the scoring of Digital Engineering (DE) maturity surveys. Given an Excel workbook that contains raw survey responses (exported from a Microsoft Forms questionnaire), clicking **Calculate Maturity** will:

- Parse every project's responses across **85 questions** grouped into four maturity levels.
- Calculate a weighted score for each level and a final composite score.
- Assign one of four maturity labels — **M1, M2, M3, M4** — to each project.
- Create a **DEMaturitySummary** sheet with aggregated results and hyperlinks.
- Create one **per-project sheet** that shows all question/response pairs, highlighting any failing answers in red.

---

## 2. Prerequisites

| Requirement | Version / Notes |
|---|---|
| **Node.js** | 12.x or later (LTS recommended) |
| **npm** | 6.x or later (bundled with Node.js) |
| **Microsoft Excel** | Desktop (Windows/Mac) **or** Excel on the Web (Office Online) |
| **Microsoft 365** | A valid M365 subscription is required to load add-ins |
| **Git** | To clone the repository |

---

## 3. Clone & Install Dependencies

```bash
git clone <your-repo-url>
cd DEMaturityCalculator
npm install
```

`npm install` downloads all dependencies declared in `package.json`, including the Webpack build chain and the `officejs.dialogs` runtime library.

---

## 4. Build the Add-in

### Production build

```bash
npm run build
```

Outputs optimised, minified bundles into the `dist/` directory.

### Development build (unminified, with source maps)

```bash
npm run build:dev
```

### Watch mode (rebuild on every file save)

```bash
npm run watch
```

---

## 5. Run Locally (Development)

The development server serves the add-in over **HTTPS on port 3000**, which is required by the Office platform.

```bash
npm start
```

This command:
1. Generates a self-signed TLS certificate via `office-addin-dev-certs` (you may be prompted to trust it).
2. Starts `webpack-dev-server` on `https://localhost:3000`.
3. Automatically sideloads the manifest into your local Excel installation (Windows/Mac desktop).

To start in web-only mode:

```bash
npm run start:web
```

To stop the development server and un-sideload the manifest:

```bash
npm stop
```

---

## 6. Sideload the Manifest into Excel

### Desktop Excel (Windows)

1. Run `npm start` — the manifest is sideloaded automatically.
2. Alternatively, follow [Microsoft's manual sideloading guide](https://docs.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

### Excel on the Web (Office Online)

1. Run `npm run start:web`.
2. Navigate to [office.com](https://office.com) → **Excel** → open a workbook.
3. **Insert** → **Add-ins** → **Upload My Add-in** → browse to `manifest.xml`.

### Validate the manifest

```bash
npm run validate
```

---

## 7. Preparing the Input Workbook

The add-in expects a specific workbook structure. It will fail gracefully with an `ItemNotFound` error if these requirements are not met.

### Required sheet: `Form1`

| Requirement | Detail |
|---|---|
| Sheet name | Exactly `Form1` (case-sensitive) |
| Table name | Exactly `Table1` (case-sensitive) |
| Header row | Column headers must be the survey question text |
| Data rows | One row per project submission |

### Required column positions (0-based index)

| Index | Field |
|---|---|
| 0 | Response ID |
| 2 | Review Date *(Excel serial date number)* |
| 3 | Respondent Email |
| 5 | Project Name |
| 6 | Resource Count |
| 7–91 | Survey question responses |

> **Note:** Columns 1, 4, and intermediate columns are reserved by the survey form structure. The add-in reads columns by their absolute index, so column order must not be changed.

### Valid response values

Every answer cell must contain **exactly** one of the following strings (case-sensitive):

| Response | Score |
|---|---|
| `Always` | 10 |
| `NA` | 10 |
| `Yes` | 10 |
| `Frequently` | 7 |
| `Sometimes` | 4 |
| `Never` | 0 |
| `No` | 0 |
| `Don't Know` | 0 |

Empty or unrecognised values are treated as **0**.

---

## 8. Running a Maturity Calculation

1. Open the prepared survey-response workbook in Excel.
2. Click the **Show DEMaturityCalculator** button in the **Home** tab ribbon (under the **DE Group** section).
3. The task pane opens on the right side of Excel, showing the Neudesic logo and the **Calculate Maturity** button.
4. Click **Calculate Maturity**.
5. A progress spinner ("Processing Project Responses") appears while calculation runs.
6. When processing completes, the **DEMaturitySummary** sheet is activated automatically.

> ⚠️ **All sheets except `Form1` and any sheet whose name begins with `_` are deleted** before new sheets are created. Save any manual notes elsewhere first.

---

## 9. Understanding the Output

### DEMaturitySummary sheet

A table named `SummaryTable` is created with the following 11 columns:

| Column | Header | Description |
|---|---|---|
| A | ID | Response ID from the original form |
| B | PROJECT | Project name (hyperlink to the project detail sheet) |
| C | REVIEW DATE | Review date formatted as `MM-DD-YYYY` |
| D | EMAIL | Respondent's email address |
| E | RESOURCE COUNT | Number of team members |
| F | LEVEL 1 SCORE | Weighted score for Level 1 (max 60) |
| G | LEVEL 2 SCORE | Weighted score for Level 2 (max 20) |
| H | LEVEL 3 SCORE | Weighted score for Level 3 (max 10) |
| I | LEVEL 4 SCORE | Weighted score for Level 4 (max 10) |
| J | FINAL SCORE | Sum of all four level scores (max 100) |
| K | MATURITY | M1 / M2 / M3 / M4 (hyperlink to the project sheet) |

**Colour coding on the summary sheet:**
- Level 1 Score cell is highlighted in **red** if the score is below 60.
- Maturity cell has a **light-yellow background** with **dark-red bold text**.

### Per-project sheets

A sheet is created for every project row, named `<ProjectName>_<ResponseID>` (project name is truncated to 25 characters and non-alphanumeric characters are removed).

Each project sheet contains:

**Summary block (A1:B11)**

| Row | Label | Value |
|---|---|---|
| 1 | ID | Response ID |
| 2 | PROJECT | Project name |
| 3 | REVIEW DATE | Formatted date |
| 4 | EMAIL | Email address |
| 5 | RESOURCE COUNT | Team size |
| 6 | LEVEL 1 SCORE | Weighted score |
| 7 | LEVEL 2 SCORE | Weighted score |
| 8 | LEVEL 3 SCORE | Weighted score |
| 9 | LEVEL 4 SCORE | Weighted score |
| 10 | FINAL SCORE | Composite score |
| 11 | MATURITY | M1 / M2 / M3 / M4 |

**Navigation link (B12)**
A yellow hyperlink labelled "Click here to go to DEMaturitySummary sheet" appears below the summary block.

**Level question tables (starting around row 14)**
Four tables are placed below the summary block in order: Level 1, Level 2, Level 3, Level 4. Each table has two columns:

| Column | Content |
|---|---|
| Question | Full question text from the form header |
| Response | The project's answer |

Rows where the response is not at full score (i.e., not `Always`, `NA`, or `Yes`) are highlighted in **red**. Each table also has an auto-filter applied to show only non-perfect responses by default (`Frequently`, `Sometimes`, `Never`, `No`, `Don't Know`).

---

## 10. Deploying to Production

1. Update `manifest.xml` — replace every `https://localhost:3000` URL with your production hosting URL (e.g., `https://myaddin.example.com`).
2. Run `npm run build` to produce the production bundle in `dist/`.
3. Deploy the contents of `dist/` to your web server / Azure Static Web App / SharePoint App Catalog.
4. Distribute `manifest.xml` to users via:
   - **Microsoft 365 Admin Center** (centralised deployment) — recommended for organisations.
   - **SharePoint App Catalog**.
   - Manual sideloading (development/testing only).

---

## 11. Debugging

| Technique | When to use |
|---|---|
| **Browser DevTools** (F12 / Cmd+Opt+I in Excel Online) | Inspect console errors and network requests in Office on the web |
| **Attach debugger from task pane** | Right-click inside the task pane → "Inspect Element" (where available) |
| **F12 Developer Tools on Windows 10** | Attach to the embedded browser process in desktop Excel |
| `OfficeExtension.config.extendedErrorLogging = true` | Already enabled in `taskpane.js`; provides detailed Office API error messages in the console |

---

## 12. Common Errors

| Error message | Cause | Fix |
|---|---|---|
| `Please run add-in on DE Survey responses` | `Form1` sheet or `Table1` table not found | Ensure the workbook has a sheet called `Form1` containing a table called `Table1` |
| `There is an error in processing project responses. Please try again or later` | Any other runtime exception | Open browser DevTools console for the full stack trace |
| Task pane shows "Please sideload your add-in" | Add-in not sideloaded | Follow [Section 6](#6-sideload-the-manifest-into-excel) |
| Certificate trust error in browser | Self-signed dev cert not trusted | Run `npx office-addin-dev-certs install` and trust the certificate |
