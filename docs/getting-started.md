# Getting Started — DEMaturityCalculator

This guide walks you through every step needed to install, configure, and run the DEMaturityCalculator Excel add-in for the first time.

---

## Table of Contents

1. [System Requirements](#1-system-requirements)
2. [Clone and Install](#2-clone-and-install)
3. [Trust the Dev Certificate](#3-trust-the-dev-certificate)
4. [Prepare Your Survey Workbook](#4-prepare-your-survey-workbook)
5. [Start the Add-in](#5-start-the-add-in)
6. [Calculate Maturity](#6-calculate-maturity)
7. [Reading the Results](#7-reading-the-results)
8. [Stopping the Add-in](#8-stopping-the-add-in)
9. [Building for Production](#9-building-for-production)
10. [Troubleshooting](#10-troubleshooting)

---

## 1. System Requirements

Before you begin, make sure the following tools are installed:

| Tool | Minimum Version | Download |
|---|---|---|
| **Node.js** | 14.x LTS or later | https://nodejs.org |
| **npm** | 6.x or later (ships with Node.js) | — |
| **Git** | Any recent version | https://git-scm.com |
| **Microsoft Excel** | Microsoft 365 (desktop or web) | — |

> **macOS users:** Excel for Microsoft 365 (subscription) is required. Perpetual-license versions of Excel (e.g., Excel 2019 without a Microsoft 365 subscription) may have limited support for task-pane add-ins and ribbon customisation via `VersionOverrides`.

---

## 2. Clone and Install

Open a terminal (PowerShell, bash, or zsh) and run:

```bash
# Clone the repository
git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
cd DEMaturityCalculator

# Install all npm dependencies
npm install
```

The install step downloads Webpack, Babel, `office-addin-debugging`, `office-addin-dev-certs`, and all other build tooling declared in `package.json`.

---

## 3. Trust the Dev Certificate

The local development server must run on HTTPS (required by the Office Add-in security model). `office-addin-dev-certs` creates a self-signed certificate the first time you start the dev server, but you must explicitly trust it.

### Windows

```powershell
npx office-addin-dev-certs install
```

You will see a Windows security prompt asking you to trust the certificate. Click **Yes**.

### macOS

```bash
npx office-addin-dev-certs install
```

You may be prompted for your macOS administrator password to add the certificate to your Keychain.

### Verify the certificate (optional)

```bash
npx office-addin-dev-certs verify
```

A successful response looks like:

```
HTTPS certificate and key exist.
Certificate:  /Users/<you>/.office-addin-dev-certs/localhost.crt
Private key:  /Users/<you>/.office-addin-dev-certs/localhost.key
```

> **Note:** You only need to do this step once per machine. The certificate is stored in your user profile and reused on subsequent starts.

---

## 4. Prepare Your Survey Workbook

The add-in reads DE survey responses from a specific location in the active Excel workbook. Your workbook **must** meet the following requirements before you run the add-in:

### Required structure

| Requirement | Detail |
|---|---|
| **Sheet name** | Exactly `Form1` (case-sensitive) |
| **Table name** | Exactly `Table1` (case-sensitive) |
| **Row 1** | Header row — question text in each column |
| **Rows 2+** | One data row per project survey submission |

### Required columns (0-based index)

| Index | Field | Notes |
|---|---|---|
| 0 | Submission ID | Unique identifier per response |
| 2 | Review Date | Must be an **Excel serial date** (numeric), not a text string |
| 3 | Email | Respondent's email address |
| 5 | Project Name | Used as the label for the per-project sheet |
| 6 | Resource Count | Team size |
| 7–91 | Survey responses | See valid values below |

### Valid response values

The cells in question columns must contain one of the following exact strings (case-sensitive):

```
NA            → 10 points
Always        → 10 points
Yes           → 10 points
Frequently    →  7 points
Sometimes     →  4 points
Never         →  0 points
No            →  0 points
Don't Know    →  0 points
```

Blank cells or unrecognized values are treated as **0 points**.

---

## 5. Start the Add-in

### Option A — Excel Desktop (recommended)

```bash
npm start
```

What happens:
1. Webpack dev server starts at `https://localhost:3000`.
2. The `office-addin-debugging` tool sideloads `manifest.xml` into your local Excel installation.
3. Excel opens (or an existing instance is used).
4. A yellow "Get started" banner appears in Excel confirming the add-in loaded.

### Option B — Excel on the Web

```bash
npm run start:web
```

The CLI will give you instructions for manually sideloading the manifest in the browser-based Excel.

### Confirm the add-in is loaded

In Excel, go to the **Home** tab. You should see a **"DE Group"** section in the ribbon with a **"Show DEMaturityCalculator"** button.

If the ribbon group is not visible, click **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs** and ensure `https://localhost:3000` is trusted, then restart Excel.

---

## 6. Calculate Maturity

1. **Open your survey workbook** — the file that contains the `Form1` sheet and `Table1`.
2. Click **Home → DE Group → Show DEMaturityCalculator** to open the task pane on the right side of the Excel window.
3. You will see the Neudesic logo, a "Welcome" header, and a blue **"Calculate Maturity"** button.
4. Click **Calculate Maturity**.
5. A "Processing Project Responses" wait spinner appears. Processing time depends on the number of project rows.
6. On completion, you are automatically switched to the **`DEMaturitySummary`** sheet.

> **If an error dialog appears:** The most common cause is that the active workbook does not contain a sheet named `Form1` or a table named `Table1`. Ensure you have the correct workbook active before clicking the button.

---

## 7. Reading the Results

### DEMaturitySummary sheet

This sheet is created (or recreated) every time you click **Calculate Maturity**. It contains a table (`SummaryTable`) with one row per project:

| Column | What it means |
|---|---|
| ID | Submission ID |
| PROJECT | Project name — click to jump to that project's detail sheet |
| REVIEW DATE | Formatted as `MM-DD-YYYY` |
| EMAIL | Respondent email |
| RESOURCE COUNT | Team headcount |
| LEVEL 1 SCORE | Weighted score, 0–70. **Red** = below 70 (critical gap) |
| LEVEL 2 SCORE | Weighted score, 0–20 |
| LEVEL 3 SCORE | Weighted score, 0–10 |
| FINAL SCORE | Total, 0–100 |
| MATURITY | M1 / M2 / M3 — click to jump to project sheet |

### Per-project detail sheets

Each project gets a sheet named `{ProjectName}_{ID}`. The sheet contains:

- **A1:B10** — Summary card with all scores and the maturity level.
- **B11** — "Click here to go to DEMaturitySummary sheet" hyperlink.
- **Level 1 / 2 / 3 Question tables** — Two-column tables (Question, Response) filtered to show only imperfect answers (`Frequently`, `Sometimes`, `Never`, `No`, `Don't Know`). Rows with imperfect answers are also coloured **red**.

### Maturity interpretation

| Rating | Final Score | What it means |
|---|---|---|
| **M1** | ≤ 70 | Foundational — significant DE practice gaps |
| **M2** | 71–90 | Developing — good foundation, room to improve |
| **M3** | > 90 | Advanced — mature, well-established DE practices |

---

## 8. Stopping the Add-in

To shut down the dev server and remove the sideloaded add-in from Excel:

```bash
npm stop
```

---

## 9. Building for Production

For a deployment build (minified, no source maps):

```bash
npm run build
```

Output is placed in the `dist/` folder. Host those static files on any HTTPS server and update the URLs in `manifest.xml` to point to your server before distributing the manifest to end users.

For a development build with source maps:

```bash
npm run build:dev
```

---

## 10. Troubleshooting

### "Please run add-in on DE Survey responses"

The workbook does not contain a sheet called `Form1` or a table called `Table1`. Open the correct survey workbook before clicking **Calculate Maturity**.

### Certificate errors / "Your connection is not private"

Re-run:

```bash
npx office-addin-dev-certs install
```

On Windows, also check that the certificate appears in **Trusted Root Certification Authorities** in the Windows Certificate Manager (`certmgr.msc`).

### Ribbon group not visible

1. Restart Excel completely (close all windows, reopen).
2. If still missing, check **File → Options → Add-ins → Manage COM Add-ins** for any disabled entry.
3. Run `npm stop` then `npm start` to force a fresh sideload.

### Port 3000 already in use

Another process is occupying port 3000. Either stop that process or change the port in `package.json`:

```json
"config": {
  "dev-server-port": 3001
}
```

Also update `manifest.xml` to replace all `localhost:3000` references with `localhost:3001`.

### Linting errors on save

Run `npm run lint:fix` to auto-fix ESLint issues, or `npm run prettier` to reformat code.

---

*Next: [Architecture Overview →](architecture.md)*
