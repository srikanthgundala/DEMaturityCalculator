# Getting Started

This guide walks you through setting up and running the **DEMaturityCalculator** Excel Add-in for the first time.

---

## Prerequisites

Before you begin, make sure you have the following installed:

| Tool | Minimum Version | Download |
|---|---|---|
| Node.js | 12.x | https://nodejs.org |
| npm | 6.x (ships with Node.js) | — |
| Microsoft Excel | Excel 2016 or Microsoft 365 | https://www.microsoft.com/en-us/microsoft-365 |

> **Tip:** Run `node -v` and `npm -v` to verify your installed versions.

---

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/srikanthgundala/DEMaturityCalculator.git
cd DEMaturityCalculator
```

### 2. Install Dependencies

```bash
npm install
```

This installs all runtime and development dependencies defined in `package.json`.

### 3. Trust the Development Certificate

The dev server uses HTTPS on `localhost:3000`. You need to trust the self-signed certificate once:

```bash
npx office-addin-dev-certs install
```

---

## Running the Add-in (Development)

```bash
npm start
```

This command:
1. Compiles the source files using webpack.
2. Starts a local HTTPS server at `https://localhost:3000`.
3. Automatically sideloads the add-in manifest into Excel (desktop).

Once Excel opens, look for the **DE Group** on the **Home** tab ribbon and click **Show DEMaturityCalculator**.

### Running in Excel Online

```bash
npm run start:web
```

Follow the on-screen instructions to sideload the manifest manually in Excel Online.

---

## Preparing the Excel Workbook

The add-in expects the following workbook structure:

| Sheet Name | Table Name | Description |
|---|---|---|
| `Form1` | `Table1` | Contains the DE survey responses exported from Microsoft Forms |

The header row of `Table1` must contain the survey question text. The data rows must contain one project response per row with the following columns at fixed positions:

| Column Index | Field |
|---|---|
| 0 | Response ID |
| 2 | Review Date (Excel date serial number) |
| 3 | Respondent Email |
| 5 | Project Name |
| 6 | Resource Count |
| 7–91 | Survey question responses |

---

## Building for Production

```bash
npm run build
```

The production bundle is output to the `dist/` folder. Host these static files on your organization's web server and update the URLs in `manifest.xml` accordingly.

---

## Validating the Manifest

```bash
npm run validate
```

This runs `office-addin-manifest validate` against `manifest.xml` and reports any schema errors.

---

## Stopping the Debug Session

```bash
npm run stop
```

---

## Next Steps

- [Architecture Overview](architecture.md) — understand how the add-in is structured.
- [API Reference](api-reference.md) — detailed description of the core functions.
