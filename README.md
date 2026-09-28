# Accounts Receivable Aging Analysis

An interactive, browser-based accounts receivable dashboard created for **ACC 455 — Audit Data Analytics**. The site reproduces a three-stage Alteryx-style audit workflow, calculates invoice aging from an adjustable analysis date, and presents transaction detail, summaries, pivots, charts, and audit observations.

**Live site:** [aging-analysis.vercel.app](https://aging-analysis.vercel.app/)

## Overview

The application loads the included `AR-Data.xlsx` workbook automatically and also accepts user-supplied Excel workbooks. All workbook parsing and analysis happen locally in the browser. The application code does not send uploaded workbook data to a server.

The dashboard provides:

- One-, two-, and three-dimensional AR aging analyses
- An adjustable analysis date, defaulting to February 1, 2026
- Automatic Excel worksheet and column detection
- Transaction-level search, filtering, and sortable columns
- Aging KPIs, summary statistics, pivot tables, and interactive charts
- Data-quality reporting for valid and excluded records
- Reconciliation checks across every analytical output
- Interactive workflow diagrams that explain each transformation step
- Responsive desktop, tablet, and mobile layouts

## Analysis workflows

### 1. One-Dimensional Aging

Prepares the complete AR population and analyzes balances by aging category.

1. Reads the transaction and customer-type source data.
2. Joins customer type to the AR transactions.
3. Calculates invoice age in calendar days.
4. Assigns each transaction to an aging category.
5. Retains and formats the audit-ready output fields.
6. Produces transaction detail, an aging summary, a distribution chart, and an audit observation.

The output includes these KPIs:

- Total receivables
- Total invoices
- Average invoice balance
- Balance and percentage over 90 days old

### 2. Two-Dimensional Aging

Transforms the processed population into a **Customer × Aging Classification** matrix.

- Rows: Customer
- Columns: Aging classification
- Values: Outstanding balance
- Aggregation: Sum

The accompanying stacked chart compares each customer's receivable balance across the four aging categories.

### 3. Three-Dimensional Aging

Analyzes balances across both aging classification and customer type.

- Rows: Aging classification
- Columns: Customer type
- Values: Outstanding balance
- Aggregation: Sum

The accompanying grouped chart compares customer-type balances across all aging categories.

## Aging logic

Invoice age is calculated using UTC calendar dates:

```text
Aging in days = Analysis date - Invoice date
```

Transactions are classified as follows:

| Aging in days | Classification |
|---:|---|
| Less than 30 | Current |
| 30–59 | 30–59 days |
| 60–89 | 60–89 days |
| 90 or more | Over 90 days |

The default analysis date is `2026-02-01`, but it can be changed directly from the dashboard. Every calculation and chart refreshes when the date changes.

## Workbook requirements

The uploader accepts `.xlsx` and `.xls` files.

The transaction worksheet must contain these fields:

| Required field | Recognized header examples |
|---|---|
| Customer | `Customer`, `Customer Number`, `Customer ID`, `Cust`, `Cust No` |
| Sales Person | `Sales Person`, `Sales Representative`, `Sales Rep` |
| Balance | `Bal.`, `Balance`, `Account Balance`, `Amount`, `Outstanding Balance`, `AR Balance` |
| Invoice Date | `Inv. Date`, `Invoice Date`, `Date of Invoice` |
| Customer Type | `Cust. Type`, `Customer Type`, `Type`, `Customer Category` |

`Customer Type` may appear in the transaction worksheet. If it does not, the workbook must include a separate customer-type worksheet containing Customer and Customer Type.

The application prefers a transaction sheet whose name resembles `Breakdown of Sales`. Otherwise, it selects the first sheet containing Customer, Balance, and Invoice Date. Customer-type data is joined by matching record position when both sheets align; otherwise, it is joined by normalized customer key.

### Supported values

- Balances may be Excel numbers or formatted text such as `$1,234.56` and `(250.00)`.
- Invoice dates may be Excel date values, `YYYY-MM-DD`, or `MM/DD/YYYY`.
- Header matching ignores capitalization, spaces, and punctuation.

### Data-quality exclusions

Rows are excluded when they contain:

- A completely blank row
- A missing customer
- A missing sales person
- An invalid balance
- An invalid invoice date
- A missing customer type

Select **Data Quality** in the dashboard to review valid and excluded record counts and the reason for each exclusion.

## Reference workbook results

Using the included workbook and the default analysis date, the application reconciles to:

| Metric | Result |
|---|---:|
| Valid records | 107 |
| Total receivables | $10,622.10 |
| Average invoice balance | $99.27 |
| Over 90 days | $3,338.38 |

All aging-summary, customer-pivot, and customer-type-pivot totals are checked against total receivables. A difference greater than half a cent is reported in the browser console.

## Using the website

1. Open the [live dashboard](https://aging-analysis.vercel.app/).
2. Review the automatically loaded reference workbook or select **Upload New Excel File**.
3. Choose the analysis date.
4. Move between the three analysis tabs.
5. Select any workflow node to read its explanation.
6. Search, filter, or sort the transaction population in Analysis 1.
7. Review the summary tables, charts, data-quality results, and audit observations.
8. Select **Reset Analysis** to restore the reference workbook and default analysis date.

## Technology

- Semantic HTML5
- Responsive CSS3
- Vanilla JavaScript
- [SheetJS](https://sheetjs.com/) for Excel parsing
- [Chart.js](https://www.chartjs.org/) for visualizations
- Vercel for static hosting and continuous deployment

The project has no package manager, framework, server, database, or build step. SheetJS and Chart.js are loaded from a CDN, so an internet connection is required to load those libraries.

## Run locally

Clone the repository:

```bash
git clone https://github.com/haseebshaik00/AR-Aging-Analysis.git
cd AR-Aging-Analysis
```

Start a local static server:

```bash
python3 -m http.server 8000
```

Then open [http://localhost:8000](http://localhost:8000). A local server is recommended because browsers may block the reference workbook request when `index.html` is opened directly from the filesystem.

## Deploy on Vercel

Import the GitHub repository into Vercel with these settings:

| Setting | Value |
|---|---|
| Framework Preset | Other |
| Root Directory | `./` |
| Build Command | Leave blank |
| Output Directory | Leave blank/default |
| Install Command | Leave blank |
| Environment Variables | None |

Vercel serves the repository root as a static site. Every push to `main` automatically creates a new production deployment.

## Project structure

```text
AR-Aging-Analysis/
├── index.html                  # Page structure and accessible interface
├── styles.css                 # Visual system and responsive layouts
├── script.js                  # Workbook processing, calculations, and rendering
├── Reference/
│   ├── AR-Data.xlsx           # Default reference dataset
│   └── Alteryx-Workbook.pdf   # Original workflow reference
└── README.md                  # Project documentation
```

## Accessibility and responsive behavior

- Semantic tab and table markup
- Keyboard navigation between analysis tabs
- Visible focus states and a skip link
- Accessible labels for controls and charts
- Reduced-motion support
- A single-screen dashboard layout on supported desktop viewports
- Reflowed controls, tables, workflows, and cards for tablet and mobile screens

## Privacy

Uploaded workbooks are read with the browser's File API and processed in memory. The application does not include analytics, authentication, server-side storage, or an upload endpoint. Refreshing or closing the page clears the uploaded data from application memory.
