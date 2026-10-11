# Inventory Optimization Dashboard

A frontend-only React demonstration for inventory analysis, old-stock review, recovery planning, and basic demand forecasting. All imports, calculations, filters, charts, and exports run locally in the browser.

The included 200-SKU dataset is synthetic. The application does not include a backend, database, authentication, or remote data endpoint.

## Requirements

- Node.js 20 or later
- pnpm 9 or later

## Install and run

```bash
pnpm install
pnpm dev
```

Open [http://127.0.0.1:4173](http://127.0.0.1:4173).

## Type checking and production build

```bash
pnpm typecheck
pnpm build
```

## Production build

```bash
pnpm build
pnpm preview
```

The production output is written to `dist/` and is intentionally excluded from Git.

## Included features

- Local CSV/XLSX import from the first worksheet, limited to 15 MB and 10,000 rows
- Required-field validation for SKU, quantity, and unit cost; duplicate-SKU and optional-field reporting
- Search plus category, supplier, warehouse, ABC, FSN, XYZ, aging, risk, stock-value, and carrying-cost filters
- A single filtered record set shared by KPIs, tables, charts, and CSV exports
- Inventory aging, risk-review, carrying-cost, classification, and correlation views
- Recovery planning that includes every filtered flagged SKU
- Separate at-risk inventory value and estimated potentially trapped capital
- Verified financial outcomes only when actual proceeds, actual costs, and evidence are recorded
- Local monthly issue aggregation and a simple linear time-trend demonstration with time-ordered validation
- Configurable display currency and a Methodology & Assumptions page
- Synthetic demo data loaded from `public/demo-data.json`

## Import columns

Column names are matched case-insensitively and tolerate spaces and punctuation. Common aliases such as `Material Code`, `Qty`, `Standard Cost`, `Vendor`, and `Storage Location` are accepted.

Required columns:

- SKU or material/item code
- Quantity or on-hand quantity
- Unit cost, price, or standard cost

Optional columns include name, category, supplier, warehouse, purchase/receipt/issue dates, average monthly consumption, annual consumption value, lead time, MOQ, requirement date, ABC/FSN/XYZ classes, inventory age, and coefficient of variation. Missing optional inputs remain explicitly unavailable; they are not presented as observed data.

## Method limits

- Risk flags are review triggers, not disposal or scrap instructions.
- Potentially trapped capital is an estimate based on stock above a simple consumption-and-lead-time target and the configured recovery assumption.
- Projected carrying cost is shown only when both a confirmed requirement date and lead time are available.
- Forecasting requires at least six dated monthly issue observations for a SKU. It uses a simple linear trend and a final-period holdout; it is illustrative, not a production forecast.
- Correlations use paired available observations only and report an unavailable state for insufficient data or zero variance.
- Changing the currency changes the display label and formatting; it does not perform foreign-exchange conversion.
- Imported data exists only in the current browser session and is not uploaded or persisted.

## Data and security

- `node_modules/`, `dist/`, local environment files, and hosting metadata are ignored.
- Never commit real credentials or confidential inventory exports.
- Replace the synthetic dataset only with data approved for the target repository's visibility.
