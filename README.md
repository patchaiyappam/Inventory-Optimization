# Inventory Optimization Dashboard

A frontend-only React demonstration for inventory analysis, old-stock review, recovery planning, and basic demand forecasting. All imports, calculations, filters, charts, and exports run locally in the browser.

The included 200-SKU dataset is synthetic. The application does not include a backend, database, authentication, or remote data endpoint.

## Requirements

- Node.js 22.6 or later (required for the built-in TypeScript test runner)
- pnpm 9 or later

## Install and run

```bash
pnpm install
pnpm dev
```

Open [http://127.0.0.1:4173](http://127.0.0.1:4173).

## Checks and production build

```bash
pnpm typecheck
pnpm test
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
- Source-currency display with optional, explicitly rated illustrative conversion
- Synthetic demo data loaded from `public/demo-data.json`

## Import columns

Column names are matched case-insensitively and tolerate spaces and punctuation. Common aliases such as `Material Code`, `Qty`, `Standard Cost`, `Vendor`, and `Storage Location` are accepted.

Required columns:

- SKU or material/item code
- Quantity or on-hand quantity
- Unit cost, price, or standard cost

Optional columns include name, category, supplier, warehouse, purchase/receipt/issue dates, average monthly consumption, annual consumption value, lead time, MOQ, requirement date, ABC/FSN/XYZ classes, inventory age, and coefficient of variation. Missing optional inputs remain explicitly unavailable; they are not presented as observed data.

An explicitly supplied valid inventory age remains usable when a purchase or receipt date is absent. The missing source date is still reported as a limitation because the supplied age cannot be independently recalculated from the import.

## Method limits

- Risk flags are review triggers. Return, exchange, and transfer must be investigated before any separately authorized scrap decision.
- Risk is `Insufficient data`, never `Low`, when material age, movement history, or FSN classification is unavailable.
- Safety time is 50% of supplier lead time. Stock coverage uses lead time plus that safety time.
- Potentially trapped capital requires a confirmed requirement date, lead time, and consumption. Missing inputs produce `Unavailable`, not zero.
- Projected carrying cost requires a confirmed requirement date and lead time, and subtracts lead time plus its 50% safety time from the holding period.
- Forecasting requires at least six observed monthly issues. Calendar gaps remain missing unless the source period is explicitly confirmed complete; only then may gaps become zero-issue months.
- MAE and RMSE use the final chronologically held-out observations. The demonstration makes no production accuracy claim.
- Numeric demand regression is separate from slow-moving, non-moving, and excess probabilities, which remain unavailable without labeled outcomes.
- Category time-series analysis is unavailable for the included uncategorized dataset.
- Correlations use paired available observations only and report an unavailable state for insufficient data or zero variance.
- Values remain in the identified source currency unless a positive, visible illustrative exchange-rate assumption is supplied.
- Imported data exists only in the current browser session and is not uploaded or persisted.

## Data and security

- `node_modules/`, `dist/`, local environment files, and hosting metadata are ignored.
- Never commit real credentials or confidential inventory exports.
- Replace the synthetic dataset only with data approved for the target repository's visibility.
