# Inventory Optimization Dashboard

Frontend-only React prototype for inventory optimization, aging analysis, carrying-cost visibility, danger-zone screening, and old-stock recovery planning.

The repository uses a synthetic 200-SKU demo dataset. It does not contain production company data, credentials, a backend, an API service, or a database.

## Requirements

- Node.js 20 or later
- pnpm 9 or later

## Install and run

```bash
pnpm install
pnpm dev
```

Open [http://127.0.0.1:4173](http://127.0.0.1:4173).

## Production build

```bash
pnpm build
pnpm preview
```

The production output is written to `dist/` and is intentionally excluded from Git.

## Prototype scope

- React, TypeScript, Vite, Recharts, and Lucide icons
- Responsive dashboard and navigation
- Inventory register, ABC/FSN/XYZ classification, aging and risk views
- Carrying-cost analysis and recovery-action planning
- CSV export and print-to-PDF support
- Synthetic demo data loaded from `public/demo-data.json`

Some screens describe future API or forecasting workflows. Those controls remain non-production placeholders in this frontend-only prototype; no backend or database is included.

## Data and security

- `node_modules/`, `dist/`, local environment files, and hosting metadata are ignored.
- Never commit real credentials or confidential inventory exports.
- Replace the synthetic dataset only with data approved for the target repository's visibility.
