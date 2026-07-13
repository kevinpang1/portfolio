# Stock Holdings Percent Site

This Sites project serves the latest combined stock holdings percent dashboard.
The root path redirects to the dashboard HTML, and archived percent-dashboard
versions are available from the dashboard's version selector.

## Prerequisites

- Node.js `>=22.13.0`

## Quick Start

```bash
node preview-site.mjs
node build-site.mjs
```

No package install is required.

## Included Files

- `public/stock_holdings_percent_dashboard.html`: latest dashboard
- `public/stock_dashboard_versions/`: archived percent dashboard pages
- `worker/index.js`: Cloudflare Worker entry point
- `build-site.mjs`: creates the Sites deployment artifact under `dist/`

Raw portfolio CSV exports are intentionally not copied into `public/`.
