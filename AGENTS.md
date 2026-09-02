# Repository Guidelines

## Project Structure & Module Organization

This static, Spanish-language browser application consults hospital leave records and produces administrative reports.

- `index.html` defines the page shell and loads browser libraries from CDNs.
- `styles.css` contains all layout, component, and print styling.
- `cs.js` handles data loading, search, rendering, absence analysis, vouchers, and PDF reports.
- `constancia.js` generates the global absence certificate workbook using ExcelJS and JSZip.
- `data.json` is the runtime data source and must remain beside `index.html`.
- `TEMPLATE_GLOBALES.xlsx` is the workbook template used by certificate generation. Other `.xlsx` files are operational source data, not application code.

Keep browser code dependency-free unless a build system is justified. For substantial features, prefer a focused new script over further expanding `cs.js`.

## Build, Test, and Development Commands

There is no compilation or package installation step. Serve the directory over HTTP because `fetch()` may not work from a `file://` URL:

```bash
python3 -m http.server 8000
```

Open `http://localhost:8000`. Check console and network errors in developer tools. Before committing, run `git diff --check` and review `git diff`.

## Coding Style & Naming Conventions

Follow the existing JavaScript style: `'use strict'`, two-space indentation, semicolons, `const` by default, and `let` only for reassigned state. Use `camelCase` for functions and variables, `UPPER_SNAKE_CASE` for constants, and underscore prefixes only for module-level internal state (for example, `_eventosCache`). Preserve Spanish domain terminology in UI text and business logic. Escape user- or data-derived HTML with `esc()` before interpolation. Use kebab-case CSS class names and reuse existing custom properties such as `var(--acc)`.

## Testing Guidelines

No automated test framework or coverage threshold is configured. Manually verify search by RFC/name/card number, report dialogs, PDF generation, Excel uploads, and certificate export. Test empty, malformed, and duplicate records; confirm generated files open correctly. Check desktop and narrow layouts.

## Commit & Pull Request Guidelines

History uses short Spanish, imperative summaries, such as `Generar Constancia Global...`; automated data refresh commits follow `Actualizar data.json automáticamente - <fecha>`. Keep functional commits separate from large `data.json` or workbook updates. Pull requests should explain the user-visible change, list manual checks, link the relevant issue, and include screenshots for UI changes. Call out modified templates or operational datasets explicitly.

## Security & Data Handling

Treat JSON and spreadsheets as sensitive personnel data. Do not add real records to screenshots, logs, fixtures, or issues. Avoid committing temporary exports or replacing workbook templates without validating formatting and formulas.
