# % of Completion Builder

A local-first single-page web app that replaces AllPro Painters' manual Excel workflow for the weekly
percent-completion (piece-rate) payroll report.

Upload the week's 6 report files (plus last week's finished workbook, optional), click **Create Report**,
review/edit the grid, and get two color-coded outputs — **TimeNavi Export** and **Pivot by Crew** — rendered
on screen (screenshot-ready) and downloadable as one styled `.xlsx`.

## Privacy

Payroll and sales data is sensitive: **everything runs client-side in your browser.**
No backend, no network calls, no telemetry. Files never leave your computer.

## Run it

```bash
npm install
npm run dev
```

Then open the printed local URL (usually http://localhost:5173).

`npm run build` produces a static bundle in `dist/` you can open from any static file server.

## The 7 upload slots

| # | File | Required |
|---|------|----------|
| 1 | QuickBooks Time (TSheets) weekly timesheet | yes |
| 2 | TimeNavi calendar export (reads the `Events` sheet) | yes |
| 3 | NV "Sales by Rep Detail" (QuickBooks Online) | yes |
| 4 | NV "Sales by Customer Summary" | yes |
| 5 | UT "Transaction List by Date" (AllPro Painters LLC) | yes |
| 6 | UT "Sales by Customer Summary" | yes |
| 7 | Last week's finished "% of Completion" workbook (`TimeNavi Export` sheet) | optional |

Every slot accepts `.xlsx` and `.csv`. Header rows are found by scanning the first 10 rows, so QBO title
rows and layout drift are tolerated; parse problems show up as a friendly error on the upload chip.

## What you can edit after Create Report

- **Include/Exclude** toggle per event (events without an AP# are excluded by default)
- **AP#** (column H) — fixing a wrong calendar AP# re-runs every lookup
- **% Completion** (column I) — pre-filled from the Description history when possible (`95`, `95%`, `HOURLY`, `-`)
- **Estimator** (column J) — dropdown override
- **Apptivo $** (column O) — defaults to Job Total $; clear the cell to reset

Everything else (Last % Paid, QBO Job Name, In QBO Time?, Job Total $, $ Difference, Invoice Date,
Date Difference, the pivot, and the Exceptions list) recomputes live.

## Settings (gear icon)

- Estimator color legend (hex fills, add/remove) — the exact colors used for row fills on screen and in the export
- Estimator order — drives row sorting and pivot column order
- Persisted to `localStorage`

The pay date (left rail) defaults to the timesheet's week-ending Friday + 7 days and drives the pivot
title (`NV PAYROLL: M/D/YYYY`) and the export filename (`% MM-DD-YYYY % of Completion.xlsx`).

## Code layout

Pure, unit-testable logic lives in `src/lib/` (no React state):

- `parsers.ts` — SheetJS reading, header scanning, date/number coercion (incl. `"=1000"` QuickBooks formula remnants)
- `apkey.ts` — the canonical AP-number key every join uses
- `joins.ts` — sales indexes + timesheet enrichment (estimator, not-on-calendar)
- `timenavi.ts` — Description → % suggestion parser and the TimeNavi Export row builder
- `pivot.ts` — the static Pivot by Crew replica (with the >40h overtime flag)
- `exceptions.ts` — the 9 exception groups
- `exportXlsx.ts` — the styled ExcelJS workbook
- `__selftest__.ts` — dev-mode assertions (apKey extraction, % parser, dates, formula strings); results log to the console when `npm run dev` starts

React components live in `src/components/`; new output tabs can be added without touching the parsers.
