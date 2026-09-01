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

## The two processes are independent

The app has two rails, and each one's outputs come **only** from its own uploads:

- **%'s Report** — slots 1–8 (timesheets, TimeNavi calendar, sales/summary files, last week's workbook).
- **Payroll** — Slot A (this week's completed %'s workbook: its `Pivot by Job` + `timesheets` sheets)
  plus the P&L slots. The payroll tabs read the AP#, Estimator, Crew Lead, WO Override, WO% and
  % of Completion straight off Slot A's `Pivot by Job`, and their hours off Slot A's `timesheets`
  sheet — never from the %'s report's output.

Jobs listed **below the Grand Total** on `Pivot by Job` (no timesheet hours this week) pay through
their Crew Lead's band, unless the row is painted the A6A6A6 gray that means "keep off payroll".

## How hours are paid

The timesheet's own `payroll item` column decides both the rate and the overtime:

| Pay item | Rate used |
|---|---|
| `Regular Pay` | hourly-jobs rate on an hourly job, else the profit-share rate |
| `Crew Lead - Hourly Wage` | the crew-lead rate (profit-share jobs) |
| `Commercial Hourly` | the person's Commercial Hourly rate |
| `Wages 2` | the person's Wages 2 rate |
| `Overtime Pay` / `Crew Lead - Hourly Wage Overtime` / `Daily Overtime` | 1.5× the matching rate |

Only the overtime rows are paid at 1.5× — the rest of that job's hours stay at straight time. When a
week carries no overtime pay items at all, the app falls back to computing overtime itself
(NV: over 8 hrs/day under the rate threshold; anyone: over 40 hrs/week).

Two painters on the same job: whoever clocked `Crew Lead - Hourly Wage` is the crew lead for that
job, and the band it pays into is whatever the `Crew Lead` column on `Pivot by Job` says.

## Settings (gear icon)

- Estimator color legend (hex fills, add/remove) — the exact colors used for row fills on screen and in the export
- Estimator order — drives row sorting and pivot column order
- Wage rates — five columns per person: PS (profit-share/work-order), hrly (T&M jobs), lead
  (a crew lead's own hours on profit-share jobs), comm (`Commercial Hourly`) and wg2 (`Wages 2`)
- Payroll line items — per-person extra rows on the Payroll Preview (Aflac, Health Insurance,
  Gas Reimb, Reimbursements), each a deduction, a reimbursement or a labeled blank note, weekly or
  only on the first pay date of the month (a pay date on the 1st–7th)
- Salaried roster — section, entity, weekly vs monthly, and a "sub" flag that sorts the row to the
  bottom of its section
- Full-materials leads — crew leads charged for **all** of a job's materials instead of half
- Persisted to a shared store (`/api/settings`, Vercel Blob) with `localStorage` as the offline cache,
  so a different browser or a cleared profile still sees the same saved Settings. Blob setup is optional —
  without it the app silently keeps using `localStorage` only.

The pay date (left rail) defaults to the timesheet's week-ending Friday + 7 days and drives the pivot
title (`NV PAYROLL: M/D/YYYY`) and the export filename (`MM-DD-YYYY Percentages of Completion.xlsx`).

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

### Running the two test suites outside the browser

Both harnesses compile to gitignored folders, so build them fresh before you
run them — an old build silently passing is the one trap here.

```sh
# 306 unit assertions
rm -rf .selftest && npx tsc src/lib/__selftest__.ts --outDir .selftest \
  --module commonjs --target es2022 --moduleResolution node \
  --skipLibCheck --esModuleInterop --strict false
echo '{"type":"commonjs"}' > .selftest/package.json
node -e "require('./.selftest/__selftest__.js').runSelfTests()"

# end-to-end: build synthetic workbooks, export, read the .xlsx back
rm -rf .itest && npx tsc scripts/integration.ts --outDir .itest \
  --module commonjs --target es2022 --moduleResolution node \
  --skipLibCheck --esModuleInterop --strict false
echo '{"type":"commonjs"}' > .itest/package.json
node .itest/scripts/integration.js
```
