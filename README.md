# % of Completion Builder

A local-first single-page web app that replaces AllPro Painters' manual Excel workflow for the weekly
percent-completion (piece-rate) payroll report.

Upload the week's 6 report files (plus last week's finished workbook, optional), click **Create Report**,
review/edit the grid, and get two color-coded outputs — **TimeNavi Export** and **Pivot by Crew** — rendered
on screen (screenshot-ready) and downloadable as one styled `.xlsx`.

## Privacy

Payroll and sales data is sensitive: **everything runs client-side in your browser.**
No backend, no network calls, no telemetry. Files never leave your computer. The working session
(uploads + manual edits) is cached in this browser's own storage so a crash does not lose it; that
cache is local too, and **Start fresh** in the left rail clears it.

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

Last % Paid is the last percentage the job was actually paid at. It comes from last week's workbook
(slot 8, its `Pivot by Job` % of Completion for the AP#) and, failing that, from the calendar
Description: the most recent % below this week's line, skipping "no pay" weeks, 0% and HOURLY weeks.
When the only earlier entries are HOURLY it reads `HOURLY`; a Description with only this week's
entry leaves it blank. A 0% in slot 8 is treated as a no-pay week, so the calendar answers instead.

On **Pivot by Job**, Notes (default `*PAID XX% LAST WK*`), WO Override, WO% (HOURLY jobs included),
% of Completion and Estimator are editable on every row, including the no-hours rows below the Grand
Total. The **Notes** and **On Calendar?** columns take typed text in every cell — job rows, customer
rows, employee rows and the Grand Total row. On a job row, clearing a typed On Calendar? brings the
calendar lookup back. A row added by hand takes its estimator's color as soon as one is typed.

Utah timesheet hours (slot 2) show under **UT PAYROLL**: each Utah job with hours gets its hours by
day and one row per employee, the same way a Nevada job does in the pivot above. They stay out of
the NV tree and its Grand Total, because the payroll side pays Utah from that section. A blank WO Override shows `suggest $X` (job total × the WO%) as a hint; the hint is not written
to the .xlsx. Every typed value is written to the download — the integration suite exports a set of
edits, reads the file back, and re-reads it as next week's slot 8 to prove it.

On the **Payroll Preview**, each value row has a memo cell beside it, and the T-SHEETS totals obey
total + taxes + NEST = net: type the net and the taxes fill in (shown as `-150.00`), or type the
taxes and the net follows. The workbook carries the same thing as a live formula.

## When one input changes mid-session

A corrected timesheet, a fixed calendar export or a re-run P&L does **not** mean re-uploading
everything and starting over. Drop the one file that changed into its slot. Only that slot goes
amber and says *pending* — every other slot keeps the file it already has, and the report on screen
does not move until you ask it to.

- **Update existing report** re-derives from the new file and keeps every manual edit.
- **Create new report** is the explicit start-over. It clears the manual edits on that rail, and
  says how many before it does.

The update only asks about one thing: a cell you hand-edited whose underlying value the new file
would change. It shows your edit, what it was, and what the new file says, and you either keep yours
or take the new one — all at once or row by row. Values you never touched are applied without a
prompt, because there is nothing there to decide. The summary line afterwards (*Updated from … · N
values changed*) opens a list of every cell the update touched — tab, AP#, column and what happened,
e.g. *TimeNavi Export · 24403 · % Completion · 77% changed to 35%* — with a filter box. Edits whose row is not in the
new file are listed separately; keeping one parks it, and it re-attaches if that row comes back.

Under the hood this runs the builders a second time with empty edit maps (`src/lib/shadowBuild.ts`),
which re-exposes the computed layer they normally fold the edits into. Diffing those two shadow
models is what makes "only what this file changed" an honest claim rather than a guess.

## If the tab dies

Uploads and edits are cached in this browser (IndexedDB) as you work, so a crash, a browser restart
or an accidental refresh comes back to the same report — there is a "Session restored" note and a
**Start fresh** button in the left rail. Closing a tab that holds manual edits also asks first.

This is browser-local only; it never goes to a server, and the privacy note above still holds. It is
a safety net, not a guarantee: a private window may refuse the cache, and a write in flight when the
tab dies is lost.

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
| `Davis Bacon` | the prevailing rate in Settings ($19.18) |
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
  (a crew lead's own hours on profit-share jobs), comm (`Commercial Hourly`) and wg2 (`Wages 2`);
  always listed A→Z
- Crew lead nicknames — a calendar parenthetical like `(Pepe)` or `(Andy)` → the crew lead it means
  (case-insensitive)
- NEST [IRA] participants — only these people get the NEST column on their T-SHEETS block
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
# 378 unit assertions
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
