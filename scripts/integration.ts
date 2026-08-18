/**
 * End-to-end pipeline test run in Node (no browser): builds synthetic versions
 * of all 7 input files in memory with SheetJS, pushes them through the exact
 * same parse → join → report → pivot → exceptions pipeline the app uses, and
 * asserts the behaviors from the spec (iteration 3).
 *
 * Run:  node node_modules/typescript/bin/tsc scripts/integration.ts --outDir .itest \
 *         --module commonjs --moduleResolution node --target es2020 --esModuleInterop --skipLibCheck
 *       node .itest/scripts/integration.js   (with .itest/package.json = {"type":"commonjs"})
 */
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { timesheetApKey } from '../src/lib/apkey';
import { buildPercentsWorkbookBuffer, buildPayrollWorkbookBuffer } from '../src/lib/exportXlsx';
import { defaultSettings, DEFAULT_LEGEND, DEFAULT_PIVOT_COLUMNS } from '../src/lib/settings';
import { buildExceptions } from '../src/lib/exceptions';
import { enrichTimesheet } from '../src/lib/joins';
import { buildJobsByCrew } from '../src/lib/jobsByCrew';
import {
  buildSalesByRepMirrorFromParsed,
  buildTimesheetMirror,
} from '../src/lib/mirrorTabs';
import {
  parseCustomerSummary,
  parseEmployeeList,
  parseLastWeek,
  parseLastWeekPivotByJob,
  parsePl,
  parseSalesDetail,
  parseTimeNavi,
  parseTimesheet,
  readWorkbook,
} from '../src/lib/parsers';
import { buildPayrollExceptions } from '../src/lib/payrollExceptions';
import { buildPayrollPreview } from '../src/lib/payrollPreview';
import { buildPivot } from '../src/lib/pivot';
import { buildPivotByJob, EMPTY_MANUAL_STATE } from '../src/lib/pivotByJob';
import { round2 } from '../src/lib/format';
import { buildReportRows, sortReportRows } from '../src/lib/timenavi';
import type { CellValue, EnrichedTimesheetRow } from '../src/lib/types';

let failures = 0;
function check(name: string, actual: unknown, expected: unknown): void {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    console.log(`  ok    ${name}`);
  } else {
    failures++;
    console.error(`  FAIL  ${name}: expected ${e}, got ${a}`);
  }
}

const isBlank = (v: unknown): boolean => v === null || v === undefined || v === '';
const fillOf = (cell: ExcelJS.Cell): string | null => {
  const f = cell.fill as ExcelJS.FillPattern | undefined;
  return f && f.type === 'pattern' && f.fgColor?.argb ? f.fgColor.argb : null;
};

function wb(sheets: [string, CellValue[][]][]): ReturnType<typeof readWorkbook> {
  const book = XLSX.utils.book_new();
  for (const [name, rows] of sheets) {
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
  }
  const buf = XLSX.write(book, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
  return readWorkbook(buf);
}

// ---------------------------------------------------------------------------
// Synthetic input files
// ---------------------------------------------------------------------------

const TS_HEADERS = [
  'username', 'payroll_id', 'fname', 'lname', 'number', 'group', 'local_date', 'local_day',
  'local_start_time', 'local_end_time', 'tz', 'hours', 'jobcode_1', 'jobcode_2', 'billable',
  'class', 'department', 'payroll item', 'service item', 'location', 'notes', 'approved_status',
  'has_flags', 'flag_types',
];

function tsRow(username: string, fname: string, lname: string, group: string, date: string, hours: number, job1: string, job2: string): CellValue[] {
  return [username, '', fname, lname, '0', group, date, 'Mon', `${date} 7:00`, `${date} 10:00`, '-7', hours, job1, job2, 'No', '', '', 'Regular Pay', '', '', '', 'Approved', 'FALSE', ''];
}

const timesheetWb = wb([
  ['Sheet1', [
    TS_HEADERS,
    tsRow('leo@x.com', 'Leonardo', 'Martinez-Rincon', "Rene Salazar's Crew", '6/29/2026', 3, 'Aaron Ward', 'AP24314 Ext'),
    tsRow('leo@x.com', 'Leonardo', 'Martinez-Rincon', "Rene Salazar's Crew", '6/30/2026', 39, 'Aaron Ward', 'AP24314 Ext'),
    tsRow('juan@x.com', 'Juan', 'Perez', "Rene Salazar's Crew", '7/1/2026', 5, 'CCSD', 'AP24403 5321 Lambrook Dr Int/Ext'),
    tsRow('leo@x.com', 'Leonardo', 'Martinez-Rincon', "Rene Salazar's Crew", '7/1/2026', 0.5, 'Aaron Ward', 'Lunch'),
    tsRow('nancy@x.com', 'Nancy', 'Duarte', 'Office (NV)', '7/2/2026', 6, 'ALLPRO OFFICE', 'OFFICE [Please put AP# in notes]'),
    tsRow('juan@x.com', 'Juan', 'Perez', "Rene Salazar's Crew", '7/3/2026', 2, 'Mystery Co', 'AP99999 Mystery Job'),
    // Hours for the job whose estimator resolves ONLY via the title fallback:
    tsRow('juan@x.com', 'Juan', 'Perez', "Rene Salazar's Crew", '7/2/2026', 4, 'Southpoint', 'AP24450 Southpoint Framing'),
    // Hours for the RF job (books-resolved):
    tsRow('juan@x.com', 'Juan', 'Perez', "Rene Salazar's Crew", '7/3/2026', 3, 'Roof Co', 'AP24500 Roof Ext'),
  ]],
]);

const TN_HEADERS = ['Event', 'Start Date', 'End Date', 'Calendar', 'Attendees', 'Location', 'Description'];
const timenaviWb = wb([
  // Decoy sheet of manually removed events — must be ignored.
  ['Sheet1', [TN_HEADERS, ['REMOVED EVENT AP11111', '6/1/2026', '6/1/2026', 'Commercial', '', '', '']]],
  ['Events', [
    TN_HEADERS,
    // Description carries NEXT week's note on top (calendar updated after the
    // pay week) — the pay-date line must win, not the topmost line.
    ['♦Int 13k (XC) Regina Conrad Ph1, AP24314 *PAID* (Daniel)', '6/29/2026 7:00', '7/1/2026 16:00', 'Residential (NV)', '', '123 Main St', '7/17/26 100% Complete\n7/10/26 30% COMPLETE -SH\n7/03/26 90% Complete'],
    ['AP24222 SNRHA *SS* [PO48670] (JS)', '6/30/2026 7:00', '7/2/2026 16:00', 'Utah Jobs', '', '', '7/8/26 90% Complete'],
    ['AP24388 Chavez Ext', '7/1/2026 7:00', '7/1/2026 16:00', 'Commercial', '', '', '7/17/26 HOURLY -SH'],
    ['Paint shop maintenance', '7/1/2026 7:00', '7/1/2026 16:00', 'Commercial', '', '', 'no status here'],
    // Round 3: description carries a WO Override line + an "@ NN% PS" line.
    ['AP24403 Lambrook (PG) (Daniel)', '8/1/2026 7:00', '8/2/2026 16:00', 'Residential (NV)', '', '', 'gate code 4321\n****PER JODI $4,715.00 is the total of work order he completed****\n7/03/26 10% COMPLETE @ 40% PS'],
    // Estimator missing from BOTH sales files — must resolve via the title
    // fallback: lowercase "(te)" tag; "(Ted/Rafael)" must never match.
    ['♦[Southpoint] 3.5k (te) Framing/DW Install, AP24450 (Ted/Rafael)', '7/1/2026 7:00', '7/2/2026 16:00', 'Commercial', '', '', '7/10/26 50% Complete'],
    // RF job (books-resolved) so the new palette entry exercises both tabs.
    ['AP24500 Roof Ext', '7/2/2026 7:00', '7/2/2026 16:00', 'Residential (NV)', '', '', '7/10/26 60% Complete'],
  ]],
  ['Summary', [['Generated by TimeNavi'], ['Range', '6/27/2026 - 7/3/2026']]],
]);

// NV Sales by Rep Detail: Column A blank, combined Customer string, NO Name 1/2
// columns (parser must derive them), oldest → newest, plus a Total junk row.
const nvSalesWb = wb([
  ['Sales by Customer Detail', [
    ['', 'Date', 'Customer', 'Transaction Type', 'Num', 'Product/Service', 'Memo/Description', 'Sales Price', 'Amount', 'Balance', 'Sales Rep', ''],
    ['', '6/10/2026', 'Chavez:AP24388 Chavez Ext', 'Invoice', '1001', 'Painting', '', '', '=1000', '', 'XC', ''],
    ['', '6/20/2026', 'CCSD:AP24403 5321 Lambrook Dr Int/Ext', 'Invoice', '1002', 'Painting', '', '', 2498, '', 'PG', ''],
    ['', '6/25/2026', 'Aaron Ward:AP24314 Ext', 'Invoice', '1003', 'Painting', '', '', 650, '', 'MH', ''],
    ['', '6/26/2026', 'Roof Co:AP24500 Roof Ext', 'Invoice', '1004', 'Painting', '', '', 3200, '', 'RF', ''],
    ['', '7/2/2026', 'Aaron Ward:AP24314 Ext', 'Invoice', '1005', 'Painting', '', '', 100, '', 'MH', ''],
    ['', '', 'Total', '', '', '', '', '', '=(I2)+(I3)', '', '', ''],
  ]],
]);

// NV Sales by Customer Summary: hierarchical; job totals include "=650" strings.
const nvSummaryWb = wb([
  ['Sales by Customer Summary', [
    ['', 'Total'],
    ['Aaron Ward', ''],
    ['   AP24314 Ext', '=650'],
    ['Total Aaron Ward', '=(B3)'],
    ['CCSD', ''],
    ['   AP24403 5321 Lambrook Dr Int/Ext', 2498],
    ['Total CCSD', '=(B6)'],
    ['Roof Co', ''],
    ['   AP24500 Roof Ext', 3200],
    ['Total Roof Co', '=(B9)'],
    // AP24388 and AP24450 deliberately missing → "#N/A" job totals.
  ]],
]);

// UT Transaction List by Date: Name combined + Name 1 / Name 2 present.
const utSalesWb = wb([
  ['Transaction List by Date', [
    ['', 'Date', 'Name', 'Name 1', 'Name 2', 'Transaction Type', 'Num', 'Split', 'Memo/Description', 'Account', 'Amount', 'Open Balance', 'Sales Rep', ''],
    ['', '6/23/2026', 'SNRHA:AP24222 Ext', 'SNRHA', 'AP24222 Ext', 'Invoice', '2001', '', '', '', 14410.63, '', 'JS', ''],
  ]],
]);

// UT Sales by Customer Summary: 3 title rows before the header (scan handles it).
const utSummaryWb = wb([
  ['Sales by Customer Summary', [
    ['AllPro Painters LLC'],
    ['Sales by Customer Summary'],
    ['June 27 - July 3, 2026'],
    ['', 'Total'],
    ['SNRHA', ''],
    ['   AP24222 Ext', 14410.63],
    ['Total SNRHA', '=(B6)'],
  ]],
]);

// Last week's finished workbook (optional slot 7). Includes a suffixed AP key
// that must round-trip intact (never collapse to the base number).
const lastWeekWb = wb([
  ['TimeNavi Export', [
    ['Event', 'Start Date', 'End Date', 'Calendar', 'Attendees', 'Location', 'Description', 'AP#', '% Completion'],
    ['old event', '6/22/2026', '6/23/2026', 'Commercial', '', '', '', '24314', 0.9],
    ['old event 2', '6/22/2026', '6/23/2026', 'Utah Jobs', '', '', '', '24222', 'HOURLY'],
    ['old event 3', '6/22/2026', '6/23/2026', 'Commercial', '', '', '', '21848-A5', 0.75],
  ]],
  ['Pivot by Crew', [['NV PAYROLL: 7/3/2026']]],
  // Round 3 §4.6: last week's Pivot by Job — one NV job partially paid with no
  // hours this week (→ BALANCES), one fully paid (dropped), one UT-section row.
  ['Pivot by Job', [
    ['NV PAYROLL: 7/3/2026', '', 'Total hrs', '', 'AP#', 'On Calendar?', 'Estimator', 'Crew Lead', 'WO Override', 'WO%', '% of Completion', 'Notes'],
    ['Sum of hours'],
    ['Row Labels'],
    ['Old Customer'],
    ['AP21848-A5 Somewhere Int', '', 12, '', '21848-A5', 'TimeNavi row 9', 'MH', 'CARLOS RUIZ CRUZ', 5000, 0.38, 0.75, 'wk1 note'],
    ['AP24314 Ext', '', 8, '', '24314', 'TimeNavi row 3', 'MH', 'DANIEL GONZALES BONILLA', '', 0.38, 1, ''],
    ['Grand Total', '', 20],
    ['NV PAYROLL (BALANCES): 7/3/2026'],
    ['Older Customer'],
    ['AP20000 Ancient Job', '', '', '', '20000', '', 'PG', 'RENE SALAZAR', 1200, 0.38, 0.5, 'still open'],
    ['UT PAYROLL: 7/3/2026'],
    ['Utah Customer'],
    ['AP24111 Provo Ext', '', '', '', '24111', '', 'JS', "PEPE'S PAINTING LLC", 900, 0.5, 0.6, 'sub job'],
  ]],
]);

// Round 3 slot 8 — raw QBO P&L by Customer (wide): 3 title rows, jobs as
// columns (customer + "Total …" columns skipped), a duplicated AP column,
// "=650" formula remnants, indented account labels, a footer line.
const plWb = wb([
  ['Profit and Loss by Customer', [
    ['AllPro Painters'],
    ['Profit and Loss by Customer'],
    ['June 27 - July 3, 2026'],
    ['', 'Aaron Ward', 'AP24314 Ext', 'Total Aaron Ward', 'AP24403 5321 Lambrook Dr Int/Ext', 'AP24450 Southpoint', 'AP24450 Southpoint (dup)', 'AP24500 Roof Ext'],
    ['Income', '', '', '', '', '', '', ''],
    ['   Construction Income', 650, 650, 650, 9000, 100, 50, 3200],
    ['Total Income', 650, '=650', 650, 9000, 100, 50, 3200],
    ['Expenses', '', '', '', '', '', '', ''],
    ['   Job Materials Purchased', 20, 20, 20, 500, 0, 0, 0],
    ['   Paint', 15.74, 15.74, 15.74, 260.1, 0, 0, 0],
    ['   Sundries -', 0, 0, 0, 0, 0, 0, 58.51],
    ['   Employee Labor', 113.04, 113.04, 113.04, 2081.09, 0, 0, 0],
    ['   Wages', 0, 0, 0, 0, 0, 0, 0],
    ['   Mileage Expense', 0, 0, 0, 0, 0, 0, 0],
    ['Friday, Jul 17, 2026 08:00 AM GMT-07:00', '', '', '', '', '', '', ''],
  ]],
]);

// Round 3 slot 9 — QB Time employee directory ("Last, First"; '*' = inactive).
const employeesWb = wb([
  ['Report', [
    ['Employee directory report'],
    ['Name', 'First Name', 'Last Name'],
    ['Martinez-Rincon, Leonardo', 'Leonardo', 'Martinez-Rincon'],
    ['Perez, Juan', 'Juan', 'Perez'],
    ['Duarte, Nancy', 'Nancy', 'Duarte'],
    ['Gonzales Bonilla, Daniel', 'Daniel', 'Gonzales Bonilla'],
    ['Frank, Ted', 'Ted', 'Frank'],
    ['Aguilar, Rafael', 'Rafael', 'Aguilar'],
    ['*Climaco, Calixto  Efrain', 'Calixto  Efrain', '*Climaco'],
  ]],
]);

// ---------------------------------------------------------------------------
// Run the pipeline exactly like the app does
// ---------------------------------------------------------------------------

console.log('parsing…');
const timesheet = parseTimesheet(timesheetWb);
const events = parseTimeNavi(timenaviWb);
const nvSales = parseSalesDetail(nvSalesWb, 'customer');
const nvTotals = parseCustomerSummary(nvSummaryWb);
const utSales = parseSalesDetail(utSalesWb, 'name');
const utTotals = parseCustomerSummary(utSummaryWb);
const lastWeek = parseLastWeek(lastWeekWb);

check('timesheet rows kept', timesheet.rows.length, 7);
check('lunch rows removed', timesheet.lunchRemoved, 1);
check('events parsed from Events sheet only', events.length, 7);
check('decoy Sheet1 ignored', events.some((e) => e.event.includes('REMOVED')), false);
check('nv sales rows (Total row skipped)', nvSales.length, 5);
check('nv sales Name2 derived from Customer', nvSales[2].name2, 'AP24314 Ext');
check('nv formula-remnant amount row still parsed (apKey)', nvSales[0].apKey, '24388');
check('nv totals: "=650" parsed', nvTotals.get('24314'), 650);
check('nv totals: subtotal formula rows skipped', nvTotals.size, 3);
check('ut totals found after title rows', utTotals.get('24222'), 14410.63);
check('last week map values as-is', [lastWeek.get('24314'), lastWeek.get('24222')], [0.9, 'HOURLY']);
check('last week suffixed key intact (…-A5)', lastWeek.get('21848-A5'), 0.75);
check('last week suffixed key does NOT collapse to base', lastWeek.get('21848'), undefined);

// Timezone-safe dates: no time component anywhere.
check(
  'event start date has no time component',
  events[0].startDate ? [events[0].startDate.getHours(), events[0].startDate.getMinutes()] : null,
  [0, 0],
);
check(
  'sales date has no time component',
  nvSales[0].date ? [nvSales[0].date.getHours(), nvSales[0].date.getMinutes()] : null,
  [0, 0],
);

// Default pay date = max local_date (7/3/2026) + 7 days = 7/10/2026.
let maxDate: Date | null = null;
for (const r of timesheet.rows) {
  if (r.localDate && (!maxDate || r.localDate.getTime() > maxDate.getTime())) maxDate = r.localDate;
}
const payDate = new Date(maxDate!.getFullYear(), maxDate!.getMonth(), maxDate!.getDate() + 7);
check('default pay date', [payDate.getMonth() + 1, payDate.getDate(), payDate.getFullYear()], [7, 10, 2026]);

const timesheetJobcodeByApKey = new Map<string, string>();
for (const r of timesheet.rows) {
  const key = timesheetApKey(r.jobcode2);
  if (key && !timesheetJobcodeByApKey.has(key)) timesheetJobcodeByApKey.set(key, r.jobcode2);
}

const knownEstimators = new Set<string>([...DEFAULT_LEGEND.map((l) => l.estimator), ...DEFAULT_PIVOT_COLUMNS]);
const deps = { nvSales, utSales, nvTotals, utTotals, lastWeekMap: lastWeek, timesheetJobcodeByApKey, knownEstimators };
const legendOrder = DEFAULT_LEGEND.map((l) => l.estimator);
const rows = sortReportRows(buildReportRows(events, {}, deps, payDate), legendOrder);

const byAp = (k: string) => rows.find((r) => r.apKey === k)!;

console.log('report rows…');
check('row count', rows.length, 7);
check('no-AP event excluded by default', rows.find((r) => r.apKey === null)!.included, false);
check('AP events included by default', rows.filter((r) => r.included).length, 6);

const r24314 = byAp('24314');
check('24314 estimator (UT first, then NV)', r24314.estimator, 'MH');
// B-1: pay-date line (7/10 → 30%) wins over the newer 7/17 100% line on top.
check('24314 suggested % is the PAY DATE line (30%, not 100%)', r24314.percent, 0.3);
check('24314 % source line', r24314.suggestedSourceLine, '7/10/26 30% COMPLETE -SH');
check('24314 last % paid (from workbook only)', r24314.lastPaid, 0.9);
check('24314 QBO job name (NV first)', r24314.qboJobName, 'AP24314 Ext');
check('24314 In QBO Time', r24314.inQboTime, 'AP24314 Ext');
check('24314 job total', r24314.jobTotal, 650);
// Round 8 §2.2 — Apptivo is never auto-filled from the job total; it stays
// blank until someone types it, so the $ Difference stays blank too.
check('24314 apptivo blank until typed', r24314.apptivo, null);
check('24314 $ difference blank while apptivo is blank', r24314.dollarDiff, null);
check('24314 invoice date = FIRST (earliest) match', [r24314.invoiceDate!.getMonth() + 1, r24314.invoiceDate!.getDate()], [6, 25]);
check('24314 date difference (R − B, whole days)', r24314.dateDiff, -4);
check('24314 title initials for mismatch check', r24314.titleInitials, 'XC');

const r24222 = byAp('24222');
check('24222 estimator resolves via UT (JS)', r24222.estimator, 'JS');
check('24222 job total from UT summary', r24222.jobTotal, 14410.63);
check('24222 NotClockedIn (no Utah crew in timesheet)', r24222.inQboTime, 'NotClockedIn');
check('24222 last % paid HOURLY (text as-is)', r24222.lastPaid, 'HOURLY');
check('24222 % from latest line ≤ pay date (7/8 → 90%)', r24222.percent, 0.9);

const r24388 = byAp('24388');
// B-1: the only status line (7/17) is AFTER pay date 7/10 → no suggestion.
check('24388 no suggestion (only line is after pay date)', r24388.percent, null);
check('24388 estimator', r24388.estimator, 'XC');
// B-3: unmatched job total renders like Excel.
check('24388 job total = #N/A (not 0)', r24388.jobTotal, '#N/A');
check('24388 apptivo blank (not 0)', r24388.apptivo, null);
check('24388 $ difference blank (not 0)', r24388.dollarDiff, null);

// §1.3: books blank → estimator pulled from the "(te)" title tag, uppercased;
// "(Ted/Rafael)" never matches.
const r24450 = byAp('24450');
check('24450 estimator via title fallback (te → TE)', r24450.estimator, 'TE');
check('24450 % pre-fill', r24450.percent, 0.5);
check('24450 job total = #N/A (missing from summaries)', r24450.jobTotal, '#N/A');
check('24450 In QBO Time', r24450.inQboTime, 'AP24450 Southpoint Framing');

const r24500 = byAp('24500');
check('24500 estimator RF from the books', r24500.estimator, 'RF');
check('24500 job total', r24500.jobTotal, 3200);

check('sorted by legend order (JS first)', rows[0].apKey, '24222');
check(
  'row order follows legend (JS, MH, PG, TE, XC, RF)',
  rows.filter((r) => r.included).map((r) => r.apKey),
  ['24222', '24314', '24403', '24450', '24388', '24500'],
);

// E-3: same TimeNavi file, pay date 7/17 → different suggestions.
const rows717 = buildReportRows(events, {}, deps, new Date(2026, 6, 17));
check('pay 7/17: 24314 suggests 100%', rows717.find((r) => r.apKey === '24314')!.percent, 1);
check('pay 7/17: 24388 suggests HOURLY', rows717.find((r) => r.apKey === '24388')!.percent, 'HOURLY');

// E-4: with no prior-week workbook, every K cell is '-' (null).
const rowsNoPrev = buildReportRows(events, {}, { ...deps, lastWeekMap: null }, payDate);
check('no prior workbook → all Last % Paid empty', rowsNoPrev.every((r) => r.lastPaid === null), true);

console.log('enrichment + pivot…');
const includedApKeys = new Set<string>();
const estimatorByApKey = new Map<string, string>();
for (const r of rows) {
  if (r.included && r.apKey) {
    includedApKeys.add(r.apKey);
    if (r.estimator && !estimatorByApKey.has(r.apKey)) estimatorByApKey.set(r.apKey, r.estimator);
  }
}
const enriched = enrichTimesheet(timesheet.rows, nvSales, utSales, includedApKeys, estimatorByApKey);

check('OFFICE → ND', enriched.find((r) => r.apKey === 'OFFICE')!.estimator, 'ND');
check('exact jobcode_2 match → MH', enriched.find((r) => r.jobcode2 === 'AP24314 Ext')!.estimator, 'MH');
check('unmatched job → #N/A', enriched.find((r) => r.apKey === '99999')!.estimator, '#N/A');
// §2.0: the RESOLVED (title-fallback) estimator drives pivot bucketing.
check('timesheet hours follow the resolved TE fallback', enriched.find((r) => r.apKey === '24450')!.estimator, 'TE');
check('99999 not on calendar', enriched.find((r) => r.apKey === '99999')!.notOnCalendar, true);
check('OFFICE never "not on calendar"', enriched.find((r) => r.apKey === 'OFFICE')!.notOnCalendar, false);

// Round 2 §4: the FULL configured set always renders (XC has zero hours this
// period yet still gets a column); RF sits immediately after TE; #N/A last.
const pivot = buildPivot(enriched, DEFAULT_PIVOT_COLUMNS, payDate);
check('pivot title', pivot.title, 'NV PAYROLL: 7/10/2026');
check('pivot estimator columns (full configured set, #N/A last)', pivot.estimators, ['MH', 'PG', 'XC', 'ND', 'TE', 'RF', '#N/A']);
check('zero-hour XC still renders', pivot.grandByEstimator['XC'], undefined);
check('RF immediately after TE', pivot.estimators.indexOf('RF'), pivot.estimators.indexOf('TE') + 1);
check('pivot grand total', pivot.grandTotal, 62);
check('pivot MH hours', pivot.grandByEstimator['MH'], 42);
check('TE hours moved out of #N/A by the fallback', pivot.grandByEstimator['TE'], 4);
check('#N/A only holds truly unresolved hours', pivot.grandByEstimator['#N/A'], 2);
check('pivot crews A→Z', pivot.crews.map((c) => c.crew), ['Office (NV)', "Rene Salazar's Crew"]);
const rene = pivot.crews[1];
// "Rene Salazar" matches no employee → the crew stays fully alphabetical.
check('no lead match → employees stay A→Z by username', rene.employees.map((e) => e.username), ['juan@x.com', 'leo@x.com']);
const juan = rene.employees[0];
const leo = rene.employees[1];
check('leo total', leo.total, 42);
check('leo overtime flag (>40)', leo.overtime, true);
check('juan not overtime', juan.overtime, false);
check('juan jobs sorted by (customer, jobcode_2)', juan.jobs.map((j) => j.customer), ['CCSD', 'Mystery Co', 'Roof Co', 'Southpoint']);
// A custom pivot order is respected, but missing defaults (XC/TE/RF) are
// healed in — RF always lands immediately after TE.
const pivotCustom = buildPivot(enriched, ['ND', 'PG', 'MH'], payDate);
check('custom pivot order healed with RF after TE', pivotCustom.estimators, ['ND', 'TE', 'RF', 'PG', 'XC', 'MH', '#N/A']);

// §5: A1 title always follows the selected pay date.
check('title follows the selected pay date', buildPivot(enriched, DEFAULT_PIVOT_COLUMNS, new Date(2026, 6, 17)).title, 'NV PAYROLL: 7/17/2026');

// Round 2 §3: crew lead first within each crew — verified against every case
// from the spec with dedicated rows (the group label is the only input).
console.log('crew-lead ordering…');
const fakeRow = (username: string, fname: string, lname: string, group: string): EnrichedTimesheetRow => ({
  username, payrollId: '', fname, lname, number: '', group,
  localDate: null, localDay: '', localStartTime: null, localEndTime: null, tz: '',
  hours: 1, jobcode1: 'Cust', jobcode2: 'AP24314 Ext', billable: '', klass: '',
  department: '', payrollItem: '', serviceItem: '', location: '', notes: '',
  approvedStatus: '', hasFlags: '', flagTypes: '',
  sourceEntity: 'NV',
  apKey: '24314', estimator: 'MH', notOnCalendar: false,
});
const leadPivot = buildPivot([
  fakeRow('ana@x.com', 'Ana', 'Bedoya Melgar', "Carlos Ruiz Cruz's Crew"),
  fakeRow('carlos@x.com', 'Carlos', 'Ruiz Cruz', "Carlos Ruiz Cruz's Crew"),
  fakeRow('armando@x.com', 'Armando', 'Lopez', "Daniel Bonilla's Crew"),
  fakeRow('daniel@x.com', 'Daniel', 'Gonzales Bonilla', "Daniel Bonilla's Crew"),
  fakeRow('al@x.com', 'Al', 'Alpha', 'Handyman'),
  fakeRow('bob@x.com', 'Bob', 'Zeta', 'Handyman'),
  fakeRow('sam@x.com', 'Samuel', 'Martin', 'Miguel Martin\u2019s Crew'),
  fakeRow('miguel@x.com', 'Miguel', 'Martin', 'Miguel Martin\u2019s Crew'),
  fakeRow('zed@x.com', 'Zed', 'Ames', "Rafael Aguilar's Crew"),
  fakeRow('rafael@x.com', 'Rafael', 'Aguilar Echevarria', "Rafael Aguilar's Crew"),
], DEFAULT_PIVOT_COLUMNS, payDate);
const crewEmps = (label: string) => leadPivot.crews.find((c) => c.crew === label)!.employees.map((e) => e.fname);
check('Carlos listed before Ana', crewEmps("Carlos Ruiz Cruz's Crew"), ['Carlos', 'Ana']);
check('Daniel Gonzales Bonilla precedes Armando', crewEmps("Daniel Bonilla's Crew"), ['Daniel', 'Armando']);
check('Miguel Martin (curly apostrophe) precedes Samuel', crewEmps('Miguel Martin\u2019s Crew'), ['Miguel', 'Samuel']);
check('extra surname match: Rafael Aguilar Echevarria first', crewEmps("Rafael Aguilar's Crew"), ['Rafael', 'Zed']);
check('Handyman (no lead) stays alphabetical', crewEmps('Handyman'), ['Al', 'Bob']);
check(
  'reordering never changes subtotals',
  leadPivot.crews.find((c) => c.crew === "Carlos Ruiz Cruz's Crew")!.employees.map((e) => e.total),
  [1, 1],
);

console.log('exceptions…');
const groups = buildExceptions({ reportRows: rows, enriched, lunchRemoved: timesheet.lunchRemoved, payDate });
const g = (key: string) => groups.find((x) => x.key === key)!;
check('excluded events', g('excluded').items.length, 1);
check(
  'not clocked in',
  g('notClockedIn')
    .items.map((i) => (i.message.match(/AP (\d+)/) as RegExpMatchArray)[1])
    .sort(),
  ['24222', '24388'],
);
check('not on calendar (AP99999)', g('notOnCalendar').items.length, 1);
check('estimator #N/A with hours', g('estimatorNA').items.length, 1);
check('mismatch: calendar (XC) vs books MH', g('mismatch').items.length, 1);
check('no duplicates this week', g('duplicates').items.length, 0);
check('date-range warning for 8/1 event', g('dateRange').items.length, 1);
check('job total not found (24388 + 24450)', g('noJobTotal').items.length, 2);
check('lunch count', g('lunch').items.length, 1);

// E-6: fixing an AP# re-runs every lookup for that row live.
console.log('edit simulation…');
const noApRow = rows.find((r) => r.apKey === null)!;
const rows2 = buildReportRows(events, { [noApRow.id]: { apKey: '24314' } }, deps, payDate);
const fixed = rows2.find((r) => r.id === noApRow.id)!;
check('edited AP# re-includes the row', fixed.included, true);
check(
  'edited AP# re-runs estimator/jobname/inqbotime/total/invoice',
  [fixed.estimator, fixed.qboJobName, fixed.inQboTime, fixed.jobTotal, fixed.invoiceDate!.getDate()],
  ['MH', 'AP24314 Ext', 'AP24314 Ext', 650, 25],
);
check(
  'duplicate AP# now detected',
  buildExceptions({ reportRows: rows2, enriched, lunchRemoved: 0, payDate }).find((x) => x.key === 'duplicates')!.items.length,
  1,
);

// ---------------------------------------------------------------------------
// Round 3 — slots 8/9, Pivot by Job, Payroll Preview, Payroll Exceptions
// ---------------------------------------------------------------------------

console.log('P&L parse (slot 8)…');
const pl = parsePl(plWb);
check('P&L job columns (customer/Total cols skipped, dup summed)', pl.byJob.size, 4);
check('P&L duplicate column flagged', pl.duplicateColumns, ['24450']);
const pl24314 = pl.byJob.get('24314')!;
check('24314 metrics (income via "=650" remnant)', [pl24314.income, pl24314.materials, pl24314.laborWagesSubs, pl24314.mileage], [650, 35.74, 113.04, 0]);
const pl24403 = pl.byJob.get('24403')!;
check('24403 metrics', [pl24403.income, round2(pl24403.materials), pl24403.laborWagesSubs], [9000, 760.1, 2081.09]);
check('duplicated 24450 columns summed', pl.byJob.get('24450')!.income, 150);
check('sundries tracked separately (24500)', pl.byJob.get('24500')!.sundries, 58.51);
// Round 7 — retain full P&L grid for PR Final Check / import mirror.
check('parsePl columnHeaders present', pl.columnHeaders.length > 1, true);
check(
  'parsePl columnHeaders include a job header',
  pl.columnHeaders.some((h) => /AP24314/i.test(h)),
  true,
);
check(
  'parsePl accountRows include Total Income',
  pl.accountRows.some((r) => /^total income$/i.test(r.label.trim())),
  true,
);
check('parsePl accountRows non-empty', pl.accountRows.length > 0, true);

console.log('employee directory (slot 9)…');
const directory = parseEmployeeList(employeesWb);
check('directory size', directory.length, 7);
check('active/inactive split', directory.filter((d) => d.active).length, 6);
const calixto = directory.find((d) => d.lastName === 'Climaco')!;
check('* stripped, doubled space collapsed, inactive', [calixto.fullName, calixto.active], ['Calixto Efrain Climaco', false]);

console.log('pivot by job…');
const settings = defaultSettings();
const lastWeekPbj = parseLastWeekPivotByJob(lastWeekWb);
check('last week Pivot by Job rows parsed', lastWeekPbj.length, 4);
const leadCtx = {
  directory,
  crewGroups: [...new Set(timesheet.rows.map((r) => r.group))],
  subcontractors: settings.subcontractors,
};
const pbj = buildPivotByJob({
  enriched,
  reportRows: rows,
  settings,
  leadCtx,
  knownEstimators,
  payDate,
  edits: {},
  qbHours: {},
  lastWeekPbj,
  balancesState: EMPTY_MANUAL_STATE,
  utState: EMPTY_MANUAL_STATE,
  pl,
  nvTotals,
  utTotals,
  nvSales,
  utSales,
});
check('pbj title', pbj.title, 'NV PAYROLL: 7/10/2026');
check('pbj distinct dates ascending (5, not assumed 7)', pbj.dates.map((d) => d.getDate()), [29, 30, 1, 2, 3]);
check('pbj grand total matches the crew pivot', pbj.grandTotal, 62);
check(
  'pbj customers A→Z (case-insensitive)',
  pbj.customers.map((c) => c.customer),
  ['Aaron Ward', 'ALLPRO OFFICE', 'CCSD', 'Mystery Co', 'Roof Co', 'Southpoint'],
);
const pbjJob = (ap: string) => pbj.customers.flatMap((c) => c.jobs).find((j) => j.analysis.apKey === ap)!;
const j24314 = pbjJob('24314');
check('AP# autofilled from jobcode_2', j24314.analysis.apKey, '24314');
check('On Calendar? = TimeNavi row {n}', j24314.analysis.onCalendar, 'TimeNavi row 3');
check('% of Completion mirrors the grid (pay-date 30%)', j24314.analysis.percent, 0.3);
check('crew lead resolved from "(Daniel)"', j24314.analysis.crewLead, 'DANIEL GONZALES BONILLA');
// Round 6 §5.3 — WO% / WO Override blank by default; description patterns are suggestions only.
check('WO% blank by default', j24314.analysis.woPct, null);
check('WO Override blank by default', j24314.analysis.woOverride, null);
check('block colored by estimator (MH yellow)', j24314.colorHex, 'FFFF99');
const j24403 = pbjJob('24403');
check(
  'WO Override description-$ is suggestion only (cell stays blank)',
  [j24403.analysis.woOverride, j24403.analysis.woOverrideSuggestion, j24403.analysis.woOverrideSuggestionSource?.includes('PER JODI')],
  [null, 4715, true],
);
check('WO% "@ 40% PS" is suggestion only (cell stays blank)', [j24403.analysis.woPct, j24403.analysis.woPctSuggestion], [null, 0.4]);
const jOffice = pbjJob('OFFICE');
check('OFFICE row: -, HOURLY lead, ND dark blue', [jOffice.analysis.onCalendar, jOffice.analysis.estimator, jOffice.analysis.crewLead, jOffice.colorHex, jOffice.analysis.hourly], ['-', '-', 'HOURLY', '4472C4', true]);
const j99999 = pbjJob('99999');
check('unmatched job: NOT ON CALENDAR, no fill', [j99999.analysis.onCalendar, j99999.colorHex], ['NOT ON CALENDAR', null]);
const j24450 = pbjJob('24450');
check('split leads joined with / and tagged', [j24450.analysis.crewLead, j24450.analysis.split], ['TED FRANK/RAFAEL AGUILAR', 2]);

// §4.6 seeds.
check(
  'BALANCES: <100%-paid NV jobs with no hours this week (notes blank — Round 7)',
  pbj.balances.map((r) => [r.apKey, r.notes]),
  [['21848-A5', ''], ['20000', '']],
);
check('BALANCES: fully-paid job NOT carried', pbj.balances.some((r) => r.apKey === '24314'), false);
check(
  'UT PAYROLL: last week UT rows + this week Utah-calendar events (deduped)',
  pbj.utPayroll.map((r) => r.apKey),
  ['24111', '24222'],
);
check('UT event seed carries resolved fields', pbj.utPayroll[1].percent, 0.9);

// Manual-state merge: delete a seed, add a row, edit a seed.
const pbj2 = buildPivotByJob({
  enriched, reportRows: rows, settings, leadCtx, knownEstimators, payDate,
  edits: { [j24314.id]: { apKey: '24450' } },
  qbHours: { [`${j24314.id}\u0000leo@x.com`]: 42.5 },
  lastWeekPbj,
  balancesState: { added: [], deleted: ['balances:20000'], edits: { 'balances:21848-A5': { notes: 'edited note' } } },
  utState: { added: [{ id: 'manual:ut:1', customer: 'X', job: 'New UT row', apKey: '24999', estimator: '', crewLead: '', woOverride: null, woPct: null, percent: null, livePercent: null, percentOverride: false, excludedFromPayroll: false, notes: '', colorHex: null }], deleted: [], edits: {} },
  pl, nvTotals, utTotals, nvSales, utSales,
});
check('editing an AP# re-runs On Calendar?/%Compl live', [pbj2.customers[0].jobs[0].analysis.onCalendar, pbj2.customers[0].jobs[0].analysis.percent], ['TimeNavi row 5', 0.5]);
check('QB Hrs override stored on the employee row', pbj2.customers[0].jobs[0].employees[0].qbHours, 42.5);
check('manual delete removes the seed', pbj2.balances.map((r) => r.apKey), ['21848-A5']);
check('manual edit applied to a seed', pbj2.balances[0].notes, 'edited note');
check('manual add appended', pbj2.utPayroll.map((r) => r.apKey), ['24111', '24222', '24999']);

console.log('payroll preview…');
const preview = buildPayrollPreview({
  pbj, enriched, reportRows: rows, pl, directory, settings, payDate,
  edits: {}, qbHours: {}, knownEstimators, leadCtx,
  nvSales, utSales, lastWeekPbj,
})!;
check('title with week range', preview.title, 'PAYROLL BREAKDOWN 7/10/2026 (6/29 - 7/3)');
check('painter blocks alphabetical by first name', preview.painters.map((b) => b.employeeName), ['Juan Perez', 'Leonardo Martinez-Rincon']);
const leoBlock = preview.painters[1];
check('OT split: 42h → 40 regular + 2 OT on the last job', leoBlock.lines.map((l) => [l.hours, l.ot]), [[40, false], [2, true]]);
check('OT line at 1.5× the rate', [leoBlock.lines[0].rate, leoBlock.lines[1].rate], [18, 27]);
check('block totals (reg+OT = actual, pay = Σ h×r)', [leoBlock.totalHours, leoBlock.totalPay], [42, 774]);
// Round 8 §9.2 — weekly and daily OT carry distinguishable tags, so the
// weekly line now reads "[OT >40 hrs/wk]".
check('weekly OT tag on the line label', leoBlock.lines[1].jobLabel.endsWith('[OT >40 hrs/wk]'), true);
const juanB = preview.painters[0];
check('unknown employee → fallback rate + flag', [juanB.lines[0].rate, juanB.lines[0].rateDefaulted], [18, true]);
check('job display name on the line', juanB.lines[0].jobLabel, 'CCSD 5321 Lambrook Dr Int/Ext [PS]');
check('lines in first-encounter order', juanB.lines.map((l) => l.apKey), ['24403', '99999', '24450', '24500']);

check(
  'lead blocks alphabetical by first name',
  preview.leads.map((b) => b.leadName),
  ['CARLOS RUIZ CRUZ', 'DANIEL GONZALES BONILLA', 'RAFAEL AGUILAR', 'RENE SALAZAR', 'TED FRANK'],
);
const leadBy = (name: string) => preview.leads.find((b) => b.leadName === name)!;
const daniel = leadBy('DANIEL GONZALES BONILLA');
check('lead jobs = piece-rate jobs he leads', daniel.jobs.map((j) => j.apKey), ['24314', '24403']);
const d24314 = daniel.jobs[0];
check('PS chain: profitShare = income × Settings crew WO% (pivot WO% blank)', d24314.profitShare, round2(650 * 0.38));
check('PS chain: WO Override blank by default', d24314.workOrderOverride, null);
check('PS chain: % Completion from Pivot by Job', d24314.percentCompletion, 0.3);
// Round 8 §8.1 (supersedes Round 6 §6.2): base × %Completion − labor − ½mat.
// Mileage stays display-only.
check(
  'PS chain: Total Profit Share = PS × %Completion − labor − ½mat',
  round2(d24314.totalProfitShare),
  round2(round2(650 * 0.38) * 0.3 - 113.04 - round2(35.74 / 2)),
);
const d24403 = daniel.jobs[1];
check('PS chain: description WO Override not auto-applied', d24403.workOrderOverride, null);
// 24403's only status line on/before the 7/10 pay date is "7/03/26 10% COMPLETE".
check('PS chain: 24403 % Completion from Pivot by Job', d24403.percentCompletion, 0.1);
check(
  'PS chain math on 24403 (Settings 38% PS × 10% complete)',
  round2(d24403.totalProfitShare),
  round2(9000 * 0.38 * 0.1 - 2081.09 - round2(760.1 / 2)),
);
check('lead Total = Σ after-split totals', round2(daniel.total), round2(d24314.totalAfterSplit + d24403.totalAfterSplit));
check('lead with no timesheet hours: no T-SHEETS block, Hourly Pay 0', [daniel.tsheets, daniel.hourlyPay], [null, 0]);
// Round 8 §8.2 — Total + Hourly Pay + Reimbursements − Deductions. Daniel's
// jobs carry no P&L sundries, so the deductions default is 0 here.
check('Total Pay = Total + Hourly + Reimb − Deductions', daniel.totalPay, daniel.total);
const ted = leadBy('TED FRANK');
const carlos = leadBy('CARLOS RUIZ CRUZ');
check('§6 carried no-timesheet job joins its lead band', carlos.jobs.map((j) => j.apKey), ['21848-A5']);
check('§6 Rene carried AP20000 is on his band', leadBy('RENE SALAZAR').jobs.some((j) => j.apKey === '20000'), true);
// Round 8 §9.4 — a lead with no WO% on file is auto-added at the seeded 38%
// default, so the split job now pays: 150 × 38% × 50% complete, ÷ 2 co-leads.
check(
  'split job: each co-lead gets ÷2 (Ted auto-added at the 38% default)',
  [ted.jobs[0].splitWith, round2(ted.jobs[0].totalAfterSplit)],
  [2, round2((150 * 0.38 * 0.5) / 2)],
);

check('office block: single Office line', [preview.office.length, preview.office[0].lines.length, preview.office[0].lines[0].jobLabel], [1, 1, 'Office']);
check('salary sections from the roster', [preview.admin.length, preview.estimators.length, preview.managers.length], [1, 5, 4]);
check(
  'inactive roster includes Calixto (directory * + persisted painters)',
  preview.inactiveZeroHours.includes('Calixto Efrain Climaco'),
  true,
);
check('inactive roster is the persisted painter list', preview.inactiveZeroHours.length >= 50, true);

// Round 4 §5.2 — one A→Z painter sequence interleaving painters and leads,
// subs last (none here). §5.3 — lead bands in Title Case. §5.5 — office PO text.
check(
  'painter sequence interleaved A→Z by first name',
  preview.painterSequence.map((e) => (e.kind === 'hourly' ? e.block.employeeName : e.block.displayName)),
  [
    'Carlos Ruiz Cruz',
    'Daniel Gonzales Bonilla',
    'Juan Perez',
    'Leonardo Martinez-Rincon',
    'Rafael Aguilar',
    'Rene Salazar',
    'Ted Frank',
  ],
);
check('week range on the model', preview.weekRange, '6/29 - 7/3');
check('office PO cell text = "(week range)"', preview.office[0].lines[0].poText, '(6/29 - 7/3)');
check('painter crew tag verbatim', preview.painters[0].crewTag, "Rene Salazar's Crew");

// Live edit: entering net computes taxes; PS-cell edits recompute the chain.
const previewEdited = buildPayrollPreview({
  pbj, enriched, reportRows: rows, pl, directory, settings, payDate,
  edits: {
    [`hb:${leoBlock.key}`]: { net: 600 },
    [`ps:${daniel.key}:${d24403.jobRowId}`]: { percentCompletion: 1 },
    [`psb:${daniel.key}`]: { reimbursements: 100 },
  },
  qbHours: {}, knownEstimators, leadCtx,
  nvSales, utSales, lastWeekPbj,
})!;
check('net edit stored (taxes/NEST/net independent; taxes not auto-derived)', previewEdited.painters[1].net, 600);
const editedDaniel = previewEdited.leads.find((b) => b.leadName === 'DANIEL GONZALES BONILLA')!;
const editedJob = editedDaniel.jobs.find((j) => j.apKey === '24403')!;
// Round 8 §8.1 — % Completion is a term now, so editing it to 100% DOES move
// Total Profit Share (up from the 10% the calendar suggested).
check('percent edit recomputes Total Profit Share', round2(editedJob.percentCompletion), 1);
check(
  'percent edit at 100% pays the full base less labor and ½mat',
  round2(editedJob.totalProfitShare),
  round2(9000 * 0.38 * 1 - 2081.09 - round2(760.1 / 2)),
);
check('reimbursement edit flows into Total Pay', round2(editedDaniel.totalPay), round2(editedDaniel.total + 100));

console.log('payroll exceptions…');
const pGroups = buildPayrollExceptions({
  pbj, preview, reportRows: rows, pl, hasDirectory: true,
  enriched, nvSales, utSales, lastWeekPbj,
});
const pg = (key: string) => pGroups.find((x) => x.key === key)!;
check('lead resolution flags (not-a-crew-lead ×4 + missing lead on 24500)', pg('leadResolution').items.length, 5);
check('hours on a job not on the calendar (99999)', pg('hoursNotOnCalendar').items.length, 1);
check('calendar job with no hours (24388, Utah excluded)', pg('calendarNoHours').items.length, 1);
check('40+ hours item (Leonardo OT)', pg('hours40or10').items.length, 1);
check('missing pay inputs (blank WO% on piece-rate jobs + default-rate employees)', pg('missingInputs').items.length, 9);
check('job missing from P&L: carried jobs with no P&L column', pg('missingFromPl').items.length, 2);
// Round 8 §8.1 flips this group: multiplying by % Completion pushes both of
// Daniel's jobs (30% and 10% complete) below the $25 threshold, while the
// §9.4 auto-added 38% lifts the split job to $28.50 and clears it.
check('low profit-share items (Daniel\u2019s two part-complete jobs)', pg('lowProfitShare').items.length, 2);
check('inactive-with-hours: none', pg('inactiveWithHours').items.length, 0);
check('duplicate P&L column surfaced', pg('duplicatePlColumn').items.length, 1);
check('UT carry-forward that matched nothing (24111)', pg('carryForward').items.length, 1);

// ---------------------------------------------------------------------------
// Export: build the styled workbook and read it back with ExcelJS
// ---------------------------------------------------------------------------

async function testExport(): Promise<void> {
  console.log('xlsx export…');
  const included = rows.filter((r) => r.included);
  const timeNaviRowByAp = new Map<string, number>();
  included.forEach((r, i) => {
    if (r.apKey) timeNaviRowByAp.set(r.apKey, i + 2); // header = row 1
  });
  const pctBuf = await buildPercentsWorkbookBuffer({
    rows,
    pivot,
    legend: DEFAULT_LEGEND,
    payDate,
    pivotByJob: pbj,
    percentExceptions: groups,
    salesByRep: buildSalesByRepMirrorFromParsed(nvSales),
    salesByCust: [],
    utSalesByRep: buildSalesByRepMirrorFromParsed(utSales),
    utSalesByCust: [],
    timesheetMirror: buildTimesheetMirror(enriched, timeNaviRowByAp),
    jobsByCrew: buildJobsByCrew(enriched),
  });
  const readBack = new ExcelJS.Workbook();
  await readBack.xlsx.load(pctBuf);

  // Round 6/7 §9 — %'s workbook: mirrors + Jobs by Crew; payroll is a separate file.
  check(
    "sheet names (Round 6/7 %'s workbook)",
    readBack.worksheets.map((w) => w.name),
    [
      'Sales by Rep',
      'Sales by Cust',
      'UT SalesbyRep',
      'UT SalesbyCust',
      'timesheets',
      'TimeNavi Export',
      'Pivot by Crew',
      'Jobs by Crew',
      'Pivot by Job',
      "%'s Exceptions",
    ],
  );

  // Round 7 — timesheets "Not on Calendar" uses `TimeNavi row {n}` when resolved.
  const tsMirror = buildTimesheetMirror(enriched, timeNaviRowByAp);
  check(
    'timesheets Not on Calendar: TimeNavi row prefix when resolved',
    tsMirror.find((r) => r.apKey === '24314')?.notOnCalendar,
    'TimeNavi row 3',
  );
  check(
    'timesheets Not on Calendar: unresolved → NotOnCalendar',
    tsMirror.find((r) => r.apKey === '99999')?.notOnCalendar,
    'NotOnCalendar',
  );

  const payBuf = await buildPayrollWorkbookBuffer({
    payDate,
    weekStart: pbj.dates[0] ?? null,
    weekEnd: pbj.dates[pbj.dates.length - 1] ?? null,
    pl,
    payrollPreview: preview,
    payrollExceptions: pGroups,
    plSourceFileName: 'Profit and Loss by Customer.xlsx',
  });
  const payWb = new ExcelJS.Workbook();
  await payWb.xlsx.load(payBuf);
  check(
    'sheet names (Round 7 payroll workbook)',
    payWb.worksheets.map((w) => w.name),
    [
      'P&L by Job Import',
      'Management Report Export',
      'PAYROLL PREVIEW',
      'Profitshare Output',
      "Input #'s",
      'NOTES',
      'FINAL PAYROLL BREAKDOWN',
      '►JOURNAL ENTRY◄',
      'JOBS',
      'Inactive Employees',
      'Payroll Exceptions',
      'PR Final Check',
    ],
  );
  // Optional Round 7 tab colors (ARGB) on a couple of named sheets.
  check(
    'tab color: PAYROLL PREVIEW',
    (payWb.getWorksheet('PAYROLL PREVIEW') as ExcelJS.Worksheet).properties.tabColor?.argb,
    'FFF79646',
  );
  check(
    'tab color: NOTES',
    (payWb.getWorksheet('NOTES') as ExcelJS.Worksheet).properties.tabColor?.argb,
    'FF000000',
  );
  {
    const notesWs = payWb.getWorksheet('NOTES') as ExcelJS.Worksheet;
    const woNotes: string[] = [];
    notesWs.eachRow((row) => {
      const v = row.getCell(6).value;
      if (typeof v === 'string' && v.startsWith('w.o. $')) woNotes.push(v);
    });
    check(
      'NOTES col F uses w.o. $X @N% PS & % complete',
      woNotes.length > 0 && woNotes.every((s) => /^w\.o\. \$[\d,.]+ @(\d+|\?)% PS & /.test(s)),
      true,
    );
  }
  check(
    'tab color: Pivot by Job',
    (readBack.getWorksheet('Pivot by Job') as ExcelJS.Worksheet).properties.tabColor?.argb,
    'FFFFFF00',
  );

  const tn = readBack.getWorksheet('TimeNavi Export') as ExcelJS.Worksheet;
  check('TimeNavi frozen at A2', [tn.views[0].state, (tn.views[0] as { ySplit?: number }).ySplit], ['frozen', 1]);
  check('TimeNavi header text', [tn.getCell('A1').value, tn.getCell('H1').value, tn.getCell('S1').value], ['Event', 'AP#', 'Date Difference']);
  // Round 2 §2: the ENTIRE header row is black — zero green cells in row 1.
  const row1Fills: (string | null)[] = [];
  tn.getRow(1).eachCell({ includeEmpty: false }, (cell) => row1Fills.push(fillOf(cell)));
  check('entire header row black A1→last header', row1Fills.every((f) => f === 'FF000000'), true);
  check('header spans all populated columns', row1Fills.length, 19);
  check('no green anywhere in row 1', row1Fills.includes('FF1BA41B'), false);
  check(
    'header white and not bold',
    [tn.getCell('A1').font?.bold === true, tn.getCell('A1').font?.color?.argb, tn.getCell('H1').font?.color?.argb],
    [false, 'FFFFFFFF', 'FFFFFFFF'],
  );
  // §1.2: E, F, G, L hidden — never deleted (values still present).
  check(
    'columns E,F,G,L hidden',
    [5, 6, 7, 12].map((c) => tn.getColumn(c).hidden === true),
    [true, true, true, true],
  );
  check('hidden column data intact (Description)', String(tn.getCell('G3').value).includes('30% COMPLETE'), true);
  check('other columns not hidden', [1, 8, 14].map((c) => tn.getColumn(c).hidden === true), [false, false, false]);
  // C-3: Event autofits (capped), Description ~40% narrower, AP# fits longest key.
  check(
    'column widths (A, B, G, H, L)',
    [1, 2, 7, 8, 12].map((c) => tn.getColumn(c).width),
    [36, 11, 7.8, 6, 39],
  );

  // Only the 6 included rows are exported; row 2 = JS (24222) after sorting.
  check('only included rows exported', tn.actualRowCount, 7);
  check('row 2 is the JS job', tn.getCell('H2').value, '24222');
  check('JS row peach fill', fillOf(tn.getCell('A2')), 'FFFABF8F');
  // §1.4: fallback-resolved TE row carries its color; RF row gets the new color.
  check('fallback TE row: J value + green fill', [tn.getCell('J5').value, fillOf(tn.getCell('A5'))], ['TE', 'FFC4D79B']);
  // Round 8 §5 — the RF seed moved from DA9694 to C4A484 (migration r8-rf-color).
  check('RF row tan fill', [tn.getCell('J7').value, fillOf(tn.getCell('A7'))], ['RF', 'FFC4A484']);
  check('light fill keeps black text', tn.getCell('A7').font?.color?.argb, 'FF000000');

  // B-4 / C-1: date cells are UTC-midnight (no time), formatted mm-dd-yy.
  const b2 = tn.getCell('B2').value as Date;
  check(
    'Start Date UTC-midnight, no time',
    [b2.getUTCFullYear(), b2.getUTCMonth() + 1, b2.getUTCDate(), b2.getUTCHours(), b2.getUTCMinutes()],
    [2026, 6, 30, 0, 0],
  );
  check('Start/End Date format mm-dd-yy', [tn.getCell('B2').numFmt, tn.getCell('C2').numFmt], ['mm-dd-yy', 'mm-dd-yy']);
  check('Invoice Date format mm/dd/yyyy', tn.getCell('R2').numFmt, 'mm/dd/yyyy');
  check('Date Difference integer format', [tn.getCell('S2').value, tn.getCell('S2').numFmt], [-7, '0']);
  check('percent stored as fraction + 0% format', [tn.getCell('I2').value, tn.getCell('I2').numFmt], [0.9, '0%']);
  check('Last % Paid HOURLY text survives', tn.getCell('K2').value, 'HOURLY');
  check(
    'money accounting format, no $',
    [tn.getCell('N2').value, String(tn.getCell('N2').numFmt).replace(/\\/g, '')],
    [14410.63, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'],
  );
  check('%compl mirrors % Completion', tn.getCell('Q2').value, 0.9);
  // Row 6 = XC (24388): #N/A / blank / blank, '-' for last paid.
  check('unmatched Job Total exports literal #N/A', tn.getCell('N6').value, '#N/A');
  check('unmatched Apptivo blank', isBlank(tn.getCell('O6').value), true);
  check('unmatched $ Difference blank', isBlank(tn.getCell('P6').value), true);
  check('dash for missing last % paid (24388 row)', tn.getCell('K6').value, '-');

  // Round 4 §1.2 — H/I/J/K and P centered (header + data).
  check(
    'TimeNavi centered columns (header)',
    ['H1', 'I1', 'J1', 'K1', 'P1'].map((a) => tn.getCell(a).alignment?.horizontal),
    ['center', 'center', 'center', 'center', 'center'],
  );
  check('TimeNavi centered data cell', tn.getCell('H2').alignment?.horizontal, 'center');
  // Round 4 §1.1 — pink flags: NotClockedIn in M (row 2 = 24222) pink; a zero
  // $ Difference stays unflagged (row 2 P = 0).
  check(
    'NotClockedIn cell pink fill + dark-red text',
    [fillOf(tn.getCell('M2')), tn.getCell('M2').font?.color?.argb],
    ['FFFFC7CE', 'FF9C0006'],
  );
  // Round 8 §2.2 — with Apptivo blank, P exports blank rather than 0; either
  // way it must keep the row fill and never take the pink flag.
  check('blank $ Difference NOT pink', [isBlank(tn.getCell('P2').value), fillOf(tn.getCell('P2'))], [true, 'FFFABF8F']);

  // ---------- Pivot by Crew ----------
  const pv = readBack.getWorksheet('Pivot by Crew') as ExcelJS.Worksheet;
  const gtColNum = 4 + pivot.estimators.length + 1; // computed, never assumed
  check('pivot frozen at A5', [pv.views[0].state, (pv.views[0] as { ySplit?: number }).ySplit], ['frozen', 4]);

  // §2.1: rows 1–3 black full-width, white text; title 17pt bold, height 31.5.
  check('row 1 title', pv.getCell('A1').value, 'NV PAYROLL: 7/10/2026');
  check('row 1 height', pv.getRow(1).height, 31.5);
  check(
    'title cell black/white 17 bold',
    [fillOf(pv.getCell('A1')), pv.getCell('A1').font?.color?.argb, pv.getCell('A1').font?.size, pv.getCell('A1').font?.bold],
    ['FF000000', 'FFFFFFFF', 17, true],
  );
  check('black band spans full width (rows 1–3)', [fillOf(pv.getCell(1, gtColNum)), fillOf(pv.getCell(2, gtColNum)), fillOf(pv.getCell(3, gtColNum))], ['FF000000', 'FF000000', 'FF000000']);
  check('row 2 Name/Job labels', [pv.getCell('C2').value, pv.getCell('D2').value], ['Name', 'Job']);
  const hoursCol = 5 + Math.floor((pivot.estimators.length - 1) / 2);
  check('row 2 Hours centered over estimator block (no merge)', [pv.getCell(2, hoursCol).value, pv.getCell(2, hoursCol).isMerged], ['Hours', false]);
  check('row 3 chrome', [pv.getCell('A3').value, pv.getCell('E3').value], ['Sum of hours', 'Column Labels']);

  // §2.2: row 4 gray labels + palette-colored estimator headers, not bold.
  check('row 4 field header', [pv.getCell('A4').value, pv.getCell('D4').value, pv.getCell(4, gtColNum).value], ['Row Labels', 'jobcode_2', 'Grand Total']);
  check('row 4 label cells gray, black text', [fillOf(pv.getCell('A4')), pv.getCell('A4').font?.color?.argb], ['FF808080', 'FF000000']);
  check('row 4 not bold', pv.getCell('A4').font?.bold === true, false);
  check('row 4 height 15', pv.getRow(4).height, 15);
  check('MH header in palette color', [pv.getCell('E4').value, fillOf(pv.getCell('E4'))], ['MH', 'FFFFFF99']);
  check('zero-hour XC still gets its header + color', [pv.getCell('G4').value, fillOf(pv.getCell('G4'))], ['XC', 'FF95B3D7']);
  check('ND header dark blue with white text', [pv.getCell('H4').value, fillOf(pv.getCell('H4')), pv.getCell('H4').font?.color?.argb], ['ND', 'FF4472C4', 'FFFFFFFF']);
  // Round 2 §4: RF right after TE (column J with the current layout).
  check('RF header tan right after TE', [pv.getCell('I4').value, pv.getCell('J4').value, fillOf(pv.getCell('J4'))], ['TE', 'RF', 'FFC4A484']);
  check('#N/A header gray, last before Grand Total', [pv.getCell('K4').value, fillOf(pv.getCell('K4'))], ['#N/A', 'FF808080']);
  check('Grand Total header gray', fillOf(pv.getCell(4, gtColNum)), 'FF808080');

  // §2.4: crew rows black full width, white text (row 5 = first crew, A→Z).
  check('first crew row (A→Z)', pv.getCell('A5').value, 'Office (NV)');
  check('crew row black band full width', [fillOf(pv.getCell('A5')), fillOf(pv.getCell(5, gtColNum)), pv.getCell('A5').font?.color?.argb], ['FF000000', 'FF000000', 'FFFFFFFF']);
  check('crew row not bold', pv.getCell('A5').font?.bold === true, false);

  // §2.3: full-height column bands incl. empty cells; gray Grand Total band;
  // no band under #N/A. Row 7 = OFFICE job row (values), row 6 = employee row (empty cells).
  check('MH band on an empty employee-row cell', fillOf(pv.getCell('E6')), 'FFFFFF99');
  check('zero-hour XC column still banded', [isBlank(pv.getCell('G7').value), fillOf(pv.getCell('G7'))], [true, 'FF95B3D7']);
  check('ND band + white value text on the OFFICE job row', [pv.getCell('H7').value, fillOf(pv.getCell('H7')), pv.getCell('H7').font?.color?.argb], [6, 'FF4472C4', 'FFFFFFFF']);
  check('#N/A column has no band', fillOf(pv.getCell('K6')), null);
  check('Grand Total column light-gray band', [fillOf(pv.getCell(6, gtColNum)), fillOf(pv.getCell(7, gtColNum))], ['FFD9D9D9', 'FFD9D9D9']);

  // Round 2 §1: bold row labels. Row 6 = employee header (Nancy), row 7 = her
  // OFFICE job row.
  check('employee row bold A:C', [1, 2, 3].map((c) => pv.getCell(6, c).font?.bold === true), [true, true, true]);
  check('job row: customer (C) bold, job code (D) regular', [pv.getCell(7, 3).font?.bold === true, pv.getCell(7, 4).font?.bold === true], [true, false]);
  check('job-row hour values regular weight', pv.getCell('H7').font?.bold === true, false);

  // Locate structural rows dynamically.
  let grandRowNum = 0;
  let leoTotalRowNum = 0;
  pv.eachRow((row, n) => {
    if (row.getCell(1).value === 'Grand Total') grandRowNum = n;
    if (row.getCell(2).value === 'Leonardo Total') leoTotalRowNum = n;
  });

  // §2.7: the >40h yellow is gone from the export — the Grand Total column
  // stays its gray band even on Leonardo's 42h total row.
  check('no overtime yellow leak in Grand Total column', [pv.getCell(leoTotalRowNum, gtColNum).value, fillOf(pv.getCell(leoTotalRowNum, gtColNum))], [42, 'FFD9D9D9']);
  // Round 2 §1: "X Total" rows are fully bold and close the block with a thin
  // bottom border.
  check(
    'total row fully bold (label + hour values)',
    [pv.getCell(leoTotalRowNum, 2).font?.bold, pv.getCell(leoTotalRowNum, 5).font?.bold, pv.getCell(leoTotalRowNum, gtColNum).font?.bold],
    [true, true, true],
  );
  // Round 4 §2.4: the separator sits on the TOP edge of the total row,
  // spanning C → Grand Total (not A/B, not the bottom edge).
  check(
    'total row thin TOP border from C through Grand Total',
    [
      pv.getCell(leoTotalRowNum, 3).border?.top?.style,
      pv.getCell(leoTotalRowNum, gtColNum).border?.top?.style,
      pv.getCell(leoTotalRowNum, 1).border?.top?.style ?? null,
      pv.getCell(leoTotalRowNum, 3).border?.bottom?.style ?? null,
    ],
    ['thin', 'thin', null, null],
  );
  // Round 4 §2.4: hour values centered in the estimator columns.
  check('hour value centered (estimator column)', pv.getCell('H7').alignment?.horizontal, 'center');

  // §2.5 + Round 2 §1: Grand Total row bold, gray/white on A + value columns,
  // B:D unfilled.
  check('grand row A gray white bold', [fillOf(pv.getCell(grandRowNum, 1)), pv.getCell(grandRowNum, 1).font?.color?.argb, pv.getCell(grandRowNum, 1).font?.bold], ['FF808080', 'FFFFFFFF', true]);
  check('grand row B:D unfilled', [2, 3, 4].map((c) => fillOf(pv.getCell(grandRowNum, c))), [null, null, null]);
  check('grand row value cells gray white bold + 0.00', [pv.getCell(grandRowNum, gtColNum).value, fillOf(pv.getCell(grandRowNum, gtColNum)), pv.getCell(grandRowNum, gtColNum).font?.bold, pv.getCell(grandRowNum, gtColNum).numFmt], [62, 'FF808080', true, '0.00']);

  // §2.7: Calibri 12 body, hours 0.00, widths.
  check('body font Calibri 12', [pv.getCell('C7').font?.name, pv.getCell('C7').font?.size], ['Calibri', 12]);
  check('hours cells 0.00', pv.getCell('H7').numFmt, '0.00');
  check('pivot widths (A, D, first est col, GT)', [pv.getColumn(1).width, pv.getColumn(4).width, pv.getColumn(5).width, pv.getColumn(gtColNum).width], [35, 39, 8, 8]);

  // ---------- Pivot by Job (Round 6: TimeNavi row label; WO blank by default) ----------
  const pj = readBack.getWorksheet('Pivot by Job') as ExcelJS.Worksheet;
  check('pbj frozen at A3', [pj.views[0].state, (pj.views[0] as { ySplit?: number }).ySplit], ['frozen', 2]);
  check('pbj A1 title', pj.getCell('A1').value, 'NV PAYROLL: 7/10/2026');
  check(
    'pbj title black/white 17 bold, height 31.5',
    [fillOf(pj.getCell('A1')), pj.getCell('A1').font?.color?.argb, pj.getCell('A1').font?.size, pj.getRow(1).height],
    ['FF000000', 'FFFFFFFF', 17, 31.5],
  );
  // 5 date cols → Total hrs at G1; % of Completion header red, Notes dark red.
  check('pbj Total hrs header position (computed from 5 dates)', pj.getCell('G1').value, 'Total hrs');
  const lastColNum = 1 + 5 + 2 + 8;
  check('pbj % of Completion header red', pj.getCell(1, lastColNum - 1).font?.color?.argb, 'FFFF0000');
  check('pbj Notes header dark red', pj.getCell(1, lastColNum).font?.color?.argb, 'FFC00000');
  check('pbj % of Completion column width', pj.getColumn(lastColNum - 1).width, 15.78);
  // Row 2: computed weekday names (Sunday would be present on an 8-date week).
  check('pbj computed day names', [pj.getCell('B2').value, pj.getCell('F2').value], ['Mon', 'Fri']);
  // Round 5 §3.1 — data begins at row 3; no Sum of hours / Row Labels rows.
  check('pbj first customer row at row 3 + MH fill', [pj.getCell('A3').value, fillOf(pj.getCell('A3'))], ['Aaron Ward', 'FFFFFF99']);
  check('pbj no Sum of hours / Column Labels artifact', typeof pj.getCell('A3').value === 'string' && !String(pj.getCell('A3').value).includes('Sum of hours'), true);
  check('pbj job row analysis: AP#, TimeNavi row, lead', [pj.getCell('I4').value, pj.getCell('J4').value, pj.getCell('L4').value], ['24314', 'TimeNavi row 3', 'DANIEL GONZALES BONILLA']);
  check('pbj WO Override blank by default', isBlank(pj.getCell('M4').value), true);
  const pctComplCell = pj.getCell('O4');
  const pctComplVal = pctComplCell.value as { formula?: string } | number;
  check(
    'pbj job % of Completion is INDEX/MATCH vs TimeNavi',
    typeof pctComplVal === 'object' &&
      typeof pctComplVal.formula === 'string' &&
      pctComplVal.formula.includes('INDEX') &&
      pctComplVal.formula.includes('MATCH') &&
      !pctComplVal.formula.includes('XLOOKUP'),
    true,
  );
  check('pbj job % of Completion 0% format', pctComplCell.numFmt, '0%');
  // Round 7 — Notes column keeps the header label; job/BALANCES/UT data cells stay empty.
  check('pbj Notes header present', pj.getCell(1, lastColNum).value, 'Notes');
  check('pbj Notes data empty on job row (Round 7)', isBlank(pj.getCell(4, lastColNum).value), true);
  check('pbj hours 0.00 format', pj.getCell('G4').numFmt, '0.00');
  check('pbj widths (A, first date, Crew Lead)', [pj.getColumn(1).width, pj.getColumn(2).width, pj.getColumn(12).width], [40, 10, 43]);
  // Outline levels: customer 0, job 1, employee 2; summaryAbove.
  check(
    'pbj outline levels (customer/job/employee)',
    [pj.getRow(3).outlineLevel, pj.getRow(4).outlineLevel, pj.getRow(5).outlineLevel],
    [0, 1, 2],
  );
  check('pbj summaryBelow false (summary above)', (pj.properties.outlineProperties as { summaryBelow?: boolean } | undefined)?.summaryBelow, false);
  // Round 5 §3.5 — thin top+bottom on A through QB Hrs; not on analysis.
  check(
    'pbj customer/job hours-band borders',
    [
      pj.getCell('A3').border?.top?.style,
      pj.getCell('H3').border?.bottom?.style,
      pj.getCell('I3').border?.top?.style ?? null,
    ],
    ['thin', 'thin', null],
  );
  check(
    'pbj analysis cells centered, Notes not',
    [
      pj.getCell('I4').alignment?.horizontal,
      pj.getCell('M4').alignment?.horizontal,
      pj.getCell('O4').alignment?.horizontal,
      pj.getCell('P4').alignment?.horizontal ?? null,
    ],
    ['center', 'center', 'center', null],
  );
  // Round 4 §3.2-1 — WO Override column keeps $-accounting format (even when blank).
  let woFmtOk = false;
  pj.eachRow((row) => {
    const cell = row.getCell(13);
    if (String(cell.numFmt ?? '').includes('"$"')) woFmtOk = true;
  });
  check('pbj WO Override uses $-accounting somewhere', woFmtOk, true);
  // Round 4 §3.2-4 — Grand Total row double top border.
  let pjGrandRow = 0;
  pj.eachRow((row, n) => {
    if (row.getCell(1).value === 'Grand Total' && pjGrandRow === 0) pjGrandRow = n;
  });
  check('pbj Grand Total double top border', pj.getCell(pjGrandRow, 1).border?.top?.style, 'double');
  // BALANCES + UT sections rendered beneath at computed positions.
  const colA = pj.getColumn(1).values as (string | undefined)[];
  const balTitleRow = colA.findIndex((v) => typeof v === 'string' && v.startsWith('NV PAYROLL (BALANCES):'));
  const utTitleRow = colA.findIndex((v) => typeof v === 'string' && v.startsWith('UT PAYROLL:'));
  check('BALANCES section below the pivot', balTitleRow > 5, true);
  check('UT section below BALANCES', utTitleRow > balTitleRow, true);
  check('BALANCES title black band', fillOf(pj.getCell(balTitleRow, 1)), 'FF000000');
  // Round 5 §3.2 — single black header row; customer+job pairs follow immediately.
  const balCustRow = balTitleRow + 1;
  check(
    'BALANCES pair: estimator fill on both rows (MH yellow, 21848-A5)',
    [fillOf(pj.getCell(balCustRow, 1)), fillOf(pj.getCell(balCustRow + 1, 1))],
    ['FFFFFF99', 'FFFFFF99'],
  );
  check(
    'BALANCES pair: thin box border (top on customer, bottom on job, edges)',
    [
      pj.getCell(balCustRow, 1).border?.top?.style,
      pj.getCell(balCustRow, 1).border?.left?.style,
      pj.getCell(balCustRow + 1, 1).border?.bottom?.style,
      pj.getCell(balCustRow + 1, lastColNum).border?.right?.style,
    ],
    ['thin', 'thin', 'thin', 'thin'],
  );
  check(
    'BALANCES outline levels (customer 0 / job 1)',
    [pj.getRow(balCustRow).outlineLevel, pj.getRow(balCustRow + 1).outlineLevel],
    [0, 1],
  );

  // ---------- timesheets mirror (Round 7 Not on Calendar label) ----------
  const tsSheet = readBack.getWorksheet('timesheets') as ExcelJS.Worksheet;
  const nocCol = TS_HEADERS.length + 5; // Apname…Not on Calendar (Round 6/7 helpers)
  check('timesheets Not on Calendar header', tsSheet.getCell(1, nocCol).value, 'Not on Calendar');
  let noc24314 = '';
  tsSheet.eachRow((row, n) => {
    if (n === 1) return;
    if (String(row.getCell(26).value ?? '') === '24314') noc24314 = String(row.getCell(nocCol).value ?? '');
  });
  check('timesheets sheet: 24314 Not on Calendar = TimeNavi row 3', noc24314, 'TimeNavi row 3');

  // ---------- Payroll Preview (Round 7 — separate payroll workbook) ----------
  const pp = payWb.getWorksheet('PAYROLL PREVIEW') as ExcelJS.Worksheet;
  check('preview title', pp.getCell('A1').value, 'PAYROLL BREAKDOWN 7/10/2026 (6/29 - 7/3)');
  check('preview title black band', fillOf(pp.getCell('A1')), 'FF000000');
  // Round 4 §5.1/§5.4 — Garamond everywhere; title merged A1:H1, centered, 13 bold.
  check(
    'preview title Garamond 13 bold, merged + centered',
    [pp.getCell('A1').font?.name, pp.getCell('A1').font?.size, pp.getCell('A1').font?.bold, pp.getCell('A1').isMerged, pp.getCell('A1').alignment?.horizontal],
    ['Garamond', 13, true, true, 'center'],
  );
  // Live formulas: find the Daniel PS block (Round 4 §5.3 Title Case band).
  const ppColA = pp.getColumn(1).values as (string | undefined)[];
  const danielRow = ppColA.findIndex((v) => v === 'Daniel Gonzales Bonilla');
  check('Daniel block present with the DDD9C4 tan band', [danielRow > 0, fillOf(pp.getCell(danielRow, 1))], [true, 'FFDDD9C4']);
  // Round 5 §5 — crew-lead band fills A through H regardless of job count.
  check(
    'Daniel crew-lead band fills A–H',
    [1, 2, 3, 4, 5, 6, 7, 8].map((c) => fillOf(pp.getCell(danielRow, c))),
    Array(8).fill('FFDDD9C4'),
  );
  // Round 4 §5.5 — labeled `PO# | Job | hrs | rate | total` header rows.
  let headerRowNum = 0;
  pp.eachRow((row, n) => {
    if (headerRowNum === 0 && row.getCell(1).value === 'PO#' && row.getCell(2).value === 'Job') headerRowNum = n;
  });
  check('PO#|Job|hrs|rate|total header row present', headerRowNum > 0, true);
  check(
    'header labels bold + underlined + centered',
    [
      pp.getCell(headerRowNum, 1).font?.bold,
      pp.getCell(headerRowNum, 1).font?.underline,
      pp.getCell(headerRowNum, 3).alignment?.horizontal,
      pp.getCell(headerRowNum, 5).value,
    ],
    [true, true, 'center', 'total'],
  );
  // Round 6 §6.1 — `(-) taxes` / `NEST [IRA]` / `net` bold labels on a job line.
  let taxesLabelOk = false;
  pp.eachRow((row) => {
    if (row.getCell(6).value === '(-) taxes' && row.getCell(7).value === 'NEST [IRA]' && row.getCell(8).value === 'net') {
      taxesLabelOk = true;
    }
  });
  check('(-) taxes / NEST / net labels present', taxesLabelOk, true);
  // taxes / NEST / net value cells blank by default (no E−G formula).
  let taxesBlankOk = false;
  pp.eachRow((row) => {
    if (row.getCell(6).value === '(-) taxes') {
      const sums = pp.getRow(row.number + 1);
      taxesBlankOk =
        isBlank(sums.getCell(6).value) && isBlank(sums.getCell(7).value) && isBlank(sums.getCell(8).value);
    }
  });
  check('taxes / NEST / net blank by default', taxesBlankOk, true);
  // §5.4 — a painter name band is merged A–H with the DDD9C4 fill (rich text
  // name + crew tag). Find Juan's band via the merged cell's rich text.
  let bandOk = false;
  pp.eachRow((row) => {
    const v = row.getCell(1).value as { richText?: { text: string }[] } | string | null;
    if (v && typeof v === 'object' && 'richText' in v && v.richText?.[0]?.text === 'Juan Perez') {
      bandOk = row.getCell(1).isMerged && fillOf(row.getCell(1)) === 'FFDDD9C4' && v.richText[1]?.text === " (Rene Salazar's Crew)";
    }
  });
  check('painter name band merged, tan, name + crew tag rich text', bandOk, true);
  // §5.6 — Hourly Pay value cell highlighted yellow.
  const hourlyPayRowNum = ppColA.findIndex((v) => v === 'Hourly Pay');
  check('Hourly Pay value cell yellow', fillOf(pp.getCell(hourlyPayRowNum, 2)), 'FFFFFF99');
  // §5.6 — double border under Deductions/Sundries.
  const deductRowNum = ppColA.findIndex((v) => v === 'Deductions/Sundries');
  check('double border under Deductions/Sundries', pp.getCell(deductRowNum, 2).border?.bottom?.style, 'double');
  const psFormulaCell = pp.getCell(danielRow + 3, 2); // Profit Share row
  check('Profit Share is a live formula (income × WO%)', typeof (psFormulaCell.value as { formula?: string })?.formula === 'string', true);
  const tpsRowNum = ppColA.findIndex((v) => v === 'Total Profit Share');
  const tpsFormula = String((pp.getCell(tpsRowNum, 2).value as { formula?: string })?.formula ?? '');
  // Round 8 §8.1 — both branches of the base multiply by the % Completion
  // cell, so the live formula now carries a '*' that Round 6's never did.
  check(
    'TPS formula = IF(N(WO)>0,WO×%,PS×%)−½mat−labor',
    [tpsFormula.includes('IF('), tpsFormula.includes('*'), tpsFormula.includes('-')],
    [true, true, true],
  );
  const totalPayRowNum = ppColA.findIndex((v) => v === 'Total Pay');
  const totalPayFormula = String((pp.getCell(totalPayRowNum, 2).value as { formula?: string })?.formula ?? '');
  check(
    'Total Pay = Total + Hourly + Reimb − Deductions',
    [totalPayFormula.includes('+'), totalPayFormula.includes('-'), totalPayFormula.startsWith('SUM')],
    [true, true, false],
  );
  // An hourly line total =C×D.
  let lineFormulaOk = false;
  pp.eachRow((row) => {
    const v = row.getCell(5).value as { formula?: string } | null;
    if (v && typeof v === 'object' && typeof v.formula === 'string' && /^C\d+\*D\d+$/.test(v.formula)) lineFormulaOk = true;
  });
  check('hourly line total = C×D live formula', lineFormulaOk, true);
  // Section headers present in order.
  const sectionRows = ['OFFICE', 'ADMIN', 'ESTIMATORS', 'MANAGER SALARY'].map((s) => ppColA.findIndex((v) => v === s));
  check('section headers in order', sectionRows.every((r, i) => r > 0 && (i === 0 || r > sectionRows[i - 1])), true);
  check('preview widths (A, B)', [pp.getColumn(1).width, pp.getColumn(2).width], [22, 41.4]);

  // ---------- Exceptions sheets (Round 6 — split across workbooks) ----------
  const pctEx = readBack.getWorksheet("%'s Exceptions") as ExcelJS.Worksheet;
  const payEx = payWb.getWorksheet('Payroll Exceptions') as ExcelJS.Worksheet;
  check("%'s Exceptions title band", fillOf(pctEx.getCell('A1')), 'FF000000');
  check('Payroll Exceptions title band', fillOf(payEx.getCell('A1')), 'FF000000');
  check("%'s Exceptions frozen at row 1", [pctEx.views[0].state, (pctEx.views[0] as { ySplit?: number }).ySplit], ['frozen', 1]);
  const pctItemCount = groups.filter((g) => g.items.length > 0).reduce((n, g) => n + g.items.length, 0);
  check(
    "%'s Exceptions title carries item count",
    String(pctEx.getCell('A1').value).includes(`(${pctItemCount} items)`),
    true,
  );
  const payItemCount = pGroups.filter((g) => g.items.length > 0).reduce((n, g) => n + g.items.length, 0);
  check(
    'Payroll Exceptions title carries item count',
    String(payEx.getCell('A1').value).includes(`(${payItemCount} items)`),
    true,
  );

  // ---------- Journal Entry (Round 7 §18 — header shell only) ----------
  const je = payWb.getWorksheet('►JOURNAL ENTRY◄') as ExcelJS.Worksheet;
  check(
    'Journal Entry headers (PO# | blank | Job)',
    [je.getCell('A1').value, je.getCell('B1').value, je.getCell('C1').value],
    ['PO#', '', 'Job'],
  );
  check('Journal Entry header only (no data rows)', je.actualRowCount, 1);

  // Round 8 §10 — Input #'s Hourly Pay row + live Profitshare refs.
  const inputNums = payWb.getWorksheet("Input #'s") as ExcelJS.Worksheet;
  let inputHourly = false;
  inputNums.eachRow((row) => {
    if (row.getCell(2).value === 'Hourly Pay') inputHourly = true;
  });
  check("Input #'s has an Hourly Pay row", inputHourly, true);
  const pso = payWb.getWorksheet('Profitshare Output') as ExcelJS.Worksheet;
  let psoHourlyFormula = '';
  let psoDeductFormula = '';
  pso.eachRow((row) => {
    const lab = String(row.getCell(1).value ?? '');
    const f = String((row.getCell(2).value as { formula?: string } | null)?.formula ?? '');
    if (lab === 'Hourly Pay') psoHourlyFormula = f;
    if (lab === 'Deductions/Sundries') psoDeductFormula = f;
  });
  check("Profitshare Hourly Pay refs Input #'s", psoHourlyFormula.includes("Input #''s"), true);
  check("Profitshare Deductions refs Input #'s", psoDeductFormula.includes("Input #''s"), true);
  check('NV-only week omits UT preview sheets', payWb.getWorksheet('Payroll Preview - UT'), undefined);

  // ---------- PR Final Check (Round 7 §2) ----------
  const prfc = payWb.getWorksheet('PR Final Check') as ExcelJS.Worksheet;
  check(
    'PR Final Check headers',
    [prfc.getCell('A1').value, prfc.getCell('B1').value, prfc.getCell('C1').value],
    ['PO#', 'Job', 'Total Income'],
  );
  check('PR Final Check has data rows from P&L', prfc.actualRowCount > 1, true);
}

testExport()
  .catch((err: unknown) => {
    failures++;
    console.error('  FAIL  export threw:', err);
  })
  .then(() => {
    console.log(failures === 0 ? '\nINTEGRATION: ALL PASSED' : `\nINTEGRATION: ${failures} FAILURES`);
    (globalThis as { process?: { exitCode?: number } }).process!.exitCode = failures === 0 ? 0 : 1;
  });
