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
import { buildWorkbookBuffer } from '../src/lib/exportXlsx';
import { DEFAULT_LEGEND, DEFAULT_PIVOT_COLUMNS } from '../src/lib/settings';
import { buildExceptions } from '../src/lib/exceptions';
import { enrichTimesheet } from '../src/lib/joins';
import {
  parseCustomerSummary,
  parseLastWeek,
  parseSalesDetail,
  parseTimeNavi,
  parseTimesheet,
  readWorkbook,
} from '../src/lib/parsers';
import { buildPivot } from '../src/lib/pivot';
import { buildReportRows, sortReportRows } from '../src/lib/timenavi';
import type { CellValue } from '../src/lib/types';

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
    ['AP24403 Lambrook (PG)', '8/1/2026 7:00', '8/2/2026 16:00', 'Residential (NV)', '', '', 'gate code 4321'],
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
check('24314 apptivo defaults to job total', r24314.apptivo, 650);
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

// D-2/§2.6: pivot columns come from the persisted setting, #N/A last.
const pivot = buildPivot(enriched, DEFAULT_PIVOT_COLUMNS, payDate);
check('pivot title', pivot.title, 'NV PAYROLL: 7/10/2026');
check('pivot estimator columns (setting order, #N/A last)', pivot.estimators, ['MH', 'PG', 'ND', 'TE', 'RF', '#N/A']);
check('pivot grand total', pivot.grandTotal, 62);
check('pivot MH hours', pivot.grandByEstimator['MH'], 42);
check('TE hours moved out of #N/A by the fallback', pivot.grandByEstimator['TE'], 4);
check('#N/A only holds truly unresolved hours', pivot.grandByEstimator['#N/A'], 2);
check('pivot crews A→Z', pivot.crews.map((c) => c.crew), ['Office (NV)', "Rene Salazar's Crew"]);
const rene = pivot.crews[1];
check('employees A→Z by username', rene.employees.map((e) => e.username), ['juan@x.com', 'leo@x.com']);
const juan = rene.employees[0];
const leo = rene.employees[1];
check('leo total', leo.total, 42);
check('leo overtime flag (>40)', leo.overtime, true);
check('juan not overtime', juan.overtime, false);
check('juan jobs sorted by (customer, jobcode_2)', juan.jobs.map((j) => j.customer), ['CCSD', 'Mystery Co', 'Roof Co', 'Southpoint']);
// A custom pivot order must reorder columns (proves it is setting-driven).
const pivotCustom = buildPivot(enriched, ['ND', 'PG', 'MH'], payDate);
check('custom pivot order respected, extras appended', pivotCustom.estimators, ['ND', 'PG', 'MH', 'TE', 'RF', '#N/A']);

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
// Export: build the styled workbook and read it back with ExcelJS
// ---------------------------------------------------------------------------

async function testExport(): Promise<void> {
  console.log('xlsx export…');
  const buf = await buildWorkbookBuffer({ rows, pivot, legend: DEFAULT_LEGEND, payDate });
  const readBack = new ExcelJS.Workbook();
  await readBack.xlsx.load(buf);

  check('sheet names', readBack.worksheets.map((w) => w.name), ['TimeNavi Export', 'Pivot by Crew']);

  const tn = readBack.getWorksheet('TimeNavi Export') as ExcelJS.Worksheet;
  check('TimeNavi frozen at A2', [tn.views[0].state, (tn.views[0] as { ySplit?: number }).ySplit], ['frozen', 1]);
  check('TimeNavi header text', [tn.getCell('A1').value, tn.getCell('H1').value, tn.getCell('S1').value], ['Event', 'AP#', 'Date Difference']);
  // §1.1: raw columns A:G stay green, appended columns H→end are black.
  check('raw header block green', [fillOf(tn.getCell('A1')), fillOf(tn.getCell('G1'))], ['FF1BA41B', 'FF1BA41B']);
  check('appended header block black', [fillOf(tn.getCell('H1')), fillOf(tn.getCell('S1'))], ['FF000000', 'FF000000']);
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
  // C-3: reference column widths.
  check(
    'column widths (A, B, G, H, L)',
    [1, 2, 7, 8, 12].map((c) => tn.getColumn(c).width),
    [60, 11, 13, 6.5, 39],
  );

  // Only the 6 included rows are exported; row 2 = JS (24222) after sorting.
  check('only included rows exported', tn.actualRowCount, 7);
  check('row 2 is the JS job', tn.getCell('H2').value, '24222');
  check('JS row peach fill', fillOf(tn.getCell('A2')), 'FFFABF8F');
  // §1.4: fallback-resolved TE row carries its color; RF row gets the new color.
  check('fallback TE row: J value + green fill', [tn.getCell('J5').value, fillOf(tn.getCell('A5'))], ['TE', 'FFC4D79B']);
  check('RF row dusty-rose fill', [tn.getCell('J7').value, fillOf(tn.getCell('A7'))], ['RF', 'FFDA9694']);
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
  check('ND header dark blue with white text', [pv.getCell('G4').value, fillOf(pv.getCell('G4')), pv.getCell('G4').font?.color?.argb], ['ND', 'FF4472C4', 'FFFFFFFF']);
  check('RF header dusty rose', [pv.getCell('I4').value, fillOf(pv.getCell('I4'))], ['RF', 'FFDA9694']);
  check('#N/A header gray, last before Grand Total', [pv.getCell('J4').value, fillOf(pv.getCell('J4'))], ['#N/A', 'FF808080']);
  check('Grand Total header gray', fillOf(pv.getCell(4, gtColNum)), 'FF808080');

  // §2.4: crew rows black full width, white text (row 5 = first crew, A→Z).
  check('first crew row (A→Z)', pv.getCell('A5').value, 'Office (NV)');
  check('crew row black band full width', [fillOf(pv.getCell('A5')), fillOf(pv.getCell(5, gtColNum)), pv.getCell('A5').font?.color?.argb], ['FF000000', 'FF000000', 'FFFFFFFF']);
  check('crew row not bold', pv.getCell('A5').font?.bold === true, false);

  // §2.3: full-height column bands incl. empty cells; gray Grand Total band;
  // no band under #N/A. Row 7 = OFFICE job row (values), row 6 = employee row (empty cells).
  check('MH band on an empty employee-row cell', fillOf(pv.getCell('E6')), 'FFFFFF99');
  check('ND band + white value text on the OFFICE job row', [pv.getCell('G7').value, fillOf(pv.getCell('G7')), pv.getCell('G7').font?.color?.argb], [6, 'FF4472C4', 'FFFFFFFF']);
  check('#N/A column has no band', fillOf(pv.getCell('J6')), null);
  check('Grand Total column light-gray band', [fillOf(pv.getCell(6, gtColNum)), fillOf(pv.getCell(7, gtColNum))], ['FFD9D9D9', 'FFD9D9D9']);

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
  check('total rows not bold', pv.getCell(leoTotalRowNum, 2).font?.bold === true, false);

  // §2.5: Grand Total row gray/white on A + value columns, B:D unfilled.
  check('grand row A gray white', [fillOf(pv.getCell(grandRowNum, 1)), pv.getCell(grandRowNum, 1).font?.color?.argb], ['FF808080', 'FFFFFFFF']);
  check('grand row B:D unfilled', [2, 3, 4].map((c) => fillOf(pv.getCell(grandRowNum, c))), [null, null, null]);
  check('grand row value cells gray white + 0.00', [pv.getCell(grandRowNum, gtColNum).value, fillOf(pv.getCell(grandRowNum, gtColNum)), pv.getCell(grandRowNum, gtColNum).numFmt], [62, 'FF808080', '0.00']);

  // §2.7: Calibri 12 body, hours 0.00, widths.
  check('body font Calibri 12', [pv.getCell('C7').font?.name, pv.getCell('C7').font?.size], ['Calibri', 12]);
  check('hours cells 0.00', pv.getCell('G7').numFmt, '0.00');
  check('pivot widths (A, D, first est col, GT)', [pv.getColumn(1).width, pv.getColumn(4).width, pv.getColumn(5).width, pv.getColumn(gtColNum).width], [35, 39, 8, 8]);
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
