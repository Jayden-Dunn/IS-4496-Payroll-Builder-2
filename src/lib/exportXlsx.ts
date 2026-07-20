import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { isDarkColor } from './format';
import { weekdayShort } from './pivotByJob';
import type {
  HourlyBlock,
  LegendEntry,
  ManualSectionRow,
  PayrollPreviewModel,
  PivotByJobModel,
  PivotModel,
  PsBlock,
  ReportRow,
  SalaryBlock,
} from './types';

const BLACK = 'FF000000';
const WHITE = 'FFFFFFFF';
const HEADER_GRAY = 'FF808080';
const GRAND_TOTAL_GRAY = 'FFD9D9D9';

// Reference formats: accounting style WITHOUT a $ symbol.
const ACCOUNTING_FMT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const START_END_DATE_FMT = 'mm-dd-yy';
const INVOICE_DATE_FMT = 'mm/dd/yyyy';
const PCT_FMT = '0%';
const HOURS_FMT = '0.00';

const FONT = 'Calibri';

function solidFill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

/** ARGB text color that stays readable on a given 6-digit hex fill. */
function textOn(hex6: string): string {
  return isDarkColor(hex6) ? WHITE : BLACK;
}

/**
 * ExcelJS serializes JS Dates to serials using their UTC components. Local
 * calendar dates therefore must be re-based to UTC midnight or the sheet
 * shows a stray time (e.g. 07:00) and can flip a day in some timezones.
 */
function excelDate(d: Date): Date {
  return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
}

export interface ExportInput {
  rows: ReportRow[]; // already sorted; only included rows are exported
  pivot: PivotModel;
  legend: LegendEntry[];
  payDate: Date | null;
  /** Round 3 sheets — optional so %-only exports still work. */
  pivotByJob?: PivotByJobModel | null;
  payrollPreview?: PayrollPreviewModel | null;
}

export function exportFilename(payDate: Date | null): string {
  const d = payDate ?? new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `% ${mm}-${dd}-${d.getFullYear()} % of Completion.xlsx`;
}

/** Pure workbook construction — testable in Node without a DOM. */
export async function buildWorkbookBuffer(input: ExportInput): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.created = new Date();
  const palette = new Map(input.legend.map((l) => [l.estimator, l.color.toUpperCase()]));
  buildTimeNaviSheet(wb, input, palette);
  buildPivotSheet(wb, input.pivot, palette);
  if (input.pivotByJob) buildPivotByJobSheet(wb, input.pivotByJob);
  if (input.payrollPreview) {
    buildPayrollPreviewSheet(wb, input.payrollPreview);
    if (input.payrollPreview.inactiveZeroHours.length > 0) {
      buildInactiveSheet(wb, input.payrollPreview.inactiveZeroHours);
    }
  }
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}

export async function exportWorkbook(input: ExportInput): Promise<void> {
  const buf = await buildWorkbookBuffer(input);
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  saveAs(blob, exportFilename(input.payDate));
}

// ---------------------------------------------------------------------------
// Sheet 1 — TimeNavi Export (columns A–S)
// ---------------------------------------------------------------------------

/** Columns that come straight from the raw TimeNavi export (A..G). */
const TN_RAW_HEADERS = ['Event', 'Start Date', 'End Date', 'Calendar', 'Attendees', 'Location', 'Description'];
/** Analysis columns the app appends after the raw export block. */
const TN_APPENDED_HEADERS = [
  'AP#', '% Completion', 'Estimator', 'Last % Paid', 'QBO Job Name', 'In QBO Time?',
  'Job Total $', 'Apptivo $', '$ Difference', '%compl', 'Invoice Date', 'Date Difference',
];
const TN_HEADERS = [...TN_RAW_HEADERS, ...TN_APPENDED_HEADERS];

/** Headers hidden in the reference (hidden columns, never deleted). */
const TN_HIDDEN_HEADERS = ['Attendees', 'Location', 'Description', 'QBO Job Name'];

// Reference column widths (A–S) so screenshots line up. G (Description) stays
// collapsed at 13 — the full text lives in the cell, wrap off.
const TN_WIDTHS = [60, 11, 13, 13, 9, 13, 13, 6.5, 11.6, 9, 10, 39, 20, 12, 14, 10.6, 9, 12.4, 13];

function buildTimeNaviSheet(wb: ExcelJS.Workbook, input: ExportInput, palette: ReadonlyMap<string, string>): void {
  const ws = wb.addWorksheet('TimeNavi Export', {
    views: [{ state: 'frozen', ySplit: 1 }], // freeze at A2
  });
  ws.columns = TN_WIDTHS.map((width) => ({ width }));

  // Header: the ENTIRE row (A1 → last populated header) is black with white
  // non-bold text — no green remains. Paint by column index so every header
  // cell is covered (ExcelJS eachCell can skip cells in some edge cases).
  const header = ws.addRow(TN_HEADERS);
  for (let c = 1; c <= TN_HEADERS.length; c++) {
    const cell = header.getCell(c);
    cell.value = TN_HEADERS[c - 1];
    cell.fill = solidFill(BLACK);
    cell.font = { name: FONT, color: { argb: WHITE }, bold: false };
    cell.alignment = { vertical: 'middle', wrapText: false };
  }

  // Hide (never delete) the reference's hidden columns, located by header text.
  for (const h of TN_HIDDEN_HEADERS) {
    const idx = TN_HEADERS.indexOf(h);
    if (idx >= 0) ws.getColumn(idx + 1).hidden = true;
  }

  for (const r of input.rows.filter((row) => row.included)) {
    const percentCell = r.percent === null ? '' : r.percent;
    const row = ws.addRow([
      r.event,
      r.startDate ? excelDate(r.startDate) : '',
      r.endDate ? excelDate(r.endDate) : '',
      r.calendar,
      r.attendees,
      r.location,
      r.description,
      r.apKey ?? '',
      percentCell,
      r.estimator,
      r.lastPaid ?? '-',
      r.qboJobName,
      r.inQboTime,
      r.jobTotal, // number, or the literal text '#N/A'
      r.apptivo ?? '', // blank when N is #N/A and not user-edited
      r.dollarDiff ?? '', // blank when N is #N/A
      percentCell, // Q mirrors I — the office PDF reads this column
      r.invoiceDate ? excelDate(r.invoiceDate) : '',
      r.dateDiff ?? '',
    ]);

    row.getCell(2).numFmt = START_END_DATE_FMT;
    row.getCell(3).numFmt = START_END_DATE_FMT;
    row.getCell(18).numFmt = INVOICE_DATE_FMT;
    for (const col of [9, 11, 17]) {
      const cell = row.getCell(col);
      if (typeof cell.value === 'number') cell.numFmt = PCT_FMT;
    }
    for (const col of [14, 15, 16]) {
      const cell = row.getCell(col);
      if (typeof cell.value === 'number') cell.numFmt = ACCOUNTING_FMT;
    }
    row.getCell(19).numFmt = '0';

    row.getCell(1).alignment = { vertical: 'top' };

    // Full-row fill keyed off the RESOLVED estimator, with readable text.
    const hex = palette.get(r.estimator);
    if (hex) {
      const font = { name: FONT, color: { argb: textOn(hex) } };
      for (let c = 1; c <= TN_HEADERS.length; c++) {
        const cell = row.getCell(c);
        cell.fill = solidFill('FF' + hex);
        cell.font = font;
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Sheet 2 — Pivot by Crew (static replica of the Excel PivotTable)
// ---------------------------------------------------------------------------

function buildPivotSheet(wb: ExcelJS.Workbook, pivot: PivotModel, palette: ReadonlyMap<string, string>): void {
  const ws = wb.addWorksheet('Pivot by Crew', {
    views: [{ state: 'frozen', ySplit: 4 }], // freeze at A5, below the header block
  });
  const estCols = pivot.estimators; // '#N/A' (when present) is always last
  // Grand Total column index is derived from however many estimator columns
  // rendered — never assumed.
  const gtCol = 4 + estCols.length + 1;

  ws.columns = [
    { width: 35 }, { width: 14 }, { width: 24 }, { width: 39 },
    ...estCols.map(() => ({ width: 8 })),
    { width: 8 },
  ];

  const baseFont = { name: FONT, size: 12 };
  const paint = (
    cell: ExcelJS.Cell,
    opts: { fill?: string; color?: string; bold?: boolean; size?: number },
  ): void => {
    if (opts.fill) cell.fill = solidFill(opts.fill);
    cell.font = {
      ...baseFont,
      ...(opts.size !== undefined ? { size: opts.size } : {}),
      ...(opts.bold ? { bold: true } : {}),
      color: { argb: opts.color ?? BLACK },
    };
  };

  const blackBand = (row: ExcelJS.Row, boldCols: number[] = []): void => {
    for (let c = 1; c <= gtCol; c++) {
      paint(row.getCell(c), { fill: BLACK, color: WHITE, bold: boldCols.includes(c) });
    }
  };

  // Rows 1–3: black title block, full table width.
  const r1 = ws.addRow([pivot.title]);
  r1.height = 31.5;
  blackBand(r1);
  paint(r1.getCell(1), { fill: BLACK, color: WHITE, bold: true, size: 17 });

  // Row 2: Name/Job labels + a single "Hours" cell centered over the
  // estimator block (no merge — the reference does not merge).
  const hoursCol = 5 + Math.floor((estCols.length - 1) / 2);
  const r2 = ws.addRow(['', '', 'Name', 'Job']);
  r2.getCell(hoursCol).value = 'Hours';
  blackBand(r2, [3, 4, hoursCol]);

  const r3 = ws.addRow(['Sum of hours', '', '', '', 'Column Labels']);
  blackBand(r3);

  // Row 4: gray label cells, estimator headers in their palette colors,
  // gray #N/A + Grand Total headers. Regular weight.
  const r4 = ws.addRow(['Row Labels', 'fname', 'lname', 'jobcode_2', ...estCols, 'Grand Total']);
  r4.height = 15;
  for (let c = 1; c <= 4; c++) paint(r4.getCell(c), { fill: HEADER_GRAY, color: BLACK });
  estCols.forEach((e, i) => {
    const hex = palette.get(e);
    paint(r4.getCell(5 + i), hex ? { fill: 'FF' + hex, color: textOn(hex) } : { fill: HEADER_GRAY, color: BLACK });
  });
  paint(r4.getCell(gtCol), { fill: HEADER_GRAY, color: BLACK });

  const hoursCells = (byEstimator: Record<string, number>, total: number): (number | '')[] => [
    ...estCols.map((e) => (byEstimator[e] !== undefined ? byEstimator[e] : '' as const)),
    total,
  ];

  /**
   * Body row: full-height estimator column bands + gray Grand Total band.
   * Bold (Round 2): employee rows bold A:C; job rows bold the customer label
   * in C only; "X Total" rows bold across, with a thin bottom border closing
   * the employee block.
   */
  const styleBodyRow = (row: ExcelJS.Row, kind: 'employee' | 'job' | 'total'): void => {
    const boldAt = (c: number): boolean =>
      kind === 'total' ? true : kind === 'employee' ? c <= 3 : c === 3;
    for (let c = 1; c <= 4; c++) paint(row.getCell(c), { bold: boldAt(c) });
    estCols.forEach((e, i) => {
      const hex = palette.get(e);
      const cell = row.getCell(5 + i);
      paint(cell, { ...(hex ? { fill: 'FF' + hex, color: textOn(hex) } : {}), bold: boldAt(5 + i) });
      if (typeof cell.value === 'number') cell.numFmt = HOURS_FMT;
    });
    const gt = row.getCell(gtCol);
    paint(gt, { fill: GRAND_TOTAL_GRAY, bold: boldAt(gtCol) });
    if (typeof gt.value === 'number') gt.numFmt = HOURS_FMT;
    if (kind === 'total') {
      for (let c = 1; c <= gtCol; c++) {
        row.getCell(c).border = { bottom: { style: 'thin' } };
      }
    }
  };

  for (const crew of pivot.crews) {
    // Crew-lead rows: full-width black band, overrides the column bands.
    blackBand(ws.addRow([crew.crew]));

    for (const emp of crew.employees) {
      styleBodyRow(ws.addRow([emp.username, emp.fname, emp.lname]), 'employee');
      for (const job of emp.jobs) {
        styleBodyRow(ws.addRow(['', '', job.customer, job.job, ...hoursCells(job.byEstimator, job.total)]), 'job');
      }
      styleBodyRow(ws.addRow(['', `${emp.fname} Total`, '', '', ...hoursCells(emp.byEstimator, emp.total)]), 'total');
    }
  }

  // Grand Total row: bold, gray/white on column A and every value column; B:D plain.
  const grandRow = ws.addRow(['Grand Total', '', '', '', ...hoursCells(pivot.grandByEstimator, pivot.grandTotal)]);
  paint(grandRow.getCell(1), { fill: HEADER_GRAY, color: WHITE, bold: true });
  for (let c = 2; c <= 4; c++) paint(grandRow.getCell(c), {});
  for (let c = 5; c <= gtCol; c++) {
    const cell = grandRow.getCell(c);
    paint(cell, { fill: HEADER_GRAY, color: WHITE, bold: true });
    if (typeof cell.value === 'number') cell.numFmt = HOURS_FMT;
  }
}

// ---------------------------------------------------------------------------
// Sheet 3 — Pivot by Job (Round 3 §4/§8)
// ---------------------------------------------------------------------------

const RED = 'FFFF0000';
const DARK_RED = 'FFC00000';
const ACCT_NO_DOLLAR = '#,##0.00';

function buildPivotByJobSheet(wb: ExcelJS.Workbook, model: PivotByJobModel): void {
  const ws = wb.addWorksheet('Pivot by Job', {
    views: [{ state: 'frozen', ySplit: 4 }], // freeze at A5
  });
  const dateCount = model.dateIsos.length;
  // A · dates · Total · QB Hrs(K) · AP# · OnCal · Est · Lead · WO$ · WO% · % · Notes
  const analysisStart = 1 + dateCount + 2 + 1; // first analysis col (AP#)
  const lastCol = analysisStart + 7;
  ws.columns = [
    { width: 40 },
    ...model.dateIsos.map(() => ({ width: 10 })),
    { width: 11 },
    { width: 7.6 },
    { width: 10.5 },
    { width: 17.8 },
    { width: 9.4 },
    { width: 43 },
    { width: 23 },
    { width: 12 },
    { width: 11 },
    { width: 18.5 },
  ];

  const baseFont = { name: FONT, size: 12 };
  const paint = (
    cell: ExcelJS.Cell,
    opts: { fill?: string; color?: string; bold?: boolean; size?: number },
  ): void => {
    if (opts.fill) cell.fill = solidFill(opts.fill);
    cell.font = {
      ...baseFont,
      ...(opts.size !== undefined ? { size: opts.size } : {}),
      ...(opts.bold ? { bold: true } : {}),
      color: { argb: opts.color ?? BLACK },
    };
  };

  const totalCol = 1 + dateCount + 1;
  const qbCol = totalCol + 1;

  const headerBlock = (title: string, withDates: boolean): void => {
    // Row 1 — black: title, dates, Total hrs, analysis headers.
    const r1Vals: (string | Date)[] = [title];
    for (let i = 0; i < dateCount; i++) {
      r1Vals.push(withDates && i < model.dates.length ? excelDate(model.dates[i]) : '');
    }
    r1Vals.push('Total hrs', '', 'AP#', 'On Calendar?', 'Estimator', 'Crew Lead', 'WO Override', 'WO%', '% of Completion', 'Notes');
    const r1 = ws.addRow(r1Vals);
    r1.height = 31.5;
    for (let c = 1; c <= lastCol; c++) {
      const color = c === lastCol ? DARK_RED : c === lastCol - 1 ? RED : WHITE;
      paint(r1.getCell(c), { fill: BLACK, color, bold: c === 1 });
      if (c === 1) paint(r1.getCell(1), { fill: BLACK, color: WHITE, bold: true, size: 17 });
      if (c > 1 && c <= 1 + dateCount && r1.getCell(c).value instanceof Date) r1.getCell(c).numFmt = 'm/d';
    }
    // Row 2 — black: computed weekday names (Sunday included).
    const r2 = ws.addRow(['']);
    for (let c = 1; c <= lastCol; c++) paint(r2.getCell(c), { fill: BLACK, color: WHITE });
    if (withDates) {
      model.dates.forEach((d, i) => {
        r2.getCell(2 + i).value = weekdayShort(d);
      });
    }
    // Row 3 — black: Sum of hours / Column Labels + repeated analysis labels.
    const r3 = ws.addRow(['Sum of hours', 'Column Labels']);
    for (let c = 1; c <= lastCol; c++) paint(r3.getCell(c), { fill: BLACK, color: WHITE });
    ['AP#', 'Estimator', 'Crew Lead', 'WO Amt', 'WO%', '% of Completion', 'Notes'].forEach((label, i) => {
      const c = analysisStart + (i === 0 ? 0 : i + 1); // skip On Calendar? on this row
      r3.getCell(c).value = label;
    });
    // Row 4 — Row Labels + dates + Grand Total.
    const r4 = ws.addRow(['Row Labels']);
    paint(r4.getCell(1), { bold: true });
    if (withDates) {
      model.dates.forEach((d, i) => {
        const cell = r4.getCell(2 + i);
        cell.value = excelDate(d);
        cell.numFmt = 'm/d';
        paint(cell, { bold: true });
      });
    }
    const gt = r4.getCell(totalCol);
    gt.value = 'Grand Total';
    paint(gt, { bold: true });
  };

  headerBlock(model.title, true);

  const hourVals = (byDate: Record<string, number>): (number | '')[] =>
    model.dateIsos.map((iso) => (byDate[iso] !== undefined ? byDate[iso] : ('' as const)));

  const fillRow = (row: ExcelJS.Row, hex: string | null, bold = false, upTo = lastCol): void => {
    for (let c = 1; c <= upTo; c++) {
      const cell = row.getCell(c);
      paint(cell, { ...(hex ? { fill: 'FF' + hex, color: textOn(hex) } : {}), bold });
      if (typeof cell.value === 'number' && c <= qbCol && c > 1) cell.numFmt = HOURS_FMT;
    }
  };

  for (const cust of model.customers) {
    const cRow = ws.addRow([cust.customer, ...hourVals(cust.byDate), cust.total]);
    fillRow(cRow, cust.colorHex, true);
    for (const job of cust.jobs) {
      const a = job.analysis;
      const jRow = ws.addRow([
        job.jobcode2,
        ...hourVals(job.byDate),
        job.total,
        '', // QB Hrs (job level blank)
        a.apKey ?? '',
        a.onCalendar,
        a.estimator,
        a.crewLead,
        a.hourly ? '-' : (a.woOverride ?? ''),
        a.hourly ? '-' : (a.woPct ?? ''),
        a.percent === null ? '' : a.percent,
        a.notes,
      ]);
      fillRow(jRow, job.colorHex);
      // Job-row analysis cells bold (§4.4) + formats (§4.5).
      for (let c = analysisStart; c <= lastCol; c++) {
        const cell = jRow.getCell(c);
        const color =
          c === lastCol
            ? job.colorHex && isDarkColor(job.colorHex)
              ? 'FFFFC7C7'
              : DARK_RED
            : job.colorHex
              ? textOn(job.colorHex)
              : BLACK;
        paint(cell, { ...(job.colorHex ? { fill: 'FF' + job.colorHex } : {}), color, bold: true });
      }
      const woCell = jRow.getCell(analysisStart + 4);
      if (typeof woCell.value === 'number') woCell.numFmt = ACCT_NO_DOLLAR;
      const woPctCell = jRow.getCell(analysisStart + 5);
      if (typeof woPctCell.value === 'number') woPctCell.numFmt = PCT_FMT;
      const pctCell = jRow.getCell(analysisStart + 6);
      if (typeof pctCell.value === 'number') pctCell.numFmt = PCT_FMT;

      for (const emp of job.employees) {
        const eRow = ws.addRow([emp.username, ...hourVals(emp.byDate), emp.total, emp.qbHours ?? '']);
        fillRow(eRow, job.colorHex);
        const qb = eRow.getCell(qbCol);
        if (typeof qb.value === 'number') {
          qb.numFmt = HOURS_FMT;
          qb.font = { ...baseFont, color: { argb: DARK_RED } };
        }
      }
    }
  }

  // Grand Total row — gray, white text on A and the analysis columns.
  const gRow = ws.addRow([
    'Grand Total',
    ...model.dateIsos.map((iso) => model.grandByDate[iso] ?? ('' as const)),
    model.grandTotal,
  ]);
  for (let c = 1; c <= lastCol; c++) {
    const cell = gRow.getCell(c);
    paint(cell, { fill: HEADER_GRAY, color: WHITE, bold: true });
    if (typeof cell.value === 'number') cell.numFmt = HOURS_FMT;
  }

  // §4.6 — BALANCES and UT sections auto-position below the pivot.
  const manualSection = (title: string, rows: ManualSectionRow[]): void => {
    ws.addRow([]);
    headerBlock(title, false);
    for (const r of rows) {
      if (r.customer) {
        const cRow = ws.addRow([r.customer]);
        fillRow(cRow, r.colorHex, true);
      }
      const jRow = ws.addRow([r.job]);
      const cells: (string | number)[] = [
        r.apKey ?? '',
        '',
        r.estimator,
        r.crewLead,
        r.woOverride ?? '',
        r.woPct ?? '',
        r.percent === null ? '' : r.percent === 'HOURLY' ? 'HOURLY' : r.percent,
        r.notes,
      ];
      cells.forEach((v, i) => {
        jRow.getCell(analysisStart + i).value = v;
      });
      fillRow(jRow, r.colorHex);
      const wo = jRow.getCell(analysisStart + 4);
      if (typeof wo.value === 'number') wo.numFmt = ACCT_NO_DOLLAR;
      for (const c of [analysisStart + 5, analysisStart + 6]) {
        const cell = jRow.getCell(c);
        if (typeof cell.value === 'number') cell.numFmt = PCT_FMT;
      }
    }
  };
  manualSection(model.title.replace('NV PAYROLL:', 'NV PAYROLL (BALANCES):'), model.balances);
  manualSection(model.title.replace('NV PAYROLL:', 'UT PAYROLL:'), model.utPayroll);
}

// ---------------------------------------------------------------------------
// Sheet 4 — Payroll Preview (Round 3 §5/§8, with live formulas)
// ---------------------------------------------------------------------------

const TAN = 'FFD6D4CA';
const NEAR_WHITE = 'FFF9F9F9';
const MANAGER_FILL = 'FFD99694';
const OFFICE_BLUE = 'FF4F81BD';
const MONEY_FMT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';

function buildPayrollPreviewSheet(wb: ExcelJS.Workbook, model: PayrollPreviewModel): void {
  const ws = wb.addWorksheet('Payroll Preview');
  ws.columns = [{ width: 22 }, { width: 41.4 }, { width: 32.4 }, { width: 16.2 }, { width: 9.4 }, { width: 12 }, { width: 12 }];
  const baseFont = { name: FONT, size: 11 };

  const band = (values: (string | number)[], fill: string, color: string, bold = true, upTo = 7): ExcelJS.Row => {
    const row = ws.addRow(values);
    for (let c = 1; c <= upTo; c++) {
      row.getCell(c).fill = solidFill(fill);
      row.getCell(c).font = { ...baseFont, bold, color: { argb: color } };
    }
    return row;
  };

  band([model.title], BLACK, WHITE);
  ws.addRow([]);

  const hourlyBlock = (block: HourlyBlock, headerFill = TAN): void => {
    band([block.title], headerFill, BLACK);
    const firstLineRow = ws.rowCount + 1;
    for (const l of block.lines) {
      const r = ws.addRow([l.apKey ? `AP${l.apKey}` : '', l.jobLabel, l.hours, l.rate]);
      const n = r.number;
      r.getCell(5).value = { formula: `C${n}*D${n}` };
      r.getCell(3).numFmt = HOURS_FMT;
      r.getCell(4).numFmt = HOURS_FMT;
      r.getCell(5).numFmt = ACCOUNTING_FMT;
      if (l.colorHex) {
        for (const c of [1, 2]) {
          r.getCell(c).fill = solidFill('FF' + l.colorHex);
          r.getCell(c).font = { ...baseFont, color: { argb: textOn(l.colorHex) } };
        }
      }
      if (l.ot) {
        r.getCell(3).fill = solidFill('FFFFFF00');
      }
    }
    const lastLineRow = ws.rowCount;
    const sums = ws.addRow(['', '']);
    const n = sums.number;
    sums.getCell(3).value = { formula: `SUM(C${firstLineRow}:C${lastLineRow})` };
    sums.getCell(5).value = { formula: `SUM(E${firstLineRow}:E${lastLineRow})` };
    sums.getCell(6).value = block.net !== null ? { formula: `E${n}-G${n}` } : '';
    sums.getCell(7).value = block.net ?? '';
    sums.getCell(3).numFmt = HOURS_FMT;
    for (const c of [5, 6, 7]) sums.getCell(c).numFmt = ACCOUNTING_FMT;
    sums.font = { ...baseFont, bold: true };
    ws.addRow([]);
  };

  const colLetter = (i: number): string => String.fromCharCode(65 + i); // 0→A

  const psBlock = (block: PsBlock): void => {
    const jobCount = block.jobs.length;
    const hdr = band(
      [block.leadName, ...block.jobs.map((j) => (j.woPct !== null ? j.woPct : '-'))],
      TAN,
      BLACK,
      true,
      1 + jobCount,
    );
    for (let i = 0; i < jobCount; i++) {
      const cell = hdr.getCell(2 + i);
      if (typeof cell.value === 'number') cell.numFmt = PCT_FMT;
    }
    const hdrRow = hdr.number;
    ws.addRow(['PO', ...block.jobs.map((j) => j.jobShort)]);
    ws.addRow(['Name', ...block.jobs.map((j) => j.customer)]);

    // Value rows first (the formula chain references them).
    const psRow = ws.addRow(['Profit Share']).number;
    const woRow = ws.addRow(['Work Order Override', ...block.jobs.map((j) => j.workOrderOverride)]).number;
    const pctRow = ws.addRow(['% Completion', ...block.jobs.map((j) => j.percentCompletion)]).number;
    const laborRow = ws.addRow(['- Labor + Wages', ...block.jobs.map((j) => j.laborWages)]).number;
    const milesRow = ws.addRow(['-Mileage', ...block.jobs.map((j) => j.mileage)]).number;
    const halfRow = ws.addRow(['-½ Materials', ...block.jobs.map((j) => j.halfMaterials)]).number;
    const tpsRow = ws.addRow(['Total Profit Share']).number;
    let splitRow = 0;
    if (block.jobs.some((j) => j.splitWith > 1)) {
      splitRow = ws.addRow(['split w/', ...block.jobs.map((j) => (j.splitWith > 1 ? j.splitWith : ''))]).number;
    }
    ws.addRow([]);
    const totalRow = ws.addRow(['Total']).number;
    const mileNewRow = ws.addRow(['Mileage - New', block.mileageNew]).number;
    const hourlyPayRow = ws.addRow(['Hourly Pay', block.hourlyPay]).number;
    const reimbRow = ws.addRow(['Reimbursements', block.reimbursements, block.reimbMemo]).number;
    const deductRow = ws.addRow(['Deductions/Sundries', block.deductions, block.deductMemo]).number;
    const totalPayRow = ws.addRow(['Total Pay']).number;
    ws.addRow([]);
    const incomeRow = ws.addRow(['Total Income', ...block.jobs.map((j) => j.income)]).number;
    const matRow = ws.addRow(['Materials', ...block.jobs.map((j) => j.materials)]).number;
    ws.addRow(['QB Labor-Wages-Subs', ...block.jobs.map((j) => j.laborWages)]);
    ws.addRow(['Mileage in Books', ...block.jobs.map((j) => j.mileage)]);
    const pctMatRow = ws.addRow(['% Materials']).number;

    for (let i = 0; i < jobCount; i++) {
      const L = colLetter(1 + i); // B, C, …
      ws.getCell(`${L}${psRow}`).value = { formula: `${L}${incomeRow}*${L}${hdrRow}` };
      const tps = ws.getCell(`${L}${tpsRow}`);
      tps.value = {
        formula: `${L}${woRow}*${L}${pctRow}-${L}${laborRow}-${L}${milesRow}-${L}${halfRow}`,
      };
      tps.font = { ...baseFont, bold: true };
      ws.getCell(`${L}${pctMatRow}`).value = {
        formula: `IF(${L}${incomeRow}=0,0,${L}${matRow}/${L}${incomeRow})`,
      };
      ws.getCell(`${L}${pctRow}`).numFmt = PCT_FMT;
      ws.getCell(`${L}${pctMatRow}`).numFmt = PCT_FMT;
      for (const rn of [psRow, woRow, laborRow, milesRow, halfRow, tpsRow, incomeRow, matRow]) {
        ws.getCell(`${L}${rn}`).numFmt = MONEY_FMT;
      }
    }

    // Total: SUM across the Total Profit Share row (÷ split when present).
    const totalCell = ws.getCell(`B${totalRow}`);
    if (splitRow > 0) {
      totalCell.value = block.total; // splits make a single SUM wrong — write the value
    } else {
      totalCell.value = { formula: `SUM(B${tpsRow}:${colLetter(jobCount)}${tpsRow})` };
    }
    ws.getCell(`B${totalPayRow}`).value = { formula: `SUM(B${totalRow}:B${deductRow})` };
    for (const rn of [totalRow, mileNewRow, hourlyPayRow, reimbRow, deductRow, totalPayRow]) {
      ws.getCell(`B${rn}`).numFmt = ACCOUNTING_FMT;
      ws.getCell(`A${rn}`).font = { ...baseFont, bold: rn === totalRow || rn === totalPayRow };
      ws.getCell(`B${rn}`).font = { ...baseFont, bold: rn === totalRow || rn === totalPayRow };
    }
    ws.addRow([]);

    if (block.tsheets) {
      band([`${block.tsheets.employeeName.split(' ')[0]}'s T-SHEETS BREAKDOWN`], NEAR_WHITE, BLACK);
      hourlyBlock(block.tsheets, NEAR_WHITE);
    }
  };

  const salaryBlock = (block: SalaryBlock, fill: string): void => {
    band([`${block.name}${block.timesheetHours > 0 ? ` (${block.timesheetHours}hrs)` : ''}`], fill, BLACK);
    const r1 = ws.addRow([block.role, '', block.amount]);
    r1.getCell(3).numFmt = ACCOUNTING_FMT;
    // Reimb already computed at the Settings mileage rate.
    const r2 = ws.addRow([`mileage reimb: ${block.miles} miles`, '', block.mileageReimb]);
    r2.getCell(3).numFmt = ACCOUNTING_FMT;
    const r3 = ws.addRow(['Total', '', '']);
    r3.getCell(3).value = { formula: `C${r1.number}+C${r2.number}` };
    r3.getCell(3).numFmt = ACCOUNTING_FMT;
    r3.font = { ...baseFont, bold: true };
    ws.addRow([]);
  };

  for (const b of model.leads) psBlock(b);
  for (const b of model.painters) hourlyBlock(b);

  band(['OFFICE'], BLACK, WHITE);
  for (const b of model.office) {
    // Office job lines fill blue per §5.1.
    const patched: HourlyBlock = {
      ...b,
      lines: b.lines.map((l) => ({ ...l, colorHex: OFFICE_BLUE.slice(2) })),
    };
    hourlyBlock(patched);
  }

  band(['ADMIN'], BLACK, WHITE);
  for (const b of model.admin) salaryBlock(b, TAN);
  band(['ESTIMATORS'], BLACK, WHITE);
  for (const b of model.estimators) salaryBlock(b, TAN);
  band(['MANAGER SALARY'], BLACK, WHITE);
  for (const b of model.managers) salaryBlock(b, MANAGER_FILL);
}

function buildInactiveSheet(wb: ExcelJS.Workbook, names: string[]): void {
  const ws = wb.addWorksheet('Inactive Employees');
  ws.columns = [{ width: 40 }];
  const hdr = ws.addRow(['Inactive Employees (zero hours this week)']);
  hdr.getCell(1).fill = solidFill(BLACK);
  hdr.getCell(1).font = { name: FONT, bold: true, color: { argb: WHITE } };
  for (const n of names) ws.addRow([n]);
}
