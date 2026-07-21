import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { isDarkColor } from './format';
import { weekdayShort } from './pivotByJob';
import { dollarDiffIsFlagged, duplicateInQboValues, percentIsFlagged } from './timenavi';
import type {
  ExceptionGroup,
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

// Round 4 §1.1 — Excel's classic "bad" style for the pink review flags.
const PINK_FILL = 'FFFFC7CE';
const PINK_TEXT = 'FF9C0006';

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
  /** Round 5 §6 — on-screen exception panels, snapshotted at download time. */
  percentExceptions?: ExceptionGroup[] | null;
  payrollExceptions?: ExceptionGroup[] | null;
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
  if (input.payrollPreview) buildPayrollPreviewSheet(wb, input.payrollPreview);
  // Round 5 §6 — exceptions sheets always written (empty panel → "No exceptions").
  buildExceptionsSheet(wb, "%'s Exceptions", input.percentExceptions ?? [], input.payDate);
  buildExceptionsSheet(wb, 'Payroll Exceptions', input.payrollExceptions ?? [], input.payDate);
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

/** Round 4 §1.2 — horizontally centered columns, resolved by header name. */
const TN_CENTERED_HEADERS = ['AP#', '% Completion', 'Estimator', 'Last % Paid', '$ Difference'];

// Reference column widths (A–S) so screenshots line up. G (Description) stays
// collapsed at 13 — the full text lives in the cell, wrap off.
const TN_WIDTHS = [60, 11, 13, 13, 9, 13, 13, 6.5, 11.6, 9, 10, 39, 20, 12, 14, 10.6, 9, 12.4, 13];

/** 1-based sheet column number for a TimeNavi header. */
function tnCol(header: string): number {
  return TN_HEADERS.indexOf(header) + 1;
}

function buildTimeNaviSheet(wb: ExcelJS.Workbook, input: ExportInput, palette: ReadonlyMap<string, string>): void {
  const ws = wb.addWorksheet('TimeNavi Export', {
    views: [{ state: 'frozen', ySplit: 1 }], // freeze at A2
  });
  // Round 5 §2.1 — AP# best-fit: max(8, longest rendered AP# + 1).
  const apIdx = TN_HEADERS.indexOf('AP#');
  const included = input.rows.filter((row) => row.included);
  let apWidth = 8;
  for (const r of included) {
    const len = (r.apKey ?? '').length + 1;
    if (len > apWidth) apWidth = len;
  }
  ws.columns = TN_WIDTHS.map((width, i) => ({ width: i === apIdx ? apWidth : width }));

  const centeredCols = TN_CENTERED_HEADERS.map(tnCol);

  // Header: the ENTIRE row (A1 → last populated header) is black with white
  // non-bold text — no green remains. Paint by column index so every header
  // cell is covered (ExcelJS eachCell can skip cells in some edge cases).
  const header = ws.addRow(TN_HEADERS);
  for (let c = 1; c <= TN_HEADERS.length; c++) {
    const cell = header.getCell(c);
    cell.value = TN_HEADERS[c - 1];
    cell.fill = solidFill(BLACK);
    cell.font = { name: FONT, color: { argb: WHITE }, bold: false };
    cell.alignment = {
      vertical: 'middle',
      wrapText: false,
      ...(centeredCols.includes(c) ? { horizontal: 'center' as const } : {}),
    };
  }

  // Hide (never delete) the reference's hidden columns, located by header text.
  for (const h of TN_HIDDEN_HEADERS) {
    const idx = TN_HEADERS.indexOf(h);
    if (idx >= 0) ws.getColumn(idx + 1).hidden = true;
  }

  const dupInQbo = duplicateInQboValues(input.rows);
  const pctCol = tnCol('% Completion');
  const inQboCol = tnCol('In QBO Time?');
  const diffCol = tnCol('$ Difference');

  for (const r of included) {
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

    // §1.2 — centered columns (over any row fill).
    for (const c of centeredCols) {
      row.getCell(c).alignment = { horizontal: 'center' };
    }

    // §1.1 — pink review flags override the row fill on just these cells.
    const pink = (c: number): void => {
      const cell = row.getCell(c);
      cell.fill = solidFill(PINK_FILL);
      cell.font = { name: FONT, color: { argb: PINK_TEXT } };
    };
    if (percentIsFlagged(r.percent)) pink(pctCol);
    if (r.inQboTime === 'NotClockedIn' || dupInQbo.has(r.inQboTime.trim())) pink(inQboCol);
    if (dollarDiffIsFlagged(r.dollarDiff)) pink(diffCol);
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
   * in C only; "X Total" rows bold across, with a thin TOP border (Round 4
   * §2.4 — the line sits between the last job row and the total) spanning
   * from C through the Grand Total column. Hour values center (§2.4); the
   * Grand Total column keeps its default right alignment.
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
      cell.alignment = { horizontal: 'center' };
    });
    const gt = row.getCell(gtCol);
    paint(gt, { fill: GRAND_TOTAL_GRAY, bold: boldAt(gtCol) });
    if (typeof gt.value === 'number') gt.numFmt = HOURS_FMT;
    if (kind === 'total') {
      for (let c = 3; c <= gtCol; c++) {
        row.getCell(c).border = { top: { style: 'thin' } };
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
    if (c < gtCol) cell.alignment = { horizontal: 'center' };
  }
}

// ---------------------------------------------------------------------------
// Sheet 3 — Pivot by Job (Round 5 §3: two-row header, outline, borders)
// ---------------------------------------------------------------------------

const RED = 'FFFF0000';
const DARK_RED = 'FFC00000';
// Round 4 §3.2-1 — WO Override uses the $ accounting variant on this tab.
const MONEY_FMT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';

function buildPivotByJobSheet(wb: ExcelJS.Workbook, model: PivotByJobModel): void {
  const ws = wb.addWorksheet('Pivot by Job', {
    views: [{ state: 'frozen', ySplit: 2 }], // Round 5 §3.1 — freeze at A3
    properties: { outlineProperties: { summaryBelow: false, summaryRight: true } },
  });
  const dateCount = model.dateIsos.length;
  // A · dates · Total · QB Hrs · AP# · OnCal · Est · Lead · WO$ · WO% · % · Notes
  const analysisStart = 1 + dateCount + 2 + 1; // first analysis col (AP#)
  const lastCol = analysisStart + 7;
  const pctComplCol = lastCol - 1;
  ws.columns = [
    { width: 40 },
    ...model.dateIsos.map(() => ({ width: 10 })),
    { width: 11 },
    { width: 7.6 },
    { width: 10.5 },
    { width: 12 },
    { width: 9.4 },
    { width: 43 },
    { width: 23 },
    { width: 12 },
    { width: 15.78 }, // Round 5 §3.4 — % of Completion fits its header
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

  const centerAnalysis = (row: ExcelJS.Row): void => {
    for (let c = analysisStart; c < lastCol; c++) {
      row.getCell(c).alignment = { horizontal: 'center' };
    }
  };

  /** Round 5 §3.5 — thin top+bottom on A through QB Hrs only. */
  const hoursBandBorder = (row: ExcelJS.Row): void => {
    for (let c = 1; c <= qbCol; c++) {
      const prev = row.getCell(c).border ?? {};
      row.getCell(c).border = { ...prev, top: { style: 'thin' }, bottom: { style: 'thin' } };
    }
  };

  /**
   * Round 5 §3.1 / §3.2 — two-row header when `withDays`, else a single black
   * section header (BALANCES / UT) mirroring row 1 only.
   */
  const headerBlock = (title: string, withDays: boolean): void => {
    const r1Vals: (string | Date)[] = [title];
    for (let i = 0; i < dateCount; i++) {
      r1Vals.push(i < model.dates.length ? excelDate(model.dates[i]) : '');
    }
    r1Vals.push('Total hrs', '', 'AP#', 'On Calendar?', 'Estimator', 'Crew Lead', 'WO Override', 'WO%', '% of Completion', 'Notes');
    const r1 = ws.addRow(r1Vals);
    r1.height = 31.5;
    for (let c = 1; c <= lastCol; c++) {
      const color = c === lastCol ? DARK_RED : c === pctComplCol ? RED : WHITE;
      paint(r1.getCell(c), { fill: BLACK, color, bold: true });
      if (c === 1) paint(r1.getCell(1), { fill: BLACK, color: WHITE, bold: true, size: withDays ? 17 : 16 });
      if (c > 1 && c <= 1 + dateCount && r1.getCell(c).value instanceof Date) {
        r1.getCell(c).numFmt = 'm/d/yyyy';
        r1.getCell(c).alignment = { horizontal: 'center' };
      }
    }
    centerAnalysis(r1);
    if (!withDays) return;
    // Row 2 — computed weekday names (Sunday included).
    const r2 = ws.addRow(['']);
    for (let c = 1; c <= lastCol; c++) paint(r2.getCell(c), { fill: BLACK, color: WHITE });
    model.dates.forEach((d, i) => {
      const cell = r2.getCell(2 + i);
      cell.value = weekdayShort(d);
      cell.alignment = { horizontal: 'center' };
    });
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

  const analysisFormats = (row: ExcelJS.Row): void => {
    const woCell = row.getCell(analysisStart + 4);
    if (typeof woCell.value === 'number') woCell.numFmt = MONEY_FMT;
    for (const c of [analysisStart + 5, analysisStart + 6]) {
      const cell = row.getCell(c);
      if (typeof cell.value === 'number') cell.numFmt = PCT_FMT;
    }
  };

  for (const cust of model.customers) {
    const cRow = ws.addRow([cust.customer, ...hourVals(cust.byDate), cust.total]);
    cRow.outlineLevel = 0;
    fillRow(cRow, cust.colorHex, true);
    hoursBandBorder(cRow);
    for (const job of cust.jobs) {
      const a = job.analysis;
      const jRow = ws.addRow([
        job.jobcode2,
        ...hourVals(job.byDate),
        job.total,
        '',
        a.apKey ?? '',
        a.onCalendar,
        a.estimator,
        a.crewLead,
        a.hourly ? '-' : (a.woOverride ?? ''),
        a.hourly ? '-' : (a.woPct ?? ''),
        a.percent === null ? '' : a.percent,
        a.notes,
      ]);
      jRow.outlineLevel = 1;
      fillRow(jRow, job.colorHex);
      // Round 5 §3.6 — bold the job label in A (and analysis cells as before).
      paint(jRow.getCell(1), { ...(job.colorHex ? { fill: 'FF' + job.colorHex, color: textOn(job.colorHex) } : {}), bold: true });
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
      analysisFormats(jRow);
      centerAnalysis(jRow);
      hoursBandBorder(jRow);

      for (const emp of job.employees) {
        const eRow = ws.addRow([emp.username, ...hourVals(emp.byDate), emp.total, emp.qbHours ?? '']);
        eRow.outlineLevel = 2;
        fillRow(eRow, job.colorHex);
        const qb = eRow.getCell(qbCol);
        if (typeof qb.value === 'number') {
          qb.numFmt = HOURS_FMT;
          qb.font = { ...baseFont, color: { argb: DARK_RED } };
        }
      }
    }
  }

  // Grand Total row — gray, white text, double top border.
  const gRow = ws.addRow([
    'Grand Total',
    ...model.dateIsos.map((iso) => model.grandByDate[iso] ?? ('' as const)),
    model.grandTotal,
  ]);
  gRow.outlineLevel = 0;
  for (let c = 1; c <= lastCol; c++) {
    const cell = gRow.getCell(c);
    paint(cell, { fill: HEADER_GRAY, color: WHITE, bold: true });
    if (typeof cell.value === 'number') cell.numFmt = HOURS_FMT;
    cell.border = { top: { style: 'double' } };
  }

  // Round 5 §3.2 — single black header row per manual section; pair box-borders kept.
  const manualSection = (title: string, rows: ManualSectionRow[]): void => {
    ws.addRow([]);
    headerBlock(title, false);
    for (const r of rows) {
      const cRow = ws.addRow([r.customer]);
      cRow.outlineLevel = 0;
      fillRow(cRow, r.colorHex, true);
      const jRow = ws.addRow([r.job]);
      jRow.outlineLevel = 1;
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
      paint(jRow.getCell(1), { ...(r.colorHex ? { fill: 'FF' + r.colorHex, color: textOn(r.colorHex) } : {}), bold: true });
      const notesCell = jRow.getCell(lastCol);
      paint(notesCell, {
        ...(r.colorHex ? { fill: 'FF' + r.colorHex } : {}),
        color: r.colorHex && isDarkColor(r.colorHex) ? 'FFFFC7C7' : DARK_RED,
      });
      analysisFormats(jRow);
      centerAnalysis(jRow);
      for (let c = 1; c <= lastCol; c++) {
        cRow.getCell(c).border = {
          top: { style: 'thin' },
          ...(c === 1 ? { left: { style: 'thin' } } : {}),
          ...(c === lastCol ? { right: { style: 'thin' } } : {}),
        };
        jRow.getCell(c).border = {
          bottom: { style: 'thin' },
          ...(c === 1 ? { left: { style: 'thin' } } : {}),
          ...(c === lastCol ? { right: { style: 'thin' } } : {}),
        };
      }
    }
  };
  manualSection(model.title.replace('NV PAYROLL:', 'NV PAYROLL (BALANCES):'), model.balances);
  manualSection(model.title.replace('NV PAYROLL:', 'UT PAYROLL:'), model.utPayroll);
}

// ---------------------------------------------------------------------------
// Sheet 4 — Payroll Preview (Round 3 §5/§8 + Round 4 §5: Garamond, bands,
// labeled columns, boxed T-SHEETS, PS-block borders/fills/formats)
// ---------------------------------------------------------------------------

// Round 4 supersession: tan is FFDDD9C4; the whole sheet is Garamond.
const TAN = 'FFDDD9C4';
const NEAR_WHITE = 'FFF9F9F9';
const MANAGER_FILL = 'FFD99694';
const OFFICE_BLUE = 'FF4F81BD';
const HOURLY_PAY_YELLOW = 'FFFFFF99';
const REIMB_BLUE = 'FF0070C0';
const DEDUCT_RED = 'FFFF0000';

const PP_FONT = 'Garamond';

function buildPayrollPreviewSheet(wb: ExcelJS.Workbook, model: PayrollPreviewModel): void {
  const ws = wb.addWorksheet('Payroll Preview');
  ws.columns = [{ width: 22 }, { width: 41.4 }, { width: 32.4 }, { width: 16.2 }, { width: 9.4 }, { width: 12 }, { width: 12 }, { width: 12 }];
  const baseFont = { name: PP_FONT, size: 12 };

  const band = (values: (string | number)[], fill: string, color: string, bold = true, upTo = 7): ExcelJS.Row => {
    const row = ws.addRow(values);
    for (let c = 1; c <= upTo; c++) {
      row.getCell(c).fill = solidFill(fill);
      row.getCell(c).font = { ...baseFont, bold, color: { argb: color } };
    }
    return row;
  };

  // §5.4 — title merged A1:H1, centered, Garamond 13 bold white on black.
  const titleRow = band([model.title], BLACK, WHITE, true, 8);
  titleRow.getCell(1).font = { ...baseFont, size: 13, bold: true, color: { argb: WHITE } };
  ws.mergeCells(titleRow.number, 1, titleRow.number, 8);
  titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.addRow([]);

  const center = { horizontal: 'center' as const };

  /** §5.5 — the labeled `PO# | Job | hrs | rate | total` header row. */
  const columnHeaderRow = (fill: string | null): ExcelJS.Row => {
    const row = ws.addRow(['PO#', 'Job', 'hrs', 'rate', 'total']);
    for (let c = 1; c <= 5; c++) {
      const cell = row.getCell(c);
      cell.font = { ...baseFont, bold: true, underline: true };
      cell.alignment = center;
      if (fill) cell.fill = solidFill(fill);
    }
    return row;
  };

  /**
   * Hourly block (§5.5 regular painters / office / §5.7 T-SHEETS variant).
   * `tsheets` switches on the near-white boxed-table styling.
   */
  const hourlyBlock = (block: HourlyBlock, opts: { tsheets?: boolean; nameBand?: boolean } = {}): void => {
    const isTs = opts.tsheets === true;
    const bodyFill = isTs ? NEAR_WHITE : null;

    if (opts.nameBand !== false && !isTs) {
      // §5.4 — merged A–H tan band; name 12 bold black + crew tag 10 bold white.
      const nameRow = band([''], TAN, BLACK, true, 8);
      const cell = nameRow.getCell(1);
      cell.value = block.crewTag
        ? {
            richText: [
              { text: block.employeeName, font: { name: PP_FONT, size: 12, bold: true, color: { argb: BLACK } } },
              { text: ` (${block.crewTag})`, font: { name: PP_FONT, size: 10, bold: true, color: { argb: WHITE } } },
            ],
          }
        : block.employeeName;
      ws.mergeCells(nameRow.number, 1, nameRow.number, 8);
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }

    const headerRow = columnHeaderRow(bodyFill);
    const firstLineRow = ws.rowCount + 1;
    block.lines.forEach((l, li) => {
      const isLast = li === block.lines.length - 1;
      const r = ws.addRow([l.poText ?? (l.apKey ? `AP${l.apKey}` : ''), l.jobLabel, l.hours, l.rate]);
      const n = r.number;
      r.getCell(5).value = { formula: `C${n}*D${n}` };
      for (const c of [3, 4, 5]) {
        r.getCell(c).numFmt = HOURS_FMT;
        r.getCell(c).alignment = center;
        r.getCell(c).font = { ...baseFont };
        if (bodyFill) r.getCell(c).fill = solidFill(bodyFill);
      }
      // Job description text 10.5pt, left.
      r.getCell(2).font = { ...baseFont, size: 10.5 };
      if (bodyFill) r.getCell(2).fill = solidFill(bodyFill);
      // §5.5 — estimator fill on the PO# cell only. Office lines: 9pt white
      // date range on the accent blue.
      const po = r.getCell(1);
      if (l.poText !== null) {
        po.fill = solidFill(OFFICE_BLUE);
        po.font = { ...baseFont, size: 9, color: { argb: WHITE } };
      } else if (l.colorHex) {
        po.fill = solidFill('FF' + l.colorHex);
        po.font = { ...baseFont, color: { argb: textOn(l.colorHex) } };
      } else if (bodyFill) {
        po.fill = solidFill(bodyFill);
        po.font = { ...baseFont };
      }
      if (l.ot) r.getCell(3).fill = solidFill('FFFFFF00');
      // §5.5 — bold `(-) taxes` / `net` labels in F/G on the LAST job line.
      if (isLast) {
        r.getCell(6).value = '(-) taxes';
        r.getCell(7).value = 'net';
        for (const c of [6, 7]) {
          r.getCell(c).font = { ...baseFont, bold: true };
          r.getCell(c).alignment = center;
        }
      }
      if (isTs) {
        // Dotted inner separators + thin outer edges (§5.7).
        r.getCell(1).border = { left: { style: 'thin' }, bottom: { style: 'dotted' } };
        for (const c of [2, 3, 4]) r.getCell(c).border = { bottom: { style: 'dotted' } };
        r.getCell(5).border = { right: { style: 'thin' }, bottom: { style: 'dotted' } };
      }
    });
    if (isTs) {
      headerRow.getCell(1).border = { left: { style: 'thin' }, top: { style: 'thin' } };
      for (const c of [2, 3, 4]) headerRow.getCell(c).border = { top: { style: 'thin' } };
      headerRow.getCell(5).border = { right: { style: 'thin' }, top: { style: 'thin' } };
    }
    const lastLineRow = ws.rowCount;
    const sums = ws.addRow(['', '']);
    const n = sums.number;
    sums.getCell(3).value = { formula: `SUM(C${firstLineRow}:C${lastLineRow})` };
    sums.getCell(5).value = { formula: `SUM(E${firstLineRow}:E${lastLineRow})` };
    sums.getCell(6).value = block.net !== null ? { formula: `E${n}-G${n}` } : '';
    sums.getCell(7).value = block.net ?? '';
    for (const c of [3, 5, 6, 7]) {
      sums.getCell(c).numFmt = HOURS_FMT;
      sums.getCell(c).alignment = center;
    }
    sums.font = { ...baseFont, bold: true };
    // Thin top border across A–E on the sums row (§5.5); box bottom for §5.7.
    for (let c = 1; c <= 5; c++) {
      sums.getCell(c).border = {
        top: { style: 'thin' },
        ...(isTs
          ? {
              bottom: { style: 'thin' },
              ...(c === 1 ? { left: { style: 'thin' } } : {}),
              ...(c === 5 ? { right: { style: 'thin' } } : {}),
            }
          : {}),
      };
      if (bodyFill) sums.getCell(c).fill = solidFill(bodyFill);
    }
    ws.addRow([]);
  };

  const colLetter = (i: number): string => String.fromCharCode(65 + i); // 0→A

  const psBlock = (block: PsBlock): void => {
    const jobCount = block.jobs.length;
    // Round 5 §5 — crew-lead band fills A through H with DDD9C4 regardless of
    // how many job columns the block uses; WO% values still render on B+.
    const hdr = band(
      [block.displayName, ...block.jobs.map((j) => (j.woPct !== null ? j.woPct : '-'))],
      TAN,
      BLACK,
      true,
      8,
    );
    for (let i = 0; i < jobCount; i++) {
      const cell = hdr.getCell(2 + i);
      if (typeof cell.value === 'number') cell.numFmt = PCT_FMT;
      cell.alignment = { horizontal: 'right' };
    }
    const hdrRow = hdr.number;
    // §5.6 — PO/Name value cells right-aligned, bold, 10.5pt, estimator fill.
    const poRow = ws.addRow(['PO', ...block.jobs.map((j) => j.jobShort)]);
    const nameRow = ws.addRow(['Name', ...block.jobs.map((j) => j.customer)]);
    for (const row of [poRow, nameRow]) {
      row.getCell(1).font = { ...baseFont };
      block.jobs.forEach((j, i) => {
        const cell = row.getCell(2 + i);
        cell.alignment = { horizontal: 'right' };
        if (j.colorHex) {
          cell.fill = solidFill('FF' + j.colorHex);
          cell.font = { ...baseFont, size: 10.5, bold: true, color: { argb: textOn(j.colorHex) } };
        } else {
          cell.font = { ...baseFont, size: 10.5, bold: true };
        }
      });
    }

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
      // §5.6 formats: $-accounting on Profit Share / Labor / Mileage / ½Mat
      // (+ the info rows); plain accounting on WO Override / Total Profit Share.
      for (const rn of [psRow, laborRow, milesRow, halfRow, incomeRow, matRow]) {
        ws.getCell(`${L}${rn}`).numFmt = MONEY_FMT;
      }
      for (const rn of [woRow, tpsRow]) {
        ws.getCell(`${L}${rn}`).numFmt = ACCOUNTING_FMT;
      }
      // §5.6 borders: thin under -½ Materials; double under Deductions.
      ws.getCell(`${L}${halfRow}`).border = { bottom: { style: 'thin' } };
    }

    // Total: SUM across the Total Profit Share row (÷ split when present).
    const totalCell = ws.getCell(`B${totalRow}`);
    if (splitRow > 0) {
      totalCell.value = block.total; // splits make a single SUM wrong — write the value
    } else {
      totalCell.value = { formula: `SUM(B${tpsRow}:${colLetter(jobCount)}${tpsRow})` };
    }
    ws.getCell(`B${totalPayRow}`).value = { formula: `SUM(B${totalRow}:B${deductRow})` };
    for (const rn of [totalRow, mileNewRow, hourlyPayRow, reimbRow, deductRow]) {
      ws.getCell(`B${rn}`).numFmt = ACCOUNTING_FMT;
      ws.getCell(`A${rn}`).font = { ...baseFont, bold: rn === totalRow };
      ws.getCell(`B${rn}`).font = { ...baseFont, bold: rn === totalRow };
    }
    // §5.6 — Total Pay in $-accounting, Hourly Pay highlighted, colored
    // Reimbursements/Deductions values, double border above Total Pay.
    ws.getCell(`A${totalPayRow}`).font = { ...baseFont, bold: true };
    const totalPayCell = ws.getCell(`B${totalPayRow}`);
    totalPayCell.numFmt = MONEY_FMT;
    totalPayCell.font = { ...baseFont, bold: true };
    ws.getCell(`B${hourlyPayRow}`).fill = solidFill(HOURLY_PAY_YELLOW);
    ws.getCell(`B${reimbRow}`).font = { ...baseFont, color: { argb: REIMB_BLUE } };
    ws.getCell(`B${deductRow}`).font = { ...baseFont, color: { argb: DEDUCT_RED } };
    for (const col of ['A', 'B']) {
      ws.getCell(`${col}${deductRow}`).border = { bottom: { style: 'double' } };
    }
    ws.addRow([]);

    if (block.tsheets) {
      // §5.7 — boxed T-SHEETS table: merged A:E title band, no extra name row.
      const tsTitle = band(
        [`${block.tsheets.employeeName.split(' ')[0]}'s T-SHEETS BREAKDOWN`],
        NEAR_WHITE,
        BLACK,
        true,
        5,
      );
      ws.mergeCells(tsTitle.number, 1, tsTitle.number, 5);
      tsTitle.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      for (let c = 1; c <= 5; c++) {
        tsTitle.getCell(c).border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          ...(c === 1 ? { left: { style: 'thin' } } : {}),
          ...(c === 5 ? { right: { style: 'thin' } } : {}),
        };
      }
      hourlyBlock(block.tsheets, { tsheets: true, nameBand: false });
    }
  };

  const salaryBlock = (block: SalaryBlock, fill: string): void => {
    band([`${block.name}${block.timesheetHours > 0 ? ` (${block.timesheetHours}hrs)` : ''}`], fill, BLACK);
    // §5.4 — the date-range cell renders 9pt like the office PO cells.
    const r1 = ws.addRow([model.weekRange ? `(${model.weekRange})` : '', block.role, block.amount]);
    r1.getCell(1).font = { ...baseFont, size: 9 };
    r1.getCell(3).numFmt = ACCOUNTING_FMT;
    // Reimb already computed at the Settings mileage rate.
    const r2 = ws.addRow([`mileage reimb: ${block.name} - ${block.miles} miles`, '', block.mileageReimb]);
    r2.getCell(3).numFmt = ACCOUNTING_FMT;
    const r3 = ws.addRow(['Total', '', '']);
    r3.getCell(3).value = { formula: `C${r1.number}+C${r2.number}` };
    r3.getCell(3).numFmt = ACCOUNTING_FMT;
    r3.font = { ...baseFont, bold: true };
    ws.addRow([]);
  };

  // §5.2 — one alphabetical painter sequence (regular painters + lead blocks
  // interleaved; subcontractor blocks last), then the black section headers.
  for (const entry of model.painterSequence) {
    if (entry.kind === 'hourly') hourlyBlock(entry.block);
    else psBlock(entry.block);
  }

  band(['OFFICE'], BLACK, WHITE);
  for (const b of model.office) hourlyBlock(b);

  band(['ADMIN'], BLACK, WHITE);
  for (const b of model.admin) salaryBlock(b, TAN);
  band(['ESTIMATORS'], BLACK, WHITE);
  for (const b of model.estimators) salaryBlock(b, TAN);
  band(['MANAGER SALARY'], BLACK, WHITE);
  for (const b of model.managers) salaryBlock(b, MANAGER_FILL);

  // Base font on every otherwise-unstyled cell (Garamond everywhere, §5.1).
  ws.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (!cell.font || !cell.font.name) cell.font = { ...baseFont, ...(cell.font ?? {}) };
    });
  });
}

// ---------------------------------------------------------------------------
// Sheets 5–6 — Exceptions panels (Round 5 §6)
// ---------------------------------------------------------------------------

function buildExceptionsSheet(
  wb: ExcelJS.Workbook,
  sheetName: string,
  groups: readonly ExceptionGroup[],
  payDate: Date | null,
): void {
  const ws = wb.addWorksheet(sheetName, {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  ws.columns = [{ width: 22 }, { width: 70 }, { width: 24 }];
  const nonEmpty = groups.filter((g) => g.items.length > 0);
  const totalItems = nonEmpty.reduce((n, g) => n + g.items.length, 0);
  const title = `${sheetName} — ${payDate ? `${payDate.getMonth() + 1}/${payDate.getDate()}/${payDate.getFullYear()}` : ''} (${totalItems} items)`.replace(
    /\s+/g,
    ' ',
  ).trim();
  const titleRow = ws.addRow([title]);
  for (let c = 1; c <= 3; c++) {
    titleRow.getCell(c).fill = solidFill(BLACK);
    titleRow.getCell(c).font = { name: FONT, size: 14, bold: true, color: { argb: WHITE } };
  }
  ws.mergeCells(1, 1, 1, 3);

  if (nonEmpty.length === 0) {
    const row = ws.addRow(['No exceptions this week']);
    row.getCell(1).font = { name: FONT, size: 12, italic: true };
    return;
  }

  for (const g of nonEmpty) {
    const hdr = ws.addRow([`${g.title} (${g.items.length})`]);
    for (let c = 1; c <= 3; c++) {
      hdr.getCell(c).fill = solidFill(HEADER_GRAY);
      hdr.getCell(c).font = { name: FONT, size: 12, bold: true, color: { argb: WHITE } };
    }
    ws.mergeCells(hdr.number, 1, hdr.number, 3);
    for (const item of g.items) {
      // Best-effort split: leading "AP NNN — …" / "Name — …" → key | rest.
      const m = /^(.+?)\s+[—–-]\s+(.+)$/.exec(item.message);
      const key = m ? m[1] : '';
      const detail = m ? m[2] : item.message;
      const row = ws.addRow([key, detail, '']);
      for (let c = 1; c <= 3; c++) {
        row.getCell(c).font = { name: FONT, size: 11 };
        row.getCell(c).alignment = { wrapText: true, vertical: 'top' };
      }
    }
  }
}
