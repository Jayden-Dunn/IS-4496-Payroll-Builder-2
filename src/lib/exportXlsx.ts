import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { isDarkColor } from './format';
import type { LegendEntry, PivotModel, ReportRow } from './types';

const GREEN = 'FF1BA41B';
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

  // Header: raw TimeNavi columns keep the green; appended analysis columns are
  // black. Both white, NOT bold.
  const header = ws.addRow(TN_HEADERS);
  header.eachCell({ includeEmpty: false }, (cell, col) => {
    cell.fill = solidFill(col <= TN_RAW_HEADERS.length ? GREEN : BLACK);
    cell.font = { name: FONT, color: { argb: WHITE } };
    cell.alignment = { vertical: 'middle', wrapText: false };
  });

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

  /** Body row: full-height estimator column bands + gray Grand Total band. */
  const styleBodyRow = (row: ExcelJS.Row): void => {
    for (let c = 1; c <= 4; c++) paint(row.getCell(c), {});
    estCols.forEach((e, i) => {
      const hex = palette.get(e);
      const cell = row.getCell(5 + i);
      paint(cell, hex ? { fill: 'FF' + hex, color: textOn(hex) } : {});
      if (typeof cell.value === 'number') cell.numFmt = HOURS_FMT;
    });
    const gt = row.getCell(gtCol);
    paint(gt, { fill: GRAND_TOTAL_GRAY });
    if (typeof gt.value === 'number') gt.numFmt = HOURS_FMT;
  };

  for (const crew of pivot.crews) {
    // Crew-lead rows: full-width black band, overrides the column bands.
    blackBand(ws.addRow([crew.crew]));

    for (const emp of crew.employees) {
      styleBodyRow(ws.addRow([emp.username, emp.fname, emp.lname]));
      for (const job of emp.jobs) {
        styleBodyRow(ws.addRow(['', '', job.customer, job.job, ...hoursCells(job.byEstimator, job.total)]));
      }
      styleBodyRow(ws.addRow(['', `${emp.fname} Total`, '', '', ...hoursCells(emp.byEstimator, emp.total)]));
    }
  }

  // Grand Total row: gray/white on column A and every value column; B:D plain.
  const grandRow = ws.addRow(['Grand Total', '', '', '', ...hoursCells(pivot.grandByEstimator, pivot.grandTotal)]);
  paint(grandRow.getCell(1), { fill: HEADER_GRAY, color: WHITE });
  for (let c = 2; c <= 4; c++) paint(grandRow.getCell(c), {});
  for (let c = 5; c <= gtCol; c++) {
    const cell = grandRow.getCell(c);
    paint(cell, { fill: HEADER_GRAY, color: WHITE });
    if (typeof cell.value === 'number') cell.numFmt = HOURS_FMT;
  }
}
