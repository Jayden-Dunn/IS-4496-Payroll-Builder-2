import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { LegendEntry, PivotModel, ReportRow } from './types';

const GREEN = 'FF1BA41B';
const OVERTIME_YELLOW = 'FFFFFF99';
const PIVOT_HEADER_FILL = 'FFDCE6F1'; // light steel blue, approximates Excel pivot banding

// Reference formats: accounting style WITHOUT a $ symbol.
const ACCOUNTING_FMT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const START_END_DATE_FMT = 'mm-dd-yy';
const INVOICE_DATE_FMT = 'mm/dd/yyyy';
const PCT_FMT = '0%';
const HOURS_FMT = '0.00';

function solidFill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
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
  buildTimeNaviSheet(wb, input);
  buildPivotSheet(wb, input.pivot);
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

const TN_HEADERS = [
  'Event', 'Start Date', 'End Date', 'Calendar', 'Attendees', 'Location', 'Description',
  'AP#', '% Completion', 'Estimator', 'Last % Paid', 'QBO Job Name', 'In QBO Time?',
  'Job Total $', 'Apptivo $', '$ Difference', '%compl', 'Invoice Date', 'Date Difference',
];

// Reference column widths (A–S) so screenshots line up. G (Description) stays
// collapsed at 13 — the full text lives in the cell, wrap off.
const TN_WIDTHS = [60, 11, 13, 13, 9, 13, 13, 6.5, 11.6, 9, 10, 39, 20, 12, 14, 10.6, 9, 12.4, 13];

function buildTimeNaviSheet(wb: ExcelJS.Workbook, input: ExportInput): void {
  const ws = wb.addWorksheet('TimeNavi Export', {
    views: [{ state: 'frozen', ySplit: 1 }], // freeze at A2
  });
  ws.columns = TN_WIDTHS.map((width) => ({ width }));

  const header = ws.addRow(TN_HEADERS);
  header.eachCell({ includeEmpty: false }, (cell) => {
    cell.fill = solidFill(GREEN);
    // Reference header: white text, NOT bold.
    cell.font = { color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', wrapText: false };
  });

  const colorByEstimator = new Map(input.legend.map((l) => [l.estimator, 'FF' + l.color.toUpperCase()]));

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

    const argb = colorByEstimator.get(r.estimator);
    if (argb) {
      for (let c = 1; c <= TN_HEADERS.length; c++) {
        row.getCell(c).fill = solidFill(argb);
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Sheet 2 — Pivot by Crew (static replica of the Excel PivotTable)
// ---------------------------------------------------------------------------

function buildPivotSheet(wb: ExcelJS.Workbook, pivot: PivotModel): void {
  const ws = wb.addWorksheet('Pivot by Crew', {
    views: [{ state: 'frozen', ySplit: 4 }], // freeze at A5, below the header block
  });
  const estCols = pivot.estimators;
  // Grand Total column index is derived from however many estimator columns
  // rendered — never assumed.
  const totalColCount = 4 + estCols.length + 1;

  ws.columns = [
    { width: 35 }, { width: 14 }, { width: 24 }, { width: 39 },
    ...estCols.map(() => ({ width: 8 })),
    { width: 8 },
  ];

  const styleHeaderBlockRow = (row: ExcelJS.Row): void => {
    for (let c = 1; c <= totalColCount; c++) {
      const cell = row.getCell(c);
      cell.fill = solidFill(PIVOT_HEADER_FILL);
      cell.font = { ...(cell.font ?? {}), bold: true };
    }
  };

  // Row 1: title.
  const r1 = ws.addRow([pivot.title]);
  styleHeaderBlockRow(r1);
  r1.getCell(1).font = { bold: true, size: 12 };

  // Row 2: Name / Job labels + "Hours" centered across the value columns.
  const r2 = ws.addRow(['', '', 'Name', 'Job']);
  ws.mergeCells(2, 5, 2, totalColCount);
  r2.getCell(5).value = 'Hours';
  r2.getCell(5).alignment = { horizontal: 'center' };
  styleHeaderBlockRow(r2);

  // Row 3: pivot chrome labels.
  const r3 = ws.addRow(['Sum of hours', '', '', '', 'Column Labels']);
  styleHeaderBlockRow(r3);

  // Row 4: field header (the reference literally labels column D "jobcode_2").
  const r4 = ws.addRow(['Row Labels', 'fname', 'lname', 'jobcode_2', ...estCols, 'Grand Total']);
  styleHeaderBlockRow(r4);

  const hoursCells = (byEstimator: Record<string, number>, total: number): (number | '')[] => [
    ...estCols.map((e) => (byEstimator[e] !== undefined ? byEstimator[e] : '' as const)),
    total,
  ];

  const setHoursFmt = (row: ExcelJS.Row): void => {
    for (let c = 5; c <= totalColCount; c++) {
      const cell = row.getCell(c);
      if (typeof cell.value === 'number') cell.numFmt = HOURS_FMT;
    }
  };

  for (const crew of pivot.crews) {
    const crewRow = ws.addRow([crew.crew]);
    styleHeaderBlockRow(crewRow);

    for (const emp of crew.employees) {
      ws.addRow([emp.username, emp.fname, emp.lname]);

      for (const job of emp.jobs) {
        const jobRow = ws.addRow(['', '', job.customer, job.job, ...hoursCells(job.byEstimator, job.total)]);
        setHoursFmt(jobRow);
      }

      const totalRow = ws.addRow(['', `${emp.fname} Total`, '', '', ...hoursCells(emp.byEstimator, emp.total)]);
      setHoursFmt(totalRow);
      totalRow.eachCell({ includeEmpty: false }, (cell) => {
        cell.font = { ...(cell.font ?? {}), bold: true };
      });
      if (emp.overtime) {
        // Hard requirement: flag >40h weeks on the employee's Grand Total cell.
        totalRow.getCell(totalColCount).fill = solidFill(OVERTIME_YELLOW);
      }
    }
  }

  const grandRow = ws.addRow(['Grand Total', '', '', '', ...hoursCells(pivot.grandByEstimator, pivot.grandTotal)]);
  setHoursFmt(grandRow);
  grandRow.eachCell({ includeEmpty: false }, (cell) => {
    cell.font = { bold: true };
  });
}
