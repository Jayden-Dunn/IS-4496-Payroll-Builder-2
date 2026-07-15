import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { LegendEntry, PivotModel, ReportRow } from './types';

const GREEN = 'FF1BA41B';
const OVERTIME_YELLOW = 'FFFFFF99';

const MONEY_FMT = '$#,##0.00';
const DATE_FMT = 'm/d/yyyy';
const PCT_FMT = '0%';
const HOURS_FMT = '0.##';

function solidFill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function styleHeaderRow(row: ExcelJS.Row): void {
  row.eachCell({ includeEmpty: false }, (cell) => {
    cell.fill = solidFill(GREEN);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', wrapText: false };
  });
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

const TN_WIDTHS = [60, 12, 12, 22, 20, 24, 80, 10, 13, 11, 12, 42, 30, 14, 14, 14, 10, 12, 13];

function buildTimeNaviSheet(wb: ExcelJS.Workbook, input: ExportInput): void {
  const ws = wb.addWorksheet('TimeNavi Export', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  ws.columns = TN_WIDTHS.map((width) => ({ width }));
  styleHeaderRow(ws.addRow(TN_HEADERS));

  const colorByEstimator = new Map(input.legend.map((l) => [l.estimator, 'FF' + l.color.toUpperCase()]));

  for (const r of input.rows.filter((row) => row.included)) {
    const percentCell = r.percent === null ? '' : r.percent;
    const row = ws.addRow([
      r.event,
      r.startDate,
      r.endDate,
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
      r.jobTotal,
      r.apptivo,
      r.dollarDiff,
      percentCell, // Q mirrors I — the office PDF reads this column
      r.invoiceDate,
      r.dateDiff ?? '',
    ]);

    for (const [col, fmt] of [[2, DATE_FMT], [3, DATE_FMT], [18, DATE_FMT]] as const) {
      row.getCell(col).numFmt = fmt;
    }
    for (const col of [9, 11, 17]) {
      const cell = row.getCell(col);
      if (typeof cell.value === 'number') cell.numFmt = PCT_FMT;
    }
    for (const col of [14, 15, 16]) row.getCell(col).numFmt = MONEY_FMT;
    row.getCell(19).numFmt = '0';

    row.getCell(7).alignment = { wrapText: true, vertical: 'top' };
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
    views: [{ state: 'frozen', ySplit: 2 }],
  });
  const estCols = pivot.estimators;
  const totalColCount = 4 + estCols.length + 1;

  ws.columns = [
    { width: 34 }, { width: 14 }, { width: 26 }, { width: 44 },
    ...estCols.map(() => ({ width: 10 })),
    { width: 12 },
  ];

  const title = ws.addRow([pivot.title]);
  title.getCell(1).font = { bold: true, size: 12 };

  styleHeaderRow(ws.addRow(['Row Labels', 'fname', 'lname', 'Job', ...estCols, 'Grand Total']));

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
    crewRow.getCell(1).font = { bold: true };

    for (const emp of crew.employees) {
      ws.addRow([emp.username, emp.fname, emp.lname]);

      for (const job of emp.jobs) {
        const jobRow = ws.addRow(['', '', job.customer, job.job, ...hoursCells(job.byEstimator, job.total)]);
        setHoursFmt(jobRow);
      }

      const totalRow = ws.addRow(['', `${emp.fname} Total`, '', '', ...hoursCells(emp.byEstimator, emp.total)]);
      totalRow.getCell(2).font = { bold: true };
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
