/**
 * Round 6 — new workbook sheets (mirror/%'s + payroll derived).
 * Round 8 §10 — Profitshare Output Hourly Pay / Deductions/Sundries are live
 * `'Input #''s'!` formulas (escaped apostrophe), not baked values.
 */
import type { CellValue, Fill, Workbook, Worksheet } from 'exceljs';
import { extractApKey } from './apkey';
import { formatCrewLeadWoNote, isDarkColor } from './format';
import type { JobsByCrewModel } from './jobsByCrew';
import {
  pxToExcelWidth,
  SALES_BY_REP_ACCOUNTING_COLS,
  SALES_BY_REP_WIDTHS_PX,
  TIMESHEET_COLLAPSED_HEADERS,
  type SalesByCustMirrorRow,
  type SalesByRepMirrorRow,
  type TimesheetMirrorRow,
} from './mirrorTabs';
import { hourSplitRows, profitshareOutputLeads } from './payroll';
import { buildPrFinalCheckRows } from './prFinalCheck';
import type {
  HourlyBlock,
  InactiveEmployeeBlock,
  PayrollPreviewModel,
  PlParseResult,
  PsBlock,
  SalaryBlock,
} from './types';

const BLACK = 'FF000000';
const WHITE = 'FFFFFFFF';
const HEADER_GRAY = 'FF808080';
const TAN = 'FFDDD9C4';
const HOURLY_PAY_YELLOW = 'FFFFFF99';
const TOTAL_PAY_CREAM = 'FFFAF5F0';
const NAME_BAND_RED = 'FFC0504D';
const NAME_BAND_GREEN = 'FF00B050';
const REIMB_BLUE = 'FF0070C0';
const DEDUCT_RED = 'FFFF0000';
const FONT = 'Calibri';
const PP_FONT = 'Garamond';
// Round 10 — negatives read red and parenthesised in the workbook too.
const ACCOUNTING_FMT = '_(* #,##0.00_);[Red]_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
// Round 10 — negatives read red and parenthesised in the workbook too.
const MONEY_FMT = '_("$"* #,##0.00_);[Red]_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
const PCT_FMT = '0%';
const HOURS_FMT = '0.00';
const DATE_FMT = 'm/d/yyyy';

function solidFill(argb: string): Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function textOn(hex6: string): string {
  return isDarkColor(hex6) ? WHITE : BLACK;
}

function excelDate(d: Date): Date {
  return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
}

function cellValue(v: string | number | boolean | Date | null | undefined): CellValue {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return excelDate(v);
  return v;
}

function boldHeader(ws: Worksheet, headers: string[], fontName = FONT): void {
  const row = ws.addRow(headers);
  for (let c = 1; c <= headers.length; c++) {
    const cell = row.getCell(c);
    cell.font = { name: fontName, bold: true, size: 11 };
  }
}

const TIMESHEET_HEADERS = [
  'username',
  'payroll_id',
  'fname',
  'lname',
  'number',
  'group',
  'local_date',
  'local_day',
  'local_start_time',
  'local_end_time',
  'tz',
  'hours',
  'jobcode_1',
  'jobcode_2',
  'billable',
  'class',
  'department',
  'payroll_item',
  'service_item',
  'location',
  'notes',
  'approved_status',
  'has_flags',
  'flag_types',
  'Apname',
  'AP#',
  'Estimator',
  'QBO Job Name',
  'Not on Calendar',
];

// ---------------------------------------------------------------------------
// %'s mirrors
// ---------------------------------------------------------------------------

const MIRROR_ARIAL = 'Arial';
const WIDE_TEXT_HEADERS = /^(customer|name(\s*[12])?|memo(\/description)?|memo)$/i;
const PRODUCT_SERVICE_HEADER = /^product\s*\/\s*service$/i;

/**
 * Round 10 — Sales by Rep (NV and UT) column widths, roughly halved so columns
 * A–M all fit on screen without scrolling right. The long free-text columns
 * are capped rather than merely halved: a 105-wide Memo column alone would
 * push column M off the edge. Nothing is truncated — the text is still in the
 * cell, it just no longer sets the column width.
 */
function salesByRepColWidth(header: string, colIndex0: number): number {
  if (colIndex0 === 0) return 7; // AP#
  if (WIDE_TEXT_HEADERS.test(header.trim())) return 16;
  if (PRODUCT_SERVICE_HEADER.test(header.trim())) return 11;
  // A date needs its full 10 characters — a clipped date cell shows as ####.
  if (/^date$/i.test(header.trim())) return 10;
  return 6;
}

export function buildSalesByRepMirrorSheet(
  wb: Workbook,
  name: string,
  data: { headers: string[]; rows: SalesByRepMirrorRow[] } | null | undefined,
): void {
  const isUt = /^ut\s*salesbyrep$/i.test(name);
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
  const headers = data?.headers?.length
    ? ['AP#', ...data.headers.slice(0, 12), 'Date', 'QBO Job Name']
    : ['AP#', 'Date', 'Customer', 'Name 1', 'Name 2', 'Transaction Type', 'Num', 'Product/Service', 'Memo/Description', 'Sales Price', 'Amount', 'Balance', 'Sales Rep', 'Date', 'QBO Job Name'];
  // Round 11 — A–M at the specified pixel widths; the trailing helper
  // columns keep their content-based widths.
  ws.columns = headers.map((h, i) => ({
    width: i < SALES_BY_REP_WIDTHS_PX.length ? pxToExcelWidth(SALES_BY_REP_WIDTHS_PX[i]) : salesByRepColWidth(h, i),
  }));
  const headerRow = ws.addRow(headers);
  for (let c = 1; c <= headers.length; c++) {
    headerRow.getCell(c).font = { name: MIRROR_ARIAL, bold: true, size: isUt ? 9 : 8 };
  }
  for (const r of data?.rows ?? []) {
    const cols = r.cols.slice(0, 12);
    while (cols.length < 12) cols.push(null);
    const row = ws.addRow([r.apKey, ...cols.map(cellValue), cellValue(r.dateCopy), cellValue(r.qboJobName)]);
    for (let c = 1; c <= headers.length; c++) {
      row.getCell(c).font = { name: MIRROR_ARIAL, size: 8 };
      const v = row.getCell(c).value;
      if (v instanceof Date) row.getCell(c).numFmt = DATE_FMT;
      // Round 11 — J, K, L show thousands separators: Excel's Accounting format.
      else if (typeof v === 'number' && SALES_BY_REP_ACCOUNTING_COLS.has(c - 1)) row.getCell(c).numFmt = ACCOUNTING_FMT;
      else if (typeof v === 'number') row.getCell(c).numFmt = HOURS_FMT;
    }
  }
}

function looksLikeCustTitleRow(text: string, index: number): boolean {
  const t = text.trim();
  if (!t) return false;
  if (index === 0 && /allpro/i.test(t)) return true;
  if (/sales\s+by\s+customer/i.test(t)) return true;
  // Date-range title line (e.g. "June 1 - June 7, 2026")
  if (index <= 2 && /\d/.test(t) && (/\s-\s/.test(t) || /\b20\d{2}\b/.test(t))) return true;
  return false;
}

export function buildSalesByCustMirrorSheet(
  wb: Workbook,
  name: string,
  rows: SalesByCustMirrorRow[] | null | undefined,
): void {
  const isUt = /^ut\s*salesbycust$/i.test(name);
  const list = rows ?? [];
  const ws = wb.addWorksheet(name, {
    views: [{ state: 'frozen', ySplit: isUt ? 5 : 1 }],
  });
  ws.columns = [{ width: 48 }, { width: 14 }, { width: 10 }, { width: 12 }];

  let dataRows = list;
  let titleCount = 0;
  while (titleCount < 3 && titleCount < list.length && looksLikeCustTitleRow(list[titleCount].text, titleCount)) {
    titleCount++;
  }

  if (isUt && titleCount > 0) {
    // Title block A1–A3; money header "Total" on B5 (not B1).
    for (let t = 0; t < titleCount; t++) {
      const size = t === 2 ? 10 : 14;
      const row = ws.addRow([list[t].text, '']);
      for (let c = 1; c <= 2; c++) {
        row.getCell(c).font = { name: MIRROR_ARIAL, bold: true, size };
        row.getCell(c).alignment = { horizontal: 'center', vertical: 'middle' };
      }
      ws.mergeCells(row.number, 1, row.number, 2);
    }
    // Pad blank rows so the Total header lands on row 5; B1 stays clear of "Total".
    while (ws.rowCount < 4) ws.addRow([]);
    const hdr = ws.addRow(['', 'Total', 'AP#', 'Job total']);
    for (let c = 1; c <= 4; c++) {
      hdr.getCell(c).font = { name: MIRROR_ARIAL, bold: true, size: 8 };
    }
    dataRows = list.slice(titleCount);
  } else {
    // NV: synthetic header; skip title rows so first customer is on row 2.
    const hdr = ws.addRow(['', 'Total', 'AP#', 'Job total']);
    for (let c = 1; c <= 4; c++) {
      hdr.getCell(c).font = { name: MIRROR_ARIAL, bold: true, size: 8 };
    }
    dataRows = list.slice(titleCount);
  }

  // Drop a leading blank data row after the header (row 2 must be first customer).
  while (dataRows.length > 0 && !dataRows[0].text.trim() && dataRows[0].value === null && !dataRows[0].apKey) {
    dataRows = dataRows.slice(1);
  }

  for (const r of dataRows) {
    const row = ws.addRow([r.text, r.value ?? '', r.apKey, r.jobTotal ?? '']);
    for (let c = 1; c <= 4; c++) {
      row.getCell(c).font = { name: MIRROR_ARIAL, size: 8, bold: c === 1 };
    }
    const aText = String(r.text ?? '');
    if (/^total/i.test(aText.trim())) {
      const thinTop = { top: { style: 'thin' as const } };
      row.getCell(1).border = { ...row.getCell(1).border, ...thinTop };
      row.getCell(2).border = { ...row.getCell(2).border, ...thinTop };
      row.getCell(2).font = { name: MIRROR_ARIAL, size: 8, bold: true };
      if (typeof row.getCell(2).value === 'number') {
        row.getCell(2).numFmt = '"$"* #,##0.00';
      }
    } else if (typeof row.getCell(2).value === 'number') {
      row.getCell(2).numFmt = HOURS_FMT;
    }
    if (typeof row.getCell(4).value === 'number') row.getCell(4).numFmt = HOURS_FMT;
  }
}

export function buildTimesheetMirrorSheet(
  wb: Workbook,
  rows: TimesheetMirrorRow[] | null | undefined,
): void {
  const ws = wb.addWorksheet('timesheets', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = TIMESHEET_HEADERS.map((h) => ({
    width: h === 'jobcode_2' || h === 'QBO Job Name' ? 36 : h === 'hours' ? 8 : 12,
  }));
  boldHeader(ws, TIMESHEET_HEADERS);
  // Round 11 — collapsed: hidden in Excel, one right-click "Unhide" away.
  TIMESHEET_HEADERS.forEach((h, i) => {
    if (TIMESHEET_COLLAPSED_HEADERS.has(h)) ws.getColumn(i + 1).hidden = true;
  });
  for (const r of rows ?? []) {
    const raw = r.raw.slice(0, 24);
    while (raw.length < 24) raw.push(null);
    const vals = [
      ...raw.map(cellValue),
      r.apname,
      r.apKey,
      r.estimator,
      r.qboJobName,
      r.notOnCalendar,
    ];
    const row = ws.addRow(vals);
    for (let c = 1; c <= TIMESHEET_HEADERS.length; c++) {
      row.getCell(c).font = { name: FONT, size: 11 };
      const v = row.getCell(c).value;
      if (v instanceof Date) row.getCell(c).numFmt = DATE_FMT;
      if (c === 12 && typeof v === 'number') row.getCell(c).numFmt = HOURS_FMT;
    }
  }
}

export function buildJobsByCrewSheet(
  wb: Workbook,
  model: JobsByCrewModel | null | undefined,
): void {
  const ws = wb.addWorksheet('Jobs by Crew');
  // Round 7 §9 — reference column widths (A spacer + B–E data).
  ws.columns = [{ width: 3 }, { width: 25.2 }, { width: 26.2 }, { width: 31.2 }, { width: 12.2 }];
  const bodyFont = { name: FONT, size: 12 };
  const OFFICE_YELLOW = 'FFFFFF99';
  const GRAND_LABEL_GRAY = 'FFBFBFBF';

  // Anchored at B2 per §3.6 — plain Calibri 12, not bold (Round 7 §9).
  ws.getCell('B2').value = 'Row Labels';
  ws.getCell('C2').value = 'jobcode_1';
  ws.getCell('D2').value = 'jobcode_2';
  ws.getCell('E2').value = 'Sum of hours';
  for (const col of ['B', 'C', 'D', 'E']) {
    ws.getCell(`${col}2`).font = { ...bodyFont };
  }

  let excelRow = 3;
  for (const g of model?.groups ?? []) {
    g.jobs.forEach((j, i) => {
      ws.getCell(`B${excelRow}`).value = i === 0 ? g.crew : '';
      ws.getCell(`C${excelRow}`).value = j.showCustomer ? j.customer : '';
      ws.getCell(`D${excelRow}`).value = j.job;
      ws.getCell(`E${excelRow}`).value = j.hours;
      ws.getCell(`E${excelRow}`).numFmt = HOURS_FMT;
      for (const col of ['B', 'C', 'D', 'E']) {
        ws.getCell(`${col}${excelRow}`).font = { ...bodyFont };
      }
      // ALLPRO OFFICE jobcode_1 cells: pale yellow + thin top/bottom border.
      const custCell = ws.getCell(`C${excelRow}`);
      if (typeof custCell.value === 'string' && custCell.value.trim().toLowerCase() === 'allpro office') {
        custCell.fill = solidFill(OFFICE_YELLOW);
        custCell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
        };
      }
      excelRow++;
    });
  }

  ws.getCell(`B${excelRow}`).value = 'Grand Total';
  ws.getCell(`E${excelRow}`).value = model?.grandTotal ?? 0;
  ws.getCell(`E${excelRow}`).numFmt = HOURS_FMT;
  ws.getCell(`B${excelRow}`).font = { ...bodyFont };
  ws.getCell(`E${excelRow}`).font = { ...bodyFont };
  // Grand Total: double top on label + hours; label grey fill.
  ws.getCell(`B${excelRow}`).fill = solidFill(GRAND_LABEL_GRAY);
  ws.getCell(`B${excelRow}`).border = { top: { style: 'double' } };
  ws.getCell(`E${excelRow}`).border = { top: { style: 'double' } };
}

// ---------------------------------------------------------------------------
// Payroll derived sheets (Garamond; values for core)
// ---------------------------------------------------------------------------

function applyGaramond(ws: Worksheet, size = 12): void {
  ws.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (!cell.font?.name) cell.font = { name: PP_FONT, size, ...(cell.font ?? {}) };
      else if (cell.font.name !== PP_FONT) cell.font = { ...cell.font, name: PP_FONT };
    });
  });
}

export function buildPlImportSheet(
  wb: Workbook,
  pl: PlParseResult | null,
  // Round 8 §7.3 — the UT tree renders identically onto its own sheet.
  sheetName = 'P&L by Job Import',
): void {
  const ws = wb.addWorksheet(sheetName);

  // Round 7 — full uploaded account tree when present; else fall back to byJob summary.
  if (pl && pl.columnHeaders.length > 0 && pl.accountRows.length > 0) {
    const headers = [...pl.columnHeaders];
    if (!headers[0]) headers[0] = 'Account';
    ws.columns = headers.map((_, i) => ({ width: i === 0 ? 28 : 14 }));
    const hdr = ws.addRow(headers);
    for (let c = 1; c <= headers.length; c++) {
      hdr.getCell(c).fill = solidFill(HEADER_GRAY);
      hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
    }
    for (const ar of pl.accountRows) {
      const cells: (string | number | null)[] = [ar.label];
      for (let c = 1; c < headers.length; c++) {
        cells.push(ar.values[c] ?? null);
      }
      const row = ws.addRow(cells);
      row.getCell(1).font = { name: PP_FONT, size: 12 };
      for (let c = 2; c <= headers.length; c++) {
        const cell = row.getCell(c);
        if (typeof cell.value === 'number') cell.numFmt = ACCOUNTING_FMT;
        cell.font = { name: PP_FONT, size: 12 };
      }
    }
    applyGaramond(ws);
    return;
  }

  const jobs = [...(pl?.byJob.entries() ?? [])];
  const headers = ['Account', ...jobs.map(([ap, m]) => m.plJobHeader || `AP${ap}`)];
  ws.columns = headers.map((_, i) => ({ width: i === 0 ? 28 : 14 }));
  const hdr = ws.addRow(headers);
  for (let c = 1; c <= headers.length; c++) {
    hdr.getCell(c).fill = solidFill(HEADER_GRAY);
    hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
  }
  const metricRows: { label: string; pick: (m: (typeof jobs)[0][1]) => number }[] = [
    { label: 'Total Income', pick: (m) => m.income },
    { label: 'Materials', pick: (m) => m.materials },
    { label: 'QB Labor-Wages-Subs', pick: (m) => m.laborWagesSubs },
    { label: 'Mileage Expense', pick: (m) => m.mileage },
    { label: 'Sundries -', pick: (m) => m.sundries },
  ];
  for (const mr of metricRows) {
    const row = ws.addRow([mr.label, ...jobs.map(([, m]) => mr.pick(m))]);
    row.getCell(1).font = { name: PP_FONT, size: 12, bold: true };
    for (let c = 2; c <= headers.length; c++) {
      row.getCell(c).numFmt = ACCOUNTING_FMT;
      row.getCell(c).font = { name: PP_FONT, size: 12 };
    }
  }
  applyGaramond(ws);
}

type PsChainStyle = 'default' | 'profitshare' | 'management';

/** 1-based column number → Excel letters ("B", "AA"). */
function colLetters(col: number): string {
  let n = Math.max(1, Math.trunc(col));
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export type InputNumsLeadRefs = {
  hourlyCell: string;
  deductCell: string;
  /** Round 10 — Reimbursements amount + its note cell. */
  reimbCell: string;
  reimbNoteCell: string;
  /** Round 10 — the Deductions/Sundries note cell. */
  deductNoteCell: string;
  /** Round 10 — the Miles: entry cell (× the mileage rate = Mileage - New). */
  milesCell: string;
};

/** Round 8 §10 — `'Input #''s'!C10` (escaped apostrophe in the sheet name). */
export function inputNumsSheetFormula(cell: string): { formula: string } {
  return { formula: `'Input #''s'!${cell}` };
}

/**
 * Round 8 §10.1/§10.2 — row map of each lead's Input #'s block. Must stay
 * byte-identical to `buildInputNumbersSheet` so Profitshare formulas resolve.
 * Layout from row `start` (name row):
 * +0 name, +1 PO, +2 WO, +3 %, +4 materials hdr, +5 blank,
 * +6 Reimbursements, +7 Hourly Pay, +8 Deductions/Sundries, +9 Miles, +10 spacer.
 */
export function inputNumbersCellRefs(leadCount: number): InputNumsLeadRefs[] {
  const out: InputNumsLeadRefs[] = [];
  let row = 3;
  for (let i = 0; i < leadCount; i++) {
    out.push({
      hourlyCell: `C${row + 7}`,
      deductCell: `C${row + 8}`,
      reimbCell: `C${row + 6}`,
      reimbNoteCell: `D${row + 6}`,
      deductNoteCell: `D${row + 8}`,
      milesCell: `C${row + 9}`,
    });
    row += 11;
  }
  return out;
}

function writePsChainValues(
  ws: Worksheet,
  block: PsBlock,
  startCol: number,
  opts: {
    skipWoBand?: boolean;
    defaultPsPct: number;
    /** Absolute start row; omit to append after current sheet content. */
    startRow?: number;
    style?: PsChainStyle;
    /** FINAL — emit Total Gross Wages (= PS Total + T-SHEETS) after Total Pay. */
    showGrossWages?: boolean;
    /**
     * Round 9 §30 — FINAL only: drop the reference rows below Total Pay
     * (Total Income, Materials, Travel, QB Labor-Wages-Subs, Mileage in
     * Books, % Materials). They stay on the main Payroll Preview.
     */
    hideBooksRows?: boolean;
    /** Round 8 §10 — live `'Input #''s'!` refs for Hourly Pay / Deductions. */
    hourlyPayRef?: string;
    deductionsRef?: string;
    /** Round 10 — live refs for Reimbursements, the two notes, and Miles. */
    reimbRef?: string;
    reimbNoteRef?: string;
    deductNoteRef?: string;
    milesRef?: string;
    /** Round 10 — $/mile applied to the Miles cell for `Mileage - New`. */
    mileageRate?: number;
  },
): number {
  const jobCount = Math.min(block.jobs.length, 6);
  const baseFont = { name: PP_FONT, size: 12 };
  const style = opts.style ?? 'default';
  let r = opts.startRow ?? ws.rowCount + 1;

  const set = (row: number, col: number, v: CellValue, fmt?: string): void => {
    const cell = ws.getCell(row, col);
    cell.value = v;
    cell.font = { ...baseFont };
    if (fmt) cell.numFmt = fmt;
  };

  const nameFill =
    style === 'profitshare' ? NAME_BAND_RED : style === 'management' ? NAME_BAND_GREEN : TAN;
  const nameFg = style === 'profitshare' || style === 'management' ? WHITE : BLACK;

  // Name / WO% band
  set(r, startCol, block.displayName);
  ws.getCell(r, startCol).fill = solidFill(nameFill);
  ws.getCell(r, startCol).font = { ...baseFont, bold: true, color: { argb: nameFg } };
  if (!opts.skipWoBand) {
    for (let i = 0; i < jobCount; i++) {
      const pct = block.jobs[i].woPct;
      set(r, startCol + 1 + i, pct !== null ? pct : '-', pct !== null ? PCT_FMT : undefined);
      ws.getCell(r, startCol + 1 + i).fill = solidFill(nameFill);
      ws.getCell(r, startCol + 1 + i).font = { ...baseFont, color: { argb: nameFg } };
    }
  } else {
    // FINAL — keep name cell only (no WO% band across job columns).
    for (let i = 0; i < jobCount; i++) {
      ws.getCell(r, startCol + 1 + i).fill = solidFill(nameFill);
    }
  }
  r++;

  set(r, startCol, 'PO');
  for (let i = 0; i < jobCount; i++) {
    const j = block.jobs[i];
    set(r, startCol + 1 + i, j.jobShort);
    if (j.colorHex) {
      ws.getCell(r, startCol + 1 + i).fill = solidFill('FF' + j.colorHex);
      ws.getCell(r, startCol + 1 + i).font = {
        ...baseFont,
        size: 10.5,
        bold: true,
        color: { argb: textOn(j.colorHex) },
      };
    }
  }
  r++;

  set(r, startCol, 'Name');
  for (let i = 0; i < jobCount; i++) {
    const j = block.jobs[i];
    set(r, startCol + 1 + i, j.customer);
    if (j.colorHex) {
      ws.getCell(r, startCol + 1 + i).fill = solidFill('FF' + j.colorHex);
      ws.getCell(r, startCol + 1 + i).font = {
        ...baseFont,
        size: 10.5,
        bold: true,
        color: { argb: textOn(j.colorHex) },
      };
    }
  }
  r++;

  // Round 8 §8.1 — SUBBED / SET RATE bands carry no auto profit share: the
  // computed cells export blank and the % cell shows the state verbatim.
  const cut = block.jobs.slice(0, jobCount);
  const rows: { label: string; vals: (number | string)[]; fmt?: string; bold?: boolean }[] = [
    { label: 'Profit Share', vals: cut.map((j) => (j.manualPay ? '' : j.profitShare)), fmt: MONEY_FMT },
    {
      label: 'Work Order Override',
      vals: cut.map((j) => (j.workOrderOverride !== null ? j.workOrderOverride : '')),
      fmt: ACCOUNTING_FMT,
    },
    {
      label: '% Completion',
      vals: cut.map((j) => (j.manualPay ? j.manualPay : j.percentCompletion)),
      fmt: PCT_FMT,
    },
    { label: '- Labor + Wages', vals: cut.map((j) => j.laborWages), fmt: MONEY_FMT },
    { label: '-Mileage', vals: cut.map((j) => j.mileage), fmt: MONEY_FMT },
    {
      // Round 9 §32 — a full-materials lead is charged all of it.
      label: block.jobs.some((j) => j.fullMaterials) ? '- ALL Materials' : '-½ Materials',
      vals: cut.map((j) => j.halfMaterials),
      fmt: MONEY_FMT,
    },
    // Round 10 §UT — the hotel charge-back, when the band carries one.
    ...(cut.some((j) => j.hotels !== null)
      ? [{ label: '- Hotels', vals: cut.map((j) => j.hotels ?? ''), fmt: MONEY_FMT }]
      : []),
    {
      label: 'Total Profit Share',
      vals: cut.map((j) => (j.manualPay ? '' : j.totalProfitShare)),
      fmt: ACCOUNTING_FMT,
      bold: true,
    },
  ];
  const materialsLabel = block.jobs.some((j) => j.fullMaterials) ? '- ALL Materials' : '-½ Materials';
  let halfMatRow = 0;
  let mileageRow = 0;
  for (const rr of rows) {
    // Round 10 — the rule sits under the LAST deduction row, which is Hotels
    // when the band has one.
    if (rr.label === materialsLabel || rr.label === '- Hotels') halfMatRow = r;
    if (rr.label === '-Mileage') mileageRow = r;
    set(r, startCol, rr.label);
    if (rr.bold) ws.getCell(r, startCol).font = { ...baseFont, bold: true };
    for (let i = 0; i < jobCount; i++) {
      set(r, startCol + 1 + i, rr.vals[i], rr.fmt);
      if (rr.bold) ws.getCell(r, startCol + 1 + i).font = { ...baseFont, bold: true };
      // FINAL — `@ NN% PS` beside each TPS amount when job WO% ≠ default.
      if (opts.skipWoBand && rr.label === 'Total Profit Share') {
        const j = block.jobs[i];
        const def = opts.defaultPsPct;
        if (j.woPct !== null && Math.abs(j.woPct - def) > 1e-9) {
          set(r, startCol + 1 + jobCount + i, `@ ${Math.round(j.woPct * 100)}% PS`);
          ws.getCell(r, startCol + 1 + jobCount + i).font = { ...baseFont, size: 10 };
        }
      }
    }
    r++;
  }
  if (halfMatRow > 0) {
    for (let c = 0; c < 5; c++) {
      ws.getCell(halfMatRow, startCol + c).border = { bottom: { style: 'thin' } };
    }
  }

  if (block.jobs.some((j) => j.splitWith > 1)) {
    set(r, startCol, 'split w/');
    for (let i = 0; i < jobCount; i++) {
      const sw = block.jobs[i].splitWith;
      set(r, startCol + 1 + i, sw > 1 ? sw : '');
    }
    r++;
  }

  r++; // blank
  const footer: {
    label: string;
    value: number | string;
    memo?: string;
    highlight?: boolean;
    cream?: boolean;
    color?: string;
    bold?: boolean;
    money?: boolean;
    formula?: string;
    /** Round 10 — a raw formula (no `Input #'s` prefixing) for this cell. */
    rawFormula?: string;
    /** Round 10 — the note cell mirrors an Input #'s note. */
    memoFormula?: string;
  }[] = [
    { label: 'Total', value: block.total, bold: true },
    {
      // Round 10 — Miles on the Input #'s tab × the mileage rate, as a live
      // formula, so typing miles there fills this in.
      label: 'Mileage - New',
      value: block.mileageNew,
      ...(opts.milesRef
        ? { rawFormula: `'Input #''s'!${opts.milesRef}*${(opts.mileageRate ?? 0.625).toFixed(3)}` }
        : {}),
    },
    { label: 'Hourly Pay', value: block.hourlyPay, highlight: true, formula: opts.hourlyPayRef },
    {
      label: 'Reimbursements/additional pay',
      value: block.reimbursements,
      memo: block.reimbMemo,
      color: REIMB_BLUE,
      // Round 10 — amount AND note come from the Input #'s tab.
      formula: opts.reimbRef,
      memoFormula: opts.reimbNoteRef,
    },
    {
      label: 'Deductions/Sundries',
      value: block.deductions,
      memo: block.deductMemo,
      color: DEDUCT_RED,
      formula: opts.deductionsRef,
      memoFormula: opts.deductNoteRef,
    },
    // Round 9 §19/§23 — configured extra rows, directly above Total Pay.
    ...block.extras.map((x) => ({
      label: x.label,
      value: x.amount ?? '',
      color: x.kind === 'deduction' ? DEDUCT_RED : x.kind === 'reimbursement' ? REIMB_BLUE : undefined,
    })),
    { label: 'Total Pay', value: block.totalPay, bold: true, money: true, cream: style === 'profitshare' },
  ];
  let deductFooterRow = 0;
  for (const f of footer) {
    if (f.label === 'Deductions/Sundries') deductFooterRow = r;
    set(r, startCol, f.label);
    if (f.formula || f.rawFormula) {
      ws.getCell(r, startCol + 1).value = f.rawFormula
        ? { formula: f.rawFormula }
        : inputNumsSheetFormula(f.formula as string);
      ws.getCell(r, startCol + 1).numFmt = f.money ? MONEY_FMT : ACCOUNTING_FMT;
      ws.getCell(r, startCol + 1).font = { ...baseFont };
    } else {
      set(r, startCol + 1, f.value, f.money ? MONEY_FMT : ACCOUNTING_FMT);
    }
    // Round 10 — the note beside the amount mirrors the Input #'s note.
    if (f.memoFormula) {
      const cell = ws.getCell(r, startCol + 2);
      cell.value = inputNumsSheetFormula(f.memoFormula);
      cell.font = { ...baseFont };
    } else if (f.memo) {
      set(r, startCol + 2, f.memo);
    }
    if (f.bold) {
      ws.getCell(r, startCol).font = { ...baseFont, bold: true };
      ws.getCell(r, startCol + 1).font = { ...baseFont, bold: true };
    }
    if (f.highlight) ws.getCell(r, startCol + 1).fill = solidFill(HOURLY_PAY_YELLOW);
    if (f.cream) {
      ws.getCell(r, startCol).fill = solidFill(TOTAL_PAY_CREAM);
      ws.getCell(r, startCol + 1).fill = solidFill(TOTAL_PAY_CREAM);
    }
    if (f.color) ws.getCell(r, startCol + 1).font = { ...baseFont, color: { argb: f.color } };
    r++;
  }
  if (deductFooterRow > 0) {
    for (let c = 0; c < 5; c++) {
      ws.getCell(deductFooterRow, startCol + c).border = { bottom: { style: 'double' } };
    }
  }
  if (opts.showGrossWages) {
    // Round 9 §31 — hours pay + profit share for this crew lead.
    const gross = block.totalGrossWages;
    set(r, startCol, 'Total Gross Wages:');
    set(r, startCol + 1, gross, MONEY_FMT);
    ws.getCell(r, startCol).font = { ...baseFont, bold: true };
    ws.getCell(r, startCol + 1).font = { ...baseFont, bold: true };
    r++;
  }

  // Round 9 §30 — the Final Payroll Breakdown stops here.
  if (opts.hideBooksRows) return r;

  r++; // blank before income block
  const info: { label: string; vals: number[]; fmt: string }[] = [
    { label: 'Total Income', vals: block.jobs.slice(0, jobCount).map((j) => j.income), fmt: MONEY_FMT },
    { label: 'Materials', vals: block.jobs.slice(0, jobCount).map((j) => j.materials), fmt: MONEY_FMT },
    // Round 8 §8.5 — informational, directly beneath Materials.
    { label: 'Travel', vals: block.jobs.slice(0, jobCount).map((j) => j.travel), fmt: MONEY_FMT },
    { label: 'QB Labor-Wages-Subs', vals: block.jobs.slice(0, jobCount).map((j) => j.laborWages), fmt: MONEY_FMT },
    { label: 'Mileage in Books', vals: block.jobs.slice(0, jobCount).map((j) => j.mileage), fmt: MONEY_FMT },
    {
      label: '% Materials',
      vals: block.jobs.slice(0, jobCount).map((j) => (j.income ? j.materials / j.income : 0)),
      fmt: PCT_FMT,
    },
  ];
  for (const ir of info) {
    set(r, startCol, ir.label);
    for (let i = 0; i < jobCount; i++) set(r, startCol + 1 + i, ir.vals[i], ir.fmt);
    // Round 10 — `-Mileage` IS the books mileage (the P&L `Mileage Expense`
    // for that AP#), so it references the row rather than repeating a value
    // that could drift once someone edits one of them.
    if (ir.label === 'Mileage in Books' && mileageRow > 0) {
      for (let i = 0; i < jobCount; i++) {
        const col = startCol + 1 + i;
        const cell = ws.getCell(mileageRow, col);
        cell.value = { formula: `${colLetters(col)}${r}` };
        cell.numFmt = MONEY_FMT;
      }
    }
    r++;
  }
  return r;
}

function notesJobText(customer: string, jobShort: string, notes: string): string {
  const base = `${customer} ${jobShort}`.trim();
  const tagged = `${base} [PS]`;
  const n = notes.trim();
  return n ? `${tagged} — ${n}` : tagged;
}

/** NOTES-only materials reconciliation table + WO note in col F. */
function writeLeadMaterialsTable(ws: Worksheet, block: PsBlock): void {
  const name = ws.addRow([block.displayName]);
  name.getCell(1).fill = solidFill('FFB8CCE4');
  name.getCell(1).font = { name: PP_FONT, bold: true, size: 12 };
  const hdr = ws.addRow(['PO#', 'Job', 'Reported Materials', 'Total Materials', 'DIFFERENCE']);
  for (let c = 1; c <= 5; c++) {
    hdr.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 12 };
  }
  for (const j of block.jobs.slice(0, 6)) {
    const wo = formatCrewLeadWoNote(j);
    const row = ws.addRow([
      j.apKey ? `AP${j.apKey}` : j.jobShort,
      notesJobText(j.customer, j.jobShort, j.notes),
      '', // Reported Materials — blank working cell
      wo.flagMissing ? '?Did we finish?' : j.materials,
      '', // DIFFERENCE — blank until reported is filled
      wo.text,
    ]);
    if (!wo.flagMissing) row.getCell(4).numFmt = MONEY_FMT;
    for (let c = 1; c <= 6; c++) {
      row.getCell(c).font = {
        name: PP_FONT,
        size: 12,
        ...(wo.flagMissing ? { color: { argb: DEDUCT_RED }, bold: true } : {}),
      };
    }
    if (j.colorHex) {
      row.getCell(1).fill = solidFill('FF' + j.colorHex);
      row.getCell(1).font = {
        name: PP_FONT,
        size: 12,
        color: { argb: wo.flagMissing ? DEDUCT_RED : textOn(j.colorHex) },
        bold: wo.flagMissing,
      };
    }
  }
  ws.addRow([]);
}

export function buildProfitshareOutputSheet(
  wb: Workbook,
  model: PayrollPreviewModel | null,
  defaultPsPct: number,
  mileageRate = 0.625,
): void {
  const ws = wb.addWorksheet('Profitshare Output');
  // 8-col bands: 7 data + 1 spacer → A, I, Q, Y…
  const allLeads = model?.leads ?? [];
  // Round 11 — hourly-only leads are left off; refs still index the Input #'s
  // tab, which lists every lead.
  const leads = profitshareOutputLeads(allLeads);
  const inputRefs = inputNumbersCellRefs(allLeads.length);
  ws.columns = Array.from({ length: Math.max(8, leads.length * 8) }, (_, i) => ({
    width: i % 8 === 7 ? 3 : 14,
  }));
  let maxEndRow = 1;
  leads.forEach(({ block, inputIndex }, bi) => {
    const startCol = 1 + bi * 8;
    const refs = inputRefs[inputIndex];
    const end = writePsChainValues(ws, block, startCol, {
      defaultPsPct,
      startRow: 1,
      style: 'profitshare',
      hourlyPayRef: refs?.hourlyCell,
      deductionsRef: refs?.deductCell,
      // Round 10 — Reimbursements, both notes and Mileage - New come live off
      // the Input #'s tab.
      reimbRef: refs?.reimbCell,
      reimbNoteRef: refs?.reimbNoteCell,
      deductNoteRef: refs?.deductNoteCell,
      milesRef: refs?.milesCell,
      mileageRate,
    });
    if (end > maxEndRow) maxEndRow = end;
  });
  if (leads.length === 0) {
    ws.addRow(['(no crew leads this week)']).font = { name: PP_FONT, italic: true };
    maxEndRow = 2;
  }
  // Footer lookups under the tallest band.
  let fr = maxEndRow + 1;
  ws.getCell(fr, 1).value = 'Lookups - DO NOT CHANGE:';
  ws.getCell(fr, 1).font = { name: PP_FONT, bold: true, size: 12 };
  fr++;
  ws.getCell(fr, 1).value = 'Mileage Rate';
  ws.getCell(fr, 2).value = mileageRate;
  ws.getCell(fr, 2).numFmt = '0.000';
  ws.getCell(fr, 1).font = { name: PP_FONT, size: 12 };
  ws.getCell(fr, 2).font = { name: PP_FONT, size: 12 };
  applyGaramond(ws);
}

export function buildManagementReportSheet(
  wb: Workbook,
  model: PayrollPreviewModel | null,
  defaultPsPct: number,
): void {
  const ws = wb.addWorksheet('Management Report Export');
  ws.columns = Array.from({ length: 8 }, () => ({ width: 16 }));
  const leads = model?.leads ?? [];
  const refs = inputNumbersCellRefs(leads.length);
  const jobCols = 6;
  const hdr = ws.addRow(['Crew', ...Array.from({ length: jobCols }, (_, i) => `Job ${i + 1}`)]);
  for (let c = 1; c <= 1 + jobCols; c++) {
    hdr.getCell(c).fill = solidFill(HEADER_GRAY);
    hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
  }
  leads.forEach((block, bi) => {
    const r = refs[bi];
    writePsChainValues(ws, block, 1, {
      defaultPsPct,
      style: 'management',
      hourlyPayRef: r?.hourlyCell,
      deductionsRef: r?.deductCell,
    });
    ws.addRow([]);
  });
  if (leads.length === 0) ws.addRow(['(no crew leads this week)']);
  applyGaramond(ws);
}

export function buildInputNumbersSheet(
  wb: Workbook,
  model: PayrollPreviewModel | null,
  /** Round 10 — $/mile, shown on the Miles row and used by Mileage - New. */
  mileageRate = 0.625,
  /** Round 10 — estimator initial → fill hex, for the PO#'s color coding. */
  estimatorColorByAp?: ReadonlyMap<string, string>,
): void {
  const ws = wb.addWorksheet("Input #'s");
  // Col A gutter ~3.4; data starts in B.
  ws.columns = [{ width: 3.4 }, ...Array.from({ length: 8 }, () => ({ width: 16 }))];
  const title = ws.getCell(1, 2);
  title.value = 'Crew Profit Share';
  title.font = { name: PP_FONT, bold: true, size: 20 };

  const leads = model?.leads ?? [];
  let row = 3;
  for (const block of leads) {
    const jobCount = Math.min(block.jobs.length, 6);
    // Red name row in B…
    const nameRow = ws.getRow(row);
    nameRow.getCell(2).value = block.displayName;
    nameRow.getCell(2).fill = solidFill(NAME_BAND_RED);
    nameRow.getCell(2).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
    for (let i = 0; i < jobCount; i++) {
      const pct = block.jobs[i].woPct;
      nameRow.getCell(3 + i).value = pct !== null ? pct : '';
      if (pct !== null) nameRow.getCell(3 + i).numFmt = PCT_FMT;
      nameRow.getCell(3 + i).fill = solidFill(NAME_BAND_RED);
      nameRow.getCell(3 + i).font = { name: PP_FONT, size: 12, color: { argb: WHITE } };
    }
    row++;

    // Grey PO labels + first PO yellow.
    const poRow = ws.getRow(row);
    poRow.getCell(2).value = "PO #'s:";
    poRow.getCell(2).fill = solidFill(HEADER_GRAY);
    poRow.getCell(2).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
    for (let i = 0; i < jobCount; i++) {
      const job = block.jobs[i];
      const cell = poRow.getCell(3 + i);
      cell.value = job.jobShort;
      /**
       * Round 10 — the PO cell wears the ESTIMATOR's color for that job (the
       * same fill the TimeNavi Export row carries), so the sheet says at a
       * glance who estimated it. Falls back to the old grey/yellow when the
       * AP# has no estimator color.
       */
      const hex = job.apKey ? estimatorColorByAp?.get(job.apKey) : undefined;
      if (hex) {
        cell.fill = solidFill('FF' + hex);
        cell.font = { name: PP_FONT, size: 11, bold: true, color: { argb: textOn(hex) } };
      } else {
        cell.font = { name: PP_FONT, size: 11, bold: true };
        cell.fill = solidFill(i === 0 ? HOURLY_PAY_YELLOW : HEADER_GRAY);
      }
    }
    row++;

    const woRow = ws.getRow(row);
    woRow.getCell(2).value = 'WO Amount/Contract Override:';
    woRow.getCell(2).font = { name: PP_FONT, size: 12 };
    for (let i = 0; i < jobCount; i++) {
      const v = block.jobs[i].workOrderOverride;
      woRow.getCell(3 + i).value = v !== null ? v : '';
      if (v !== null) woRow.getCell(3 + i).numFmt = ACCOUNTING_FMT;
      woRow.getCell(3 + i).font = { name: PP_FONT, size: 12 };
    }
    row++;

    const pctRow = ws.getRow(row);
    pctRow.getCell(2).value = '% Completion:';
    pctRow.getCell(2).font = { name: PP_FONT, size: 12 };
    for (let i = 0; i < jobCount; i++) {
      const j = block.jobs[i];
      pctRow.getCell(3 + i).value = j.manualPay ? j.manualPay : j.percentCompletion;
      pctRow.getCell(3 + i).numFmt = PCT_FMT;
      pctRow.getCell(3 + i).font = { name: PP_FONT, size: 12 };
    }
    row++;

    const matHdr = ws.getRow(row);
    matHdr.getCell(2).value = 'Reported Materials:';
    matHdr.getCell(3).value = 'Amount';
    matHdr.getCell(4).value = 'Notes';
    for (let c = 2; c <= 4; c++) matHdr.getCell(c).font = { name: PP_FONT, size: 12 };
    row++;
    row++; // blank reported-materials entry row

    const reimb = ws.getRow(row);
    reimb.getCell(2).value = 'Reimbursements/additional pay';
    reimb.getCell(3).value = block.reimbursements;
    reimb.getCell(3).numFmt = ACCOUNTING_FMT;
    reimb.getCell(4).value = block.reimbMemo;
    reimb.getCell(3).font = { name: PP_FONT, size: 12, color: { argb: REIMB_BLUE } };
    row++;

    // Round 8 §10.2 — Hourly Pay sits between Reimbursements and Deductions.
    const hourly = ws.getRow(row);
    hourly.getCell(2).value = 'Hourly Pay';
    hourly.getCell(3).value = block.hourlyPay;
    hourly.getCell(3).numFmt = ACCOUNTING_FMT;
    hourly.getCell(3).fill = solidFill(HOURLY_PAY_YELLOW);
    hourly.getCell(3).font = { name: PP_FONT, size: 12 };
    row++;

    const deduct = ws.getRow(row);
    deduct.getCell(2).value = 'Deductions/Sundries';
    deduct.getCell(3).value = block.deductions;
    deduct.getCell(3).numFmt = ACCOUNTING_FMT;
    deduct.getCell(4).value = block.deductMemo;
    deduct.getCell(3).font = { name: PP_FONT, size: 12, color: { argb: DEDUCT_RED } };
    row++;

    // Round 10 — the miles cell is the entry point for `Mileage - New` on the
    // Profitshare Output, which multiplies it by the rate shown here.
    const miles = ws.getRow(row);
    miles.getCell(2).value = `Miles: (@ $${mileageRate.toFixed(3)}/mi)`;
    miles.getCell(2).font = { name: PP_FONT, size: 12 };
    miles.getCell(3).numFmt = HOURS_FMT;
    miles.getCell(3).font = { name: PP_FONT, size: 12 };
    row += 2; // spacer between lead blocks
  }
  if (leads.length === 0) {
    ws.getCell(3, 2).value = '(no crew leads this week)';
    ws.getCell(3, 2).font = { name: PP_FONT, italic: true, size: 12 };
  }
  applyGaramond(ws);
}

function writeHourlyBlockValues(
  ws: Worksheet,
  block: HourlyBlock,
  opts?: { tsheets?: boolean; notesLayout?: boolean },
): void {
  const notesLayout = opts?.notesLayout === true;
  const colCount = notesLayout ? 7 : 5;
  if (opts?.tsheets) {
    const title = ws.addRow([`${block.employeeName.split(' ')[0]}'s T-SHEETS BREAKDOWN`]);
    for (let c = 1; c <= colCount; c++) {
      title.getCell(c).fill = solidFill('FFF9F9F9');
      title.getCell(c).font = { name: PP_FONT, bold: true, size: 12 };
      title.getCell(c).border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        ...(c === 1 ? { left: { style: 'thin' as const } } : {}),
        ...(c === colCount ? { right: { style: 'thin' as const } } : {}),
      };
    }
  } else {
    const nameRow = ws.addRow([
      block.crewTag ? `${block.employeeName} (${block.crewTag})` : block.employeeName,
    ]);
    nameRow.getCell(1).fill = solidFill(notesLayout ? HEADER_GRAY : TAN);
    nameRow.getCell(1).font = {
      name: PP_FONT,
      bold: true,
      size: 12,
      ...(notesLayout ? { color: { argb: WHITE } } : {}),
    };
  }
  const hdr = ws.addRow(
    notesLayout
      ? ['PO#', 'Job', 'hrs', 'rate', 'total', '(-) taxes', 'net']
      : ['PO#', 'Job', 'hrs', 'rate', 'total'],
  );
  for (let c = 1; c <= colCount; c++) {
    hdr.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 12 };
    hdr.getCell(c).alignment = { horizontal: 'center' };
  }
  for (const l of block.lines) {
    const total = Math.round(l.hours * l.rate * 100) / 100;
    const row = ws.addRow(
      notesLayout
        ? [l.poText ?? (l.apKey ? `AP${l.apKey}` : ''), l.jobLabel, l.hours, l.rate, total, '', '']
        : [l.poText ?? (l.apKey ? `AP${l.apKey}` : ''), l.jobLabel, l.hours, l.rate, total],
    );
    for (const c of [3, 4, 5]) {
      row.getCell(c).numFmt = HOURS_FMT;
      row.getCell(c).alignment = { horizontal: 'center' };
      row.getCell(c).font = { name: PP_FONT, size: 12 };
    }
    row.getCell(2).font = { name: PP_FONT, size: 10.5 };
  }
  const sums = notesLayout
    ? ws.addRow(['', '', block.totalHours, '', block.totalPay, block.taxes ?? '', block.net ?? ''])
    : ws.addRow(['', '', block.totalHours, '', block.totalPay]);
  sums.font = { name: PP_FONT, bold: true, size: 12 };
  sums.getCell(3).numFmt = HOURS_FMT;
  sums.getCell(5).numFmt = HOURS_FMT;
  if (notesLayout) {
    if (typeof block.taxes === 'number') sums.getCell(6).numFmt = HOURS_FMT;
    if (typeof block.net === 'number') sums.getCell(7).numFmt = HOURS_FMT;
    // Round 9 §3 — the NOTES log spells out the hour split: regular vs daily
    // overtime vs weekly overtime, taken from the timesheet's own pay items.
    for (const split of hourSplitRows(block)) {
      const row = ws.addRow(['', split.label, split.hours, '', split.pay, '', '']);
      row.getCell(2).font = { name: PP_FONT, size: 10.5, italic: true };
      for (const c of [3, 5]) {
        row.getCell(c).numFmt = HOURS_FMT;
        row.getCell(c).alignment = { horizontal: 'center' };
        row.getCell(c).font = { name: PP_FONT, size: 11 };
      }
    }
  }
  // Round 9 §19/§23 — configured extra rows under an hourly block.
  for (const x of block.extras) {
    const row = ws.addRow(
      notesLayout
        ? [x.label, '', '', '', x.amount ?? '', '', '']
        : [x.label, '', '', '', x.amount ?? ''],
    );
    row.getCell(5).numFmt = ACCOUNTING_FMT;
    row.getCell(5).alignment = { horizontal: 'center' };
  }
  if (block.extras.some((x) => x.kind !== 'note')) {
    const row = ws.addRow(
      notesLayout ? ['Total Pay', '', '', '', block.netPay, '', ''] : ['Total Pay', '', '', '', block.netPay],
    );
    row.font = { name: PP_FONT, bold: true, size: 12 };
    row.getCell(5).numFmt = ACCOUNTING_FMT;
    row.getCell(5).alignment = { horizontal: 'center' };
  }
  ws.addRow([]);
}


function writeSalaryBlockValues(ws: Worksheet, block: SalaryBlock, weekRange: string, fill: string): void {
  const band = ws.addRow([
    `${block.name}${block.timesheetHours > 0 ? ` (${block.timesheetHours}hrs)` : ''}` +
      // Round 9 §26/§33 — subcontractors in the salary sections.
      (block.sub ? ' (subcontractor)' : ''),
  ]);
  band.getCell(1).fill = solidFill(fill);
  band.getCell(1).font = { name: PP_FONT, bold: true, size: 12 };
  const r1 = ws.addRow([weekRange ? `(${weekRange})` : '', block.role, block.amount]);
  r1.getCell(3).numFmt = ACCOUNTING_FMT;
  // Round 9 §28 — flat labeled reimbursements row instead of mileage reimb.
  // Round 10 — some people carry no mileage row at all.
  const r2 = block.reimbursementsLabel
    ? ws.addRow([block.reimbursementsLabel, '', block.reimbursements])
    : block.noMileageRow
      ? null
      : ws.addRow([`mileage reimb: ${block.name} - ${block.miles} miles`, '', block.mileageReimb]);
  if (r2) r2.getCell(3).numFmt = ACCOUNTING_FMT;
  // Round 9 §19/§23/§24/§29 — configured extra rows.
  for (const x of block.extras) {
    const rx = ws.addRow([x.label, '', x.amount ?? '']);
    rx.getCell(3).numFmt = ACCOUNTING_FMT;
  }
  const r3 = ws.addRow(['Total', '', block.total]);
  r3.getCell(3).numFmt = ACCOUNTING_FMT;
  r3.font = { name: PP_FONT, bold: true, size: 12 };
  // Round 9 §34 — pre-filled taxes / net when the roster carries them.
  if (block.taxes !== null || block.net !== null) {
    const rt = ws.addRow(['(-) taxes', '', block.taxes ?? '']);
    rt.getCell(3).numFmt = ACCOUNTING_FMT;
    const rn = ws.addRow(['net', '', block.net ?? '']);
    rn.getCell(3).numFmt = ACCOUNTING_FMT;
    rn.font = { name: PP_FONT, bold: true, size: 12 };
  }
  ws.addRow([]);
}

/**
 * NOTES vs FINAL share painters / office / admin / estimators / managers,
 * but lead sections differ:
 * - NOTES: materials reconciliation + T-SHEETS (working log; no PS WO% band)
 * - FINAL: PS chain without WO% band, `@ NN% PS` notes, Total Gross Wages; no materials table
 */
function writePreviewSequence(
  ws: Worksheet,
  model: PayrollPreviewModel,
  opts: { mode: 'notes' | 'final'; defaultPsPct: number },
): number {
  let grossSum = 0;
  for (const entry of model.painterSequence) {
    if (entry.kind === 'hourly') {
      writeHourlyBlockValues(ws, entry.block, { notesLayout: opts.mode === 'notes' });
      grossSum += entry.block.totalPay;
    } else if (entry.kind === 'salary') {
      // Round 10 — a subcontractor's salary block, in with the subs.
      writeSalaryBlockValues(ws, entry.block, model.weekRange, TAN);
      grossSum += entry.block.total;
    } else if (opts.mode === 'notes') {
      // Working log: materials table → T-SHEETS (not the full PS WO% band).
      writeLeadMaterialsTable(ws, entry.block);
      if (entry.block.tsheets) {
        writeHourlyBlockValues(ws, entry.block.tsheets, { tsheets: true, notesLayout: true });
      }
      grossSum += entry.block.totalPay;
    } else {
      writePsChainValues(ws, entry.block, 1, {
        skipWoBand: true,
        defaultPsPct: opts.defaultPsPct,
        showGrossWages: true,
        // Round 9 §30 — FINAL drops the books/reference rows.
        hideBooksRows: true,
      });
      ws.addRow([]);
      if (entry.block.tsheets) writeHourlyBlockValues(ws, entry.block.tsheets, { tsheets: true });
      grossSum += entry.block.totalPay;
    }
  }
  const section = (title: string): void => {
    const row = ws.addRow([title]);
    row.getCell(1).fill = solidFill(BLACK);
    row.getCell(1).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
  };
  section('OFFICE');
  for (const b of model.office) {
    writeHourlyBlockValues(ws, b, { notesLayout: opts.mode === 'notes' });
    grossSum += b.totalPay;
  }
  section('ADMIN');
  for (const b of model.admin) {
    writeSalaryBlockValues(ws, b, model.weekRange, TAN);
    grossSum += b.total;
  }
  section('ESTIMATORS');
  for (const b of model.estimators) {
    writeSalaryBlockValues(ws, b, model.weekRange, TAN);
    grossSum += b.total;
  }
  section('MANAGER SALARY');
  for (const b of model.managers) {
    writeSalaryBlockValues(ws, b, model.weekRange, 'FFD99694');
    grossSum += b.total;
  }
  return grossSum;
}

function payrollBreakdownTitle(
  payDate: Date | null,
  weekStart: Date | null,
  weekEnd: Date | null,
  withGuillemet: boolean,
): string {
  const pay = payDate
    ? `${payDate.getMonth() + 1}/${payDate.getDate()}/${payDate.getFullYear()}`
    : '';
  const range =
    weekStart && weekEnd
      ? `${weekStart.getMonth() + 1}/${weekStart.getDate()} - ${weekEnd.getMonth() + 1}/${weekEnd.getDate()}`
      : '';
  const core = `PAYROLL BREAKDOWN ${pay}${range ? ` (${range})` : ''}`.trim();
  return withGuillemet ? `«${core}` : core;
}

export function buildNotesSheet(
  wb: Workbook,
  input: {
    payDate: Date | null;
    weekStart: Date | null;
    weekEnd: Date | null;
    payrollPreview?: PayrollPreviewModel | null;
    defaultProfitSharePct?: number;
  },
): void {
  const ws = wb.addWorksheet('NOTES');
  ws.columns = Array.from({ length: 8 }, () => ({ width: 16 }));
  const title = ws.addRow([
    payrollBreakdownTitle(input.payDate, input.weekStart, input.weekEnd, true),
  ]);
  title.getCell(1).font = { name: PP_FONT, bold: true, size: 13 };
  const model = input.payrollPreview ?? null;
  if (!model) {
    ws.addRow(['(no payroll preview)']);
    applyGaramond(ws);
    return;
  }
  const gross = writePreviewSequence(ws, model, {
    mode: 'notes',
    defaultPsPct: input.defaultProfitSharePct ?? 0.38,
  });
  // Grand-total cell on row 1 (computed value — never a cross-sheet #REF!).
  ws.getCell(1, 10).value = gross;
  ws.getCell(1, 10).numFmt = MONEY_FMT;
  ws.getCell(1, 10).font = { name: PP_FONT, bold: true, size: 12 };
  applyGaramond(ws);
}

export function buildFinalPayrollSheet(
  wb: Workbook,
  input: {
    payDate: Date | null;
    weekStart: Date | null;
    weekEnd: Date | null;
    payrollPreview?: PayrollPreviewModel | null;
    defaultProfitSharePct?: number;
  },
): void {
  const ws = wb.addWorksheet('FINAL PAYROLL BREAKDOWN');
  ws.columns = Array.from({ length: 10 }, () => ({ width: 16 }));
  const title = ws.addRow([
    payrollBreakdownTitle(input.payDate, input.weekStart, input.weekEnd, false),
  ]);
  title.getCell(1).font = { name: PP_FONT, bold: true, size: 13 };
  const model = input.payrollPreview ?? null;
  if (!model) {
    ws.addRow(['(no payroll preview)']);
    applyGaramond(ws);
    return;
  }
  writePreviewSequence(ws, model, {
    mode: 'final',
    defaultPsPct: input.defaultProfitSharePct ?? 0.38,
  });
  applyGaramond(ws);
}

/** Round 7 §18 — header shell only (PO# | blank | Job). */
export function buildJournalEntrySheet(wb: Workbook): void {
  const ws = wb.addWorksheet('►JOURNAL ENTRY◄');
  ws.columns = [{ width: 12 }, { width: 12 }, { width: 48 }];
  const headers = ['PO#', '', 'Job'];
  const row = ws.addRow(headers);
  for (let c = 1; c <= headers.length; c++) {
    row.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 10.5 };
  }
}

export function buildJobsSheet(
  wb: Workbook,
  input: {
    jobs?: { poLabel: string; jobLabel: string; apKey?: string | null }[] | null;
    payrollPreview?: PayrollPreviewModel | null;
    /** Round 7 §19 — apKey → estimator initials for PO#/Job fill. */
    estimatorByApKey?: Map<string, string>;
    legend?: { estimator: string; color: string }[];
  },
): void {
  const ws = wb.addWorksheet('JOBS');
  ws.columns = [{ width: 12 }, { width: 48 }];
  const font = { name: PP_FONT, size: 10.5 };
  boldHeader(ws, ['PO#', 'Job'], PP_FONT);
  ws.getRow(1).eachCell((c) => {
    c.font = { ...font, bold: true };
  });

  const colorByEstimator = new Map(
    (input.legend ?? []).map((l) => [l.estimator, l.color.replace(/^#/, '').toUpperCase()]),
  );

  let rows = input.jobs ?? null;
  if (!rows) {
    const map = new Map<string, { jobLabel: string; apKey: string }>();
    for (const lead of input.payrollPreview?.leads ?? []) {
      for (const j of lead.jobs) {
        if (!j.apKey) continue;
        const po = `AP${j.apKey}`;
        if (!map.has(po)) map.set(po, { jobLabel: `${j.customer} ${j.jobShort}`.trim(), apKey: j.apKey });
      }
    }
    for (const entry of input.payrollPreview?.painterSequence ?? []) {
      const block =
        entry.kind === 'hourly' ? entry.block : entry.kind === 'ps' ? entry.block.tsheets : null;
      if (!block) continue;
      for (const l of block.lines) {
        if (!l.apKey) continue;
        const po = `AP${l.apKey}`;
        if (!map.has(po)) map.set(po, { jobLabel: l.jobLabel, apKey: l.apKey });
      }
    }
    rows = [...map.entries()]
      .map(([poLabel, v]) => ({ poLabel, jobLabel: v.jobLabel, apKey: v.apKey }))
      .sort((a, b) => a.jobLabel.localeCompare(b.jobLabel, 'en', { sensitivity: 'base' }));
  }

  for (const r of rows) {
    const row = ws.addRow([r.poLabel, r.jobLabel]);
    row.getCell(1).font = { ...font };
    row.getCell(2).font = { ...font };
    const apKey = r.apKey ?? extractApKey(r.poLabel);
    const estimator = apKey ? input.estimatorByApKey?.get(apKey) : undefined;
    const hex = estimator ? colorByEstimator.get(estimator) : undefined;
    if (hex) {
      const fill = solidFill(hex.length === 8 ? hex : `FF${hex}`);
      const fg = { argb: textOn(hex.replace(/^FF/i, '').slice(0, 6)) };
      row.getCell(1).fill = fill;
      row.getCell(2).fill = fill;
      row.getCell(1).font = { ...font, color: fg };
      row.getCell(2).font = { ...font, color: fg };
    }
  }
  applyGaramond(ws, 10.5);
}

export function buildInactiveEmployeesSheet(
  wb: Workbook,
  input: {
    inactiveEmployees?: InactiveEmployeeBlock[] | null;
    payrollPreview?: PayrollPreviewModel | null;
  },
): void {
  const ws = wb.addWorksheet('Inactive Employees');
  ws.columns = [{ width: 36 }, { width: 16 }];
  const list =
    input.inactiveEmployees ??
    input.payrollPreview?.inactiveEmployees ??
    (input.payrollPreview?.inactiveZeroHours ?? []).map((name) => ({
      name,
      displayName: name,
      rateNote: '',
      isLead: false,
      directoryInactive: true,
    }));
  const sorted = [...list].sort((a, b) => a.displayName.localeCompare(b.displayName, 'en', { sensitivity: 'base' }));
  for (const p of sorted) {
    const bandLabel = p.rateNote ? `${p.displayName}  ${p.rateNote}` : p.displayName;
    const band = ws.addRow([bandLabel]);
    band.getCell(1).fill = solidFill(TAN);
    band.getCell(1).font = { name: PP_FONT, bold: true, size: 12 };
    if (p.isLead) {
      const matHdr = ws.addRow(['PO#', 'Job', 'Reported Materials', 'Total Materials', 'DIFFERENCE']);
      for (let c = 1; c <= 5; c++) {
        matHdr.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 12 };
      }
      ws.addRow([]);
    }
    const hdr = ws.addRow(['PO#', 'Job', 'hrs', 'rate', 'total']);
    for (let c = 1; c <= 5; c++) {
      hdr.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 12 };
    }
    const rateNum = Number.parseFloat(p.rateNote.replace(/[^0-9.]/g, '')) || 0;
    ws.addRow(['', '', 0, rateNum, 0]);
    ws.addRow([]);
  }
  if (sorted.length === 0) {
    ws.addRow(['(no inactive employees without hours)']).font = {
      name: PP_FONT,
      italic: true,
      size: 12,
    };
  }
  applyGaramond(ws);
}

/** Round 7 §2 — PR Final Check (last payroll sheet). Dynamic from P&L byJob. */
export function buildPrFinalCheckSheet(
  wb: Workbook,
  pl: PlParseResult | null,
  sourceFileName?: string,
): void {
  const ws = wb.addWorksheet('PR Final Check', { views: [{ state: 'frozen', ySplit: 1 }] });
  const headers = [
    'PO#',
    'Job',
    'Total Income',
    'Materials',
    'QB Labor-Wages-Subs',
    'Mileage Expense',
    'Sundries -',
  ];
  ws.columns = headers.map((h) => ({ width: h === 'Job' ? 36 : 16 }));
  const hdr = ws.addRow(headers);
  for (let c = 1; c <= headers.length; c++) {
    hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 11 };
  }
  // Legend cell names the P&L input file.
  const legendCol = headers.length + 2;
  ws.getCell(1, legendCol).value = sourceFileName?.trim() ? sourceFileName : '';
  ws.getCell(1, legendCol).font = { name: PP_FONT, italic: true, size: 10 };
  ws.getColumn(legendCol).width = 28;

  const rows = pl ? buildPrFinalCheckRows(pl) : [];
  for (const r of rows) {
    const job =
      r.customerName && r.jobHeader && !r.jobHeader.includes(r.customerName)
        ? `${r.customerName} ${r.jobHeader}`.trim()
        : r.customerName || r.jobHeader || `AP${r.apKey}`;
    const row = ws.addRow([
      `AP${r.apKey}`,
      job,
      r.income,
      r.materials,
      r.laborWagesSubs,
      r.mileage,
      r.sundries,
    ]);
    row.getCell(1).font = { name: PP_FONT, size: 11 };
    row.getCell(2).font = { name: PP_FONT, size: 11 };
    for (let c = 3; c <= 7; c++) {
      row.getCell(c).numFmt = ACCOUNTING_FMT;
      row.getCell(c).font = { name: PP_FONT, size: 11 };
    }
  }
  applyGaramond(ws, 11);
}
