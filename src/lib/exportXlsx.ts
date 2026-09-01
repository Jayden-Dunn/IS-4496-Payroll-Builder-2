import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { isDarkColor, percentExcelValue } from './format';
import type { JobsByCrewModel } from './jobsByCrew';
import type {
  SalesByCustMirrorRow,
  SalesByRepMirrorRow,
  TimesheetMirrorRow,
} from './mirrorTabs';
import {
  buildFinalPayrollSheet,
  buildInactiveEmployeesSheet,
  buildInputNumbersSheet,
  buildJobsByCrewSheet,
  buildJobsSheet,
  buildJournalEntrySheet,
  buildManagementReportSheet,
  buildNotesSheet,
  buildPlImportSheet,
  buildPrFinalCheckSheet,
  buildProfitshareOutputSheet,
  buildSalesByCustMirrorSheet,
  buildSalesByRepMirrorSheet,
  buildTimesheetMirrorSheet,
  inputNumbersCellRefs,
  inputNumsSheetFormula,
} from './exportRound6Sheets';
import { weekdayShort } from './pivotByJob';
import {
  dollarDiffIsFlagged,
  duplicateInQboValues,
  percentIsFlagged,
  percentIsRedFlagged,
} from './timenavi';
import type {
  ExceptionGroup,
  HourlyBlock,
  InactiveEmployeeBlock,
  LegendEntry,
  ManualSectionRow,
  PayrollPreviewModel,
  PercentValue,
  PivotByJobModel,
  PivotModel,
  PlParseResult,
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

// Round 8 §2.4 — a % Completion cell with no valid state.
const RED_FILL = 'FFFF0000';

// Reference formats: accounting style WITHOUT a $ symbol.
// Round 10 — negatives read red and parenthesised in the workbook too.
const ACCOUNTING_FMT = '_(* #,##0.00_);[Red]_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const START_END_DATE_FMT = 'mm-dd-yy';
const INVOICE_DATE_FMT = 'mm/dd/yyyy';
const PCT_FMT = '0%';
const HOURS_FMT = '0.00';

/**
 * Round 9 §38 — accounting format WITH the "$" symbol. Used by the payroll
 * bands and by the TimeNavi `Apptivo $` column, whose values are dollars.
 */
const CURRENCY_FMT = '_("$"* #,##0.00_);[Red]_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';

/**
 * Round 9 §2/§36 — the AP# column is autofit, then widened by this factor so
 * every AP number (suffixes included) fits without truncation.
 */
const TN_AP_WIDTH_SCALE = 1.35;

/**
 * Round 9 §37 — TimeNavi rows are pinned to a compact single-line height.
 * Nothing on the sheet wraps, so the tall auto-height rows are unnecessary;
 * 12.75 is the shortest height that still shows the 10pt Calibri text in full.
 */
const TN_ROW_HEIGHT = 12.75;
const TN_HEADER_ROW_HEIGHT = 12.75;

const FONT = 'Calibri';

function solidFill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

// Round 8 §4.1 — carried jobs painted this gray must not flow to payroll.
const EXCLUSION_GRAY_HEX = 'A6A6A6';

/** Text format, so AP# columns compare as text on both sides of a MATCH. */
const TEXT_FMT = '@';

/**
 * Round 10 — AP# → estimator fill hex, from the resolved estimator map and the
 * Settings palette. Used to color the Input #'s `PO #'s:` cells.
 */
function estimatorFillByApKey(
  estimatorByApKey: ReadonlyMap<string, string> | undefined,
  legend: readonly LegendEntry[] | undefined,
): Map<string, string> {
  const palette = new Map((legend ?? []).map((l) => [l.estimator, l.color.toUpperCase()]));
  const out = new Map<string, string>();
  for (const [ap, est] of estimatorByApKey ?? []) {
    const hex = palette.get(est);
    if (hex) out.set(ap, hex);
  }
  return out;
}

/** 1-based column number → Excel letters ("H", "AA"). */
export function excelColumnLetter(col: number): string {
  let n = Math.max(1, Math.trunc(col));
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/**
 * Round 8 §3.2 — cross-sheet live link for a `% of Completion` cell.
 * INDEX/MATCH only (XLOOKUP is not safe in LibreOffice); the AP# columns on
 * both sheets hold text so MATCH resolves. Duplicate AP# → first match.
 */
export function timeNaviPercentFormula(opts: {
  sheetName: string;
  apColumn: string;
  pctColumn: string;
  apCellRef: string;
}): string {
  const sheet = `'${opts.sheetName.replace(/'/g, "''")}'`;
  return `IFERROR(INDEX(${sheet}!${opts.pctColumn}:${opts.pctColumn}, MATCH(${opts.apCellRef}, ${sheet}!${opts.apColumn}:${opts.apColumn}, 0)), "")`;
}

interface TimeNaviLink {
  sheetName: string;
  apColumn: string;
  pctColumn: string;
  /** AP#s actually present on the sheet — a formula is only worth writing for these. */
  apKeys: Set<string>;
}

/**
 * Resolve the TimeNavi Export link columns AT EXPORT TIME by header text, so
 * reordering that sheet can never point the formulas at the wrong column.
 */
function resolveTimeNaviLink(wb: ExcelJS.Workbook): TimeNaviLink | null {
  const ws = wb.getWorksheet(TIMENAVI_SHEET_NAME);
  if (!ws) return null;
  const header = ws.getRow(1);
  let apCol = 0;
  let pctCol = 0;
  for (let c = 1; c <= header.cellCount; c++) {
    const v = header.getCell(c).value;
    const text = typeof v === 'string' ? v.trim() : '';
    if (text === 'AP#' && apCol === 0) apCol = c;
    else if (text === '% Completion' && pctCol === 0) pctCol = c;
  }
  if (apCol === 0 || pctCol === 0) return null;
  const apKeys = new Set<string>();
  for (let r = 2; r <= ws.rowCount; r++) {
    const v = ws.getRow(r).getCell(apCol).value;
    const text = typeof v === 'string' ? v.trim() : typeof v === 'number' ? String(v) : '';
    if (text) apKeys.add(text);
  }
  return {
    sheetName: TIMENAVI_SHEET_NAME,
    apColumn: excelColumnLetter(apCol),
    pctColumn: excelColumnLetter(pctCol),
    apKeys,
  };
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

/** Round 6 §9 — %'s workbook export input. */
export interface PercentsExportInput {
  rows: ReportRow[]; // already sorted; only included rows are exported
  pivot: PivotModel;
  legend: LegendEntry[];
  payDate: Date | null;
  pivotByJob?: PivotByJobModel | null;
  percentExceptions?: ExceptionGroup[] | null;
  /** §3.1–§3.5 mirrors (pre-built via mirrorTabs.ts). */
  salesByRep?: { headers: string[]; rows: SalesByRepMirrorRow[] } | null;
  salesByCust?: SalesByCustMirrorRow[] | null;
  utSalesByRep?: { headers: string[]; rows: SalesByRepMirrorRow[] } | null;
  utSalesByCust?: SalesByCustMirrorRow[] | null;
  timesheetMirror?: TimesheetMirrorRow[] | null;
  jobsByCrew?: JobsByCrewModel | null;
}

/** Round 6 §9 — payroll workbook export input. */
export interface PayrollExportInput {
  payDate: Date | null;
  weekStart: Date | null;
  weekEnd: Date | null;
  /** Slot B — the NV P&L. */
  pl?: PlParseResult | null;
  /** Round 8 §7.2 — Slot C, the optional UT P&L. Its sheets are omitted when absent. */
  plUt?: PlParseResult | null;
  /** P&L upload filename for PR Final Check legend cell. */
  plSourceFileName?: string | null;
  /** The NV `PAYROLL PREVIEW` model. */
  payrollPreview?: PayrollPreviewModel | null;
  /** Round 8 §7.3 — the `Payroll Preview - UT` model; null on an NV-only week. */
  payrollPreviewUt?: PayrollPreviewModel | null;
  /**
   * Round 8 §7.3 — model backing the shared downstream sheets. Defaults to
   * `payrollPreview`; pass `mergePayrollPreviews(nv, ut)` to show both entities.
   */
  payrollPreviewAll?: PayrollPreviewModel | null;
  payrollExceptions?: ExceptionGroup[] | null;
  /**
   * Canonical Jobs sheet rows (`PO#` + display name). When omitted, derived
   * from `payrollPreview` lead/painter AP keys.
   */
  jobs?: { poLabel: string; jobLabel: string; apKey?: string | null }[] | null;
  /** Inactive Employees tab rows; falls back to `payrollPreview.inactiveEmployees`. */
  inactiveEmployees?: InactiveEmployeeBlock[] | null;
  mileageRate?: number;
  defaultProfitSharePct?: number;
  /** Round 7 §19 — apKey → estimator initials for JOBS sheet fills. */
  estimatorByApKey?: Map<string, string>;
  legend?: LegendEntry[];
}

/** ExcelJS tab colors (ARGB). Sheets omitted / `null` keep the default (no color). */
function applyTabColors(wb: ExcelJS.Workbook, colors: ReadonlyMap<string, string>): void {
  for (const ws of wb.worksheets) {
    const argb = colors.get(ws.name);
    if (argb) ws.properties.tabColor = { argb };
  }
}

const PERCENTS_TAB_COLORS = new Map<string, string>([
  ['Sales by Rep', 'FF00B050'],
  ['Sales by Cust', 'FF00B050'],
  ['UT SalesbyRep', 'FF00B050'],
  ['UT SalesbyCust', 'FF00B050'],
  ['timesheets', 'FF7030A0'],
  ['TimeNavi Export', 'FFFFFF00'],
  ['Pivot by Crew', 'FFC00000'],
  // Jobs by Crew — none
  ['Pivot by Job', 'FFFFFF00'],
  // %'s Exceptions — none
]);

const PAYROLL_TAB_COLORS = new Map<string, string>([
  ['P&L by Job Import', 'FFC4BD97'],
  ['P&L by Job Import - UT', 'FFC4BD97'],
  ['Management Report Export', 'FF76933C'],
  ['PAYROLL PREVIEW', 'FFF79646'],
  ['Payroll Preview - UT', 'FFF79646'],
  ['Profitshare Output', 'FF0070C0'],
  ["Input #'s", 'FFFF0000'],
  ['NOTES', 'FF000000'],
  ['FINAL PAYROLL BREAKDOWN', 'FFFFFF00'],
  // ►JOURNAL ENTRY◄ — none
  ['JOBS', 'FF33CCCC'],
  // Inactive Employees / Payroll Exceptions / PR Final Check — none
]);

/** @deprecated Prefer PercentsExportInput — kept for Round 1–5 call sites. */
export type ExportInput = PercentsExportInput & {
  payrollPreview?: PayrollPreviewModel | null;
  payrollExceptions?: ExceptionGroup[] | null;
};

/**
 * Round 10 — `% of Completion MM-DD-YYYY.xlsx`: the `%` symbol leads, and the
 * spelled-out "Percentages" is gone. (Supersedes Round 8 §13, which spelled it
 * out to dodge filename mangling.)
 */
export function percentsExportFilename(payDate: Date | null): string {
  const d = payDate ?? new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `% of Completion ${mm}-${dd}-${d.getFullYear()}.xlsx`;
}

/** Alias — same as percentsExportFilename. */
export function exportFilename(payDate: Date | null): string {
  return percentsExportFilename(payDate);
}

/** `MM-DD-YYYY payroll preview (M-D to M-D).xlsx` */
export function payrollExportFilename(
  payDate: Date | null,
  weekStart: Date | null,
  weekEnd: Date | null,
): string {
  const d = payDate ?? new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const pay = `${mm}-${dd}-${d.getFullYear()}`;
  const range =
    weekStart && weekEnd
      ? `${weekStart.getMonth() + 1}-${weekStart.getDate()} to ${weekEnd.getMonth() + 1}-${weekEnd.getDate()}`
      : 'week';
  return `${pay} payroll preview (${range}).xlsx`;
}

function downloadBuffer(buf: ArrayBuffer, filename: string): void {
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  saveAs(blob, filename);
}

/**
 * Round 6 §9 — %'s workbook sheet order:
 * Sales by Rep | Sales by Cust | UT SalesbyRep | UT SalesbyCust | timesheets |
 * TimeNavi Export | Pivot by Crew | Jobs by Crew | Pivot by Job | %'s Exceptions
 */
export async function buildPercentsWorkbookBuffer(input: PercentsExportInput): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.created = new Date();
  const palette = new Map(input.legend.map((l) => [l.estimator, l.color.toUpperCase()]));

  buildSalesByRepMirrorSheet(wb, 'Sales by Rep', input.salesByRep);
  buildSalesByCustMirrorSheet(wb, 'Sales by Cust', input.salesByCust);
  buildSalesByRepMirrorSheet(wb, 'UT SalesbyRep', input.utSalesByRep);
  buildSalesByCustMirrorSheet(wb, 'UT SalesbyCust', input.utSalesByCust);
  buildTimesheetMirrorSheet(wb, input.timesheetMirror);
  buildTimeNaviSheet(wb, input, palette);
  buildPivotSheet(wb, input.pivot, palette);
  buildJobsByCrewSheet(wb, input.jobsByCrew);
  if (input.pivotByJob) buildPivotByJobSheet(wb, input.pivotByJob);
  buildExceptionsSheet(wb, "%'s Exceptions", input.percentExceptions ?? [], input.payDate);
  applyTabColors(wb, PERCENTS_TAB_COLORS);

  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}

/**
 * Round 6/7/8 — payroll workbook sheet order:
 * P&L by Job Import | [P&L by Job Import - UT] | Management Report Export |
 * PAYROLL PREVIEW | [Payroll Preview - UT] | Profitshare Output | Input #'s |
 * NOTES | FINAL PAYROLL BREAKDOWN | ►JOURNAL ENTRY◄ | JOBS |
 * Inactive Employees | Payroll Exceptions | PR Final Check
 *
 * Round 8 §7.3 — the two bracketed UT sheets appear only when the week has UT
 * uploads, so an NV-only week produces exactly the Round 7 workbook.
 * Split from `buildPayrollWorkbookBuffer` so the sheet order is assertable
 * without serializing a whole workbook.
 */
export function assemblePayrollWorkbook(input: PayrollExportInput): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  wb.created = new Date();
  const shared = input.payrollPreviewAll ?? input.payrollPreview ?? null;
  const sharedInput: PayrollExportInput = { ...input, payrollPreview: shared };
  const leadList = shared?.leads ?? [];
  const refs = inputNumbersCellRefs(leadList.length);
  const refsByLeadKey = new Map(leadList.map((b, i) => [b.key, refs[i]!]));

  buildPlImportSheet(wb, input.pl ?? null);
  if (input.plUt) buildPlImportSheet(wb, input.plUt, 'P&L by Job Import - UT');
  buildManagementReportSheet(wb, shared, input.defaultProfitSharePct ?? 0.38);
  if (input.payrollPreview) buildPayrollPreviewSheet(wb, input.payrollPreview, 'PAYROLL PREVIEW', refsByLeadKey);
  else wb.addWorksheet('PAYROLL PREVIEW');
  if (input.payrollPreviewUt) {
    buildPayrollPreviewSheet(wb, input.payrollPreviewUt, 'Payroll Preview - UT', refsByLeadKey);
  }
  buildProfitshareOutputSheet(
    wb,
    shared,
    input.defaultProfitSharePct ?? 0.38,
    input.mileageRate ?? 0.625,
  );
  buildInputNumbersSheet(
    wb,
    shared,
    input.mileageRate ?? 0.625,
    // Round 10 — PO#'s cells wear the estimator's color for that AP#.
    estimatorFillByApKey(input.estimatorByApKey, input.legend),
  );
  buildNotesSheet(wb, sharedInput);
  buildFinalPayrollSheet(wb, sharedInput);
  buildJournalEntrySheet(wb);
  buildJobsSheet(wb, sharedInput);
  buildInactiveEmployeesSheet(wb, sharedInput);
  buildExceptionsSheet(wb, 'Payroll Exceptions', input.payrollExceptions ?? [], input.payDate);
  buildPrFinalCheckSheet(wb, input.pl ?? null, input.plSourceFileName ?? undefined);
  applyTabColors(wb, PAYROLL_TAB_COLORS);
  return wb;
}

export async function buildPayrollWorkbookBuffer(input: PayrollExportInput): Promise<ArrayBuffer> {
  return (await assemblePayrollWorkbook(input).xlsx.writeBuffer()) as ArrayBuffer;
}

/** @deprecated Prefer buildPercentsWorkbookBuffer. */
export async function buildWorkbookBuffer(input: ExportInput): Promise<ArrayBuffer> {
  return buildPercentsWorkbookBuffer(input);
}

export async function exportPercentsWorkbook(input: PercentsExportInput): Promise<void> {
  downloadBuffer(await buildPercentsWorkbookBuffer(input), percentsExportFilename(input.payDate));
}

export async function exportPayrollWorkbook(input: PayrollExportInput): Promise<void> {
  downloadBuffer(
    await buildPayrollWorkbookBuffer(input),
    payrollExportFilename(input.payDate, input.weekStart, input.weekEnd),
  );
}

/** @deprecated Prefer exportPercentsWorkbook. */
export async function exportWorkbook(input: ExportInput): Promise<void> {
  return exportPercentsWorkbook(input);
}

// ---------------------------------------------------------------------------
// Sheet 1 — TimeNavi Export (columns A–S)
// ---------------------------------------------------------------------------

const TIMENAVI_SHEET_NAME = 'TimeNavi Export';

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

// Default column widths (A–S). Event and AP# are overwritten with autofit
// below; Description is then scaled to ~60% of this stub (a 40% cut).
const TN_WIDTHS = [36, 11, 13, 13, 9, 13, 13, 6.5, 11.6, 9, 10, 39, 20, 12, 14, 10.6, 9, 12.4, 13];
const TN_EVENT_MAX = 36;
const TN_DESC_SCALE = 0.6;

/** Excel column width ≈ character count + pad, clamped. */
function autofitColWidth(values: readonly string[], opts?: { min?: number; max?: number; pad?: number }): number {
  const pad = opts?.pad ?? 1;
  const min = opts?.min ?? 4;
  const max = opts?.max ?? 80;
  let longest = 0;
  for (const v of values) {
    if (v.length > longest) longest = v.length;
  }
  return Math.min(max, Math.max(min, longest + pad));
}

/** 1-based sheet column number for a TimeNavi header. */
function tnCol(header: string): number {
  return TN_HEADERS.indexOf(header) + 1;
}

function buildTimeNaviSheet(wb: ExcelJS.Workbook, input: ExportInput, palette: ReadonlyMap<string, string>): void {
  const ws = wb.addWorksheet(TIMENAVI_SHEET_NAME, {
    views: [{ state: 'frozen', ySplit: 1 }], // freeze at A2
  });
  const apIdx = TN_HEADERS.indexOf('AP#');
  const descIdx = TN_HEADERS.indexOf('Description');
  const included = input.rows.filter((row) => row.included);
  const eventWidth = autofitColWidth(
    ['Event', ...included.map((r) => r.event)],
    { min: 12, max: TN_EVENT_MAX },
  );
  const descWidth = Math.round(TN_WIDTHS[descIdx] * TN_DESC_SCALE * 10) / 10;
  // Round 9 §2/§36 — autofit, then 35% wider so no AP number is clipped.
  const apWidth =
    Math.round(
      autofitColWidth(['AP#', ...included.map((r) => r.apKey ?? '')], { min: 4, max: 18 }) *
        TN_AP_WIDTH_SCALE *
        10,
    ) / 10;
  ws.columns = TN_WIDTHS.map((width, i) => {
    if (i === 0) return { width: eventWidth };
    if (i === descIdx) return { width: descWidth };
    if (i === apIdx) return { width: apWidth };
    return { width };
  });

  const centeredCols = TN_CENTERED_HEADERS.map(tnCol);

  // Header: the ENTIRE row (A1 → last populated header) is black with white
  // non-bold text — no green remains. Paint by column index so every header
  // cell is covered (ExcelJS eachCell can skip cells in some edge cases).
  // Round 7 — Calibri 10 for all used TimeNavi cells.
  const tnFont = { name: FONT, size: 10 };
  const header = ws.addRow(TN_HEADERS);
  header.height = TN_HEADER_ROW_HEIGHT; // Round 9 §37
  for (let c = 1; c <= TN_HEADERS.length; c++) {
    const cell = header.getCell(c);
    cell.value = TN_HEADERS[c - 1];
    cell.fill = solidFill(BLACK);
    cell.font = { ...tnFont, color: { argb: WHITE }, bold: false };
    cell.alignment = {
      vertical: 'middle',
      wrapText: false,
      ...(centeredCols.includes(c) ? { horizontal: 'center' as const } : {}),
    };
  }

  // Round 9 §38 — the whole `Apptivo $` column is typed as dollars, so a value
  // typed into a blank cell later still formats as currency.
  ws.getColumn(tnCol('Apptivo $')).numFmt = CURRENCY_FMT;

  // Hide (never delete) the reference's hidden columns, located by header text.
  for (const h of TN_HIDDEN_HEADERS) {
    const idx = TN_HEADERS.indexOf(h);
    if (idx >= 0) ws.getColumn(idx + 1).hidden = true;
  }

  const dupInQbo = duplicateInQboValues(input.rows);
  const pctCol = tnCol('% Completion');
  const apCol = tnCol('AP#');
  const inQboCol = tnCol('In QBO Time?');
  const diffCol = tnCol('$ Difference');

  for (const r of included) {
    const percentCell = percentExcelValue(r.percent);
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
      r.apptivo ?? '', // §2.2 — blank unless the user typed a value
      r.dollarDiff ?? '', // replaced by the §6 formula below when both sides exist
      percentCell, // Q mirrors I — the office PDF reads this column
      r.invoiceDate ? excelDate(r.invoiceDate) : '',
      r.dateDiff ?? '',
    ]);

    // Round 10 — `$ Difference` carries the formula into the workbook:
    // Job Total $ − Apptivo $, live for whatever the office types afterwards.
    {
      const jobTotalRef = `${excelColumnLetter(tnCol('Job Total $'))}${row.number}`;
      const apptivoRef = `${excelColumnLetter(tnCol('Apptivo $'))}${row.number}`;
      const diffCell = row.getCell(tnCol('$ Difference'));
      diffCell.value = {
        formula: `IF(OR(NOT(ISNUMBER(${jobTotalRef})),${apptivoRef}=""),"",${jobTotalRef}-${apptivoRef})`,
        result: r.dollarDiff ?? undefined,
      };
    }

    row.getCell(2).numFmt = START_END_DATE_FMT;
    row.getCell(3).numFmt = START_END_DATE_FMT;
    row.getCell(18).numFmt = INVOICE_DATE_FMT;
    // §3.2 — Pivot by Job MATCHes against this column; keep it text on both sides.
    row.getCell(apCol).numFmt = TEXT_FMT;
    for (const col of [9, 11, 17]) {
      const cell = row.getCell(col);
      if (typeof cell.value === 'number') cell.numFmt = PCT_FMT;
    }
    for (const col of [14, 16]) {
      const cell = row.getCell(col);
      if (typeof cell.value === 'number') cell.numFmt = ACCOUNTING_FMT;
    }
    // Round 9 §38 — `Apptivo $` holds dollars, so it carries the $ format.
    {
      const apptivoCell = row.getCell(tnCol('Apptivo $'));
      if (typeof apptivoCell.value === 'number') apptivoCell.numFmt = CURRENCY_FMT;
    }
    row.getCell(19).numFmt = '0';

    // Round 9 §37 — halve the row height; nothing on this sheet wraps.
    row.height = TN_ROW_HEIGHT;
    row.getCell(1).alignment = { vertical: 'top' };

    // Full-row fill keyed off the RESOLVED estimator, with readable text.
    const hex = palette.get(r.estimator);
    for (let c = 1; c <= TN_HEADERS.length; c++) {
      const cell = row.getCell(c);
      if (hex) {
        cell.fill = solidFill('FF' + hex);
        cell.font = { ...tnFont, color: { argb: textOn(hex) } };
      } else {
        cell.font = { ...tnFont };
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
      cell.font = { ...tnFont, color: { argb: PINK_TEXT } };
    };
    // §2.4 — a missing/invalid % outranks the pink review style.
    if (percentIsRedFlagged(r.percent)) {
      const cell = row.getCell(pctCol);
      cell.fill = solidFill(RED_FILL);
      cell.font = { ...tnFont, color: { argb: WHITE }, bold: true };
    } else if (percentIsFlagged(r.percent)) {
      pink(pctCol);
    }
    // Round 10 — the `In QBO Time?` flag is CONDITIONAL formatting now (below),
    // so deleting a duplicate clears the highlight and the cell falls back to
    // its estimator fill instead of staying pink forever.
    if (dollarDiffIsFlagged(r.dollarDiff)) pink(diffCol);
  }
  void dupInQbo;

  /**
   * Round 10 — live duplicate / not-clocked-in highlighting on `In QBO Time?`.
   * The rule re-evaluates in Excel, so removing one of two identical values
   * un-highlights the one that remains.
   */
  const lastRow = ws.rowCount;
  if (lastRow >= 2) {
    const col = excelColumnLetter(inQboCol);
    const range = `${col}2:${col}${lastRow}`;
    ws.addConditionalFormatting({
      ref: range,
      rules: [
        {
          type: 'expression',
          priority: 1,
          formulae: [
            `AND($${col}2<>"",OR($${col}2="NotClockedIn",COUNTIF($${col}$2:$${col}$${lastRow},$${col}2)>1))`,
          ],
          style: {
            fill: solidFill(PINK_FILL),
            font: { name: FONT, size: 10, color: { argb: PINK_TEXT } },
          },
        },
      ],
    });
  }
}

// ---------------------------------------------------------------------------
// Sheet 2 — Pivot by Crew (static replica of the Excel PivotTable)
// ---------------------------------------------------------------------------

function buildPivotSheet(wb: ExcelJS.Workbook, pivot: PivotModel, palette: ReadonlyMap<string, string>): void {
  const ws = wb.addWorksheet('Pivot by Crew', {
    views: [{ state: 'frozen', ySplit: 4 }], // freeze at A5, below the header block
    // Round 7 §7 — crew summary rows sit above their collapsible children.
    properties: { outlineProperties: { summaryBelow: false, summaryRight: true } },
  });
  const estCols = pivot.estimators; // '#N/A' (when present) is always last
  // Grand Total column index is derived from however many estimator columns
  // rendered — never assumed. Round 7 §7 — Estimator helper sits after it.
  const gtCol = 4 + estCols.length + 1;
  const estHelperCol = gtCol + 1;
  const BLUE_EMAIL = 'FF0000FF';

  ws.columns = [
    { width: 35 }, { width: 14 }, { width: 24 }, { width: 39 },
    ...estCols.map(() => ({ width: 8 })),
    { width: 8 },
    { width: 10 },
  ];

  // Round 7 §7 — Pivot by Crew stays at Calibri 12.
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
    for (let c = 1; c <= estHelperCol; c++) {
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
  // gray #N/A + Grand Total + Estimator helper. Regular weight.
  const r4 = ws.addRow(['Row Labels', 'fname', 'lname', 'jobcode_2', ...estCols, 'Grand Total', 'Estimator']);
  r4.height = 15;
  for (let c = 1; c <= 4; c++) paint(r4.getCell(c), { fill: HEADER_GRAY, color: BLACK });
  estCols.forEach((e, i) => {
    const hex = palette.get(e);
    paint(r4.getCell(5 + i), hex ? { fill: 'FF' + hex, color: textOn(hex) } : { fill: HEADER_GRAY, color: BLACK });
  });
  paint(r4.getCell(gtCol), { fill: HEADER_GRAY, color: BLACK });
  paint(r4.getCell(estHelperCol), { fill: HEADER_GRAY, color: BLACK });

  // Round 7 §7 — AutoFilter over the full header row (incl. Estimator helper).
  ws.autoFilter = {
    from: { row: 4, column: 1 },
    to: { row: 4, column: estHelperCol },
  };

  const hoursCells = (byEstimator: Record<string, number>, total: number): (number | '')[] => [
    ...estCols.map((e) => (byEstimator[e] !== undefined ? byEstimator[e] : '' as const)),
    total,
  ];

  /** Estimator initials with hours on this row (column order), for AutoFilter. */
  const estimatorsLabel = (byEstimator: Record<string, number>): string =>
    estCols.filter((e) => {
      const v = byEstimator[e];
      return v !== undefined && v !== 0;
    }).join(', ');

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
    paint(row.getCell(estHelperCol), {});
    if (kind === 'total') {
      for (let c = 3; c <= gtCol; c++) {
        row.getCell(c).border = { top: { style: 'thin' } };
      }
    }
  };

  for (const crew of pivot.crews) {
    // Crew-name row: outline level 0 summary above its children (Round 7 §7).
    const crewRow = ws.addRow([crew.crew]);
    blackBand(crewRow);
    crewRow.outlineLevel = 0;

    for (const emp of crew.employees) {
      // Painter email row = level 1; jobs + "… Total" nest under it at level 2.
      const empRow = ws.addRow([emp.username, emp.fname, emp.lname]);
      styleBodyRow(empRow, 'employee');
      empRow.outlineLevel = 1;
      // Round 7 §8 — thin blue top border on email cells in column A only.
      const aVal = empRow.getCell(1).value;
      if (typeof aVal === 'string' && aVal.includes('@')) {
        empRow.getCell(1).border = {
          top: { style: 'thin', color: { argb: BLUE_EMAIL } },
        };
      }

      for (const job of emp.jobs) {
        const jobRow = ws.addRow([
          '',
          '',
          job.customer,
          job.job,
          ...hoursCells(job.byEstimator, job.total),
          estimatorsLabel(job.byEstimator),
        ]);
        styleBodyRow(jobRow, 'job');
        jobRow.outlineLevel = 2;
      }
      const totRow = ws.addRow([
        '',
        `${emp.fname} Total`,
        '',
        '',
        ...hoursCells(emp.byEstimator, emp.total),
        '',
      ]);
      styleBodyRow(totRow, 'total');
      totRow.outlineLevel = 2;
    }
  }

  // Grand Total row: bold, gray/white on column A and every value column; B:D plain.
  // Estimator helper left blank (Round 7 §7).
  const grandRow = ws.addRow([
    'Grand Total',
    '',
    '',
    '',
    ...hoursCells(pivot.grandByEstimator, pivot.grandTotal),
    '',
  ]);
  grandRow.outlineLevel = 0;
  paint(grandRow.getCell(1), { fill: HEADER_GRAY, color: WHITE, bold: true });
  for (let c = 2; c <= 4; c++) paint(grandRow.getCell(c), {});
  for (let c = 5; c <= gtCol; c++) {
    const cell = grandRow.getCell(c);
    paint(cell, { fill: HEADER_GRAY, color: WHITE, bold: true });
    if (typeof cell.value === 'number') cell.numFmt = HOURS_FMT;
    if (c < gtCol) cell.alignment = { horizontal: 'center' };
  }
  paint(grandRow.getCell(estHelperCol), {});
}

// ---------------------------------------------------------------------------
// Sheet 3 — Pivot by Job (Round 5 §3: two-row header, outline, borders)
// ---------------------------------------------------------------------------

const RED = 'FFFF0000';
const DARK_RED = 'FFC00000';
// Round 4 §3.2-1 — WO Override uses the $ accounting variant on this tab.
const MONEY_FMT = CURRENCY_FMT;

function buildPivotByJobSheet(wb: ExcelJS.Workbook, model: PivotByJobModel): void {
  const ws = wb.addWorksheet('Pivot by Job', {
    views: [{ state: 'frozen', ySplit: 2 }], // Round 5 §3.1 — freeze at A3
    properties: { outlineProperties: { summaryBelow: false, summaryRight: true } },
  });
  const dateCount = model.dateIsos.length;
  // A · dates · Total hrs · (blank manual / QB on emp rows) · AP# · OnCal · Est · Lead · WO$ · WO% · % · Notes
  const analysisStart = 1 + dateCount + 2 + 1; // first analysis col (AP#)
  const lastCol = analysisStart + 7;
  const pctComplCol = lastCol - 1;
  // Round 7 — blank column (~width 7) between Total hrs and AP#.
  // With a 7-date week, L = On Calendar? → width 16.90 (Calibri 12).
  ws.columns = [
    { width: 40 },
    ...model.dateIsos.map(() => ({ width: 10 })),
    { width: 11 },
    { width: 7 }, // blank manual column (QB hours on employee rows only)
    { width: 10.5 },
    { width: 16.9 }, // On Calendar? (column L on a 7-date week)
    { width: 9.4 },
    { width: 43 },
    { width: 23 },
    { width: 12 },
    { width: 15.78 }, // Round 5 §3.4 — % of Completion fits its header
    { width: 18.5 },
  ];

  // Round 10 — every value on this sheet is 8pt (was 12pt, with a 17pt title).
  const baseFont = { name: FONT, size: 8 };
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
  const apColLetter = excelColumnLetter(analysisStart);
  const link = resolveTimeNaviLink(wb);

  /**
   * Round 8 §3.2 — write `% of Completion` as a live cross-sheet reference to
   * this week's TimeNavi value. Manual overrides (and AP#s the calendar never
   * saw) export as VALUES: the formula would resolve to "" and lose them.
   */
  const writePercentCell = (
    row: ExcelJS.Row,
    opts: { apKey: string | null; percent: PercentValue; override: boolean },
  ): void => {
    const cell = row.getCell(pctComplCol);
    const value = percentExcelValue(opts.percent);
    const linked = !opts.override && link !== null && opts.apKey !== null && link.apKeys.has(opts.apKey);
    if (linked && link) {
      cell.value = {
        formula: timeNaviPercentFormula({
          sheetName: link.sheetName,
          apColumn: link.apColumn,
          pctColumn: link.pctColumn,
          apCellRef: `${apColLetter}${row.number}`,
        }),
        result: value,
      };
      cell.numFmt = PCT_FMT; // the linked value is a fraction (text states ignore it)
      return;
    }
    cell.value = value;
    if (typeof value === 'number') cell.numFmt = PCT_FMT;
  };

  /** AP# stays text so the MATCH against TimeNavi Export resolves. */
  const textApCell = (row: ExcelJS.Row): void => {
    row.getCell(analysisStart).numFmt = TEXT_FMT;
  };

  const centerAnalysis = (row: ExcelJS.Row): void => {
    for (let c = analysisStart; c < lastCol; c++) {
      row.getCell(c).alignment = { horizontal: 'center' };
    }
  };

  /**
   * Round 10 — job separators. The old per-row top+bottom rules over A:I made
   * the sheet look ruled; a job block now gets ONE line above it and ONE below
   * it, drawn across every column, so the color-coded blocks read as blocks.
   */
  const edgeBorder = (row: ExcelJS.Row, edge: 'top' | 'bottom'): void => {
    for (let c = 1; c <= lastCol; c++) {
      const prev = row.getCell(c).border ?? {};
      row.getCell(c).border = { ...prev, [edge]: { style: 'thin' } };
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
      if (c === 1) paint(r1.getCell(1), { fill: BLACK, color: WHITE, bold: true });
      if (c > 1 && c <= 1 + dateCount && r1.getCell(c).value instanceof Date) {
        // Round 6 — short date in B1–H1 (no year), e.g. 6/27.
        r1.getCell(c).numFmt = 'm/d';
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
      // Round 10 — the analysis block (AP# → Notes, columns K–R on a 7-date
      // week) is bold on every row.
      const boldCell = bold || c >= analysisStart;
      paint(cell, { ...(hex ? { fill: 'FF' + hex, color: textOn(hex) } : {}), bold: boldCell });
      if (typeof cell.value === 'number' && c <= qbCol && c > 1) cell.numFmt = HOURS_FMT;
    }
  };

  const analysisFormats = (row: ExcelJS.Row): void => {
    const woCell = row.getCell(analysisStart + 4);
    // Always apply $-accounting so blank cells format when the user types.
    if (woCell.value !== '-') woCell.numFmt = MONEY_FMT;
    for (const c of [analysisStart + 5, analysisStart + 6]) {
      const cell = row.getCell(c);
      if (typeof cell.value === 'number') cell.numFmt = PCT_FMT;
    }
  };

  for (const cust of model.customers) {
    const cRow = ws.addRow([cust.customer, ...hourVals(cust.byDate), cust.total]);
    cRow.outlineLevel = 0;
    fillRow(cRow, cust.colorHex, true);
    // Round 10 — one line above the customer's first job block.
    edgeBorder(cRow, 'top');
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
        percentExcelValue(a.percent),
        '', // Round 7 — Notes data cleared; keep centered "Notes" header only
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
      writePercentCell(jRow, { apKey: a.apKey, percent: a.percent, override: a.percentOverride });
      textApCell(jRow);
      centerAnalysis(jRow);

      let lastRowOfJob: ExcelJS.Row = jRow;
      for (const emp of job.employees) {
        const eRow = ws.addRow([emp.username, ...hourVals(emp.byDate), emp.total, emp.qbHours ?? '']);
        eRow.outlineLevel = 2;
        fillRow(eRow, job.colorHex);
        const qb = eRow.getCell(qbCol);
        if (typeof qb.value === 'number') {
          qb.numFmt = HOURS_FMT;
          qb.font = { ...baseFont, color: { argb: DARK_RED } };
        }
        lastRowOfJob = eRow;
      }
      // Round 10 — one line closing the job block, across every column.
      edgeBorder(lastRowOfJob, 'bottom');
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
      // §4.1 — an excluded carried job is painted A6A6A6 across the pair, so
      // next week's reader picks the signal back up from the fill.
      const rowHex = r.excludedFromPayroll ? EXCLUSION_GRAY_HEX : r.colorHex;
      const cRow = ws.addRow([r.customer]);
      cRow.outlineLevel = 0;
      fillRow(cRow, rowHex, true);
      const jRow = ws.addRow([r.job]);
      jRow.outlineLevel = 1;
      const cells: (string | number)[] = [
        r.apKey ?? '',
        '',
        r.estimator,
        r.crewLead,
        r.woOverride ?? '',
        r.woPct ?? '',
        percentExcelValue(r.percent),
        '', // Round 7 — clear BALANCES/UT Notes data; keep header label
      ];
      cells.forEach((v, i) => {
        jRow.getCell(analysisStart + i).value = v;
      });
      fillRow(jRow, rowHex);
      paint(jRow.getCell(1), { ...(rowHex ? { fill: 'FF' + rowHex, color: textOn(rowHex) } : {}), bold: true });
      const notesCell = jRow.getCell(lastCol);
      paint(notesCell, {
        ...(rowHex ? { fill: 'FF' + rowHex } : {}),
        color: rowHex && isDarkColor(rowHex) ? 'FFFFC7C7' : DARK_RED,
      });
      analysisFormats(jRow);
      writePercentCell(jRow, { apKey: r.apKey, percent: r.percent, override: r.percentOverride });
      textApCell(jRow);
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

function buildPayrollPreviewSheet(
  wb: ExcelJS.Workbook,
  model: PayrollPreviewModel,
  // Round 8 §7.3 — `Payroll Preview - UT` renders through this exact writer.
  sheetName = 'PAYROLL PREVIEW',
  inputRefsByLeadKey?: ReadonlyMap<string, { hourlyCell: string; deductCell: string }>,
): void {
  const ws = wb.addWorksheet(sheetName);
  /** Round 10 — Travel is a UT-only row. */
  const isUtSheet = model.entity === 'UT';
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
      // Round 6 §6.1 — bold `(-) taxes` / `NEST [IRA]` / `net` labels on the LAST job line.
      if (isLast) {
        r.getCell(6).value = '(-) taxes';
        r.getCell(7).value = 'NEST [IRA]';
        r.getCell(8).value = 'net';
        for (const c of [6, 7, 8]) {
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
    sums.getCell(3).value = { formula: `SUM(C${firstLineRow}:C${lastLineRow})` };
    sums.getCell(5).value = { formula: `SUM(E${firstLineRow}:E${lastLineRow})` };
    // Round 6 §6.1 — taxes / NEST / net blank by default (independent manual cells).
    sums.getCell(6).value = block.taxes ?? '';
    sums.getCell(7).value = block.nest ?? '';
    sums.getCell(8).value = block.net ?? '';
    for (const c of [3, 5, 6, 7, 8]) {
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
    // Round 9 §19/§23 — extra rows under an hourly block's totals.
    if (block.extras.length > 0) {
      const extraRows = block.extras.map((x) => ({
        extra: x,
        row: ws.addRow([x.label, '', '', '', x.amount ?? '']).number,
      }));
      for (const r of extraRows) {
        ws.getCell(`E${r.row}`).numFmt = ACCOUNTING_FMT;
        ws.getCell(`E${r.row}`).alignment = center;
        ws.getCell(`E${r.row}`).font = {
          ...baseFont,
          ...(r.extra.kind === 'deduction'
            ? { color: { argb: DEDUCT_RED } }
            : r.extra.kind === 'reimbursement'
              ? { color: { argb: REIMB_BLUE } }
              : {}),
        };
      }
      const moneyExtras = extraRows.filter((r) => r.extra.kind !== 'note');
      if (moneyExtras.length > 0) {
        const terms = moneyExtras
          .map((r) => `${r.extra.kind === 'deduction' ? '-' : '+'}N(E${r.row})`)
          .join('');
        const netRow = ws.addRow(['Total Pay']);
        netRow.getCell(5).value = { formula: `E${sums.number}${terms}` };
        netRow.getCell(5).numFmt = ACCOUNTING_FMT;
        netRow.getCell(5).alignment = center;
        netRow.font = { ...baseFont, bold: true };
      }
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
    const woRow = ws.addRow([
      'Work Order Override',
      ...block.jobs.map((j) => (j.workOrderOverride !== null && j.workOrderOverride !== undefined ? j.workOrderOverride : '')),
    ]).number;
    // Round 8 §8.1 — SUBBED / SET RATE show the state instead of a %; their
    // chain cells stay blank so the payroll clerk types the pay in.
    const pctRow = ws.addRow([
      '% Completion',
      ...block.jobs.map((j) => (j.manualPay ? j.manualPay : j.percentCompletion)),
    ]).number;
    const laborRow = ws.addRow(['- Labor + Wages', ...block.jobs.map((j) => j.laborWages)]).number;
    const milesRow = ws.addRow(['-Mileage', ...block.jobs.map((j) => j.mileage)]).number;
    // Round 9 §32 — a full-materials lead is charged all of it, and the row says so.
    // Round 10 — a full-materials band reads "- ALL Materials".
    const materialsLabel = block.jobs.some((j) => j.fullMaterials) ? '- ALL Materials' : '-½ Materials';
    const halfRow = ws.addRow([materialsLabel, ...block.jobs.map((j) => j.halfMaterials)]).number;
    /**
     * Round 10 §UT — the painter pays their own lodging, so the job's travel is
     * charged back here. `halfMaterials` already carries the hotel amount into
     * the chain, so this row is the visible half of that deduction.
     */
    const hotelsRow = block.jobs.some((j) => j.hotels !== null)
      ? ws.addRow(['- Hotels', ...block.jobs.map((j) => j.hotels ?? '')]).number
      : 0;
    const tpsRow = ws.addRow(['Total Profit Share']).number;
    let splitRow = 0;
    if (block.jobs.some((j) => j.splitWith > 1)) {
      splitRow = ws.addRow(['split w/', ...block.jobs.map((j) => (j.splitWith > 1 ? j.splitWith : ''))]).number;
    }
    ws.addRow([]);
    const totalRow = ws.addRow(['Total']).number;
    // Round 10 — one column per job, like every other chain row.
    const mileNewRow = ws.addRow([
      'Mileage - New',
      block.mileageNew,
      ...block.jobs.slice(1).map(() => ''),
    ]).number;
    const hourlyPayRow = ws.addRow(['Hourly Pay']).number;
    const reimbRow = ws.addRow(['Reimbursements', block.reimbursements, block.reimbMemo]).number;
    const deductRow = ws.addRow(['Deductions/Sundries']).number;
    // Round 9 §19/§23 — configured extra rows sit directly above Total Pay so
    // the formula below can pick them up.
    const extraRows = block.extras.map((x) => ({
      extra: x,
      row: ws.addRow([x.label, x.amount ?? '']).number,
    }));
    const totalPayRow = ws.addRow(['Total Pay']).number;
    const inputRef = inputRefsByLeadKey?.get(block.key);
    if (inputRef) {
      ws.getCell(`B${hourlyPayRow}`).value = inputNumsSheetFormula(inputRef.hourlyCell);
      ws.getCell(`B${deductRow}`).value = inputNumsSheetFormula(inputRef.deductCell);
    } else {
      ws.getCell(`B${hourlyPayRow}`).value = block.hourlyPay;
      ws.getCell(`B${deductRow}`).value = block.deductions;
    }
    ws.getCell(`C${deductRow}`).value = block.deductMemo;
    ws.addRow([]);
    const incomeRow = ws.addRow(['Total Income', ...block.jobs.map((j) => j.income)]).number;
    const matRow = ws.addRow(['Materials', ...block.jobs.map((j) => j.materials)]).number;
    // Round 8 §8.5 — informational, directly beneath Materials.
    // Round 10 — Travel is UT-only and sits BELOW QB Labor-Wages-Subs.
    ws.addRow(['QB Labor-Wages-Subs', ...block.jobs.map((j) => j.laborWages)]);
    if (isUtSheet) ws.addRow(['Travel', ...block.jobs.map((j) => j.travel)]);
    // Round 10 — `-Mileage` above IS this books figure (the P&L Mileage
    // Expense for the AP#), so it points here instead of repeating the number.
    const booksMileageRow = ws.addRow(['Mileage in Books', ...block.jobs.map((j) => j.mileage)]).number;
    for (let i = 0; i < jobCount; i++) {
      const L = colLetter(1 + i);
      const cell = ws.getCell(`${L}${milesRow}`);
      cell.value = { formula: `${L}${booksMileageRow}` };
      cell.numFmt = MONEY_FMT;
    }
    const pctMatRow = ws.addRow(['% Materials']).number;

    for (let i = 0; i < jobCount; i++) {
      const L = colLetter(1 + i); // B, C, …
      const job = block.jobs[i];
      const tps = ws.getCell(`${L}${tpsRow}`);
      if (job?.manualPay) {
        // SUBBED / SET RATE — no auto share; leave the cells open for typing.
        ws.getCell(`${L}${psRow}`).value = null;
        tps.value = null;
      } else {
        ws.getCell(`${L}${psRow}`).value = { formula: `${L}${incomeRow}*${L}${hdrRow}` };
        // Round 8 §8.1 (SUPERSEDES Round 6 §6.2 / Round 7 §21) — % Completion
        // now multiplies BOTH branches of the base, not just the override.
        // Mileage stays display-only (not a term).
        tps.value = {
          // N() so the blank-override cell compares as 0 — in Excel raw text
          // sorts above every number, so a bare `>0` would take the WO branch.
          formula: `IF(N(${L}${woRow})>0,${L}${woRow}*${L}${pctRow},${L}${psRow}*${L}${pctRow})-${L}${halfRow}-${L}${laborRow}${hotelsRow > 0 ? `-N(${L}${hotelsRow})` : ''}`,
        };
      }
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
      if (hotelsRow > 0) ws.getCell(`${L}${hotelsRow}`).numFmt = MONEY_FMT;
      for (const rn of [woRow, tpsRow]) {
        ws.getCell(`${L}${rn}`).numFmt = ACCOUNTING_FMT;
      }
      // §5.6 / Round 7 — thin under the last deduction row (Hotels when the
      // band has one), i.e. the row directly above Total Profit Share.
      ws.getCell(`${L}${hotelsRow > 0 ? hotelsRow : halfRow}`).border = { bottom: { style: 'thin' } };
    }
    // Thin bottom border A:E on the last deduction row; double A:E on
    // Deductions/Sundries.
    for (let c = 1; c <= 5; c++) {
      ws.getCell(hotelsRow > 0 ? hotelsRow : halfRow, c).border = { bottom: { style: 'thin' } };
    }

    // Total: SUM across the Total Profit Share row (÷ split when present).
    const totalCell = ws.getCell(`B${totalRow}`);
    if (splitRow > 0) {
      totalCell.value = block.total; // splits make a single SUM wrong — write the value
    } else if (jobCount > 0) {
      totalCell.value = { formula: `SUM(B${tpsRow}:${colLetter(jobCount)}${tpsRow})` };
    } else {
      totalCell.value = 0;
    }
    // Round 8 §8.2 — Total + Hourly + Reimbursements − Deductions/Sundries.
    // Mileage - New is display-only, and deductions are entered positive.
    const extraTerms = extraRows
      .filter((r) => r.extra.kind !== 'note')
      .map((r) => `${r.extra.kind === 'deduction' ? '-' : '+'}N(B${r.row})`)
      .join('');
    ws.getCell(`B${totalPayRow}`).value = {
      formula: `B${totalRow}+B${hourlyPayRow}+B${reimbRow}-B${deductRow}${extraTerms}`,
    };
    for (const r of extraRows) {
      ws.getCell(`B${r.row}`).numFmt = ACCOUNTING_FMT;
      ws.getCell(`A${r.row}`).font = { ...baseFont };
      ws.getCell(`B${r.row}`).font = {
        ...baseFont,
        ...(r.extra.kind === 'deduction'
          ? { color: { argb: DEDUCT_RED } }
          : r.extra.kind === 'reimbursement'
            ? { color: { argb: REIMB_BLUE } }
            : {}),
      };
    }
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
    for (let c = 1; c <= 5; c++) {
      ws.getCell(deductRow, c).border = { bottom: { style: 'double' } };
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
    band(
      [
        `${block.name}${block.timesheetHours > 0 ? ` (${block.timesheetHours}hrs)` : ''}` +
          // Round 9 §26/§33 — subcontractors in the salary sections.
          (block.sub ? ' (subcontractor)' : ''),
      ],
      fill,
      BLACK,
    );
    // §5.4 — the date-range cell renders 9pt like the office PO cells.
    const r1 = ws.addRow([model.weekRange ? `(${model.weekRange})` : '', block.role, block.amount]);
    r1.getCell(1).font = { ...baseFont, size: 9 };
    r1.getCell(3).numFmt = ACCOUNTING_FMT;
    // Round 9 §28 — a flat labeled reimbursements row replaces mileage reimb
    // for the people configured that way (Jodi Dunn).
    // Round 10 — some people carry no mileage row at all.
    const r2 = block.reimbursementsLabel
      ? ws.addRow([block.reimbursementsLabel, '', block.reimbursements])
      : block.noMileageRow
        ? null
        : ws.addRow([`mileage reimb: ${block.name} - ${block.miles} miles`, '', block.mileageReimb]);
    if (r2) r2.getCell(3).numFmt = ACCOUNTING_FMT;
    // Round 9 §19/§23/§24/§29 — configured extra rows.
    const extraRows = block.extras.map((x) => ({
      extra: x,
      row: ws.addRow([x.label, '', x.amount ?? '']).number,
    }));
    for (const r of extraRows) {
      ws.getCell(`C${r.row}`).numFmt = ACCOUNTING_FMT;
      ws.getCell(`C${r.row}`).font = {
        ...baseFont,
        ...(r.extra.kind === 'deduction'
          ? { color: { argb: DEDUCT_RED } }
          : r.extra.kind === 'reimbursement'
            ? { color: { argb: REIMB_BLUE } }
            : {}),
      };
    }
    const extraTerms = extraRows
      .filter((r) => r.extra.kind !== 'note')
      .map((r) => `${r.extra.kind === 'deduction' ? '-' : '+'}N(C${r.row})`)
      .join('');
    const r3 = ws.addRow(['Total', '', '']);
    r3.getCell(3).value = {
      formula: `C${r1.number}${r2 ? `+C${r2.number}` : ''}${extraTerms}`,
    };
    r3.getCell(3).numFmt = ACCOUNTING_FMT;
    r3.font = { ...baseFont, bold: true };
    // Round 9 §34 — pre-filled taxes / net when the roster carries them.
    if (block.taxes !== null || block.net !== null) {
      const rt = ws.addRow(['(-) taxes', '', block.taxes ?? '']);
      rt.getCell(3).numFmt = ACCOUNTING_FMT;
      const rn = ws.addRow(['net', '', block.net ?? '']);
      rn.getCell(3).numFmt = ACCOUNTING_FMT;
      rn.font = { ...baseFont, bold: true };
    }
    ws.addRow([]);
  };

  // §5.2 — one alphabetical painter sequence (regular painters + lead blocks
  // interleaved; subcontractor blocks last), then the black section headers.
  for (const entry of model.painterSequence) {
    if (entry.kind === 'hourly') hourlyBlock(entry.block);
    // Round 10 — subcontractor salary blocks render with the subs.
    else if (entry.kind === 'salary') salaryBlock(entry.block, TAN);
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
