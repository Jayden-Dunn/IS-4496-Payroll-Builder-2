/**
 * Round 6 §3.1–§3.5 — input mirror models for the %'s workbook helper tabs.
 * Read-only mirrors of the uploads + the same helper columns the human keeps.
 */
import { extractApKey, timesheetApKey } from './apkey';
import type { CellValue, EnrichedTimesheetRow, SalesRow, SheetData } from './types';

/**
 * Round 11 — Sales by Rep / UT Sales by Rep column widths A–M, in screen
 * pixels as specified. The web table uses them as-is; the workbook converts
 * them to Excel's character units (see `pxToExcelWidth`).
 */
export const SALES_BY_REP_WIDTHS_PX: readonly number[] = [77, 110, 716, 357, 176, 66, 176, 121, 176, 108, 132, 132, 66];

/** Round 11 — columns J, K, L (0-based 9–11) use the Accounting format. */
export const SALES_BY_REP_ACCOUNTING_COLS: ReadonlySet<number> = new Set([9, 10, 11]);

/**
 * Pixels → Excel column width. Excel measures widths in "characters" of the
 * default font (Calibri 11: 7px per digit plus 5px of cell padding).
 */
export function pxToExcelWidth(px: number): number {
  return Math.round(((px - 5) / 7) * 100) / 100;
}

/**
 * Round 11 — timesheet columns nobody reads during the weekly run. They stay
 * in the workbook (hidden, like TimeNavi's Attendees/Location) and collapse to
 * thin stubs on screen, with the value still in the tooltip.
 */
export const TIMESHEET_COLLAPSED_HEADERS: ReadonlySet<string> = new Set([
  'number',
  'billable',
  'class',
  'department',
  'service_item',
  'has_flags',
  'flag_types',
]);

export interface SalesByRepMirrorRow {
  apKey: string;
  cols: CellValue[]; // B–M (12 upload columns in file order)
  dateCopy: CellValue; // N
  qboJobName: CellValue; // O
}

export interface SalesByCustMirrorRow {
  text: string;
  value: number | null;
  apKey: string;
  jobTotal: number | null;
}

export interface TimesheetMirrorRow {
  raw: CellValue[]; // A–X (24)
  apname: string;
  apKey: string;
  estimator: string;
  qboJobName: string;
  notOnCalendar: string; // `TimeNavi row {n}`, or NotOnCalendar
}

/** §3.1 / §3.2 — Sales by Rep mirrors from already-parsed SalesRow aren't enough
 *  for full column fidelity; build from the raw sheet when available. */
export function buildSalesByRepMirrorFromSheet(
  sheet: SheetData | null,
  kind: 'nv' | 'ut',
): { headers: string[]; rows: SalesByRepMirrorRow[] } {
  if (!sheet || sheet.rows.length === 0) return { headers: [], rows: [] };
  // Find header row containing Date + Sales Rep-ish columns.
  let headerIdx = -1;
  for (let i = 0; i < Math.min(sheet.rows.length, 30); i++) {
    const joined = sheet.rows[i].map((c) => String(c ?? '').toLowerCase()).join('|');
    if (joined.includes('date') && (joined.includes('sales rep') || joined.includes('amount'))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) headerIdx = 0;
  const headerRow = sheet.rows[headerIdx].map((c) => String(c ?? '').trim());
  // Round 7 — drop a leading empty / "Col 1" so upload cols start at Date (or Name for UT).
  const h0 = headerRow[0] ?? '';
  const dateOrNameAt1 = /^date$/i.test(headerRow[1] ?? '') || /^name$/i.test(headerRow[1] ?? '');
  let startCol = 0;
  if (!h0.trim() || /^col\s*1$/i.test(h0) || (!h0.trim() && dateOrNameAt1)) {
    startCol = 1;
  }
  const salesRepIdx = headerRow.findIndex((h) => /^sales\s*rep$/i.test(h));
  // 12 upload columns from startCol; always include Sales Rep when present in the file.
  const colIndexes: number[] = [];
  for (let i = 0; i < 12; i++) colIndexes.push(startCol + i);
  if (salesRepIdx >= 0 && !colIndexes.includes(salesRepIdx)) {
    colIndexes[colIndexes.length - 1] = salesRepIdx;
  }
  const headers = colIndexes.map((i) => headerRow[i] || `Col ${i + 1}`);
  const name2Idx = headers.findIndex((h) => /^name\s*2$/i.test(h));
  const customerIdx = headers.findIndex((h) => /^customer$/i.test(h) || /^name$/i.test(h));
  const dateIdx = headers.findIndex((h) => /^date$/i.test(h));
  const rows: SalesByRepMirrorRow[] = [];
  for (let r = headerIdx + 1; r < sheet.rows.length; r++) {
    const raw = sheet.rows[r];
    if (!raw || raw.every((c) => c === null || c === '')) continue;
    const cols: CellValue[] = colIndexes.map((i) => raw[i] ?? null);
    const name2 = name2Idx >= 0 ? String(cols[name2Idx] ?? '') : '';
    const customer = customerIdx >= 0 ? String(cols[customerIdx] ?? '') : '';
    const apKey = extractApKey(name2) ?? extractApKey(customer) ?? '';
    rows.push({
      apKey,
      cols,
      dateCopy: dateIdx >= 0 ? cols[dateIdx] : null,
      qboJobName: name2Idx >= 0 ? cols[name2Idx] : null,
    });
  }
  void kind;
  return { headers, rows };
}

/** Fallback from parsed SalesRow when raw sheet isn't kept. */
export function buildSalesByRepMirrorFromParsed(sales: SalesRow[]): {
  headers: string[];
  rows: SalesByRepMirrorRow[];
} {
  const headers = [
    'Date',
    'Customer',
    'Name 1',
    'Name 2',
    'Transaction Type',
    'Num',
    'Product/Service',
    'Memo/Description',
    'Sales Price',
    'Amount',
    'Balance',
    'Sales Rep',
  ];
  const rows: SalesByRepMirrorRow[] = sales.map((s) => {
    const cols: CellValue[] = [
      s.date,
      s.name1,
      s.name1,
      s.name2,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      s.salesRep,
    ];
    return { apKey: s.apKey ?? '', cols, dateCopy: s.date, qboJobName: s.name2 };
  });
  return { headers, rows };
}

function parseCustMoney(v: CellValue): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const m = /^-?\d+(?:\.\d+)?/.exec(v.replace(/[,$=()]/g, '').replace(/^\(/, '-'));
    if (m) return Number(m[0]);
    const n = Number(String(v).replace(/[,$]/g, '').replace(/^=/, ''));
    if (Number.isFinite(n)) return n;
  }
  return null;
}

function isBlankCustRow(row: CellValue[] | undefined): boolean {
  if (!row) return true;
  return row.every((c) => c === null || c === undefined || c === '');
}

/** True when A is empty and B is the "Total" column header (not a "Total …" subtotal). */
function isTotalHeaderRow(row: CellValue[]): boolean {
  const a = String(row[0] ?? '').trim();
  const b = String(row[1] ?? '').trim();
  return a === '' && /^total$/i.test(b);
}

export function buildSalesByCustMirror(sheet: SheetData | null): SalesByCustMirrorRow[] {
  if (!sheet) return [];
  const out: SalesByCustMirrorRow[] = [];
  let totalHeaderIdx = -1;
  for (let i = 0; i < Math.min(sheet.rows.length, 15); i++) {
    const row = sheet.rows[i];
    if (row && isTotalHeaderRow(row)) {
      totalHeaderIdx = i;
      break;
    }
  }
  // Keep title rows above the Total header (UT SalesbyCust uses them in export).
  const dataStart = totalHeaderIdx >= 0 ? totalHeaderIdx + 1 : 0;
  for (let i = 0; i < dataStart; i++) {
    if (i === totalHeaderIdx) continue;
    const row = sheet.rows[i];
    if (!row || isBlankCustRow(row)) continue;
    const text = String(row[0] ?? '');
    if (text === '' && (row[1] === null || row[1] === undefined || row[1] === '')) continue;
    if (isTotalHeaderRow(row)) continue;
    const value = parseCustMoney(row[1]);
    const apKey = extractApKey(text) ?? '';
    out.push({ text, value, apKey, jobTotal: apKey && value !== null ? value : null });
  }
  // Round 7 — skip blank row(s) immediately after the Total header so export row 2 is the first customer.
  let i = dataStart;
  while (i < sheet.rows.length && isBlankCustRow(sheet.rows[i])) i++;
  for (; i < sheet.rows.length; i++) {
    const row = sheet.rows[i];
    if (!row || isBlankCustRow(row)) continue;
    if (isTotalHeaderRow(row)) continue;
    const text = String(row[0] ?? '');
    if (text === '' && (row[1] === null || row[1] === undefined || row[1] === '')) continue;
    const value = parseCustMoney(row[1]);
    const apKey = extractApKey(text) ?? '';
    out.push({
      text,
      value,
      apKey,
      jobTotal: apKey && value !== null ? value : null,
    });
  }
  return out;
}

/**
 * §3.5 timesheets mirror helpers. `timeNaviRowByAp` maps apKey → 1-based sheet
 * row number on TimeNavi Export (header = 1).
 */
export function buildTimesheetMirror(
  enriched: EnrichedTimesheetRow[],
  timeNaviRowByAp: ReadonlyMap<string, number>,
): TimesheetMirrorRow[] {
  return enriched.map((r) => {
    const apKey =
      r.apKey === 'OFFICE' || /^office$/i.test(r.jobcode2.trim())
        ? 'OFFICE'
        : (r.apKey ?? timesheetApKey(r.jobcode2) ?? '');
    const apname = r.jobcode2.trim().split(/\s+/)[0] ?? '';
    const tnRow = apKey && apKey !== 'OFFICE' ? timeNaviRowByAp.get(apKey) : undefined;
    const cal = tnRow !== undefined ? `TimeNavi row ${tnRow}` : 'NotOnCalendar';
    const raw: CellValue[] = [
      r.username,
      r.payrollId,
      r.fname,
      r.lname,
      r.number,
      r.group,
      r.localDate,
      r.localDay,
      r.localStartTime,
      r.localEndTime,
      r.tz,
      r.hours,
      r.jobcode1,
      r.jobcode2,
      r.billable,
      r.klass,
      r.department,
      r.payrollItem,
      r.serviceItem,
      r.location,
      r.notes,
      r.approvedStatus,
      r.hasFlags,
      r.flagTypes,
    ];
    return {
      raw,
      apname,
      apKey,
      estimator: r.estimator,
      qboJobName: r.jobcode2,
      notOnCalendar: cal,
    };
  });
}
