import * as XLSX from 'xlsx';
import { extractApKey, normalizeApInput } from './apkey';
import { suggestPercent } from './timenavi';
import type {
  CellValue,
  PrevPercent,
  SalesRow,
  SheetData,
  TimeNaviEvent,
  TimesheetParseResult,
  TimesheetRow,
  WorkbookData,
} from './types';

// ---------------------------------------------------------------------------
// Workbook reading (SheetJS handles both .xlsx and .csv transparently)
// ---------------------------------------------------------------------------

export function readWorkbook(buffer: ArrayBuffer): WorkbookData {
  const wb = readRaw(buffer);
  const sheets: SheetData[] = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    if (ws) restoreFormulaRemnants(ws);
    const raw = ws
      ? (XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: true }) as unknown[][])
      : [];
    return { name, rows: raw.map((r) => r.map(toCellValue)) };
  });
  if (sheets.length === 0) throw new Error('The file contains no sheets.');
  return { sheets };
}

/**
 * Text files (CSV) without a byte-order mark are decoded by SheetJS as
 * cp1252, which garbles UTF-8 characters like the ♦ in TimeNavi titles.
 * If the bytes are not a binary workbook and decode cleanly as UTF-8, read
 * them as a string instead.
 */
function readRaw(buffer: ArrayBuffer): XLSX.WorkBook {
  const bytes = new Uint8Array(buffer);
  const isZip = bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b; // .xlsx
  const isCfb = bytes.length >= 2 && bytes[0] === 0xd0 && bytes[1] === 0xcf; // legacy .xls
  const hasBom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  if (!isZip && !isCfb && !hasBom) {
    try {
      const text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
      return XLSX.read(text, { type: 'string', cellDates: true });
    } catch {
      // Not valid UTF-8 — fall through to the binary reader (cp1252 etc.).
    }
  }
  return XLSX.read(bytes, { type: 'array', cellDates: true });
}

/**
 * QuickBooks exports leave literal strings like "=650" or "=(B2)+(B3)" in
 * cells. When such a file round-trips through CSV, SheetJS parses them as
 * formulas with no cached value, losing the text. Put the "=…" text back so
 * parseNumberLoose sees the same thing for .csv and .xlsx inputs.
 */
function restoreFormulaRemnants(ws: XLSX.WorkSheet): void {
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith('!')) continue;
    const cell = ws[addr] as XLSX.CellObject;
    if (cell && typeof cell === 'object' && typeof cell.f === 'string' && (cell.v === undefined || cell.v === null)) {
      cell.v = `=${cell.f}`;
      cell.t = 's';
    }
  }
}

function toCellValue(v: unknown): CellValue {
  if (v == null) return null;
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  if (v instanceof Date) return v;
  return String(v);
}

// ---------------------------------------------------------------------------
// Shared low-level helpers
// ---------------------------------------------------------------------------

/**
 * Parse a date from anything the source files throw at us: real Date cells,
 * Excel serial numbers, "6/29/2026", "01/01/2024", "2026-06-23", and the same
 * with trailing times ("6/29/2026 7:00", "2026-06-23T07:00").
 */
export function parseDate(v: CellValue): Date | null {
  if (v == null || v === '' || typeof v === 'boolean') return null;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') {
    // Excel 1900-system serial number (days since 1899-12-30).
    if (v < 1 || v > 200000) return null;
    const days = Math.floor(v);
    const frac = v - days;
    const seconds = Math.round(frac * 86400);
    return new Date(1899, 11, 30 + days, 0, 0, seconds);
  }
  const s = v.trim();
  if (!s) return null;

  // M/D/YYYY (or M/D/YY) with optional time.
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[\sT]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?)?$/i.exec(s);
  if (m) {
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    let hour = m[4] !== undefined ? Number(m[4]) : 0;
    const ampm = m[7]?.toUpperCase();
    if (ampm === 'PM' && hour < 12) hour += 12;
    if (ampm === 'AM' && hour === 12) hour = 0;
    const d = new Date(year, Number(m[1]) - 1, Number(m[2]), hour, m[5] !== undefined ? Number(m[5]) : 0, m[6] !== undefined ? Number(m[6]) : 0);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  // ISO: 2026-06-23 with optional time.
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[\sT]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(s);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), m[4] !== undefined ? Number(m[4]) : 0, m[5] !== undefined ? Number(m[5]) : 0, m[6] !== undefined ? Number(m[6]) : 0);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

/**
 * Loose number parsing for QuickBooks exports.
 * - real numbers pass through
 * - "1,234.50", "$650" → number
 * - "=1000" (formula remnant of a plain number) → 1000
 * - "=(B2)+(B3)" (subtotal formula) → 'FORMULA' (caller should skip the row)
 * - anything else → null
 */
export function parseNumberLoose(v: CellValue): number | null | 'FORMULA' {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v !== 'string') return null;
  const s = v.trim();
  if (!s) return null;
  if (s.startsWith('=')) {
    const rest = s.slice(1).trim();
    if (/^-?\$?[\d,]+(?:\.\d+)?$/.test(rest)) {
      const n = Number(rest.replace(/[$,]/g, ''));
      return Number.isFinite(n) ? n : null;
    }
    return 'FORMULA';
  }
  const cleaned = s.replace(/[$,]/g, '').replace(/^\((.*)\)$/, '-$1');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function asString(v: CellValue): string {
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v).trim();
}

function numOrZero(v: CellValue): number {
  const n = parseNumberLoose(v);
  return typeof n === 'number' ? n : 0;
}

function isEmptyRow(row: CellValue[]): boolean {
  return row.every((c) => c == null || (typeof c === 'string' && c.trim() === ''));
}

function cellAt(row: CellValue[], col: number | undefined): CellValue {
  if (col === undefined) return null;
  return row[col] ?? null;
}

interface HeaderMatch {
  rowIndex: number;
  colMap: Map<string, number>;
}

const HEADER_SCAN_ROWS = 10;

/**
 * Scan the first 10 rows for a header row containing every `required` header
 * (case-insensitive, trimmed). Extra/unknown columns are ignored.
 */
function findHeaderRow(rows: CellValue[][], required: string[], optional: string[] = []): HeaderMatch | null {
  for (let i = 0; i < Math.min(HEADER_SCAN_ROWS, rows.length); i++) {
    const row = rows[i];
    if (!row) continue;
    const found = new Map<string, number>();
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (typeof v !== 'string') continue;
      const key = v.trim().toLowerCase();
      if (key && !found.has(key)) found.set(key, c);
    }
    if (required.every((h) => found.has(h))) {
      const colMap = new Map<string, number>();
      for (const h of [...required, ...optional]) {
        const c = found.get(h);
        if (c !== undefined) colMap.set(h, c);
      }
      return { rowIndex: i, colMap };
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Slot 1 — QuickBooks Time (TSheets) timesheet
// ---------------------------------------------------------------------------

const TS_REQUIRED = ['username', 'group', 'local_date', 'hours', 'jobcode_1', 'jobcode_2'];
const TS_OPTIONAL = [
  'payroll_id', 'fname', 'lname', 'number', 'local_day', 'local_start_time', 'local_end_time',
  'tz', 'billable', 'class', 'department', 'payroll item', 'service item', 'location', 'notes',
  'approved_status', 'has_flags', 'flag_types',
];

export function parseTimesheet(wbd: WorkbookData): TimesheetParseResult {
  const sheet = wbd.sheets[0];
  const hdr = findHeaderRow(sheet.rows, TS_REQUIRED, TS_OPTIONAL);
  if (!hdr) {
    throw new Error(
      'Header row not found. Expected a QuickBooks Time export with columns like username, group, local_date, hours, jobcode_1, jobcode_2 in the first 10 rows.',
    );
  }
  const rows: TimesheetRow[] = [];
  let lunchRemoved = 0;
  for (let i = hdr.rowIndex + 1; i < sheet.rows.length; i++) {
    const r = sheet.rows[i];
    if (!r || isEmptyRow(r)) continue;
    const get = (name: string): CellValue => cellAt(r, hdr.colMap.get(name));
    const username = asString(get('username'));
    const jobcode2 = asString(get('jobcode_2'));
    const payrollItem = asString(get('payroll item'));
    const hours = numOrZero(get('hours'));
    if (!username && !jobcode2 && hours === 0) continue; // stray footer/total rows
    if (/lunch/i.test(jobcode2) || /lunch/i.test(payrollItem)) {
      lunchRemoved++;
      continue;
    }
    rows.push({
      username,
      payrollId: asString(get('payroll_id')),
      fname: asString(get('fname')),
      lname: asString(get('lname')),
      number: asString(get('number')),
      group: asString(get('group')),
      localDate: parseDate(get('local_date')),
      localDay: asString(get('local_day')),
      localStartTime: parseDate(get('local_start_time')),
      localEndTime: parseDate(get('local_end_time')),
      tz: asString(get('tz')),
      hours,
      jobcode1: asString(get('jobcode_1')),
      jobcode2,
      billable: asString(get('billable')),
      klass: asString(get('class')),
      department: asString(get('department')),
      payrollItem,
      serviceItem: asString(get('service item')),
      location: asString(get('location')),
      notes: asString(get('notes')),
      approvedStatus: asString(get('approved_status')),
      hasFlags: asString(get('has_flags')),
      flagTypes: asString(get('flag_types')),
    });
  }
  if (rows.length === 0) throw new Error('No timesheet rows found under the header row.');
  return { rows, lunchRemoved };
}

// ---------------------------------------------------------------------------
// Slot 2 — TimeNavi calendar export
// ---------------------------------------------------------------------------

const TN_REQUIRED = ['event', 'start date'];
const TN_OPTIONAL = ['end date', 'calendar', 'attendees', 'location', 'description'];

export function parseTimeNavi(wbd: WorkbookData): TimeNaviEvent[] {
  // Read the sheet named "Events" if present, otherwise the first sheet.
  // Other sheets (Sheet1 of removed events, Summary metadata) are ignored.
  const sheet = wbd.sheets.find((s) => s.name.trim().toLowerCase() === 'events') ?? wbd.sheets[0];
  const hdr = findHeaderRow(sheet.rows, TN_REQUIRED, TN_OPTIONAL);
  if (!hdr) {
    throw new Error(
      `Header row not found on sheet "${sheet.name}". Expected TimeNavi columns: Event, Start Date, End Date, Calendar, Attendees, Location, Description.`,
    );
  }
  const events: TimeNaviEvent[] = [];
  for (let i = hdr.rowIndex + 1; i < sheet.rows.length; i++) {
    const r = sheet.rows[i];
    if (!r || isEmptyRow(r)) continue;
    const get = (name: string): CellValue => cellAt(r, hdr.colMap.get(name));
    const title = asString(get('event'));
    if (!title) continue;
    const descriptionRaw = get('description');
    const description = descriptionRaw == null ? '' : String(descriptionRaw).trim();
    const suggestion = suggestPercent(description);
    events.push({
      index: events.length,
      event: title,
      startDate: parseDate(get('start date')),
      endDate: parseDate(get('end date')),
      calendar: asString(get('calendar')),
      attendees: asString(get('attendees')),
      location: asString(get('location')),
      description,
      apKey: extractApKey(title),
      suggestedPercent: suggestion.value,
      suggestedSourceLine: suggestion.line,
    });
  }
  if (events.length === 0) throw new Error(`No events found on sheet "${sheet.name}".`);
  return events;
}

// ---------------------------------------------------------------------------
// Slots 3 & 5 — sales / transaction detail reports (NV & UT)
// ---------------------------------------------------------------------------

/**
 * @param combinedHeader 'customer' for the NV Sales by Rep Detail,
 *                       'name' for the UT Transaction List by Date.
 */
export function parseSalesDetail(wbd: WorkbookData, combinedHeader: 'customer' | 'name'): SalesRow[] {
  const sheet = wbd.sheets[0];
  const hdr = findHeaderRow(sheet.rows, ['date', combinedHeader, 'sales rep'], ['name 1', 'name 2']);
  if (!hdr) {
    throw new Error(
      `Header row not found. Expected columns Date, ${combinedHeader === 'customer' ? 'Customer' : 'Name'}, Sales Rep in the first 10 rows.`,
    );
  }
  const out: SalesRow[] = [];
  for (let i = hdr.rowIndex + 1; i < sheet.rows.length; i++) {
    const r = sheet.rows[i];
    if (!r || isEmptyRow(r)) continue;
    const get = (name: string): CellValue => cellAt(r, hdr.colMap.get(name));
    const combined = asString(get(combinedHeader));
    let name1 = asString(get('name 1'));
    let name2 = asString(get('name 2'));
    if (!combined && !name1 && !name2) continue;
    if (/^total\b/i.test(combined)) continue; // QBO subtotal/total rows
    if (!name1 && !name2) {
      // Manual process adds Name 1/Name 2 via Text-to-Columns; derive them here.
      const idx = combined.indexOf(':');
      if (idx >= 0) {
        name1 = combined.slice(0, idx).trim();
        name2 = combined.slice(idx + 1).trim();
      } else {
        name1 = combined;
        name2 = combined;
      }
    }
    const apKey = extractApKey(name2) ?? extractApKey(combined);
    out.push({
      apKey,
      date: parseDate(get('date')),
      name1,
      name2,
      salesRep: asString(get('sales rep')),
      rowIndex: out.length,
    });
  }
  if (out.length === 0) throw new Error('No data rows found under the header row.');
  return out; // file order preserved: first match = earliest transaction
}

// ---------------------------------------------------------------------------
// Slots 4 & 6 — Sales by Customer Summary (hierarchical, NV & UT)
// ---------------------------------------------------------------------------

export function parseCustomerSummary(wbd: WorkbookData): Map<string, number> {
  const sheet = wbd.sheets[0];
  // The header row has "Total" in it (usually column B). Title rows before it
  // ("AllPro Painters LLC", "Sales by Customer Summary", the date range) never
  // consist of the exact word "Total".
  let headerRow = -1;
  let totalCol = -1;
  for (let i = 0; i < Math.min(HEADER_SCAN_ROWS, sheet.rows.length); i++) {
    const r = sheet.rows[i];
    if (!r) continue;
    for (let c = 0; c < r.length; c++) {
      const v = r[c];
      if (typeof v === 'string' && v.trim().toLowerCase() === 'total') {
        headerRow = i;
        totalCol = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow < 0) {
    throw new Error('Header row not found. Expected a Sales by Customer Summary with a "Total" column header in the first 10 rows.');
  }
  const map = new Map<string, number>();
  for (let i = headerRow + 1; i < sheet.rows.length; i++) {
    const r = sheet.rows[i];
    if (!r) continue;
    const rawName = r[0];
    const name = rawName == null ? '' : String(rawName);
    const trimmed = name.trim();
    if (!trimmed) continue;
    if (/^total\b/i.test(trimmed)) continue; // "Total <customer>" subtotal rows
    if (!/^AP/i.test(trimmed)) continue; // customer rows (no leading AP…)
    const apKey = extractApKey(trimmed);
    if (!apKey) continue;
    const val = parseNumberLoose(cellAt(r, totalCol));
    if (val === 'FORMULA' || val === null) continue;
    if (!map.has(apKey)) map.set(apKey, val);
  }
  if (map.size === 0) throw new Error('No job rows (names starting with "AP…") with totals were found.');
  return map;
}

// ---------------------------------------------------------------------------
// Slot 7 — last week's finished "% of Completion" workbook (optional)
// ---------------------------------------------------------------------------

export function parseLastWeek(wbd: WorkbookData): Map<string, PrevPercent> {
  // Read only the "TimeNavi Export" sheet. A single-sheet file (e.g. a CSV
  // export of that sheet) is accepted as-is.
  const sheet =
    wbd.sheets.find((s) => s.name.trim().toLowerCase() === 'timenavi export') ??
    (wbd.sheets.length === 1 ? wbd.sheets[0] : undefined);
  if (!sheet) {
    throw new Error('No sheet named "TimeNavi Export" found — is this last week\'s finished workbook?');
  }
  // Prefer real headers; fall back to fixed columns H (AP#) / I (% Completion).
  let apCol = 7;
  let pctCol = 8;
  let startRow = 1;
  const hdr = findHeaderRow(sheet.rows, ['ap#'], ['% completion']);
  if (hdr) {
    apCol = hdr.colMap.get('ap#') ?? 7;
    pctCol = hdr.colMap.get('% completion') ?? apCol + 1;
    startRow = hdr.rowIndex + 1;
  }
  const map = new Map<string, PrevPercent>();
  for (let i = startRow; i < sheet.rows.length; i++) {
    const r = sheet.rows[i];
    if (!r || isEmptyRow(r)) continue;
    const apRaw = cellAt(r, apCol);
    if (apRaw == null || apRaw === '') continue;
    const key = normalizeApInput(String(apRaw));
    if (!key) continue;
    const v = cellAt(r, pctCol);
    let val: PrevPercent | null = null;
    if (typeof v === 'number') {
      val = v > 1 ? v / 100 : v;
    } else if (typeof v === 'string') {
      const t = v.trim();
      if (/hourly/i.test(t)) {
        val = 'HOURLY';
      } else {
        const m = /^(\d+(?:\.\d+)?)\s*%?$/.exec(t);
        if (m) {
          const n = Number(m[1]);
          val = n > 1 ? n / 100 : n;
        }
      }
    }
    if (val !== null && !map.has(key)) map.set(key, val);
  }
  if (map.size === 0) {
    throw new Error('No AP# / % Completion pairs found on the "TimeNavi Export" sheet.');
  }
  return map;
}
