// ---------------------------------------------------------------------------
// Workbook primitives (what SheetJS hands us after sanitizing)
// ---------------------------------------------------------------------------

export type CellValue = string | number | boolean | Date | null;

export interface SheetData {
  name: string;
  rows: CellValue[][];
}

export interface WorkbookData {
  sheets: SheetData[];
}

// ---------------------------------------------------------------------------
// Parsed input rows
// ---------------------------------------------------------------------------

/** One row of the QuickBooks Time (TSheets) weekly timesheet (Slot 1). */
export interface TimesheetRow {
  username: string;
  payrollId: string;
  fname: string;
  lname: string;
  number: string;
  group: string; // crew name
  localDate: Date | null;
  localDay: string;
  localStartTime: Date | null;
  localEndTime: Date | null;
  tz: string;
  hours: number;
  jobcode1: string; // customer
  jobcode2: string; // job string containing the AP number
  billable: string;
  klass: string;
  department: string;
  payrollItem: string;
  serviceItem: string;
  location: string;
  notes: string;
  approvedStatus: string;
  hasFlags: string;
  flagTypes: string;
}

export interface TimesheetParseResult {
  rows: TimesheetRow[];
  lunchRemoved: number;
}

/** % Completion cell: fraction (0.95), the literal HOURLY, or blank. */
export type PercentValue = number | 'HOURLY' | null;

/** Last-week "% Completion" values we can carry forward. */
export type PrevPercent = number | 'HOURLY';

/** Job Total $ when no customer-summary match. */
export type JobTotalValue = number | '#N/A';

/** One TimeNavi calendar event (Slot 2). */
export interface TimeNaviEvent {
  index: number; // stable position in the file
  event: string;
  startDate: Date | null;
  endDate: Date | null;
  calendar: string;
  attendees: string;
  location: string;
  description: string;
  apKey: string | null;
}

/** One row of a sales/transaction detail report (Slots 3 and 5), in file order. */
export interface SalesRow {
  apKey: string | null;
  date: Date | null;
  name1: string; // customer
  name2: string; // job (AP…)
  salesRep: string; // estimator initials
  rowIndex: number;
}

/** Timesheet row + the derived helper columns (mirrors workbook cols AA–AE). */
export interface EnrichedTimesheetRow extends TimesheetRow {
  apKey: string | null;
  estimator: string; // '#N/A' when unresolved
  notOnCalendar: boolean;
}

// ---------------------------------------------------------------------------
// User edits & report rows
// ---------------------------------------------------------------------------

/** Sparse per-row user overrides, keyed by ReportRow.id (apKey + row index). */
export interface RowEdit {
  apKey?: string | null;
  percent?: PercentValue;
  estimator?: string;
  apptivo?: number;
  included?: boolean;
}

export type EditsMap = Record<string, RowEdit>;

/** One row of the "TimeNavi Export" output table (columns A–S). */
export interface ReportRow {
  id: string;
  eventIndex: number;
  included: boolean;
  event: string; // A
  startDate: Date | null; // B
  endDate: Date | null; // C
  calendar: string; // D
  attendees: string; // E
  location: string; // F
  description: string; // G
  apKey: string | null; // H (editable)
  percent: PercentValue; // I (editable)
  suggestedSourceLine: string | null; // tooltip for I
  estimator: string; // J (editable)
  lastPaid: PrevPercent | null; // K
  qboJobName: string; // L
  inQboTime: string; // M
  jobTotal: JobTotalValue; // N
  jobTotalFound: boolean;
  apptivo: number | null; // O (editable; blank when N is #N/A)
  apptivoEdited: boolean;
  dollarDiff: number | null; // P = N − O (blank when N is #N/A)
  invoiceDate: Date | null; // R
  dateDiff: number | null; // S = whole days R − B
  titleInitials: string | null; // (XX) from the event title, for mismatch check
}

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

export interface LegendEntry {
  estimator: string;
  color: string; // 6-digit hex, no '#'
}

export interface AppSettings {
  legend: LegendEntry[];
  /** Pivot estimator column order (separate from legend row-sort order). */
  pivotColumnOrder: string[];
}

// ---------------------------------------------------------------------------
// Pivot by Crew model
// ---------------------------------------------------------------------------

export interface PivotJobRow {
  customer: string; // jobcode_1
  job: string; // jobcode_2
  byEstimator: Record<string, number>;
  total: number;
}

export interface PivotEmployee {
  username: string;
  fname: string;
  lname: string;
  jobs: PivotJobRow[];
  byEstimator: Record<string, number>;
  total: number;
  overtime: boolean; // grand total > 40
}

export interface PivotCrew {
  crew: string;
  employees: PivotEmployee[];
}

export interface PivotModel {
  title: string;
  estimators: string[]; // ordered column list
  crews: PivotCrew[];
  grandByEstimator: Record<string, number>;
  grandTotal: number;
}

// ---------------------------------------------------------------------------
// Exceptions panel
// ---------------------------------------------------------------------------

export interface ExceptionItem {
  message: string;
  /** When set, clicking the item scrolls to this TimeNavi Export row. */
  rowId?: string;
}

export interface ExceptionGroup {
  key: string;
  title: string;
  note?: string;
  items: ExceptionItem[];
}
