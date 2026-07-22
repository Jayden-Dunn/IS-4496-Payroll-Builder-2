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
  /** Round 6 §2.1 — which timesheet upload produced this row. */
  sourceEntity: 'NV' | 'UT';
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
// Slot 8 — P&L "Profit and Loss by Customer" (QBO, wide layout)
// ---------------------------------------------------------------------------

export interface PlJobMetrics {
  income: number; // 'Total Income'
  materials: number; // equipment + job materials + other job costs + paint + sundries + tools
  laborWagesSubs: number; // Employee Labor + Subcontractor Labor + Wages
  mileage: number; // Mileage Expense
  sundries: number; // the 'Sundries -' row alone (double-count hint, §6.8)
  plJobHeader: string; // the raw column header text
  /** Round 5 §4 step 5 — nearest preceding non-Total, non-AP customer header. */
  customerName: string;
}

export interface PlParseResult {
  byJob: Map<string, PlJobMetrics>;
  /** apKeys that appeared in more than one column (values summed). */
  duplicateColumns: string[];
}

// ---------------------------------------------------------------------------
// Slot 9 — Employee Name List (QB Time directory)
// ---------------------------------------------------------------------------

export interface DirectoryEmployee {
  fullName: string; // "First M Last", single-spaced, '*' stripped
  firstName: string;
  lastName: string;
  active: boolean; // false when the raw name started with '*'
}

/** Round 4 §4.2 — the employee directory persisted to localStorage. */
export interface SavedDirectory {
  employees: DirectoryEmployee[];
  sourceFilename: string;
  uploadedAt: string; // ISO timestamp
  activeCount: number;
  inactiveCount: number;
}

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

export interface LegendEntry {
  estimator: string;
  color: string; // 6-digit hex, no '#'
}

/** Crew profit-share map entry ({crew lead → WO%}, §7.1). Pct is a fraction. */
export interface ProfitShareEntry {
  lead: string;
  pct: number; // 0.38 = 38%
}

/** Wage-rate map entry (§7.2). */
export interface WageRateEntry {
  name: string;
  rate: number;
  /** Optional different rate on hourly (T&M) jobs. */
  hourlyRate?: number;
}

export type SalarySection = 'ADMIN' | 'ESTIMATORS' | 'MANAGER SALARY';

/** Salaried roster entry (§7.3). */
export interface SalariedEntry {
  name: string;
  section: SalarySection;
  role: string;
  amount: number; // weekly
}

export interface AppSettings {
  legend: LegendEntry[];
  /** Pivot estimator column order (separate from legend row-sort order). */
  pivotColumnOrder: string[];
  profitShare: ProfitShareEntry[];
  defaultProfitSharePct: number; // fraction; seed 0.38
  wageRates: WageRateEntry[];
  defaultWageRate: number; // seed 18
  salariedRoster: SalariedEntry[];
  mileageRate: number; // $/mile; seed 0.625
  /** Round 6 §7.3 — office mileage rate (PS footer); seed 0.50. */
  officeMileageRate: number;
  subcontractors: string[]; // §7.4 — valid crew leads without directory matches
  /** Round 4/6 — Davis-Bacon job patterns (Settings UI label). */
  prevailingWagePatterns: string[];
  /** Round 4/6 — Davis-Bacon minimum hourly rate; seed 19.18. */
  prevailingMinRate: number;
  /** Round 6 §7.2 — negotiated job-specific $/hr rules. */
  jobRateRules: JobRateRule[];
}

/** Round 6 §7.2 — job-specific pay rule. */
export interface JobRateRule {
  id: string;
  apKey: string;
  rate: number;
  /** Empty = all painters. */
  employeeNames: string[];
  includeSuffixes: boolean;
  note: string;
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

// ---------------------------------------------------------------------------
// Crew-lead resolution (§3.2)
// ---------------------------------------------------------------------------

export type CrewLeadFlagKind =
  | 'noMatch'
  | 'ambiguous'
  | 'notACrewLead'
  | 'inactiveMatch';

export interface CrewLeadResolution {
  /** ALL-CAPS resolved display value ("CARLOS RUIZ CRUZ" or "A/B" joins). */
  display: string;
  /** Individual resolved names (ALL CAPS); empty when nothing matched. */
  names: string[];
  /** >1 when the parenthetical listed several leads ("Ted/Rafael"). */
  split: number;
  /** True when the lead resolved through the Subcontractors list. */
  sub: boolean;
  /** Raw parenthetical text, for exception messages/tooltips. */
  raw: string;
  flags: { kind: CrewLeadFlagKind; candidate: string }[];
}

// ---------------------------------------------------------------------------
// Pivot by Job (Round 3 §4)
// ---------------------------------------------------------------------------

/** Per-job analysis fields shown on job rows (all editable). */
export interface JobAnalysis {
  apKey: string | null;
  /** 'NOT ON CALENDAR' | '-' | 'TimeNavi row {n}' */
  onCalendar: string;
  estimator: string;
  crewLead: string; // resolved display, editable
  crewLeadResolution: CrewLeadResolution | null;
  /** Manual WO Override $ — blank by default (Round 6 §5.3). */
  woOverride: number | null;
  /** Calendar description-$ suggestion (tooltip only; not written into the cell). */
  woOverrideSuggestion: number | null;
  woOverrideSuggestionSource: string | null;
  /** Manual WO% — blank by default (Round 6 §5.3). */
  woPct: number | null;
  /** Calendar @ NN% PS suggestion (tooltip only). */
  woPctSuggestion: number | null;
  woPctSuggestionSource: string | null;
  percent: PercentValue; // live mirror of the TimeNavi grid value
  notes: string;
  hourly: boolean; // job is hourly (vs piece-rate)
  sub: boolean; // paid to a subcontractor
  split: number; // co-lead count (≥1)
  /** True when job text matches a Davis-Bacon pattern (before any override). */
  dbPatternMatch: boolean;
  /** Effective DB flag: pattern match and not forced off via `dbOverride`. */
  isDbJob: boolean;
  /** True when the user forced non-DB on a pattern-matched job (§7.1). */
  dbOverride: boolean;
}

export interface PbjEmployeeRow {
  username: string;
  byDate: Record<string, number>; // ISO date → hours
  total: number;
  qbHours: number | null; // column K override (editable)
}

export interface PbjJobRow {
  /** Stable id for edits: `${customer}\u0000${jobcode2}`. */
  id: string;
  jobcode2: string;
  byDate: Record<string, number>;
  total: number;
  analysis: JobAnalysis;
  employees: PbjEmployeeRow[];
  colorHex: string | null; // estimator fill for the block
}

export interface PbjCustomerBlock {
  customer: string;
  byDate: Record<string, number>;
  total: number;
  jobs: PbjJobRow[];
  colorHex: string | null; // first job's color
}

/** Manual BALANCES / UT PAYROLL section row (§4.6). */
export interface ManualSectionRow {
  id: string;
  customer: string;
  job: string;
  apKey: string | null;
  estimator: string;
  crewLead: string;
  woOverride: number | null;
  woPct: number | null;
  percent: PercentValue;
  notes: string;
  colorHex: string | null;
}

export interface PivotByJobModel {
  title: string;
  dates: Date[]; // distinct local_dates ascending
  dateIsos: string[]; // same, as ISO keys
  customers: PbjCustomerBlock[];
  grandByDate: Record<string, number>;
  grandTotal: number;
  balances: ManualSectionRow[];
  utPayroll: ManualSectionRow[];
}

/** Sparse user edits on Pivot by Job rows (keyed by PbjJobRow.id). */
export interface PbjJobEdit {
  apKey?: string | null;
  crewLead?: string;
  woOverride?: number | null;
  woPct?: number | null;
  /** Seeded from completed %'s (Slot A) or last-week carry; otherwise mirrors TimeNavi. */
  percent?: PercentValue;
  notes?: string;
  /** Round 6 §7.1 — when true, force non–Davis-Bacon even if the job text matches a pattern. */
  dbOverride?: boolean;
}

/** QB Hrs overrides keyed by `${jobRowId}\u0000${username}`. */
export type QbHoursMap = Record<string, number>;

export type PbjEditsMap = Record<string, PbjJobEdit>;

// ---------------------------------------------------------------------------
// Payroll Preview (Round 3 §5)
// ---------------------------------------------------------------------------

/** One job line inside an hourly (T-SHEETS) block. */
export interface HourlyLine {
  id: string; // for edits: employee + jobRowId (+ ':OT')
  apKey: string | null;
  /** PO# cell text override (office lines: the "(M/D - M/D)" range or "Allpro Office"). */
  poText: string | null;
  jobLabel: string; // display name (+ ' [hrly]' / ' [OT]' tags)
  hours: number;
  rate: number;
  rateDefaulted: boolean; // true when no configured wage rate was found
  /** True when the job matched a Davis-Bacon pattern (§7.1 Branch 1). */
  prevailing: boolean;
  /** True when Branch 1 fired but the rendered rate is below the DB minimum. */
  belowPrevailing: boolean;
  /** Round 6 §7.2 — rate came from a job-specific rule (pre–Branch-1). */
  fromJobRule?: boolean;
  /** Optional note from the matched job rule. */
  ruleNote?: string;
  ot: boolean;
  hourlyJob: boolean;
  colorHex: string | null; // estimator fill for the PO# cell
}

export interface HourlyBlock {
  key: string; // employee username or full name
  title: string; // "{Full Name} ({group})" or office variants
  employeeName: string; // display name (single-letter tokens stripped)
  crewTag: string; // timesheet `group` label verbatim ('' for office variants)
  lines: HourlyLine[];
  totalHours: number;
  totalPay: number;
  /** Round 6 §6.1 — independent manual cells; blank by default. */
  taxes: number | null;
  nest: number | null;
  net: number | null;
  inactive: boolean;
}

/** One job column inside a profit-share block. */
export interface PsJobColumn {
  jobRowId: string;
  apKey: string | null;
  colorHex: string | null; // estimator fill for the PO/Name cells (§5.6)
  jobShort: string; // "AP24403 5321 Lambrook Dr Int/Ext" style
  customer: string;
  woPct: number | null;
  profitShare: number; // income × woPct
  workOrderOverride: number | null; // blank unless typed (Round 6 §6.2)
  percentCompletion: number; // fraction (HOURLY jobs excluded from PS blocks) — displayed, not in TPS chain
  laborWages: number;
  mileage: number; // displayed, not in TPS chain
  halfMaterials: number;
  totalProfitShare: number; // (WO Override|PS) − labor − ½materials (Round 6 §6.2)
  splitWith: number; // 1 = no split
  totalAfterSplit: number;
  /** Round 6 §7.1 Branch 2 — Davis-Bacon floor raised this lead's per-job pay. */
  dbFloorApplied?: boolean;
  /** Floor amount (hours × DB rate) when Branch 2 was evaluated. */
  dbFloor?: number;
  /** Per-job pay before Branch 2 raise (for tooltip / exception). */
  dbFloorPriorPay?: number;
  /** Timesheet hours used for the Branch 2 floor. */
  dbFloorHours?: number;
  missingFromPl: boolean;
  income: number;
  materials: number;
}

export interface PsBlock {
  key: string;
  leadName: string; // ALL CAPS (matching key)
  /** Title Case band name as stored in the directory / subs list (§5.3). */
  displayName: string;
  sub: boolean;
  jobs: PsJobColumn[];
  total: number; // Σ totalAfterSplit
  mileageNew: number; // editable, default 0
  hourlyPay: number; // own T-SHEETS gross
  reimbursements: number; // editable
  reimbMemo: string;
  deductions: number; // editable (negative)
  deductMemo: string;
  totalPay: number;
  tsheets: HourlyBlock | null; // the lead's own hourly block
}

/** Salary block (§5.4/§5.5). */
export interface SalaryBlock {
  key: string;
  name: string;
  role: string;
  amount: number;
  section: SalarySection;
  timesheetHours: number; // shown as "(Nhrs)" when > 0
  miles: number; // editable
  mileageReimb: number; // miles × rate
}

/** Round 4 §5.2 — the single A→Z painter/crew-lead sequence. */
export type PainterSectionEntry =
  | { kind: 'hourly'; block: HourlyBlock }
  | { kind: 'ps'; block: PsBlock };

/** Round 6 §4.10 — empty template block for someone without hours this week. */
export interface InactiveEmployeeBlock {
  name: string;
  displayName: string;
  /** e.g. "$25/hr" from the wage map. */
  rateNote: string;
  /** Crew-lead / subcontractor → materials + T-SHEETS blocks; else painter T-SHEETS only. */
  isLead: boolean;
  /** Directory row marked inactive (*). */
  directoryInactive: boolean;
}

export interface PayrollPreviewModel {
  title: string;
  /** "M/D - M/D" week range (office PO cells, salary date-range lines). */
  weekRange: string;
  painters: HourlyBlock[]; // non-lead timesheet employees
  leads: PsBlock[]; // crew-lead PS blocks
  /** §5.2 — painters + lead blocks interleaved A→Z by first name; subs last. */
  painterSequence: PainterSectionEntry[];
  office: HourlyBlock[];
  admin: SalaryBlock[];
  estimators: SalaryBlock[];
  managers: SalaryBlock[];
  /** §4.10 — full inactive tab (directory + absent leads/subs). */
  inactiveEmployees: InactiveEmployeeBlock[];
  inactiveZeroHours: string[]; // §5.7 — *-inactive directory names without hours (exceptions)
}

/** Sparse edits for Payroll Preview (keyed by block/line ids). */
export interface PayrollEdit {
  rate?: number;
  /** Round 6 §6.1 — manual; never auto-computed. */
  taxes?: number | null;
  nest?: number | null;
  net?: number | null;
  workOrderOverride?: number | null;
  percentCompletion?: number;
  laborWages?: number;
  mileage?: number;
  halfMaterials?: number;
  mileageNew?: number;
  reimbursements?: number;
  reimbMemo?: string;
  deductions?: number;
  deductMemo?: string;
  income?: number;
  materials?: number;
  miles?: number;
  hours?: number;
}

export type PayrollEditsMap = Record<string, PayrollEdit>;
