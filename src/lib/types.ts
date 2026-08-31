// ---------------------------------------------------------------------------
// Workbook primitives (what SheetJS hands us after sanitizing)
// ---------------------------------------------------------------------------

export type CellValue = string | number | boolean | Date | null;

/**
 * Round 8 §4.1 — a cell's foreground fill as the reader found it. Excel stores
 * hand-picked swatches as `theme` + `tint`, not as a literal color, so both
 * encodings are kept (see `fills.ts`).
 */
export interface CellFill {
  /** Literal color when the file carries one ("FFA6A6A6" / "A6A6A6"). */
  rgb?: string | null;
  /** Theme color index (0 = Background 1 / window). */
  theme?: number | null;
  /** Theme tint in [-1, 1]; −0.35 darkens white to A6A6A6. */
  tint?: number | null;
  patternType?: string | null;
}

export interface SheetData {
  name: string;
  rows: CellValue[][];
  /**
   * Per-cell fills aligned to `rows` — only present when the reader captured
   * styles and the sheet actually had any.
   */
  fills?: (CellFill | null)[][];
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

/**
 * Round 8 §2.3 — free-text `% Completion` value. Anything the user typed that
 * is not a percent, HOURLY, SUBBED or SET RATE is kept (and exported) verbatim.
 */
export interface PercentText {
  kind: 'text';
  text: string;
}

/**
 * % Completion cell: fraction (0.95), one of the literal states
 * (HOURLY / SUBBED / SET RATE), free text, or blank.
 *
 * Use the helpers in `format.ts` (`isPercentNumber`, `isPercentText`,
 * `percentText`, `percentIsValidState`, `percentExcelValue`) instead of
 * matching on the union members by hand.
 */
export type PercentValue = number | 'HOURLY' | 'SUBBED' | 'SET RATE' | PercentText | null;

/** The literal (non-numeric, non-text) `% Completion` states. */
export type PercentState = 'HOURLY' | 'SUBBED' | 'SET RATE';

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

/**
 * Sparse per-row user overrides, keyed by ReportRow.id (apKey + row index).
 *
 * Round 10 — every displayed column is overridable, so the whole TimeNavi
 * Export grid can be corrected on screen before the workbook is downloaded.
 * The derived columns ($ Difference, Date Difference) recompute from these.
 */
export interface RowEdit {
  apKey?: string | null;
  percent?: PercentValue;
  estimator?: string;
  apptivo?: number;
  included?: boolean;
  event?: string;
  calendar?: string;
  attendees?: string;
  location?: string;
  description?: string;
  /** Round 10 — the `Last % Paid` cell (from last week's `% Completion`). */
  lastPaid?: PrevPercent | null;
  qboJobName?: string;
  inQboTime?: string;
  jobTotal?: JobTotalValue;
  invoiceDate?: Date | null;
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
  apptivo: number | null; // O (editable; §2.2 — blank until typed, never auto-filled)
  apptivoEdited: boolean;
  dollarDiff: number | null; // P = N − O (blank while O is blank or N is #N/A)
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
  sundries: number; // the 'Sundries -' / 'Sundries' row (feeds Deductions, not materials)
  /** Round 8 §8.5 — Travel Expenses / Travel Expense; informational. */
  travel: number;
  plJobHeader: string; // the raw column header text
  /** Round 5 §4 step 5 — nearest preceding non-Total, non-AP customer header. */
  customerName: string;
  /** Round 8 §7.1 — which upload slot this column came from. */
  plEntity: EntityCode;
}

export interface PlParseResult {
  byJob: Map<string, PlJobMetrics>;
  /** apKeys that appeared in more than one column (values summed). */
  duplicateColumns: string[];
  /** Round 7 — header row as parsed (col A often blank; then customer | job | Total …). */
  columnHeaders: string[];
  /** Round 7 — full account tree in file order; values aligned to columnHeaders. */
  accountRows: { label: string; values: (number | null)[] }[];
  /** Round 8 §7.1 — the entity whose slot produced this file. */
  entity: EntityCode;
  /**
   * Round 8 §7.1 — Cost of Goods Sold rows the Settings mapping does not
   * mention, so a new QBO account can never silently drop out of the math.
   */
  unmappedCogsRows: string[];
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
  /**
   * Round 9 §18 — the person's rate on timesheet lines whose pay type is
   * `Commercial Hourly`. Falls back to their `rate` when unset.
   */
  commercialHourlyRate?: number;
  /** Round 9 §18 — the person's rate on `Wages 2` timesheet lines. */
  wages2Rate?: number;
}

/**
 * Round 9 §19/§23/§24/§29 — a person-scoped extra row on their payroll block
 * (Aflac, health insurance, gas/other reimbursements). Seeded in Settings and
 * fully editable there; the amount stays editable on the preview too.
 */
export interface PayrollLineItem {
  id: string;
  /** Person the row belongs to (token match against the block name). */
  person: string;
  /** Row label exactly as it should read on the Payroll Preview. */
  label: string;
  /** Default amount; `null` renders a blank, editable cell. */
  amount: number | null;
  /** `deduction` subtracts from the block total, `reimbursement` adds, `note` neither. */
  kind: 'deduction' | 'reimbursement' | 'note';
  /**
   * `weekly` — every payroll run.
   * `firstPayDateOfMonth` — only when the pay date falls on the 1st–7th.
   */
  frequency: 'weekly' | 'firstPayDateOfMonth';
}

/** Round 9 — one resolved extra row on a rendered payroll block. */
export interface PayrollExtraLine {
  /** Edit key: `xl:{blockKey}:{lineItemId}`. */
  key: string;
  label: string;
  amount: number | null;
  kind: 'deduction' | 'reimbursement' | 'note';
}

/**
 * Round 8 §9.1 — person-scoped rate that only applies while that person works
 * a job led by `condition.leadName`.
 */
export interface ConditionalRateRule {
  id: string;
  person: string;
  condition: { kind: 'onSameJobAsLead'; leadName: string };
  rate: number;
  note?: string;
}

/**
 * Round 8 §9.1 — person-scoped per-job rate. `apKeyPattern` is a substring
 * match against the AP key (so '24097' also covers change-order suffixes).
 * Distinct from `JobRateRule` (Round 6 §7.2), which is job-scoped.
 */
export interface PerJobRateOverride {
  id: string;
  person: string;
  apKeyPattern: string;
  rate: number;
  note?: string;
}

/**
 * Round 10 — `SUBCONTRACTORS` blocks render with the subcontractor crew-lead
 * bands (below Pepe's Painting LLC), not in their own black-header section.
 */
export type SalarySection = 'ADMIN' | 'ESTIMATORS' | 'MANAGER SALARY' | 'SUBCONTRACTORS';

/** Round 6 §2.1 / Round 8 §7.3 — which company a row belongs to. */
export type EntityCode = 'NV' | 'UT';

/** Salaried roster entry (§7.3). */
export interface SalariedEntry {
  name: string;
  section: SalarySection;
  role: string;
  amount: number; // weekly (or monthly when frequency is firstPayDateOfMonth)
  /** Round 8 §7.3 — owning entity; legacy rows migrate to 'NV'. */
  entity: EntityCode;
  /**
   * Round 9 §25 — `firstPayDateOfMonth` pays once a month, on the first pay
   * date of that month (pay date on the 1st–7th). Defaults to weekly.
   */
  frequency?: 'weekly' | 'firstPayDateOfMonth';
  /** Round 9 §26 — subcontractor: rendered last inside the section. */
  sub?: boolean;
  /** Round 10 — this person's block shows no mileage-reimbursement row. */
  noMileageRow?: boolean;
  /** Round 9 §28 — replaces the "mileage reimb" row with a flat labeled row. */
  reimbursementsLabel?: string;
  /** Round 9 §34 — pre-filled taxes / net rows for this person. */
  taxes?: number;
  net?: number;
}

/**
 * Round 8 §7.1 — P&L account labels that feed each derived metric, per entity.
 * Matching is by account-row label. Sundries is deliberately NOT part of
 * `materials` (Round 8 §8.3 supersedes Round 3 §2).
 */
export interface PlAccountMapping {
  income: string[];
  materials: string[];
  laborWagesSubs: string[];
  sundries: string[];
  travel: string[];
  mileage: string[];
}

/**
 * Round 8 §11 — the painter roster with an active flag. `note` carries display
 * metadata from the source list (crew name, `*split*`, `{Reno}`).
 */
export interface PainterRosterEntry {
  name: string;
  active: boolean;
  note?: string;
}

export interface AppSettings {
  legend: LegendEntry[];
  /** Pivot estimator column order (separate from legend row-sort order). */
  pivotColumnOrder: string[];
  profitShare: ProfitShareEntry[];
  defaultProfitSharePct: number; // fraction; seed 0.38
  wageRates: WageRateEntry[];
  defaultWageRate: number; // seed 18
  /** Round 8 §9.1 — "person X earns Y while on lead Z's job" rules. */
  conditionalRateRules: ConditionalRateRule[];
  /** Round 8 §9.1 — person + AP# rate overrides (person-scoped). */
  perJobRateOverrides: PerJobRateOverride[];
  /** Round 8 §9.2 — NV daily-OT applies only at/below this base rate; seed 18. */
  nvDailyOtRateThreshold: number;
  /** Round 8 §9.2 — hours per day before daily OT starts; seed 8. */
  dailyOtHoursThreshold: number;
  /** Round 8 §7.1 — P&L account labels per metric, per entity. */
  plAccountMapping: Record<EntityCode, PlAccountMapping>;
  /** Round 8 §11 — painter roster with Active/Inactive flags. */
  painterRoster: PainterRosterEntry[];
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
  /**
   * Round 9 §19/§23/§24/§29 — person-scoped extra rows (Aflac, health
   * insurance, gas/other reimbursements) on the Payroll Preview blocks.
   */
  payrollLineItems: PayrollLineItem[];
  /**
   * Round 9 §32 — crew leads / subcontractors responsible for ALL of a job's
   * materials instead of half. Matching is by token subset on the lead name.
   */
  fullMaterialsLeads: string[];
  /**
   * Round 9 §27 — people the NV daily-overtime rule never applies to (e.g.
   * Utah crews on the NV timesheet).
   */
  dailyOtExemptNames: string[];
  /**
   * One-time data migrations already applied to this settings blob (e.g.
   * 'r8-rf-color'). Persisted so each runs exactly once per stored settings.
   */
  appliedMigrations: string[];
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
  /** Effective % — the manual override when set, else `livePercent`. */
  percent: PercentValue;
  /**
   * Round 8 §3.1 — THIS week's TimeNavi Export `% Completion` for the same
   * AP#. Null when no included calendar row matches.
   */
  livePercent: PercentValue;
  /** True when a manual pivot edit replaced the live TimeNavi value. */
  percentOverride: boolean;
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
  /** Effective % — the manual override when set, else `livePercent`. */
  percent: PercentValue;
  /** Round 8 §3.1/§3.3 — THIS week's TimeNavi `% Completion` for this AP#. */
  livePercent: PercentValue;
  /** True when a manual edit on this row replaced the live TimeNavi value. */
  percentOverride: boolean;
  /**
   * Round 8 §4.1 — the A6A6A6 gray flag: this job must NOT flow to payroll.
   * Seeded from last week's fill, toggleable in-app, painted on export so the
   * signal round-trips into next week's seed.
   */
  excludedFromPayroll: boolean;
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
  /**
   * Round 8 §4.3 — carried AP#s whose customer resolved from neither the Sales
   * by Customer Summary uploads nor the P&L header structure.
   */
  unresolvedCarriedCustomers: string[];
}

/** Sparse user edits on Pivot by Job rows (keyed by PbjJobRow.id). */
export interface PbjJobEdit {
  apKey?: string | null;
  crewLead?: string;
  woOverride?: number | null;
  woPct?: number | null;
  /**
   * Round 8 §3.1 — MANUAL % override. Wins over the live TimeNavi value; set
   * to `undefined` to relink the cell.
   */
  percent?: PercentValue;
  /**
   * Seed from completed %'s (Slot A). Used only when this week's TimeNavi grid
   * has no row for the AP# — never outranks the live link or a manual edit.
   */
  percentSeed?: PercentValue;
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
  jobLabel: string; // display name (+ ' [Hourly]' / ' [PS]' / ' [Daily OT]' tags)
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
  /** Round 8 §9.1 — rate came from a person + AP# override. */
  fromPerJobOverride?: boolean;
  /** Round 8 §9.1 — rate came from a "person on lead Z's job" rule. */
  fromConditionalRule?: boolean;
  /** Round 8 §9.1 — the lead's own crew-lead rate on a profit-share job. */
  fromCrewLeadRate?: boolean;
  /** Weekly (>40 hrs/wk) overtime line. */
  ot: boolean;
  /**
   * Round 9 §16 — the timesheet `payroll item` this line came from
   * ("Regular Pay", "Crew Lead - Hourly Wage", "Overtime Pay", …). Blank when
   * the line was synthesized (office roll-ups, computed OT).
   */
  payItem?: string;
  /**
   * Round 9 §16 — the OT line came from the timesheet's own overtime pay
   * item rather than from the computed >40 / >8 rules.
   */
  fromTimesheetOt?: boolean;
  /** Round 8 §9.2 — NV daily-overtime line (hours past 8 in one day). */
  dailyOt?: boolean;
  /** "7/14" — the day the daily-OT hours came from. */
  dailyOtDate?: string;
  /** Full `[Daily OT — …]` tag text, reused for the cell tooltip. */
  dailyOtNote?: string;
  hourlyJob: boolean;
  colorHex: string | null; // estimator fill for the PO# cell
  /** Pivot by Job Notes, shown on the NOTES tab Job column. */
  jobNotes?: string;
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
  /** Round 9 §19/§23 — extra rows (Gas Reimb, Aflac, …) under the totals. */
  extras: PayrollExtraLine[];
  /**
   * Round 9 §23 — `totalPay` + reimbursements − deductions from `extras`.
   * Equals `totalPay` when the block has no money-moving extras.
   */
  netPay: number;
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
  /** Fraction actually used by the §8.1 chain (0 when blank/invalid). */
  percentCompletion: number;
  /** Round 8 §8.1 — the raw `% of Completion` cell, states and all. */
  percent: PercentValue;
  /** Round 8 §8.1 — SUBBED / SET RATE: no auto profit share, pay entered by hand. */
  manualPay: 'SUBBED' | 'SET RATE' | null;
  /** Round 8 §8.1 — blank/invalid % on a piece-rate job; the chain paid 0. */
  missingPercent: boolean;
  laborWages: number;
  mileage: number; // displayed, not in TPS chain
  /** Round 8 §8.5 — P&L travel, informational only. */
  travel: number;
  /** Round 8 §8.3 — P&L sundries; seeds Deductions/Sundries, not materials. */
  sundries: number;
  halfMaterials: number;
  /**
   * Round 9 §32 / Round 10 — this band is responsible for ALL of the job's
   * materials, so the row reads "- ALL Materials" and carries the full amount.
   */
  fullMaterials: boolean;
  /**
   * Round 10 — Utah only: the painter pays their own hotel costs, so the job's
   * P&L travel is charged back as a "Hotels" line under the materials row.
   * Null when the band has no hotel charge-back.
   */
  hotels: number | null;
  /**
   * Round 9 §1 — labor computed from the timesheet hours on this AP# (rate ×
   * hours for every painter on the job), which is what `laborWages` defaults
   * to. Kept for the tooltip / exceptions when a manual edit replaces it.
   */
  timesheetLabor: number;
  /** Round 8 §8.1 — base × %Completion − labor − ½materials (0 when manual). */
  totalProfitShare: number;
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
  /** Pivot by Job Notes, shown on the NOTES materials Job column. */
  notes: string;
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
  /** Round 8 §8.2 — entered POSITIVE and SUBTRACTED from Total Pay. */
  deductions: number;
  deductMemo: string;
  /** Round 8 §8.3 — Σ P&L sundries, the default this block's deductions used. */
  sundriesDefault: number;
  /** Round 8 §8.2 — Total + Hourly Pay + Reimbursements − Deductions. */
  totalPay: number;
  tsheets: HourlyBlock | null; // the lead's own hourly block
  /** Round 9 §19/§23 — extra rows (Aflac, Gas Reimb, …) above Total Pay. */
  extras: PayrollExtraLine[];
  /** Round 9 §31 — Hourly Pay + Total profit share (Final Payroll Breakdown). */
  totalGrossWages: number;
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
  /** Round 9 §26 — subcontractor row, rendered last inside its section. */
  sub: boolean;
  /** Round 10 — no mileage-reimbursement row on this block. */
  noMileageRow: boolean;
  /**
   * Round 9 §28 — when set, the mileage row is replaced by a single flat
   * editable row with this label (Jodi Dunn's "Reimbursements").
   */
  reimbursementsLabel?: string;
  /** Round 9 §28 — the flat reimbursement amount when `reimbursementsLabel` is set. */
  reimbursements: number;
  /** Round 9 §19/§23/§24/§29 — extra rows (Aflac, Health Insurance, …). */
  extras: PayrollExtraLine[];
  /** Round 9 §34 — pre-filled taxes / net rows; null hides them. */
  taxes: number | null;
  net: number | null;
  /** amount + reimbursements + extra reimbursements − extra deductions. */
  total: number;
}

/**
 * Round 4 §5.2 — the single A→Z painter/crew-lead sequence.
 * Round 10 — a `salary` entry carries a subcontractor's salary block, which
 * renders with the subcontractor bands at the end of the sequence.
 */
export type PainterSectionEntry =
  | { kind: 'hourly'; block: HourlyBlock }
  | { kind: 'ps'; block: PsBlock }
  | { kind: 'salary'; block: SalaryBlock };

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
  /** Round 8 §7.3 — which entity this model was built from. */
  entity: EntityCode;
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
  /**
   * Round 8 §9.4 — leads that resolved from input data with no crew
   * profit-share % on file, auto-added at the seeded default. Building the
   * preview stays pure: the App layer persists these to the shared settings
   * store, which is what makes the auto-add stick.
   */
  newLeadsToAdd: NewLeadDefault[];
  /**
   * Round 10 — painter-roster names that were flagged inactive but worked this
   * week. Hours make someone active, so the App layer flips the roster flag
   * and the "inactive (*)" marker never shows for them again.
   */
  activatedPainters: string[];
  /** Round 8 §9.2 — every daily-OT chunk applied this week (exceptions/tags). */
  dailyOtNotes: DailyOtNote[];
}

/** Round 8 §9.4 — a crew lead the preview auto-added to the WO% map. */
export interface NewLeadDefault {
  /** Title Case display name, as it should be stored in Settings. */
  lead: string;
  pct: number;
}

/** Round 8 §9.2 — one applied daily-OT chunk, for tags and exceptions. */
export interface DailyOtNote {
  employeeName: string;
  /** "7/14" */
  date: string;
  hours: number;
  rate: number;
  jobLabel: string;
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
  /** Round 8 §10.2 — override the lead's computed T-SHEETS hourly pay. */
  hourlyPay?: number;
  reimbursements?: number;
  reimbMemo?: string;
  deductions?: number;
  deductMemo?: string;
  income?: number;
  materials?: number;
  miles?: number;
  hours?: number;
  /** Round 9 §14 — the crew profit-share % on the band header. */
  woPct?: number | null;
  /** Round 9 §19/§23 — an extra line's amount (`xl:` keys). */
  amount?: number | null;
}

export type PayrollEditsMap = Record<string, PayrollEdit>;
