/**
 * The "what would the file alone produce?" build.
 *
 * `buildReportRows` and friends fold the user's edits into their output, so a
 * finished model cannot tell you what the inputs said underneath. These are
 * pure functions though, so running the whole chain again with EMPTY edit maps
 * re-exposes the computed layer. Doing that twice — once on the committed
 * uploads, once with the pending one merged in — is what lets "Update existing
 * report" name the manual edits a new file would overwrite, and stay silent
 * about everything else.
 *
 * This file deliberately mirrors the derivation chain in `App.tsx`. It calls
 * the SAME builders with the SAME arguments; only the inputs and the (empty)
 * edit maps differ. Keep the two in step: if a builder gains an argument that
 * changes its numbers, it has to be added here too, or the review will compare
 * against a model the app never shows.
 */
import { timesheetApKey } from './apkey';
import { enrichTimesheet } from './joins';
import { buildSalesByCustMirror } from './mirrorTabs';
import { directoryFromTimesheet, type LeadResolverContext } from './payroll';
import { buildPayrollPivotByJob } from './payrollPivot';
import { buildPayrollPreview, hasUtPayrollData } from './payrollPreview';
import { buildPivotByJob, EMPTY_MANUAL_STATE, type ManualSectionState } from './pivotByJob';
import { buildReportRows, lastPaidMapFromPivot, sortReportRows } from './timenavi';
import type { CompletedPercentsResult } from './completedPercents';
import type {
  AppSettings,
  DirectoryEmployee,
  PayrollPreviewModel,
  PivotByJobModel,
  PlParseResult,
  PrevPercent,
  ReportRow,
  SalesRow,
  SheetData,
  TimeNaviEvent,
  TimesheetParseResult,
} from './types';
import type { LastWeekPbjRow } from './parsers';

/** The slice of App's `ParsedData` the two rails' models are built from. */
export interface ShadowParsed {
  timesheetNv: TimesheetParseResult | null;
  timesheetUt: TimesheetParseResult | null;
  completedPercents: CompletedPercentsResult | null;
  events: TimeNaviEvent[] | null;
  nvSales: SalesRow[] | null;
  nvTotals: Map<string, number> | null;
  utSales: SalesRow[] | null;
  utTotals: Map<string, number> | null;
  lastWeek: Map<string, PrevPercent> | null;
  lastWeekPbj: LastWeekPbjRow[] | null;
  pl: PlParseResult | null;
  plUt: PlParseResult | null;
  directory: DirectoryEmployee[] | null;
  nvSummarySheet: SheetData | null;
  utSummarySheet: SheetData | null;
}

export interface ShadowDeps {
  settings: AppSettings;
  /** `parsed.directory ?? savedDirectory?.employees ?? null`, as App resolves it. */
  directory: DirectoryEmployee[] | null;
  payDate: Date | null;
  knownEstimators: ReadonlySet<string>;
  legendOrder: string[];
  /**
   * The manual BALANCES / UT PAYROLL section state. These are themselves user
   * edits, but they ADD and REMOVE rows rather than overriding a computed
   * value, so the shadow build keeps them: without them the shadow pivot would
   * be missing rows the real one has, and every edit on a hand-added row would
   * look orphaned.
   */
  balancesState?: ManualSectionState;
  utState?: ManualSectionState;
}

export interface ShadowModels {
  reportRows: ReportRow[];
  pivotByJob: PivotByJobModel | null;
  payroll: PayrollPreviewModel | null;
  payrollUt: PayrollPreviewModel | null;
}

/** Mirrors App's `combinedTimesheet`. */
function combineTimesheets(p: ShadowParsed): TimesheetParseResult | null {
  const nv = p.timesheetNv;
  const ut = p.timesheetUt;
  if (!nv && !ut) return null;
  return {
    rows: [...(nv?.rows ?? []), ...(ut?.rows ?? [])],
    lunchRemoved: (nv?.lunchRemoved ?? 0) + (ut?.lunchRemoved ?? 0),
  };
}

function leadContext(
  rows: readonly { group: string }[],
  directory: DirectoryEmployee[] | null,
  settings: AppSettings,
): LeadResolverContext {
  const crewGroups = [...new Set(rows.map((r) => r.group).filter(Boolean))];
  return {
    directory: directory ?? directoryFromTimesheet(rows as never),
    crewGroups,
    subcontractors: settings.subcontractors,
    aliases: settings.crewLeadAliases,
  };
}

/** The %'s rail: TimeNavi grid + Pivot by Job, computed values only. */
export function buildShadowPercents(p: ShadowParsed, deps: ShadowDeps): {
  reportRows: ReportRow[];
  pivotByJob: PivotByJobModel | null;
} {
  const combined = combineTimesheets(p);
  if (!p.events || !p.nvSales || !p.utSales || !p.nvTotals || !p.utTotals) {
    return { reportRows: [], pivotByJob: null };
  }

  const timesheetJobcodeByApKey = new Map<string, string>();
  for (const r of combined?.rows ?? []) {
    const key = timesheetApKey(r.jobcode2);
    if (key && !timesheetJobcodeByApKey.has(key)) timesheetJobcodeByApKey.set(key, r.jobcode2);
  }

  const reportRows = sortReportRows(
    buildReportRows(
      p.events,
      {}, // the point of the whole exercise
      {
        nvSales: p.nvSales,
        utSales: p.utSales,
        nvTotals: p.nvTotals,
        utTotals: p.utTotals,
        lastWeekMap: lastPaidMapFromPivot(p.lastWeekPbj),
        timesheetJobcodeByApKey,
        knownEstimators: deps.knownEstimators,
      },
      deps.payDate,
    ),
    deps.legendOrder,
  );

  const includedApKeys = new Set<string>();
  const estimatorByApKey = new Map<string, string>();
  for (const r of reportRows) {
    if (!r.included || !r.apKey) continue;
    includedApKeys.add(r.apKey);
    if (r.estimator && !estimatorByApKey.has(r.apKey)) estimatorByApKey.set(r.apKey, r.estimator);
  }

  const enriched = combined
    ? enrichTimesheet(combined.rows, p.nvSales ?? [], p.utSales ?? [], includedApKeys, estimatorByApKey)
    : [];

  const pivotByJob = buildPivotByJob({
    enriched,
    reportRows,
    settings: deps.settings,
    leadCtx: leadContext(combined?.rows ?? [], deps.directory, deps.settings),
    knownEstimators: deps.knownEstimators,
    payDate: deps.payDate,
    edits: {},
    qbHours: {},
    lastWeekPbj: p.lastWeekPbj,
    balancesState: deps.balancesState ?? EMPTY_MANUAL_STATE,
    utState: deps.utState ?? EMPTY_MANUAL_STATE,
    pl: null,
    nvTotals: p.nvTotals,
    utTotals: p.utTotals,
    nvSales: p.nvSales,
    utSales: p.utSales,
    nvSalesByCust: buildSalesByCustMirror(p.nvSummarySheet),
    utSalesByCust: buildSalesByCustMirror(p.utSummarySheet),
  });

  return { reportRows, pivotByJob };
}

/** The payroll rail: Slot A + the P&Ls, computed values only. */
export function buildShadowPayroll(p: ShadowParsed, deps: ShadowDeps): {
  payroll: PayrollPreviewModel | null;
  payrollUt: PayrollPreviewModel | null;
} {
  const completed = p.completedPercents;
  if (!completed || !p.pl) return { payroll: null, payrollUt: null };

  const pbj = buildPayrollPivotByJob({ completed, settings: deps.settings });
  const timesheet = completed.timesheet;

  const estimatorByApKey = new Map<string, string>();
  const apKeys = new Set<string>();
  for (const j of completed.jobs) {
    if (!j.apKey) continue;
    apKeys.add(j.apKey);
    if (j.estimator && !estimatorByApKey.has(j.apKey)) estimatorByApKey.set(j.apKey, j.estimator);
  }

  const enriched = timesheet ? enrichTimesheet(timesheet.rows, [], [], apKeys, estimatorByApKey) : [];
  const leadCtx = leadContext(timesheet?.rows ?? [], deps.directory, deps.settings);

  // Slot A's own `QB Hrs` column — an upload value, not this rail's edit state.
  const qbHours: Record<string, number> = {};
  for (const j of completed.jobs) {
    const id = `${j.customer || '(blank)'}\u0000${j.job}`;
    for (const e of j.employees) {
      if (e.qbHours != null) qbHours[`${id}\u0000${e.username}`] = e.qbHours;
    }
  }

  const common = {
    pbj,
    enriched,
    reportRows: [] as ReportRow[],
    pl: p.pl,
    plUt: p.plUt,
    directory: deps.directory,
    settings: deps.settings,
    payDate: deps.payDate,
    edits: {},
    qbHours,
    knownEstimators: deps.knownEstimators,
    leadCtx,
    nvSales: null,
    utSales: null,
    lastWeekPbj: completed.lastWeekPbj ?? null,
    customerNameByAp: completed.customerNameByAp ?? null,
  } as const;

  const payroll = buildPayrollPreview({ ...common, entity: 'NV' });
  const payrollUt = hasUtPayrollData({ enriched, pbj, plUt: p.plUt })
    ? buildPayrollPreview({ ...common, entity: 'UT' })
    : null;

  return { payroll, payrollUt };
}

export function buildShadowModels(p: ShadowParsed, deps: ShadowDeps): ShadowModels {
  const percents = buildShadowPercents(p, deps);
  const payroll = buildShadowPayroll(p, deps);
  return { ...percents, ...payroll };
}
