/**
 * Round 3 §4 — "Pivot by Job": customer → job → employee hour tree with the
 * analysis columns, plus the BALANCES and UT PAYROLL manual sections that
 * auto-position below the pivot (obsoleting the Excel move-refresh-move
 * workaround).
 */
import { timesheetApKey } from './apkey';
import { buildCustomerByApResolver } from './customerByAp';
import { dateToIso, fmtDate, round2 } from './format';
import {
  carryDisplayMap,
  jobDisplayPair,
  salesDisplayMap,
  timesheetDisplayMap,
  type JobDisplaySources,
} from './jobNames';
import type { SalesByCustMirrorRow } from './mirrorTabs';
import type { LastWeekPbjRow } from './parsers';
import {
  extractLeadParenthetical,
  extractWoOverride,
  extractWoPct,
  isPrevailingWageJob,
  resolveCrewLead,
  type LeadResolverContext,
} from './payroll';
import type {
  AppSettings,
  EnrichedTimesheetRow,
  JobAnalysis,
  ManualSectionRow,
  PercentValue,
  PbjCustomerBlock,
  PbjEditsMap,
  PbjEmployeeRow,
  PbjJobRow,
  PivotByJobModel,
  PlParseResult,
  QbHoursMap,
  ReportRow,
  SalesRow,
} from './types';

/** Round 6 §5.1 — display form for a matched TimeNavi Export sheet row. */
export function formatTimeNaviRow(sheetRow: number): string {
  return `TimeNavi row ${sheetRow}`;
}

/** Parse `TimeNavi row 19` or a bare `19` into the 1-based sheet row number. */
export function parseTimeNaviRowRef(text: string): number | null {
  const t = text.trim();
  const labeled = /^TimeNavi\s+row\s+(\d+)$/i.exec(t);
  if (labeled) return Number(labeled[1]);
  if (/^\d+$/.test(t)) return Number(t);
  return null;
}

export interface ManualSectionState {
  added: ManualSectionRow[];
  deleted: string[];
  edits: Record<string, Partial<ManualSectionRow>>;
}

export const EMPTY_MANUAL_STATE: ManualSectionState = { added: [], deleted: [], edits: {} };

export interface PivotByJobInput {
  enriched: EnrichedTimesheetRow[];
  reportRows: ReportRow[]; // sorted TimeNavi grid (source of % + row refs)
  settings: AppSettings;
  leadCtx: LeadResolverContext;
  knownEstimators: ReadonlySet<string>;
  payDate: Date | null;
  edits: PbjEditsMap;
  qbHours: QbHoursMap;
  lastWeekPbj: LastWeekPbjRow[] | null;
  balancesState: ManualSectionState;
  utState: ManualSectionState;
  /** Round 5 §3.7 — P&L income for the WO Override default. */
  pl: PlParseResult | null;
  /** Round 5 §3.7 / §4 — customer-summary totals + sales for naming/WO$. */
  nvTotals: ReadonlyMap<string, number> | null;
  utTotals: ReadonlyMap<string, number> | null;
  nvSales: readonly SalesRow[] | null;
  utSales: readonly SalesRow[] | null;
  /**
   * Round 8 §4.3 — Sales by Customer Summary mirrors (already parsed by
   * `mirrorTabs`); the carried-job customer resolver reads their hierarchy.
   */
  nvSalesByCust?: readonly SalesByCustMirrorRow[] | null;
  utSalesByCust?: readonly SalesByCustMirrorRow[] | null;
}

/**
 * Round 8 §4.2 — the Reno, NV estimator's initial. Reno is a Nevada branch, so
 * his carried jobs belong to NV PAYROLL (BALANCES) no matter which calendar
 * the event came from.
 */
export const RENO_ESTIMATOR = 'RF';

/** Section a carried job lands in: the estimator outranks the source calendar. */
export function carriedSectionFor(
  defaultSection: 'balances' | 'ut',
  estimator: string | null | undefined,
): 'balances' | 'ut' {
  return (estimator ?? '').trim().toUpperCase() === RENO_ESTIMATOR ? 'balances' : defaultSection;
}

export function pbjJobId(customer: string, jobcode2: string): string {
  return `${customer}\u0000${jobcode2}`;
}

/** A job is hourly when its % says HOURLY, its lead is HOURLY, or it's tagged T&M/[hrly]. */
export function isHourlyJob(percent: JobAnalysis['percent'], crewLead: string, jobText: string): boolean {
  if (percent === 'HOURLY') return true;
  if (/^hourly$/i.test(crewLead.trim())) return true;
  return /\bT\s*&\s*M\b|\[hrly\]/i.test(jobText);
}

const byText = (a: string, b: string): number => a.localeCompare(b, 'en', { sensitivity: 'base' });

export function buildPivotByJob(input: PivotByJobInput): PivotByJobModel {
  const { enriched, reportRows, settings, leadCtx, knownEstimators, payDate, edits, qbHours } = input;
  const colorByEstimator = new Map(settings.legend.map((l) => [l.estimator, l.color]));

  // TimeNavi lookup: first included row per apKey + its 1-based sheet row
  // number (header = row 1) for the "On Calendar?" MATCH-style reference.
  const included = reportRows.filter((r) => r.included);
  const eventByApKey = new Map<string, { row: ReportRow; sheetRow: number }>();
  included.forEach((r, i) => {
    if (r.apKey && !eventByApKey.has(r.apKey)) eventByApKey.set(r.apKey, { row: r, sheetRow: i + 2 });
  });

  const nameSources: JobDisplaySources = {
    timesheetByAp: timesheetDisplayMap(enriched),
    carryByAp: carryDisplayMap(input.lastWeekPbj),
    salesByAp: salesDisplayMap(input.nvSales, input.utSales),
    pl: input.pl,
  };

  // ----- dates -----------------------------------------------------------
  const dateMs = new Set<number>();
  let hasUndated = false;
  for (const r of enriched) {
    // NV pivot date columns follow NV hours; UT section headers reuse the same axis.
    if (r.localDate) dateMs.add(r.localDate.getTime());
    else if (r.hours !== 0) hasUndated = true;
  }
  const dates = [...dateMs].sort((a, b) => a - b).map((ms) => new Date(ms));
  const dateIsos = dates.map(dateToIso);
  if (hasUndated) dateIsos.push(''); // "(blank)" column only when needed

  // ----- customer → job → employee tree ----------------------------------
  interface EmpAcc {
    username: string;
    byDate: Map<string, number>;
    total: number;
    firstIndex: number;
  }
  interface JobAcc {
    jobcode2: string;
    byDate: Map<string, number>;
    total: number;
    employees: Map<string, EmpAcc>;
    rows: EnrichedTimesheetRow[];
  }
  const custs = new Map<string, Map<string, JobAcc>>();
  enriched.forEach((r, idx) => {
    if (r.hours === 0) return;
    // Round 6 — UT timesheet hours belong under UT PAYROLL, not the NV pivot tree.
    if (r.sourceEntity === 'UT') return;
    const customer = r.jobcode1 || '(blank)';
    let jobs = custs.get(customer);
    if (!jobs) {
      jobs = new Map();
      custs.set(customer, jobs);
    }
    let job = jobs.get(r.jobcode2);
    if (!job) {
      job = { jobcode2: r.jobcode2, byDate: new Map(), total: 0, employees: new Map(), rows: [] };
      jobs.set(r.jobcode2, job);
    }
    const iso = r.localDate ? dateToIso(r.localDate) : '';
    job.byDate.set(iso, (job.byDate.get(iso) ?? 0) + r.hours);
    job.total += r.hours;
    job.rows.push(r);
    const key = r.username || `${r.fname} ${r.lname}`;
    let emp = job.employees.get(key);
    if (!emp) {
      emp = { username: key, byDate: new Map(), total: 0, firstIndex: idx };
      job.employees.set(key, emp);
    }
    emp.byDate.set(iso, (emp.byDate.get(iso) ?? 0) + r.hours);
    emp.total += r.hours;
  });

  const toRecord = (m: Map<string, number>): Record<string, number> => {
    const rec: Record<string, number> = {};
    for (const iso of dateIsos) {
      const v = m.get(iso);
      if (v !== undefined && round2(v) !== 0) rec[iso] = round2(v);
    }
    return rec;
  };

  const grandByDateMap = new Map<string, number>();
  let grandTotal = 0;

  const customers: PbjCustomerBlock[] = [...custs.entries()]
    .sort(([a], [b]) => byText(a, b))
    .map(([customer, jobsMap]) => {
      const custByDate = new Map<string, number>();
      let custTotal = 0;
      const jobs: PbjJobRow[] = [...jobsMap.values()]
        .sort((a, b) => byText(a.jobcode2, b.jobcode2))
        .map((job) => {
          const id = pbjJobId(customer, job.jobcode2);
          const edit = edits[id] ?? {};
          const defaultAp = timesheetApKey(job.jobcode2);
          const apKey = edit.apKey !== undefined ? edit.apKey : defaultAp;
          const isOffice = apKey === 'OFFICE';
          const match = apKey && !isOffice ? eventByApKey.get(apKey) : undefined;

          // Round 6 §5.1 — `TimeNavi row {n}` (MATCH semantics).
          const onCalendar = isOffice ? '-' : match ? formatTimeNaviRow(match.sheetRow) : 'NOT ON CALENDAR';
          // An edited AP# re-resolves the estimator from the matched event
          // (acceptance §10.3); otherwise the enriched-timesheet value rules.
          const baseEstimator = job.rows[0]?.estimator ?? '#N/A';
          const estimator = isOffice
            ? 'ND'
            : edit.apKey !== undefined && match
              ? match.row.estimator || baseEstimator
              : baseEstimator;
          // Round 8 §3.1 — the cell is a live reference to THIS week's TimeNavi
          // `% Completion` for the same AP#; a manual pivot edit overrides it.
          // Slot A seeds only fill in when the grid has no row (payroll-only).
          const livePercent: PercentValue = isOffice ? null : match ? match.row.percent : null;
          const percentOverride = !isOffice && edit.percent !== undefined;
          const percent: PercentValue = isOffice
            ? null
            : percentOverride
              ? (edit.percent as PercentValue)
              : match
                ? livePercent
                : (edit.percentSeed ?? null);

          // Crew lead — from the matched event title (§3.2); OFFICE → HOURLY.
          const resolution = match
            ? resolveCrewLead(extractLeadParenthetical(match.row.event, knownEstimators), leadCtx)
            : null;
          const crewLead =
            edit.crewLead !== undefined ? edit.crewLead : isOffice ? 'HOURLY' : (resolution?.display ?? '');

          const jobText = `${job.jobcode2} ${match?.row.event ?? ''}`;
          const hourly = isHourlyJob(percent, crewLead, jobText);

          // Round 6 §5.3 — WO% / WO Override blank by default; calendar patterns
          // are suggestions only (tooltips / click-to-accept). Slot A seeds via edits.
          const pctSrc = match ? extractWoPct(match.row.description) : null;
          const woSrc = match ? extractWoOverride(match.row.description) : null;
          const woPct = hourly ? null : edit.woPct !== undefined ? edit.woPct : null;
          const woOverride = hourly ? null : edit.woOverride !== undefined ? edit.woOverride : null;

          // Round 7 — notes stay blank unless the user edited; do not carry last-week / *PAID* text.
          const notes = edit.notes !== undefined ? edit.notes : '';

          // Round 6 §7.1 — DB chip + optional force-non-DB override.
          const dbPatternMatch = isPrevailingWageJob(jobText, settings.prevailingWagePatterns);
          const dbOverride = edit.dbOverride === true;
          const isDbJob = dbPatternMatch && !dbOverride;

          const analysis: JobAnalysis = {
            apKey,
            onCalendar,
            estimator: isOffice ? '-' : estimator,
            crewLead,
            crewLeadResolution: resolution,
            woOverride,
            woOverrideSuggestion: hourly ? null : (woSrc?.amount ?? null),
            woOverrideSuggestionSource: hourly ? null : (woSrc?.line ?? null),
            woPct,
            woPctSuggestion: hourly ? null : (pctSrc?.pct ?? null),
            woPctSuggestionSource: hourly ? null : (pctSrc?.line ?? null),
            percent,
            livePercent,
            percentOverride,
            notes,
            hourly,
            sub: resolution?.sub ?? false,
            split: resolution?.split ?? 1,
            dbPatternMatch,
            isDbJob,
            dbOverride,
          };

          const employees: PbjEmployeeRow[] = [...job.employees.values()]
            .sort((a, b) => a.firstIndex - b.firstIndex)
            .map((e) => ({
              username: e.username,
              byDate: toRecord(e.byDate),
              total: round2(e.total),
              qbHours: qbHours[`${id}\u0000${e.username}`] ?? null,
            }));

          for (const [iso, v] of job.byDate) {
            custByDate.set(iso, (custByDate.get(iso) ?? 0) + v);
            grandByDateMap.set(iso, (grandByDateMap.get(iso) ?? 0) + v);
          }
          custTotal += job.total;
          grandTotal += job.total;

          const colorHex = isOffice
            ? (colorByEstimator.get('ND') ?? null)
            : (colorByEstimator.get(estimator) ?? null);

          return {
            id,
            jobcode2: job.jobcode2,
            byDate: toRecord(job.byDate),
            total: round2(job.total),
            analysis,
            employees,
            colorHex,
          };
        });

      return {
        customer,
        byDate: toRecord(custByDate),
        total: round2(custTotal),
        jobs,
        colorHex: jobs[0]?.colorHex ?? null,
      };
    });

  // ----- §4.6 manual sections ---------------------------------------------
  const timesheetApKeys = new Set<string>();
  for (const c of customers) for (const j of c.jobs) if (j.analysis.apKey) timesheetApKeys.add(j.analysis.apKey);

  const balancesSeeds: ManualSectionRow[] = [];
  const utSeeds: ManualSectionRow[] = [];
  const seenBalancesAp = new Set<string>();
  const seenUtAp = new Set<string>();

  // Round 8 §4.3 — every carried job resolves its OWN customer by AP key
  // (Sales by Customer Summary first, P&L header structure second).
  const customerResolver = buildCustomerByApResolver({
    nvSalesByCust: input.nvSalesByCust,
    utSalesByCust: input.utSalesByCust,
    pl: input.pl,
  });
  const unresolvedCarriedCustomers: string[] = [];
  const resolveCarriedCustomer = (
    apKey: string | null,
    section: 'balances' | 'ut',
    fallback: string,
    { flag }: { flag: boolean } = { flag: true },
  ): string => {
    const hit = customerResolver.resolve(apKey, section === 'ut' ? 'UT' : 'NV');
    if (hit.customer) return hit.customer;
    if (flag && apKey && !unresolvedCarriedCustomers.includes(apKey)) {
      unresolvedCarriedCustomers.push(apKey);
    }
    return fallback;
  };

  /** §3.1/§3.3 — this week's TimeNavi `% Completion` for an AP#, if any. */
  const liveLink = (apKey: string | null): { linked: boolean; percent: PercentValue } => {
    const hit = apKey ? eventByApKey.get(apKey) : undefined;
    return hit ? { linked: true, percent: hit.row.percent } : { linked: false, percent: null };
  };

  const pushSeed = (section: 'balances' | 'ut', seed: ManualSectionRow): void => {
    if (section === 'balances') {
      seenBalancesAp.add(seed.apKey ?? seed.id);
      balancesSeeds.push(seed);
    } else {
      seenUtAp.add(seed.apKey ?? seed.id);
      utSeeds.push(seed);
    }
  };

  for (const row of input.lastWeekPbj ?? []) {
    // §4.2 — RF (Reno, NV) carried jobs go to NV BALANCES, never UT PAYROLL.
    const section = carriedSectionFor(row.section === 'ut' ? 'ut' : 'balances', row.estimator);
    if (seenBalancesAp.has(row.apKey) || seenUtAp.has(row.apKey)) continue;
    if (row.section !== 'ut') {
      // NV pivot/balances rows: carry when paid < 100% and no hours this week.
      const partial = typeof row.percent === 'number' && row.percent < 1;
      const open = row.section === 'balances' ? row.percent !== 1 : partial;
      if (!open || timesheetApKeys.has(row.apKey)) continue;
    }
    // Round 7 — no carried notes / *PAID* text in BALANCES/UT.
    pushSeed(
      section,
      manualRowFromCarry({
        section,
        row,
        notes: '',
        colorByEstimator,
        live: liveLink(row.apKey),
        customer: resolveCarriedCustomer(row.apKey, section, row.customer),
      }),
    );
  }

  // Seeds from this week's included Utah/Reno-calendar events (dedupe by apKey).
  for (const r of included) {
    if (!/^utah\b|^reno\b/i.test(r.calendar.trim())) continue;
    const apKey = r.apKey;
    if (!apKey || seenUtAp.has(apKey) || seenBalancesAp.has(apKey)) continue;
    const section = carriedSectionFor('ut', r.estimator);
    const resolution = resolveCrewLead(extractLeadParenthetical(r.event, knownEstimators), leadCtx);
    // Round 6 §5.3 — blank WO defaults (suggestions live on calendar-matched NV jobs only).
    const names = jobDisplayPair(apKey, { ...nameSources, eventTitle: r.event });
    pushSeed(section, {
      id: `${section}:${apKey}`,
      customer: resolveCarriedCustomer(apKey, section, names.customerName, { flag: false }),
      job: names.poLabel,
      apKey,
      estimator: r.estimator,
      crewLead: resolution.display,
      woOverride: null,
      woPct: null,
      percent: r.percent,
      livePercent: r.percent,
      percentOverride: false,
      excludedFromPayroll: false,
      notes: '',
      colorHex: colorByEstimator.get(r.estimator) ?? null,
    });
  }

  // UT timesheet (`sourceEntity: 'UT'`) jobs also seed UT PAYROLL when not already present.
  for (const r of enriched) {
    if (r.sourceEntity !== 'UT' || r.hours === 0) continue;
    const apKey = r.apKey;
    if (!apKey || apKey === 'OFFICE' || seenUtAp.has(apKey) || seenBalancesAp.has(apKey)) continue;
    seenUtAp.add(apKey);
    const match = eventByApKey.get(apKey);
    const resolution = match
      ? resolveCrewLead(extractLeadParenthetical(match.row.event, knownEstimators), leadCtx)
      : null;
    const names = jobDisplayPair(apKey, {
      ...nameSources,
      eventTitle: match?.row.event ?? null,
    });
    const live = liveLink(apKey);
    utSeeds.push({
      id: `ut:${apKey}`,
      customer:
        resolveCarriedCustomer(apKey, 'ut', names.customerName, { flag: false }) ||
        r.jobcode1 ||
        '(blank)',
      job: names.poLabel || r.jobcode2,
      apKey,
      estimator: r.estimator,
      crewLead: resolution?.display ?? '',
      woOverride: null,
      woPct: null,
      percent: live.percent,
      livePercent: live.percent,
      percentOverride: false,
      excludedFromPayroll: false,
      notes: '',
      colorHex: colorByEstimator.get(r.estimator) ?? null,
    });
  }

  const balances = applyManualState(balancesSeeds, input.balancesState);
  const utPayroll = applyManualState(utSeeds, input.utState);

  const grandByDate: Record<string, number> = {};
  for (const iso of dateIsos) {
    const v = grandByDateMap.get(iso);
    if (v !== undefined) grandByDate[iso] = round2(v);
  }

  return {
    title: `NV PAYROLL: ${payDate ? fmtDate(payDate) : ''}`.trim(),
    dates,
    dateIsos,
    customers,
    grandByDate,
    grandTotal: round2(grandTotal),
    balances,
    utPayroll,
    unresolvedCarriedCustomers,
  };
}

function manualRowFromCarry(args: {
  section: 'balances' | 'ut';
  row: LastWeekPbjRow;
  notes: string;
  colorByEstimator: ReadonlyMap<string, string>;
  /** This week's TimeNavi link for the same AP# (§3.3). */
  live: { linked: boolean; percent: PercentValue };
  /** Customer resolved by AP key (§4.3); falls back to the carried text. */
  customer: string;
}): ManualSectionRow {
  const { section, row, notes, colorByEstimator, live, customer } = args;
  return {
    id: `${section}:${row.apKey}`,
    customer,
    job: row.job,
    apKey: row.apKey,
    estimator: row.estimator,
    crewLead: row.crewLead,
    woOverride: row.woOverride,
    woPct: row.woPct,
    // §3.3 — CURRENT week in the column; last week's number stays in Last % Paid.
    percent: live.linked ? live.percent : row.percent,
    livePercent: live.percent,
    percentOverride: false,
    excludedFromPayroll: row.excludedFromPayroll,
    notes,
    colorHex: colorByEstimator.get(row.estimator) ?? null,
  };
}

/**
 * Merge a seed with the user's patch. Keys explicitly set to `undefined` mean
 * "drop the edit" (e.g. relinking % to TimeNavi), so they never overwrite.
 */
function applyRowEdits(
  seed: ManualSectionRow,
  patch: Partial<ManualSectionRow> | undefined,
): ManualSectionRow {
  const out: ManualSectionRow = { ...seed };
  if (!patch) return out;
  for (const key of Object.keys(patch) as (keyof ManualSectionRow)[]) {
    const value = patch[key];
    if (value === undefined) continue;
    Object.assign(out, { [key]: value });
  }
  // §3.1 — a typed % on a carried row wins over the live TimeNavi value.
  if (patch.percent !== undefined) out.percentOverride = true;
  return out;
}

/** Merge seeds with user adds/deletes/edits (rows persist keyed by id). */
export function applyManualState(seeds: ManualSectionRow[], state: ManualSectionState): ManualSectionRow[] {
  const deleted = new Set(state.deleted);
  const out: ManualSectionRow[] = [];
  const seen = new Set<string>();
  for (const seed of seeds) {
    if (deleted.has(seed.id)) continue;
    seen.add(seed.id);
    out.push(applyRowEdits(seed, state.edits[seed.id]));
  }
  for (const added of state.added) {
    if (deleted.has(added.id) || seen.has(added.id)) continue;
    out.push(applyRowEdits(added, state.edits[added.id]));
  }
  return out;
}

/** Weekday short name for the §4.1 day-name row (computed, never typed). */
export function weekdayShort(d: Date): string {
  return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][d.getDay()];
}

/** "M/D" short date used in headers. */
export function fmtMonthDay(d: Date): string {
  return `${d.getMonth() + 1}/${d.getDate()}`;
}
