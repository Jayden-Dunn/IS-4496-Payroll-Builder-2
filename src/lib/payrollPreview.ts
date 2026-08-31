/**
 * Round 3 §5 / Round 4 §5 — "Payroll Preview": one block per person (hourly
 * painters + crew-lead profit-share blocks in a single A→Z sequence, then
 * office/admin/estimator/manager sections). Every derived number is an
 * editable default; edits recompute live.
 */
import { extractApKey } from './apkey';
import { fmtDate, round2 } from './format';
import {
  carryDisplayMap,
  jobDisplayPair,
  salesDisplayMap,
  timesheetDisplayMap,
} from './jobNames';
import {
  applyDavisBaconLeadFloor,
  computeTotalPay,
  computeTotalProfitShare,
  directoryFromTimesheet,
  extractLeadParenthetical,
  extractWoOverride,
  extractWoPct,
  isPrevailingWageJob,
  jobDisplayName,
  resolveCrewLead,
  resolveLineRate,
  splitDailyOvertime,
  splitOvertime,
  stripSingleLetterTokens,
  titleCaseLeadName,
  tokensSubset,
  wageRateFor,
  type DailyOtSegment,
  type LeadResolverContext,
  type LineRate,
  type OtJobHours,
} from './payroll';
import { classifyPayItem, payItemTag, type PayRateKind } from './payTypes';
import { isHourlyJob, fmtMonthDay } from './pivotByJob';
import type {
  AppSettings,
  DailyOtNote,
  DirectoryEmployee,
  EnrichedTimesheetRow,
  EntityCode,
  HourlyBlock,
  HourlyLine,
  InactiveEmployeeBlock,
  ManualSectionRow,
  NewLeadDefault,
  PainterSectionEntry,
  PayrollEditsMap,
  PayrollExtraLine,
  PayrollPreviewModel,
  PercentValue,
  PivotByJobModel,
  PlAccountMapping,
  PlParseResult,
  PsBlock,
  PsJobColumn,
  QbHoursMap,
  ReportRow,
  SalaryBlock,
  SalesRow,
} from './types';
import type { LastWeekPbjRow } from './parsers';

export interface PayrollPreviewInput {
  pbj: PivotByJobModel;
  enriched: EnrichedTimesheetRow[];
  reportRows: ReportRow[];
  /** Slot B — the NV P&L. */
  pl: PlParseResult | null;
  /** Round 8 §7.2 — Slot C, the optional UT P&L. */
  plUt?: PlParseResult | null;
  /**
   * Round 8 §7.3 — which entity's model to build. `'NV'` (the default) keeps
   * the Round 7 tab exactly as it was, minus UT-timesheet people.
   */
  entity?: EntityCode;
  directory: DirectoryEmployee[] | null; // Slot 9 / saved (null → timesheet fallback)
  settings: AppSettings;
  payDate: Date | null;
  edits: PayrollEditsMap;
  qbHours: QbHoursMap;
  knownEstimators: ReadonlySet<string>;
  leadCtx: LeadResolverContext;
  /** Round 5 §4 — naming chain inputs. */
  nvSales?: readonly SalesRow[] | null;
  utSales?: readonly SalesRow[] | null;
  lastWeekPbj?: LastWeekPbjRow[] | null;
}

const byText = (a: string, b: string): number => a.localeCompare(b, 'en', { sensitivity: 'base' });

/** Display name of a painter-sequence entry, whatever kind it is. */
function sequenceName(e: PainterSectionEntry): string {
  return e.kind === 'hourly' ? e.block.employeeName : e.kind === 'ps' ? e.block.displayName : e.block.name;
}

/** First name for alphabetical block ordering. */
function firstNameOf(name: string): string {
  return name.trim().split(/\s+/)[0] ?? '';
}

// ---------------------------------------------------------------------------
// Round 8 §8.3 / §8.4 / §8.5 — P&L metrics, derived from the Settings mapping
// ---------------------------------------------------------------------------

/**
 * Per-job P&L numbers as the payroll chain consumes them. Differs from
 * `PlJobMetrics` in three Round 8 ways: `materials` excludes sundries (§8.3),
 * `sundries` seeds Deductions/Sundries instead, and `travel` is carried (§8.5).
 */
export interface PlDerivedMetrics {
  income: number;
  /** §8.3 — sundries NO LONGER included, so ½ Materials and % Materials shrink. */
  materials: number;
  /** §8.4 — includes NV `Subcontractor Labor` / UT `Subcontractors`. */
  laborWagesSubs: number;
  sundries: number;
  travel: number;
  mileage: number;
}

const ENTITIES: readonly EntityCode[] = ['NV', 'UT'];

/**
 * Round 8 §8.4 — read the account→metric mapping out of Settings rather than
 * the hardcoded label lists in `parsePl`, so `Subcontractor Labor` (NV) and
 * `Subcontractors` (UT) land in `laborWagesSubs` only and never in materials.
 *
 * Pass `entity` to read one entity's labels — that is what keeps UT's
 * `Subcontractors` / `Sundries` / absent mileage row distinct from NV's
 * `Subcontractor Labor` / `Sundries -` / `Mileage Expense`. With `entity`
 * omitted the two label sets are unioned, which is how Round 7 callers that
 * predate the second upload slot still behave.
 *
 * When the raw account tree is unavailable this falls back to `pl.byJob` and
 * simply un-folds sundries back out of materials.
 */
export function plMetricsByJob(
  pl: PlParseResult | null,
  mapping: Record<EntityCode, PlAccountMapping>,
  entity?: EntityCode,
): Map<string, PlDerivedMetrics> {
  const out = new Map<string, PlDerivedMetrics>();
  if (!pl) return out;

  if (pl.columnHeaders.length === 0 || pl.accountRows.length === 0) {
    for (const [ap, m] of pl.byJob) {
      out.set(ap, {
        income: m.income,
        materials: m.materials - m.sundries,
        laborWagesSubs: m.laborWagesSubs,
        sundries: m.sundries,
        travel: m.travel,
        mileage: m.mileage,
      });
    }
    return out;
  }

  const scope = entity ? [entity] : ENTITIES;
  const labelsFor = (pick: (m: PlAccountMapping) => string[]): string[] => {
    const set = new Set<string>();
    for (const e of scope) {
      for (const label of pick(mapping[e]) ?? []) {
        const t = label.trim().toLowerCase();
        if (t) set.add(t);
      }
    }
    return [...set];
  };
  const laborLabels = labelsFor((m) => m.laborWagesSubs);
  const laborLabelSet = new Set(laborLabels);
  const metricLabels = {
    income: labelsFor((m) => m.income),
    // Round 10 — subcontractor labor is a LABOR line, never a materials line.
    materials: labelsFor((m) => m.materials).filter((l) => !laborLabelSet.has(l)),
    laborWagesSubs: laborLabels,
    sundries: labelsFor((m) => m.sundries),
    travel: labelsFor((m) => m.travel),
    mileage: labelsFor((m) => m.mileage),
  };

  // First occurrence per label, matching how `parsePl` reads the tree.
  const rowByLabel = new Map<string, (number | null)[]>();
  for (const row of pl.accountRows) {
    const label = row.label.trim().toLowerCase();
    if (label && !rowByLabel.has(label)) rowByLabel.set(label, row.values);
  }
  const sumAt = (labels: string[], col: number): number => {
    let acc = 0;
    for (const label of labels) {
      const v = rowByLabel.get(label)?.[col];
      if (typeof v === 'number') acc += v;
    }
    return acc;
  };

  for (let col = 1; col < pl.columnHeaders.length; col++) {
    const header = pl.columnHeaders[col];
    if (!header || /^total\b/i.test(header)) continue;
    const apKey = extractApKey(header);
    if (!apKey) continue;
    const prev = out.get(apKey);
    const next: PlDerivedMetrics = {
      income: sumAt(metricLabels.income, col),
      materials: sumAt(metricLabels.materials, col),
      laborWagesSubs: sumAt(metricLabels.laborWagesSubs, col),
      sundries: sumAt(metricLabels.sundries, col),
      travel: sumAt(metricLabels.travel, col),
      mileage: sumAt(metricLabels.mileage, col),
    };
    // Duplicate AP# columns sum, same as `parsePl`.
    out.set(
      apKey,
      prev
        ? {
            income: prev.income + next.income,
            materials: prev.materials + next.materials,
            laborWagesSubs: prev.laborWagesSubs + next.laborWagesSubs,
            sundries: prev.sundries + next.sundries,
            travel: prev.travel + next.travel,
            mileage: prev.mileage + next.mileage,
          }
        : next,
    );
  }
  return out;
}

/** Split a resolved crew-lead display ("A" / "A/B") into matching keys. */
function leadNamesFromDisplay(display: string): string[] {
  return display
    .split('/')
    .map((s) => s.replace(/\s+/g, ' ').trim().toUpperCase())
    .filter((s) => s !== '');
}

export function buildPayrollPreview(input: PayrollPreviewInput): PayrollPreviewModel {
  const { pbj, enriched, reportRows, settings, payDate, edits, qbHours } = input;
  // Round 8 §7.2/§7.3 — one builder, two tabs. The entity decides which
  // timesheet rows count, which jobs the bands carry, and which P&L the job
  // lookups resolve against.
  const entity: EntityCode = input.entity ?? 'NV';
  const isUt = entity === 'UT';
  const pl = isUt ? (input.plUt ?? null) : input.pl;
  const entityRows = enriched.filter((r) => r.sourceEntity === entity);
  const directory = input.directory ?? directoryFromTimesheet(enriched);
  const colorByEstimator = new Map(settings.legend.map((l) => [l.estimator, l.color]));

  /**
   * Round 8 §7.3 — the UT tab's job list. UT timesheet hours never enter the
   * NV pivot tree (Round 6 §2.1), so a UT job reaches payroll as a UT PAYROLL
   * section row. The §4.1 gray flag still means "keep this out of payroll".
   */
  const utJobs = isUt
    ? pbj.utPayroll.filter((r) => !r.excludedFromPayroll && r.apKey !== 'OFFICE')
    : [];
  /**
   * Round 8 §6 — non-gray NV PAYROLL (BALANCES) rows join this tab as their
   * own job columns, even with zero timesheet hours. Gray-excluded rows stay
   * pivot-only. The UT tab reads `utPayroll` instead.
   */
  const nvBalanceJobs = !isUt
    ? pbj.balances.filter((r) => !r.excludedFromPayroll && r.apKey !== 'OFFICE')
    : [];
  const utJobByAp = new Map<string, ManualSectionRow>();
  for (const r of utJobs) {
    if (r.apKey && !utJobByAp.has(r.apKey)) utJobByAp.set(r.apKey, r);
  }
  const nvJobByAp = new Map<string, ManualSectionRow>();
  for (const r of nvBalanceJobs) {
    if (r.apKey && !nvJobByAp.has(r.apKey)) nvJobByAp.set(r.apKey, r);
  }
  const pbjJobByAp = new Map<string, (typeof pbj.customers)[number]['jobs'][number]>();
  for (const c of pbj.customers) {
    for (const j of c.jobs) {
      if (j.analysis.apKey && !pbjJobByAp.has(j.analysis.apKey)) pbjJobByAp.set(j.analysis.apKey, j);
    }
  }
  const carriedIsHourly = (row: ManualSectionRow): boolean =>
    isHourlyJob(row.percent, row.crewLead, `${row.customer} ${row.job}`);

  const nameSources = {
    timesheetByAp: timesheetDisplayMap(enriched),
    carryByAp: carryDisplayMap(input.lastWeekPbj),
    salesByAp: salesDisplayMap(input.nvSales, input.utSales),
    // Fall back to the NV tree so a UT-only week still resolves shared names.
    pl: pl ?? input.pl,
  };

  // ----- week range --------------------------------------------------------
  let weekStart: Date | null = null;
  let weekEnd: Date | null = null;
  for (const r of enriched) {
    if (!r.localDate) continue;
    if (!weekStart || r.localDate.getTime() < weekStart.getTime()) weekStart = r.localDate;
    if (!weekEnd || r.localDate.getTime() > weekEnd.getTime()) weekEnd = r.localDate;
  }
  const weekRange = weekStart && weekEnd ? `${fmtMonthDay(weekStart)} - ${fmtMonthDay(weekEnd)}` : '';
  const range = weekRange ? `(${weekRange})` : '';
  const title = `${isUt ? 'UT ' : ''}PAYROLL BREAKDOWN ${payDate ? fmtDate(payDate) : ''} ${range}`
    .replace(/\s+/g, ' ')
    .trim();

  // ----- employee snapshot --------------------------------------------------
  interface Emp {
    username: string;
    fname: string;
    lname: string;
    group: string;
    fullName: string; // directory name when matched (raw, for matching)
    displayName: string; // §3.2-2 single-letter tokens stripped
    inactive: boolean;
  }
  /**
   * Round 10 — anyone with hours this week is active, full stop. A painter who
   * was on the inactive roster and shows up on the timesheet has come back, so
   * no "inactive (*)" marker is drawn and the roster flag is flipped (the App
   * layer persists `activatedPainters`).
   */
  const usernamesWithHours = new Set<string>();
  for (const r of entityRows) {
    if (r.hours > 0) usernamesWithHours.add(r.username || `${r.fname} ${r.lname}`);
  }
  const activatedPainters = new Set<string>();

  const emps = new Map<string, Emp>();
  for (const r of entityRows) {
    const key = r.username || `${r.fname} ${r.lname}`;
    if (emps.has(key)) continue;
    const tsName = `${r.fname} ${r.lname}`.replace(/\s+/g, ' ').trim();
    const dirMatch = directory.find((d) => tokensSubset(tsName, d.fullName) || tokensSubset(d.fullName, tsName));
    const fullName = dirMatch?.fullName ?? tsName;
    const rosterInactive = settings.painterRoster.some(
      (p) => !p.active && (tokensSubset(p.name, fullName) || tokensSubset(fullName, p.name)),
    );
    emps.set(key, {
      username: key,
      fname: r.fname,
      lname: r.lname,
      group: r.group,
      fullName,
      displayName: stripSingleLetterTokens(fullName),
      inactive:
        !usernamesWithHours.has(key) && ((dirMatch ? !dirMatch.active : false) || rosterInactive),
    });
    if (usernamesWithHours.has(key) && rosterInactive) {
      const rosterRow = settings.painterRoster.find(
        (pr) => !pr.active && (tokensSubset(pr.name, fullName) || tokensSubset(fullName, pr.name)),
      );
      if (rosterRow) activatedPainters.add(rosterRow.name);
    }
  }

  // ----- per (employee, job) QB-adjusted hours ------------------------------
  /**
   * Round 9 §16/§18 — one bucket of an employee's hours on a job that share a
   * timesheet pay item. QuickBooks Time already splits overtime onto its own
   * rows ("Overtime Pay", "Crew Lead - Hourly Wage Overtime", "Daily
   * Overtime"), so a segment IS the payroll line: hours, the rate slot the pay
   * item selects, and whether those hours are overtime.
   */
  interface EmpJobSeg {
    /** Line-id suffix; '' for plain straight-time Regular Pay. */
    key: string;
    payItem: string;
    rateKind: PayRateKind;
    overtime: boolean;
    otReason: 'weekly' | 'daily' | null;
    hours: number;
    firstIndex: number;
  }

  interface EmpJob {
    jobId: string;
    apKey: string | null;
    customer: string;
    jobcode2: string;
    hours: number;
    hourly: boolean;
    colorHex: string | null;
    firstIndex: number;
    lastDate: Date | null;
    lastRowIndex: number;
    office: boolean;
    /** Text used for prevailing-wage pattern matching (§5.8). */
    jobText: string;
    /** Pivot §7.1 — user forced non-DB on a pattern-matched job. */
    forceNonDb: boolean;
    /** Round 8 §9.1 — crew leads on this job, for the conditional rate rules. */
    leadNames: string[];
    /** Pivot by Job Notes for this AP#. */
    notes: string;
    /** Round 9 §16 — per-pay-item hour buckets, in first-encounter order. */
    segments: EmpJobSeg[];
  }
  /** Round 8 §9.2 — one (job, day) bucket of an employee's hours. */
  interface EmpDayRow {
    jobId: string;
    date: Date | null;
    hours: number;
    rowIndex: number;
    /** NV timesheet upload — UT rows never get daily OT. */
    nv: boolean;
  }
  const empJobs = new Map<string, EmpJob[]>(); // username → jobs in first-encounter order
  const empDayRows = new Map<string, EmpDayRow[]>(); // username → (job, day) buckets
  const jobRowById = new Map(pbj.customers.flatMap((c) => c.jobs.map((j) => [j.id, j] as const)));
  const eventByApKey = new Map<string, ReportRow>();
  for (const r of reportRows) {
    if (r.included && r.apKey && !eventByApKey.has(r.apKey)) eventByApKey.set(r.apKey, r);
  }

  /** Round 9 §16 — line-id suffix + label tag for a classified pay item. */
  const segKeyFor = (cls: ReturnType<typeof classifyPayItem>): string => {
    if (cls.kind === 'overtime') {
      if (cls.reason === 'daily') return ':DOT';
      return cls.rate === 'crewLead' ? ':CLOT' : ':OT';
    }
    switch (cls.rate) {
      case 'crewLead':
        return ':CL';
      case 'commercialHourly':
        return ':CH';
      case 'wages2':
        return ':W2';
      case 'salary':
        return ':SAL';
      default:
        return '';
    }
  };

  /**
   * Round 9 §13 — who clocked CREW-LEAD hours on which job. When two leads
   * work the same job, the one on "Crew Lead - Hourly Wage" is the lead for
   * that job; the other is paid their own regular rate.
   */
  const crewLeadUsersByJob = new Map<string, Set<string>>();

  entityRows.forEach((r, idx) => {
    if (r.hours === 0) return;
    const username = r.username || `${r.fname} ${r.lname}`;
    const jobId = `${r.jobcode1 || '(blank)'}\u0000${r.jobcode2}`;
    let list = empJobs.get(username);
    if (!list) {
      list = [];
      empJobs.set(username, list);
    }
    let j = list.find((x) => x.jobId === jobId);
    if (!j) {
      const pbjJob = jobRowById.get(jobId) ?? (r.apKey ? pbjJobByAp.get(r.apKey) : undefined);
      // UT / carried NV jobs have no pivot-tree row — analysis lives on the section row.
      const utRow = r.apKey ? utJobByAp.get(r.apKey) : undefined;
      const nvRow = r.apKey ? nvJobByAp.get(r.apKey) : undefined;
      const eventTitle = r.apKey ? (eventByApKey.get(r.apKey)?.event ?? '') : '';
      const res = pbjJob?.analysis.crewLeadResolution;
      const hourly =
        pbjJob?.analysis.hourly ??
        (utRow ? carriedIsHourly(utRow) : undefined) ??
        (nvRow ? carriedIsHourly(nvRow) : undefined) ??
        false;
      j = {
        jobId,
        apKey: r.apKey,
        customer: r.jobcode1,
        jobcode2: r.jobcode2,
        hours: 0,
        hourly,
        colorHex:
          r.apKey === 'OFFICE'
            ? '4F81BD'
            : (pbjJob?.colorHex ?? utRow?.colorHex ?? nvRow?.colorHex ?? colorByEstimator.get(r.estimator) ?? null),
        firstIndex: idx,
        lastDate: r.localDate,
        lastRowIndex: idx,
        office: r.apKey === 'OFFICE',
        jobText: `${r.jobcode1} ${r.jobcode2} ${eventTitle}`,
        forceNonDb: pbjJob?.analysis.dbOverride === true,
        leadNames:
          res && res.names.length > 0
            ? res.names
            : pbjJob?.analysis.crewLead
              ? [pbjJob.analysis.crewLead]
              : utRow
                ? leadNamesFromDisplay(utRow.crewLead)
                : nvRow
                  ? leadNamesFromDisplay(nvRow.crewLead)
                  : [],
        notes: pbjJob?.analysis.notes || utRow?.notes || nvRow?.notes || '',
        segments: [],
      };
      list.push(j);
    }
    j.hours += r.hours;

    // Round 9 §16 — bucket by pay item so the timesheet's own overtime rows
    // stay separate lines at their own rate.
    const cls = classifyPayItem(r.payrollItem);
    const segKey = segKeyFor(cls);
    let seg = j.segments.find((x) => x.key === segKey);
    if (!seg) {
      seg = {
        key: segKey,
        payItem: r.payrollItem.trim(),
        rateKind: cls.rate,
        overtime: cls.kind === 'overtime',
        otReason: cls.kind === 'overtime' ? cls.reason : null,
        hours: 0,
        firstIndex: idx,
      };
      j.segments.push(seg);
    }
    seg.hours += r.hours;
    if (cls.rate === 'crewLead') {
      let set = crewLeadUsersByJob.get(jobId);
      if (!set) {
        set = new Set<string>();
        crewLeadUsersByJob.set(jobId, set);
      }
      set.add(username);
    }

    const t = r.localDate ? r.localDate.getTime() : -Infinity;
    const bt = j.lastDate ? j.lastDate.getTime() : -Infinity;
    if (t > bt || (t === bt && idx > j.lastRowIndex)) {
      j.lastDate = r.localDate;
      j.lastRowIndex = idx;
    }

    // Round 8 §9.2 — same hours bucketed by day, so daily OT can be measured.
    let dayList = empDayRows.get(username);
    if (!dayList) {
      dayList = [];
      empDayRows.set(username, dayList);
    }
    const dayIso = r.localDate ? r.localDate.toDateString() : '';
    let dr = dayList.find((x) => x.jobId === jobId && (x.date ? x.date.toDateString() : '') === dayIso);
    if (!dr) {
      dr = { jobId, date: r.localDate, hours: 0, rowIndex: idx, nv: r.sourceEntity !== 'UT' };
      dayList.push(dr);
    }
    dr.hours += r.hours;
    dr.rowIndex = Math.max(dr.rowIndex, idx);
    if (r.sourceEntity === 'UT') dr.nv = false;
  });

  // Apply QB Hrs overrides (§3.3): keyed `${jobRowId}\u0000${username}`.
  // The day buckets and pay-item segments are scaled with them, so both daily
  // OT and the timesheet's own overtime lines reflect the same hours.
  for (const [username, jobs] of empJobs) {
    for (const j of jobs) {
      const o = qbHours[`${j.jobId}\u0000${username}`];
      if (o === undefined) continue;
      const dayRows = (empDayRows.get(username) ?? []).filter((d) => d.jobId === j.jobId);
      const raw = dayRows.reduce((a, d) => a + d.hours, 0);
      if (raw > 0) {
        for (const d of dayRows) d.hours = round2((d.hours / raw) * o);
      } else if (dayRows.length > 0) {
        dayRows[dayRows.length - 1].hours = o;
      }
      const segRaw = j.segments.reduce((a, seg) => a + seg.hours, 0);
      if (segRaw > 0) {
        for (const seg of j.segments) seg.hours = round2((seg.hours / segRaw) * o);
      } else if (j.segments.length > 0) {
        j.segments[j.segments.length - 1].hours = o;
      }
      j.hours = o;
    }
  }

  // ----- crew leads this week ----------------------------------------------
  // Names resolved on piece-rate NV pivot jobs + sub-only calendar jobs.
  const leadNames = new Map<string, { sub: boolean }>(); // UPPER full name
  const isSubName = (name: string): boolean =>
    settings.subcontractors.some((s) => tokensSubset(s, name) || tokensSubset(name, s));
  // The pivot tree is NV-only by construction, so the UT tab starts empty here
  // and fills its bands from `utJobs` instead.
  const pieceRateJobs = isUt
    ? []
    : pbj.customers.flatMap((c) =>
        c.jobs
          .filter((j) => !j.analysis.hourly && j.analysis.apKey !== 'OFFICE')
          .map((j) => ({ job: j, customer: c.customer })),
      );
  for (const row of [...utJobs, ...nvBalanceJobs]) {
    if (carriedIsHourly(row)) continue;
    for (const n of leadNamesFromDisplay(row.crewLead)) {
      if (!leadNames.has(n)) leadNames.set(n, { sub: isSubName(n) });
    }
  }
  for (const { job } of pieceRateJobs) {
    const res = job.analysis.crewLeadResolution;
    const names = res && res.names.length > 0 ? res.names : job.analysis.crewLead ? [job.analysis.crewLead] : [];
    for (const n of names) {
      const up = n.toUpperCase();
      if (!leadNames.has(up)) leadNames.set(up, { sub: res?.sub ?? false });
    }
  }
  // Sub/no-hours calendar jobs (e.g. Pepe's Painting): included NV piece-rate
  // events with no timesheet hours whose lead resolves.
  const timesheetApKeys = new Set<string>();
  for (const c of pbj.customers) for (const j of c.jobs) if (j.analysis.apKey) timesheetApKeys.add(j.analysis.apKey);
  interface CalJob {
    row: ReportRow;
    lead: string[];
    sub: boolean;
    split: number;
    woOverride: number | null;
    woPct: number | null;
  }
  const calOnlyJobs: CalJob[] = [];
  // NV only — Utah/Reno calendar events already seeded the UT PAYROLL section.
  for (const r of isUt ? [] : reportRows) {
    if (!r.included || !r.apKey || timesheetApKeys.has(r.apKey)) continue;
    if (/^utah\b|^reno\b/i.test(r.calendar.trim())) continue; // UT pays via Pivot by Job
    const res = resolveCrewLead(extractLeadParenthetical(r.event, input.knownEstimators), input.leadCtx);
    if (res.names.length === 0) continue;
    if (isHourlyJob(r.percent, res.display, r.event)) continue;
    const woSrc = extractWoOverride(r.description);
    const pctSrc = extractWoPct(r.description);
    void woSrc;
    void pctSrc;
    calOnlyJobs.push({
      row: r,
      lead: res.names,
      sub: res.sub,
      split: res.split,
      woOverride: null,
      woPct: null,
    });
    for (const n of res.names) {
      const up = n.toUpperCase();
      if (!leadNames.has(up)) leadNames.set(up, { sub: res.sub });
    }
  }

  // Round 7 §20.1 — timesheet crew-group owners are leads even when no
  // piece-rate parenthetical named them this week (e.g. "Antonio Zuniga's Crew").
  const crewOwnerOf = (group: string): string | null => {
    const m = /^(.*?)['\u2019]s\s+crew\s*$/i.exec(group.trim());
    return m ? m[1].trim() : null;
  };
  const addLeadName = (name: string, sub = false): void => {
    const up = name.replace(/\s+/g, ' ').trim().toUpperCase();
    if (!up) return;
    if (!leadNames.has(up)) leadNames.set(up, { sub });
  };
  const nameMatches = (a: string, b: string): boolean =>
    tokensSubset(a, b) ||
    tokensSubset(b, a) ||
    a.replace(/\s+/g, ' ').trim().toUpperCase() === b.replace(/\s+/g, ' ').trim().toUpperCase();
  const seenCrewOwners = new Set<string>();
  for (const emp of emps.values()) {
    const owner = crewOwnerOf(emp.group);
    if (!owner) continue;
    const ownerKey = owner.toUpperCase();
    if (seenCrewOwners.has(ownerKey)) continue;
    seenCrewOwners.add(ownerKey);
    const match = [...emps.values()].find((e) => nameMatches(e.fullName, owner));
    /**
     * Round 9 §18 — owning a crew label is not by itself proof of crew-lead
     * work. Samuel Brewster's whole week is clocked as "Commercial Hourly", so
     * he is an hourly painter this week, not a profit-share lead.
     *
     * The Round 7 §20.1 promotion still stands by default; it is only skipped
     * when the timesheet positively says otherwise — the person clocked a
     * PERSONAL pay type (Commercial Hourly / Wages 2) and never a crew-lead
     * one, and nothing else names them as a lead. A blank or plain
     * "Regular Pay" week leaves the promotion exactly as it was.
     */
    const ownerUser = match?.username ?? null;
    const ownerJobs = ownerUser ? (empJobs.get(ownerUser) ?? []) : [];
    const clockedCrewLead = ownerJobs.some((j) => j.segments.some((sg) => sg.rateKind === 'crewLead'));
    const clockedOwnPayType = ownerJobs.some((j) =>
      j.segments.some((sg) => sg.rateKind === 'commercialHourly' || sg.rateKind === 'wages2'),
    );
    if (clockedOwnPayType && !clockedCrewLead) {
      const ownerName = match?.fullName ?? owner;
      const namedOnPivot = pieceRateJobs.some((pj) => {
        const res = pj.job.analysis.crewLeadResolution;
        const names =
          res && res.names.length > 0
            ? res.names
            : pj.job.analysis.crewLead
              ? [pj.job.analysis.crewLead]
              : [];
        return names.some((n) => nameMatches(n, ownerName));
      });
      const inProfitShareMap = settings.profitShare.some((pentry) => nameMatches(pentry.lead, ownerName));
      if (!namedOnPivot && !inProfitShareMap) continue;
    }
    addLeadName(match?.fullName ?? owner);
  }
  // Optional boost: Settings profitShare leads with hours this week.
  for (const p of settings.profitShare) {
    const lead = p.lead.trim();
    if (!lead) continue;
    const match = [...emps.values()].find((e) => nameMatches(e.fullName, lead));
    if (!match) continue;
    const hrs = (empJobs.get(match.username) ?? []).reduce((a, j) => a + j.hours, 0);
    if (hrs > 0) addLeadName(match.fullName);
  }

  /**
   * Round 9 — one band per human. "Rene Salazar" (crew label / Settings) and
   * "Rene Salazar M" (timesheet + directory) are the same person, and each
   * used to raise its own band that paid the same jobs twice. Names that
   * token-match are collapsed onto the most specific spelling — the one that
   * matches a timesheet employee when there is one.
   */
  {
    const entries = [...leadNames.entries()];
    const canonical: { key: string; meta: { sub: boolean } }[] = [];
    const matchesEmp = (n: string): boolean =>
      [...emps.values()].some((e) => nameMatches(e.fullName, n));
    for (const [name, meta] of entries) {
      const at = canonical.findIndex(
        (c) => tokensSubset(c.key, name) || tokensSubset(name, c.key),
      );
      if (at < 0) {
        canonical.push({ key: name, meta });
        continue;
      }
      const kept = canonical[at].key;
      // Prefer the spelling that matches a real timesheet employee; when both
      // (or neither) do, prefer the longer, more specific one.
      const preferNew =
        (matchesEmp(name) && !matchesEmp(kept)) ||
        (matchesEmp(name) === matchesEmp(kept) && name.length > kept.length);
      canonical[at] = {
        key: preferNew ? name : kept,
        meta: { sub: canonical[at].meta.sub || meta.sub },
      };
    }
    if (canonical.length !== entries.length) {
      leadNames.clear();
      for (const c of canonical) leadNames.set(c.key, c.meta);
    }
  }

  /** Is this timesheet employee one of the resolved crew leads? */
  const isLead = (fullName: string): boolean => {
    for (const lead of leadNames.keys()) {
      if (tokensSubset(fullName, lead) || tokensSubset(lead, fullName)) return true;
    }
    return false;
  };

  /** §5.9 — salaried-roster people never get an hourly block. */
  const isSalaried = (fullName: string): boolean =>
    settings.salariedRoster.some(
      (s) => tokensSubset(s.name, fullName) || tokensSubset(fullName, s.name),
    );

  // ----- Round 9 §19/§23/§24/§29 — person-scoped extra rows ------------------
  /** True when the pay date is the FIRST pay date of its month (1st–7th). */
  const isFirstPayDateOfMonth = payDate !== null && payDate.getDate() <= 7;

  /** The extra rows configured for one person, resolved for this pay date. */
  const extrasFor = (fullName: string, blockKey: string): PayrollExtraLine[] => {
    const out: PayrollExtraLine[] = [];
    for (const item of settings.payrollLineItems ?? []) {
      if (!(tokensSubset(item.person, fullName) || tokensSubset(fullName, item.person))) continue;
      if (item.frequency === 'firstPayDateOfMonth' && !isFirstPayDateOfMonth) continue;
      const key = `xl:${blockKey}:${item.id}`;
      const e = edits[key] ?? {};
      out.push({
        key,
        label: item.label,
        amount: e.amount !== undefined ? e.amount : item.amount,
        kind: item.kind,
      });
    }
    return out;
  };

  /** Reimbursements − deductions across a block's extra rows. */
  const extrasNet = (extras: readonly PayrollExtraLine[]): number =>
    extras.reduce((acc, x) => {
      if (x.amount === null) return acc;
      if (x.kind === 'reimbursement') return acc + x.amount;
      if (x.kind === 'deduction') return acc - x.amount;
      return acc;
    }, 0);

  /**
   * Round 9 §1 — labor dollars per AP#, computed from the TIMESHEET: every
   * painter's hours on the job at their resolved line rate (overtime lines at
   * 1.5×). This is what the crew lead's "- Labor + Wages" row is charged, and
   * it comes from the timesheets tab of this side's input, never from the P&L.
   */
  const timesheetLaborByAp = new Map<string, number>();
  {
    const empByUsername = new Map([...emps.values()].map((e) => [e.username, e] as const));
    for (const [username, jobs] of empJobs) {
      const emp = empByUsername.get(username);
      if (!emp) continue;
      const empIsLead = isLead(emp.fullName);
      for (const j of jobs) {
        if (!j.apKey || j.apKey === 'OFFICE') continue;
        const leadUsers = crewLeadUsersByJob.get(j.jobId);
        const someoneElseLeads =
          leadUsers !== undefined && leadUsers.size > 0 && !leadUsers.has(username);
        const segs = j.segments.length > 0
          ? j.segments
          : [{ key: '', payItem: '', rateKind: 'regular' as PayRateKind, overtime: false, otReason: null, hours: j.hours, firstIndex: j.firstIndex }];
        let acc = 0;
        for (const sg of segs) {
          const info = resolveLineRate(emp.fullName, j.jobText, j.hourly, {
            ...settings,
            apKey: j.apKey,
            jobRateRules: settings.jobRateRules,
            forceNonDb: j.forceNonDb,
            isCrewLead: sg.rateKind === 'crewLead' ? true : someoneElseLeads ? false : empIsLead,
            jobLeadNames: j.leadNames,
            payRateKind: sg.rateKind,
          });
          acc += sg.hours * (sg.overtime ? info.rate * 1.5 : info.rate);
        }
        timesheetLaborByAp.set(j.apKey, round2((timesheetLaborByAp.get(j.apKey) ?? 0) + acc));
      }
    }
  }

  // ----- hourly block construction ------------------------------------------
  const dailyOtNotes: DailyOtNote[] = [];

  /** Round 9 §27 — people the NV daily-OT rule never applies to. */
  const dailyOtExempt = (fullName: string): boolean =>
    (settings.dailyOtExemptNames ?? []).some(
      (n) => tokensSubset(n, fullName) || tokensSubset(fullName, n),
    );

  const buildHourlyBlock = (emp: Emp, officeStyle: boolean, isLead = false): HourlyBlock => {
    const jobs = empJobs.get(emp.username) ?? [];

    /**
     * Round 9 §13/§18 — the rate is resolved PER SEGMENT, because the
     * timesheet pay item decides which rate slot the hours are paid from
     * (Regular Pay / Crew Lead - Hourly Wage / Commercial Hourly / Wages 2).
     * `isCrewLead` still applies when the pay item itself is not crew-lead
     * but this IS the lead's own block and no one else on the job clocked
     * crew-lead hours.
     */
    const rateFor = (j: EmpJob, seg: EmpJobSeg | null): LineRate => {
      const payRateKind: PayRateKind = seg?.rateKind ?? 'regular';
      // §13 — when someone else on this job clocked "Crew Lead - Hourly Wage"
      // and this person did not, this person is NOT the lead on that job.
      const leadUsers = crewLeadUsersByJob.get(j.jobId);
      const someoneElseLeads =
        leadUsers !== undefined && leadUsers.size > 0 && !leadUsers.has(emp.username);
      const isCrewLeadHere =
        payRateKind === 'crewLead' ? true : someoneElseLeads ? false : isLead;
      return resolveLineRate(emp.fullName, j.jobText, j.hourly, {
        ...settings,
        apKey: j.apKey,
        jobRateRules: settings.jobRateRules,
        forceNonDb: j.forceNonDb,
        isCrewLead: isCrewLeadHere,
        jobLeadNames: j.leadNames,
        payRateKind,
      });
    };

    /** Straight-time rate on a job — what daily-OT eligibility is measured on. */
    const straightRate = (jobId: string): number => {
      const j = jobs.find((x) => x.jobId === jobId);
      if (!j) return settings.defaultWageRate;
      const seg = j.segments.find((sg) => !sg.overtime) ?? null;
      const e = edits[`hl:${emp.username}\u0000${jobId}${seg?.key ?? ''}`] ?? {};
      if (e.rate !== undefined) return e.rate;
      return rateFor(j, seg).rate;
    };

    // Round 9 §16 — the timesheet's own overtime rows are authoritative. Only
    // when the week carries NONE do the computed >8/day and >40/week rules run,
    // so a QuickBooks file without OT items still splits overtime correctly.
    const hasTimesheetOt = jobs.some((j) => j.segments.some((sg) => sg.overtime));

    const daily = hasTimesheetOt
      ? { otByJob: new Map<string, number>(), segments: [] as DailyOtSegment[], totalOtHours: 0 }
      : splitDailyOvertime(
          (empDayRows.get(emp.username) ?? []).map((d) => ({
            jobId: d.jobId,
            date: d.date,
            hours: d.hours,
            rowIndex: d.rowIndex,
            eligible:
              d.nv &&
              !dailyOtExempt(emp.fullName) &&
              straightRate(d.jobId) < settings.nvDailyOtRateThreshold,
          })),
          settings.dailyOtHoursThreshold,
        );

    // §3.4 weekly overtime on the QB-adjusted hours (computed fallback only).
    const ot = hasTimesheetOt
      ? { regularByJob: new Map<string, number>(), otJobId: null as string | null, otHours: 0 }
      : splitOvertime(
          jobs.map(
            (j): OtJobHours => ({
              jobId: j.jobId,
              hours: j.hours,
              lastDate: j.lastDate,
              lastRowIndex: j.lastRowIndex,
            }),
          ),
          { dailyOtHours: daily.totalOtHours },
        );

    const lines: HourlyLine[] = [];
    const pushLine = (
      j: EmpJob,
      seg: EmpJobSeg | null,
      hours: number,
      kind: 'regular' | 'weeklyOt' | 'dailyOt',
      segment?: DailyOtSegment,
    ): void => {
      const isOt = kind !== 'regular';
      const computedSuffix =
        kind === 'weeklyOt' ? ':OT' : kind === 'dailyOt' ? `:DOT:${segment?.dateIso ?? ''}` : '';
      // Timesheet-driven lines key off the pay item; computed OT keeps the
      // Round 8 suffixes so existing manual edits still land.
      const suffix = seg ? seg.key : computedSuffix;
      const lineId = `${emp.username}\u0000${j.jobId}${suffix}`;
      const e = edits[`hl:${lineId}`] ?? {};
      const rateInfo = rateFor(j, seg);
      const baseRate = e.rate !== undefined ? e.rate : rateInfo.rate;
      const rate = isOt ? round2(baseRate * 1.5) : baseRate;
      const lineHours = e.hours !== undefined ? e.hours : round2(hours);
      // Round 8 §9.2 — the two OT reasons carry distinguishable tags.
      const dayLabel = segment?.date ? fmtMonthDay(segment.date) : '';
      const dailyTag = segment
        ? `[Daily OT — ${dayLabel}: ${round2(hours)} hrs over ${settings.dailyOtHoursThreshold} @ 1.5× (rate < $${settings.nvDailyOtRateThreshold}, NV)]`
        : seg?.overtime
          ? `[${seg.payItem} — timesheet overtime line, paid at 1.5× $${baseRate.toFixed(2)}]`
          : undefined;
      const otTag = seg
        ? payItemTag(seg.payItem)
        : kind === 'weeklyOt'
          ? ' [OT]'
          : kind === 'dailyOt'
            ? ' [Daily OT]'
            : '';
      const isOfficeLine = officeStyle || j.office;
      const notesSuffix = !isOfficeLine && j.notes.trim() ? ` — ${j.notes.trim()}` : '';
      const typeTag = isOfficeLine ? '' : j.hourly ? ' [Hourly]' : ' [PS]';
      const label = isOfficeLine
        ? `Office${otTag}`
        : `${jobDisplayName(j.customer, j.jobcode2)}${typeTag}${otTag}${notesSuffix}`;
      // §5.5 — office lines carry the week range (or "Allpro Office" inside a
      // mixed painter block) on the blue PO cell instead of an AP number.
      const poText = officeStyle ? `(${weekRange})` : j.office ? 'Allpro Office' : null;
      const isDailyOtLine = kind === 'dailyOt' || (seg?.overtime === true && seg.otReason === 'daily');
      if (isDailyOtLine && lineHours > 0) {
        dailyOtNotes.push({
          employeeName: emp.displayName,
          date: dayLabel,
          hours: lineHours,
          rate,
          jobLabel: isOfficeLine ? 'Office' : jobDisplayName(j.customer, j.jobcode2),
        });
      }
      lines.push({
        id: lineId,
        apKey: j.apKey,
        poText,
        jobLabel: label,
        hours: lineHours,
        rate,
        rateDefaulted: rateInfo.defaulted && e.rate === undefined,
        prevailing: rateInfo.prevailing,
        belowPrevailing: rateInfo.prevailing && baseRate < settings.prevailingMinRate,
        fromJobRule: rateInfo.fromJobRule && e.rate === undefined ? true : undefined,
        ruleNote:
          e.rate === undefined &&
          (rateInfo.fromJobRule || rateInfo.fromPerJobOverride || rateInfo.fromConditionalRule)
            ? rateInfo.ruleNote
            : undefined,
        fromPerJobOverride: rateInfo.fromPerJobOverride && e.rate === undefined ? true : undefined,
        fromConditionalRule: rateInfo.fromConditionalRule && e.rate === undefined ? true : undefined,
        fromCrewLeadRate: rateInfo.fromCrewLeadRate && e.rate === undefined ? true : undefined,
        ot: isOt && !isDailyOtLine,
        dailyOt: isDailyOtLine ? true : undefined,
        dailyOtDate: isDailyOtLine && dayLabel ? dayLabel : undefined,
        dailyOtNote: dailyTag,
        payItem: seg?.payItem || undefined,
        fromTimesheetOt: seg?.overtime ? true : undefined,
        hourlyJob: j.hourly,
        colorHex: j.colorHex,
        jobNotes: j.notes.trim() || undefined,
      });
    };

    /** Segments in a stable order: straight time first, overtime last. */
    const orderedSegments = (j: EmpJob): EmpJobSeg[] =>
      j.segments
        .filter((sg) => round2(sg.hours) !== 0)
        .slice()
        .sort((a, b) => Number(a.overtime) - Number(b.overtime) || a.firstIndex - b.firstIndex);

    if (officeStyle && jobs.length > 0) {
      // §5.3 — one "Office" line per pay item covering the week, so an office
      // overtime row still shows separately at its own rate.
      const rep = jobs[0];
      const merged = new Map<string, EmpJobSeg>();
      for (const j of jobs) {
        for (const sg of orderedSegments(j)) {
          const at = merged.get(sg.key);
          if (at) at.hours = round2(at.hours + sg.hours);
          else merged.set(sg.key, { ...sg });
        }
      }
      const officeSegs = [...merged.values()].sort(
        (a, b) => Number(a.overtime) - Number(b.overtime) || a.firstIndex - b.firstIndex,
      );
      if (hasTimesheetOt || officeSegs.length > 1) {
        for (const sg of officeSegs) pushLine(rep, sg, sg.hours, sg.overtime ? 'weeklyOt' : 'regular');
      } else {
        const totalHrs = jobs.reduce((a, j) => a + j.hours, 0);
        const regular = round2(Math.max(0, totalHrs - daily.totalOtHours - ot.otHours));
        pushLine(rep, officeSegs[0] ?? null, regular, 'regular');
        for (const seg of daily.segments) pushLine(rep, null, seg.otHours, 'dailyOt', seg);
        if (ot.otHours > 0) pushLine(rep, null, ot.otHours, 'weeklyOt');
      }
    } else {
      for (const j of jobs.slice().sort((a, b) => a.firstIndex - b.firstIndex)) {
        const segs = orderedSegments(j);
        if (hasTimesheetOt) {
          // Round 9 §10/§16 — straight-time and overtime hours keep their own
          // lines exactly as the timesheet split them; nothing is re-derived,
          // so a job never pays ALL of its hours at the overtime rate.
          for (const sg of segs) pushLine(j, sg, sg.hours, sg.overtime ? 'weeklyOt' : 'regular');
          continue;
        }
        // Computed fallback: carve the day's / week's overtime out of this
        // job's straight-time segments, last segment first.
        const dailyOnJob = daily.otByJob.get(j.jobId) ?? 0;
        const weeklyOnJob = ot.otJobId === j.jobId ? ot.otHours : 0;
        let toRemove = round2(dailyOnJob + weeklyOnJob);
        const straight = segs.filter((sg) => !sg.overtime);
        const regularBySeg = new Map<string, number>(straight.map((sg) => [sg.key, round2(sg.hours)]));
        for (let i = straight.length - 1; i >= 0 && toRemove > 0; i--) {
          const sg = straight[i];
          const have = regularBySeg.get(sg.key) ?? 0;
          const take = round2(Math.min(have, toRemove));
          regularBySeg.set(sg.key, round2(have - take));
          toRemove = round2(toRemove - take);
        }
        for (const sg of straight) pushLine(j, sg, regularBySeg.get(sg.key) ?? 0, 'regular');
        if (straight.length === 0 && segs.length === 0) {
          pushLine(j, null, round2(Math.max(0, j.hours - dailyOnJob - weeklyOnJob)), 'regular');
        }
        for (const seg of daily.segments.filter((sd) => sd.jobId === j.jobId)) {
          pushLine(j, null, seg.otHours, 'dailyOt', seg);
        }
        if (weeklyOnJob > 0) pushLine(j, null, weeklyOnJob, 'weeklyOt');
      }
    }

    const totalHours = round2(lines.reduce((a, l) => a + l.hours, 0));
    const totalPay = round2(lines.reduce((a, l) => a + l.hours * l.rate, 0));
    const blockEdit = edits[`hb:${emp.username}`] ?? {};
    // Round 9 §19/§23 — person-scoped extra rows (Gas Reimb, Aflac, …). A crew
    // lead's own T-SHEETS block skips them: the PS band directly above it
    // already carries that person's rows, and a deduction must only bite once.
    const extras = isLead ? [] : extrasFor(emp.fullName, `hb:${emp.username}`);
    return {
      key: emp.username,
      title: officeStyle ? emp.displayName : `${emp.displayName}${emp.group ? ` (${emp.group})` : ''}`,
      employeeName: emp.displayName,
      crewTag: officeStyle ? '' : emp.group,
      lines,
      totalHours,
      totalPay,
      taxes: blockEdit.taxes !== undefined ? blockEdit.taxes : null,
      nest: blockEdit.nest !== undefined ? blockEdit.nest : null,
      net: blockEdit.net !== undefined ? blockEdit.net : null,
      inactive: emp.inactive,
      extras,
      netPay: round2(totalPay + extrasNet(extras)),
    };
  };

  // ----- sections ------------------------------------------------------------
  const officeEmps: Emp[] = [];
  const painterEmps: Emp[] = [];
  const leadEmps: Emp[] = [];
  for (const emp of emps.values()) {
    if (isSalaried(emp.fullName)) continue; // §5.9 — roster block only
    if (/office/i.test(emp.group)) officeEmps.push(emp);
    else if (isLead(emp.fullName)) leadEmps.push(emp);
    else painterEmps.push(emp);
  }

  const painters = painterEmps
    .map((e) => buildHourlyBlock(e, false))
    .sort((a, b) => byText(firstNameOf(a.employeeName), firstNameOf(b.employeeName)));
  const office = officeEmps
    .map((e) => buildHourlyBlock(e, true))
    .sort((a, b) => byText(firstNameOf(a.employeeName), firstNameOf(b.employeeName)));

  // ----- PS blocks -------------------------------------------------------------
  // Round 8 §8.3–§8.5 — metrics come from the Settings account mapping, so
  // sundries are out of materials and travel is available.
  const plByJob = plMetricsByJob(pl, settings.plAccountMapping, entity);

  /**
   * Round 8 §9.4 — a lead with no entry in the crew profit-share % map is
   * auto-added at the seeded default instead of leaving a blank WO%. The
   * addition is recorded here and persisted by the App layer, keeping this
   * build a pure function.
   */
  const newLeadsToAdd = new Map<string, NewLeadDefault>();
  const leadMapPct = (leadUpper: string): number => {
    const entry = settings.profitShare.find(
      (e) => tokensSubset(e.lead, leadUpper) || tokensSubset(leadUpper, e.lead),
    );
    if (entry) return entry.pct;
    if (!newLeadsToAdd.has(leadUpper)) {
      newLeadsToAdd.set(leadUpper, {
        lead: titleCaseLeadName(leadUpper, directory, settings.subcontractors),
        pct: settings.defaultProfitSharePct,
      });
    }
    return settings.defaultProfitSharePct;
  };

  /** Lead's own timesheet hours on an AP# (QB-adjusted). */
  const leadHoursOnAp = (leadUpper: string, apKey: string | null): number => {
    if (!apKey) return 0;
    const emp = leadEmps.find(
      (e) => tokensSubset(e.fullName, leadUpper) || tokensSubset(leadUpper, e.fullName),
    );
    if (!emp) return 0;
    return round2((empJobs.get(emp.username) ?? []).filter((j) => j.apKey === apKey).reduce((a, j) => a + j.hours, 0));
  };

  const jobTextForAp = (apKey: string | null, customer: string, jobShort: string): string => {
    const eventTitle = apKey ? (eventByApKey.get(apKey)?.event ?? '') : '';
    return `${customer} ${jobShort} ${eventTitle}`;
  };

  const buildPsJob = (
    blockKey: string,
    jobRowId: string,
    apKey: string | null,
    jobShort: string,
    customer: string,
    pivotWoPct: number | null,
    pivotWoOverride: number | null,
    leadDefaultWoPct: number | null,
    percent: PercentValue,
    split: number,
    colorHex: string | null,
    leadUpper: string,
    isSub: boolean,
  ): PsJobColumn => {
    const e = edits[`ps:${blockKey}:${jobRowId}`] ?? {};
    const metrics = apKey ? plByJob.get(apKey) : undefined;
    const missingFromPl = !metrics;
    const income = e.income !== undefined ? e.income : (metrics?.income ?? 0);
    const materials = e.materials !== undefined ? e.materials : (metrics?.materials ?? 0);
    // Round 9 §1 — labor comes from the TIMESHEET hours on this AP# (hours ×
    // resolved rate, overtime at 1.5×). The P&L `laborWagesSubs` stays the
    // fallback for a job with no timesheet hours this week, and is still shown
    // verbatim on the `QB Labor-Wages-Subs` row below.
    const timesheetLabor = apKey ? (timesheetLaborByAp.get(apKey) ?? 0) : 0;
    const laborWages =
      e.laborWages !== undefined
        ? e.laborWages
        : timesheetLabor > 0
          ? timesheetLabor
          : (metrics?.laborWagesSubs ?? 0);
    const mileage = e.mileage !== undefined ? e.mileage : (metrics?.mileage ?? 0);
    const travel = metrics?.travel ?? 0;
    const sundries = metrics?.sundries ?? 0;
    // Round 9 §32 — a lead responsible for ALL of a job's materials is charged
    // the full amount instead of half.
    const fullMaterials = (settings.fullMaterialsLeads ?? []).some(
      (n) => tokensSubset(n, leadUpper) || tokensSubset(leadUpper, n),
    );
    /**
     * Round 10 — Utah hotel charge-back. On the UT tab a full-materials lead
     * also carries their own lodging, so the job's P&L travel becomes a
     * "Hotels" line that comes OUT of the chain.
     */
    const hotels = isUt && fullMaterials ? travel : null;
    const halfMaterials =
      e.halfMaterials !== undefined ? e.halfMaterials : fullMaterials ? materials : materials / 2;
    // Round 6 §6.2 — WO% from the Pivot by Job `WO%` cell (Round 9 §15: that
    // column is the source of truth), else the Settings crew map (Round 8 §9.4:
    // the crew map now always answers, auto-adding if needed). Round 9 §14 —
    // a typed % on the preview band header wins over both.
    const effectiveWoPct =
      e.woPct !== undefined ? e.woPct : (pivotWoPct ?? leadDefaultWoPct);
    const pct = effectiveWoPct ?? 0;
    const workOrderOverride =
      e.workOrderOverride !== undefined ? e.workOrderOverride : pivotWoOverride;
    // Round 9 §8/§20 — the Pivot by Job `WO Override` amount IS the work-order
    // dollar figure, so it drives Profit Share whenever it is filled in. Only
    // an empty override falls back to Total Income × WO%, which is what left
    // AP24097-E at $0 when its P&L income was zero.
    const profitShare =
      workOrderOverride !== null && workOrderOverride !== undefined
        ? workOrderOverride
        : income * pct;
    // A typed % Completion on the payroll band overrides the grid value.
    const effectivePercent: PercentValue =
      e.percentCompletion !== undefined ? e.percentCompletion : percent;
    const chain = computeTotalProfitShare({
      profitShare,
      workOrderOverride,
      laborWages,
      // Round 10 — the hotel charge-back rides with the materials deduction.
      halfMaterials: halfMaterials + (hotels ?? 0),
      percent: effectivePercent,
      mileage,
    });
    const percentCompletion = chain.percentCompletion;
    // SUBBED / SET RATE pay nothing automatically; the band cell renders blank
    // and the amount is entered by hand.
    let totalProfitShare = chain.total ?? 0;
    const splitWith = Math.max(1, split);
    let totalAfterSplit = totalProfitShare / splitWith;

    // Round 6 §7.1 Branch 2 — Davis-Bacon floor after chain + split.
    const jobText = jobTextForAp(apKey, customer, jobShort);
    const pbjJobForDb = apKey
      ? pbj.customers.flatMap((c) => c.jobs).find((j) => j.analysis.apKey === apKey)
      : undefined;
    const isDbJob =
      pbjJobForDb !== undefined
        ? pbjJobForDb.analysis.isDbJob
        : isPrevailingWageJob(jobText, settings.prevailingWagePatterns);
    const hoursOnJob = leadHoursOnAp(leadUpper, apKey);
    const floorResult = applyDavisBaconLeadFloor({
      // A hand-paid SUBBED / SET RATE job has no computed chain to floor.
      isDbJob: isDbJob && chain.manualPay === null,
      isSubcontractor: isSub,
      timesheetHoursOnJob: hoursOnJob,
      profitSharePay: totalAfterSplit,
      dbRate: settings.prevailingMinRate,
    });
    if (floorResult.floorApplied) {
      totalAfterSplit = floorResult.pay;
      if (splitWith <= 1) totalProfitShare = floorResult.pay;
    }

    const pbjNotes = (() => {
      if (apKey) {
        const job = pbjJobByAp.get(apKey);
        if (job?.analysis.notes) return job.analysis.notes;
        const carried = utJobByAp.get(apKey) ?? nvJobByAp.get(apKey);
        if (carried?.notes) return carried.notes;
      }
      return '';
    })();

    return {
      jobRowId,
      apKey,
      colorHex,
      jobShort,
      customer,
      woPct: effectiveWoPct,
      profitShare,
      workOrderOverride,
      percentCompletion,
      percent: effectivePercent,
      manualPay: chain.manualPay,
      missingPercent: chain.missingPercent,
      laborWages,
      mileage,
      travel,
      sundries,
      halfMaterials,
      fullMaterials,
      hotels,
      timesheetLabor,
      totalProfitShare,
      splitWith,
      totalAfterSplit,
      ...(floorResult.floorApplied
        ? {
            dbFloorApplied: true,
            dbFloor: floorResult.floor,
            dbFloorPriorPay: floorResult.priorPay,
            dbFloorHours: hoursOnJob,
          }
        : {}),
      missingFromPl,
      income,
      materials,
      notes: pbjNotes,
    };
  };

  /**
   * Round 9 §4/§7 — a band cell must never be blank. The PO label falls back
   * through the display chain → the Pivot by Job job text → `AP{key}`, and the
   * customer name prefers the Pivot by Job customer column (which is where the
   * UT rows carry their customer) before the resolved chain. A customer with
   * several AP jobs therefore labels every column, not just the first.
   */
  const bandNames = (
    apKey: string | null,
    display: { poLabel: string; customerName: string },
    pivotJobText: string,
    pivotCustomer: string,
  ): { poLabel: string; customerName: string } => {
    const po =
      display.poLabel.trim() ||
      pivotJobText.trim() ||
      (apKey && apKey !== 'OFFICE' ? `AP${apKey}` : '');
    const cust =
      pivotCustomer.trim() && pivotCustomer.trim() !== '(blank)'
        ? pivotCustomer.trim()
        : display.customerName.trim() ||
          (apKey ? (plByJob.has(apKey) ? (pl?.byJob.get(apKey)?.customerName ?? '') : '') : '');
    return { poLabel: po, customerName: cust };
  };

  const leads: PsBlock[] = [];
  const leadMatchesName = (candidate: string, leadUpper: string): boolean =>
    candidate.toUpperCase() === leadUpper ||
    tokensSubset(candidate, leadUpper) ||
    tokensSubset(leadUpper, candidate);

  for (const [leadUpper, meta] of leadNames) {
    // A lead working in both states gets a band on each tab; the UT band needs
    // its own edit namespace so its Reimbursements/Deductions stay separate.
    const blockKey = isUt ? `lead:UT:${leadUpper}` : `lead:${leadUpper}`;
    const jobs: PsJobColumn[] = [];
    const seenJobIds = new Set<string>();
    const seenAps = new Set<string>();
    const markJob = (jobId: string, apKey: string | null): boolean => {
      if (seenJobIds.has(jobId)) return false;
      if (apKey && seenAps.has(apKey)) return false;
      seenJobIds.add(jobId);
      if (apKey) seenAps.add(apKey);
      return true;
    };

    for (const { job } of pieceRateJobs) {
      const names =
        job.analysis.crewLeadResolution && job.analysis.crewLeadResolution.names.length > 0
          ? job.analysis.crewLeadResolution.names
          : job.analysis.crewLead
            ? [job.analysis.crewLead.toUpperCase()]
            : [];
      if (!names.some((n) => leadMatchesName(n, leadUpper))) continue;
      if (!markJob(job.id, job.analysis.apKey)) continue;
      const display = bandNames(
        job.analysis.apKey,
        jobDisplayPair(job.analysis.apKey, {
          ...nameSources,
          eventTitle: eventByApKey.get(job.analysis.apKey ?? '')?.event,
        }),
        job.jobcode2,
        job.id.split('\u0000')[0],
      );
      jobs.push(
        buildPsJob(
          blockKey,
          job.id,
          job.analysis.apKey,
          display.poLabel,
          display.customerName,
          job.analysis.woPct,
          job.analysis.woOverride,
          leadMapPct(leadUpper),
          job.analysis.percent,
          job.analysis.split,
          job.colorHex,
          leadUpper,
          meta.sub,
        ),
      );
    }
    for (const cj of calOnlyJobs) {
      if (!cj.lead.some((n) => leadMatchesName(n, leadUpper))) continue;
      const calId = `cal:${cj.row.id}`;
      if (!markJob(calId, cj.row.apKey)) continue;
      const display = bandNames(
        cj.row.apKey,
        jobDisplayPair(cj.row.apKey, { ...nameSources, eventTitle: cj.row.event }),
        '',
        '',
      );
      jobs.push(
        buildPsJob(
          blockKey,
          calId,
          cj.row.apKey,
          display.poLabel,
          display.customerName,
          cj.woPct,
          cj.woOverride,
          leadMapPct(leadUpper),
          cj.row.percent,
          cj.split,
          colorByEstimator.get(cj.row.estimator) ?? null,
          leadUpper,
          meta.sub,
        ),
      );
    }

    // Round 8 §6 / §7.3 — carried no-timesheet jobs (NV BALANCES or UT PAYROLL)
    // join this lead's band. HOURLY jobs stay off the PS chain.
    for (const row of isUt ? utJobs : nvBalanceJobs) {
      if (carriedIsHourly(row)) continue;
      const names = leadNamesFromDisplay(row.crewLead);
      if (!names.some((n) => leadMatchesName(n, leadUpper))) continue;
      if (!markJob(row.id, row.apKey)) continue;
      const display = bandNames(
        row.apKey,
        jobDisplayPair(row.apKey, {
          ...nameSources,
          eventTitle: eventByApKey.get(row.apKey ?? '')?.event,
        }),
        row.job,
        row.customer,
      );
      jobs.push(
        buildPsJob(
          blockKey,
          row.id,
          row.apKey,
          display.poLabel,
          display.customerName,
          row.woPct,
          row.woOverride,
          leadMapPct(leadUpper),
          row.percent,
          Math.max(1, names.length),
          row.colorHex,
          leadUpper,
          meta.sub,
        ),
      );
    }

    const leadEmp = leadEmps.find(
      (e) => tokensSubset(e.fullName, leadUpper) || tokensSubset(leadUpper, e.fullName),
    );

    // Crew-owner / profitShare leads with no parenthetical jobs: attach their
    // own piece-rate timesheet APs so they still get a PS band (not a painter block).
    if (jobs.length === 0 && leadEmp) {
      for (const ej of empJobs.get(leadEmp.username) ?? []) {
        if (ej.hourly || ej.office || !ej.apKey) continue;
        const pbjJob = jobRowById.get(ej.jobId);
        const utRow = utJobByAp.get(ej.apKey);
        if (pbjJob && (pbjJob.analysis.hourly || pbjJob.analysis.apKey === 'OFFICE')) continue;
        if (!pbjJob && !utRow) continue;
        if (!markJob(ej.jobId, ej.apKey)) continue;
        const apKey = pbjJob?.analysis.apKey ?? ej.apKey;
        const display = bandNames(
          apKey,
          jobDisplayPair(apKey, {
            ...nameSources,
            eventTitle: eventByApKey.get(apKey ?? '')?.event,
          }),
          pbjJob?.jobcode2 || utRow?.job || ej.jobcode2,
          utRow?.customer || ej.customer,
        );
        jobs.push(
          buildPsJob(
            blockKey,
            ej.jobId,
            apKey,
            display.poLabel,
            display.customerName,
            pbjJob?.analysis.woPct ?? utRow?.woPct ?? null,
            pbjJob?.analysis.woOverride ?? utRow?.woOverride ?? null,
            leadMapPct(leadUpper),
            pbjJob?.analysis.percent ?? utRow?.percent ?? null,
            pbjJob?.analysis.split ?? 1,
            pbjJob?.colorHex ?? utRow?.colorHex ?? ej.colorHex,
            leadUpper,
            meta.sub,
          ),
        );
      }
    }

    // Skip pure name-only entries with neither jobs nor timesheet presence.
    if (jobs.length === 0 && !leadEmp) continue;

    const tsheets = leadEmp && !meta.sub ? buildHourlyBlock(leadEmp, false, true) : null;
    const be = edits[`psb:${blockKey}`] ?? {};
    const total = jobs.reduce((a, j) => a + j.totalAfterSplit, 0);
    const mileageNew = be.mileageNew !== undefined ? be.mileageNew : 0;
    const hourlyPay =
      be.hourlyPay !== undefined ? be.hourlyPay : tsheets ? tsheets.totalPay : 0;
    const reimbursements = be.reimbursements !== undefined ? be.reimbursements : 0;
    // Round 8 §8.3 — the P&L sundries left materials; they are the DEFAULT
    // Deductions/Sundries entry now (a manual edit still wins).
    const sundriesDefault = round2(jobs.reduce((a, j) => a + j.sundries, 0));
    const deductions = be.deductions !== undefined ? be.deductions : sundriesDefault;
    const deductMemo =
      be.deductMemo ?? (be.deductions === undefined && sundriesDefault !== 0 ? 'Sundries (P&L)' : '');
    // Round 9 §19/§23 — extra rows on the lead's band (Aflac, Gas Reimb, …).
    const leadDisplayName = titleCaseLeadName(leadUpper, directory, settings.subcontractors);
    const extras = extrasFor(leadDisplayName || leadUpper, `psb:${blockKey}`);
    leads.push({
      key: blockKey,
      leadName: leadUpper,
      displayName: leadDisplayName,
      sub: meta.sub,
      jobs,
      total,
      mileageNew,
      hourlyPay,
      reimbursements,
      reimbMemo: be.reimbMemo ?? '',
      deductions,
      deductMemo,
      sundriesDefault,
      // Round 8 §8.2 — deductions are entered POSITIVE and SUBTRACTED; every
      // mileage row (Mileage - New, -Mileage, Mileage in Books) is display-only.
      // Round 9 §23 — configured deductions/reimbursements move Total Pay too.
      totalPay: round2(
        computeTotalPay({ total, hourlyPay, reimbursements, deductions }) + extrasNet(extras),
      ),
      tsheets,
      extras,
      // Round 9 §31 — hours pay + profit share, for the Final Payroll Breakdown.
      totalGrossWages: round2(total + hourlyPay),
    });
  }
  leads.sort((a, b) => byText(firstNameOf(a.displayName), firstNameOf(b.displayName)));

  // §5.2 — one alphabetical sequence: painters + lead blocks interleaved A→Z
  // by first name; subcontractor lead blocks last (A→Z among themselves).
  const painterSequence: PainterSectionEntry[] = [
    ...painters.map((b): PainterSectionEntry => ({ kind: 'hourly', block: b })),
    ...leads.filter((b) => !b.sub).map((b): PainterSectionEntry => ({ kind: 'ps', block: b })),
  ].sort((a, b) => byText(firstNameOf(sequenceName(a)), firstNameOf(sequenceName(b))));
  for (const b of leads.filter((x) => x.sub)) painterSequence.push({ kind: 'ps', block: b });

  // ----- salary sections ---------------------------------------------------
  const salaryBlock = (entry: AppSettings['salariedRoster'][number]): SalaryBlock => {
    const key = `sal:${entry.name}`;
    const e = edits[key] ?? {};
    const tsEmp = [...emps.values()].find(
      (x) => tokensSubset(x.fullName, entry.name) || tokensSubset(entry.name, x.fullName),
    );
    const hours = tsEmp
      ? round2((empJobs.get(tsEmp.username) ?? []).reduce((a, j) => a + j.hours, 0))
      : 0;
    const miles = e.miles !== undefined ? e.miles : 0;
    const mileageReimb = round2(miles * settings.mileageRate);
    // Round 9 §28 — Jodi Dunn's row reads "Reimbursements" and carries a flat
    // amount instead of miles × rate.
    const flatReimb = entry.reimbursementsLabel
      ? (e.reimbursements !== undefined ? e.reimbursements : 0)
      : 0;
    const extras = extrasFor(entry.name, key);
    return {
      key,
      name: stripSingleLetterTokens(entry.name),
      role: entry.role,
      amount: entry.amount,
      section: entry.section,
      timesheetHours: hours,
      miles,
      mileageReimb,
      sub: entry.sub === true,
      // Round 10 — these people carry no mileage-reimbursement row at all.
      noMileageRow: entry.noMileageRow === true,
      ...(entry.reimbursementsLabel ? { reimbursementsLabel: entry.reimbursementsLabel } : {}),
      reimbursements: flatReimb,
      extras,
      taxes: e.taxes !== undefined ? e.taxes : (entry.taxes ?? null),
      net: e.net !== undefined ? e.net : (entry.net ?? null),
      total: round2(
        entry.amount +
          (entry.reimbursementsLabel ? flatReimb : entry.noMileageRow ? 0 : mileageReimb) +
          extrasNet(extras),
      ),
    };
  };
  // Round 8 §7.3 — each tab shows only its own entity's salaried people, so the
  // UT tab is painters/leads-only until someone is tagged UT in Settings.
  const roster = settings.salariedRoster.filter(
    (r) =>
      r.entity === entity &&
      // Round 9 §25 — monthly people appear only on the month's first pay date.
      (r.frequency !== 'firstPayDateOfMonth' || isFirstPayDateOfMonth),
  );
  // Round 9 §26 — subcontractor rows sort to the very bottom of their section.
  const sortSalary = (a: SalaryBlock, b: SalaryBlock): number =>
    Number(a.sub) - Number(b.sub) || byText(firstNameOf(a.name), firstNameOf(b.name));
  const admin = roster.filter((r) => r.section === 'ADMIN').map(salaryBlock).sort(sortSalary);
  const estimators = roster.filter((r) => r.section === 'ESTIMATORS').map(salaryBlock).sort(sortSalary);
  const managers = roster.filter((r) => r.section === 'MANAGER SALARY').map(salaryBlock).sort(sortSalary);
  // Round 10 — subcontractor salary rows render with the subcontractor crew-lead
  // bands at the end of the painter sequence, not in their own black section.
  const subcontractorSalaries = roster
    .filter((r) => r.section === 'SUBCONTRACTORS')
    .map(salaryBlock)
    .sort(sortSalary);
  // Round 10 — Kristina Benson & co. render right under Pepe's Painting LLC.
  for (const b of subcontractorSalaries) painterSequence.push({ kind: 'salary', block: b });

  // ----- Round 6 §4.10 Inactive Employees tab --------------------------------
  const norm = (s: string): string => s.trim().toLowerCase();
  const namesWithHours = new Set<string>();
  for (const b of painters) {
    if (b.totalHours > 0) namesWithHours.add(norm(b.employeeName));
  }
  for (const b of office) {
    if (b.totalHours > 0) namesWithHours.add(norm(b.employeeName));
  }
  for (const b of leads) namesWithHours.add(norm(b.leadName));
  for (const entry of painterSequence) {
    if (entry.kind === 'ps') namesWithHours.add(norm(entry.block.leadName));
    else if (entry.kind === 'salary') namesWithHours.add(norm(entry.block.name));
    else if (entry.block.totalHours > 0) namesWithHours.add(norm(entry.block.employeeName));
  }

  const salariedNames = new Set(settings.salariedRoster.map((r) => norm(r.name)));
  const leadNameSet = new Set<string>([
    ...settings.profitShare.map((p) => norm(p.lead)),
    ...settings.subcontractors.map((s) => norm(s)),
  ]);
  const isKnownLead = (name: string): boolean => leadNameSet.has(norm(name));

  interface InactiveCand {
    name: string;
    directoryInactive: boolean;
  }
  const inactiveCands = new Map<string, InactiveCand>();

  for (const d of directory) {
    if (d.active) continue;
    const key = norm(d.fullName);
    if (namesWithHours.has(key) || salariedNames.has(key)) continue;
    inactiveCands.set(key, { name: d.fullName, directoryInactive: true });
  }
  // Round 8 §11 — the Inactive Employees sheet is the persisted painter roster,
  // regenerated every week until someone toggles a name in Settings.
  for (const p of settings.painterRoster) {
    if (p.active) continue;
    const key = norm(p.name);
    if (namesWithHours.has(key) || salariedNames.has(key)) continue;
    if (!inactiveCands.has(key)) inactiveCands.set(key, { name: p.name, directoryInactive: true });
  }

  // The Inactive Employees tab is company-wide, so only the NV model fills it;
  // otherwise the merged export would list every absent painter twice.
  const inactiveEmployees: InactiveEmployeeBlock[] = (isUt ? [] : [...inactiveCands.values()])
    .map((c) => {
      const rateInfo = wageRateFor(c.name, settings.wageRates, settings.defaultWageRate, false);
      return {
        name: c.name,
        displayName: stripSingleLetterTokens(c.name),
        rateNote: `$${rateInfo.rate.toFixed(2)}/hr`,
        isLead: isKnownLead(c.name),
        directoryInactive: c.directoryInactive,
      };
    })
    .sort((a, b) => byText(a.displayName, b.displayName));

  const inactiveZeroHours = inactiveEmployees
    .filter((e) => e.directoryInactive)
    .map((e) => e.name)
    .sort(byText);

  return {
    entity,
    title,
    weekRange,
    painters,
    leads,
    painterSequence,
    office,
    admin,
    estimators,
    managers,
    inactiveEmployees,
    inactiveZeroHours,
    newLeadsToAdd: [...newLeadsToAdd.values()].sort((a, b) => byText(a.lead, b.lead)),
    // Round 10 — roster names that worked this week and should stop being
    // listed as inactive.
    activatedPainters: [...activatedPainters].sort(byText),
    dailyOtNotes,
  };
}

/**
 * Round 8 §7.3 — does this week have any Utah payroll at all? Drives both the
 * `Payroll Preview - UT` empty state and the byte-parity rule that an NV-only
 * week emits no UT sheets.
 */
export function hasUtPayrollData(args: {
  enriched: readonly EnrichedTimesheetRow[];
  pbj: PivotByJobModel;
  plUt: PlParseResult | null;
}): boolean {
  if (args.plUt) return true;
  if (args.pbj.utPayroll.some((r) => !r.excludedFromPayroll)) return true;
  return args.enriched.some((r) => r.sourceEntity === 'UT' && r.hours !== 0);
}

/**
 * Round 8 §7.3 — one model carrying BOTH entities' people, for the downstream
 * shared tabs (Profitshare Output, Input #'s, NOTES, Management Report Export,
 * …). Every band keeps the numbers its own entity's P&L produced; only the
 * ordering is redone so the combined sheets still read A→Z with subs last.
 */
export function mergePayrollPreviews(
  nv: PayrollPreviewModel,
  ut: PayrollPreviewModel | null,
): PayrollPreviewModel {
  if (!ut) return nv;
  const leads = [...nv.leads, ...ut.leads];
  const painters = [...nv.painters, ...ut.painters].sort((a, b) =>
    byText(firstNameOf(a.employeeName), firstNameOf(b.employeeName)),
  );
  const painterSequence: PainterSectionEntry[] = [
    ...painters.map((b): PainterSectionEntry => ({ kind: 'hourly', block: b })),
    ...leads.filter((b) => !b.sub).map((b): PainterSectionEntry => ({ kind: 'ps', block: b })),
  ].sort((a, b) => byText(firstNameOf(sequenceName(a)), firstNameOf(sequenceName(b))));
  for (const b of leads.filter((x) => x.sub)) painterSequence.push({ kind: 'ps', block: b });
  // Round 10 — subcontractor salary blocks stay with the subcontractor bands.
  for (const m of [nv, ut]) {
    for (const e of m.painterSequence) if (e.kind === 'salary') painterSequence.push(e);
  }
  const byFirstName = (a: SalaryBlock, b: SalaryBlock): number =>
    // Round 9 §26 — subcontractor rows stay at the bottom of the merged section.
    Number(a.sub) - Number(b.sub) || byText(firstNameOf(a.name), firstNameOf(b.name));
  const seenLead = new Set(nv.newLeadsToAdd.map((n) => n.lead.trim().toLowerCase()));
  return {
    ...nv,
    leads,
    painters,
    painterSequence,
    office: [...nv.office, ...ut.office].sort((a, b) =>
      byText(firstNameOf(a.employeeName), firstNameOf(b.employeeName)),
    ),
    admin: [...nv.admin, ...ut.admin].sort(byFirstName),
    estimators: [...nv.estimators, ...ut.estimators].sort(byFirstName),
    managers: [...nv.managers, ...ut.managers].sort(byFirstName),
    newLeadsToAdd: [
      ...nv.newLeadsToAdd,
      ...ut.newLeadsToAdd.filter((n) => !seenLead.has(n.lead.trim().toLowerCase())),
    ],
    activatedPainters: [...new Set([...nv.activatedPainters, ...ut.activatedPainters])].sort(byText),
    dailyOtNotes: [...nv.dailyOtNotes, ...ut.dailyOtNotes],
  };
}
