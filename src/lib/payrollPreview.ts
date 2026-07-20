/**
 * Round 3 §5 / Round 4 §5 — "Payroll Preview": one block per person (hourly
 * painters + crew-lead profit-share blocks in a single A→Z sequence, then
 * office/admin/estimator/manager sections). Every derived number is an
 * editable default; edits recompute live.
 */
import { fmtDate, round2 } from './format';
import {
  directoryFromTimesheet,
  extractLeadParenthetical,
  extractWoOverride,
  extractWoPct,
  jobDisplayName,
  profitSharePctFor,
  resolveCrewLead,
  resolveLineRate,
  splitOvertime,
  stripSingleLetterTokens,
  titleCaseLeadName,
  tokensSubset,
  type LeadResolverContext,
  type OtJobHours,
} from './payroll';
import { isHourlyJob, fmtMonthDay } from './pivotByJob';
import type {
  AppSettings,
  DirectoryEmployee,
  EnrichedTimesheetRow,
  HourlyBlock,
  HourlyLine,
  PainterSectionEntry,
  PayrollEditsMap,
  PayrollPreviewModel,
  PivotByJobModel,
  PlJobMetrics,
  PlParseResult,
  PsBlock,
  PsJobColumn,
  QbHoursMap,
  ReportRow,
  SalaryBlock,
} from './types';

export interface PayrollPreviewInput {
  pbj: PivotByJobModel;
  enriched: EnrichedTimesheetRow[];
  reportRows: ReportRow[];
  pl: PlParseResult | null;
  directory: DirectoryEmployee[] | null; // Slot 9 / saved (null → timesheet fallback)
  settings: AppSettings;
  payDate: Date | null;
  edits: PayrollEditsMap;
  qbHours: QbHoursMap;
  knownEstimators: ReadonlySet<string>;
  leadCtx: LeadResolverContext;
}

const byText = (a: string, b: string): number => a.localeCompare(b, 'en', { sensitivity: 'base' });

/** First name for alphabetical block ordering. */
function firstNameOf(name: string): string {
  return name.trim().split(/\s+/)[0] ?? '';
}

export function buildPayrollPreview(input: PayrollPreviewInput): PayrollPreviewModel {
  const { pbj, enriched, reportRows, pl, settings, payDate, edits, qbHours } = input;
  const directory = input.directory ?? directoryFromTimesheet(enriched);
  const colorByEstimator = new Map(settings.legend.map((l) => [l.estimator, l.color]));

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
  const title = `PAYROLL BREAKDOWN ${payDate ? fmtDate(payDate) : ''} ${range}`.replace(/\s+/g, ' ').trim();

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
  const emps = new Map<string, Emp>();
  for (const r of enriched) {
    const key = r.username || `${r.fname} ${r.lname}`;
    if (emps.has(key)) continue;
    const tsName = `${r.fname} ${r.lname}`.replace(/\s+/g, ' ').trim();
    const dirMatch = directory.find((d) => tokensSubset(tsName, d.fullName) || tokensSubset(d.fullName, tsName));
    const fullName = dirMatch?.fullName ?? tsName;
    emps.set(key, {
      username: key,
      fname: r.fname,
      lname: r.lname,
      group: r.group,
      fullName,
      displayName: stripSingleLetterTokens(fullName),
      inactive: dirMatch ? !dirMatch.active : false,
    });
  }

  // ----- per (employee, job) QB-adjusted hours ------------------------------
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
  }
  const empJobs = new Map<string, EmpJob[]>(); // username → jobs in first-encounter order
  const jobRowById = new Map(pbj.customers.flatMap((c) => c.jobs.map((j) => [j.id, j] as const)));
  const eventByApKey = new Map<string, ReportRow>();
  for (const r of reportRows) {
    if (r.included && r.apKey && !eventByApKey.has(r.apKey)) eventByApKey.set(r.apKey, r);
  }

  enriched.forEach((r, idx) => {
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
      const pbjJob = jobRowById.get(jobId);
      const eventTitle = r.apKey ? (eventByApKey.get(r.apKey)?.event ?? '') : '';
      j = {
        jobId,
        apKey: r.apKey,
        customer: r.jobcode1,
        jobcode2: r.jobcode2,
        hours: 0,
        hourly: pbjJob?.analysis.hourly ?? false,
        colorHex: r.apKey === 'OFFICE' ? '4F81BD' : (pbjJob?.colorHex ?? colorByEstimator.get(r.estimator) ?? null),
        firstIndex: idx,
        lastDate: r.localDate,
        lastRowIndex: idx,
        office: r.apKey === 'OFFICE',
        jobText: `${r.jobcode1} ${r.jobcode2} ${eventTitle}`,
      };
      list.push(j);
    }
    j.hours += r.hours;
    const t = r.localDate ? r.localDate.getTime() : -Infinity;
    const bt = j.lastDate ? j.lastDate.getTime() : -Infinity;
    if (t > bt || (t === bt && idx > j.lastRowIndex)) {
      j.lastDate = r.localDate;
      j.lastRowIndex = idx;
    }
  });

  // Apply QB Hrs overrides (§3.3): keyed `${jobRowId}\u0000${username}`.
  for (const [username, jobs] of empJobs) {
    for (const j of jobs) {
      const o = qbHours[`${j.jobId}\u0000${username}`];
      if (o !== undefined) j.hours = o;
    }
  }

  // ----- crew leads this week ----------------------------------------------
  // Names resolved on piece-rate NV pivot jobs + sub-only calendar jobs.
  const leadNames = new Map<string, { sub: boolean }>(); // UPPER full name
  const pieceRateJobs = pbj.customers.flatMap((c) =>
    c.jobs
      .filter((j) => !j.analysis.hourly && j.analysis.apKey !== 'OFFICE')
      .map((j) => ({ job: j, customer: c.customer })),
  );
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
  for (const r of reportRows) {
    if (!r.included || !r.apKey || timesheetApKeys.has(r.apKey)) continue;
    if (/^utah\b|^reno\b/i.test(r.calendar.trim())) continue; // UT pays via Pivot by Job
    const res = resolveCrewLead(extractLeadParenthetical(r.event, input.knownEstimators), input.leadCtx);
    if (res.names.length === 0) continue;
    if (isHourlyJob(r.percent, res.display, r.event)) continue;
    const woSrc = extractWoOverride(r.description);
    const pctSrc = extractWoPct(r.description);
    calOnlyJobs.push({
      row: r,
      lead: res.names,
      sub: res.sub,
      split: res.split,
      woOverride: woSrc?.amount ?? null,
      woPct:
        pctSrc?.pct ??
        profitSharePctFor(res.names[0], settings.profitShare, settings.defaultProfitSharePct),
    });
    for (const n of res.names) {
      const up = n.toUpperCase();
      if (!leadNames.has(up)) leadNames.set(up, { sub: res.sub });
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

  // ----- hourly block construction ------------------------------------------
  const buildHourlyBlock = (emp: Emp, officeStyle: boolean): HourlyBlock => {
    const jobs = empJobs.get(emp.username) ?? [];
    // §3.4 overtime on the QB-adjusted hours.
    const ot = splitOvertime(
      jobs.map(
        (j): OtJobHours => ({ jobId: j.jobId, hours: j.hours, lastDate: j.lastDate, lastRowIndex: j.lastRowIndex }),
      ),
    );
    const lines: HourlyLine[] = [];
    const pushLine = (j: EmpJob, hours: number, isOt: boolean): void => {
      const lineId = `${emp.username}\u0000${j.jobId}${isOt ? ':OT' : ''}`;
      const e = edits[`hl:${lineId}`] ?? {};
      const rateInfo = resolveLineRate(emp.fullName, j.jobText, j.hourly, settings);
      const baseRate = e.rate !== undefined ? e.rate : rateInfo.rate;
      const rate = isOt ? round2(baseRate * 1.5) : baseRate;
      const isOfficeLine = officeStyle || j.office;
      const label = isOfficeLine
        ? `Office${isOt ? ' [OT]' : ''}`
        : `${jobDisplayName(j.customer, j.jobcode2)}${j.hourly ? ' [hrly]' : ''}${isOt ? ' [OT]' : ''}`;
      // §5.5 — office lines carry the week range (or "Allpro Office" inside a
      // mixed painter block) on the blue PO cell instead of an AP number.
      const poText = officeStyle ? `(${weekRange})` : j.office ? 'Allpro Office' : null;
      lines.push({
        id: lineId,
        apKey: j.apKey,
        poText,
        jobLabel: label,
        hours: e.hours !== undefined ? e.hours : round2(hours),
        rate,
        rateDefaulted: rateInfo.defaulted && e.rate === undefined,
        prevailing: rateInfo.prevailing,
        belowPrevailing: rateInfo.prevailing && baseRate < settings.prevailingMinRate,
        ot: isOt,
        hourlyJob: j.hourly,
        colorHex: j.colorHex,
      });
    };

    if (officeStyle && jobs.length > 0) {
      // §5.3 — single "Office" line covering the week (+ OT).
      const totalHrs = jobs.reduce((a, j) => a + j.hours, 0);
      const regular = ot.otHours > 0 ? round2(totalHrs - ot.otHours) : round2(totalHrs);
      const rep = jobs[0];
      pushLine(rep, regular, false);
      if (ot.otHours > 0) pushLine(rep, ot.otHours, true);
    } else {
      for (const j of jobs.slice().sort((a, b) => a.firstIndex - b.firstIndex)) {
        pushLine(j, ot.regularByJob.get(j.jobId) ?? j.hours, false);
        if (ot.otJobId === j.jobId && ot.otHours > 0) pushLine(j, ot.otHours, true);
      }
    }

    const totalHours = round2(lines.reduce((a, l) => a + l.hours, 0));
    const totalPay = round2(lines.reduce((a, l) => a + l.hours * l.rate, 0));
    const blockEdit = edits[`hb:${emp.username}`] ?? {};
    return {
      key: emp.username,
      title: officeStyle ? emp.displayName : `${emp.displayName}${emp.group ? ` (${emp.group})` : ''}`,
      employeeName: emp.displayName,
      crewTag: officeStyle ? '' : emp.group,
      lines,
      totalHours,
      totalPay,
      net: blockEdit.net !== undefined ? blockEdit.net : null,
      inactive: emp.inactive,
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
  const plByJob = pl?.byJob ?? new Map<string, PlJobMetrics>();

  const buildPsJob = (
    blockKey: string,
    jobRowId: string,
    apKey: string | null,
    jobShort: string,
    customer: string,
    woPct: number | null,
    woOverride: number | null,
    percent: number,
    split: number,
    colorHex: string | null,
  ): PsJobColumn => {
    const e = edits[`ps:${blockKey}:${jobRowId}`] ?? {};
    const metrics = apKey ? plByJob.get(apKey) : undefined;
    const missingFromPl = !metrics;
    const income = e.income !== undefined ? e.income : (metrics?.income ?? 0);
    const materials = e.materials !== undefined ? e.materials : (metrics?.materials ?? 0);
    const laborWages = e.laborWages !== undefined ? e.laborWages : (metrics?.laborWagesSubs ?? 0);
    const mileage = e.mileage !== undefined ? e.mileage : (metrics?.mileage ?? 0);
    // §5.6 precision — full-precision derived values; formats round the display.
    const halfMaterials = e.halfMaterials !== undefined ? e.halfMaterials : materials / 2;
    const pct = woPct ?? 0;
    const profitShare = income * pct;
    const workOrderOverride =
      e.workOrderOverride !== undefined ? e.workOrderOverride : (woOverride ?? profitShare);
    const percentCompletion = e.percentCompletion !== undefined ? e.percentCompletion : percent;
    const totalProfitShare =
      workOrderOverride * percentCompletion - laborWages - mileage - halfMaterials;
    return {
      jobRowId,
      apKey,
      colorHex,
      jobShort,
      customer,
      woPct,
      profitShare,
      workOrderOverride,
      percentCompletion,
      laborWages,
      mileage,
      halfMaterials,
      totalProfitShare,
      splitWith: split,
      totalAfterSplit: totalProfitShare / Math.max(1, split),
      missingFromPl,
      income,
      materials,
    };
  };

  const leads: PsBlock[] = [];
  for (const [leadUpper, meta] of leadNames) {
    const blockKey = `lead:${leadUpper}`;
    const jobs: PsJobColumn[] = [];

    for (const { job } of pieceRateJobs) {
      const names =
        job.analysis.crewLeadResolution && job.analysis.crewLeadResolution.names.length > 0
          ? job.analysis.crewLeadResolution.names
          : job.analysis.crewLead
            ? [job.analysis.crewLead.toUpperCase()]
            : [];
      if (!names.some((n) => n.toUpperCase() === leadUpper)) continue;
      const pctNum = typeof job.analysis.percent === 'number' ? job.analysis.percent : 0;
      jobs.push(
        buildPsJob(
          blockKey,
          job.id,
          job.analysis.apKey,
          job.jobcode2,
          job.id.split('\u0000')[0],
          job.analysis.woPct,
          job.analysis.woOverride,
          pctNum,
          job.analysis.split,
          job.colorHex,
        ),
      );
    }
    for (const cj of calOnlyJobs) {
      if (!cj.lead.some((n) => n.toUpperCase() === leadUpper)) continue;
      const pctNum = typeof cj.row.percent === 'number' ? cj.row.percent : 0;
      jobs.push(
        buildPsJob(
          blockKey,
          `cal:${cj.row.id}`,
          cj.row.apKey,
          cj.row.qboJobName || cj.row.event,
          cj.row.qboJobName,
          cj.woPct,
          cj.woOverride,
          pctNum,
          cj.split,
          colorByEstimator.get(cj.row.estimator) ?? null,
        ),
      );
    }
    if (jobs.length === 0) continue;

    const leadEmp = leadEmps.find(
      (e) => tokensSubset(e.fullName, leadUpper) || tokensSubset(leadUpper, e.fullName),
    );
    const tsheets = leadEmp && !meta.sub ? buildHourlyBlock(leadEmp, false) : null;
    const be = edits[`psb:${blockKey}`] ?? {};
    const total = jobs.reduce((a, j) => a + j.totalAfterSplit, 0);
    const mileageNew = be.mileageNew !== undefined ? be.mileageNew : 0;
    const hourlyPay = tsheets ? tsheets.totalPay : 0;
    const reimbursements = be.reimbursements !== undefined ? be.reimbursements : 0;
    const deductions = be.deductions !== undefined ? be.deductions : 0;
    leads.push({
      key: blockKey,
      leadName: leadUpper,
      displayName: titleCaseLeadName(leadUpper, directory, settings.subcontractors),
      sub: meta.sub,
      jobs,
      total,
      mileageNew,
      hourlyPay,
      reimbursements,
      reimbMemo: be.reimbMemo ?? '',
      deductions,
      deductMemo: be.deductMemo ?? '',
      totalPay: total + mileageNew + hourlyPay + reimbursements + deductions,
      tsheets,
    });
  }
  leads.sort((a, b) => byText(firstNameOf(a.displayName), firstNameOf(b.displayName)));

  // §5.2 — one alphabetical sequence: painters + lead blocks interleaved A→Z
  // by first name; subcontractor lead blocks last (A→Z among themselves).
  const painterSequence: PainterSectionEntry[] = [
    ...painters.map((b): PainterSectionEntry => ({ kind: 'hourly', block: b })),
    ...leads.filter((b) => !b.sub).map((b): PainterSectionEntry => ({ kind: 'ps', block: b })),
  ].sort((a, b) =>
    byText(
      firstNameOf(a.kind === 'hourly' ? a.block.employeeName : a.block.displayName),
      firstNameOf(b.kind === 'hourly' ? b.block.employeeName : b.block.displayName),
    ),
  );
  for (const b of leads.filter((x) => x.sub)) painterSequence.push({ kind: 'ps', block: b });

  // ----- salary sections ---------------------------------------------------
  const salaryBlock = (entry: AppSettings['salariedRoster'][number]): SalaryBlock => {
    const e = edits[`sal:${entry.name}`] ?? {};
    const tsEmp = [...emps.values()].find(
      (x) => tokensSubset(x.fullName, entry.name) || tokensSubset(entry.name, x.fullName),
    );
    const hours = tsEmp
      ? round2((empJobs.get(tsEmp.username) ?? []).reduce((a, j) => a + j.hours, 0))
      : 0;
    const miles = e.miles !== undefined ? e.miles : 0;
    return {
      key: `sal:${entry.name}`,
      name: stripSingleLetterTokens(entry.name),
      role: entry.role,
      amount: entry.amount,
      section: entry.section,
      timesheetHours: hours,
      miles,
      mileageReimb: round2(miles * settings.mileageRate),
    };
  };
  const roster = settings.salariedRoster;
  const admin = roster.filter((r) => r.section === 'ADMIN').map(salaryBlock)
    .sort((a, b) => byText(firstNameOf(a.name), firstNameOf(b.name)));
  const estimators = roster.filter((r) => r.section === 'ESTIMATORS').map(salaryBlock)
    .sort((a, b) => byText(firstNameOf(a.name), firstNameOf(b.name)));
  const managers = roster.filter((r) => r.section === 'MANAGER SALARY').map(salaryBlock)
    .sort((a, b) => byText(firstNameOf(a.name), firstNameOf(b.name)));

  // ----- inactive employees with zero hours (feeds Payroll Exceptions only) --
  const namesWithHours = new Set([...emps.values()].map((e) => e.fullName.toLowerCase()));
  const inactiveZeroHours = directory
    .filter((d) => !d.active && !namesWithHours.has(d.fullName.toLowerCase()))
    .map((d) => d.fullName)
    .sort(byText);

  return {
    title,
    weekRange,
    painters,
    leads,
    painterSequence,
    office,
    admin,
    estimators,
    managers,
    inactiveZeroHours,
  };
}
