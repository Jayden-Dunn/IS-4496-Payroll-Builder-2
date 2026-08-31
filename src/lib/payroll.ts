/**
 * Round 3 shared payroll derivations (§3): job display names, crew-lead
 * extraction/resolution, overtime split, WO pre-fills. All pure functions —
 * nothing dataset-specific lives here; tribal knowledge comes from Settings.
 */
import { AP_RE, extractApKey, normalizeApInput } from './apkey';
import { isPercentNumber, isPercentState, round2 } from './format';
import { baseApKey } from './jobNames';
import type { PayRateKind } from './payTypes';
import type {
  ConditionalRateRule,
  CrewLeadResolution,
  DirectoryEmployee,
  JobRateRule,
  PercentValue,
  PerJobRateOverride,
  TimesheetRow,
} from './types';

// ---------------------------------------------------------------------------
// §3.1 Job display name
// ---------------------------------------------------------------------------

/** "Aaron Ward" + "AP24314 Ext" → "Aaron Ward Ext" (single-spaced). */
export function jobDisplayName(jobcode1: string, jobcode2: string): string {
  const jobText = jobcode2.replace(AP_RE, ' ').replace(/\s+/g, ' ').trim();
  return `${jobcode1.trim()} ${jobText}`.replace(/\s+/g, ' ').trim();
}

// ---------------------------------------------------------------------------
// Name tokens (shared by lead matching and directory lookup)
// ---------------------------------------------------------------------------

/** Lower-cased name tokens; hyphens split; apostrophes kept inside tokens. */
export function nameTokens(s: string): string[] {
  return s
    .toLowerCase()
    .split(/[\s\-,]+/)
    .map((t) => t.replace(/[.*]/g, ''))
    .filter(Boolean);
}

/**
 * Token set of a full name with trailing single-letter middle initials
 * treated as optional ("Antonio J Zuniga" → required {antonio, zuniga},
 * optional {j}).
 */
export function nameTokenSets(fullName: string): { all: Set<string>; required: Set<string> } {
  const tokens = nameTokens(fullName);
  const all = new Set(tokens);
  const required = new Set(tokens.filter((t) => t.length > 1));
  return { all, required };
}

/** Do all candidate tokens appear in the employee's token set? */
export function tokensSubset(candidate: string, employeeFullName: string): boolean {
  const cand = nameTokens(candidate);
  if (cand.length === 0) return false;
  const { all } = nameTokenSets(employeeFullName);
  return cand.every((t) => all.has(t));
}

/**
 * Round 4 §3.2-2 — display names drop single-letter tokens (optionally
 * followed by '.') anywhere in the name: "Rene Salazar M" → "Rene Salazar",
 * "Adriana G Lopez Barajas" → "Adriana Lopez Barajas". Matching logic keeps
 * the raw tokens; only display changes.
 */
export function stripSingleLetterTokens(name: string): string {
  const kept = name
    .split(/\s+/)
    .filter((t) => !/^[A-Za-z]\.?$/.test(t));
  return (kept.length > 0 ? kept : name.split(/\s+/)).join(' ').trim();
}

/**
 * Resolve an ALL-CAPS lead key (possibly "A/B" joined) back to the stored
 * Title Case names from the directory / subcontractors list (§5.3). Never
 * invents casing — an unmatched part keeps its original text.
 */
export function titleCaseLeadName(
  leadDisplay: string,
  directory: readonly DirectoryEmployee[],
  subcontractors: readonly string[],
): string {
  return leadDisplay
    .split('/')
    .map((part) => {
      const p = part.trim();
      if (!p) return p;
      const dir = directory.find((d) => d.fullName.toUpperCase() === p.toUpperCase());
      if (dir) return stripSingleLetterTokens(dir.fullName);
      const sub = subcontractors.find((s) => s.toUpperCase() === p.toUpperCase());
      if (sub) return sub;
      // No directory hit — still Title Case so auto-added leads aren't ALL CAPS.
      const titled = p.replace(/[A-Za-z]+(?:'[A-Za-z]+)?/g, (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
      return stripSingleLetterTokens(titled);
    })
    .join('/');
}

// ---------------------------------------------------------------------------
// §3.2 Crew-lead extraction from the event title
// ---------------------------------------------------------------------------

/**
 * The calendar names the crew lead in the LAST parenthesized group of the
 * event title — unless that group is the estimator tag (2–3 letters matching
 * a known estimator). Returns the raw inner text, or null.
 */
export function extractLeadParenthetical(title: string, knownEstimators: ReadonlySet<string>): string | null {
  const groups = [...title.matchAll(/\(([^()]*)\)/g)].map((m) => m[1].trim()).filter(Boolean);
  if (groups.length === 0) return null;
  const last = groups[groups.length - 1];
  if (/^[A-Za-z]{2,3}$/.test(last) && knownEstimators.has(last.toUpperCase())) return null;
  return last;
}

/** Strip decorations ("hrly_", "*") and split into lead-name candidates. */
export function splitLeadCandidates(raw: string): string[] {
  return raw
    .split(/[/,&]/)
    .map((s) => s.replace(/\b[a-z0-9]+_/gi, '').replace(/\*/g, '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

// ---------------------------------------------------------------------------
// §3.2 Full-name resolution
// ---------------------------------------------------------------------------

export interface LeadResolverContext {
  directory: DirectoryEmployee[];
  /** Timesheet `group` labels (for the "is a crew lead" preference). */
  crewGroups: string[];
  subcontractors: string[];
}

/** Strip the "'s Crew" suffix from a group label; null when absent. */
function crewLabelName(label: string): string | null {
  const m = /^(.*?)['\u2019]s\s+crew\s*$/i.exec(label.trim());
  return m ? m[1].trim() : null;
}

/** Employee name token-matches some crew-group label (Round-2 lead rule). */
function isKnownCrewLead(employeeFullName: string, crewGroups: string[]): boolean {
  const { all } = nameTokenSets(employeeFullName);
  for (const g of crewGroups) {
    const name = crewLabelName(g);
    if (!name) continue;
    const labelTokens = nameTokens(name);
    if (labelTokens.length > 0 && labelTokens.every((t) => all.has(t))) return true;
  }
  return false;
}

/**
 * Resolve one candidate (a first name, full name, or business name) to a
 * directory employee or subcontractor. Never invents a name.
 */
function resolveCandidate(
  candidate: string,
  ctx: LeadResolverContext,
): {
  name: string | null;
  sub: boolean;
  flags: CrewLeadResolution['flags'];
} {
  const flags: CrewLeadResolution['flags'] = [];
  const candLower = candidate.toLowerCase();
  const candTokens = nameTokens(candidate);

  let matches = ctx.directory.filter((e) => {
    if (candTokens.length > 1) return tokensSubset(candidate, e.fullName);
    const firstTok = nameTokens(e.firstName)[0] ?? '';
    return firstTok === candLower;
  });

  if (matches.length === 0) {
    // §3.2.4 — Subcontractors list by token match.
    const subMatch = ctx.subcontractors.find(
      (s) => tokensSubset(candidate, s) || s.toLowerCase() === candLower,
    );
    if (subMatch) return { name: subMatch.toUpperCase(), sub: true, flags };
    flags.push({ kind: 'noMatch', candidate });
    return { name: null, sub: false, flags };
  }

  // Prefer active matches; the only match being inactive is used but flagged.
  const actives = matches.filter((e) => e.active);
  if (actives.length > 0) {
    matches = actives;
  } else {
    flags.push({ kind: 'inactiveMatch', candidate });
  }

  // Among the pool, prefer known crew leads.
  const leads = matches.filter((e) => isKnownCrewLead(e.fullName, ctx.crewGroups));
  if (leads.length > 0) {
    matches = leads;
  } else {
    flags.push({ kind: 'notACrewLead', candidate });
  }

  if (matches.length > 1) {
    flags.push({ kind: 'ambiguous', candidate });
    matches = [...matches].sort((a, b) => a.fullName.localeCompare(b.fullName, 'en', { sensitivity: 'base' }));
  }
  // §3.2-2 display rule: single-letter tokens dropped ("Rene Salazar M" →
  // "RENE SALAZAR"); matching elsewhere is token-based so nothing breaks.
  return { name: stripSingleLetterTokens(matches[0].fullName).toUpperCase(), sub: false, flags };
}

/**
 * Resolve a raw lead parenthetical ("Ted/Rafael", "Pepe's Painting"…) into
 * the display value + flags. `raw === null` → empty resolution (no flags —
 * the missing-lead case is reported by the caller with job context).
 */
export function resolveCrewLead(raw: string | null, ctx: LeadResolverContext): CrewLeadResolution {
  if (!raw) return { display: '', names: [], split: 1, sub: false, raw: '', flags: [] };
  const candidates = splitLeadCandidates(raw);
  const names: string[] = [];
  const flags: CrewLeadResolution['flags'] = [];
  let sub = false;
  for (const cand of candidates) {
    const res = resolveCandidate(cand, ctx);
    flags.push(...res.flags);
    if (res.name) {
      names.push(res.name);
      if (res.sub) sub = true;
    }
  }
  return {
    display: names.join('/'),
    names,
    split: Math.max(1, names.length),
    sub,
    raw,
    flags,
  };
}

/** Directory fallback built from the timesheet when Slot 9 is empty. */
export function directoryFromTimesheet(rows: TimesheetRow[]): DirectoryEmployee[] {
  const seen = new Map<string, DirectoryEmployee>();
  for (const r of rows) {
    const fullName = `${r.fname} ${r.lname}`.replace(/\s+/g, ' ').trim();
    if (!fullName || seen.has(fullName.toLowerCase())) continue;
    seen.set(fullName.toLowerCase(), {
      fullName,
      firstName: r.fname.trim(),
      lastName: r.lname.trim(),
      active: true,
    });
  }
  return [...seen.values()];
}

// ---------------------------------------------------------------------------
// §3.4 Overtime split
// ---------------------------------------------------------------------------

export interface OtSplit {
  /** Per-job regular hours after the OT reduction, keyed by job-row id. */
  regularByJob: Map<string, number>;
  /** Job-row id carrying the OT line, or null when total ≤ 40. */
  otJobId: string | null;
  otHours: number;
}

export interface OtJobHours {
  jobId: string;
  hours: number;
  /** Latest local_date the employee worked this job. */
  lastDate: Date | null;
  /** File order of the last row for this job (tie-break within the day). */
  lastRowIndex: number;
}

/**
 * If the employee's week exceeds 40 hours, move the overage to an
 * [OT >40 hrs/wk] line on the LAST job worked (latest date; last by file order
 * within the day), reducing that job's regular hours so regular + OT = actual.
 *
 * Round 8 §9.2 — `dailyOtHours` are hours already paid at 1.5× by the NV
 * daily-OT rule. They are excluded here so no hour is paid overtime twice;
 * each hour gets the single highest applicable OT treatment once.
 */
export function splitOvertime(jobs: OtJobHours[], opts?: { dailyOtHours?: number }): OtSplit {
  const regularByJob = new Map<string, number>(jobs.map((j) => [j.jobId, round2(j.hours)]));
  const total = round2(jobs.reduce((acc, j) => acc + j.hours, 0));
  const dailyOtHours = round2(Math.max(0, opts?.dailyOtHours ?? 0));
  if (total <= 40 || jobs.length === 0) return { regularByJob, otJobId: null, otHours: 0 };
  const otHours = round2(Math.max(0, total - 40 - dailyOtHours));
  if (otHours <= 0) return { regularByJob, otJobId: null, otHours: 0 };
  const last = jobs.reduce((best, cur) => {
    const bt = best.lastDate ? best.lastDate.getTime() : -Infinity;
    const ct = cur.lastDate ? cur.lastDate.getTime() : -Infinity;
    if (ct !== bt) return ct > bt ? cur : best;
    return cur.lastRowIndex >= best.lastRowIndex ? cur : best;
  });
  regularByJob.set(last.jobId, round2((regularByJob.get(last.jobId) ?? 0) - otHours));
  return { regularByJob, otJobId: last.jobId, otHours };
}

// ---------------------------------------------------------------------------
// Round 8 §9.2 — NV daily overtime
// ---------------------------------------------------------------------------

/** One (job, day) bucket of an employee's hours. */
export interface DailyOtRow {
  jobId: string;
  date: Date | null;
  hours: number;
  /** File order — the day's overage lands on the latest rows. */
  rowIndex: number;
  /** NV row whose resolved line rate is under `nvDailyOtRateThreshold`. */
  eligible: boolean;
}

/** One daily-OT chunk: hours over the threshold on a single day and job. */
export interface DailyOtSegment {
  jobId: string;
  dateIso: string;
  date: Date | null;
  /** The employee's total hours that day (all jobs). */
  dayHours: number;
  otHours: number;
}

export interface DailyOtSplit {
  /** Hours to move onto a daily-OT line, keyed by job-row id. */
  otByJob: Map<string, number>;
  segments: DailyOtSegment[];
  totalOtHours: number;
}

function dayKeyOf(d: Date | null): string {
  if (!d) return '';
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${d.getFullYear()}-${mm}-${dd}`;
}

/**
 * Round 8 §9.2 — NV only: a painter whose resolved line rate is under the
 * threshold earns 1.5× for hours past `thresholdHours` in a single day.
 *
 * The overage is taken from the day's LAST rows backwards (mirroring the
 * weekly rule's "last job worked"). Hours on an ineligible line — a UT row, or
 * one already at/above the threshold rate — still consume the overage but
 * never become daily OT, so a mixed day cannot invent overtime for the wrong
 * line and UT rows are untouched.
 */
export function splitDailyOvertime(
  rows: readonly DailyOtRow[],
  thresholdHours: number,
): DailyOtSplit {
  const otByJob = new Map<string, number>();
  const segments: DailyOtSegment[] = [];
  let totalOtHours = 0;
  if (!(thresholdHours > 0)) return { otByJob, segments, totalOtHours };

  const byDay = new Map<string, DailyOtRow[]>();
  for (const r of rows) {
    const key = dayKeyOf(r.date);
    if (!key || r.hours <= 0) continue;
    const list = byDay.get(key);
    if (list) list.push(r);
    else byDay.set(key, [r]);
  }

  for (const key of [...byDay.keys()].sort()) {
    const list = byDay.get(key) ?? [];
    const dayHours = round2(list.reduce((a, r) => a + r.hours, 0));
    if (dayHours <= thresholdHours) continue;
    let remaining = round2(dayHours - thresholdHours);
    for (const r of [...list].sort((a, b) => b.rowIndex - a.rowIndex)) {
      if (remaining <= 0) break;
      const take = round2(Math.min(remaining, r.hours));
      remaining = round2(remaining - take);
      if (!r.eligible || take <= 0) continue;
      otByJob.set(r.jobId, round2((otByJob.get(r.jobId) ?? 0) + take));
      segments.push({ jobId: r.jobId, dateIso: key, date: r.date, dayHours, otHours: take });
      totalOtHours = round2(totalOtHours + take);
    }
  }
  return { otByJob, segments, totalOtHours };
}

// ---------------------------------------------------------------------------
// §4.3 WO Override / WO% pre-fills from the event description
// ---------------------------------------------------------------------------

const WO_BEFORE_RE = /(w\.?o\.?|work\s*order)[^\n$]*\$\s*([\d,]+(?:\.\d+)?)/i;
const WO_AFTER_RE = /\$\s*([\d,]+(?:\.\d+)?)[^\n]*?work\s*order/i;

/** Explicit work-order dollar amount in the description, with source line. */
export function extractWoOverride(description: string): { amount: number; line: string } | null {
  for (const raw of description.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line) continue;
    let m = WO_BEFORE_RE.exec(line);
    let amtText = m ? m[2] : null;
    if (!amtText) {
      m = WO_AFTER_RE.exec(line);
      amtText = m ? m[1] : null;
    }
    if (amtText) {
      const n = Number(amtText.replace(/,/g, ''));
      if (Number.isFinite(n) && n > 0) return { amount: n, line };
    }
  }
  return null;
}

const WO_PCT_RE = /@\s*(\d{1,3}(?:\.\d+)?)\s*%\s*PS\b/i;

/** "@ 40% PS" status-line override for the crew profit-share % (§4.3). */
export function extractWoPct(description: string): { pct: number; line: string } | null {
  for (const raw of description.split(/\r?\n/)) {
    const line = raw.trim();
    const m = WO_PCT_RE.exec(line);
    if (m) {
      const n = Number(m[1]);
      if (n >= 0 && n <= 100) return { pct: n / 100, line };
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Profit-share % lookup (Settings §7.1)
// ---------------------------------------------------------------------------

/** Look up a lead's WO% from the Settings map by token match. */
export function profitSharePctFor(
  leadName: string,
  map: { lead: string; pct: number }[],
  fallback: number,
): number {
  const entry = map.find(
    (e) => tokensSubset(e.lead, leadName) || tokensSubset(leadName, e.lead),
  );
  return entry ? entry.pct : fallback;
}

/** Wage rate lookup (Settings §7.2) by token match. */
export function wageRateFor(
  employeeName: string,
  rates: { name: string; rate: number; hourlyRate?: number }[],
  fallback: number,
  hourlyJob: boolean,
): { rate: number; defaulted: boolean } {
  const entry = rates.find(
    (e) => tokensSubset(e.name, employeeName) || tokensSubset(employeeName, e.name),
  );
  if (!entry) return { rate: fallback, defaulted: true };
  return { rate: hourlyJob && entry.hourlyRate !== undefined ? entry.hourlyRate : entry.rate, defaulted: false };
}

// ---------------------------------------------------------------------------
// Round 8 §8.1 — profit-share chain (pure)
// ---------------------------------------------------------------------------

/** A `% Completion` state that means "pay is entered by hand" (Round 8 §8.1). */
export type ManualPayState = 'SUBBED' | 'SET RATE';

export interface TotalProfitShareResult {
  /**
   * null on SUBBED / SET RATE — there is no auto profit share, so the band
   * cells render blank and stay editable for a hand-entered amount.
   */
  total: number | null;
  /** WO Override when entered, else Profit Share (= Total Income × WO%). */
  base: number;
  /** The fraction actually multiplied in (0 when the % is blank or invalid). */
  percentCompletion: number;
  manualPay: ManualPayState | null;
  /**
   * Blank / free-text / HOURLY % on a piece-rate job. The chain pays 0 — a
   * deliberately conservative choice, since silently paying 100% would
   * overpay a job that has barely started.
   */
  missingPercent: boolean;
}

/**
 * Round 8 §8.1 — SUPERSEDES Round 6 §6.2 and Round 7 §21: BOTH branches now
 * multiply the base by % Completion (Round 5's override-only multiply is
 * generalized).
 *
 *   base  = WO Override when entered, else Profit Share (Total Income × WO%)
 *   total = base × %Completion − (− Labor + Wages) − (½ Materials)
 *
 * `-Mileage` is still DISPLAYED on the band but is not a term. HOURLY jobs
 * never reach this chain (they pay hourly per §9).
 */
export function computeTotalProfitShare(args: {
  profitShare: number;
  workOrderOverride: number | null | undefined;
  laborWages: number;
  halfMaterials: number;
  /** The job's `% of Completion` cell for the current week. */
  percent?: PercentValue;
  /** Display-only — must not affect the result. */
  mileage?: number;
}): TotalProfitShareResult {
  void args.mileage;
  const base =
    args.workOrderOverride !== null && args.workOrderOverride !== undefined
      ? args.workOrderOverride
      : args.profitShare;
  const percent = args.percent ?? null;

  if (isPercentState(percent) && percent !== 'HOURLY') {
    return { total: null, base, percentCompletion: 0, manualPay: percent, missingPercent: false };
  }

  const pct = isPercentNumber(percent) ? percent : 0;
  return {
    total: base * pct - args.laborWages - args.halfMaterials,
    base,
    percentCompletion: pct,
    manualPay: null,
    missingPercent: !isPercentNumber(percent),
  };
}

/**
 * Round 8 §8.2 — `Total (post-split) + Hourly Pay + Reimbursements −
 * Deductions/Sundries`.
 *
 * Deductions/Sundries is entered as a POSITIVE amount and subtracted. The
 * mileage rows (`Mileage - New`, `-Mileage`, `Mileage in Books`) are displayed
 * on the band but are not terms.
 */
export function computeTotalPay(args: {
  total: number;
  hourlyPay: number;
  reimbursements: number;
  deductions: number;
}): number {
  return args.total + args.hourlyPay + args.reimbursements - args.deductions;
}

// ---------------------------------------------------------------------------
// Round 6 §7.1 / §7.2 — Davis-Bacon + job-rate resolution
// ---------------------------------------------------------------------------

/** Case-insensitive substring match of any configured pattern on the job text. */
export function isPrevailingWageJob(jobText: string, patterns: readonly string[]): boolean {
  const t = jobText.toLowerCase();
  return patterns.some((p) => p.trim() !== '' && t.includes(p.trim().toLowerCase()));
}

export interface LineRate {
  rate: number;
  defaulted: boolean;
  /** Branch 1 considered: job matched a Davis-Bacon pattern. */
  prevailing: boolean;
  /** A §7.2 job-specific rule supplied the pre–Branch-1 rate. */
  fromJobRule?: boolean;
  /** Optional note from the matched job rule (tooltip). */
  ruleNote?: string;
  /** Round 8 §9.1 — a person + AP# override supplied the rate. */
  fromPerJobOverride?: boolean;
  /** Round 8 §9.1 — a "person X on lead Z's job" rule supplied the rate. */
  fromConditionalRule?: boolean;
  /** Round 8 §9.1 — the lead's own crew-lead rate on a profit-share job. */
  fromCrewLeadRate?: boolean;
  /** Round 9 §18 — a per-painter pay-item rate (Commercial Hourly / Wages 2). */
  fromPayItemRate?: 'commercialHourly' | 'wages2';
}

/** How a job-rate rule's AP# matches a timesheet/calendar apKey. */
function jobRuleApMatch(
  rule: JobRateRule,
  jobApKey: string,
): 'exact' | 'suffix' | null {
  const ruleKey = (normalizeApInput(rule.apKey) ?? rule.apKey.trim()).toUpperCase();
  const jobKey = jobApKey.trim().toUpperCase();
  if (!ruleKey || !jobKey) return null;
  if (ruleKey === jobKey) return 'exact';
  if (!rule.includeSuffixes) return null;
  const ruleBase = baseApKey(ruleKey) ?? ruleKey;
  const jobBase = baseApKey(jobKey) ?? jobKey;
  return ruleBase === jobBase ? 'suffix' : null;
}

function jobRuleAppliesToEmployee(rule: JobRateRule, employeeName: string): boolean {
  if (rule.employeeNames.length === 0) return true;
  return rule.employeeNames.some(
    (n) => tokensSubset(n, employeeName) || tokensSubset(employeeName, n),
  );
}

/**
 * Specificity: exact+named (4) > exact+all (3) > suffix+named (2) > suffix+all (1).
 * Returns null when no rule matches.
 */
export function matchJobRateRule(
  employeeName: string,
  jobApKey: string | null | undefined,
  rules: readonly JobRateRule[] | undefined,
): { rule: JobRateRule; specificity: number } | null {
  if (!jobApKey || !rules || rules.length === 0) return null;
  let best: { rule: JobRateRule; specificity: number } | null = null;
  for (const rule of rules) {
    const ap = jobRuleApMatch(rule, jobApKey);
    if (!ap) continue;
    if (!jobRuleAppliesToEmployee(rule, employeeName)) continue;
    const named = rule.employeeNames.length > 0;
    const specificity =
      ap === 'exact' ? (named ? 4 : 3) : named ? 2 : 1;
    if (!best || specificity > best.specificity) best = { rule, specificity };
  }
  return best;
}

/** Wage-map row as the rate engine needs it (Settings `WageRateEntry`). */
export interface WageRateLike {
  name: string;
  rate: number;
  hourlyRate?: number;
  /** Round 9 §18 — per-painter `Commercial Hourly` rate. */
  commercialHourlyRate?: number;
  /** Round 9 §18 — per-painter `Wages 2` rate. */
  wages2Rate?: number;
}

/** Which rate slot supplied a generic rate. */
export type RateKind =
  | 'crewLead'
  | 'hourlyJob'
  | 'painter'
  | 'commercialHourly'
  | 'wages2'
  | 'default';

/**
 * The rate lookup, resolved PER LINE by job type:
 *
 * - job is HOURLY → the person's `hourlyRate` (the global fallback when they
 *   have none)
 * - everyone else, crew leads included → their profit-share `rate`
 *
 * Round 10 — the separate crew-lead rate is GONE. A crew lead working a job
 * they lead is on a profit-share job by definition, so the PS rate already is
 * their crew-lead rate; keeping two numbers only let them drift apart.
 */
export function resolveWageRate(
  employeeName: string,
  opts: {
    wageRates: readonly WageRateLike[];
    defaultWageRate: number;
    hourlyJob: boolean;
    /** These are the crew lead's OWN timesheet hours (kept for callers/tags). */
    isCrewLead?: boolean;
    /**
     * Round 9 §18 — the timesheet pay item's rate slot. `commercialHourly` and
     * `wages2` are per-painter rates that outrank the hourly-job slot, because
     * the pay item is what the painter actually clocked into.
     */
    payRateKind?: PayRateKind;
  },
): { rate: number; defaulted: boolean; kind: RateKind } {
  const entry = opts.wageRates.find(
    (e) => tokensSubset(e.name, employeeName) || tokensSubset(employeeName, e.name),
  );
  const defaulted = !entry;

  // Round 9 §18 — a pay item with its own rate slot wins outright. With no
  // configured rate for that pay type the line falls back the same way an
  // ordinary line would (hourly job → hourly rate, else the PS rate), so
  // Round 9 §35 still pays Xavier Green $18 on an hourly job and $25 on a
  // work-order job even though both are clocked as Commercial Hourly.
  const genericFallback = (opts.hourlyJob ? entry?.hourlyRate : undefined) ?? entry?.rate ?? opts.defaultWageRate;
  if (opts.payRateKind === 'commercialHourly') {
    return {
      rate: entry?.commercialHourlyRate ?? genericFallback,
      defaulted,
      kind: entry?.commercialHourlyRate !== undefined ? 'commercialHourly' : 'painter',
    };
  }
  if (opts.payRateKind === 'wages2') {
    return {
      rate: entry?.wages2Rate ?? genericFallback,
      defaulted,
      kind: entry?.wages2Rate !== undefined ? 'wages2' : 'painter',
    };
  }
  if (opts.hourlyJob) {
    // No hourly-jobs rate on file falls through to the global default, NOT to
    // the person's profit-share rate — a piece-rate number would understate an
    // hourly job.
    const hourly = entry?.hourlyRate;
    return {
      rate: hourly ?? opts.defaultWageRate,
      defaulted,
      kind: hourly !== undefined ? 'hourlyJob' : 'default',
    };
  }
  // Round 10 — crew-lead hours on a profit-share job pay the PS rate, which is
  // the same row everyone else reads. `isCrewLead` only tags the line now.
  const isCrewLead = opts.payRateKind === 'crewLead' ? true : opts.isCrewLead === true;
  return {
    rate: entry?.rate ?? opts.defaultWageRate,
    defaulted,
    kind: isCrewLead ? 'crewLead' : entry ? 'painter' : 'default',
  };
}

/**
 * Round 8 §9.1 — person + AP# override. `apKeyPattern` is a substring match,
 * so '24097' also covers change-order suffixes; the longest matching pattern
 * wins when several apply.
 */
export function matchPerJobRateOverride(
  employeeName: string,
  jobApKey: string | null | undefined,
  overrides: readonly PerJobRateOverride[] | undefined,
): PerJobRateOverride | null {
  if (!jobApKey || !overrides || overrides.length === 0) return null;
  const key = jobApKey.trim().toUpperCase();
  let best: PerJobRateOverride | null = null;
  let bestLen = -1;
  for (const o of overrides) {
    const pattern = o.apKeyPattern.trim().toUpperCase();
    if (!pattern || !key.includes(pattern)) continue;
    if (!(tokensSubset(o.person, employeeName) || tokensSubset(employeeName, o.person))) continue;
    if (pattern.length > bestLen) {
      best = o;
      bestLen = pattern.length;
    }
  }
  return best;
}

/** Round 8 §9.1 — "person X earns Y while working a job led by Z". */
export function matchConditionalRateRule(
  employeeName: string,
  jobLeadNames: readonly string[] | undefined,
  rules: readonly ConditionalRateRule[] | undefined,
): ConditionalRateRule | null {
  if (!rules || rules.length === 0 || !jobLeadNames || jobLeadNames.length === 0) return null;
  for (const rule of rules) {
    if (rule.condition.kind !== 'onSameJobAsLead') continue;
    if (!(tokensSubset(rule.person, employeeName) || tokensSubset(employeeName, rule.person))) continue;
    const lead = rule.condition.leadName;
    if (jobLeadNames.some((n) => tokensSubset(lead, n) || tokensSubset(n, lead))) return rule;
  }
  return null;
}

/**
 * Round 8 §9.1 (extends Round 6 §7.2 + §7.1 Branch 1). Manual line edits are
 * applied by the caller before/after this. Precedence, most specific first:
 *
 *   1. person + AP# override (`perJobRateOverrides`)
 *   2. job-scoped rule (`jobRateRules`, Round 6 §7.2)
 *   3. conditional rule ("X on lead Z's job")
 *   4. the three generic rates (hourly-jobs / crew-lead / profit-share)
 *
 * then the Davis-Bacon max-wrap on DB jobs: `rate = max(resolved, dbRate)` —
 * it never LOWERS a rate, so a $25 painter stays at $25 on an SNRHA job.
 */
export function resolveLineRate(
  employeeName: string,
  jobText: string,
  hourlyJob: boolean,
  settings: {
    wageRates: WageRateLike[];
    defaultWageRate: number;
    prevailingWagePatterns: string[];
    prevailingMinRate: number;
    jobRateRules?: JobRateRule[];
    apKey?: string | null;
    /** Pivot-by-Job §7.1 toggle: skip Davis-Bacon floor even when patterns match. */
    forceNonDb?: boolean;
    conditionalRateRules?: readonly ConditionalRateRule[];
    perJobRateOverrides?: readonly PerJobRateOverride[];
    /** True when these are the crew lead's own timesheet hours. */
    isCrewLead?: boolean;
    /** Crew-lead names resolved on this job (drives the conditional rules). */
    jobLeadNames?: readonly string[];
    /** Round 9 §18 — rate slot selected by the timesheet pay item. */
    payRateKind?: PayRateKind;
  },
): LineRate {
  const apKey = settings.apKey ?? extractApKey(jobText);
  const perJob = matchPerJobRateOverride(employeeName, apKey, settings.perJobRateOverrides);
  const jobRule = perJob ? null : matchJobRateRule(employeeName, apKey, settings.jobRateRules);
  const conditional =
    perJob || jobRule
      ? null
      : matchConditionalRateRule(employeeName, settings.jobLeadNames, settings.conditionalRateRules);

  let rate: number;
  let defaulted = false;
  let kind: RateKind | null = null;
  let ruleNote: string | undefined;

  if (perJob) {
    rate = perJob.rate;
    if (perJob.note?.trim()) ruleNote = perJob.note.trim();
  } else if (jobRule) {
    rate = jobRule.rule.rate;
    if (jobRule.rule.note.trim()) ruleNote = jobRule.rule.note.trim();
  } else if (conditional) {
    rate = conditional.rate;
    if (conditional.note?.trim()) ruleNote = conditional.note.trim();
  } else {
    const info = resolveWageRate(employeeName, {
      wageRates: settings.wageRates,
      defaultWageRate: settings.defaultWageRate,
      hourlyJob,
      isCrewLead: settings.isCrewLead,
      payRateKind: settings.payRateKind,
    });
    rate = info.rate;
    defaulted = info.defaulted;
    kind = info.kind;
  }

  const isDb =
    !settings.forceNonDb && isPrevailingWageJob(jobText, settings.prevailingWagePatterns);
  if (isDb) {
    rate = Math.max(rate, settings.prevailingMinRate);
  }

  return {
    rate,
    defaulted,
    prevailing: isDb,
    ...(jobRule ? { fromJobRule: true, ...(ruleNote ? { ruleNote } : {}) } : {}),
    ...(perJob ? { fromPerJobOverride: true, ...(ruleNote ? { ruleNote } : {}) } : {}),
    ...(conditional ? { fromConditionalRule: true, ...(ruleNote ? { ruleNote } : {}) } : {}),
    ...(kind === 'crewLead' ? { fromCrewLeadRate: true } : {}),
    ...(kind === 'commercialHourly' || kind === 'wages2' ? { fromPayItemRate: kind } : {}),
  };
}

/**
 * Round 9 §3 — regular / daily-OT / weekly-OT hour buckets for one hourly
 * block, for the NOTES log's hour split. Returns nothing when the week is
 * straight time only, so a plain block keeps its single totals row.
 */
export function hourSplitRows(
  block: { lines: readonly { hours: number; rate: number; ot: boolean; dailyOt?: boolean }[] },
): { label: string; hours: number; pay: number }[] {
  const buckets: {
    label: string;
    match: (l: { ot: boolean; dailyOt?: boolean }) => boolean;
  }[] = [
    { label: 'regular hrs', match: (l) => !l.ot && !l.dailyOt },
    { label: 'daily OT hrs (1.5x)', match: (l) => l.dailyOt === true },
    { label: 'overtime hrs (1.5x)', match: (l) => l.ot && !l.dailyOt },
  ];
  const out: { label: string; hours: number; pay: number }[] = [];
  for (const b of buckets) {
    const lines = block.lines.filter(b.match);
    const hours = lines.reduce((a, l) => a + l.hours, 0);
    if (hours === 0) continue;
    out.push({
      label: b.label,
      hours: round2(hours),
      pay: round2(lines.reduce((a, l) => a + l.hours * l.rate, 0)),
    });
  }
  return out.length > 1 ? out : [];
}

/**
 * Round 6 §7.1 Branch 2 — crew-lead Davis-Bacon floor on per-job pay after
 * the §6.2 chain and any split. Never lowers pay; subcontractors exempt.
 */
export function applyDavisBaconLeadFloor(args: {
  isDbJob: boolean;
  isSubcontractor: boolean;
  timesheetHoursOnJob: number;
  profitSharePay: number;
  dbRate: number;
}): { pay: number; floorApplied: boolean; floor: number; priorPay: number } {
  const priorPay = args.profitSharePay;
  if (!args.isDbJob || args.isSubcontractor) {
    return { pay: priorPay, floorApplied: false, floor: 0, priorPay };
  }
  const hours = args.timesheetHoursOnJob;
  if (!Number.isFinite(hours) || hours <= 0) {
    return { pay: priorPay, floorApplied: false, floor: 0, priorPay };
  }
  const floor = round2(hours * args.dbRate);
  if (priorPay < floor) {
    return { pay: floor, floorApplied: true, floor, priorPay };
  }
  return { pay: priorPay, floorApplied: false, floor, priorPay };
}
