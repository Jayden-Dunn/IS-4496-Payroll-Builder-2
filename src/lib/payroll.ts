/**
 * Round 3 shared payroll derivations (§3): job display names, crew-lead
 * extraction/resolution, overtime split, WO pre-fills. All pure functions —
 * nothing dataset-specific lives here; tribal knowledge comes from Settings.
 */
import { AP_RE } from './apkey';
import { round2 } from './format';
import type {
  CrewLeadResolution,
  DirectoryEmployee,
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
  return { name: matches[0].fullName.toUpperCase(), sub: false, flags };
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
 * If the employee's week exceeds 40 hours, move the overage to an [OT] line
 * on the LAST job worked (latest date; last by file order within the day),
 * reducing that job's regular hours so regular + OT = actual.
 */
export function splitOvertime(jobs: OtJobHours[]): OtSplit {
  const regularByJob = new Map<string, number>(jobs.map((j) => [j.jobId, round2(j.hours)]));
  const total = round2(jobs.reduce((acc, j) => acc + j.hours, 0));
  if (total <= 40 || jobs.length === 0) return { regularByJob, otJobId: null, otHours: 0 };
  const otHours = round2(total - 40);
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
