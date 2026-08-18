"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.jobDisplayName = jobDisplayName;
exports.nameTokens = nameTokens;
exports.nameTokenSets = nameTokenSets;
exports.tokensSubset = tokensSubset;
exports.stripSingleLetterTokens = stripSingleLetterTokens;
exports.titleCaseLeadName = titleCaseLeadName;
exports.extractLeadParenthetical = extractLeadParenthetical;
exports.splitLeadCandidates = splitLeadCandidates;
exports.resolveCrewLead = resolveCrewLead;
exports.directoryFromTimesheet = directoryFromTimesheet;
exports.splitOvertime = splitOvertime;
exports.splitDailyOvertime = splitDailyOvertime;
exports.extractWoOverride = extractWoOverride;
exports.extractWoPct = extractWoPct;
exports.profitSharePctFor = profitSharePctFor;
exports.wageRateFor = wageRateFor;
exports.computeTotalProfitShare = computeTotalProfitShare;
exports.computeTotalPay = computeTotalPay;
exports.isPrevailingWageJob = isPrevailingWageJob;
exports.matchJobRateRule = matchJobRateRule;
exports.resolveWageRate = resolveWageRate;
exports.matchPerJobRateOverride = matchPerJobRateOverride;
exports.matchConditionalRateRule = matchConditionalRateRule;
exports.resolveLineRate = resolveLineRate;
exports.applyDavisBaconLeadFloor = applyDavisBaconLeadFloor;
/**
 * Round 3 shared payroll derivations (§3): job display names, crew-lead
 * extraction/resolution, overtime split, WO pre-fills. All pure functions —
 * nothing dataset-specific lives here; tribal knowledge comes from Settings.
 */
const apkey_1 = require("./apkey");
const format_1 = require("./format");
const jobNames_1 = require("./jobNames");
// ---------------------------------------------------------------------------
// §3.1 Job display name
// ---------------------------------------------------------------------------
/** "Aaron Ward" + "AP24314 Ext" → "Aaron Ward Ext" (single-spaced). */
function jobDisplayName(jobcode1, jobcode2) {
    const jobText = jobcode2.replace(apkey_1.AP_RE, ' ').replace(/\s+/g, ' ').trim();
    return `${jobcode1.trim()} ${jobText}`.replace(/\s+/g, ' ').trim();
}
// ---------------------------------------------------------------------------
// Name tokens (shared by lead matching and directory lookup)
// ---------------------------------------------------------------------------
/** Lower-cased name tokens; hyphens split; apostrophes kept inside tokens. */
function nameTokens(s) {
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
function nameTokenSets(fullName) {
    const tokens = nameTokens(fullName);
    const all = new Set(tokens);
    const required = new Set(tokens.filter((t) => t.length > 1));
    return { all, required };
}
/** Do all candidate tokens appear in the employee's token set? */
function tokensSubset(candidate, employeeFullName) {
    const cand = nameTokens(candidate);
    if (cand.length === 0)
        return false;
    const { all } = nameTokenSets(employeeFullName);
    return cand.every((t) => all.has(t));
}
/**
 * Round 4 §3.2-2 — display names drop single-letter tokens (optionally
 * followed by '.') anywhere in the name: "Rene Salazar M" → "Rene Salazar",
 * "Adriana G Lopez Barajas" → "Adriana Lopez Barajas". Matching logic keeps
 * the raw tokens; only display changes.
 */
function stripSingleLetterTokens(name) {
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
function titleCaseLeadName(leadDisplay, directory, subcontractors) {
    return leadDisplay
        .split('/')
        .map((part) => {
        const p = part.trim();
        if (!p)
            return p;
        const dir = directory.find((d) => d.fullName.toUpperCase() === p.toUpperCase());
        if (dir)
            return stripSingleLetterTokens(dir.fullName);
        const sub = subcontractors.find((s) => s.toUpperCase() === p.toUpperCase());
        if (sub)
            return sub;
        return stripSingleLetterTokens(p);
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
function extractLeadParenthetical(title, knownEstimators) {
    const groups = [...title.matchAll(/\(([^()]*)\)/g)].map((m) => m[1].trim()).filter(Boolean);
    if (groups.length === 0)
        return null;
    const last = groups[groups.length - 1];
    if (/^[A-Za-z]{2,3}$/.test(last) && knownEstimators.has(last.toUpperCase()))
        return null;
    return last;
}
/** Strip decorations ("hrly_", "*") and split into lead-name candidates. */
function splitLeadCandidates(raw) {
    return raw
        .split(/[/,&]/)
        .map((s) => s.replace(/\b[a-z0-9]+_/gi, '').replace(/\*/g, '').replace(/\s+/g, ' ').trim())
        .filter(Boolean);
}
/** Strip the "'s Crew" suffix from a group label; null when absent. */
function crewLabelName(label) {
    const m = /^(.*?)['\u2019]s\s+crew\s*$/i.exec(label.trim());
    return m ? m[1].trim() : null;
}
/** Employee name token-matches some crew-group label (Round-2 lead rule). */
function isKnownCrewLead(employeeFullName, crewGroups) {
    const { all } = nameTokenSets(employeeFullName);
    for (const g of crewGroups) {
        const name = crewLabelName(g);
        if (!name)
            continue;
        const labelTokens = nameTokens(name);
        if (labelTokens.length > 0 && labelTokens.every((t) => all.has(t)))
            return true;
    }
    return false;
}
/**
 * Resolve one candidate (a first name, full name, or business name) to a
 * directory employee or subcontractor. Never invents a name.
 */
function resolveCandidate(candidate, ctx) {
    const flags = [];
    const candLower = candidate.toLowerCase();
    const candTokens = nameTokens(candidate);
    let matches = ctx.directory.filter((e) => {
        if (candTokens.length > 1)
            return tokensSubset(candidate, e.fullName);
        const firstTok = nameTokens(e.firstName)[0] ?? '';
        return firstTok === candLower;
    });
    if (matches.length === 0) {
        // §3.2.4 — Subcontractors list by token match.
        const subMatch = ctx.subcontractors.find((s) => tokensSubset(candidate, s) || s.toLowerCase() === candLower);
        if (subMatch)
            return { name: subMatch.toUpperCase(), sub: true, flags };
        flags.push({ kind: 'noMatch', candidate });
        return { name: null, sub: false, flags };
    }
    // Prefer active matches; the only match being inactive is used but flagged.
    const actives = matches.filter((e) => e.active);
    if (actives.length > 0) {
        matches = actives;
    }
    else {
        flags.push({ kind: 'inactiveMatch', candidate });
    }
    // Among the pool, prefer known crew leads.
    const leads = matches.filter((e) => isKnownCrewLead(e.fullName, ctx.crewGroups));
    if (leads.length > 0) {
        matches = leads;
    }
    else {
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
function resolveCrewLead(raw, ctx) {
    if (!raw)
        return { display: '', names: [], split: 1, sub: false, raw: '', flags: [] };
    const candidates = splitLeadCandidates(raw);
    const names = [];
    const flags = [];
    let sub = false;
    for (const cand of candidates) {
        const res = resolveCandidate(cand, ctx);
        flags.push(...res.flags);
        if (res.name) {
            names.push(res.name);
            if (res.sub)
                sub = true;
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
function directoryFromTimesheet(rows) {
    const seen = new Map();
    for (const r of rows) {
        const fullName = `${r.fname} ${r.lname}`.replace(/\s+/g, ' ').trim();
        if (!fullName || seen.has(fullName.toLowerCase()))
            continue;
        seen.set(fullName.toLowerCase(), {
            fullName,
            firstName: r.fname.trim(),
            lastName: r.lname.trim(),
            active: true,
        });
    }
    return [...seen.values()];
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
function splitOvertime(jobs, opts) {
    const regularByJob = new Map(jobs.map((j) => [j.jobId, (0, format_1.round2)(j.hours)]));
    const total = (0, format_1.round2)(jobs.reduce((acc, j) => acc + j.hours, 0));
    const dailyOtHours = (0, format_1.round2)(Math.max(0, opts?.dailyOtHours ?? 0));
    if (total <= 40 || jobs.length === 0)
        return { regularByJob, otJobId: null, otHours: 0 };
    const otHours = (0, format_1.round2)(Math.max(0, total - 40 - dailyOtHours));
    if (otHours <= 0)
        return { regularByJob, otJobId: null, otHours: 0 };
    const last = jobs.reduce((best, cur) => {
        const bt = best.lastDate ? best.lastDate.getTime() : -Infinity;
        const ct = cur.lastDate ? cur.lastDate.getTime() : -Infinity;
        if (ct !== bt)
            return ct > bt ? cur : best;
        return cur.lastRowIndex >= best.lastRowIndex ? cur : best;
    });
    regularByJob.set(last.jobId, (0, format_1.round2)((regularByJob.get(last.jobId) ?? 0) - otHours));
    return { regularByJob, otJobId: last.jobId, otHours };
}
function dayKeyOf(d) {
    if (!d)
        return '';
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
function splitDailyOvertime(rows, thresholdHours) {
    const otByJob = new Map();
    const segments = [];
    let totalOtHours = 0;
    if (!(thresholdHours > 0))
        return { otByJob, segments, totalOtHours };
    const byDay = new Map();
    for (const r of rows) {
        const key = dayKeyOf(r.date);
        if (!key || r.hours <= 0)
            continue;
        const list = byDay.get(key);
        if (list)
            list.push(r);
        else
            byDay.set(key, [r]);
    }
    for (const key of [...byDay.keys()].sort()) {
        const list = byDay.get(key) ?? [];
        const dayHours = (0, format_1.round2)(list.reduce((a, r) => a + r.hours, 0));
        if (dayHours <= thresholdHours)
            continue;
        let remaining = (0, format_1.round2)(dayHours - thresholdHours);
        for (const r of [...list].sort((a, b) => b.rowIndex - a.rowIndex)) {
            if (remaining <= 0)
                break;
            const take = (0, format_1.round2)(Math.min(remaining, r.hours));
            remaining = (0, format_1.round2)(remaining - take);
            if (!r.eligible || take <= 0)
                continue;
            otByJob.set(r.jobId, (0, format_1.round2)((otByJob.get(r.jobId) ?? 0) + take));
            segments.push({ jobId: r.jobId, dateIso: key, date: r.date, dayHours, otHours: take });
            totalOtHours = (0, format_1.round2)(totalOtHours + take);
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
function extractWoOverride(description) {
    for (const raw of description.split(/\r?\n/)) {
        const line = raw.trim();
        if (!line)
            continue;
        let m = WO_BEFORE_RE.exec(line);
        let amtText = m ? m[2] : null;
        if (!amtText) {
            m = WO_AFTER_RE.exec(line);
            amtText = m ? m[1] : null;
        }
        if (amtText) {
            const n = Number(amtText.replace(/,/g, ''));
            if (Number.isFinite(n) && n > 0)
                return { amount: n, line };
        }
    }
    return null;
}
const WO_PCT_RE = /@\s*(\d{1,3}(?:\.\d+)?)\s*%\s*PS\b/i;
/** "@ 40% PS" status-line override for the crew profit-share % (§4.3). */
function extractWoPct(description) {
    for (const raw of description.split(/\r?\n/)) {
        const line = raw.trim();
        const m = WO_PCT_RE.exec(line);
        if (m) {
            const n = Number(m[1]);
            if (n >= 0 && n <= 100)
                return { pct: n / 100, line };
        }
    }
    return null;
}
// ---------------------------------------------------------------------------
// Profit-share % lookup (Settings §7.1)
// ---------------------------------------------------------------------------
/** Look up a lead's WO% from the Settings map by token match. */
function profitSharePctFor(leadName, map, fallback) {
    const entry = map.find((e) => tokensSubset(e.lead, leadName) || tokensSubset(leadName, e.lead));
    return entry ? entry.pct : fallback;
}
/** Wage rate lookup (Settings §7.2) by token match. */
function wageRateFor(employeeName, rates, fallback, hourlyJob) {
    const entry = rates.find((e) => tokensSubset(e.name, employeeName) || tokensSubset(employeeName, e.name));
    if (!entry)
        return { rate: fallback, defaulted: true };
    return { rate: hourlyJob && entry.hourlyRate !== undefined ? entry.hourlyRate : entry.rate, defaulted: false };
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
function computeTotalProfitShare(args) {
    void args.mileage;
    const base = args.workOrderOverride !== null && args.workOrderOverride !== undefined
        ? args.workOrderOverride
        : args.profitShare;
    const percent = args.percent ?? null;
    if ((0, format_1.isPercentState)(percent) && percent !== 'HOURLY') {
        return { total: null, base, percentCompletion: 0, manualPay: percent, missingPercent: false };
    }
    const pct = (0, format_1.isPercentNumber)(percent) ? percent : 0;
    return {
        total: base * pct - args.laborWages - args.halfMaterials,
        base,
        percentCompletion: pct,
        manualPay: null,
        missingPercent: !(0, format_1.isPercentNumber)(percent),
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
function computeTotalPay(args) {
    return args.total + args.hourlyPay + args.reimbursements - args.deductions;
}
// ---------------------------------------------------------------------------
// Round 6 §7.1 / §7.2 — Davis-Bacon + job-rate resolution
// ---------------------------------------------------------------------------
/** Case-insensitive substring match of any configured pattern on the job text. */
function isPrevailingWageJob(jobText, patterns) {
    const t = jobText.toLowerCase();
    return patterns.some((p) => p.trim() !== '' && t.includes(p.trim().toLowerCase()));
}
/** How a job-rate rule's AP# matches a timesheet/calendar apKey. */
function jobRuleApMatch(rule, jobApKey) {
    const ruleKey = ((0, apkey_1.normalizeApInput)(rule.apKey) ?? rule.apKey.trim()).toUpperCase();
    const jobKey = jobApKey.trim().toUpperCase();
    if (!ruleKey || !jobKey)
        return null;
    if (ruleKey === jobKey)
        return 'exact';
    if (!rule.includeSuffixes)
        return null;
    const ruleBase = (0, jobNames_1.baseApKey)(ruleKey) ?? ruleKey;
    const jobBase = (0, jobNames_1.baseApKey)(jobKey) ?? jobKey;
    return ruleBase === jobBase ? 'suffix' : null;
}
function jobRuleAppliesToEmployee(rule, employeeName) {
    if (rule.employeeNames.length === 0)
        return true;
    return rule.employeeNames.some((n) => tokensSubset(n, employeeName) || tokensSubset(employeeName, n));
}
/**
 * Specificity: exact+named (4) > exact+all (3) > suffix+named (2) > suffix+all (1).
 * Returns null when no rule matches.
 */
function matchJobRateRule(employeeName, jobApKey, rules) {
    if (!jobApKey || !rules || rules.length === 0)
        return null;
    let best = null;
    for (const rule of rules) {
        const ap = jobRuleApMatch(rule, jobApKey);
        if (!ap)
            continue;
        if (!jobRuleAppliesToEmployee(rule, employeeName))
            continue;
        const named = rule.employeeNames.length > 0;
        const specificity = ap === 'exact' ? (named ? 4 : 3) : named ? 2 : 1;
        if (!best || specificity > best.specificity)
            best = { rule, specificity };
    }
    return best;
}
/**
 * Round 8 §9.1 — the three-rate lookup, resolved PER LINE by job type:
 *
 * - job is HOURLY → the person's `hourlyRate` (their own rate, then the global
 *   fallback, when they have none)
 * - crew lead's own hours on a profit-share job → their `crewLeadRate`, else
 *   the global `defaultCrewLeadRate`
 * - everyone else → their profit-share `rate`
 */
function resolveWageRate(employeeName, opts) {
    const entry = opts.wageRates.find((e) => tokensSubset(e.name, employeeName) || tokensSubset(employeeName, e.name));
    const defaulted = !entry;
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
    if (opts.isCrewLead) {
        return {
            rate: entry?.crewLeadRate ?? opts.defaultCrewLeadRate ?? entry?.rate ?? opts.defaultWageRate,
            defaulted,
            kind: 'crewLead',
        };
    }
    return {
        rate: entry?.rate ?? opts.defaultWageRate,
        defaulted,
        kind: entry ? 'painter' : 'default',
    };
}
/**
 * Round 8 §9.1 — person + AP# override. `apKeyPattern` is a substring match,
 * so '24097' also covers change-order suffixes; the longest matching pattern
 * wins when several apply.
 */
function matchPerJobRateOverride(employeeName, jobApKey, overrides) {
    if (!jobApKey || !overrides || overrides.length === 0)
        return null;
    const key = jobApKey.trim().toUpperCase();
    let best = null;
    let bestLen = -1;
    for (const o of overrides) {
        const pattern = o.apKeyPattern.trim().toUpperCase();
        if (!pattern || !key.includes(pattern))
            continue;
        if (!(tokensSubset(o.person, employeeName) || tokensSubset(employeeName, o.person)))
            continue;
        if (pattern.length > bestLen) {
            best = o;
            bestLen = pattern.length;
        }
    }
    return best;
}
/** Round 8 §9.1 — "person X earns Y while working a job led by Z". */
function matchConditionalRateRule(employeeName, jobLeadNames, rules) {
    if (!rules || rules.length === 0 || !jobLeadNames || jobLeadNames.length === 0)
        return null;
    for (const rule of rules) {
        if (rule.condition.kind !== 'onSameJobAsLead')
            continue;
        if (!(tokensSubset(rule.person, employeeName) || tokensSubset(employeeName, rule.person)))
            continue;
        const lead = rule.condition.leadName;
        if (jobLeadNames.some((n) => tokensSubset(lead, n) || tokensSubset(n, lead)))
            return rule;
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
function resolveLineRate(employeeName, jobText, hourlyJob, settings) {
    const apKey = settings.apKey ?? (0, apkey_1.extractApKey)(jobText);
    const perJob = matchPerJobRateOverride(employeeName, apKey, settings.perJobRateOverrides);
    const jobRule = perJob ? null : matchJobRateRule(employeeName, apKey, settings.jobRateRules);
    const conditional = perJob || jobRule
        ? null
        : matchConditionalRateRule(employeeName, settings.jobLeadNames, settings.conditionalRateRules);
    let rate;
    let defaulted = false;
    let kind = null;
    let ruleNote;
    if (perJob) {
        rate = perJob.rate;
        if (perJob.note?.trim())
            ruleNote = perJob.note.trim();
    }
    else if (jobRule) {
        rate = jobRule.rule.rate;
        if (jobRule.rule.note.trim())
            ruleNote = jobRule.rule.note.trim();
    }
    else if (conditional) {
        rate = conditional.rate;
        if (conditional.note?.trim())
            ruleNote = conditional.note.trim();
    }
    else {
        const info = resolveWageRate(employeeName, {
            wageRates: settings.wageRates,
            defaultWageRate: settings.defaultWageRate,
            defaultCrewLeadRate: settings.defaultCrewLeadRate,
            hourlyJob,
            isCrewLead: settings.isCrewLead,
        });
        rate = info.rate;
        defaulted = info.defaulted;
        kind = info.kind;
    }
    const isDb = !settings.forceNonDb && isPrevailingWageJob(jobText, settings.prevailingWagePatterns);
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
    };
}
/**
 * Round 6 §7.1 Branch 2 — crew-lead Davis-Bacon floor on per-job pay after
 * the §6.2 chain and any split. Never lowers pay; subcontractors exempt.
 */
function applyDavisBaconLeadFloor(args) {
    const priorPay = args.profitSharePay;
    if (!args.isDbJob || args.isSubcontractor) {
        return { pay: priorPay, floorApplied: false, floor: 0, priorPay };
    }
    const hours = args.timesheetHoursOnJob;
    if (!Number.isFinite(hours) || hours <= 0) {
        return { pay: priorPay, floorApplied: false, floor: 0, priorPay };
    }
    const floor = (0, format_1.round2)(hours * args.dbRate);
    if (priorPay < floor) {
        return { pay: floor, floorApplied: true, floor, priorPay };
    }
    return { pay: priorPay, floorApplied: false, floor, priorPay };
}
