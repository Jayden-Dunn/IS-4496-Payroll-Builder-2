"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.jobDisplayName = jobDisplayName;
exports.nameTokens = nameTokens;
exports.nameTokenSets = nameTokenSets;
exports.tokensSubset = tokensSubset;
exports.extractLeadParenthetical = extractLeadParenthetical;
exports.splitLeadCandidates = splitLeadCandidates;
exports.resolveCrewLead = resolveCrewLead;
exports.directoryFromTimesheet = directoryFromTimesheet;
exports.splitOvertime = splitOvertime;
exports.extractWoOverride = extractWoOverride;
exports.extractWoPct = extractWoPct;
exports.profitSharePctFor = profitSharePctFor;
exports.wageRateFor = wageRateFor;
/**
 * Round 3 shared payroll derivations (§3): job display names, crew-lead
 * extraction/resolution, overtime split, WO pre-fills. All pure functions —
 * nothing dataset-specific lives here; tribal knowledge comes from Settings.
 */
const apkey_1 = require("./apkey");
const format_1 = require("./format");
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
    return { name: matches[0].fullName.toUpperCase(), sub: false, flags };
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
 * If the employee's week exceeds 40 hours, move the overage to an [OT] line
 * on the LAST job worked (latest date; last by file order within the day),
 * reducing that job's regular hours so regular + OT = actual.
 */
function splitOvertime(jobs) {
    const regularByJob = new Map(jobs.map((j) => [j.jobId, (0, format_1.round2)(j.hours)]));
    const total = (0, format_1.round2)(jobs.reduce((acc, j) => acc + j.hours, 0));
    if (total <= 40 || jobs.length === 0)
        return { regularByJob, otJobId: null, otHours: 0 };
    const otHours = (0, format_1.round2)(total - 40);
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
