"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.RENO_ESTIMATOR = exports.EMPTY_MANUAL_STATE = void 0;
exports.formatTimeNaviRow = formatTimeNaviRow;
exports.parseTimeNaviRowRef = parseTimeNaviRowRef;
exports.carriedSectionFor = carriedSectionFor;
exports.pbjJobId = pbjJobId;
exports.isHourlyJob = isHourlyJob;
exports.buildPivotByJob = buildPivotByJob;
exports.applyManualState = applyManualState;
exports.weekdayShort = weekdayShort;
exports.fmtMonthDay = fmtMonthDay;
/**
 * Round 3 §4 — "Pivot by Job": customer → job → employee hour tree with the
 * analysis columns, plus the BALANCES and UT PAYROLL manual sections that
 * auto-position below the pivot (obsoleting the Excel move-refresh-move
 * workaround).
 */
const apkey_1 = require("./apkey");
const customerByAp_1 = require("./customerByAp");
const format_1 = require("./format");
const jobNames_1 = require("./jobNames");
const payroll_1 = require("./payroll");
/** Round 6 §5.1 — display form for a matched TimeNavi Export sheet row. */
function formatTimeNaviRow(sheetRow) {
    return `TimeNavi row ${sheetRow}`;
}
/** Parse `TimeNavi row 19` or a bare `19` into the 1-based sheet row number. */
function parseTimeNaviRowRef(text) {
    const t = text.trim();
    const labeled = /^TimeNavi\s+row\s+(\d+)$/i.exec(t);
    if (labeled)
        return Number(labeled[1]);
    if (/^\d+$/.test(t))
        return Number(t);
    return null;
}
exports.EMPTY_MANUAL_STATE = { added: [], deleted: [], edits: {} };
/**
 * Round 8 §4.2 — the Reno, NV estimator's initial. Reno is a Nevada branch, so
 * his carried jobs belong to NV PAYROLL (BALANCES) no matter which calendar
 * the event came from.
 */
exports.RENO_ESTIMATOR = 'RF';
/** Section a carried job lands in: the estimator outranks the source calendar. */
function carriedSectionFor(defaultSection, estimator) {
    return (estimator ?? '').trim().toUpperCase() === exports.RENO_ESTIMATOR ? 'balances' : defaultSection;
}
function pbjJobId(customer, jobcode2) {
    return `${customer}\u0000${jobcode2}`;
}
/** A job is hourly when its % says HOURLY, its lead is HOURLY, or it's tagged T&M/[hrly]. */
function isHourlyJob(percent, crewLead, jobText) {
    if (percent === 'HOURLY')
        return true;
    if (/^hourly$/i.test(crewLead.trim()))
        return true;
    return /\bT\s*&\s*M\b|\[hrly\]/i.test(jobText);
}
const byText = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' });
function buildPivotByJob(input) {
    const { enriched, reportRows, settings, leadCtx, knownEstimators, payDate, edits, qbHours } = input;
    const colorByEstimator = new Map(settings.legend.map((l) => [l.estimator, l.color]));
    // TimeNavi lookup: first included row per apKey + its 1-based sheet row
    // number (header = row 1) for the "On Calendar?" MATCH-style reference.
    const included = reportRows.filter((r) => r.included);
    const eventByApKey = new Map();
    included.forEach((r, i) => {
        if (r.apKey && !eventByApKey.has(r.apKey))
            eventByApKey.set(r.apKey, { row: r, sheetRow: i + 2 });
    });
    const nameSources = {
        timesheetByAp: (0, jobNames_1.timesheetDisplayMap)(enriched),
        carryByAp: (0, jobNames_1.carryDisplayMap)(input.lastWeekPbj),
        salesByAp: (0, jobNames_1.salesDisplayMap)(input.nvSales, input.utSales),
        pl: input.pl,
    };
    // ----- dates -----------------------------------------------------------
    const dateMs = new Set();
    let hasUndated = false;
    for (const r of enriched) {
        // NV pivot date columns follow NV hours; UT section headers reuse the same axis.
        if (r.localDate)
            dateMs.add(r.localDate.getTime());
        else if (r.hours !== 0)
            hasUndated = true;
    }
    const dates = [...dateMs].sort((a, b) => a - b).map((ms) => new Date(ms));
    const dateIsos = dates.map(format_1.dateToIso);
    if (hasUndated)
        dateIsos.push(''); // "(blank)" column only when needed
    const custs = new Map();
    enriched.forEach((r, idx) => {
        if (r.hours === 0)
            return;
        // Round 6 — UT timesheet hours belong under UT PAYROLL, not the NV pivot tree.
        if (r.sourceEntity === 'UT')
            return;
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
        const iso = r.localDate ? (0, format_1.dateToIso)(r.localDate) : '';
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
    const toRecord = (m) => {
        const rec = {};
        for (const iso of dateIsos) {
            const v = m.get(iso);
            if (v !== undefined && (0, format_1.round2)(v) !== 0)
                rec[iso] = (0, format_1.round2)(v);
        }
        return rec;
    };
    const grandByDateMap = new Map();
    let grandTotal = 0;
    const customers = [...custs.entries()]
        .sort(([a], [b]) => byText(a, b))
        .map(([customer, jobsMap]) => {
        const custByDate = new Map();
        let custTotal = 0;
        const jobs = [...jobsMap.values()]
            .sort((a, b) => byText(a.jobcode2, b.jobcode2))
            .map((job) => {
            const id = pbjJobId(customer, job.jobcode2);
            const edit = edits[id] ?? {};
            const defaultAp = (0, apkey_1.timesheetApKey)(job.jobcode2);
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
            const livePercent = isOffice ? null : match ? match.row.percent : null;
            const percentOverride = !isOffice && edit.percent !== undefined;
            const percent = isOffice
                ? null
                : percentOverride
                    ? edit.percent
                    : match
                        ? livePercent
                        : (edit.percentSeed ?? null);
            // Crew lead — from the matched event title (§3.2); OFFICE → HOURLY.
            const resolution = match
                ? (0, payroll_1.resolveCrewLead)((0, payroll_1.extractLeadParenthetical)(match.row.event, knownEstimators), leadCtx)
                : null;
            const crewLead = edit.crewLead !== undefined ? edit.crewLead : isOffice ? 'HOURLY' : (resolution?.display ?? '');
            const jobText = `${job.jobcode2} ${match?.row.event ?? ''}`;
            const hourly = isHourlyJob(percent, crewLead, jobText);
            // Round 6 §5.3 — WO% / WO Override blank by default; calendar patterns
            // are suggestions only (tooltips / click-to-accept). Slot A seeds via edits.
            const pctSrc = match ? (0, payroll_1.extractWoPct)(match.row.description) : null;
            const woSrc = match ? (0, payroll_1.extractWoOverride)(match.row.description) : null;
            const woPct = hourly ? null : edit.woPct !== undefined ? edit.woPct : null;
            const woOverride = hourly ? null : edit.woOverride !== undefined ? edit.woOverride : null;
            // Round 7 — notes stay blank unless the user edited; do not carry last-week / *PAID* text.
            const notes = edit.notes !== undefined ? edit.notes : '';
            // Round 6 §7.1 — DB chip + optional force-non-DB override.
            const dbPatternMatch = (0, payroll_1.isPrevailingWageJob)(jobText, settings.prevailingWagePatterns);
            const dbOverride = edit.dbOverride === true;
            const isDbJob = dbPatternMatch && !dbOverride;
            const analysis = {
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
            const employees = [...job.employees.values()]
                .sort((a, b) => a.firstIndex - b.firstIndex)
                .map((e) => ({
                username: e.username,
                byDate: toRecord(e.byDate),
                total: (0, format_1.round2)(e.total),
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
                total: (0, format_1.round2)(job.total),
                analysis,
                employees,
                colorHex,
            };
        });
        return {
            customer,
            byDate: toRecord(custByDate),
            total: (0, format_1.round2)(custTotal),
            jobs,
            colorHex: jobs[0]?.colorHex ?? null,
        };
    });
    // ----- §4.6 manual sections ---------------------------------------------
    const timesheetApKeys = new Set();
    for (const c of customers)
        for (const j of c.jobs)
            if (j.analysis.apKey)
                timesheetApKeys.add(j.analysis.apKey);
    const balancesSeeds = [];
    const utSeeds = [];
    const seenBalancesAp = new Set();
    const seenUtAp = new Set();
    // Round 8 §4.3 — every carried job resolves its OWN customer by AP key
    // (Sales by Customer Summary first, P&L header structure second).
    const customerResolver = (0, customerByAp_1.buildCustomerByApResolver)({
        nvSalesByCust: input.nvSalesByCust,
        utSalesByCust: input.utSalesByCust,
        pl: input.pl,
    });
    const unresolvedCarriedCustomers = [];
    const resolveCarriedCustomer = (apKey, section, fallback, { flag } = { flag: true }) => {
        const hit = customerResolver.resolve(apKey, section === 'ut' ? 'UT' : 'NV');
        if (hit.customer)
            return hit.customer;
        if (flag && apKey && !unresolvedCarriedCustomers.includes(apKey)) {
            unresolvedCarriedCustomers.push(apKey);
        }
        return fallback;
    };
    /** §3.1/§3.3 — this week's TimeNavi `% Completion` for an AP#, if any. */
    const liveLink = (apKey) => {
        const hit = apKey ? eventByApKey.get(apKey) : undefined;
        return hit ? { linked: true, percent: hit.row.percent } : { linked: false, percent: null };
    };
    const pushSeed = (section, seed) => {
        if (section === 'balances') {
            seenBalancesAp.add(seed.apKey ?? seed.id);
            balancesSeeds.push(seed);
        }
        else {
            seenUtAp.add(seed.apKey ?? seed.id);
            utSeeds.push(seed);
        }
    };
    for (const row of input.lastWeekPbj ?? []) {
        // §4.2 — RF (Reno, NV) carried jobs go to NV BALANCES, never UT PAYROLL.
        const section = carriedSectionFor(row.section === 'ut' ? 'ut' : 'balances', row.estimator);
        if (seenBalancesAp.has(row.apKey) || seenUtAp.has(row.apKey))
            continue;
        if (row.section !== 'ut') {
            // NV pivot/balances rows: carry when paid < 100% and no hours this week.
            const partial = typeof row.percent === 'number' && row.percent < 1;
            const open = row.section === 'balances' ? row.percent !== 1 : partial;
            if (!open || timesheetApKeys.has(row.apKey))
                continue;
        }
        // Round 7 — no carried notes / *PAID* text in BALANCES/UT.
        pushSeed(section, manualRowFromCarry({
            section,
            row,
            notes: '',
            colorByEstimator,
            live: liveLink(row.apKey),
            customer: resolveCarriedCustomer(row.apKey, section, row.customer),
        }));
    }
    // Seeds from this week's included Utah/Reno-calendar events (dedupe by apKey).
    for (const r of included) {
        if (!/^utah\b|^reno\b/i.test(r.calendar.trim()))
            continue;
        const apKey = r.apKey;
        if (!apKey || seenUtAp.has(apKey) || seenBalancesAp.has(apKey))
            continue;
        const section = carriedSectionFor('ut', r.estimator);
        const resolution = (0, payroll_1.resolveCrewLead)((0, payroll_1.extractLeadParenthetical)(r.event, knownEstimators), leadCtx);
        // Round 6 §5.3 — blank WO defaults (suggestions live on calendar-matched NV jobs only).
        const names = (0, jobNames_1.jobDisplayPair)(apKey, { ...nameSources, eventTitle: r.event });
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
        if (r.sourceEntity !== 'UT' || r.hours === 0)
            continue;
        const apKey = r.apKey;
        if (!apKey || apKey === 'OFFICE' || seenUtAp.has(apKey) || seenBalancesAp.has(apKey))
            continue;
        seenUtAp.add(apKey);
        const match = eventByApKey.get(apKey);
        const resolution = match
            ? (0, payroll_1.resolveCrewLead)((0, payroll_1.extractLeadParenthetical)(match.row.event, knownEstimators), leadCtx)
            : null;
        const names = (0, jobNames_1.jobDisplayPair)(apKey, {
            ...nameSources,
            eventTitle: match?.row.event ?? null,
        });
        const live = liveLink(apKey);
        utSeeds.push({
            id: `ut:${apKey}`,
            customer: resolveCarriedCustomer(apKey, 'ut', names.customerName, { flag: false }) ||
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
    const grandByDate = {};
    for (const iso of dateIsos) {
        const v = grandByDateMap.get(iso);
        if (v !== undefined)
            grandByDate[iso] = (0, format_1.round2)(v);
    }
    return {
        title: `NV PAYROLL: ${payDate ? (0, format_1.fmtDate)(payDate) : ''}`.trim(),
        dates,
        dateIsos,
        customers,
        grandByDate,
        grandTotal: (0, format_1.round2)(grandTotal),
        balances,
        utPayroll,
        unresolvedCarriedCustomers,
    };
}
function manualRowFromCarry(args) {
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
function applyRowEdits(seed, patch) {
    const out = { ...seed };
    if (!patch)
        return out;
    for (const key of Object.keys(patch)) {
        const value = patch[key];
        if (value === undefined)
            continue;
        Object.assign(out, { [key]: value });
    }
    // §3.1 — a typed % on a carried row wins over the live TimeNavi value.
    if (patch.percent !== undefined)
        out.percentOverride = true;
    return out;
}
/** Merge seeds with user adds/deletes/edits (rows persist keyed by id). */
function applyManualState(seeds, state) {
    const deleted = new Set(state.deleted);
    const out = [];
    const seen = new Set();
    for (const seed of seeds) {
        if (deleted.has(seed.id))
            continue;
        seen.add(seed.id);
        out.push(applyRowEdits(seed, state.edits[seed.id]));
    }
    for (const added of state.added) {
        if (deleted.has(added.id) || seen.has(added.id))
            continue;
        out.push(applyRowEdits(added, state.edits[added.id]));
    }
    return out;
}
/** Weekday short name for the §4.1 day-name row (computed, never typed). */
function weekdayShort(d) {
    return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][d.getDay()];
}
/** "M/D" short date used in headers. */
function fmtMonthDay(d) {
    return `${d.getMonth() + 1}/${d.getDate()}`;
}
