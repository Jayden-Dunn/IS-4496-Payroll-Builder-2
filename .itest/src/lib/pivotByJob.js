"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.fmtPercent = exports.EMPTY_MANUAL_STATE = void 0;
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
const format_1 = require("./format");
Object.defineProperty(exports, "fmtPercent", { enumerable: true, get: function () { return format_1.fmtPercent; } });
const payroll_1 = require("./payroll");
exports.EMPTY_MANUAL_STATE = { added: [], deleted: [], edits: {} };
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
    // Last-week notes / carry rows keyed by apKey.
    const lastWeekByAp = new Map();
    for (const r of input.lastWeekPbj ?? []) {
        if (!lastWeekByAp.has(r.apKey))
            lastWeekByAp.set(r.apKey, r);
    }
    // ----- dates -----------------------------------------------------------
    const dateMs = new Set();
    let hasUndated = false;
    for (const r of enriched) {
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
            const onCalendar = isOffice ? '-' : match ? `TimeNavi row ${match.sheetRow}` : 'NOT ON CALENDAR';
            // An edited AP# re-resolves the estimator from the matched event
            // (acceptance §10.3); otherwise the enriched-timesheet value rules.
            const baseEstimator = job.rows[0]?.estimator ?? '#N/A';
            const estimator = isOffice
                ? 'ND'
                : edit.apKey !== undefined && match
                    ? match.row.estimator || baseEstimator
                    : baseEstimator;
            const percent = isOffice ? null : (match ? match.row.percent : null);
            // Crew lead — from the matched event title (§3.2); OFFICE → HOURLY.
            const resolution = match
                ? (0, payroll_1.resolveCrewLead)((0, payroll_1.extractLeadParenthetical)(match.row.event, knownEstimators), leadCtx)
                : null;
            const crewLead = edit.crewLead !== undefined ? edit.crewLead : isOffice ? 'HOURLY' : (resolution?.display ?? '');
            const hourly = isHourlyJob(percent, crewLead, `${job.jobcode2} ${match?.row.event ?? ''}`);
            // WO Override — explicit $ in the description only.
            const woSrc = match ? (0, payroll_1.extractWoOverride)(match.row.description) : null;
            const woOverride = edit.woOverride !== undefined ? edit.woOverride : (woSrc?.amount ?? null);
            // WO% — "@ NN% PS" line beats the Settings map default.
            const pctSrc = match ? (0, payroll_1.extractWoPct)(match.row.description) : null;
            const defaultWoPct = hourly
                ? null
                : pctSrc
                    ? pctSrc.pct
                    : crewLead
                        ? (0, payroll_1.profitSharePctFor)(crewLead.split('/')[0], settings.profitShare, settings.defaultProfitSharePct)
                        : settings.defaultProfitSharePct;
            const woPct = edit.woPct !== undefined ? edit.woPct : defaultWoPct;
            const carried = apKey ? lastWeekByAp.get(apKey) : undefined;
            const notes = edit.notes !== undefined ? edit.notes : (carried?.notes ?? '');
            const analysis = {
                apKey,
                onCalendar,
                estimator: isOffice ? '-' : estimator,
                crewLead,
                crewLeadResolution: resolution,
                woOverride,
                woOverrideSource: woSrc?.line ?? null,
                woPct,
                woPctSource: pctSrc?.line ?? null,
                percent,
                notes,
                hourly,
                sub: resolution?.sub ?? false,
                split: resolution?.split ?? 1,
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
    const carryNote = (row) => {
        const paid = typeof row.percent === 'number' ? ` *PAID ${Math.round(row.percent * 100)}% LAST WK*` : '';
        return `${row.notes}${paid}`.trim();
    };
    for (const row of input.lastWeekPbj ?? []) {
        if (row.section === 'ut') {
            if (seenUtAp.has(row.apKey))
                continue;
            seenUtAp.add(row.apKey);
            utSeeds.push(manualRowFromCarry('ut', row, carryNote(row), colorByEstimator));
        }
        else {
            // NV pivot/balances rows: carry when paid < 100% and no hours this week.
            const partial = typeof row.percent === 'number' && row.percent < 1;
            const open = row.section === 'balances' ? row.percent !== 1 : partial;
            if (!open || timesheetApKeys.has(row.apKey) || seenBalancesAp.has(row.apKey))
                continue;
            seenBalancesAp.add(row.apKey);
            balancesSeeds.push(manualRowFromCarry('balances', row, carryNote(row), colorByEstimator));
        }
    }
    // UT seeds from this week's included Utah-calendar events (dedupe by apKey).
    for (const r of included) {
        if (!/^utah\b|^reno\b/i.test(r.calendar.trim()))
            continue;
        const apKey = r.apKey;
        if (!apKey || seenUtAp.has(apKey))
            continue;
        seenUtAp.add(apKey);
        const resolution = (0, payroll_1.resolveCrewLead)((0, payroll_1.extractLeadParenthetical)(r.event, knownEstimators), leadCtx);
        const woSrc = (0, payroll_1.extractWoOverride)(r.description);
        const pctSrc = (0, payroll_1.extractWoPct)(r.description);
        const hourly = isHourlyJob(r.percent, resolution.display, r.event);
        utSeeds.push({
            id: `ut:${apKey}`,
            customer: r.qboJobName || '',
            job: r.event,
            apKey,
            estimator: r.estimator,
            crewLead: resolution.display,
            woOverride: woSrc?.amount ?? null,
            woPct: hourly
                ? null
                : (pctSrc?.pct ??
                    (resolution.display
                        ? (0, payroll_1.profitSharePctFor)(resolution.display.split('/')[0], settings.profitShare, settings.defaultProfitSharePct)
                        : settings.defaultProfitSharePct)),
            percent: r.percent,
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
    };
}
function manualRowFromCarry(section, row, notes, colorByEstimator) {
    return {
        id: `${section}:${row.apKey}`,
        customer: row.customer,
        job: row.job,
        apKey: row.apKey,
        estimator: row.estimator,
        crewLead: row.crewLead,
        woOverride: row.woOverride,
        woPct: row.woPct,
        percent: row.percent,
        notes,
        colorHex: colorByEstimator.get(row.estimator) ?? null,
    };
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
        out.push({ ...seed, ...state.edits[seed.id] });
    }
    for (const added of state.added) {
        if (deleted.has(added.id) || seen.has(added.id))
            continue;
        out.push({ ...added, ...state.edits[added.id] });
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
