"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.plMetricsByJob = plMetricsByJob;
exports.buildPayrollPreview = buildPayrollPreview;
exports.hasUtPayrollData = hasUtPayrollData;
exports.mergePayrollPreviews = mergePayrollPreviews;
/**
 * Round 3 §5 / Round 4 §5 — "Payroll Preview": one block per person (hourly
 * painters + crew-lead profit-share blocks in a single A→Z sequence, then
 * office/admin/estimator/manager sections). Every derived number is an
 * editable default; edits recompute live.
 */
const apkey_1 = require("./apkey");
const format_1 = require("./format");
const jobNames_1 = require("./jobNames");
const payroll_1 = require("./payroll");
const payTypes_1 = require("./payTypes");
const pivotByJob_1 = require("./pivotByJob");
const byText = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' });
/** First name for alphabetical block ordering. */
function firstNameOf(name) {
    return name.trim().split(/\s+/)[0] ?? '';
}
const ENTITIES = ['NV', 'UT'];
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
function plMetricsByJob(pl, mapping, entity) {
    const out = new Map();
    if (!pl)
        return out;
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
    const labelsFor = (pick) => {
        const set = new Set();
        for (const e of scope) {
            for (const label of pick(mapping[e]) ?? []) {
                const t = label.trim().toLowerCase();
                if (t)
                    set.add(t);
            }
        }
        return [...set];
    };
    const metricLabels = {
        income: labelsFor((m) => m.income),
        materials: labelsFor((m) => m.materials),
        laborWagesSubs: labelsFor((m) => m.laborWagesSubs),
        sundries: labelsFor((m) => m.sundries),
        travel: labelsFor((m) => m.travel),
        mileage: labelsFor((m) => m.mileage),
    };
    // First occurrence per label, matching how `parsePl` reads the tree.
    const rowByLabel = new Map();
    for (const row of pl.accountRows) {
        const label = row.label.trim().toLowerCase();
        if (label && !rowByLabel.has(label))
            rowByLabel.set(label, row.values);
    }
    const sumAt = (labels, col) => {
        let acc = 0;
        for (const label of labels) {
            const v = rowByLabel.get(label)?.[col];
            if (typeof v === 'number')
                acc += v;
        }
        return acc;
    };
    for (let col = 1; col < pl.columnHeaders.length; col++) {
        const header = pl.columnHeaders[col];
        if (!header || /^total\b/i.test(header))
            continue;
        const apKey = (0, apkey_1.extractApKey)(header);
        if (!apKey)
            continue;
        const prev = out.get(apKey);
        const next = {
            income: sumAt(metricLabels.income, col),
            materials: sumAt(metricLabels.materials, col),
            laborWagesSubs: sumAt(metricLabels.laborWagesSubs, col),
            sundries: sumAt(metricLabels.sundries, col),
            travel: sumAt(metricLabels.travel, col),
            mileage: sumAt(metricLabels.mileage, col),
        };
        // Duplicate AP# columns sum, same as `parsePl`.
        out.set(apKey, prev
            ? {
                income: prev.income + next.income,
                materials: prev.materials + next.materials,
                laborWagesSubs: prev.laborWagesSubs + next.laborWagesSubs,
                sundries: prev.sundries + next.sundries,
                travel: prev.travel + next.travel,
                mileage: prev.mileage + next.mileage,
            }
            : next);
    }
    return out;
}
/** Split a resolved crew-lead display ("A" / "A/B") into matching keys. */
function leadNamesFromDisplay(display) {
    return display
        .split('/')
        .map((s) => s.replace(/\s+/g, ' ').trim().toUpperCase())
        .filter((s) => s !== '');
}
function buildPayrollPreview(input) {
    const { pbj, enriched, reportRows, settings, payDate, edits, qbHours } = input;
    // Round 8 §7.2/§7.3 — one builder, two tabs. The entity decides which
    // timesheet rows count, which jobs the bands carry, and which P&L the job
    // lookups resolve against.
    const entity = input.entity ?? 'NV';
    const isUt = entity === 'UT';
    const pl = isUt ? (input.plUt ?? null) : input.pl;
    const entityRows = enriched.filter((r) => r.sourceEntity === entity);
    const directory = input.directory ?? (0, payroll_1.directoryFromTimesheet)(enriched);
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
    const utJobByAp = new Map();
    for (const r of utJobs) {
        if (r.apKey && !utJobByAp.has(r.apKey))
            utJobByAp.set(r.apKey, r);
    }
    const nvJobByAp = new Map();
    for (const r of nvBalanceJobs) {
        if (r.apKey && !nvJobByAp.has(r.apKey))
            nvJobByAp.set(r.apKey, r);
    }
    const pbjJobByAp = new Map();
    for (const c of pbj.customers) {
        for (const j of c.jobs) {
            if (j.analysis.apKey && !pbjJobByAp.has(j.analysis.apKey))
                pbjJobByAp.set(j.analysis.apKey, j);
        }
    }
    const carriedIsHourly = (row) => (0, pivotByJob_1.isHourlyJob)(row.percent, row.crewLead, `${row.customer} ${row.job}`);
    const nameSources = {
        timesheetByAp: (0, jobNames_1.timesheetDisplayMap)(enriched),
        carryByAp: (0, jobNames_1.carryDisplayMap)(input.lastWeekPbj),
        salesByAp: (0, jobNames_1.salesDisplayMap)(input.nvSales, input.utSales),
        // Fall back to the NV tree so a UT-only week still resolves shared names.
        pl: pl ?? input.pl,
    };
    // ----- week range --------------------------------------------------------
    let weekStart = null;
    let weekEnd = null;
    for (const r of enriched) {
        if (!r.localDate)
            continue;
        if (!weekStart || r.localDate.getTime() < weekStart.getTime())
            weekStart = r.localDate;
        if (!weekEnd || r.localDate.getTime() > weekEnd.getTime())
            weekEnd = r.localDate;
    }
    const weekRange = weekStart && weekEnd ? `${(0, pivotByJob_1.fmtMonthDay)(weekStart)} - ${(0, pivotByJob_1.fmtMonthDay)(weekEnd)}` : '';
    const range = weekRange ? `(${weekRange})` : '';
    const title = `${isUt ? 'UT ' : ''}PAYROLL BREAKDOWN ${payDate ? (0, format_1.fmtDate)(payDate) : ''} ${range}`
        .replace(/\s+/g, ' ')
        .trim();
    const emps = new Map();
    for (const r of entityRows) {
        const key = r.username || `${r.fname} ${r.lname}`;
        if (emps.has(key))
            continue;
        const tsName = `${r.fname} ${r.lname}`.replace(/\s+/g, ' ').trim();
        const dirMatch = directory.find((d) => (0, payroll_1.tokensSubset)(tsName, d.fullName) || (0, payroll_1.tokensSubset)(d.fullName, tsName));
        const fullName = dirMatch?.fullName ?? tsName;
        const rosterInactive = settings.painterRoster.some((p) => !p.active && ((0, payroll_1.tokensSubset)(p.name, fullName) || (0, payroll_1.tokensSubset)(fullName, p.name)));
        emps.set(key, {
            username: key,
            fname: r.fname,
            lname: r.lname,
            group: r.group,
            fullName,
            displayName: (0, payroll_1.stripSingleLetterTokens)(fullName),
            inactive: (dirMatch ? !dirMatch.active : false) || rosterInactive,
        });
    }
    const empJobs = new Map(); // username → jobs in first-encounter order
    const empDayRows = new Map(); // username → (job, day) buckets
    const jobRowById = new Map(pbj.customers.flatMap((c) => c.jobs.map((j) => [j.id, j])));
    const eventByApKey = new Map();
    for (const r of reportRows) {
        if (r.included && r.apKey && !eventByApKey.has(r.apKey))
            eventByApKey.set(r.apKey, r);
    }
    /** Round 9 §16 — line-id suffix + label tag for a classified pay item. */
    const segKeyFor = (cls) => {
        if (cls.kind === 'overtime') {
            if (cls.reason === 'daily')
                return ':DOT';
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
    const crewLeadUsersByJob = new Map();
    entityRows.forEach((r, idx) => {
        if (r.hours === 0)
            return;
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
            const hourly = pbjJob?.analysis.hourly ??
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
                colorHex: r.apKey === 'OFFICE'
                    ? '4F81BD'
                    : (pbjJob?.colorHex ?? utRow?.colorHex ?? nvRow?.colorHex ?? colorByEstimator.get(r.estimator) ?? null),
                firstIndex: idx,
                lastDate: r.localDate,
                lastRowIndex: idx,
                office: r.apKey === 'OFFICE',
                jobText: `${r.jobcode1} ${r.jobcode2} ${eventTitle}`,
                forceNonDb: pbjJob?.analysis.dbOverride === true,
                leadNames: res && res.names.length > 0
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
        const cls = (0, payTypes_1.classifyPayItem)(r.payrollItem);
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
                set = new Set();
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
        if (r.sourceEntity === 'UT')
            dr.nv = false;
    });
    // Apply QB Hrs overrides (§3.3): keyed `${jobRowId}\u0000${username}`.
    // The day buckets and pay-item segments are scaled with them, so both daily
    // OT and the timesheet's own overtime lines reflect the same hours.
    for (const [username, jobs] of empJobs) {
        for (const j of jobs) {
            const o = qbHours[`${j.jobId}\u0000${username}`];
            if (o === undefined)
                continue;
            const dayRows = (empDayRows.get(username) ?? []).filter((d) => d.jobId === j.jobId);
            const raw = dayRows.reduce((a, d) => a + d.hours, 0);
            if (raw > 0) {
                for (const d of dayRows)
                    d.hours = (0, format_1.round2)((d.hours / raw) * o);
            }
            else if (dayRows.length > 0) {
                dayRows[dayRows.length - 1].hours = o;
            }
            const segRaw = j.segments.reduce((a, seg) => a + seg.hours, 0);
            if (segRaw > 0) {
                for (const seg of j.segments)
                    seg.hours = (0, format_1.round2)((seg.hours / segRaw) * o);
            }
            else if (j.segments.length > 0) {
                j.segments[j.segments.length - 1].hours = o;
            }
            j.hours = o;
        }
    }
    // ----- crew leads this week ----------------------------------------------
    // Names resolved on piece-rate NV pivot jobs + sub-only calendar jobs.
    const leadNames = new Map(); // UPPER full name
    const isSubName = (name) => settings.subcontractors.some((s) => (0, payroll_1.tokensSubset)(s, name) || (0, payroll_1.tokensSubset)(name, s));
    // The pivot tree is NV-only by construction, so the UT tab starts empty here
    // and fills its bands from `utJobs` instead.
    const pieceRateJobs = isUt
        ? []
        : pbj.customers.flatMap((c) => c.jobs
            .filter((j) => !j.analysis.hourly && j.analysis.apKey !== 'OFFICE')
            .map((j) => ({ job: j, customer: c.customer })));
    for (const row of [...utJobs, ...nvBalanceJobs]) {
        if (carriedIsHourly(row))
            continue;
        for (const n of leadNamesFromDisplay(row.crewLead)) {
            if (!leadNames.has(n))
                leadNames.set(n, { sub: isSubName(n) });
        }
    }
    for (const { job } of pieceRateJobs) {
        const res = job.analysis.crewLeadResolution;
        const names = res && res.names.length > 0 ? res.names : job.analysis.crewLead ? [job.analysis.crewLead] : [];
        for (const n of names) {
            const up = n.toUpperCase();
            if (!leadNames.has(up))
                leadNames.set(up, { sub: res?.sub ?? false });
        }
    }
    // Sub/no-hours calendar jobs (e.g. Pepe's Painting): included NV piece-rate
    // events with no timesheet hours whose lead resolves.
    const timesheetApKeys = new Set();
    for (const c of pbj.customers)
        for (const j of c.jobs)
            if (j.analysis.apKey)
                timesheetApKeys.add(j.analysis.apKey);
    const calOnlyJobs = [];
    // NV only — Utah/Reno calendar events already seeded the UT PAYROLL section.
    for (const r of isUt ? [] : reportRows) {
        if (!r.included || !r.apKey || timesheetApKeys.has(r.apKey))
            continue;
        if (/^utah\b|^reno\b/i.test(r.calendar.trim()))
            continue; // UT pays via Pivot by Job
        const res = (0, payroll_1.resolveCrewLead)((0, payroll_1.extractLeadParenthetical)(r.event, input.knownEstimators), input.leadCtx);
        if (res.names.length === 0)
            continue;
        if ((0, pivotByJob_1.isHourlyJob)(r.percent, res.display, r.event))
            continue;
        const woSrc = (0, payroll_1.extractWoOverride)(r.description);
        const pctSrc = (0, payroll_1.extractWoPct)(r.description);
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
            if (!leadNames.has(up))
                leadNames.set(up, { sub: res.sub });
        }
    }
    // Round 7 §20.1 — timesheet crew-group owners are leads even when no
    // piece-rate parenthetical named them this week (e.g. "Antonio Zuniga's Crew").
    const crewOwnerOf = (group) => {
        const m = /^(.*?)['\u2019]s\s+crew\s*$/i.exec(group.trim());
        return m ? m[1].trim() : null;
    };
    const addLeadName = (name, sub = false) => {
        const up = name.replace(/\s+/g, ' ').trim().toUpperCase();
        if (!up)
            return;
        if (!leadNames.has(up))
            leadNames.set(up, { sub });
    };
    const nameMatches = (a, b) => (0, payroll_1.tokensSubset)(a, b) ||
        (0, payroll_1.tokensSubset)(b, a) ||
        a.replace(/\s+/g, ' ').trim().toUpperCase() === b.replace(/\s+/g, ' ').trim().toUpperCase();
    const seenCrewOwners = new Set();
    for (const emp of emps.values()) {
        const owner = crewOwnerOf(emp.group);
        if (!owner)
            continue;
        const ownerKey = owner.toUpperCase();
        if (seenCrewOwners.has(ownerKey))
            continue;
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
        const clockedOwnPayType = ownerJobs.some((j) => j.segments.some((sg) => sg.rateKind === 'commercialHourly' || sg.rateKind === 'wages2'));
        if (clockedOwnPayType && !clockedCrewLead) {
            const ownerName = match?.fullName ?? owner;
            const namedOnPivot = pieceRateJobs.some((pj) => {
                const res = pj.job.analysis.crewLeadResolution;
                const names = res && res.names.length > 0
                    ? res.names
                    : pj.job.analysis.crewLead
                        ? [pj.job.analysis.crewLead]
                        : [];
                return names.some((n) => nameMatches(n, ownerName));
            });
            const inProfitShareMap = settings.profitShare.some((pentry) => nameMatches(pentry.lead, ownerName));
            if (!namedOnPivot && !inProfitShareMap)
                continue;
        }
        addLeadName(match?.fullName ?? owner);
    }
    // Optional boost: Settings profitShare leads with hours this week.
    for (const p of settings.profitShare) {
        const lead = p.lead.trim();
        if (!lead)
            continue;
        const match = [...emps.values()].find((e) => nameMatches(e.fullName, lead));
        if (!match)
            continue;
        const hrs = (empJobs.get(match.username) ?? []).reduce((a, j) => a + j.hours, 0);
        if (hrs > 0)
            addLeadName(match.fullName);
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
        const canonical = [];
        const matchesEmp = (n) => [...emps.values()].some((e) => nameMatches(e.fullName, n));
        for (const [name, meta] of entries) {
            const at = canonical.findIndex((c) => (0, payroll_1.tokensSubset)(c.key, name) || (0, payroll_1.tokensSubset)(name, c.key));
            if (at < 0) {
                canonical.push({ key: name, meta });
                continue;
            }
            const kept = canonical[at].key;
            // Prefer the spelling that matches a real timesheet employee; when both
            // (or neither) do, prefer the longer, more specific one.
            const preferNew = (matchesEmp(name) && !matchesEmp(kept)) ||
                (matchesEmp(name) === matchesEmp(kept) && name.length > kept.length);
            canonical[at] = {
                key: preferNew ? name : kept,
                meta: { sub: canonical[at].meta.sub || meta.sub },
            };
        }
        if (canonical.length !== entries.length) {
            leadNames.clear();
            for (const c of canonical)
                leadNames.set(c.key, c.meta);
        }
    }
    /** Is this timesheet employee one of the resolved crew leads? */
    const isLead = (fullName) => {
        for (const lead of leadNames.keys()) {
            if ((0, payroll_1.tokensSubset)(fullName, lead) || (0, payroll_1.tokensSubset)(lead, fullName))
                return true;
        }
        return false;
    };
    /** §5.9 — salaried-roster people never get an hourly block. */
    const isSalaried = (fullName) => settings.salariedRoster.some((s) => (0, payroll_1.tokensSubset)(s.name, fullName) || (0, payroll_1.tokensSubset)(fullName, s.name));
    // ----- Round 9 §19/§23/§24/§29 — person-scoped extra rows ------------------
    /** True when the pay date is the FIRST pay date of its month (1st–7th). */
    const isFirstPayDateOfMonth = payDate !== null && payDate.getDate() <= 7;
    /** The extra rows configured for one person, resolved for this pay date. */
    const extrasFor = (fullName, blockKey) => {
        const out = [];
        for (const item of settings.payrollLineItems ?? []) {
            if (!((0, payroll_1.tokensSubset)(item.person, fullName) || (0, payroll_1.tokensSubset)(fullName, item.person)))
                continue;
            if (item.frequency === 'firstPayDateOfMonth' && !isFirstPayDateOfMonth)
                continue;
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
    const extrasNet = (extras) => extras.reduce((acc, x) => {
        if (x.amount === null)
            return acc;
        if (x.kind === 'reimbursement')
            return acc + x.amount;
        if (x.kind === 'deduction')
            return acc - x.amount;
        return acc;
    }, 0);
    /**
     * Round 9 §1 — labor dollars per AP#, computed from the TIMESHEET: every
     * painter's hours on the job at their resolved line rate (overtime lines at
     * 1.5×). This is what the crew lead's "- Labor + Wages" row is charged, and
     * it comes from the timesheets tab of this side's input, never from the P&L.
     */
    const timesheetLaborByAp = new Map();
    {
        const empByUsername = new Map([...emps.values()].map((e) => [e.username, e]));
        for (const [username, jobs] of empJobs) {
            const emp = empByUsername.get(username);
            if (!emp)
                continue;
            const empIsLead = isLead(emp.fullName);
            for (const j of jobs) {
                if (!j.apKey || j.apKey === 'OFFICE')
                    continue;
                const leadUsers = crewLeadUsersByJob.get(j.jobId);
                const someoneElseLeads = leadUsers !== undefined && leadUsers.size > 0 && !leadUsers.has(username);
                const segs = j.segments.length > 0
                    ? j.segments
                    : [{ key: '', payItem: '', rateKind: 'regular', overtime: false, otReason: null, hours: j.hours, firstIndex: j.firstIndex }];
                let acc = 0;
                for (const sg of segs) {
                    const info = (0, payroll_1.resolveLineRate)(emp.fullName, j.jobText, j.hourly, {
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
                timesheetLaborByAp.set(j.apKey, (0, format_1.round2)((timesheetLaborByAp.get(j.apKey) ?? 0) + acc));
            }
        }
    }
    // ----- hourly block construction ------------------------------------------
    const dailyOtNotes = [];
    /** Round 9 §27 — people the NV daily-OT rule never applies to. */
    const dailyOtExempt = (fullName) => (settings.dailyOtExemptNames ?? []).some((n) => (0, payroll_1.tokensSubset)(n, fullName) || (0, payroll_1.tokensSubset)(fullName, n));
    const buildHourlyBlock = (emp, officeStyle, isLead = false) => {
        const jobs = empJobs.get(emp.username) ?? [];
        /**
         * Round 9 §13/§18 — the rate is resolved PER SEGMENT, because the
         * timesheet pay item decides which rate slot the hours are paid from
         * (Regular Pay / Crew Lead - Hourly Wage / Commercial Hourly / Wages 2).
         * `isCrewLead` still applies when the pay item itself is not crew-lead
         * but this IS the lead's own block and no one else on the job clocked
         * crew-lead hours.
         */
        const rateFor = (j, seg) => {
            const payRateKind = seg?.rateKind ?? 'regular';
            // §13 — when someone else on this job clocked "Crew Lead - Hourly Wage"
            // and this person did not, this person is NOT the lead on that job.
            const leadUsers = crewLeadUsersByJob.get(j.jobId);
            const someoneElseLeads = leadUsers !== undefined && leadUsers.size > 0 && !leadUsers.has(emp.username);
            const isCrewLeadHere = payRateKind === 'crewLead' ? true : someoneElseLeads ? false : isLead;
            return (0, payroll_1.resolveLineRate)(emp.fullName, j.jobText, j.hourly, {
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
        const straightRate = (jobId) => {
            const j = jobs.find((x) => x.jobId === jobId);
            if (!j)
                return settings.defaultWageRate;
            const seg = j.segments.find((sg) => !sg.overtime) ?? null;
            const e = edits[`hl:${emp.username}\u0000${jobId}${seg?.key ?? ''}`] ?? {};
            if (e.rate !== undefined)
                return e.rate;
            return rateFor(j, seg).rate;
        };
        // Round 9 §16 — the timesheet's own overtime rows are authoritative. Only
        // when the week carries NONE do the computed >8/day and >40/week rules run,
        // so a QuickBooks file without OT items still splits overtime correctly.
        const hasTimesheetOt = jobs.some((j) => j.segments.some((sg) => sg.overtime));
        const daily = hasTimesheetOt
            ? { otByJob: new Map(), segments: [], totalOtHours: 0 }
            : (0, payroll_1.splitDailyOvertime)((empDayRows.get(emp.username) ?? []).map((d) => ({
                jobId: d.jobId,
                date: d.date,
                hours: d.hours,
                rowIndex: d.rowIndex,
                eligible: d.nv &&
                    !dailyOtExempt(emp.fullName) &&
                    straightRate(d.jobId) < settings.nvDailyOtRateThreshold,
            })), settings.dailyOtHoursThreshold);
        // §3.4 weekly overtime on the QB-adjusted hours (computed fallback only).
        const ot = hasTimesheetOt
            ? { regularByJob: new Map(), otJobId: null, otHours: 0 }
            : (0, payroll_1.splitOvertime)(jobs.map((j) => ({
                jobId: j.jobId,
                hours: j.hours,
                lastDate: j.lastDate,
                lastRowIndex: j.lastRowIndex,
            })), { dailyOtHours: daily.totalOtHours });
        const lines = [];
        const pushLine = (j, seg, hours, kind, segment) => {
            const isOt = kind !== 'regular';
            const computedSuffix = kind === 'weeklyOt' ? ':OT' : kind === 'dailyOt' ? `:DOT:${segment?.dateIso ?? ''}` : '';
            // Timesheet-driven lines key off the pay item; computed OT keeps the
            // Round 8 suffixes so existing manual edits still land.
            const suffix = seg ? seg.key : computedSuffix;
            const lineId = `${emp.username}\u0000${j.jobId}${suffix}`;
            const e = edits[`hl:${lineId}`] ?? {};
            const rateInfo = rateFor(j, seg);
            const baseRate = e.rate !== undefined ? e.rate : rateInfo.rate;
            const rate = isOt ? (0, format_1.round2)(baseRate * 1.5) : baseRate;
            const lineHours = e.hours !== undefined ? e.hours : (0, format_1.round2)(hours);
            // Round 8 §9.2 — the two OT reasons carry distinguishable tags.
            const dayLabel = segment?.date ? (0, pivotByJob_1.fmtMonthDay)(segment.date) : '';
            const dailyTag = segment
                ? `[Daily OT — ${dayLabel}: ${(0, format_1.round2)(hours)} hrs over ${settings.dailyOtHoursThreshold} @ 1.5× (rate < $${settings.nvDailyOtRateThreshold}, NV)]`
                : seg?.overtime
                    ? `[${seg.payItem} — timesheet overtime line, paid at 1.5× $${baseRate.toFixed(2)}]`
                    : undefined;
            const otTag = seg
                ? (0, payTypes_1.payItemTag)(seg.payItem)
                : kind === 'weeklyOt'
                    ? ' [OT >40 hrs/wk]'
                    : kind === 'dailyOt'
                        ? ' [Daily OT]'
                        : '';
            const isOfficeLine = officeStyle || j.office;
            const notesSuffix = !isOfficeLine && j.notes.trim() ? ` — ${j.notes.trim()}` : '';
            const typeTag = isOfficeLine ? '' : j.hourly ? ' [Hourly]' : ' [PS]';
            const label = isOfficeLine
                ? `Office${otTag}`
                : `${(0, payroll_1.jobDisplayName)(j.customer, j.jobcode2)}${typeTag}${otTag}${notesSuffix}`;
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
                    jobLabel: isOfficeLine ? 'Office' : (0, payroll_1.jobDisplayName)(j.customer, j.jobcode2),
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
                ruleNote: e.rate === undefined &&
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
        const orderedSegments = (j) => j.segments
            .filter((sg) => (0, format_1.round2)(sg.hours) !== 0)
            .slice()
            .sort((a, b) => Number(a.overtime) - Number(b.overtime) || a.firstIndex - b.firstIndex);
        if (officeStyle && jobs.length > 0) {
            // §5.3 — one "Office" line per pay item covering the week, so an office
            // overtime row still shows separately at its own rate.
            const rep = jobs[0];
            const merged = new Map();
            for (const j of jobs) {
                for (const sg of orderedSegments(j)) {
                    const at = merged.get(sg.key);
                    if (at)
                        at.hours = (0, format_1.round2)(at.hours + sg.hours);
                    else
                        merged.set(sg.key, { ...sg });
                }
            }
            const officeSegs = [...merged.values()].sort((a, b) => Number(a.overtime) - Number(b.overtime) || a.firstIndex - b.firstIndex);
            if (hasTimesheetOt || officeSegs.length > 1) {
                for (const sg of officeSegs)
                    pushLine(rep, sg, sg.hours, sg.overtime ? 'weeklyOt' : 'regular');
            }
            else {
                const totalHrs = jobs.reduce((a, j) => a + j.hours, 0);
                const regular = (0, format_1.round2)(Math.max(0, totalHrs - daily.totalOtHours - ot.otHours));
                pushLine(rep, officeSegs[0] ?? null, regular, 'regular');
                for (const seg of daily.segments)
                    pushLine(rep, null, seg.otHours, 'dailyOt', seg);
                if (ot.otHours > 0)
                    pushLine(rep, null, ot.otHours, 'weeklyOt');
            }
        }
        else {
            for (const j of jobs.slice().sort((a, b) => a.firstIndex - b.firstIndex)) {
                const segs = orderedSegments(j);
                if (hasTimesheetOt) {
                    // Round 9 §10/§16 — straight-time and overtime hours keep their own
                    // lines exactly as the timesheet split them; nothing is re-derived,
                    // so a job never pays ALL of its hours at the overtime rate.
                    for (const sg of segs)
                        pushLine(j, sg, sg.hours, sg.overtime ? 'weeklyOt' : 'regular');
                    continue;
                }
                // Computed fallback: carve the day's / week's overtime out of this
                // job's straight-time segments, last segment first.
                const dailyOnJob = daily.otByJob.get(j.jobId) ?? 0;
                const weeklyOnJob = ot.otJobId === j.jobId ? ot.otHours : 0;
                let toRemove = (0, format_1.round2)(dailyOnJob + weeklyOnJob);
                const straight = segs.filter((sg) => !sg.overtime);
                const regularBySeg = new Map(straight.map((sg) => [sg.key, (0, format_1.round2)(sg.hours)]));
                for (let i = straight.length - 1; i >= 0 && toRemove > 0; i--) {
                    const sg = straight[i];
                    const have = regularBySeg.get(sg.key) ?? 0;
                    const take = (0, format_1.round2)(Math.min(have, toRemove));
                    regularBySeg.set(sg.key, (0, format_1.round2)(have - take));
                    toRemove = (0, format_1.round2)(toRemove - take);
                }
                for (const sg of straight)
                    pushLine(j, sg, regularBySeg.get(sg.key) ?? 0, 'regular');
                if (straight.length === 0 && segs.length === 0) {
                    pushLine(j, null, (0, format_1.round2)(Math.max(0, j.hours - dailyOnJob - weeklyOnJob)), 'regular');
                }
                for (const seg of daily.segments.filter((sd) => sd.jobId === j.jobId)) {
                    pushLine(j, null, seg.otHours, 'dailyOt', seg);
                }
                if (weeklyOnJob > 0)
                    pushLine(j, null, weeklyOnJob, 'weeklyOt');
            }
        }
        const totalHours = (0, format_1.round2)(lines.reduce((a, l) => a + l.hours, 0));
        const totalPay = (0, format_1.round2)(lines.reduce((a, l) => a + l.hours * l.rate, 0));
        const blockEdit = edits[`hb:${emp.username}`] ?? {};
        // Round 9 §19/§23 — person-scoped extra rows (Gas Reimb, Aflac, …).
        const extras = extrasFor(emp.fullName, `hb:${emp.username}`);
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
            netPay: (0, format_1.round2)(totalPay + extrasNet(extras)),
        };
    };
    // ----- sections ------------------------------------------------------------
    const officeEmps = [];
    const painterEmps = [];
    const leadEmps = [];
    for (const emp of emps.values()) {
        if (isSalaried(emp.fullName))
            continue; // §5.9 — roster block only
        if (/office/i.test(emp.group))
            officeEmps.push(emp);
        else if (isLead(emp.fullName))
            leadEmps.push(emp);
        else
            painterEmps.push(emp);
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
    const newLeadsToAdd = new Map();
    const leadMapPct = (leadUpper) => {
        const entry = settings.profitShare.find((e) => (0, payroll_1.tokensSubset)(e.lead, leadUpper) || (0, payroll_1.tokensSubset)(leadUpper, e.lead));
        if (entry)
            return entry.pct;
        if (!newLeadsToAdd.has(leadUpper)) {
            newLeadsToAdd.set(leadUpper, {
                lead: (0, payroll_1.titleCaseLeadName)(leadUpper, directory, settings.subcontractors),
                pct: settings.defaultProfitSharePct,
            });
        }
        return settings.defaultProfitSharePct;
    };
    /** Lead's own timesheet hours on an AP# (QB-adjusted). */
    const leadHoursOnAp = (leadUpper, apKey) => {
        if (!apKey)
            return 0;
        const emp = leadEmps.find((e) => (0, payroll_1.tokensSubset)(e.fullName, leadUpper) || (0, payroll_1.tokensSubset)(leadUpper, e.fullName));
        if (!emp)
            return 0;
        return (0, format_1.round2)((empJobs.get(emp.username) ?? []).filter((j) => j.apKey === apKey).reduce((a, j) => a + j.hours, 0));
    };
    const jobTextForAp = (apKey, customer, jobShort) => {
        const eventTitle = apKey ? (eventByApKey.get(apKey)?.event ?? '') : '';
        return `${customer} ${jobShort} ${eventTitle}`;
    };
    const buildPsJob = (blockKey, jobRowId, apKey, jobShort, customer, pivotWoPct, pivotWoOverride, leadDefaultWoPct, percent, split, colorHex, leadUpper, isSub) => {
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
        const laborWages = e.laborWages !== undefined
            ? e.laborWages
            : timesheetLabor > 0
                ? timesheetLabor
                : (metrics?.laborWagesSubs ?? 0);
        const mileage = e.mileage !== undefined ? e.mileage : (metrics?.mileage ?? 0);
        const travel = metrics?.travel ?? 0;
        const sundries = metrics?.sundries ?? 0;
        // Round 9 §32 — a lead responsible for ALL of a job's materials is charged
        // the full amount instead of half.
        const fullMaterials = (settings.fullMaterialsLeads ?? []).some((n) => (0, payroll_1.tokensSubset)(n, leadUpper) || (0, payroll_1.tokensSubset)(leadUpper, n));
        const halfMaterials = e.halfMaterials !== undefined ? e.halfMaterials : fullMaterials ? materials : materials / 2;
        // Round 6 §6.2 — WO% from the Pivot by Job `WO%` cell (Round 9 §15: that
        // column is the source of truth), else the Settings crew map (Round 8 §9.4:
        // the crew map now always answers, auto-adding if needed). Round 9 §14 —
        // a typed % on the preview band header wins over both.
        const effectiveWoPct = e.woPct !== undefined ? e.woPct : (pivotWoPct ?? leadDefaultWoPct);
        const pct = effectiveWoPct ?? 0;
        const workOrderOverride = e.workOrderOverride !== undefined ? e.workOrderOverride : pivotWoOverride;
        // Round 9 §8/§20 — the Pivot by Job `WO Override` amount IS the work-order
        // dollar figure, so it drives Profit Share whenever it is filled in. Only
        // an empty override falls back to Total Income × WO%, which is what left
        // AP24097-E at $0 when its P&L income was zero.
        const profitShare = workOrderOverride !== null && workOrderOverride !== undefined
            ? workOrderOverride
            : income * pct;
        // A typed % Completion on the payroll band overrides the grid value.
        const effectivePercent = e.percentCompletion !== undefined ? e.percentCompletion : percent;
        const chain = (0, payroll_1.computeTotalProfitShare)({
            profitShare,
            workOrderOverride,
            laborWages,
            halfMaterials,
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
        const isDbJob = pbjJobForDb !== undefined
            ? pbjJobForDb.analysis.isDbJob
            : (0, payroll_1.isPrevailingWageJob)(jobText, settings.prevailingWagePatterns);
        const hoursOnJob = leadHoursOnAp(leadUpper, apKey);
        const floorResult = (0, payroll_1.applyDavisBaconLeadFloor)({
            // A hand-paid SUBBED / SET RATE job has no computed chain to floor.
            isDbJob: isDbJob && chain.manualPay === null,
            isSubcontractor: isSub,
            timesheetHoursOnJob: hoursOnJob,
            profitSharePay: totalAfterSplit,
            dbRate: settings.prevailingMinRate,
        });
        if (floorResult.floorApplied) {
            totalAfterSplit = floorResult.pay;
            if (splitWith <= 1)
                totalProfitShare = floorResult.pay;
        }
        const pbjNotes = (() => {
            if (apKey) {
                const job = pbjJobByAp.get(apKey);
                if (job?.analysis.notes)
                    return job.analysis.notes;
                const carried = utJobByAp.get(apKey) ?? nvJobByAp.get(apKey);
                if (carried?.notes)
                    return carried.notes;
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
    const bandNames = (apKey, display, pivotJobText, pivotCustomer) => {
        const po = display.poLabel.trim() ||
            pivotJobText.trim() ||
            (apKey && apKey !== 'OFFICE' ? `AP${apKey}` : '');
        const cust = pivotCustomer.trim() && pivotCustomer.trim() !== '(blank)'
            ? pivotCustomer.trim()
            : display.customerName.trim() ||
                (apKey ? (plByJob.has(apKey) ? (pl?.byJob.get(apKey)?.customerName ?? '') : '') : '');
        return { poLabel: po, customerName: cust };
    };
    const leads = [];
    const leadMatchesName = (candidate, leadUpper) => candidate.toUpperCase() === leadUpper ||
        (0, payroll_1.tokensSubset)(candidate, leadUpper) ||
        (0, payroll_1.tokensSubset)(leadUpper, candidate);
    for (const [leadUpper, meta] of leadNames) {
        // A lead working in both states gets a band on each tab; the UT band needs
        // its own edit namespace so its Reimbursements/Deductions stay separate.
        const blockKey = isUt ? `lead:UT:${leadUpper}` : `lead:${leadUpper}`;
        const jobs = [];
        const seenJobIds = new Set();
        const seenAps = new Set();
        const markJob = (jobId, apKey) => {
            if (seenJobIds.has(jobId))
                return false;
            if (apKey && seenAps.has(apKey))
                return false;
            seenJobIds.add(jobId);
            if (apKey)
                seenAps.add(apKey);
            return true;
        };
        for (const { job } of pieceRateJobs) {
            const names = job.analysis.crewLeadResolution && job.analysis.crewLeadResolution.names.length > 0
                ? job.analysis.crewLeadResolution.names
                : job.analysis.crewLead
                    ? [job.analysis.crewLead.toUpperCase()]
                    : [];
            if (!names.some((n) => leadMatchesName(n, leadUpper)))
                continue;
            if (!markJob(job.id, job.analysis.apKey))
                continue;
            const display = bandNames(job.analysis.apKey, (0, jobNames_1.jobDisplayPair)(job.analysis.apKey, {
                ...nameSources,
                eventTitle: eventByApKey.get(job.analysis.apKey ?? '')?.event,
            }), job.jobcode2, job.id.split('\u0000')[0]);
            jobs.push(buildPsJob(blockKey, job.id, job.analysis.apKey, display.poLabel, display.customerName, job.analysis.woPct, job.analysis.woOverride, leadMapPct(leadUpper), job.analysis.percent, job.analysis.split, job.colorHex, leadUpper, meta.sub));
        }
        for (const cj of calOnlyJobs) {
            if (!cj.lead.some((n) => leadMatchesName(n, leadUpper)))
                continue;
            const calId = `cal:${cj.row.id}`;
            if (!markJob(calId, cj.row.apKey))
                continue;
            const display = bandNames(cj.row.apKey, (0, jobNames_1.jobDisplayPair)(cj.row.apKey, { ...nameSources, eventTitle: cj.row.event }), '', '');
            jobs.push(buildPsJob(blockKey, calId, cj.row.apKey, display.poLabel, display.customerName, cj.woPct, cj.woOverride, leadMapPct(leadUpper), cj.row.percent, cj.split, colorByEstimator.get(cj.row.estimator) ?? null, leadUpper, meta.sub));
        }
        // Round 8 §6 / §7.3 — carried no-timesheet jobs (NV BALANCES or UT PAYROLL)
        // join this lead's band. HOURLY jobs stay off the PS chain.
        for (const row of isUt ? utJobs : nvBalanceJobs) {
            if (carriedIsHourly(row))
                continue;
            const names = leadNamesFromDisplay(row.crewLead);
            if (!names.some((n) => leadMatchesName(n, leadUpper)))
                continue;
            if (!markJob(row.id, row.apKey))
                continue;
            const display = bandNames(row.apKey, (0, jobNames_1.jobDisplayPair)(row.apKey, {
                ...nameSources,
                eventTitle: eventByApKey.get(row.apKey ?? '')?.event,
            }), row.job, row.customer);
            jobs.push(buildPsJob(blockKey, row.id, row.apKey, display.poLabel, display.customerName, row.woPct, row.woOverride, leadMapPct(leadUpper), row.percent, Math.max(1, names.length), row.colorHex, leadUpper, meta.sub));
        }
        const leadEmp = leadEmps.find((e) => (0, payroll_1.tokensSubset)(e.fullName, leadUpper) || (0, payroll_1.tokensSubset)(leadUpper, e.fullName));
        // Crew-owner / profitShare leads with no parenthetical jobs: attach their
        // own piece-rate timesheet APs so they still get a PS band (not a painter block).
        if (jobs.length === 0 && leadEmp) {
            for (const ej of empJobs.get(leadEmp.username) ?? []) {
                if (ej.hourly || ej.office || !ej.apKey)
                    continue;
                const pbjJob = jobRowById.get(ej.jobId);
                const utRow = utJobByAp.get(ej.apKey);
                if (pbjJob && (pbjJob.analysis.hourly || pbjJob.analysis.apKey === 'OFFICE'))
                    continue;
                if (!pbjJob && !utRow)
                    continue;
                if (!markJob(ej.jobId, ej.apKey))
                    continue;
                const apKey = pbjJob?.analysis.apKey ?? ej.apKey;
                const display = bandNames(apKey, (0, jobNames_1.jobDisplayPair)(apKey, {
                    ...nameSources,
                    eventTitle: eventByApKey.get(apKey ?? '')?.event,
                }), pbjJob?.jobcode2 || utRow?.job || ej.jobcode2, utRow?.customer || ej.customer);
                jobs.push(buildPsJob(blockKey, ej.jobId, apKey, display.poLabel, display.customerName, pbjJob?.analysis.woPct ?? utRow?.woPct ?? null, pbjJob?.analysis.woOverride ?? utRow?.woOverride ?? null, leadMapPct(leadUpper), pbjJob?.analysis.percent ?? utRow?.percent ?? null, pbjJob?.analysis.split ?? 1, pbjJob?.colorHex ?? utRow?.colorHex ?? ej.colorHex, leadUpper, meta.sub));
            }
        }
        // Skip pure name-only entries with neither jobs nor timesheet presence.
        if (jobs.length === 0 && !leadEmp)
            continue;
        const tsheets = leadEmp && !meta.sub ? buildHourlyBlock(leadEmp, false, true) : null;
        const be = edits[`psb:${blockKey}`] ?? {};
        const total = jobs.reduce((a, j) => a + j.totalAfterSplit, 0);
        const mileageNew = be.mileageNew !== undefined ? be.mileageNew : 0;
        const hourlyPay = be.hourlyPay !== undefined ? be.hourlyPay : tsheets ? tsheets.totalPay : 0;
        const reimbursements = be.reimbursements !== undefined ? be.reimbursements : 0;
        // Round 8 §8.3 — the P&L sundries left materials; they are the DEFAULT
        // Deductions/Sundries entry now (a manual edit still wins).
        const sundriesDefault = (0, format_1.round2)(jobs.reduce((a, j) => a + j.sundries, 0));
        const deductions = be.deductions !== undefined ? be.deductions : sundriesDefault;
        const deductMemo = be.deductMemo ?? (be.deductions === undefined && sundriesDefault !== 0 ? 'Sundries (P&L)' : '');
        // Round 9 §19/§23 — extra rows on the lead's band (Aflac, Gas Reimb, …).
        const leadDisplayName = (0, payroll_1.titleCaseLeadName)(leadUpper, directory, settings.subcontractors);
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
            totalPay: (0, format_1.round2)((0, payroll_1.computeTotalPay)({ total, hourlyPay, reimbursements, deductions }) + extrasNet(extras)),
            tsheets,
            extras,
            // Round 9 §31 — hours pay + profit share, for the Final Payroll Breakdown.
            totalGrossWages: (0, format_1.round2)(total + hourlyPay),
        });
    }
    leads.sort((a, b) => byText(firstNameOf(a.displayName), firstNameOf(b.displayName)));
    // §5.2 — one alphabetical sequence: painters + lead blocks interleaved A→Z
    // by first name; subcontractor lead blocks last (A→Z among themselves).
    const painterSequence = [
        ...painters.map((b) => ({ kind: 'hourly', block: b })),
        ...leads.filter((b) => !b.sub).map((b) => ({ kind: 'ps', block: b })),
    ].sort((a, b) => byText(firstNameOf(a.kind === 'hourly' ? a.block.employeeName : a.block.displayName), firstNameOf(b.kind === 'hourly' ? b.block.employeeName : b.block.displayName)));
    for (const b of leads.filter((x) => x.sub))
        painterSequence.push({ kind: 'ps', block: b });
    // ----- salary sections ---------------------------------------------------
    const salaryBlock = (entry) => {
        const key = `sal:${entry.name}`;
        const e = edits[key] ?? {};
        const tsEmp = [...emps.values()].find((x) => (0, payroll_1.tokensSubset)(x.fullName, entry.name) || (0, payroll_1.tokensSubset)(entry.name, x.fullName));
        const hours = tsEmp
            ? (0, format_1.round2)((empJobs.get(tsEmp.username) ?? []).reduce((a, j) => a + j.hours, 0))
            : 0;
        const miles = e.miles !== undefined ? e.miles : 0;
        const mileageReimb = (0, format_1.round2)(miles * settings.mileageRate);
        // Round 9 §28 — Jodi Dunn's row reads "Reimbursements" and carries a flat
        // amount instead of miles × rate.
        const flatReimb = entry.reimbursementsLabel
            ? (e.reimbursements !== undefined ? e.reimbursements : 0)
            : 0;
        const extras = extrasFor(entry.name, key);
        return {
            key,
            name: (0, payroll_1.stripSingleLetterTokens)(entry.name),
            role: entry.role,
            amount: entry.amount,
            section: entry.section,
            timesheetHours: hours,
            miles,
            mileageReimb,
            sub: entry.sub === true,
            ...(entry.reimbursementsLabel ? { reimbursementsLabel: entry.reimbursementsLabel } : {}),
            reimbursements: flatReimb,
            extras,
            taxes: e.taxes !== undefined ? e.taxes : (entry.taxes ?? null),
            net: e.net !== undefined ? e.net : (entry.net ?? null),
            total: (0, format_1.round2)(entry.amount +
                (entry.reimbursementsLabel ? flatReimb : mileageReimb) +
                extrasNet(extras)),
        };
    };
    // Round 8 §7.3 — each tab shows only its own entity's salaried people, so the
    // UT tab is painters/leads-only until someone is tagged UT in Settings.
    const roster = settings.salariedRoster.filter((r) => r.entity === entity &&
        // Round 9 §25 — monthly people appear only on the month's first pay date.
        (r.frequency !== 'firstPayDateOfMonth' || isFirstPayDateOfMonth));
    // Round 9 §26 — subcontractor rows sort to the very bottom of their section.
    const sortSalary = (a, b) => Number(a.sub) - Number(b.sub) || byText(firstNameOf(a.name), firstNameOf(b.name));
    const admin = roster.filter((r) => r.section === 'ADMIN').map(salaryBlock).sort(sortSalary);
    const estimators = roster.filter((r) => r.section === 'ESTIMATORS').map(salaryBlock).sort(sortSalary);
    const managers = roster.filter((r) => r.section === 'MANAGER SALARY').map(salaryBlock).sort(sortSalary);
    // ----- Round 6 §4.10 Inactive Employees tab --------------------------------
    const norm = (s) => s.trim().toLowerCase();
    const namesWithHours = new Set();
    for (const b of painters) {
        if (b.totalHours > 0)
            namesWithHours.add(norm(b.employeeName));
    }
    for (const b of office) {
        if (b.totalHours > 0)
            namesWithHours.add(norm(b.employeeName));
    }
    for (const b of leads)
        namesWithHours.add(norm(b.leadName));
    for (const entry of painterSequence) {
        if (entry.kind === 'ps')
            namesWithHours.add(norm(entry.block.leadName));
        else if (entry.block.totalHours > 0)
            namesWithHours.add(norm(entry.block.employeeName));
    }
    const salariedNames = new Set(settings.salariedRoster.map((r) => norm(r.name)));
    const leadNameSet = new Set([
        ...settings.profitShare.map((p) => norm(p.lead)),
        ...settings.subcontractors.map((s) => norm(s)),
    ]);
    const isKnownLead = (name) => leadNameSet.has(norm(name));
    const inactiveCands = new Map();
    for (const d of directory) {
        if (d.active)
            continue;
        const key = norm(d.fullName);
        if (namesWithHours.has(key) || salariedNames.has(key))
            continue;
        inactiveCands.set(key, { name: d.fullName, directoryInactive: true });
    }
    // Round 8 §11 — the Inactive Employees sheet is the persisted painter roster,
    // regenerated every week until someone toggles a name in Settings.
    for (const p of settings.painterRoster) {
        if (p.active)
            continue;
        const key = norm(p.name);
        if (namesWithHours.has(key) || salariedNames.has(key))
            continue;
        if (!inactiveCands.has(key))
            inactiveCands.set(key, { name: p.name, directoryInactive: true });
    }
    // The Inactive Employees tab is company-wide, so only the NV model fills it;
    // otherwise the merged export would list every absent painter twice.
    const inactiveEmployees = (isUt ? [] : [...inactiveCands.values()])
        .map((c) => {
        const rateInfo = (0, payroll_1.wageRateFor)(c.name, settings.wageRates, settings.defaultWageRate, false);
        return {
            name: c.name,
            displayName: (0, payroll_1.stripSingleLetterTokens)(c.name),
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
        dailyOtNotes,
    };
}
/**
 * Round 8 §7.3 — does this week have any Utah payroll at all? Drives both the
 * `Payroll Preview - UT` empty state and the byte-parity rule that an NV-only
 * week emits no UT sheets.
 */
function hasUtPayrollData(args) {
    if (args.plUt)
        return true;
    if (args.pbj.utPayroll.some((r) => !r.excludedFromPayroll))
        return true;
    return args.enriched.some((r) => r.sourceEntity === 'UT' && r.hours !== 0);
}
/**
 * Round 8 §7.3 — one model carrying BOTH entities' people, for the downstream
 * shared tabs (Profitshare Output, Input #'s, NOTES, Management Report Export,
 * …). Every band keeps the numbers its own entity's P&L produced; only the
 * ordering is redone so the combined sheets still read A→Z with subs last.
 */
function mergePayrollPreviews(nv, ut) {
    if (!ut)
        return nv;
    const leads = [...nv.leads, ...ut.leads];
    const painters = [...nv.painters, ...ut.painters].sort((a, b) => byText(firstNameOf(a.employeeName), firstNameOf(b.employeeName)));
    const painterSequence = [
        ...painters.map((b) => ({ kind: 'hourly', block: b })),
        ...leads.filter((b) => !b.sub).map((b) => ({ kind: 'ps', block: b })),
    ].sort((a, b) => byText(firstNameOf(a.kind === 'hourly' ? a.block.employeeName : a.block.displayName), firstNameOf(b.kind === 'hourly' ? b.block.employeeName : b.block.displayName)));
    for (const b of leads.filter((x) => x.sub))
        painterSequence.push({ kind: 'ps', block: b });
    const byFirstName = (a, b) => 
    // Round 9 §26 — subcontractor rows stay at the bottom of the merged section.
    Number(a.sub) - Number(b.sub) || byText(firstNameOf(a.name), firstNameOf(b.name));
    const seenLead = new Set(nv.newLeadsToAdd.map((n) => n.lead.trim().toLowerCase()));
    return {
        ...nv,
        leads,
        painters,
        painterSequence,
        office: [...nv.office, ...ut.office].sort((a, b) => byText(firstNameOf(a.employeeName), firstNameOf(b.employeeName))),
        admin: [...nv.admin, ...ut.admin].sort(byFirstName),
        estimators: [...nv.estimators, ...ut.estimators].sort(byFirstName),
        managers: [...nv.managers, ...ut.managers].sort(byFirstName),
        newLeadsToAdd: [
            ...nv.newLeadsToAdd,
            ...ut.newLeadsToAdd.filter((n) => !seenLead.has(n.lead.trim().toLowerCase())),
        ],
        dailyOtNotes: [...nv.dailyOtNotes, ...ut.dailyOtNotes],
    };
}
