"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildPayrollPreview = buildPayrollPreview;
/**
 * Round 3 §5 — "Payroll Preview": one block per person (hourly painters,
 * crew-lead profit-share blocks, office/admin/estimator/manager sections).
 * Every derived number is an editable default; edits recompute live.
 */
const format_1 = require("./format");
const payroll_1 = require("./payroll");
const pivotByJob_1 = require("./pivotByJob");
const byText = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' });
/** First name for alphabetical block ordering. */
function firstNameOf(name) {
    return name.trim().split(/\s+/)[0] ?? '';
}
function buildPayrollPreview(input) {
    const { pbj, enriched, reportRows, pl, settings, payDate, edits, qbHours } = input;
    const directory = input.directory ?? (0, payroll_1.directoryFromTimesheet)(enriched);
    const colorByEstimator = new Map(settings.legend.map((l) => [l.estimator, l.color]));
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
    const range = weekStart && weekEnd ? `(${(0, pivotByJob_1.fmtMonthDay)(weekStart)} - ${(0, pivotByJob_1.fmtMonthDay)(weekEnd)})` : '';
    const title = `PAYROLL BREAKDOWN ${payDate ? (0, format_1.fmtDate)(payDate) : ''} ${range}`.replace(/\s+/g, ' ').trim();
    const emps = new Map();
    for (const r of enriched) {
        const key = r.username || `${r.fname} ${r.lname}`;
        if (emps.has(key))
            continue;
        const tsName = `${r.fname} ${r.lname}`.replace(/\s+/g, ' ').trim();
        const dirMatch = directory.find((d) => (0, payroll_1.tokensSubset)(tsName, d.fullName) || (0, payroll_1.tokensSubset)(d.fullName, tsName));
        emps.set(key, {
            username: key,
            fname: r.fname,
            lname: r.lname,
            group: r.group,
            fullName: dirMatch?.fullName ?? tsName,
            inactive: dirMatch ? !dirMatch.active : false,
        });
    }
    const empJobs = new Map(); // username → jobs in first-encounter order
    const jobRowById = new Map(pbj.customers.flatMap((c) => c.jobs.map((j) => [j.id, j])));
    enriched.forEach((r, idx) => {
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
            const pbjJob = jobRowById.get(jobId);
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
            if (o !== undefined)
                j.hours = o;
        }
    }
    // ----- crew leads this week ----------------------------------------------
    // Names resolved on piece-rate NV pivot jobs + sub-only calendar jobs.
    const leadNames = new Map(); // UPPER full name
    const pieceRateJobs = pbj.customers.flatMap((c) => c.jobs
        .filter((j) => !j.analysis.hourly && j.analysis.apKey !== 'OFFICE')
        .map((j) => ({ job: j, customer: c.customer })));
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
    for (const r of reportRows) {
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
        calOnlyJobs.push({
            row: r,
            lead: res.names,
            sub: res.sub,
            split: res.split,
            woOverride: woSrc?.amount ?? null,
            woPct: pctSrc?.pct ??
                (0, payroll_1.profitSharePctFor)(res.names[0], settings.profitShare, settings.defaultProfitSharePct),
        });
        for (const n of res.names) {
            const up = n.toUpperCase();
            if (!leadNames.has(up))
                leadNames.set(up, { sub: res.sub });
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
    // ----- hourly block construction ------------------------------------------
    const crewSuffix = (group) => {
        const m = /^(.*?)['\u2019]s\s+crew\s*$/i.exec(group.trim());
        return m ? ` (${m[1]}'s Crew)` : group ? ` (${group})` : '';
    };
    const buildHourlyBlock = (emp, officeStyle) => {
        const jobs = empJobs.get(emp.username) ?? [];
        // §3.4 overtime on the QB-adjusted hours.
        const ot = (0, payroll_1.splitOvertime)(jobs.map((j) => ({ jobId: j.jobId, hours: j.hours, lastDate: j.lastDate, lastRowIndex: j.lastRowIndex })));
        const lines = [];
        const pushLine = (j, hours, isOt) => {
            const lineId = `${emp.username}\u0000${j.jobId}${isOt ? ':OT' : ''}`;
            const e = edits[`hl:${lineId}`] ?? {};
            const rateInfo = (0, payroll_1.wageRateFor)(emp.fullName, settings.wageRates, settings.defaultWageRate, j.hourly);
            const baseRate = e.rate !== undefined ? e.rate : rateInfo.rate;
            const rate = isOt ? (0, format_1.round2)(baseRate * 1.5) : baseRate;
            const label = officeStyle
                ? `Office${isOt ? ' [OT]' : ''}`
                : `${(0, payroll_1.jobDisplayName)(j.customer, j.jobcode2)}${j.hourly ? ' [hrly]' : ''}${isOt ? ' [OT]' : ''}`;
            lines.push({
                id: lineId,
                apKey: j.apKey,
                jobLabel: label,
                hours: e.hours !== undefined ? e.hours : (0, format_1.round2)(hours),
                rate,
                rateDefaulted: rateInfo.defaulted && e.rate === undefined,
                ot: isOt,
                hourlyJob: j.hourly,
                colorHex: j.colorHex,
            });
        };
        if (officeStyle && jobs.length > 0) {
            // §5.3 — single "Office" line covering the week (+ OT).
            const totalHrs = jobs.reduce((a, j) => a + j.hours, 0);
            const regular = ot.otHours > 0 ? (0, format_1.round2)(totalHrs - ot.otHours) : (0, format_1.round2)(totalHrs);
            const rep = jobs[0];
            pushLine(rep, regular, false);
            if (ot.otHours > 0)
                pushLine(rep, ot.otHours, true);
        }
        else {
            for (const j of jobs.slice().sort((a, b) => a.firstIndex - b.firstIndex)) {
                pushLine(j, ot.regularByJob.get(j.jobId) ?? j.hours, false);
                if (ot.otJobId === j.jobId && ot.otHours > 0)
                    pushLine(j, ot.otHours, true);
            }
        }
        const totalHours = (0, format_1.round2)(lines.reduce((a, l) => a + l.hours, 0));
        const totalPay = (0, format_1.round2)(lines.reduce((a, l) => a + l.hours * l.rate, 0));
        const blockEdit = edits[`hb:${emp.username}`] ?? {};
        return {
            key: emp.username,
            title: officeStyle ? emp.fullName : `${emp.fullName}${crewSuffix(emp.group)}`,
            employeeName: emp.fullName,
            lines,
            totalHours,
            totalPay,
            net: blockEdit.net !== undefined ? blockEdit.net : null,
            inactive: emp.inactive,
        };
    };
    // ----- sections ------------------------------------------------------------
    const officeEmps = [];
    const painterEmps = [];
    const leadEmps = [];
    for (const emp of emps.values()) {
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
    const plByJob = pl?.byJob ?? new Map();
    const buildPsJob = (blockKey, jobRowId, apKey, jobShort, customer, woPct, woOverride, percent, split) => {
        const e = edits[`ps:${blockKey}:${jobRowId}`] ?? {};
        const metrics = apKey ? plByJob.get(apKey) : undefined;
        const missingFromPl = !metrics;
        const income = e.income !== undefined ? e.income : (metrics?.income ?? 0);
        const materials = e.materials !== undefined ? e.materials : (metrics?.materials ?? 0);
        const laborWages = e.laborWages !== undefined ? e.laborWages : (metrics?.laborWagesSubs ?? 0);
        const mileage = e.mileage !== undefined ? e.mileage : (metrics?.mileage ?? 0);
        const halfMaterials = e.halfMaterials !== undefined ? e.halfMaterials : (0, format_1.round2)(materials / 2);
        const pct = woPct ?? 0;
        const profitShare = (0, format_1.round2)(income * pct);
        const workOrderOverride = e.workOrderOverride !== undefined ? e.workOrderOverride : (woOverride ?? profitShare);
        const percentCompletion = e.percentCompletion !== undefined ? e.percentCompletion : percent;
        const totalProfitShare = (0, format_1.round2)(workOrderOverride * percentCompletion - laborWages - mileage - halfMaterials);
        return {
            jobRowId,
            apKey,
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
            totalAfterSplit: (0, format_1.round2)(totalProfitShare / Math.max(1, split)),
            missingFromPl,
            income,
            materials,
        };
    };
    const leads = [];
    for (const [leadUpper, meta] of leadNames) {
        const blockKey = `lead:${leadUpper}`;
        const jobs = [];
        for (const { job } of pieceRateJobs) {
            const names = job.analysis.crewLeadResolution && job.analysis.crewLeadResolution.names.length > 0
                ? job.analysis.crewLeadResolution.names
                : job.analysis.crewLead
                    ? [job.analysis.crewLead.toUpperCase()]
                    : [];
            if (!names.some((n) => n.toUpperCase() === leadUpper))
                continue;
            const pctNum = typeof job.analysis.percent === 'number' ? job.analysis.percent : 0;
            jobs.push(buildPsJob(blockKey, job.id, job.analysis.apKey, job.jobcode2, job.id.split('\u0000')[0], job.analysis.woPct, job.analysis.woOverride, pctNum, job.analysis.split));
        }
        for (const cj of calOnlyJobs) {
            if (!cj.lead.some((n) => n.toUpperCase() === leadUpper))
                continue;
            const pctNum = typeof cj.row.percent === 'number' ? cj.row.percent : 0;
            jobs.push(buildPsJob(blockKey, `cal:${cj.row.id}`, cj.row.apKey, cj.row.qboJobName || cj.row.event, cj.row.qboJobName, cj.woPct, cj.woOverride, pctNum, cj.split));
        }
        if (jobs.length === 0)
            continue;
        const leadEmp = leadEmps.find((e) => (0, payroll_1.tokensSubset)(e.fullName, leadUpper) || (0, payroll_1.tokensSubset)(leadUpper, e.fullName));
        const tsheets = leadEmp && !meta.sub ? buildHourlyBlock(leadEmp, false) : null;
        const be = edits[`psb:${blockKey}`] ?? {};
        const total = (0, format_1.round2)(jobs.reduce((a, j) => a + j.totalAfterSplit, 0));
        const mileageNew = be.mileageNew !== undefined ? be.mileageNew : 0;
        const hourlyPay = tsheets ? tsheets.totalPay : 0;
        const reimbursements = be.reimbursements !== undefined ? be.reimbursements : 0;
        const deductions = be.deductions !== undefined ? be.deductions : 0;
        leads.push({
            key: blockKey,
            leadName: leadUpper,
            sub: meta.sub,
            jobs,
            total,
            mileageNew,
            hourlyPay,
            reimbursements,
            reimbMemo: be.reimbMemo ?? '',
            deductions,
            deductMemo: be.deductMemo ?? '',
            totalPay: (0, format_1.round2)(total + mileageNew + hourlyPay + reimbursements + deductions),
            tsheets,
        });
    }
    leads.sort((a, b) => byText(firstNameOf(a.leadName), firstNameOf(b.leadName)));
    // ----- salary sections ---------------------------------------------------
    const salaryBlock = (entry) => {
        const e = edits[`sal:${entry.name}`] ?? {};
        const tsEmp = [...emps.values()].find((x) => (0, payroll_1.tokensSubset)(x.fullName, entry.name) || (0, payroll_1.tokensSubset)(entry.name, x.fullName));
        const hours = tsEmp
            ? (0, format_1.round2)((empJobs.get(tsEmp.username) ?? []).reduce((a, j) => a + j.hours, 0))
            : 0;
        const miles = e.miles !== undefined ? e.miles : 0;
        return {
            key: `sal:${entry.name}`,
            name: entry.name,
            role: entry.role,
            amount: entry.amount,
            section: entry.section,
            timesheetHours: hours,
            miles,
            mileageReimb: (0, format_1.round2)(miles * settings.mileageRate),
        };
    };
    const roster = settings.salariedRoster;
    const admin = roster.filter((r) => r.section === 'ADMIN').map(salaryBlock)
        .sort((a, b) => byText(firstNameOf(a.name), firstNameOf(b.name)));
    const estimators = roster.filter((r) => r.section === 'ESTIMATORS').map(salaryBlock)
        .sort((a, b) => byText(firstNameOf(a.name), firstNameOf(b.name)));
    const managers = roster.filter((r) => r.section === 'MANAGER SALARY').map(salaryBlock)
        .sort((a, b) => byText(firstNameOf(a.name), firstNameOf(b.name)));
    // ----- §5.7 inactive employees with zero hours ---------------------------
    const namesWithHours = new Set([...emps.values()].map((e) => e.fullName.toLowerCase()));
    const inactiveZeroHours = directory
        .filter((d) => !d.active && !namesWithHours.has(d.fullName.toLowerCase()))
        .map((d) => d.fullName)
        .sort(byText);
    return { title, painters, leads, office, admin, estimators, managers, inactiveZeroHours };
}
