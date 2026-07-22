"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildPayrollExceptions = buildPayrollExceptions;
/**
 * Round 3 §6 — "Payroll Exceptions": the standing red-flag rules from the
 * payroll process doc plus this round's checks. Same grouped/linked shape as
 * the %'s Exceptions panel.
 */
const format_1 = require("./format");
const jobNames_1 = require("./jobNames");
const PS_THRESHOLD_DEFAULT = 25;
function buildPayrollExceptions(input) {
    const { pbj, preview, reportRows, pl } = input;
    const groups = [];
    const jobs = pbj.customers.flatMap((c) => c.jobs.map((j) => ({ job: j, customer: c.customer })));
    const included = reportRows.filter((r) => r.included);
    // 1. Crew-lead resolution problems.
    {
        const items = [];
        for (const { job } of jobs) {
            if (job.analysis.apKey === 'OFFICE' || job.analysis.hourly)
                continue;
            const res = job.analysis.crewLeadResolution;
            if (!res) {
                if (job.analysis.onCalendar !== 'NOT ON CALENDAR' && !job.analysis.crewLead) {
                    items.push({ message: `${job.jobcode2} — no crew-lead parenthetical on the calendar event` });
                }
                continue;
            }
            if (!res.raw && !job.analysis.crewLead) {
                items.push({ message: `${job.jobcode2} — the event title names no crew lead` });
                continue;
            }
            for (const f of res.flags) {
                const label = f.kind === 'noMatch'
                    ? `no matching employee or subcontractor for "${f.candidate}"`
                    : f.kind === 'ambiguous'
                        ? `multiple equally-preferred matches for "${f.candidate}"`
                        : f.kind === 'notACrewLead'
                            ? `"${f.candidate}" matched an employee who is not a known crew lead`
                            : `only match for "${f.candidate}" is marked inactive (*) — used, verify`;
                items.push({ message: `${job.jobcode2} — (${res.raw}): ${label}` });
            }
        }
        groups.push({
            key: 'leadResolution',
            title: 'Crew-lead resolution',
            note: input.hasDirectory
                ? 'Raw parenthetical shown; fix the Crew Lead cell on Pivot by Job.'
                : 'Employee Name List not saved in Settings — resolution used timesheet names only.',
            items,
        });
    }
    // 2. Hours on a job not on the calendar (piece-rate only).
    groups.push({
        key: 'hoursNotOnCalendar',
        title: 'Hours on a job not on the calendar',
        note: 'Red flag: jobs with hours not on the entry (calendar). Verify the hours are on the right job.',
        items: jobs
            .filter(({ job }) => job.analysis.onCalendar === 'NOT ON CALENDAR' && !job.analysis.hourly && job.analysis.apKey !== 'OFFICE')
            .map(({ job }) => ({
            message: `${job.jobcode2} — ${(0, format_1.fmtHours)(job.total)} h, AP ${job.analysis.apKey ?? '—'} matches no included TimeNavi event`,
        })),
    });
    // 3. Calendar job with no hours (piece-rate, not subbed).
    {
        const timesheetAps = new Set(jobs.map(({ job }) => job.analysis.apKey).filter(Boolean));
        const subAps = new Set();
        for (const b of preview.leads) {
            if (!b.sub)
                continue;
            for (const j of b.jobs)
                if (j.apKey)
                    subAps.add(j.apKey);
        }
        groups.push({
            key: 'calendarNoHours',
            title: 'Calendar job with no hours',
            note: 'Red flag: job on entry with no hours. Subcontracted jobs are excluded (they never clock in).',
            items: included
                .filter((r) => r.apKey &&
                !timesheetAps.has(r.apKey) &&
                !subAps.has(r.apKey) &&
                r.percent !== 'HOURLY' &&
                !/^utah\b|^reno\b/i.test(r.calendar.trim()))
                .map((r) => ({ message: `AP ${r.apKey} — "${r.event}" has no timesheet hours`, rowId: r.id })),
        });
    }
    // 4. Hours 40+ / under 10.
    {
        const items = [];
        const everyBlock = [...preview.painters, ...preview.office, ...preview.leads.flatMap((l) => (l.tsheets ? [l.tsheets] : []))];
        for (const b of everyBlock) {
            const otLine = b.lines.find((l) => l.ot);
            if (otLine) {
                items.push({
                    message: `${b.employeeName} — ${(0, format_1.fmtHours)(b.totalHours)} h total: ${(0, format_1.fmtHours)(otLine.hours)} h OT at 1.5× on "${otLine.jobLabel.replace(' [OT]', '')}"`,
                });
            }
        }
        for (const b of preview.painters) {
            if (!b.inactive && b.totalHours > 0 && b.totalHours < 10) {
                items.push({ message: `${b.employeeName} — only ${(0, format_1.fmtHours)(b.totalHours)} h this week (under 10)` });
            }
        }
        groups.push({
            key: 'hours40or10',
            title: 'Hours 40+ / under 10',
            note: 'Overtime is split automatically (1.5× on the last job worked); under-10 painters may be missing punches.',
            items,
        });
    }
    // 5. Missing pay inputs.
    {
        const items = [];
        for (const { job } of jobs) {
            if (job.analysis.hourly || job.analysis.apKey === 'OFFICE')
                continue;
            if (job.analysis.woPct === null) {
                items.push({ message: `${job.jobcode2} — no WO% (hourly? set '-' or a % on Pivot by Job)` });
            }
            if (job.analysis.woOverride === null && job.analysis.apKey && input.pl && !input.pl.byJob.get(job.analysis.apKey)?.income) {
                items.push({ message: `${job.jobcode2} — no WO Override and no P&L income (the "????" case)` });
            }
            if (typeof job.analysis.percent !== 'number') {
                items.push({ message: `${job.jobcode2} — piece-rate job with no % of Completion on the TimeNavi grid` });
            }
        }
        const everyBlock = [...preview.painters, ...preview.office, ...preview.leads.flatMap((l) => (l.tsheets ? [l.tsheets] : []))];
        for (const b of everyBlock) {
            if (b.lines.some((l) => l.rateDefaulted)) {
                items.push({ message: `${b.employeeName} — no configured wage rate (default $ applied); add them in Settings` });
            }
        }
        groups.push({
            key: 'missingInputs',
            title: 'Missing pay inputs',
            note: 'Piece-rate jobs missing WO% / WO Override + Income, and employees with no configured wage rate.',
            items,
        });
    }
    // 5b (Round 4/6 §7.1). Davis-Bacon job paid below the DB minimum (manual edit).
    {
        const items = [];
        const everyBlock = [...preview.painters, ...preview.office, ...preview.leads.flatMap((l) => (l.tsheets ? [l.tsheets] : []))];
        for (const b of everyBlock) {
            for (const l of b.lines) {
                if (l.belowPrevailing) {
                    items.push({
                        message: `${b.employeeName} — "${l.jobLabel}" is a Davis-Bacon job but pays ${l.rate.toFixed(2)}/hr (below the Davis-Bacon rate)`,
                    });
                }
            }
        }
        groups.push({
            key: 'prevailingWage',
            title: 'Davis-Bacon job paid below DB rate',
            note: 'The job matched a Davis-Bacon pattern in Settings but the rendered rate sits below the Davis-Bacon minimum (usually after a manual edit).',
            items,
        });
    }
    // 5c (Round 6 §7.1 Branch 2). Davis-Bacon floor applied on crew-lead pay.
    {
        const items = [];
        for (const b of preview.leads) {
            for (const j of b.jobs) {
                if (!j.dbFloorApplied)
                    continue;
                const ps = (j.dbFloorPriorPay ?? j.totalAfterSplit).toFixed(2);
                const hrs = j.dbFloorHours ?? 0;
                const rate = j.dbFloor && hrs > 0 ? (j.dbFloor / hrs).toFixed(2) : '?';
                items.push({
                    message: `Davis-Bacon floor applied: ${b.leadName} on ${j.apKey ? `AP${j.apKey}` : '—'} — profit share ${ps} < ${hrs}h × ${rate}`,
                });
            }
        }
        groups.push({
            key: 'dbFloorApplied',
            title: 'Davis-Bacon floor applied',
            note: 'Informational: crew-lead per-job pay was raised to hours × Davis-Bacon rate on a DB job.',
            items,
        });
    }
    // 5d (Round 6 §7.2). Job-rate rule applied (informational).
    {
        const items = [];
        const everyBlock = [...preview.painters, ...preview.office, ...preview.leads.flatMap((l) => (l.tsheets ? [l.tsheets] : []))];
        for (const b of everyBlock) {
            for (const l of b.lines) {
                if (!l.fromJobRule)
                    continue;
                const note = l.ruleNote ? ` (${l.ruleNote})` : '';
                items.push({
                    message: `${b.employeeName} — "${l.jobLabel}" @ $${l.rate.toFixed(2)}/hr${note}`,
                });
            }
        }
        groups.push({
            key: 'jobRateRule',
            title: 'Job-rate rule applied',
            note: 'Informational: hourly lines paid from a Settings job-specific pay rule this week.',
            items,
        });
    }
    // 5e (Round 6 §6.2). No work-order % for lead (pivot blank + no Settings crew map entry).
    {
        const items = [];
        const seen = new Set();
        for (const b of preview.leads) {
            if (b.jobs.some((j) => j.woPct === null) && !seen.has(b.leadName)) {
                seen.add(b.leadName);
                items.push({ message: `no work-order % for ${b.leadName}` });
            }
        }
        groups.push({
            key: 'noWorkOrderPct',
            title: 'No work-order %',
            note: 'Pivot WO% blank and the lead has no Work-order % entry in Settings — Profit Share uses 0 until one is set.',
            items,
        });
    }
    // 6. Job missing from P&L.
    groups.push({
        key: 'missingFromPl',
        title: 'Job missing from P&L',
        note: 'Income/Labor default to 0 — the profit share cannot be computed until the P&L includes the job (or you edit the numbers).',
        items: pl
            ? preview.leads
                .flatMap((b) => b.jobs.filter((j) => j.missingFromPl).map((j) => ({ b, j })))
                .map(({ b, j }) => ({ message: `AP ${j.apKey ?? '—'} (${b.leadName}) — no P&L column` }))
            : [],
    });
    // 7. Profit share ≤ 0 or near-zero.
    {
        const threshold = input.psThreshold ?? PS_THRESHOLD_DEFAULT;
        groups.push({
            key: 'lowProfitShare',
            title: `Profit share below $${threshold}`,
            note: 'Often means the crew went way over on the work order — consider an Addition/adjustment.',
            items: preview.leads.flatMap((b) => b.jobs
                .filter((j) => j.totalProfitShare < threshold && !j.missingFromPl)
                .map((j) => ({
                message: `${b.leadName} — AP ${j.apKey ?? '—'} Total Profit Share ${j.totalProfitShare.toFixed(2)}`,
            }))),
        });
    }
    // 8. Sundry double-count hint.
    {
        const items = [];
        if (pl) {
            for (const b of preview.leads) {
                if (b.deductions === 0 && !b.deductMemo)
                    continue;
                for (const j of b.jobs) {
                    const metrics = j.apKey ? pl.byJob.get(j.apKey) : undefined;
                    if (metrics && metrics.sundries !== 0) {
                        items.push({
                            message: `${b.leadName} — AP ${j.apKey}: P&L shows a "Sundries -" amount (${metrics.sundries.toFixed(2)}) and the lead also has a Deductions/Sundries entry — remove it from ½-materials to avoid double-counting`,
                        });
                    }
                }
            }
        }
        groups.push({
            key: 'sundryHint',
            title: 'Sundry double-count hint',
            note: 'A sundry deducted from the painter should also come out of the ½-materials line.',
            items,
        });
    }
    // 9. Inactive employee has hours.
    groups.push({
        key: 'inactiveWithHours',
        title: 'Inactive employee has hours',
        items: [...preview.painters, ...preview.office]
            .filter((b) => b.inactive && b.totalHours > 0)
            .map((b) => ({ message: `${b.employeeName} — marked inactive (*) in the directory but has ${(0, format_1.fmtHours)(b.totalHours)} h` })),
    });
    // 10. Duplicate P&L job column.
    groups.push({
        key: 'duplicatePlColumn',
        title: 'Duplicate P&L job column',
        note: 'The same AP number appeared in more than one P&L column — values were summed.',
        items: (pl?.duplicateColumns ?? []).map((ap) => ({ message: `AP ${ap} appears in multiple P&L columns` })),
    });
    // 11. Balances / UT carry-forward attention.
    {
        const items = [];
        const thisWeekAps = new Set(included.map((r) => r.apKey).filter(Boolean));
        for (const row of [...pbj.balances, ...pbj.utPayroll]) {
            if (typeof row.percent === 'number' && row.percent >= 1) {
                items.push({ message: `${row.job || row.apKey} — carried forward but % of Completion is already 100% — likely fully paid; delete the row if so` });
            }
            else if (row.apKey && !thisWeekAps.has(row.apKey) && row.id.startsWith('ut:')) {
                items.push({ message: `${row.job || row.apKey} — UT carry-forward matched nothing this week` });
            }
        }
        groups.push({
            key: 'carryForward',
            title: 'Balances / UT carry-forward attention',
            note: 'Rows carried from last week that look fully paid or matched nothing this week.',
            items,
        });
    }
    // 12 (Round 5 §4). Naming-chain informational flags.
    {
        const baseKeyItems = [];
        const unresolvedItems = [];
        const sources = {
            timesheetByAp: (0, jobNames_1.timesheetDisplayMap)(input.enriched ?? []),
            carryByAp: (0, jobNames_1.carryDisplayMap)(input.lastWeekPbj),
            salesByAp: (0, jobNames_1.salesDisplayMap)(input.nvSales, input.utSales),
            pl,
        };
        const seen = new Set();
        const consider = (apKey, eventTitle) => {
            if (!apKey || apKey === 'OFFICE' || seen.has(apKey))
                return;
            seen.add(apKey);
            const pair = (0, jobNames_1.jobDisplayPair)(apKey, { ...sources, eventTitle });
            if (pair.usedBaseKey) {
                baseKeyItems.push({
                    message: `AP ${apKey} — suffix not found in QBO; names derived from base AP# (${pair.poLabel} / ${pair.customerName || '—'})`,
                });
            }
            if (pair.unresolved) {
                unresolvedItems.push({ message: `AP ${apKey} — job name unresolved` });
            }
        };
        for (const { job } of jobs)
            consider(job.analysis.apKey);
        for (const r of [...pbj.balances, ...pbj.utPayroll])
            consider(r.apKey);
        for (const r of included) {
            if (r.apKey && !seen.has(r.apKey))
                consider(r.apKey, r.event);
        }
        for (const lead of preview.leads) {
            for (const j of lead.jobs)
                consider(j.apKey);
        }
        groups.push({
            key: 'apSuffixBaseKey',
            title: 'AP# suffix not found in QBO; names derived from base AP#',
            note: 'A suffixed calendar AP# missed the sales lookup; display names came from the base number. Edit the AP# cell if the suffix is wrong.',
            items: baseKeyItems,
        });
        groups.push({
            key: 'jobNameUnresolved',
            title: 'Job name unresolved',
            note: 'No timesheet, carry-forward, QBO, or P&L name was found — the PO shows AP{key} with a blank customer.',
            items: unresolvedItems,
        });
    }
    return groups;
}
