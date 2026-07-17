"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.orderLeadFirst = orderLeadFirst;
exports.buildPivot = buildPivot;
const format_1 = require("./format");
// ---------------------------------------------------------------------------
// Crew-lead detection (dynamic — derived from the group label, never names)
// ---------------------------------------------------------------------------
/** Case-insensitive name tokens; hyphens split ("Martinez-Rincon" → 2 tokens). */
function nameTokens(s) {
    return s.toLowerCase().split(/[\s-]+/).filter(Boolean);
}
/**
 * Order employees so the crew lead comes first, then the rest in their
 * existing (alphabetical) order. The lead's name is derived by stripping the
 * trailing "'s Crew" / "’s Crew" from the group label; an employee is the
 * lead when every label token appears in their fname+lname token set (this
 * tolerates middle/extra surnames, e.g. "Daniel Bonilla" → Daniel Gonzales
 * Bonilla). Ties prefer the exact fname+lname match, then the fewest extra
 * name tokens, then the existing order. Groups without a "'s Crew" suffix or
 * with no matching employee are left untouched.
 */
function orderLeadFirst(crewLabel, employees) {
    const m = /^(.*?)['\u2019]s\s+crew\s*$/i.exec(crewLabel.trim());
    if (!m)
        return employees;
    const labelTokens = nameTokens(m[1]);
    if (labelTokens.length === 0)
        return employees;
    const empTokens = (e) => nameTokens(`${e.fname} ${e.lname}`);
    const candidates = employees.filter((e) => {
        const set = new Set(empTokens(e));
        return labelTokens.every((t) => set.has(t));
    });
    if (candidates.length === 0)
        return employees;
    const label = labelTokens.join(' ');
    const lead = [...candidates].sort((a, b) => {
        const exactA = empTokens(a).join(' ') === label ? 0 : 1;
        const exactB = empTokens(b).join(' ') === label ? 0 : 1;
        if (exactA !== exactB)
            return exactA - exactB;
        const extraA = empTokens(a).length - labelTokens.length;
        const extraB = empTokens(b).length - labelTokens.length;
        if (extraA !== extraB)
            return extraA - extraB;
        return 0; // stable sort keeps the existing (alphabetical) order
    })[0];
    return [lead, ...employees.filter((e) => e !== lead)];
}
/**
 * Static replica of the manual Excel PivotTable: crews → employees → jobs,
 * hours bucketed into one column per estimator.
 *
 * Ordering is deterministic at every level: crews A→Z; employees within a
 * crew crew-lead first, then A→Z by username; job rows A→Z by (customer,
 * then jobcode_2). The configured estimator columns always render (in the
 * persisted setting order); extras found in the data append after, and
 * `#N/A` (only when unresolved hours exist) renders last before Grand Total.
 */
function buildPivot(rows, pivotColumnOrder, payDate) {
    const crews = new Map();
    const grand = new Map();
    let grandTotal = 0;
    for (const r of rows) {
        if (r.hours === 0)
            continue;
        const crewName = r.group || '(no crew)';
        let crew = crews.get(crewName);
        if (!crew) {
            crew = { crew: crewName, employees: new Map() };
            crews.set(crewName, crew);
        }
        const empKey = r.username || `${r.fname} ${r.lname}`;
        let emp = crew.employees.get(empKey);
        if (!emp) {
            emp = { username: r.username, fname: r.fname, lname: r.lname, jobs: new Map(), byEstimator: new Map(), total: 0 };
            crew.employees.set(empKey, emp);
        }
        const jobKey = `${r.jobcode1}\u0000${r.jobcode2}`;
        let job = emp.jobs.get(jobKey);
        if (!job) {
            job = { customer: r.jobcode1, job: r.jobcode2, byEstimator: new Map(), total: 0 };
            emp.jobs.set(jobKey, job);
        }
        const est = r.estimator;
        job.byEstimator.set(est, (job.byEstimator.get(est) ?? 0) + r.hours);
        job.total += r.hours;
        emp.byEstimator.set(est, (emp.byEstimator.get(est) ?? 0) + r.hours);
        emp.total += r.hours;
        grand.set(est, (grand.get(est) ?? 0) + r.hours);
        grandTotal += r.hours;
    }
    // Columns: the configured estimator set ALWAYS renders (even with zero
    // hours this period); estimators discovered in the data outside the set
    // append after it; #N/A renders only when unresolved hours exist, last.
    const present = [...grand.keys()].filter((e) => (0, format_1.round2)(grand.get(e) ?? 0) !== 0);
    const hasNA = present.includes('#N/A');
    const configured = pivotColumnOrder.filter((e) => e !== '#N/A');
    const extras = present.filter((e) => e !== '#N/A' && !configured.includes(e));
    const estimators = [...configured, ...extras, ...(hasNA ? ['#N/A'] : [])];
    const toRecord = (m) => {
        const rec = {};
        for (const e of estimators) {
            const v = m.get(e);
            if (v !== undefined && (0, format_1.round2)(v) !== 0)
                rec[e] = (0, format_1.round2)(v);
        }
        return rec;
    };
    const byText = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' });
    const crewModels = [...crews.values()]
        .sort((a, b) => byText(a.crew, b.crew))
        .map((crew) => {
        const sortedEmps = [...crew.employees.values()].sort((a, b) => byText(a.username || `${a.fname} ${a.lname}`, b.username || `${b.fname} ${b.lname}`));
        const employees = orderLeadFirst(crew.crew, sortedEmps)
            .map((emp) => {
            const jobs = [...emp.jobs.values()]
                .sort((a, b) => byText(a.customer, b.customer) || byText(a.job, b.job))
                .map((job) => ({
                customer: job.customer,
                job: job.job,
                byEstimator: toRecord(job.byEstimator),
                total: (0, format_1.round2)(job.total),
            }));
            const total = (0, format_1.round2)(emp.total);
            return {
                username: emp.username,
                fname: emp.fname,
                lname: emp.lname,
                jobs,
                byEstimator: toRecord(emp.byEstimator),
                total,
                overtime: total > 40,
            };
        });
        return { crew: crew.crew, employees };
    });
    return {
        title: `NV PAYROLL: ${payDate ? (0, format_1.fmtDate)(payDate) : ''}`.trim(),
        estimators,
        crews: crewModels,
        grandByEstimator: toRecord(grand),
        grandTotal: (0, format_1.round2)(grandTotal),
    };
}
