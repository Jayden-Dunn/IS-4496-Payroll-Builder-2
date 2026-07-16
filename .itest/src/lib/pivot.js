"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildPivot = buildPivot;
const format_1 = require("./format");
const settings_1 = require("./settings");
/**
 * Static replica of the manual Excel PivotTable: crews → employees → jobs,
 * hours bucketed into one column per estimator.
 *
 * Ordering is deterministic at every level: crews A→Z, employees within a
 * crew A→Z by username, job rows A→Z by (customer, then jobcode_2).
 * Estimator columns follow the persisted pivot-column-order setting; any
 * estimator not in the stored list is appended, and `#N/A` (when unresolved
 * hours exist) always renders last before Grand Total.
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
    // Columns: only estimators with hours, in the stored setting order (extras
    // appended); #N/A always last when present.
    const present = [...grand.keys()].filter((e) => (0, format_1.round2)(grand.get(e) ?? 0) !== 0);
    const hasNA = present.includes('#N/A');
    const estimators = [
        ...(0, settings_1.mergePivotColumnOrder)(pivotColumnOrder, present),
        ...(hasNA ? ['#N/A'] : []),
    ];
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
        const employees = [...crew.employees.values()]
            .sort((a, b) => byText(a.username || `${a.fname} ${a.lname}`, b.username || `${b.fname} ${b.lname}`))
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
