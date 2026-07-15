import { fmtDate, round2 } from './format';
import type { EnrichedTimesheetRow, PivotCrew, PivotEmployee, PivotJobRow, PivotModel } from './types';

interface JobAcc {
  customer: string;
  job: string;
  byEstimator: Map<string, number>;
  total: number;
}

interface EmpAcc {
  username: string;
  fname: string;
  lname: string;
  jobs: Map<string, JobAcc>;
  jobOrder: string[];
  byEstimator: Map<string, number>;
  total: number;
}

interface CrewAcc {
  crew: string;
  employees: Map<string, EmpAcc>;
  empOrder: string[];
}

/**
 * Static replica of the manual Excel PivotTable: crews → employees → jobs,
 * hours bucketed into one column per estimator. Crews and employees appear in
 * first-encounter (file) order; estimator columns follow the Settings legend
 * order, with any extras (e.g. #N/A) appended.
 */
export function buildPivot(
  rows: EnrichedTimesheetRow[],
  estimatorOrder: string[],
  payDate: Date | null,
): PivotModel {
  const crews = new Map<string, CrewAcc>();
  const crewOrder: string[] = [];
  const grand = new Map<string, number>();
  let grandTotal = 0;

  for (const r of rows) {
    if (r.hours === 0) continue;
    const crewName = r.group || '(no crew)';
    let crew = crews.get(crewName);
    if (!crew) {
      crew = { crew: crewName, employees: new Map(), empOrder: [] };
      crews.set(crewName, crew);
      crewOrder.push(crewName);
    }
    const empKey = r.username || `${r.fname} ${r.lname}`;
    let emp = crew.employees.get(empKey);
    if (!emp) {
      emp = { username: r.username, fname: r.fname, lname: r.lname, jobs: new Map(), jobOrder: [], byEstimator: new Map(), total: 0 };
      crew.employees.set(empKey, emp);
      crew.empOrder.push(empKey);
    }
    const jobKey = `${r.jobcode1}\u0000${r.jobcode2}`;
    let job = emp.jobs.get(jobKey);
    if (!job) {
      job = { customer: r.jobcode1, job: r.jobcode2, byEstimator: new Map(), total: 0 };
      emp.jobs.set(jobKey, job);
      emp.jobOrder.push(jobKey);
    }
    const est = r.estimator;
    job.byEstimator.set(est, (job.byEstimator.get(est) ?? 0) + r.hours);
    job.total += r.hours;
    emp.byEstimator.set(est, (emp.byEstimator.get(est) ?? 0) + r.hours);
    emp.total += r.hours;
    grand.set(est, (grand.get(est) ?? 0) + r.hours);
    grandTotal += r.hours;
  }

  // Estimator columns: only those with hours, in legend order, extras last.
  const present = [...grand.keys()].filter((e) => round2(grand.get(e) ?? 0) !== 0);
  const estimators = [
    ...estimatorOrder.filter((e) => present.includes(e)),
    ...present.filter((e) => !estimatorOrder.includes(e)),
  ];

  const toRecord = (m: Map<string, number>): Record<string, number> => {
    const rec: Record<string, number> = {};
    for (const e of estimators) {
      const v = m.get(e);
      if (v !== undefined && round2(v) !== 0) rec[e] = round2(v);
    }
    return rec;
  };

  const crewModels: PivotCrew[] = crewOrder.map((crewName) => {
    const crew = crews.get(crewName) as CrewAcc;
    const employees: PivotEmployee[] = crew.empOrder.map((empKey) => {
      const emp = crew.employees.get(empKey) as EmpAcc;
      const jobs: PivotJobRow[] = emp.jobOrder.map((jobKey) => {
        const job = emp.jobs.get(jobKey) as JobAcc;
        return {
          customer: job.customer,
          job: job.job,
          byEstimator: toRecord(job.byEstimator),
          total: round2(job.total),
        };
      });
      const total = round2(emp.total);
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
    return { crew: crewName, employees };
  });

  return {
    title: `NV PAYROLL: ${payDate ? fmtDate(payDate) : ''}`.trim(),
    estimators,
    crews: crewModels,
    grandByEstimator: toRecord(grand),
    grandTotal: round2(grandTotal),
  };
}
