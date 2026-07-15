import { fmtHours } from '../lib/format';
import type { PivotModel } from '../lib/types';

interface PivotTableProps {
  pivot: PivotModel;
}

const OVERTIME_BG = '#FFFF99';

export function PivotTable({ pivot }: PivotTableProps): JSX.Element {
  const estCols = pivot.estimators;

  const hoursCells = (byEstimator: Record<string, number>, bold: boolean): JSX.Element[] => [
    ...estCols.map((e) => (
      <td key={e} className={`xcell text-right ${bold ? 'font-semibold' : ''}`}>
        {byEstimator[e] !== undefined ? fmtHours(byEstimator[e]) : ''}
      </td>
    )),
  ];

  return (
    <table className="w-max min-w-full border-collapse bg-white text-xs">
      <thead>
        <tr>
          <th className="xhead sticky top-0 z-10 min-w-[220px]" colSpan={4 + estCols.length + 1}>
            {pivot.title}
          </th>
        </tr>
        <tr>
          <th className="xhead sticky top-[26px] z-10 min-w-[220px]">Row Labels</th>
          <th className="xhead sticky top-[26px] z-10 min-w-[90px]">fname</th>
          <th className="xhead sticky top-[26px] z-10 min-w-[160px]">lname</th>
          <th className="xhead sticky top-[26px] z-10 min-w-[260px]">Job</th>
          {estCols.map((e) => (
            <th key={e} className="xhead sticky top-[26px] z-10 min-w-[64px] text-right">
              {e}
            </th>
          ))}
          <th className="xhead sticky top-[26px] z-10 min-w-[80px] text-right">Grand Total</th>
        </tr>
      </thead>
      <tbody>
        {pivot.crews.map((crew) => (
          <CrewRows key={crew.crew} crew={crew} estCols={estCols} hoursCells={hoursCells} />
        ))}
        <tr>
          <td className="xcell font-bold" colSpan={4}>
            Grand Total
          </td>
          {hoursCells(pivot.grandByEstimator, true)}
          <td className="xcell text-right font-bold">{fmtHours(pivot.grandTotal)}</td>
        </tr>
      </tbody>
    </table>
  );
}

function CrewRows({
  crew,
  estCols,
  hoursCells,
}: {
  crew: PivotModel['crews'][number];
  estCols: string[];
  hoursCells: (byEstimator: Record<string, number>, bold: boolean) => JSX.Element[];
}): JSX.Element {
  return (
    <>
      <tr className="bg-neutral-200">
        <td className="xcell font-bold" colSpan={4 + estCols.length + 1}>
          {crew.crew}
        </td>
      </tr>
      {crew.employees.map((emp) => (
        <EmployeeRows key={`${crew.crew}-${emp.username}-${emp.lname}`} emp={emp} estCols={estCols} hoursCells={hoursCells} />
      ))}
    </>
  );
}

function EmployeeRows({
  emp,
  estCols,
  hoursCells,
}: {
  emp: PivotModel['crews'][number]['employees'][number];
  estCols: string[];
  hoursCells: (byEstimator: Record<string, number>, bold: boolean) => JSX.Element[];
}): JSX.Element {
  return (
    <>
      <tr>
        <td className="xcell">{emp.username}</td>
        <td className="xcell">{emp.fname}</td>
        <td className="xcell">{emp.lname}</td>
        <td className="xcell" />
        {estCols.map((e) => (
          <td key={e} className="xcell" />
        ))}
        <td className="xcell" />
      </tr>
      {emp.jobs.map((job) => (
        <tr key={`${job.customer}\u0000${job.job}`}>
          <td className="xcell" />
          <td className="xcell" />
          <td className="xcell">{job.customer}</td>
          <td className="xcell">{job.job}</td>
          {hoursCells(job.byEstimator, false)}
          <td className="xcell text-right">{fmtHours(job.total)}</td>
        </tr>
      ))}
      <tr>
        <td className="xcell" />
        <td className="xcell font-semibold">{emp.fname} Total</td>
        <td className="xcell" />
        <td className="xcell" />
        {hoursCells(emp.byEstimator, true)}
        <td
          className="xcell text-right font-semibold"
          style={emp.overtime ? { backgroundColor: OVERTIME_BG } : undefined}
          title={emp.overtime ? 'Overtime: grand total exceeds 40 hours' : undefined}
        >
          {fmtHours(emp.total)}
        </td>
      </tr>
    </>
  );
}
