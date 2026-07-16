import { fmtHoursFixed } from '../lib/format';
import type { PivotModel } from '../lib/types';

interface PivotTableProps {
  pivot: PivotModel;
}

const OVERTIME_BG = '#FFFF99';

export function PivotTable({ pivot }: PivotTableProps): JSX.Element {
  const estCols = pivot.estimators;
  const totalCols = 4 + estCols.length + 1;

  const hoursCells = (byEstimator: Record<string, number>, bold: boolean): JSX.Element[] =>
    estCols.map((e) => (
      <td key={e} className={`xcell text-right ${bold ? 'font-semibold' : ''}`}>
        {byEstimator[e] !== undefined ? fmtHoursFixed(byEstimator[e]) : ''}
      </td>
    ));

  return (
    <table className="w-max min-w-full border-collapse bg-white text-xs">
      <thead>
        {/* Four-row header block mirroring the exported sheet. */}
        <tr>
          <th className="phead text-left text-sm" colSpan={totalCols}>
            {pivot.title}
          </th>
        </tr>
        <tr>
          <th className="phead" />
          <th className="phead" />
          <th className="phead text-left">Name</th>
          <th className="phead text-left">Job</th>
          <th className="phead text-center" colSpan={estCols.length + 1}>
            Hours
          </th>
        </tr>
        <tr>
          <th className="phead text-left">Sum of hours</th>
          <th className="phead" />
          <th className="phead" />
          <th className="phead" />
          <th className="phead text-left" colSpan={estCols.length + 1}>
            Column Labels
          </th>
        </tr>
        <tr>
          <th className="phead sticky top-0 z-10 min-w-[220px] text-left">Row Labels</th>
          <th className="phead sticky top-0 z-10 min-w-[90px] text-left">fname</th>
          <th className="phead sticky top-0 z-10 min-w-[160px] text-left">lname</th>
          <th className="phead sticky top-0 z-10 min-w-[260px] text-left">jobcode_2</th>
          {estCols.map((e) => (
            <th key={e} className="phead sticky top-0 z-10 min-w-[64px] text-right">
              {e}
            </th>
          ))}
          <th className="phead sticky top-0 z-10 min-w-[80px] text-right">Grand Total</th>
        </tr>
      </thead>
      <tbody>
        {pivot.crews.map((crew) => (
          <CrewRows key={crew.crew} crew={crew} estCols={estCols} totalCols={totalCols} hoursCells={hoursCells} />
        ))}
        <tr>
          <td className="xcell font-bold" colSpan={4}>
            Grand Total
          </td>
          {hoursCells(pivot.grandByEstimator, true)}
          <td className="xcell text-right font-bold">{fmtHoursFixed(pivot.grandTotal)}</td>
        </tr>
      </tbody>
    </table>
  );
}

function CrewRows({
  crew,
  estCols,
  totalCols,
  hoursCells,
}: {
  crew: PivotModel['crews'][number];
  estCols: string[];
  totalCols: number;
  hoursCells: (byEstimator: Record<string, number>, bold: boolean) => JSX.Element[];
}): JSX.Element {
  return (
    <>
      <tr>
        <td className="xcell font-bold" style={{ backgroundColor: '#DCE6F1' }} colSpan={totalCols}>
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
          <td className="xcell text-right">{fmtHoursFixed(job.total)}</td>
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
          {fmtHoursFixed(emp.total)}
        </td>
      </tr>
    </>
  );
}
