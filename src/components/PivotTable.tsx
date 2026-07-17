import type { CSSProperties } from 'react';
import { fmtHoursFixed, isDarkColor } from '../lib/format';
import type { PivotModel } from '../lib/types';

interface PivotTableProps {
  pivot: PivotModel;
  colorByEstimator: ReadonlyMap<string, string>;
}

const BLACK_BAND: CSSProperties = { backgroundColor: '#000', color: '#fff' };
const HEADER_GRAY: CSSProperties = { backgroundColor: '#808080', color: '#000' };
const GRAND_GRAY: CSSProperties = { backgroundColor: '#808080', color: '#fff' };
const GT_BAND: CSSProperties = { backgroundColor: '#D9D9D9' };
const OVERTIME_BG = '#FFFF99';

function bandStyle(hex: string | undefined): CSSProperties | undefined {
  if (!hex) return undefined;
  return { backgroundColor: `#${hex}`, color: isDarkColor(hex) ? '#fff' : '#000' };
}

export function PivotTable({ pivot, colorByEstimator }: PivotTableProps): JSX.Element {
  const estCols = pivot.estimators;
  const totalCols = 4 + estCols.length + 1;
  // "Hours" sits over the middle of the estimator block (no merge).
  const hoursIdx = Math.floor((estCols.length - 1) / 2);

  const hoursCells = (byEstimator: Record<string, number>, cls = 'xcell'): JSX.Element[] =>
    estCols.map((e) => (
      <td key={e} className={`${cls} text-right`} style={bandStyle(colorByEstimator.get(e))}>
        {byEstimator[e] !== undefined ? fmtHoursFixed(byEstimator[e]) : ''}
      </td>
    ));

  return (
    <table className="w-max min-w-full border-collapse bg-white text-xs">
      <thead>
        {/* Black title block (rows 1–3), mirroring the exported sheet. */}
        <tr>
          <th className="xcell2 text-left text-base font-bold" style={BLACK_BAND} colSpan={totalCols}>
            {pivot.title}
          </th>
        </tr>
        <tr>
          <th className="xcell2" style={BLACK_BAND} />
          <th className="xcell2" style={BLACK_BAND} />
          <th className="xcell2 text-left font-bold" style={BLACK_BAND}>
            Name
          </th>
          <th className="xcell2 text-left font-bold" style={BLACK_BAND}>
            Job
          </th>
          {estCols.map((e, i) => (
            <th key={e} className="xcell2 text-center font-bold" style={BLACK_BAND}>
              {i === hoursIdx ? 'Hours' : ''}
            </th>
          ))}
          <th className="xcell2" style={BLACK_BAND} />
        </tr>
        <tr>
          <th className="xcell2 text-left" style={BLACK_BAND}>
            Sum of hours
          </th>
          <th className="xcell2" style={BLACK_BAND} />
          <th className="xcell2" style={BLACK_BAND} />
          <th className="xcell2" style={BLACK_BAND} />
          <th className="xcell2 text-left" style={BLACK_BAND} colSpan={estCols.length + 1}>
            Column Labels
          </th>
        </tr>
        {/* Row 4: gray labels + palette-colored estimator headers. */}
        <tr>
          <th className="xcell2 sticky top-0 z-10 min-w-[220px] text-left" style={HEADER_GRAY}>
            Row Labels
          </th>
          <th className="xcell2 sticky top-0 z-10 min-w-[90px] text-left" style={HEADER_GRAY}>
            fname
          </th>
          <th className="xcell2 sticky top-0 z-10 min-w-[160px] text-left" style={HEADER_GRAY}>
            lname
          </th>
          <th className="xcell2 sticky top-0 z-10 min-w-[260px] text-left" style={HEADER_GRAY}>
            jobcode_2
          </th>
          {estCols.map((e) => (
            <th
              key={e}
              className="xcell2 sticky top-0 z-10 min-w-[64px] text-right"
              style={bandStyle(colorByEstimator.get(e)) ?? HEADER_GRAY}
            >
              {e}
            </th>
          ))}
          <th className="xcell2 sticky top-0 z-10 min-w-[80px] text-right" style={HEADER_GRAY}>
            Grand Total
          </th>
        </tr>
      </thead>
      <tbody>
        {pivot.crews.map((crew) => (
          <CrewRows
            key={crew.crew}
            crew={crew}
            estCols={estCols}
            totalCols={totalCols}
            colorByEstimator={colorByEstimator}
            hoursCells={hoursCells}
          />
        ))}
        <tr>
          <td className="xcell font-bold" style={GRAND_GRAY}>
            Grand Total
          </td>
          <td className="xcell" />
          <td className="xcell" />
          <td className="xcell" />
          {estCols.map((e) => (
            <td key={e} className="xcell text-right font-bold" style={GRAND_GRAY}>
              {pivot.grandByEstimator[e] !== undefined ? fmtHoursFixed(pivot.grandByEstimator[e]) : ''}
            </td>
          ))}
          <td className="xcell text-right font-bold" style={GRAND_GRAY}>
            {fmtHoursFixed(pivot.grandTotal)}
          </td>
        </tr>
      </tbody>
    </table>
  );
}

function CrewRows({
  crew,
  estCols,
  totalCols,
  colorByEstimator,
  hoursCells,
}: {
  crew: PivotModel['crews'][number];
  estCols: string[];
  totalCols: number;
  colorByEstimator: ReadonlyMap<string, string>;
  hoursCells: (byEstimator: Record<string, number>, cls?: string) => JSX.Element[];
}): JSX.Element {
  return (
    <>
      {/* Crew-lead row: full-width black band, overrides the column bands. */}
      <tr>
        <td className="xcell" style={BLACK_BAND} colSpan={totalCols}>
          {crew.crew}
        </td>
      </tr>
      {crew.employees.map((emp) => (
        <EmployeeRows
          key={`${crew.crew}-${emp.username}-${emp.lname}`}
          emp={emp}
          estCols={estCols}
          colorByEstimator={colorByEstimator}
          hoursCells={hoursCells}
        />
      ))}
    </>
  );
}

function EmployeeRows({
  emp,
  estCols,
  colorByEstimator,
  hoursCells,
}: {
  emp: PivotModel['crews'][number]['employees'][number];
  estCols: string[];
  colorByEstimator: ReadonlyMap<string, string>;
  hoursCells: (byEstimator: Record<string, number>, cls?: string) => JSX.Element[];
}): JSX.Element {
  const emptyBand = estCols.map((e) => (
    <td key={e} className="xcell" style={bandStyle(colorByEstimator.get(e))} />
  ));
  // "X Total" rows close the employee block with a thin bottom border.
  const totalCls = 'xcell border-b-2 border-b-neutral-400 font-bold';
  return (
    <>
      <tr>
        <td className="xcell font-bold">{emp.username}</td>
        <td className="xcell font-bold">{emp.fname}</td>
        <td className="xcell font-bold">{emp.lname}</td>
        <td className="xcell" />
        {emptyBand}
        <td className="xcell" style={GT_BAND} />
      </tr>
      {emp.jobs.map((job) => (
        <tr key={`${job.customer}\u0000${job.job}`}>
          <td className="xcell" />
          <td className="xcell" />
          <td className="xcell font-bold">{job.customer}</td>
          <td className="xcell">{job.job}</td>
          {hoursCells(job.byEstimator)}
          <td className="xcell text-right" style={GT_BAND}>
            {fmtHoursFixed(job.total)}
          </td>
        </tr>
      ))}
      <tr>
        <td className={totalCls} />
        <td className={totalCls}>{emp.fname} Total</td>
        <td className={totalCls} />
        <td className={totalCls} />
        {hoursCells(emp.byEstimator, totalCls)}
        <td
          className={`${totalCls} text-right`}
          // On-screen review aid only (not exported): flag >40h weeks.
          style={emp.overtime ? { ...GT_BAND, backgroundColor: OVERTIME_BG } : GT_BAND}
          title={emp.overtime ? 'Overtime: grand total exceeds 40 hours (on-screen flag only)' : undefined}
        >
          {fmtHoursFixed(emp.total)}
        </td>
      </tr>
    </>
  );
}
