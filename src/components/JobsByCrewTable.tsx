import { fmtHoursFixed } from '../lib/format';
import type { JobsByCrewModel } from '../lib/jobsByCrew';

interface JobsByCrewTableProps {
  model: JobsByCrewModel;
}

/** Round 6 §3.6 — Jobs by Crew pivot (read-only). */
export function JobsByCrewTable({ model }: JobsByCrewTableProps): JSX.Element {
  return (
    <div className="p-2">
      <table className="w-max min-w-[480px] border-collapse bg-white text-xs">
        <thead>
          <tr>
            <th className="xcell2 min-w-[12px]" />
            <th className="xcell2 min-w-[220px] text-left font-bold">Row Labels</th>
            <th className="xcell2 min-w-[160px] text-left font-bold">jobcode_1</th>
            <th className="xcell2 min-w-[260px] text-left font-bold">jobcode_2</th>
            <th className="xcell2 min-w-[90px] text-right font-bold">Sum of hours</th>
          </tr>
        </thead>
        <tbody>
          {model.groups.map((g) =>
            g.jobs.map((j, i) => (
              <tr key={`${g.crew}\u0000${j.customer}\u0000${j.job}`}>
                <td className="xcell" />
                <td className="xcell font-bold">{i === 0 ? g.crew : ''}</td>
                <td className="xcell">{j.showCustomer ? j.customer : ''}</td>
                <td className="xcell">{j.job}</td>
                <td className="xcell text-right">{fmtHoursFixed(j.hours)}</td>
              </tr>
            )),
          )}
          <tr>
            <td className="xcell" />
            <td className="xcell font-bold">Grand Total</td>
            <td className="xcell" />
            <td className="xcell" />
            <td className="xcell text-right font-bold">{fmtHoursFixed(model.grandTotal)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}
