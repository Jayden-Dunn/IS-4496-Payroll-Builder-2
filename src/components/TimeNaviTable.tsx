import { useState } from 'react';
import { normalizeApInput } from '../lib/apkey';
import { fmtDate, fmtMoney, fmtPercent, parseMoneyInput, parsePercentInput } from '../lib/format';
import type { ReportRow, RowEdit } from '../lib/types';

interface TimeNaviTableProps {
  rows: ReportRow[];
  colorByEstimator: ReadonlyMap<string, string>;
  estimatorOptions: string[];
  onEdit: (rowId: string, patch: RowEdit) => void;
}

const HEADERS = [
  'Include', 'Event', 'Start Date', 'End Date', 'Calendar', 'Attendees', 'Location', 'Description',
  'AP#', '% Completion', 'Estimator', 'Last % Paid', 'QBO Job Name', 'In QBO Time?',
  'Job Total $', 'Apptivo $', '$ Difference', '%compl', 'Invoice Date', 'Date Difference',
];

/** Small uncontrolled-ish text cell that commits on blur / Enter. */
function EditableCell({
  display,
  onCommit,
  className,
  title,
}: {
  display: string;
  onCommit: (text: string) => void;
  className?: string;
  title?: string;
}): JSX.Element {
  const [draft, setDraft] = useState<string | null>(null);
  return (
    <input
      className={`cell-input ${className ?? ''}`}
      value={draft ?? display}
      title={title}
      onChange={(e) => setDraft(e.target.value)}
      onFocus={() => setDraft(display)}
      onBlur={() => {
        if (draft !== null && draft !== display) onCommit(draft);
        setDraft(null);
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') e.currentTarget.blur();
        if (e.key === 'Escape') {
          setDraft(null);
          e.currentTarget.blur();
        }
      }}
    />
  );
}

export function TimeNaviTable({ rows, colorByEstimator, estimatorOptions, onEdit }: TimeNaviTableProps): JSX.Element {
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(new Set());

  const toggleExpanded = (id: string): void => {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  return (
    <table className="w-max min-w-full border-collapse bg-white text-xs">
      <thead>
        <tr>
          {HEADERS.map((h) => (
            <th key={h} className="xhead sticky top-0 z-10 whitespace-nowrap">
              {h}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((r) => {
          const hex = r.included ? colorByEstimator.get(r.estimator) : undefined;
          const isExpanded = expanded.has(r.id);
          const options = estimatorOptions.includes(r.estimator)
            ? estimatorOptions
            : [...estimatorOptions, r.estimator];
          return (
            <tr
              key={r.id}
              id={`tn-${r.id}`}
              className={r.included ? '' : 'bg-neutral-100 text-neutral-400'}
              style={hex ? { backgroundColor: `#${hex}` } : undefined}
            >
              <td className="xcell text-center">
                <input
                  type="checkbox"
                  checked={r.included}
                  onChange={() => onEdit(r.id, { included: !r.included })}
                  title={r.included ? 'Exclude this event from the report' : 'Include this event in the report'}
                />
              </td>
              <td className="xcell min-w-[260px] max-w-[340px]">{r.event}</td>
              <td className="xcell whitespace-nowrap">{fmtDate(r.startDate)}</td>
              <td className="xcell whitespace-nowrap">{fmtDate(r.endDate)}</td>
              <td className="xcell whitespace-nowrap">{r.calendar}</td>
              <td className="xcell max-w-[140px] truncate">{r.attendees}</td>
              <td className="xcell max-w-[160px] truncate" title={r.location}>
                {r.location}
              </td>
              <td className="xcell min-w-[240px] max-w-[320px]">
                <div
                  className={`cursor-pointer whitespace-pre-line ${isExpanded ? '' : 'line-clamp-2'}`}
                  onClick={() => toggleExpanded(r.id)}
                  title={isExpanded ? 'Click to collapse' : 'Click to expand'}
                >
                  {r.description}
                </div>
              </td>
              <td className="xcell w-20 font-mono">
                <EditableCell
                  display={r.apKey ?? ''}
                  onCommit={(text) => onEdit(r.id, { apKey: normalizeApInput(text) })}
                  title="AP# — edit to fix a wrong calendar AP#; all lookups re-run"
                />
              </td>
              <td className="xcell w-20 text-right">
                <EditableCell
                  display={fmtPercent(r.percent)}
                  onCommit={(text) => onEdit(r.id, { percent: parsePercentInput(text) })}
                  className="text-right"
                  title={r.suggestedSourceLine ? `Suggested from: "${r.suggestedSourceLine}"` : 'No suggestion found in Description'}
                />
              </td>
              <td className="xcell w-16">
                <select
                  className="cell-input"
                  value={r.estimator}
                  onChange={(e) => onEdit(r.id, { estimator: e.target.value })}
                  title="Estimator override"
                >
                  <option value="">—</option>
                  {options
                    .filter((o) => o !== '')
                    .map((o) => (
                      <option key={o} value={o}>
                        {o}
                      </option>
                    ))}
                </select>
              </td>
              <td className="xcell w-16 text-right">{r.lastPaid === null ? '-' : fmtPercent(r.lastPaid)}</td>
              <td className="xcell max-w-[240px] truncate" title={r.qboJobName}>
                {r.qboJobName}
              </td>
              <td
                className={`xcell max-w-[200px] truncate ${r.inQboTime === 'NotClockedIn' ? 'font-semibold text-red-700' : ''}`}
                title={r.inQboTime}
              >
                {r.inQboTime}
              </td>
              <td className="xcell whitespace-nowrap text-right">{fmtMoney(r.jobTotal)}</td>
              <td className="xcell w-24 text-right">
                <EditableCell
                  display={fmtMoney(r.apptivo)}
                  onCommit={(text) => {
                    const n = parseMoneyInput(text);
                    onEdit(r.id, { apptivo: n === null ? undefined : n });
                  }}
                  className="text-right"
                  title="Apptivo $ — defaults to Job Total; adjust when Apptivo disagrees (clear to reset)"
                />
              </td>
              <td className={`xcell whitespace-nowrap text-right ${r.dollarDiff !== 0 ? 'font-semibold' : ''}`}>
                {fmtMoney(r.dollarDiff)}
              </td>
              <td className="xcell w-14 text-right">{fmtPercent(r.percent)}</td>
              <td className="xcell whitespace-nowrap">{fmtDate(r.invoiceDate)}</td>
              <td className="xcell text-right">{r.dateDiff === null ? '' : r.dateDiff}</td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}
