import { useMemo, useState, type CSSProperties } from 'react';
import { normalizeApInput } from '../lib/apkey';
import {
  fmtAccounting,
  fmtDate,
  fmtDateShort,
  fmtJobTotal,
  fmtPercent,
  isDarkColor,
  parseMoneyInput,
  parsePercentInput,
} from '../lib/format';
import {
  dollarDiffIsFlagged,
  duplicateInQboValues,
  percentIsFlagged,
  percentIsRedFlagged,
} from '../lib/timenavi';
import type { ReportRow, RowEdit } from '../lib/types';
import { CellInput } from './CellInput';
import { withScrollKept } from './cellNav';

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

/** Round 4 §1.2 — horizontally centered columns (header + data). */
const CENTERED_HEADERS = new Set(['AP#', '% Completion', 'Estimator', 'Last % Paid', '$ Difference']);

/**
 * Round 8 §2.1 — hidden in the export, so collapsed to a stub on screen too.
 * The data stays in the DOM (and in the model); the header toggle expands it.
 */
const COLLAPSIBLE_HEADERS = ['Attendees', 'Location'] as const;
type CollapsibleHeader = (typeof COLLAPSIBLE_HEADERS)[number];

/** Round 4 §1.1 — Excel's classic "bad" style. */
const PINK: CSSProperties = { backgroundColor: '#FFC7CE', color: '#9C0006' };

/** Round 8 §2.4 — no valid % state (FFFF0000 + white bold, matching the export). */
const RED: CSSProperties = { backgroundColor: '#FF0000', color: '#FFFFFF', fontWeight: 700 };

export function TimeNaviTable({ rows, colorByEstimator, estimatorOptions, onEdit }: TimeNaviTableProps): JSX.Element {
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(new Set());
  const [shownCols, setShownCols] = useState<ReadonlySet<string>>(new Set());
  const dupInQbo = useMemo(() => duplicateInQboValues(rows), [rows]);

  const toggleExpanded = (id: string): void => {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const isCollapsed = (h: CollapsibleHeader): boolean => !shownCols.has(h);

  const toggleColumn = (h: CollapsibleHeader): void => {
    setShownCols((prev) => {
      const next = new Set(prev);
      if (next.has(h)) next.delete(h);
      else next.add(h);
      return next;
    });
  };

  /**
   * A collapsed column keeps its text in the cell, clipped to a hairline. The
   * inner div (not the td) carries the width — table cells treat width as a
   * suggestion and would grow for long unbreakable values like emails.
   */
  const stubCell = (text: string): JSX.Element => (
    <td className="xcell p-0" title={text}>
      <div className="w-[7px] overflow-hidden text-transparent">{text}</div>
    </td>
  );

  return (
    <table className="w-max min-w-full border-collapse bg-white text-xs">
      <thead>
        <tr>
          {HEADERS.map((h) => {
            const collapsible = (COLLAPSIBLE_HEADERS as readonly string[]).includes(h);
            const collapsed = collapsible && isCollapsed(h as CollapsibleHeader);
            return (
              <th
                key={h}
                className={`xhead sticky top-0 z-10 whitespace-nowrap ${
                  CENTERED_HEADERS.has(h) ? 'text-center' : ''
                } ${collapsed ? 'p-0' : ''}`}
                title={collapsed ? `${h} — hidden (click to expand)` : undefined}
              >
                {collapsible ? (
                  <button
                    type="button"
                    className={`cursor-pointer overflow-hidden text-left ${collapsed ? 'w-[7px]' : 'w-full'}`}
                    onClick={() => toggleColumn(h as CollapsibleHeader)}
                    title={collapsed ? `Expand ${h}` : `Collapse ${h}`}
                    aria-label={collapsed ? `Expand ${h}` : `Collapse ${h}`}
                    aria-expanded={!collapsed}
                  >
                    {collapsed ? '\u22ee' : `${h} \u2212`}
                  </button>
                ) : (
                  h
                )}
              </th>
            );
          })}
        </tr>
      </thead>
      <tbody>
        {rows.map((r) => {
          const hex = r.included ? colorByEstimator.get(r.estimator) : undefined;
          const isExpanded = expanded.has(r.id);
          const options = estimatorOptions.includes(r.estimator)
            ? estimatorOptions
            : [...estimatorOptions, r.estimator];
          const pctStyle = !r.included
            ? undefined
            : percentIsRedFlagged(r.percent)
              ? RED
              : percentIsFlagged(r.percent)
                ? PINK
                : undefined;
          return (
            <tr
              key={r.id}
              id={`tn-${r.id}`}
              className={r.included ? '' : 'bg-neutral-100 text-neutral-400'}
              style={hex ? { backgroundColor: `#${hex}`, color: isDarkColor(hex) ? '#fff' : '#000' } : undefined}
            >
              <td className="xcell text-center">
                <input
                  type="checkbox"
                  checked={r.included}
                  onChange={(e) => withScrollKept(e.currentTarget, () => onEdit(r.id, { included: !r.included }))}
                  title={r.included ? 'Exclude this event from the report' : 'Include this event in the report'}
                />
              </td>
              <td className="xcell min-w-[260px] max-w-[340px]">{r.event}</td>
              <td className="xcell whitespace-nowrap">{fmtDateShort(r.startDate)}</td>
              <td className="xcell whitespace-nowrap">{fmtDateShort(r.endDate)}</td>
              <td className="xcell whitespace-nowrap">{r.calendar}</td>
              {isCollapsed('Attendees') ? (
                stubCell(r.attendees)
              ) : (
                <td className="xcell max-w-[140px] truncate" title={r.attendees}>
                  {r.attendees}
                </td>
              )}
              {isCollapsed('Location') ? (
                stubCell(r.location)
              ) : (
                <td className="xcell max-w-[160px] truncate" title={r.location}>
                  {r.location}
                </td>
              )}
              <td className="xcell min-w-[240px] max-w-[320px]">
                <div
                  className={`cursor-pointer whitespace-pre-line ${isExpanded ? '' : 'line-clamp-2'}`}
                  onClick={() => toggleExpanded(r.id)}
                  title={isExpanded ? 'Click to collapse' : 'Click to expand'}
                >
                  {r.description}
                </div>
              </td>
              <td className="xcell min-w-[5.5rem] text-center font-mono">
                <CellInput
                  value={r.apKey ?? ''}
                  onCommit={(text) => onEdit(r.id, { apKey: normalizeApInput(text) })}
                  className="text-center"
                  title="AP# — edit to fix a wrong calendar AP#; all lookups re-run"
                />
              </td>
              {/* pctStyle also goes on the input: inline colors outrank
                  cell-input's focus:bg-white, so the flag stays readable
                  while the cell is being edited. */}
              <td className="xcell w-20 text-center" style={pctStyle}>
                <CellInput
                  value={fmtPercent(r.percent)}
                  onCommit={(text) => onEdit(r.id, { percent: parsePercentInput(text) })}
                  className="text-center"
                  style={pctStyle}
                  title={
                    r.suggestedSourceLine
                      ? `Suggested from: "${r.suggestedSourceLine}" — type a %, HOURLY, SUBBED, SET RATE or any note`
                      : 'No suggestion found in Description — type a %, HOURLY, SUBBED, SET RATE or any note'
                  }
                />
              </td>
              <td className="xcell w-16 text-center">
                <select
                  className="cell-input text-center"
                  value={r.estimator}
                  onChange={(e) => {
                    const value = e.target.value;
                    withScrollKept(e.currentTarget, () => onEdit(r.id, { estimator: value }));
                  }}
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
              <td className="xcell w-16 text-center">{r.lastPaid === null ? '-' : fmtPercent(r.lastPaid)}</td>
              <td className="xcell max-w-[240px] truncate" title={r.qboJobName}>
                {r.qboJobName}
              </td>
              <td
                className="xcell max-w-[200px] truncate"
                style={
                  r.included && (r.inQboTime === 'NotClockedIn' || dupInQbo.has(r.inQboTime.trim()))
                    ? PINK
                    : undefined
                }
                title={r.inQboTime}
              >
                {r.inQboTime}
              </td>
              <td className="xcell whitespace-nowrap text-right">{fmtJobTotal(r.jobTotal)}</td>
              <td className="xcell w-24 text-right">
                <CellInput
                  value={r.apptivo === null ? '' : fmtAccounting(r.apptivo)}
                  onCommit={(text) => {
                    const n = parseMoneyInput(text);
                    onEdit(r.id, { apptivo: n === null ? undefined : n });
                  }}
                  className="text-right"
                  title="Apptivo $ — blank by default; type the Apptivo amount to compare against Job Total (clear to blank again)"
                />
              </td>
              <td
                className="xcell whitespace-nowrap text-center"
                style={r.included && dollarDiffIsFlagged(r.dollarDiff) ? PINK : undefined}
              >
                {r.dollarDiff === null ? '' : fmtAccounting(r.dollarDiff)}
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
