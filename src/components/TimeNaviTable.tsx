import { useMemo, useState, type CSSProperties } from 'react';
import { normalizeApInput } from '../lib/apkey';
import { parseDate } from '../lib/parsers';
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
const COLLAPSIBLE_HEADERS = ['Attendees', 'Location', 'QBO Job Name'] as const;
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
      {/* Round 10 — `whitespace-nowrap` matters: without it a long value (an
          attendee email, a QBO job name) wrapped inside the 7px stub and made
          the whole ROW as tall as the wrapped text. */}
      <div className="w-[7px] overflow-hidden whitespace-nowrap text-transparent">{text}</div>
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
              {/* Round 9 §37 / Round 10 — one line, sized to the content; the
                  full title stays in the tooltip and is editable in place. */}
              <td className="xcell max-w-[12rem] truncate whitespace-nowrap" title={r.event}>
                <CellInput
                  value={r.event}
                  onCommit={(text) => onEdit(r.id, { event: text })}
                  title={r.event}
                />
              </td>
              <td className="xcell w-px whitespace-nowrap">
                <CellInput
                  value={fmtDateShort(r.startDate)}
                  title="Start Date — type to override (M/D/YYYY)"
                  onCommit={(text) => onEdit(r.id, { startDate: parseDate(text.trim() || null) })}
                />
              </td>
              <td className="xcell w-px whitespace-nowrap">
                <CellInput
                  value={fmtDateShort(r.endDate)}
                  title="End Date — type to override (M/D/YYYY)"
                  onCommit={(text) => onEdit(r.id, { endDate: parseDate(text.trim() || null) })}
                />
              </td>
              <td className="xcell max-w-[7rem] truncate whitespace-nowrap" title={r.calendar}>
                <CellInput value={r.calendar} onCommit={(text) => onEdit(r.id, { calendar: text })} />
              </td>
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
              <td className="xcell min-w-[7rem] max-w-[9rem]">
                <div
                  className={`cursor-pointer whitespace-pre-line ${isExpanded ? '' : 'line-clamp-1'}`}
                  onClick={() => toggleExpanded(r.id)}
                  title={isExpanded ? 'Click to collapse' : 'Click to expand'}
                >
                  {r.description}
                </div>
              </td>
              {/* Round 10 — the AP# column is never collapsed or clipped. */}
              <td className="xcell w-px min-w-[5.5rem] whitespace-nowrap text-center font-mono">
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
              {/* Round 10 — carried from last week's `% Completion` for this
                  AP#, and editable when the carry-forward needs correcting. */}
              <td className="xcell w-16 text-center">
                <CellInput
                  value={r.lastPaid === null ? '' : fmtPercent(r.lastPaid)}
                  placeholder="-"
                  className="text-center"
                  title="Last % Paid — last week's % Completion for this AP#"
                  onCommit={(text) => {
                    const t = text.trim();
                    if (t === '' || t === '-') {
                      onEdit(r.id, { lastPaid: null });
                      return;
                    }
                    if (/^hourly$/i.test(t)) {
                      onEdit(r.id, { lastPaid: 'HOURLY' });
                      return;
                    }
                    const p = parsePercentInput(t);
                    onEdit(r.id, { lastPaid: typeof p === 'number' ? p : null });
                  }}
                />
              </td>
              {isCollapsed('QBO Job Name') ? (
                stubCell(r.qboJobName)
              ) : (
                <td className="xcell max-w-[10rem] truncate" title={r.qboJobName}>
                  <CellInput
                    value={r.qboJobName}
                    onCommit={(text) => onEdit(r.id, { qboJobName: text })}
                  />
                </td>
              )}
              <td
                className="xcell max-w-[9rem] truncate"
                style={
                  r.included && (r.inQboTime === 'NotClockedIn' || dupInQbo.has(r.inQboTime.trim()))
                    ? PINK
                    : undefined
                }
                title={r.inQboTime}
              >
                <CellInput
                  value={r.inQboTime}
                  onCommit={(text) => onEdit(r.id, { inQboTime: text })}
                  style={
                    r.included && (r.inQboTime === 'NotClockedIn' || dupInQbo.has(r.inQboTime.trim()))
                      ? PINK
                      : undefined
                  }
                />
              </td>
              <td className="xcell w-24 whitespace-nowrap text-right">
                <CellInput
                  value={fmtJobTotal(r.jobTotal)}
                  className="text-right"
                  title="Job Total $ — from the Sales by Customer Summary; type to override"
                  onCommit={(text) => {
                    const t = text.trim();
                    if (t === '' || t.toUpperCase() === '#N/A') {
                      onEdit(r.id, { jobTotal: '#N/A' });
                      return;
                    }
                    const n = parseMoneyInput(t);
                    onEdit(r.id, { jobTotal: n === null ? '#N/A' : n });
                  }}
                />
              </td>
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
              <td className="xcell w-px whitespace-nowrap">
                <CellInput
                  value={fmtDate(r.invoiceDate)}
                  title="Invoice Date — from the books; type to override (M/D/YYYY)"
                  onCommit={(text) => onEdit(r.id, { invoiceDate: parseDate(text.trim() || null) })}
                />
              </td>
              <td className="xcell w-px text-right">{r.dateDiff === null ? '' : r.dateDiff}</td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}
