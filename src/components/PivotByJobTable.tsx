import { useState, type CSSProperties } from 'react';
import { fmtAccounting, fmtHoursFixed, fmtPercent, isDarkColor, parseMoneyInput, parsePercentInput } from '../lib/format';
import {
  fmtMonthDay,
  PBJ_GRAND_ROW_ID,
  pbjCustomerRowId,
  pbjEmployeeRowId,
  weekdayShort,
} from '../lib/pivotByJob';
import type { ManualSectionRow, PbjJobEdit, PercentValue, PivotByJobModel } from '../lib/types';
import { CellInput } from './CellInput';

interface PivotByJobTableProps {
  model: PivotByJobModel;
  onEditJob: (jobId: string, patch: PbjJobEdit) => void;
  onEditQbHours: (jobId: string, username: string, value: number | null) => void;
  onManualEdit: (section: 'balances' | 'ut', rowId: string, patch: Partial<ManualSectionRow>) => void;
  onManualAdd: (section: 'balances' | 'ut') => void;
  onManualDelete: (section: 'balances' | 'ut', rowId: string) => void;
}

const BLACK: CSSProperties = { backgroundColor: '#000', color: '#fff' };
const GRAY: CSSProperties = { backgroundColor: '#808080', color: '#fff' };
/** Round 8 §4.1 — the "excluded from payroll" swatch, on screen and on export. */
const EXCLUSION_GRAY: CSSProperties = { backgroundColor: '#A6A6A6', color: '#000' };
/**
 * Round 11 — `% of Completion` is black on every row. On a dark estimator band
 * (ND's blue) the cell used to inherit the band's white text.
 */
const PCT_TEXT: CSSProperties = { color: '#000' };

/** The analysis block: AP# · On Calendar? · Estimator · Crew Lead · WO Override · WO% · % · Notes. */
const ANALYSIS_COLS = 8;

/** Tooltip for a % cell that mirrors this week's TimeNavi value (§3.1). */
function percentLinkTitle(live: PercentValue, override: boolean): string {
  const shown = fmtPercent(live) || '(blank)';
  return override
    ? `Manual override — this week's TimeNavi Export says ${shown}`
    : `Live from this week's TimeNavi Export (${shown}) — type here to override`;
}

/** Marker + one-click relink shown on a manually overridden % cell. */
function OverrideBadge({ live, onRelink }: { live: PercentValue; onRelink: () => void }): JSX.Element {
  return (
    <button
      type="button"
      className="ml-1 inline-flex h-4 shrink-0 items-center rounded-xs border border-amber-700 bg-amber-100 px-1 text-[9px] font-bold leading-none text-amber-900"
      title={`Manual override — click to relink to TimeNavi (${fmtPercent(live) || 'blank'})`}
      aria-label="Manual override — click to relink to TimeNavi"
      onClick={(e) => {
        e.stopPropagation();
        onRelink();
      }}
    >
      M
    </button>
  );
}

function bandStyle(hex: string | null): CSSProperties | undefined {
  if (!hex) return undefined;
  return { backgroundColor: `#${hex}`, color: isDarkColor(hex) ? '#fff' : '#000' };
}

function Toggle({ open, onClick, label }: { open: boolean; onClick: () => void; label: string }): JSX.Element {
  return (
    <button
      type="button"
      className="mr-1 inline-flex h-4 w-4 shrink-0 items-center justify-center rounded-xs border border-neutral-500 bg-white text-[10px] leading-none text-neutral-800"
      onClick={(e) => {
        e.stopPropagation();
        onClick();
      }}
      title={open ? `Collapse ${label}` : `Expand ${label}`}
      aria-label={open ? `Collapse ${label}` : `Expand ${label}`}
      aria-expanded={open}
    >
      {open ? '−' : '+'}
    </button>
  );
}

/** Editable text cell that commits on blur/Enter (Round 8 §12.2 — see CellInput). */
function EditCell(props: {
  value: string;
  onCommit: (text: string) => void;
  className?: string;
  title?: string;
  placeholder?: string;
  style?: CSSProperties;
}): JSX.Element {
  return <CellInput {...props} />;
}

/**
 * Round 12 — the analysis block of a row that is not a job (customer,
 * employee, Grand Total). Every cell of `On Calendar?` and `Notes` is
 * editable, so these rows get the two inputs in their own columns instead of
 * one blank merged cell. `leading` is the QB Hrs cell when the row has one.
 */
function RowNoteCells({
  onCalendar,
  notes,
  onEdit,
  cellStyle,
  notesColor,
  leading,
  bold,
}: {
  onCalendar: string;
  notes: string;
  onEdit: (patch: { onCalendar?: string; notes?: string }) => void;
  cellStyle?: CSSProperties;
  notesColor: string;
  leading?: JSX.Element;
  bold?: boolean;
}): JSX.Element {
  const weight = bold ? 'font-bold' : '';
  return (
    <>
      {leading ?? <td className="xcell" style={cellStyle} />}
      <td className="xcell" style={cellStyle} />
      <td className={`xcell text-center ${weight}`} style={cellStyle}>
        <EditCell
          value={onCalendar}
          onCommit={(t) => onEdit({ onCalendar: t.trim() })}
          className="text-center focus:text-black"
          title="On Calendar? — type to add text on this row"
        />
      </td>
      <td className="xcell" style={cellStyle} colSpan={ANALYSIS_COLS - 3} />
      <td className={`xcell ${weight}`} style={{ ...cellStyle, color: notesColor }}>
        <EditCell
          value={notes}
          onCommit={(t) => onEdit({ notes: t })}
          className="focus:text-[#C00000]"
          title="Notes — type to add a note on this row"
        />
      </td>
      <td className="xcell" style={cellStyle} />
    </>
  );
}

/** Notes text: dark red, lightened on a dark estimator band so it stays readable. */
const notesColorOn = (hex: string | null): string => (hex && isDarkColor(hex) ? '#FFC7C7' : '#C00000');

const money = (v: number | null): string => (v !== null ? `$${fmtAccounting(v)}` : '');
const pctText = (v: number | null): string => (v !== null ? `${Math.round(v * 100)}%` : '');

/**
 * Round 11 — the WO Override placeholder. A suggested WO% is turned into
 * dollars against the Job Total $ by the model, so the cell reads
 * "suggest $1,234.56" whichever way the calendar phrased it.
 */
function woOverridePlaceholder(suggestion: number | null, hourly: boolean): string {
  if (suggestion !== null) return `suggest $${fmtAccounting(suggestion)}`;
  return hourly ? '-' : '—';
}

export function PivotByJobTable(props: PivotByJobTableProps): JSX.Element {
  const { model, onEditJob, onEditQbHours } = props;
  const dateCount = model.dateIsos.length;
  /**
   * Round 11 — every section lives in ONE table, so every column lines up
   * across the pivot, BALANCES and UT PAYROLL. They used to be three tables,
   * each sizing its own columns, so the bottom sections never aligned.
   */
  const totalCols = 1 + dateCount + 2 + ANALYSIS_COLS + 1;
  // Round 5 §3.3 — view-only collapse state (never affects totals/exports).
  const [collapsedCust, setCollapsedCust] = useState<ReadonlySet<string>>(new Set());
  const [collapsedJob, setCollapsedJob] = useState<ReadonlySet<string>>(new Set());

  const toggle = (setter: typeof setCollapsedCust, key: string): void => {
    setter((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const hourCells = (byDate: Record<string, number>, border?: CSSProperties): JSX.Element[] =>
    model.dateIsos.map((iso) => (
      <td key={iso || 'blank'} className="xcell text-right" style={border}>
        {byDate[iso] !== undefined ? fmtHoursFixed(byDate[iso]) : ''}
      </td>
    ));

  const hb: CSSProperties = { borderTop: '1px solid #000', borderBottom: '1px solid #000' };

  /** Row-1 cells shared by the top header and each bottom section's header. */
  const headerRow = (title: string, sticky: boolean): JSX.Element => {
    const stickyCls = sticky ? 'sticky top-0 z-10' : '';
    return (
      <tr>
        <th className={`xcell2 ${stickyCls} min-w-[260px] text-left text-base font-bold`} style={BLACK}>
          {title}
        </th>
        {model.dates.map((d) => (
          <th key={d.getTime()} className={`xcell2 ${stickyCls} min-w-[72px] text-center font-bold`} style={BLACK}>
            {fmtMonthDay(d)}
          </th>
        ))}
        {model.dateIsos.length > model.dates.length && <th className={`xcell2 ${stickyCls}`} style={BLACK} />}
        <th className={`xcell2 ${stickyCls} min-w-[70px] text-center font-bold`} style={BLACK}>
          Total hrs
        </th>
        <th className={`xcell2 ${stickyCls} min-w-[56px]`} style={BLACK} title="QB Hrs override (optional)" />
        <th className={`xcell2 ${stickyCls} min-w-[76px] text-center`} style={BLACK}>AP#</th>
        <th className={`xcell2 ${stickyCls} min-w-[136px] text-center`} style={BLACK}>On Calendar?</th>
        <th className={`xcell2 ${stickyCls} min-w-[64px] text-center`} style={BLACK}>Estimator</th>
        <th className={`xcell2 ${stickyCls} min-w-[220px] text-center`} style={BLACK}>Crew Lead</th>
        <th className={`xcell2 ${stickyCls} min-w-[130px] text-center`} style={BLACK}>WO Override</th>
        <th className={`xcell2 ${stickyCls} min-w-[64px] text-center`} style={BLACK}>WO%</th>
        <th className={`xcell2 ${stickyCls} min-w-[126px] text-center`} style={{ ...BLACK, color: '#FF0000' }}>
          % of Completion
        </th>
        <th className={`xcell2 ${stickyCls} min-w-[180px] text-left`} style={{ ...BLACK, color: '#C00000' }}>
          Notes
        </th>
        <th className={`xcell2 ${stickyCls}`} style={BLACK} />
      </tr>
    );
  };

  return (
    <div className="p-2">
      <table className="w-max min-w-full border-collapse bg-white text-xs">
        <thead>
          {headerRow(model.title, true)}
          {/* Row 2 — computed weekday names (all 7 days, including Sun). */}
          <tr>
            <th className="xcell2" style={BLACK} />
            {model.dates.map((d) => (
              <th key={d.getTime()} className="xcell2 text-center font-normal" style={BLACK}>
                {weekdayShort(d)}
              </th>
            ))}
            {model.dateIsos.length > model.dates.length && (
              <th className="xcell2 font-normal" style={BLACK}>
                (blank)
              </th>
            )}
            <th className="xcell2" style={BLACK} colSpan={2 + ANALYSIS_COLS + 1} />
          </tr>
        </thead>
        <tbody>
          {model.customers.map((cust) => {
            const custOpen = !collapsedCust.has(cust.customer);
            const custStyle = { ...bandStyle(cust.colorHex), ...hb };
            return (
              <CustomerBlock
                key={cust.customer}
                cust={cust}
                custOpen={custOpen}
                onToggleCust={() => toggle(setCollapsedCust, cust.customer)}
                collapsedJob={collapsedJob}
                onToggleJob={(id) => toggle(setCollapsedJob, id)}
                hourCells={hourCells}
                onEditJob={onEditJob}
                onEditQbHours={onEditQbHours}
                dateIsos={model.dateIsos}
                hb={hb}
                custStyle={custStyle}
              />
            );
          })}
          <tr>
            <td className="xcell font-bold" style={GRAY}>
              Grand Total
            </td>
            {model.dateIsos.map((iso) => (
              <td key={iso || 'blank'} className="xcell text-right font-bold" style={GRAY}>
                {model.grandByDate[iso] !== undefined ? fmtHoursFixed(model.grandByDate[iso]) : ''}
              </td>
            ))}
            <td className="xcell text-right font-bold" style={GRAY}>
              {fmtHoursFixed(model.grandTotal)}
            </td>
            <RowNoteCells
              onCalendar={model.grandOnCalendar ?? ''}
              notes={model.grandNotes ?? ''}
              onEdit={(patch) => onEditJob(PBJ_GRAND_ROW_ID, patch)}
              cellStyle={GRAY}
              notesColor="#fff"
              bold
            />
          </tr>
        </tbody>

        <ManualSection
          section="balances"
          title={model.title.replace('NV PAYROLL:', 'NV PAYROLL (BALANCES):')}
          rows={model.balances}
          dateCount={dateCount}
          dateIsos={model.dateIsos}
          totalCols={totalCols}
          headerRow={headerRow}
          onEditJob={onEditJob}
          onEditQbHours={onEditQbHours}
          onManualEdit={props.onManualEdit}
          onManualAdd={props.onManualAdd}
          onManualDelete={props.onManualDelete}
        />
        <ManualSection
          section="ut"
          title={model.title.replace('NV PAYROLL:', 'UT PAYROLL:')}
          rows={model.utPayroll}
          dateCount={dateCount}
          dateIsos={model.dateIsos}
          totalCols={totalCols}
          headerRow={headerRow}
          onEditJob={onEditJob}
          onEditQbHours={onEditQbHours}
          onManualEdit={props.onManualEdit}
          onManualAdd={props.onManualAdd}
          onManualDelete={props.onManualDelete}
        />
      </table>
      <p className="mt-2 text-[11px] italic text-neutral-500">
        Balances &amp; UT sections auto-position below the pivot — the Excel move-refresh-move workaround is gone.
      </p>
    </div>
  );
}

function CustomerBlock({
  cust,
  custOpen,
  onToggleCust,
  collapsedJob,
  onToggleJob,
  hourCells,
  onEditJob,
  onEditQbHours,
  dateIsos,
  hb,
  custStyle,
}: {
  cust: PivotByJobModel['customers'][number];
  custOpen: boolean;
  onToggleCust: () => void;
  collapsedJob: ReadonlySet<string>;
  onToggleJob: (id: string) => void;
  hourCells: (byDate: Record<string, number>, border?: CSSProperties) => JSX.Element[];
  onEditJob: PivotByJobTableProps['onEditJob'];
  onEditQbHours: PivotByJobTableProps['onEditQbHours'];
  dateIsos: string[];
  hb: CSSProperties;
  custStyle: CSSProperties;
}): JSX.Element {
  return (
    <>
      <tr style={custStyle}>
        <td className="xcell font-bold" style={hb}>
          <span className="inline-flex items-center">
            <Toggle open={custOpen} onClick={onToggleCust} label={cust.customer} />
            {cust.customer}
          </span>
        </td>
        {hourCells(cust.byDate, hb)}
        <td className="xcell text-right font-bold" style={hb}>
          {fmtHoursFixed(cust.total)}
        </td>
        <RowNoteCells
          onCalendar={cust.onCalendar ?? ''}
          notes={cust.notes ?? ''}
          onEdit={(patch) => onEditJob(pbjCustomerRowId(cust.customer), patch)}
          cellStyle={hb}
          notesColor={notesColorOn(cust.colorHex)}
          bold
        />
      </tr>
      {custOpen &&
        cust.jobs.map((job) => (
          <JobBlock
            key={job.id}
            job={job}
            jobOpen={!collapsedJob.has(job.id)}
            onToggleJob={() => onToggleJob(job.id)}
            hourCells={hourCells}
            onEditJob={onEditJob}
            onEditQbHours={onEditQbHours}
            dateIsos={dateIsos}
            hb={hb}
          />
        ))}
    </>
  );
}

function JobBlock({
  job,
  jobOpen,
  onToggleJob,
  hourCells,
  onEditJob,
  onEditQbHours,
  dateIsos,
  hb,
}: {
  job: PivotByJobModel['customers'][number]['jobs'][number];
  jobOpen: boolean;
  onToggleJob: () => void;
  hourCells: (byDate: Record<string, number>, border?: CSSProperties) => JSX.Element[];
  onEditJob: PivotByJobTableProps['onEditJob'];
  onEditQbHours: PivotByJobTableProps['onEditQbHours'];
  dateIsos: string[];
  hb: CSSProperties;
}): JSX.Element {
  const a = job.analysis;
  const style = bandStyle(job.colorHex);
  const rowStyle = { ...style, ...hb };
  const edit = (patch: PbjJobEdit): void => onEditJob(job.id, patch);
  return (
    <>
      <tr id={`pbj-${job.id}`} style={rowStyle}>
        <td className="xcell pl-4 font-bold" style={hb}>
          <span className="inline-flex items-center gap-1">
            <Toggle open={jobOpen} onClick={onToggleJob} label={job.jobcode2} />
            {job.jobcode2}
            {a.dbPatternMatch && (
              <button
                type="button"
                className={`ml-1 inline-flex h-4 items-center rounded-xs px-1 text-[9px] font-bold leading-none ${
                  a.dbOverride
                    ? 'border border-neutral-400 bg-neutral-100 text-neutral-500 line-through'
                    : 'border border-amber-700 bg-amber-100 text-amber-900'
                }`}
                title={
                  a.dbOverride
                    ? 'Davis-Bacon pattern matched — click to re-enable DB floor'
                    : 'Davis-Bacon job — click to force non-DB for this row'
                }
                onClick={(e) => {
                  e.stopPropagation();
                  edit({ dbOverride: !a.dbOverride });
                }}
              >
                DB
              </button>
            )}
          </span>
        </td>
        {hourCells(job.byDate, hb)}
        <td className="xcell text-right font-bold" style={hb}>
          {fmtHoursFixed(job.total)}
        </td>
        <td className="xcell" style={hb} />
        <td className="xcell text-center font-bold">
          <EditCell
            value={a.apKey ?? ''}
            onCommit={(t) => edit({ apKey: t.trim() === '' ? null : t.trim().toUpperCase() })}
            className="text-center"
          />
        </td>
        {/* Round 12 — On Calendar? is editable; clearing it brings the computed value back. */}
        <td className={`xcell text-center font-bold ${a.onCalendar === 'NOT ON CALENDAR' ? 'text-red-700' : ''}`}>
          <EditCell
            value={a.onCalendar}
            onCommit={(t) => edit({ onCalendar: t.trim() === '' ? undefined : t.trim() })}
            className="text-center"
            title="Type to override; clear the cell to go back to the calendar lookup"
          />
        </td>
        {/* Round 11 — Estimator is editable on every row. */}
        <td className="xcell text-center font-bold">
          <EditCell value={a.estimator} onCommit={(t) => edit({ estimator: t.trim().toUpperCase() })} className="text-center" />
        </td>
        <td className="xcell text-center font-bold" title={a.crewLeadResolution?.raw ? `From "(${a.crewLeadResolution.raw})"` : undefined}>
          <EditCell value={a.crewLead} onCommit={(t) => edit({ crewLead: t })} className="text-center" placeholder="—" />
        </td>
        {/* Round 11 — WO Override / WO% stay editable on HOURLY jobs. */}
        <td
          className="xcell text-center font-bold"
          title={
            a.woOverrideSuggestion !== null
              ? `Suggested $${fmtAccounting(a.woOverrideSuggestion)}${
                  a.woOverrideSuggestionSource ? ` — ${a.woOverrideSuggestionSource}` : ''
                }`
              : undefined
          }
        >
          <EditCell
            value={money(a.woOverride)}
            onCommit={(t) => edit({ woOverride: t.trim() === '' || t.trim() === '-' ? null : parseMoneyInput(t) })}
            className="text-center"
            placeholder={woOverridePlaceholder(a.woOverrideSuggestion, a.hourly)}
          />
        </td>
        <td
          className="xcell text-center font-bold"
          title={
            a.woPctSuggestion !== null
              ? `Calendar suggests ${Math.round(a.woPctSuggestion * 100)}%${
                  a.woPctSuggestionSource ? ` — ${a.woPctSuggestionSource}` : ''
                }`
              : undefined
          }
        >
          <EditCell
            value={pctText(a.woPct)}
            onCommit={(t) => {
              const p = parsePercentInput(t);
              edit({ woPct: typeof p === 'number' ? p : null });
            }}
            className="text-center"
            placeholder={
              a.woPctSuggestion !== null ? `suggest ${Math.round(a.woPctSuggestion * 100)}%` : a.hourly ? '-' : '—'
            }
          />
        </td>
        <td className="xcell text-center font-bold" style={PCT_TEXT} title={percentLinkTitle(a.livePercent, a.percentOverride)}>
          <span className="inline-flex w-full items-center justify-center">
            <EditCell
              value={fmtPercent(a.percent)}
              onCommit={(t) => edit({ percent: t.trim() === '' ? undefined : parsePercentInput(t) })}
              className="text-center"
              style={PCT_TEXT}
              placeholder="—"
            />
            {a.percentOverride && <OverrideBadge live={a.livePercent} onRelink={() => edit({ percent: undefined })} />}
          </span>
        </td>
        <td className="xcell font-bold" style={{ color: notesColorOn(job.colorHex) }}>
          <EditCell value={a.notes} onCommit={(t) => edit({ notes: t })} className="focus:text-[#C00000]" title="Notes" />
        </td>
        <td className="xcell" />
      </tr>
      {jobOpen &&
        job.employees.map((emp) => (
          <tr key={emp.username} style={style}>
            <td className="xcell pl-8">{emp.username}</td>
            {dateIsos.map((iso) => (
              <td key={iso || 'blank'} className="xcell text-right">
                {emp.byDate[iso] !== undefined ? fmtHoursFixed(emp.byDate[iso]) : ''}
              </td>
            ))}
            <td className="xcell text-right">{fmtHoursFixed(emp.total)}</td>
            <RowNoteCells
              onCalendar={emp.onCalendar ?? ''}
              notes={emp.notes ?? ''}
              onEdit={(patch) => onEditJob(pbjEmployeeRowId(job.id, emp.username), patch)}
              notesColor={notesColorOn(job.colorHex)}
              leading={
                <td className="xcell text-right" title="QB Hrs override — used downstream instead of the computed hours">
                  <EditCell
                    value={emp.qbHours !== null ? String(emp.qbHours) : ''}
                    onCommit={(t) => {
                      const n = Number(t.trim());
                      onEditQbHours(job.id, emp.username, t.trim() === '' || !Number.isFinite(n) ? null : n);
                    }}
                    className="text-right font-semibold"
                  />
                </td>
              }
            />
          </tr>
        ))}
    </>
  );
}

/**
 * Round 5 §3.2 — one black header row (mirrors row 1), then customer+job pairs.
 * Round 11 — a <tbody> of the main table rather than a table of its own, so its
 * columns line up with the pivot above.
 */
function ManualSection({
  section,
  title,
  rows,
  dateCount,
  dateIsos,
  totalCols,
  headerRow,
  onEditJob,
  onEditQbHours,
  onManualEdit,
  onManualAdd,
  onManualDelete,
}: {
  section: 'balances' | 'ut';
  title: string;
  rows: ManualSectionRow[];
  dateCount: number;
  dateIsos: string[];
  totalCols: number;
  headerRow: (title: string, sticky: boolean) => JSX.Element;
} & Pick<
  PivotByJobTableProps,
  'onEditJob' | 'onEditQbHours' | 'onManualEdit' | 'onManualAdd' | 'onManualDelete'
>): JSX.Element {
  return (
    <tbody>
      {/* Visual gap between sections, inside the same column grid. */}
      <tr aria-hidden="true">
        <td colSpan={totalCols} className="h-4 border-0 bg-transparent" />
      </tr>
      {headerRow(title, false)}
      {rows.length === 0 && (
        <tr>
          <td className="xcell text-neutral-400" colSpan={totalCols}>
            No rows.{' '}
            {section === 'balances'
              ? 'Partially-paid NV jobs from last week (slot 8) seed here.'
              : 'Utah timesheet jobs (slot 2), Utah-calendar events and last week’s open UT rows seed here.'}
          </td>
        </tr>
      )}
      {rows.map((r) => (
        <ManualPair
          key={r.id}
          section={section}
          r={r}
          style={r.excludedFromPayroll ? EXCLUSION_GRAY : bandStyle(r.colorHex)}
          dateCount={dateCount}
          dateIsos={dateIsos}
          onEditJob={onEditJob}
          onEditQbHours={onEditQbHours}
          onManualEdit={onManualEdit}
          onManualDelete={onManualDelete}
        />
      ))}
      <tr>
        <td className="xcell" colSpan={totalCols}>
          <button type="button" className="btn text-xs" onClick={() => onManualAdd(section)}>
            + Add row
          </button>
        </td>
      </tr>
    </tbody>
  );
}

function ManualPair({
  section,
  r,
  style,
  dateCount,
  dateIsos,
  onEditJob,
  onEditQbHours,
  onManualEdit,
  onManualDelete,
}: {
  section: 'balances' | 'ut';
  r: ManualSectionRow;
  style: CSSProperties | undefined;
  dateCount: number;
  dateIsos: string[];
  onEditJob: PivotByJobTableProps['onEditJob'];
  onEditQbHours: PivotByJobTableProps['onEditQbHours'];
  onManualEdit: PivotByJobTableProps['onManualEdit'];
  onManualDelete: PivotByJobTableProps['onManualDelete'];
}): JSX.Element {
  const edit = (patch: Partial<ManualSectionRow>): void => onManualEdit(section, r.id, patch);
  /**
   * Round 12 — a Utah job with timesheet hours shows them the way the pivot
   * above does: hours by day on the customer and job rows, then one row per
   * employee. A no-hours job keeps the merged free-text cell.
   */
  const hours = r.byDate && r.total !== undefined ? { byDate: r.byDate, total: r.total } : null;
  const hourCells = (byDate: Record<string, number>, bold: boolean): JSX.Element[] =>
    dateIsos.map((iso) => (
      <td key={iso || 'blank'} className={`xcell text-right ${bold ? 'font-bold' : ''}`}>
        {byDate[iso] !== undefined ? fmtHoursFixed(byDate[iso]) : ''}
      </td>
    ));
  return (
    <>
      <tr style={style}>
        <td className="xcell font-bold">
          <EditCell value={r.customer} onCommit={(t) => edit({ customer: t })} />
        </td>
        {hours ? (
          <>
            {hourCells(hours.byDate, true)}
            <td className="xcell text-right font-bold">{fmtHoursFixed(hours.total)}</td>
          </>
        ) : (
          <td className="xcell" colSpan={dateCount + 1} />
        )}
        <RowNoteCells
          onCalendar={r.customerOnCalendar ?? ''}
          notes={r.customerNotes ?? ''}
          onEdit={(patch) =>
            edit({
              ...(patch.onCalendar !== undefined ? { customerOnCalendar: patch.onCalendar } : {}),
              ...(patch.notes !== undefined ? { customerNotes: patch.notes } : {}),
            })
          }
          notesColor={r.excludedFromPayroll ? '#C00000' : notesColorOn(r.colorHex)}
          bold
        />
      </tr>
      <tr style={style}>
        <td className="xcell pl-4 font-bold">
          <EditCell value={r.job} onCommit={(t) => edit({ job: t })} />
        </td>
        {/*
          Round 11 — a no-hours job has nothing under its dates, so the date
          columns and Total hrs are one merged, centered cell the user can type
          in. It exports merged the same way.
        */}
        {hours ? (
          <>
            {hourCells(hours.byDate, true)}
            <td className="xcell text-right font-bold">{fmtHoursFixed(hours.total)}</td>
          </>
        ) : (
          <td className="xcell text-center" colSpan={dateCount + 1}>
            <EditCell
              value={r.hoursNote ?? ''}
              onCommit={(t) => edit({ hoursNote: t })}
              className="text-center"
              placeholder="—"
            />
          </td>
        )}
        <td className="xcell" />
        <td className="xcell text-center">
          <EditCell value={r.apKey ?? ''} onCommit={(t) => edit({ apKey: t.trim() || null })} className="text-center" />
        </td>
        <td className={`xcell text-center ${r.onCalendar === 'NOT ON CALENDAR' ? 'text-red-700' : ''}`}>
          <EditCell
            value={r.onCalendar ?? ''}
            onCommit={(t) => edit({ onCalendar: t.trim() === '' ? undefined : t.trim() })}
            className="text-center"
            title="Type to override; clear the cell to go back to the calendar lookup"
          />
        </td>
        <td className="xcell text-center">
          <EditCell value={r.estimator} onCommit={(t) => edit({ estimator: t.trim().toUpperCase() })} className="text-center" />
        </td>
        <td className="xcell text-center">
          <EditCell value={r.crewLead} onCommit={(t) => edit({ crewLead: t })} className="text-center" />
        </td>
        <td className="xcell text-center">
          <EditCell
            value={money(r.woOverride)}
            onCommit={(t) => edit({ woOverride: t.trim() === '' ? null : parseMoneyInput(t) })}
            className="text-center"
            placeholder="—"
          />
        </td>
        <td className="xcell text-center">
          <EditCell
            value={pctText(r.woPct)}
            onCommit={(t) => {
              const p = parsePercentInput(t);
              edit({ woPct: typeof p === 'number' ? p : null });
            }}
            className="text-center"
            placeholder="—"
          />
        </td>
        <td className="xcell text-center" style={PCT_TEXT} title={percentLinkTitle(r.livePercent, r.percentOverride)}>
          <span className="inline-flex w-full items-center justify-center">
            <EditCell
              value={fmtPercent(r.percent)}
              onCommit={(t) => edit({ percent: t.trim() === '' ? undefined : parsePercentInput(t) })}
              className="text-center"
              style={PCT_TEXT}
            />
            {r.percentOverride && <OverrideBadge live={r.livePercent} onRelink={() => edit({ percent: undefined })} />}
          </span>
        </td>
        <td className="xcell" style={{ color: r.excludedFromPayroll ? '#C00000' : notesColorOn(r.colorHex) }}>
          <EditCell value={r.notes} onCommit={(t) => edit({ notes: t })} className="focus:text-[#C00000]" title="Notes" />
        </td>
        <td className="xcell text-center">
          <span className="inline-flex items-center gap-1">
            <button
              type="button"
              className={`inline-flex h-4 items-center rounded-xs border px-1 text-[9px] font-bold leading-none ${
                r.excludedFromPayroll
                  ? 'border-neutral-700 bg-[#A6A6A6] text-neutral-900'
                  : 'border-neutral-300 bg-white text-neutral-500'
              }`}
              title={
                r.excludedFromPayroll
                  ? 'Excluded from payroll (A6A6A6) — click to include'
                  : 'Click to exclude this job from payroll (paints A6A6A6 here and on export)'
              }
              aria-pressed={r.excludedFromPayroll}
              onClick={() => edit({ excludedFromPayroll: !r.excludedFromPayroll })}
            >
              EX
            </button>
            <button
              type="button"
              className="text-neutral-400 hover:text-red-600"
              onClick={() => onManualDelete(section, r.id)}
              aria-label="Delete row"
            >
              🗑
            </button>
          </span>
        </td>
      </tr>
      {(r.employees ?? []).map((emp) => (
        <tr key={emp.username} style={style}>
          <td className="xcell pl-8">{emp.username}</td>
          {hourCells(emp.byDate, false)}
          <td className="xcell text-right">{fmtHoursFixed(emp.total)}</td>
          <RowNoteCells
            onCalendar={emp.onCalendar ?? ''}
            notes={emp.notes ?? ''}
            onEdit={(patch) => onEditJob(pbjEmployeeRowId(r.id, emp.username), patch)}
            notesColor={r.excludedFromPayroll ? '#C00000' : notesColorOn(r.colorHex)}
            leading={
              <td className="xcell text-right" title="QB Hrs override — used downstream instead of the computed hours">
                <EditCell
                  value={emp.qbHours !== null ? String(emp.qbHours) : ''}
                  onCommit={(t) => {
                    const n = Number(t.trim());
                    onEditQbHours(r.id, emp.username, t.trim() === '' || !Number.isFinite(n) ? null : n);
                  }}
                  className="text-right font-semibold"
                />
              </td>
            }
          />
        </tr>
      ))}
    </>
  );
}
