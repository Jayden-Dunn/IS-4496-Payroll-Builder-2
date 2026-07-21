import { useState, type CSSProperties } from 'react';
import { fmtAccounting, fmtHoursFixed, fmtPercent, isDarkColor, parseMoneyInput, parsePercentInput } from '../lib/format';
import { fmtMonthDay, weekdayShort } from '../lib/pivotByJob';
import type { ManualSectionRow, PbjJobEdit, PivotByJobModel } from '../lib/types';

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

/** Thin top+bottom borders spanning A through the QB-Hrs column (§3.5). */
function hoursBorderStyle(dateCount: number): CSSProperties {
  // A + dates + Total hrs + QB Hrs = 1 + dateCount + 2 columns.
  void dateCount;
  return { borderTop: '1px solid #000', borderBottom: '1px solid #000' };
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

/** Editable text cell that commits on blur/Enter. */
function EditCell({
  value,
  onCommit,
  className,
  title,
  placeholder,
}: {
  value: string;
  onCommit: (text: string) => void;
  className?: string;
  title?: string;
  placeholder?: string;
}): JSX.Element {
  const [text, setText] = useState<string | null>(null);
  return (
    <input
      className={`cell-input ${className ?? ''}`}
      value={text ?? value}
      title={title}
      placeholder={placeholder}
      onChange={(e) => setText(e.target.value)}
      onFocus={() => setText(value)}
      onBlur={() => {
        if (text !== null && text !== value) onCommit(text);
        setText(null);
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
        if (e.key === 'Escape') setText(null);
      }}
    />
  );
}

export function PivotByJobTable(props: PivotByJobTableProps): JSX.Element {
  const { model, onEditJob, onEditQbHours } = props;
  const dateCount = model.dateIsos.length;
  // Round 5 §3.3 — view-only collapse state (never affects totals/exports).
  const [collapsedCust, setCollapsedCust] = useState<ReadonlySet<string>>(new Set());
  const [collapsedJob, setCollapsedJob] = useState<ReadonlySet<string>>(new Set());

  const toggleCust = (key: string): void => {
    setCollapsedCust((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };
  const toggleJob = (key: string): void => {
    setCollapsedJob((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const hourCells = (byDate: Record<string, number>, cls = 'xcell', border?: CSSProperties): JSX.Element[] =>
    model.dateIsos.map((iso) => (
      <td key={iso || 'blank'} className={`${cls} text-right`} style={border}>
        {byDate[iso] !== undefined ? fmtHoursFixed(byDate[iso]) : ''}
      </td>
    ));

  const emptyAnalysis = <>{Array.from({ length: 8 }, (_, i) => <td key={i} className="xcell" />)}</>;
  const hb = hoursBorderStyle(dateCount);

  /** Shared analysis-header cells for row 1 and BALANCES/UT section headers. */
  const analysisHeaders = (sticky = false): JSX.Element => {
    const stickyCls = sticky ? 'sticky top-0 z-10' : '';
    return (
      <>
        <th className={`xcell2 ${stickyCls} min-w-[76px] text-center`} style={BLACK}>AP#</th>
        <th className={`xcell2 ${stickyCls} min-w-[136px] text-center`} style={BLACK}>On Calendar?</th>
        <th className={`xcell2 ${stickyCls} min-w-[64px] text-center`} style={BLACK}>Estimator</th>
        <th className={`xcell2 ${stickyCls} min-w-[220px] text-center`} style={BLACK}>Crew Lead</th>
        <th className={`xcell2 ${stickyCls} min-w-[110px] text-center`} style={BLACK}>WO Override</th>
        <th className={`xcell2 ${stickyCls} min-w-[64px] text-center`} style={BLACK}>WO%</th>
        <th className={`xcell2 ${stickyCls} min-w-[126px] text-center`} style={{ ...BLACK, color: '#FF0000' }}>
          % of Completion
        </th>
        <th className={`xcell2 ${stickyCls} min-w-[160px] text-left`} style={{ ...BLACK, color: '#C00000' }}>
          Notes
        </th>
      </>
    );
  };

  return (
    <div className="p-2">
      <table className="w-max min-w-full border-collapse bg-white text-xs">
        <thead>
          {/* Round 5 §3.1 / Round 6 — row 1: title, short m/d dates, Total hrs, analysis. */}
          <tr>
            <th className="xcell2 sticky top-0 z-10 min-w-[260px] text-left text-base font-bold" style={BLACK}>
              {model.title}
            </th>
            {model.dates.map((d) => (
              <th key={d.getTime()} className="xcell2 sticky top-0 z-10 min-w-[72px] text-center font-bold" style={BLACK}>
                {fmtMonthDay(d)}
              </th>
            ))}
            {model.dateIsos.length > model.dates.length && (
              <th className="xcell2 sticky top-0 z-10" style={BLACK} />
            )}
            <th className="xcell2 sticky top-0 z-10 min-w-[70px] text-center font-bold" style={BLACK}>
              Total hrs
            </th>
            <th className="xcell2 sticky top-0 z-10 min-w-[56px]" style={BLACK} title="QB Hrs override (optional)" />
            {analysisHeaders(true)}
          </tr>
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
            <th className="xcell2" style={BLACK} colSpan={10} />
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
                onToggleCust={() => toggleCust(cust.customer)}
                collapsedJob={collapsedJob}
                onToggleJob={toggleJob}
                hourCells={hourCells}
                emptyAnalysis={emptyAnalysis}
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
            <td className="xcell" style={GRAY} colSpan={9} />
          </tr>
        </tbody>
      </table>

      <ManualSection
        section="balances"
        title={model.title.replace('NV PAYROLL:', 'NV PAYROLL (BALANCES):')}
        rows={model.balances}
        model={model}
        analysisHeaders={analysisHeaders}
        onManualEdit={props.onManualEdit}
        onManualAdd={props.onManualAdd}
        onManualDelete={props.onManualDelete}
      />
      <ManualSection
        section="ut"
        title={model.title.replace('NV PAYROLL:', 'UT PAYROLL:')}
        rows={model.utPayroll}
        model={model}
        analysisHeaders={analysisHeaders}
        onManualEdit={props.onManualEdit}
        onManualAdd={props.onManualAdd}
        onManualDelete={props.onManualDelete}
      />
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
  emptyAnalysis,
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
  hourCells: (byDate: Record<string, number>, cls?: string, border?: CSSProperties) => JSX.Element[];
  emptyAnalysis: JSX.Element;
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
        {hourCells(cust.byDate, 'xcell', hb)}
        <td className="xcell text-right font-bold" style={hb}>
          {fmtHoursFixed(cust.total)}
        </td>
        <td className="xcell" style={hb} />
        {emptyAnalysis}
      </tr>
      {custOpen &&
        cust.jobs.map((job) => {
          const jobOpen = !collapsedJob.has(job.id);
          const style = bandStyle(job.colorHex);
          const a = job.analysis;
          return (
            <JobBlock
              key={job.id}
              job={job}
              jobOpen={jobOpen}
              onToggleJob={() => onToggleJob(job.id)}
              style={style}
              a={a}
              hourCells={hourCells}
              onEditJob={onEditJob}
              onEditQbHours={onEditQbHours}
              dateIsos={dateIsos}
              hb={hb}
            />
          );
        })}
    </>
  );
}

function JobBlock({
  job,
  jobOpen,
  onToggleJob,
  style,
  a,
  hourCells,
  onEditJob,
  onEditQbHours,
  dateIsos,
  hb,
}: {
  job: PivotByJobModel['customers'][number]['jobs'][number];
  jobOpen: boolean;
  onToggleJob: () => void;
  style: CSSProperties | undefined;
  a: PivotByJobModel['customers'][number]['jobs'][number]['analysis'];
  hourCells: (byDate: Record<string, number>, cls?: string, border?: CSSProperties) => JSX.Element[];
  onEditJob: PivotByJobTableProps['onEditJob'];
  onEditQbHours: PivotByJobTableProps['onEditQbHours'];
  dateIsos: string[];
  hb: CSSProperties;
}): JSX.Element {
  const rowStyle = { ...style, ...hb };
  return (
    <>
      <tr id={`pbj-${job.id}`} style={rowStyle}>
        <td className="xcell pl-4 font-bold" style={hb}>
          <span className="inline-flex items-center">
            <Toggle open={jobOpen} onClick={onToggleJob} label={job.jobcode2} />
            {job.jobcode2}
          </span>
        </td>
        {hourCells(job.byDate, 'xcell', hb)}
        <td className="xcell text-right font-bold" style={hb}>
          {fmtHoursFixed(job.total)}
        </td>
        <td className="xcell" style={hb} />
        <td className="xcell text-center font-bold">
          <EditCell
            value={a.apKey ?? ''}
            onCommit={(t) => onEditJob(job.id, { apKey: t.trim() === '' ? null : t.trim().toUpperCase() })}
            className="text-center"
          />
        </td>
        <td className={`xcell text-center font-bold ${a.onCalendar === 'NOT ON CALENDAR' ? 'text-red-700' : ''}`}>
          {a.onCalendar}
        </td>
        <td className="xcell text-center font-bold">{a.estimator}</td>
        <td className="xcell text-center font-bold" title={a.crewLeadResolution?.raw ? `From "(${a.crewLeadResolution.raw})"` : undefined}>
          <EditCell value={a.crewLead} onCommit={(t) => onEditJob(job.id, { crewLead: t })} className="text-center" placeholder="—" />
        </td>
        <td className="xcell text-center font-bold" title={a.woOverrideSource ?? undefined}>
          {a.hourly ? (
            '-'
          ) : (
            <EditCell
              value={a.woOverride !== null ? `$${fmtAccounting(a.woOverride)}` : ''}
              onCommit={(t) => onEditJob(job.id, { woOverride: parseMoneyInput(t) })}
              className="text-center"
              placeholder="—"
            />
          )}
        </td>
        <td className="xcell text-center font-bold" title={a.woPctSource ?? undefined}>
          {a.hourly ? (
            '-'
          ) : (
            <EditCell
              value={a.woPct !== null ? `${Math.round(a.woPct * 100)}%` : ''}
              onCommit={(t) => {
                const p = parsePercentInput(t);
                onEditJob(job.id, { woPct: typeof p === 'number' ? p : null });
              }}
              className="text-center"
            />
          )}
        </td>
        <td className="xcell text-center font-bold" title="Mirrors the TimeNavi Export grid — edit it there">
          {fmtPercent(a.percent)}
        </td>
        <td className="xcell font-bold" style={{ color: style && isDarkColor(job.colorHex ?? 'FFFFFF') ? '#FFC7C7' : '#C00000' }}>
          <EditCell value={a.notes} onCommit={(t) => onEditJob(job.id, { notes: t })} />
        </td>
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
            <td className="xcell" colSpan={8} />
          </tr>
        ))}
    </>
  );
}

/**
 * Round 5 §3.2 — one black header row (mirrors row 1), then customer+job pairs.
 * No day-name / Sum of hours / Row Labels artifact rows.
 */
function ManualSection({
  section,
  title,
  rows,
  model,
  analysisHeaders,
  onManualEdit,
  onManualAdd,
  onManualDelete,
}: {
  section: 'balances' | 'ut';
  title: string;
  rows: ManualSectionRow[];
  model: PivotByJobModel;
  analysisHeaders: (sticky?: boolean) => JSX.Element;
} & Pick<PivotByJobTableProps, 'onManualEdit' | 'onManualAdd' | 'onManualDelete'>): JSX.Element {
  const dateCount = model.dateIsos.length;
  return (
    <table className="mt-4 w-max min-w-full border-collapse bg-white text-xs">
      <thead>
        <tr>
          <th className="xcell2 text-left text-base font-bold" style={BLACK}>
            {title}
          </th>
          {model.dates.map((d) => (
            <th key={d.getTime()} className="xcell2 text-center font-bold" style={BLACK}>
              {fmtMonthDay(d)}
            </th>
          ))}
          {model.dateIsos.length > model.dates.length && <th className="xcell2" style={BLACK} />}
          <th className="xcell2 text-center font-bold" style={BLACK}>
            Total hrs
          </th>
          <th className="xcell2" style={BLACK} />
          {analysisHeaders()}
          <th className="xcell2" style={BLACK} />
        </tr>
      </thead>
      <tbody>
        {rows.length === 0 && (
          <tr>
            <td className="xcell text-neutral-400" colSpan={1 + dateCount + 2 + 8 + 1}>
              No rows.{' '}
              {section === 'balances'
                ? 'Partially-paid NV jobs from last week (slot 7) seed here.'
                : 'Utah-calendar events and last week\u2019s open UT rows seed here.'}
            </td>
          </tr>
        )}
        {rows.map((r) => {
          const style = bandStyle(r.colorHex);
          return (
            <ManualPair
              key={r.id}
              section={section}
              r={r}
              style={style}
              dateCount={dateCount}
              onManualEdit={onManualEdit}
              onManualDelete={onManualDelete}
            />
          );
        })}
        <tr>
          <td className="xcell" colSpan={1 + dateCount + 2 + 8 + 1}>
            <button type="button" className="btn text-xs" onClick={() => onManualAdd(section)}>
              + Add row
            </button>
          </td>
        </tr>
      </tbody>
    </table>
  );
}

function ManualPair({
  section,
  r,
  style,
  dateCount,
  onManualEdit,
  onManualDelete,
}: {
  section: 'balances' | 'ut';
  r: ManualSectionRow;
  style: CSSProperties | undefined;
  dateCount: number;
  onManualEdit: PivotByJobTableProps['onManualEdit'];
  onManualDelete: PivotByJobTableProps['onManualDelete'];
}): JSX.Element {
  const blankHours = Array.from({ length: dateCount + 2 }, (_, i) => <td key={i} className="xcell" />);
  return (
    <>
      <tr style={style}>
        <td className="xcell font-bold">
          <EditCell value={r.customer} onCommit={(t) => onManualEdit(section, r.id, { customer: t })} />
        </td>
        {blankHours}
        <td className="xcell" colSpan={8} />
        <td className="xcell" />
      </tr>
      <tr style={style}>
        <td className="xcell pl-4 font-bold">
          <EditCell value={r.job} onCommit={(t) => onManualEdit(section, r.id, { job: t })} />
        </td>
        {blankHours}
        <td className="xcell text-center">
          <EditCell
            value={r.apKey ?? ''}
            onCommit={(t) => onManualEdit(section, r.id, { apKey: t.trim() || null })}
            className="text-center"
          />
        </td>
        <td className="xcell text-center" />
        <td className="xcell text-center">
          <EditCell value={r.estimator} onCommit={(t) => onManualEdit(section, r.id, { estimator: t })} className="text-center" />
        </td>
        <td className="xcell text-center">
          <EditCell value={r.crewLead} onCommit={(t) => onManualEdit(section, r.id, { crewLead: t })} className="text-center" />
        </td>
        <td className="xcell text-center">
          <EditCell
            value={r.woOverride !== null ? `$${fmtAccounting(r.woOverride)}` : ''}
            onCommit={(t) => onManualEdit(section, r.id, { woOverride: parseMoneyInput(t) })}
            className="text-center"
          />
        </td>
        <td className="xcell text-center">
          <EditCell
            value={r.woPct !== null ? `${Math.round(r.woPct * 100)}%` : ''}
            onCommit={(t) => {
              const p = parsePercentInput(t);
              onManualEdit(section, r.id, { woPct: typeof p === 'number' ? p : null });
            }}
            className="text-center"
          />
        </td>
        <td className="xcell text-center">
          <EditCell
            value={fmtPercent(r.percent)}
            onCommit={(t) => onManualEdit(section, r.id, { percent: parsePercentInput(t) })}
            className="text-center"
          />
        </td>
        <td className="xcell" style={{ color: '#C00000' }}>
          <EditCell value={r.notes} onCommit={(t) => onManualEdit(section, r.id, { notes: t })} />
        </td>
        <td className="xcell text-center">
          <button
            type="button"
            className="text-neutral-400 hover:text-red-600"
            onClick={() => onManualDelete(section, r.id)}
            aria-label="Delete row"
          >
            🗑
          </button>
        </td>
      </tr>
    </>
  );
}
