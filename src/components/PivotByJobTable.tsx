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

function bandStyle(hex: string | null): CSSProperties | undefined {
  if (!hex) return undefined;
  return { backgroundColor: `#${hex}`, color: isDarkColor(hex) ? '#fff' : '#000' };
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
  // Columns: A · dates · Total · QB Hrs · AP# · On Cal? · Est · Lead · WO$ · WO% · %Compl · Notes
  const totalCols = 1 + dateCount + 2 + 8;

  const hourCells = (byDate: Record<string, number>, cls = 'xcell'): JSX.Element[] =>
    model.dateIsos.map((iso) => (
      <td key={iso || 'blank'} className={`${cls} text-right`}>
        {byDate[iso] !== undefined ? fmtHoursFixed(byDate[iso]) : ''}
      </td>
    ));

  const emptyAnalysis = <>{Array.from({ length: 8 }, (_, i) => <td key={i} className="xcell" />)}</>;

  return (
    <div className="p-2">
      <table className="w-max min-w-full border-collapse bg-white text-xs">
        <thead>
          {/* Row 1 — black title band: title, dates, headers. */}
          <tr>
            <th className="xcell2 text-left text-base font-bold" style={BLACK}>
              {model.title}
            </th>
            {model.dates.map((d) => (
              <th key={d.getTime()} className="xcell2 text-right" style={BLACK}>
                {fmtMonthDay(d)}
              </th>
            ))}
            {model.dateIsos.length > model.dates.length && <th className="xcell2" style={BLACK} />}
            <th className="xcell2 text-right" style={BLACK}>Total hrs</th>
            <th className="xcell2" style={BLACK} title="QB Hrs override (optional)" />
            <th className="xcell2 text-center" style={BLACK}>AP#</th>
            <th className="xcell2 text-center" style={BLACK}>On Calendar?</th>
            <th className="xcell2 text-center" style={BLACK}>Estimator</th>
            <th className="xcell2 text-center" style={BLACK}>Crew Lead</th>
            <th className="xcell2 text-center" style={BLACK}>WO Override</th>
            <th className="xcell2 text-center" style={BLACK}>WO%</th>
            <th className="xcell2 text-center" style={{ ...BLACK, color: '#FF0000' }}>% of Completion</th>
            <th className="xcell2 text-left" style={{ ...BLACK, color: '#FF6A6A' }}>Notes</th>
          </tr>
          {/* Row 2 — computed weekday names (Sunday included, unlike the hand-typed reference). */}
          <tr>
            <th className="xcell2" style={BLACK} />
            {model.dates.map((d) => (
              <th key={d.getTime()} className="xcell2 text-right font-normal" style={BLACK}>
                {weekdayShort(d)}
              </th>
            ))}
            {model.dateIsos.length > model.dates.length && <th className="xcell2 font-normal" style={BLACK}>(blank)</th>}
            <th className="xcell2" style={BLACK} colSpan={10} />
          </tr>
          {/* Row 4-style label row. */}
          <tr>
            <th className="xcell2 sticky top-0 z-10 min-w-[260px] text-left" style={GRAY}>Row Labels</th>
            {model.dateIsos.map((iso, i) => (
              <th key={iso || 'blank'} className="xcell2 sticky top-0 z-10 min-w-[64px] text-right" style={GRAY}>
                {i < model.dates.length ? fmtMonthDay(model.dates[i]) : '(blank)'}
              </th>
            ))}
            <th className="xcell2 sticky top-0 z-10 min-w-[70px] text-right" style={GRAY}>Grand Total</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[56px] text-right" style={GRAY} title="Optional QuickBooks-approved hours override">QB Hrs</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[76px] text-center" style={GRAY}>AP#</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[120px] text-center" style={GRAY}>On Calendar?</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[64px] text-center" style={GRAY}>Estimator</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[220px] text-center" style={GRAY}>Crew Lead</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[110px] text-center" style={GRAY}>WO Override</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[64px] text-center" style={GRAY}>WO%</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[76px] text-center" style={GRAY}>% of Compl.</th>
            <th className="xcell2 sticky top-0 z-10 min-w-[160px] text-left" style={GRAY}>Notes</th>
          </tr>
        </thead>
        <tbody>
          {model.customers.map((cust) => (
            <CustomerRows key={cust.customer} cust={cust} hourCells={hourCells} emptyAnalysis={emptyAnalysis} onEditJob={onEditJob} onEditQbHours={onEditQbHours} dateIsos={model.dateIsos} />
          ))}
          {/* Grand Total */}
          <tr>
            <td className="xcell font-bold" style={GRAY}>Grand Total</td>
            {model.dateIsos.map((iso) => (
              <td key={iso || 'blank'} className="xcell text-right font-bold" style={GRAY}>
                {model.grandByDate[iso] !== undefined ? fmtHoursFixed(model.grandByDate[iso]) : ''}
              </td>
            ))}
            <td className="xcell text-right font-bold" style={GRAY}>{fmtHoursFixed(model.grandTotal)}</td>
            <td className="xcell" style={GRAY} colSpan={9} />
          </tr>
        </tbody>
      </table>

      <ManualSection
        section="balances"
        title={model.title.replace('NV PAYROLL:', 'NV PAYROLL (BALANCES):')}
        rows={model.balances}
        totalCols={totalCols}
        {...props}
      />
      <ManualSection
        section="ut"
        title={model.title.replace('NV PAYROLL:', 'UT PAYROLL:')}
        rows={model.utPayroll}
        totalCols={totalCols}
        {...props}
      />
      <p className="mt-2 text-[11px] italic text-neutral-500">
        Balances &amp; UT sections auto-position below the pivot — the Excel move-refresh-move workaround is gone.
      </p>
    </div>
  );
}

function CustomerRows({
  cust,
  hourCells,
  emptyAnalysis,
  onEditJob,
  onEditQbHours,
  dateIsos,
}: {
  cust: PivotByJobModel['customers'][number];
  hourCells: (byDate: Record<string, number>, cls?: string) => JSX.Element[];
  emptyAnalysis: JSX.Element;
  onEditJob: PivotByJobTableProps['onEditJob'];
  onEditQbHours: PivotByJobTableProps['onEditQbHours'];
  dateIsos: string[];
}): JSX.Element {
  const custStyle = bandStyle(cust.colorHex);
  return (
    <>
      <tr style={custStyle}>
        <td className="xcell font-bold">{cust.customer}</td>
        {hourCells(cust.byDate)}
        <td className="xcell text-right font-bold">{fmtHoursFixed(cust.total)}</td>
        <td className="xcell" />
        {emptyAnalysis}
      </tr>
      {cust.jobs.map((job) => {
        const style = bandStyle(job.colorHex);
        const a = job.analysis;
        return (
          <JobRows key={job.id} job={job} style={style} a={a} hourCells={hourCells} onEditJob={onEditJob} onEditQbHours={onEditQbHours} dateIsos={dateIsos} />
        );
      })}
    </>
  );
}

function JobRows({
  job,
  style,
  a,
  hourCells,
  onEditJob,
  onEditQbHours,
  dateIsos,
}: {
  job: PivotByJobModel['customers'][number]['jobs'][number];
  style: CSSProperties | undefined;
  a: PivotByJobModel['customers'][number]['jobs'][number]['analysis'];
  hourCells: (byDate: Record<string, number>, cls?: string) => JSX.Element[];
  onEditJob: PivotByJobTableProps['onEditJob'];
  onEditQbHours: PivotByJobTableProps['onEditQbHours'];
  dateIsos: string[];
}): JSX.Element {
  return (
    <>
      <tr id={`pbj-${job.id}`} style={style}>
        <td className="xcell pl-4 font-bold">{job.jobcode2}</td>
        {hourCells(job.byDate)}
        <td className="xcell text-right font-bold">{fmtHoursFixed(job.total)}</td>
        <td className="xcell" />
        <td className="xcell text-center font-bold">
          <EditCell
            value={a.apKey ?? ''}
            onCommit={(t) => onEditJob(job.id, { apKey: t.trim() === '' ? null : t.trim().toUpperCase() })}
            className="text-center"
          />
        </td>
        <td className={`xcell text-center font-bold ${a.onCalendar === 'NOT ON CALENDAR' ? 'text-red-700' : ''}`}>{a.onCalendar}</td>
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
      {job.employees.map((emp) => (
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

function ManualSection({
  section,
  title,
  rows,
  totalCols,
  onManualEdit,
  onManualAdd,
  onManualDelete,
}: {
  section: 'balances' | 'ut';
  title: string;
  rows: ManualSectionRow[];
  totalCols: number;
} & Pick<PivotByJobTableProps, 'onManualEdit' | 'onManualAdd' | 'onManualDelete'>): JSX.Element {
  return (
    <table className="mt-4 w-max min-w-full border-collapse bg-white text-xs">
      <thead>
        <tr>
          <th className="xcell2 text-left text-sm font-bold" style={BLACK} colSpan={totalCols}>
            {title}
          </th>
        </tr>
        <tr>
          {['Customer', 'Job', 'AP#', 'Estimator', 'Crew Lead', 'WO Override', 'WO%', '% of Compl.', 'Notes', ''].map((h) => (
            <th key={h || 'del'} className="xcell2 text-left" style={GRAY}>
              {h}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.length === 0 && (
          <tr>
            <td className="xcell text-neutral-400" colSpan={10}>
              No rows. {section === 'balances' ? 'Partially-paid NV jobs from last week (slot 7) seed here.' : 'Utah-calendar events and last week\u2019s open UT rows seed here.'}
            </td>
          </tr>
        )}
        {rows.map((r) => {
          const style = bandStyle(r.colorHex);
          const cell = (field: keyof ManualSectionRow, className?: string): JSX.Element => (
            <EditCell
              value={String(r[field] ?? '')}
              onCommit={(t) => onManualEdit(section, r.id, { [field]: t } as Partial<ManualSectionRow>)}
              className={className}
            />
          );
          return (
            <tr key={r.id} style={style}>
              <td className="xcell min-w-[160px]">{cell('customer')}</td>
              <td className="xcell min-w-[260px]">{cell('job')}</td>
              <td className="xcell min-w-[76px] text-center">
                <EditCell value={r.apKey ?? ''} onCommit={(t) => onManualEdit(section, r.id, { apKey: t.trim() || null })} className="text-center" />
              </td>
              <td className="xcell min-w-[64px] text-center">{cell('estimator', 'text-center')}</td>
              <td className="xcell min-w-[200px] text-center">{cell('crewLead', 'text-center')}</td>
              <td className="xcell min-w-[100px] text-center">
                <EditCell
                  value={r.woOverride !== null ? `$${fmtAccounting(r.woOverride)}` : ''}
                  onCommit={(t) => onManualEdit(section, r.id, { woOverride: parseMoneyInput(t) })}
                  className="text-center"
                />
              </td>
              <td className="xcell min-w-[56px] text-center">
                <EditCell
                  value={r.woPct !== null ? `${Math.round(r.woPct * 100)}%` : ''}
                  onCommit={(t) => {
                    const p = parsePercentInput(t);
                    onManualEdit(section, r.id, { woPct: typeof p === 'number' ? p : null });
                  }}
                  className="text-center"
                />
              </td>
              <td className="xcell min-w-[64px] text-center">
                <EditCell
                  value={fmtPercent(r.percent)}
                  onCommit={(t) => onManualEdit(section, r.id, { percent: parsePercentInput(t) })}
                  className="text-center"
                />
              </td>
              <td className="xcell min-w-[160px]" style={{ color: '#C00000' }}>{cell('notes')}</td>
              <td className="xcell text-center">
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => onManualDelete(section, r.id)} aria-label="Delete row">
                  🗑
                </button>
              </td>
            </tr>
          );
        })}
        <tr>
          <td className="xcell" colSpan={10}>
            <button type="button" className="btn text-xs" onClick={() => onManualAdd(section)}>
              + Add row
            </button>
          </td>
        </tr>
      </tbody>
    </table>
  );
}
