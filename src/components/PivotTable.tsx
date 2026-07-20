import { useEffect, useRef, useState, type CSSProperties } from 'react';
import { fmtHoursFixed, isDarkColor } from '../lib/format';
import { isPivotViewActive, EMPTY_PIVOT_VIEW, type PivotViewState, type RowField, type SortDir } from '../lib/pivotView';
import type { PivotModel } from '../lib/types';

interface PivotTableProps {
  /** Filtered view model (what renders). */
  pivot: PivotModel;
  /** Unfiltered model — source for the filter checkbox lists. */
  fullPivot: PivotModel;
  colorByEstimator: ReadonlyMap<string, string>;
  view: PivotViewState;
  onViewChange: (view: PivotViewState) => void;
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

// ---------------------------------------------------------------------------
// §2.1/§2.2 — the Excel-style filter popover
// ---------------------------------------------------------------------------

interface FilterPopoverProps {
  items: string[]; // all candidate values (unfiltered)
  hidden: ReadonlySet<string>;
  sort: SortDir;
  onSort: (dir: SortDir) => void;
  onToggle: (item: string, checked: boolean) => void;
  onSetAll: (checked: boolean) => void;
  onClear: () => void;
  onClose: () => void;
  /** Row Labels variant: the "Select field" dropdown. */
  field?: RowField;
  onFieldChange?: (f: RowField) => void;
}

function FilterPopover(props: FilterPopoverProps): JSX.Element {
  const { items, hidden, onToggle, onSetAll, onClear, onClose } = props;
  const [search, setSearch] = useState('');
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const onDown = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [onClose]);

  const visibleItems = items.filter((i) => i.toLowerCase().includes(search.trim().toLowerCase()));
  const allChecked = items.every((i) => !hidden.has(i));

  return (
    <div
      ref={ref}
      className="absolute left-0 top-full z-30 mt-1 w-64 rounded border border-neutral-300 bg-white p-2 text-left text-xs font-normal text-neutral-900 shadow-lg"
      onClick={(e) => e.stopPropagation()}
    >
      {props.field !== undefined && props.onFieldChange && (
        <label className="mb-2 flex items-center gap-2">
          <span className="text-neutral-500">Select field:</span>
          <select
            className="flex-1 rounded border border-neutral-300 px-1 py-0.5"
            value={props.field}
            onChange={(e) => props.onFieldChange?.(e.target.value as RowField)}
          >
            <option value="group">group</option>
            <option value="username">username</option>
          </select>
        </label>
      )}
      <button type="button" className="block w-full rounded px-1 py-0.5 text-left hover:bg-neutral-100" onClick={() => props.onSort('asc')}>
        ↓ Sort A to Z{props.sort === 'asc' ? ' ✓' : ''}
      </button>
      <button type="button" className="block w-full rounded px-1 py-0.5 text-left hover:bg-neutral-100" onClick={() => props.onSort('desc')}>
        ↑ Sort Z to A{props.sort === 'desc' ? ' ✓' : ''}
      </button>
      <button type="button" className="mb-1 block w-full rounded px-1 py-0.5 text-left hover:bg-neutral-100" onClick={onClear}>
        ✕ Clear Filter
      </button>
      <input
        className="mb-1 w-full rounded border border-neutral-300 px-1.5 py-1"
        placeholder="Search"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />
      <div className="max-h-48 overflow-y-auto rounded border border-neutral-200 p-1">
        <label className="flex items-center gap-1.5 px-1 py-0.5 hover:bg-neutral-50">
          <input type="checkbox" checked={allChecked} onChange={(e) => onSetAll(e.target.checked)} />
          <span className="font-semibold">(Select All)</span>
        </label>
        {visibleItems.map((item) => (
          <label key={item} className="flex items-center gap-1.5 px-1 py-0.5 hover:bg-neutral-50">
            <input type="checkbox" checked={!hidden.has(item)} onChange={(e) => onToggle(item, e.target.checked)} />
            <span className="truncate">{item}</span>
          </label>
        ))}
        {visibleItems.length === 0 && <div className="px-1 py-1 text-neutral-400">No matches.</div>}
      </div>
    </div>
  );
}

/** The small ▼ filter button that sits on the Column/Row Labels cells. */
function FilterButton({ active, onClick }: { active: boolean; onClick: () => void }): JSX.Element {
  return (
    <button
      type="button"
      className={`ml-1 inline-flex h-4 w-4 items-center justify-center rounded-xs border text-[9px] leading-none ${
        active ? 'border-amber-400 bg-amber-200 text-black' : 'border-neutral-400 bg-white text-neutral-700'
      }`}
      onClick={(e) => {
        e.stopPropagation();
        onClick();
      }}
      title={active ? 'Filter active' : 'Filter'}
      aria-label="Filter"
    >
      ▼
    </button>
  );
}

export function PivotTable({ pivot, fullPivot, colorByEstimator, view, onViewChange }: PivotTableProps): JSX.Element {
  const [openPopover, setOpenPopover] = useState<'col' | 'row' | null>(null);

  const estCols = pivot.estimators;
  const totalCols = 4 + estCols.length + 1;
  // "Hours" sits over the middle of the estimator block (no merge).
  const hoursIdx = Math.floor((estCols.length - 1) / 2);

  const allEstimators = fullPivot.estimators;
  const allGroups = fullPivot.crews.map((c) => c.crew);
  const allUsernames = [
    ...new Set(fullPivot.crews.flatMap((c) => c.employees.map((e) => e.username))),
  ].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));

  const toggleList = (list: string[], item: string, checked: boolean): string[] =>
    checked ? list.filter((x) => x !== item) : list.includes(item) ? list : [...list, item];

  const filtered = isPivotViewActive(view);

  const hoursCells = (byEstimator: Record<string, number>, cls = 'xcell'): JSX.Element[] =>
    estCols.map((e) => (
      <td key={e} className={`${cls} text-center`} style={bandStyle(colorByEstimator.get(e))}>
        {byEstimator[e] !== undefined ? fmtHoursFixed(byEstimator[e]) : ''}
      </td>
    ));

  return (
    <div>
      {filtered && (
        <div className="flex items-center gap-2 border-b border-amber-300 bg-amber-50 px-3 py-1.5 text-xs">
          <span className="font-semibold text-amber-800">Filters active — this view only.</span>
          <span className="text-amber-700">Payroll math, exceptions and the default export are unaffected.</span>
          <button
            type="button"
            className="ml-auto rounded border border-amber-400 bg-white px-2 py-0.5 font-semibold text-amber-800 hover:bg-amber-100"
            onClick={() => onViewChange({ ...EMPTY_PIVOT_VIEW })}
          >
            ✕ Clear all filters
          </button>
        </div>
      )}
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
            {/* §2.1 — Column Labels filter button (E3 position). */}
            <th className="xcell2 relative text-left" style={BLACK_BAND} colSpan={estCols.length + 1}>
              Column Labels
              <FilterButton
                active={view.hiddenEstimators.length > 0 || view.columnSort !== 'none'}
                onClick={() => setOpenPopover(openPopover === 'col' ? null : 'col')}
              />
              {openPopover === 'col' && (
                <FilterPopover
                  items={allEstimators}
                  hidden={new Set(view.hiddenEstimators)}
                  sort={view.columnSort}
                  onSort={(dir) => onViewChange({ ...view, columnSort: view.columnSort === dir ? 'none' : dir })}
                  onToggle={(item, checked) =>
                    onViewChange({ ...view, hiddenEstimators: toggleList(view.hiddenEstimators, item, checked) })
                  }
                  onSetAll={(checked) =>
                    onViewChange({ ...view, hiddenEstimators: checked ? [] : [...allEstimators] })
                  }
                  onClear={() => onViewChange({ ...view, hiddenEstimators: [], columnSort: 'none' })}
                  onClose={() => setOpenPopover(null)}
                />
              )}
            </th>
          </tr>
          {/* Row 4: gray labels + palette-colored estimator headers. */}
          <tr>
            <th className="xcell2 sticky top-0 z-10 min-w-[220px] text-left" style={HEADER_GRAY}>
              <span className="relative inline-flex items-center">
                Row Labels
                {/* §2.2 — Row Labels filter button (A4 position). */}
                <FilterButton
                  active={view.hiddenGroups.length > 0 || view.hiddenUsernames.length > 0 || view.rowSort !== 'none'}
                  onClick={() => setOpenPopover(openPopover === 'row' ? null : 'row')}
                />
                {openPopover === 'row' && (
                  <FilterPopover
                    items={view.rowField === 'group' ? allGroups : allUsernames}
                    hidden={new Set(view.rowField === 'group' ? view.hiddenGroups : view.hiddenUsernames)}
                    sort={view.rowSort}
                    field={view.rowField}
                    onFieldChange={(f) => onViewChange({ ...view, rowField: f, rowSort: 'none' })}
                    onSort={(dir) => onViewChange({ ...view, rowSort: view.rowSort === dir ? 'none' : dir })}
                    onToggle={(item, checked) =>
                      view.rowField === 'group'
                        ? onViewChange({ ...view, hiddenGroups: toggleList(view.hiddenGroups, item, checked) })
                        : onViewChange({ ...view, hiddenUsernames: toggleList(view.hiddenUsernames, item, checked) })
                    }
                    onSetAll={(checked) =>
                      view.rowField === 'group'
                        ? onViewChange({ ...view, hiddenGroups: checked ? [] : [...allGroups] })
                        : onViewChange({ ...view, hiddenUsernames: checked ? [] : [...allUsernames] })
                    }
                    onClear={() =>
                      onViewChange({ ...view, hiddenGroups: [], hiddenUsernames: [], rowSort: 'none' })
                    }
                    onClose={() => setOpenPopover(null)}
                  />
                )}
              </span>
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
                className="xcell2 sticky top-0 z-10 min-w-[64px] text-center"
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
              <td key={e} className="xcell text-center font-bold" style={GRAND_GRAY}>
                {pivot.grandByEstimator[e] !== undefined ? fmtHoursFixed(pivot.grandByEstimator[e]) : ''}
              </td>
            ))}
            <td className="xcell text-right font-bold" style={GRAND_GRAY}>
              {fmtHoursFixed(pivot.grandTotal)}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
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
  // §2.4 — the thin separator sits on the TOP edge of the "X Total" row (the
  // line between the last job row and the total), from C through Grand Total.
  const totalCls = 'xcell font-bold';
  const totalBorderCls = 'xcell border-t-2 border-t-neutral-400 font-bold';
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
        <td className={totalBorderCls} />
        <td className={totalBorderCls} />
        {hoursCells(emp.byEstimator, totalBorderCls)}
        <td
          className={`${totalBorderCls} text-right`}
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
