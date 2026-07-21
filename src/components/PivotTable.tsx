import { useEffect, useRef, useState, type CSSProperties, type Dispatch, type SetStateAction } from 'react';
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

function CollapseToggle({
  open,
  onClick,
  label,
  light,
}: {
  open: boolean;
  onClick: () => void;
  label: string;
  /** White chrome on the black crew band. */
  light?: boolean;
}): JSX.Element {
  return (
    <button
      type="button"
      className={`mr-1 inline-flex h-4 w-4 shrink-0 items-center justify-center rounded-xs border text-[10px] leading-none ${
        light
          ? 'border-neutral-300 bg-neutral-800 text-white hover:bg-neutral-700'
          : 'border-neutral-500 bg-white text-neutral-800 hover:bg-neutral-100'
      }`}
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

// ---------------------------------------------------------------------------
// §2.1/§2.2 — the Excel-style filter popover (Column Labels / Row Labels)
// ---------------------------------------------------------------------------

interface FilterPopoverProps {
  items: string[];
  /** Optional display labels (e.g. painter "Fname Lname" for a username key). */
  labels?: ReadonlyMap<string, string>;
  hidden: ReadonlySet<string>;
  sort: SortDir;
  onSort: (dir: SortDir) => void;
  onToggle: (item: string, checked: boolean) => void;
  onSetAll: (checked: boolean) => void;
  onClear: () => void;
  onClose: () => void;
  field?: RowField;
  onFieldChange?: (f: RowField) => void;
}

function FilterPopover(props: FilterPopoverProps): JSX.Element {
  const { items, labels, hidden, onToggle, onSetAll, onClear, onClose } = props;
  const [search, setSearch] = useState('');
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const onDown = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [onClose]);

  const labelOf = (item: string): string => labels?.get(item) ?? item;
  const visibleItems = items.filter((i) => labelOf(i).toLowerCase().includes(search.trim().toLowerCase()));
  const allChecked = items.length > 0 && items.every((i) => !hidden.has(i));

  return (
    <div
      ref={ref}
      className="absolute left-0 top-full z-40 mt-1 w-72 rounded border border-neutral-300 bg-white p-2 text-left text-xs font-normal text-neutral-900 shadow-lg"
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
            <option value="username">painter</option>
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
            <span className="truncate" title={labelOf(item)}>
              {labelOf(item)}
            </span>
          </label>
        ))}
        {visibleItems.length === 0 && <div className="px-1 py-1 text-neutral-400">No matches.</div>}
      </div>
    </div>
  );
}

/** The small ▼ filter button on Column Labels (E3) / Row Labels (A4). */
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
  // View-only collapse — never affects totals, export, or other tabs.
  const [collapsedCrew, setCollapsedCrew] = useState<ReadonlySet<string>>(new Set());
  const [collapsedEmp, setCollapsedEmp] = useState<ReadonlySet<string>>(new Set());
  const [collapsedJob, setCollapsedJob] = useState<ReadonlySet<string>>(new Set());

  const toggleKey = (
    setter: Dispatch<SetStateAction<ReadonlySet<string>>>,
    key: string,
  ): void => {
    setter((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const estCols = pivot.estimators;
  const totalCols = 4 + estCols.length + 1;
  const hoursIdx = Math.floor((estCols.length - 1) / 2);

  const allEstimators = fullPivot.estimators;
  const allGroups = fullPivot.crews.map((c) => c.crew);
  const allUsernames = [
    ...new Set(fullPivot.crews.flatMap((c) => c.employees.map((e) => e.username))),
  ].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
  const painterLabels = new Map<string, string>();
  for (const c of fullPivot.crews) {
    for (const e of c.employees) {
      if (!painterLabels.has(e.username)) {
        const name = `${e.fname} ${e.lname}`.trim();
        painterLabels.set(e.username, name || e.username);
      }
    }
  }

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
            {/* Estimator column filter — Excel E3 "Column Labels". */}
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
          <tr>
            {/* Painter / crew row filter — Excel A4 "Row Labels". */}
            <th className="xcell2 sticky top-0 z-10 min-w-[220px] text-left" style={HEADER_GRAY}>
              <span className="relative inline-flex items-center">
                Row Labels
                <FilterButton
                  active={view.hiddenGroups.length > 0 || view.hiddenUsernames.length > 0 || view.rowSort !== 'none'}
                  onClick={() => setOpenPopover(openPopover === 'row' ? null : 'row')}
                />
                {openPopover === 'row' && (
                  <FilterPopover
                    items={view.rowField === 'group' ? allGroups : allUsernames}
                    labels={view.rowField === 'username' ? painterLabels : undefined}
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
          {pivot.crews.map((crew) => {
            const crewOpen = !collapsedCrew.has(crew.crew);
            return (
              <CrewRows
                key={crew.crew}
                crew={crew}
                crewOpen={crewOpen}
                onToggleCrew={() => toggleKey(setCollapsedCrew, crew.crew)}
                collapsedEmp={collapsedEmp}
                onToggleEmp={(key) => toggleKey(setCollapsedEmp, key)}
                collapsedJob={collapsedJob}
                onToggleJob={(key) => toggleKey(setCollapsedJob, key)}
                estCols={estCols}
                totalCols={totalCols}
                colorByEstimator={colorByEstimator}
                hoursCells={hoursCells}
              />
            );
          })}
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
  crewOpen,
  onToggleCrew,
  collapsedEmp,
  onToggleEmp,
  collapsedJob,
  onToggleJob,
  estCols,
  totalCols,
  colorByEstimator,
  hoursCells,
}: {
  crew: PivotModel['crews'][number];
  crewOpen: boolean;
  onToggleCrew: () => void;
  collapsedEmp: ReadonlySet<string>;
  onToggleEmp: (key: string) => void;
  collapsedJob: ReadonlySet<string>;
  onToggleJob: (key: string) => void;
  estCols: string[];
  totalCols: number;
  colorByEstimator: ReadonlyMap<string, string>;
  hoursCells: (byEstimator: Record<string, number>, cls?: string) => JSX.Element[];
}): JSX.Element {
  return (
    <>
      <tr>
        <td className="xcell" style={BLACK_BAND} colSpan={totalCols}>
          <span className="inline-flex items-center">
            <CollapseToggle open={crewOpen} onClick={onToggleCrew} label={crew.crew} light />
            {crew.crew}
          </span>
        </td>
      </tr>
      {crewOpen &&
        crew.employees.map((emp) => {
          const empKey = `${crew.crew}\u0000${emp.username}`;
          const empOpen = !collapsedEmp.has(empKey);
          return (
            <EmployeeRows
              key={empKey}
              emp={emp}
              empKey={empKey}
              empOpen={empOpen}
              onToggleEmp={() => onToggleEmp(empKey)}
              collapsedJob={collapsedJob}
              onToggleJob={onToggleJob}
              estCols={estCols}
              colorByEstimator={colorByEstimator}
              hoursCells={hoursCells}
            />
          );
        })}
    </>
  );
}

function EmployeeRows({
  emp,
  empKey,
  empOpen,
  onToggleEmp,
  collapsedJob,
  onToggleJob,
  estCols,
  colorByEstimator,
  hoursCells,
}: {
  emp: PivotModel['crews'][number]['employees'][number];
  empKey: string;
  empOpen: boolean;
  onToggleEmp: () => void;
  collapsedJob: ReadonlySet<string>;
  onToggleJob: (key: string) => void;
  estCols: string[];
  colorByEstimator: ReadonlyMap<string, string>;
  hoursCells: (byEstimator: Record<string, number>, cls?: string) => JSX.Element[];
}): JSX.Element {
  const emptyBand = estCols.map((e) => (
    <td key={e} className="xcell" style={bandStyle(colorByEstimator.get(e))} />
  ));
  const totalCls = 'xcell font-bold';
  const totalBorderCls = 'xcell border-t-2 border-t-neutral-400 font-bold';
  return (
    <>
      <tr>
        <td className="xcell font-bold">
          <span className="inline-flex items-center">
            <CollapseToggle open={empOpen} onClick={onToggleEmp} label={emp.username} />
            {emp.username}
          </span>
        </td>
        <td className="xcell font-bold">
          <span className="inline-flex items-center">
            <CollapseToggle open={empOpen} onClick={onToggleEmp} label={emp.fname || 'fname'} />
            {emp.fname}
          </span>
        </td>
        <td className="xcell font-bold">
          <span className="inline-flex items-center">
            <CollapseToggle open={empOpen} onClick={onToggleEmp} label={emp.lname || 'lname'} />
            {emp.lname}
          </span>
        </td>
        <td className="xcell" />
        {emptyBand}
        <td className="xcell" style={GT_BAND} />
      </tr>
      {empOpen &&
        emp.jobs.map((job) => {
          const jobKey = `${empKey}\u0000${job.customer}\u0000${job.job}`;
          const jobOpen = !collapsedJob.has(jobKey);
          if (!jobOpen) {
            return (
              <tr key={jobKey}>
                <td className="xcell" />
                <td className="xcell" />
                <td className="xcell font-bold">
                  <span className="inline-flex items-center">
                    <CollapseToggle open={false} onClick={() => onToggleJob(jobKey)} label={job.customer} />
                    {job.customer}
                  </span>
                </td>
                <td className="xcell">
                  <span className="inline-flex items-center">
                    <CollapseToggle open={false} onClick={() => onToggleJob(jobKey)} label={job.job} />
                    {job.job}
                  </span>
                </td>
                {estCols.map((e) => (
                  <td key={e} className="xcell" style={bandStyle(colorByEstimator.get(e))} />
                ))}
                <td className="xcell" style={GT_BAND} />
              </tr>
            );
          }
          return (
            <tr key={jobKey}>
              <td className="xcell" />
              <td className="xcell" />
              <td className="xcell font-bold">
                <span className="inline-flex items-center">
                  <CollapseToggle open={jobOpen} onClick={() => onToggleJob(jobKey)} label={job.customer} />
                  {job.customer}
                </span>
              </td>
              <td className="xcell">
                <span className="inline-flex items-center">
                  <CollapseToggle open={jobOpen} onClick={() => onToggleJob(jobKey)} label={job.job} />
                  {job.job}
                </span>
              </td>
              {hoursCells(job.byEstimator)}
              <td className="xcell text-right" style={GT_BAND}>
                {fmtHoursFixed(job.total)}
              </td>
            </tr>
          );
        })}
      {empOpen && (
        <tr>
          <td className={totalCls} />
          <td className={totalCls}>{emp.fname} Total</td>
          <td className={totalBorderCls} />
          <td className={totalBorderCls} />
          {hoursCells(emp.byEstimator, totalBorderCls)}
          <td
            className={`${totalBorderCls} text-right`}
            style={emp.overtime ? { ...GT_BAND, backgroundColor: OVERTIME_BG } : GT_BAND}
            title={emp.overtime ? 'Overtime: grand total exceeds 40 hours (on-screen flag only)' : undefined}
          >
            {fmtHoursFixed(emp.total)}
          </td>
        </tr>
      )}
    </>
  );
}
