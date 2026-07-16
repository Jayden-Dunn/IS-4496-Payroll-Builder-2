import { useState } from 'react';
import { DEFAULT_LEGEND, DEFAULT_PIVOT_COLUMNS } from '../lib/settings';
import type { LegendEntry } from '../lib/types';

interface SettingsModalProps {
  legend: LegendEntry[];
  pivotColumnOrder: string[];
  onSave: (legend: LegendEntry[], pivotColumnOrder: string[]) => void;
  onClose: () => void;
}

const HEX_RE = /^[0-9A-Fa-f]{6}$/;

function moveItem<T>(list: T[], i: number, dir: -1 | 1): T[] {
  const j = i + dir;
  if (j < 0 || j >= list.length) return list;
  const next = [...list];
  const tmp = next[i];
  next[i] = next[j];
  next[j] = tmp;
  return next;
}

export function SettingsModal({ legend, pivotColumnOrder, onSave, onClose }: SettingsModalProps): JSX.Element {
  const [entries, setEntries] = useState<LegendEntry[]>(legend.map((l) => ({ ...l })));
  const [pivotCols, setPivotCols] = useState<string[]>([...pivotColumnOrder]);
  const [newPivotCol, setNewPivotCol] = useState('');

  const updateEntry = (i: number, patch: Partial<LegendEntry>): void => {
    setEntries((prev) => prev.map((e, j) => (j === i ? { ...e, ...patch } : e)));
  };

  const cleanedLegend = entries
    .map((e) => ({ estimator: e.estimator.trim().toUpperCase(), color: e.color.trim().toUpperCase() }))
    .filter((e) => e.estimator !== '');
  const legendValid =
    cleanedLegend.length > 0 &&
    cleanedLegend.every((e) => HEX_RE.test(e.color)) &&
    new Set(cleanedLegend.map((e) => e.estimator)).size === cleanedLegend.length;

  const cleanedPivot = pivotCols.map((e) => e.trim().toUpperCase()).filter((e) => e !== '');
  const pivotValid = cleanedPivot.length > 0 && new Set(cleanedPivot).size === cleanedPivot.length;

  const addPivotCol = (): void => {
    const t = newPivotCol.trim().toUpperCase();
    if (t && !pivotCols.includes(t)) setPivotCols((prev) => [...prev, t]);
    setNewPivotCol('');
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div
        className="w-full max-w-md rounded-lg bg-white shadow-xl"
        onClick={(e) => e.stopPropagation()}
        role="dialog"
        aria-label="Settings"
      >
        <header className="flex items-center justify-between border-b border-neutral-200 px-4 py-3">
          <h2 className="text-sm font-bold">Settings</h2>
          <button type="button" className="text-neutral-400 hover:text-neutral-700" onClick={onClose} aria-label="Close">
            ✕
          </button>
        </header>
        <div className="max-h-[65vh] overflow-y-auto p-4">
          <h3 className="mb-1 text-xs font-bold uppercase tracking-wide text-neutral-600">
            Estimator color legend &amp; row-sort order
          </h3>
          <p className="mb-3 text-[11px] text-neutral-500">
            This order drives row sorting on the TimeNavi Export. Colors are the row fills estimators use to find
            their jobs.
          </p>
          <div className="flex flex-col gap-1.5">
            {entries.map((e, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <div className="flex flex-col">
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === 0} onClick={() => setEntries((prev) => moveItem(prev, i, -1))} aria-label="Move up">
                    ▲
                  </button>
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === entries.length - 1} onClick={() => setEntries((prev) => moveItem(prev, i, 1))} aria-label="Move down">
                    ▼
                  </button>
                </div>
                <input
                  className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-xs font-semibold uppercase"
                  value={e.estimator}
                  maxLength={6}
                  placeholder="Init."
                  onChange={(ev) => updateEntry(i, { estimator: ev.target.value.toUpperCase() })}
                />
                <input
                  type="color"
                  className="h-7 w-9 cursor-pointer rounded border border-neutral-300"
                  value={HEX_RE.test(e.color) ? `#${e.color}` : '#cccccc'}
                  onChange={(ev) => updateEntry(i, { color: ev.target.value.slice(1).toUpperCase() })}
                />
                <input
                  className={`w-20 rounded border px-1.5 py-1 font-mono text-xs uppercase ${HEX_RE.test(e.color) ? 'border-neutral-300' : 'border-red-400 bg-red-50'}`}
                  value={e.color}
                  maxLength={6}
                  onChange={(ev) => updateEntry(i, { color: ev.target.value })}
                />
                <div className="h-6 flex-1 rounded border border-neutral-200" style={{ backgroundColor: HEX_RE.test(e.color) ? `#${e.color}` : undefined }} />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setEntries((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove estimator">
                  🗑
                </button>
              </div>
            ))}
          </div>
          <div className="mt-3 flex gap-2">
            <button type="button" className="btn text-xs" onClick={() => setEntries((prev) => [...prev, { estimator: '', color: 'CCCCCC' }])}>
              + Add estimator
            </button>
            <button type="button" className="btn text-xs" onClick={() => setEntries(DEFAULT_LEGEND.map((l) => ({ ...l })))}>
              Reset to defaults
            </button>
          </div>

          <h3 className="mb-1 mt-6 text-xs font-bold uppercase tracking-wide text-neutral-600">
            Pivot estimator column order
          </h3>
          <p className="mb-3 text-[11px] text-neutral-500">
            Column order on the Pivot by Crew tab and sheet (separate from the row-sort order above). Only estimators
            present in the data are rendered; new ones found in the data are appended automatically.
          </p>
          <div className="flex flex-col gap-1.5">
            {pivotCols.map((col, i) => (
              <div key={`${col}-${i}`} className="flex items-center gap-1.5">
                <div className="flex flex-col">
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === 0} onClick={() => setPivotCols((prev) => moveItem(prev, i, -1))} aria-label="Move up">
                    ▲
                  </button>
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === pivotCols.length - 1} onClick={() => setPivotCols((prev) => moveItem(prev, i, 1))} aria-label="Move down">
                    ▼
                  </button>
                </div>
                <span className="w-16 rounded border border-neutral-200 bg-neutral-50 px-1.5 py-1 text-xs font-semibold">{col}</span>
                <span className="flex-1 text-[10px] text-neutral-400">{i === 0 ? 'first column' : ''}</span>
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setPivotCols((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove column">
                  🗑
                </button>
              </div>
            ))}
          </div>
          <div className="mt-3 flex items-center gap-2">
            <input
              className="w-20 rounded border border-neutral-300 px-1.5 py-1 text-xs font-semibold uppercase"
              value={newPivotCol}
              maxLength={6}
              placeholder="Init."
              onChange={(e) => setNewPivotCol(e.target.value.toUpperCase())}
              onKeyDown={(e) => {
                if (e.key === 'Enter') addPivotCol();
              }}
            />
            <button type="button" className="btn text-xs" onClick={addPivotCol}>
              + Add
            </button>
            <button type="button" className="btn text-xs" onClick={() => setPivotCols([...DEFAULT_PIVOT_COLUMNS])}>
              Reset to defaults
            </button>
          </div>
        </div>
        <footer className="flex justify-end gap-2 border-t border-neutral-200 px-4 py-3">
          <button type="button" className="btn" onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            disabled={!legendValid || !pivotValid}
            title={
              legendValid && pivotValid
                ? undefined
                : 'Each estimator needs unique initials; colors need a 6-digit hex; pivot columns must be unique'
            }
            onClick={() => onSave(cleanedLegend, cleanedPivot)}
          >
            Save
          </button>
        </footer>
      </div>
    </div>
  );
}
