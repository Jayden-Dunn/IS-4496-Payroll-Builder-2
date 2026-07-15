import { useState } from 'react';
import { DEFAULT_LEGEND } from '../lib/settings';
import type { LegendEntry } from '../lib/types';

interface SettingsModalProps {
  legend: LegendEntry[];
  onSave: (legend: LegendEntry[]) => void;
  onClose: () => void;
}

const HEX_RE = /^[0-9A-Fa-f]{6}$/;

export function SettingsModal({ legend, onSave, onClose }: SettingsModalProps): JSX.Element {
  const [entries, setEntries] = useState<LegendEntry[]>(legend.map((l) => ({ ...l })));

  const update = (i: number, patch: Partial<LegendEntry>): void => {
    setEntries((prev) => prev.map((e, j) => (j === i ? { ...e, ...patch } : e)));
  };
  const move = (i: number, dir: -1 | 1): void => {
    setEntries((prev) => {
      const j = i + dir;
      if (j < 0 || j >= prev.length) return prev;
      const next = [...prev];
      const tmp = next[i];
      next[i] = next[j];
      next[j] = tmp;
      return next;
    });
  };
  const remove = (i: number): void => {
    setEntries((prev) => prev.filter((_, j) => j !== i));
  };
  const add = (): void => {
    setEntries((prev) => [...prev, { estimator: '', color: 'CCCCCC' }]);
  };

  const cleaned = entries
    .map((e) => ({ estimator: e.estimator.trim().toUpperCase(), color: e.color.trim().toUpperCase() }))
    .filter((e) => e.estimator !== '');
  const valid =
    cleaned.length > 0 &&
    cleaned.every((e) => HEX_RE.test(e.color)) &&
    new Set(cleaned.map((e) => e.estimator)).size === cleaned.length;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div
        className="w-full max-w-md rounded-lg bg-white shadow-xl"
        onClick={(e) => e.stopPropagation()}
        role="dialog"
        aria-label="Settings"
      >
        <header className="flex items-center justify-between border-b border-neutral-200 px-4 py-3">
          <h2 className="text-sm font-bold">Estimator color legend &amp; order</h2>
          <button type="button" className="text-neutral-400 hover:text-neutral-700" onClick={onClose} aria-label="Close">
            ✕
          </button>
        </header>
        <div className="max-h-[60vh] overflow-y-auto p-4">
          <p className="mb-3 text-[11px] text-neutral-500">
            The order below drives row sorting on the TimeNavi Export and column order in the pivot. Colors are the row
            fills estimators use to find their jobs.
          </p>
          <div className="flex flex-col gap-1.5">
            {entries.map((e, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <div className="flex flex-col">
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === 0} onClick={() => move(i, -1)} aria-label="Move up">
                    ▲
                  </button>
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === entries.length - 1} onClick={() => move(i, 1)} aria-label="Move down">
                    ▼
                  </button>
                </div>
                <input
                  className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-xs font-semibold uppercase"
                  value={e.estimator}
                  maxLength={6}
                  placeholder="Init."
                  onChange={(ev) => update(i, { estimator: ev.target.value.toUpperCase() })}
                />
                <input
                  type="color"
                  className="h-7 w-9 cursor-pointer rounded border border-neutral-300"
                  value={HEX_RE.test(e.color) ? `#${e.color}` : '#cccccc'}
                  onChange={(ev) => update(i, { color: ev.target.value.slice(1).toUpperCase() })}
                />
                <input
                  className={`w-20 rounded border px-1.5 py-1 font-mono text-xs uppercase ${HEX_RE.test(e.color) ? 'border-neutral-300' : 'border-red-400 bg-red-50'}`}
                  value={e.color}
                  maxLength={6}
                  onChange={(ev) => update(i, { color: ev.target.value })}
                />
                <div className="h-6 flex-1 rounded border border-neutral-200" style={{ backgroundColor: HEX_RE.test(e.color) ? `#${e.color}` : undefined }} />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => remove(i)} aria-label="Remove estimator">
                  🗑
                </button>
              </div>
            ))}
          </div>
          <div className="mt-3 flex gap-2">
            <button type="button" className="btn text-xs" onClick={add}>
              + Add estimator
            </button>
            <button
              type="button"
              className="btn text-xs"
              onClick={() => setEntries(DEFAULT_LEGEND.map((l) => ({ ...l })))}
            >
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
            disabled={!valid}
            title={valid ? undefined : 'Each estimator needs unique initials and a 6-digit hex color'}
            onClick={() => onSave(cleaned)}
          >
            Save
          </button>
        </footer>
      </div>
    </div>
  );
}
