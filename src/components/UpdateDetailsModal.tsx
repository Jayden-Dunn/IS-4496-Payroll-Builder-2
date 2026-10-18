import { useMemo, useState } from 'react';
import type { UpdateDetail, UpdateDetailGroup } from '../lib/reportDiff';

/**
 * Round 11 — opened from the "Updated from … · N values changed" line, after
 * an "Update existing report" has been applied. Lists every cell the update
 * touched: the tab, the row (AP# first), the column, and what happened there
 * ("40% changed to 38%").
 *
 * Read-only. The decisions were made in the review; this is the receipt.
 */

const GROUPS: { id: UpdateDetailGroup; title: string; blurb: string }[] = [
  {
    id: 'replaced',
    title: 'Manual edits replaced by the new file',
    blurb: 'You chose "new" for these in the review, so your typed value was dropped.',
  },
  {
    id: 'kept',
    title: 'Manual edits kept',
    blurb: 'The new file has a different value here, but your typed value still wins.',
  },
  {
    id: 'orphanDiscarded',
    title: 'Edits discarded (row not in the new file)',
    blurb: '',
  },
  {
    id: 'orphanKept',
    title: 'Edits parked (row not in the new file)',
    blurb: 'Not applied while the row is missing; they come back if the row returns.',
  },
  {
    id: 'changed',
    title: 'Values the new file changed',
    blurb: 'Cells nobody had edited, recalculated from the new upload.',
  },
];

export interface UpdateDetailsModalProps {
  details: readonly UpdateDetail[];
  /** Filename(s) the update came from, for the header line. */
  sourceLabel: string;
  onClose: () => void;
}

export function UpdateDetailsModal({ details, sourceLabel, onClose }: UpdateDetailsModalProps): JSX.Element {
  const [query, setQuery] = useState('');

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return details;
    return details.filter((d) => `${d.tab} ${d.row} ${d.column} ${d.change}`.toLowerCase().includes(q));
  }, [details, query]);

  const byGroup = useMemo(() => {
    const m = new Map<UpdateDetailGroup, UpdateDetail[]>();
    for (const d of filtered) {
      const list = m.get(d.group);
      if (list) list.push(d);
      else m.set(d.group, [d]);
    }
    return m;
  }, [filtered]);

  return (
    <div
      className="fixed inset-0 z-[70] flex items-center justify-center bg-black/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="update-details-title"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="flex max-h-[85vh] w-full max-w-4xl flex-col rounded-lg border border-neutral-300 bg-white shadow-xl">
        <div className="border-b border-neutral-200 px-5 py-3">
          <h2 id="update-details-title" className="text-base font-bold text-neutral-900">
            What the update changed
          </h2>
          <p className="mt-1 text-xs text-neutral-600">
            From {sourceLabel} · {details.length} {details.length === 1 ? 'change' : 'changes'}
          </p>
          {details.length > 0 && (
            <input
              type="search"
              className="mt-2 w-full rounded border border-neutral-300 px-2 py-1 text-xs"
              placeholder="Filter by tab, AP#, column or value…"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          )}
        </div>

        <div className="min-h-0 flex-1 overflow-y-auto px-5 py-3">
          {details.length === 0 && (
            <p className="text-xs text-neutral-600">
              The new upload did not change any value on the report&rsquo;s editable cells.
            </p>
          )}
          {details.length > 0 && filtered.length === 0 && (
            <p className="text-xs text-neutral-600">Nothing matches &ldquo;{query}&rdquo;.</p>
          )}
          {GROUPS.map((g) => {
            const rows = byGroup.get(g.id);
            if (!rows || rows.length === 0) return null;
            return (
              <section key={g.id} className="mb-4">
                <h3 className="text-xs font-bold uppercase tracking-wide text-neutral-700">
                  {g.title} ({rows.length})
                </h3>
                {g.blurb && <p className="mb-1 mt-0.5 text-[11px] text-neutral-500">{g.blurb}</p>}
                <table className="w-full border-collapse text-xs">
                  <thead>
                    <tr className="bg-neutral-100 text-left">
                      <th className="border border-neutral-300 px-2 py-1">Tab</th>
                      <th className="border border-neutral-300 px-2 py-1">AP# · Row</th>
                      <th className="border border-neutral-300 px-2 py-1">Column</th>
                      <th className="border border-neutral-300 px-2 py-1">Change</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((d, i) => (
                      <tr key={i} className={g.id === 'replaced' ? 'bg-amber-50' : ''}>
                        <td className="whitespace-nowrap border border-neutral-300 px-2 py-1">{d.tab}</td>
                        <td className="border border-neutral-300 px-2 py-1">{d.row}</td>
                        <td className="whitespace-nowrap border border-neutral-300 px-2 py-1">{d.column}</td>
                        <td className="border border-neutral-300 px-2 py-1 font-semibold">{d.change}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>
            );
          })}
        </div>

        <div className="flex justify-end border-t border-neutral-200 px-5 py-3">
          <button type="button" className="btn-primary px-4 py-1.5 text-sm" onClick={onClose}>
            Close
          </button>
        </div>
      </div>
    </div>
  );
}
