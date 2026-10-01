import { useMemo, useState } from 'react';
import type { Conflict, OrphanedEdit, UpdateReview } from '../lib/reportDiff';

/**
 * Shown by "Update existing report" when — and only when — a re-uploaded file
 * would overwrite something the user typed by hand.
 *
 * Values the new file changes that nobody edited never reach this screen; they
 * are applied and counted. That restraint is the whole point: a review that
 * listed every moved number would be too long to read, and the one thing worth
 * a decision would be buried in it.
 */

export type Resolution = 'keep' | 'take';

export interface UpdateReviewModalProps {
  review: UpdateReview;
  /** Filename(s) being applied, for the header line. */
  sourceLabel: string;
  onCancel: () => void;
  /**
   * `take` lists the conflicts whose manual edit should be dropped (relinking
   * the cell to the new computed value); `discardOrphans` lists orphaned edit
   * keys to delete outright. Everything not named is left exactly as it is.
   */
  onApply: (choices: { take: Conflict[]; discardOrphans: OrphanedEdit[] }) => void;
}

const conflictId = (c: Conflict): string => `${c.scope}\u0000${c.editKey}\u0000${c.field}`;

export function UpdateReviewModal({
  review,
  sourceLabel,
  onCancel,
  onApply,
}: UpdateReviewModalProps): JSX.Element {
  // Default is to keep every manual edit: it is the work being protected, and
  // a wrong "keep" is visible in the cell, while a wrong "take" is silent.
  const [taken, setTaken] = useState<Set<string>>(new Set());
  const [discarded, setDiscarded] = useState<Set<string>>(new Set());

  const allTaken = useMemo(
    () => review.conflicts.length > 0 && taken.size === review.conflicts.length,
    [review.conflicts.length, taken.size],
  );

  const toggle = (c: Conflict, next: Resolution): void => {
    setTaken((prev) => {
      const s = new Set(prev);
      if (next === 'take') s.add(conflictId(c));
      else s.delete(conflictId(c));
      return s;
    });
  };

  const apply = (): void => {
    onApply({
      take: review.conflicts.filter((c) => taken.has(conflictId(c))),
      discardOrphans: review.orphans.filter((o) => discarded.has(`${o.scope}\u0000${o.editKey}`)),
    });
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="update-review-title"
    >
      <div className="flex max-h-[85vh] w-full max-w-3xl flex-col rounded-lg border border-neutral-300 bg-white shadow-xl">
        <div className="border-b border-neutral-200 px-5 py-3">
          <h2 id="update-review-title" className="text-base font-bold text-neutral-900">
            {review.conflicts.length === 0
              ? 'Edits with no matching row'
              : `${review.conflicts.length} manual ${review.conflicts.length === 1 ? 'edit' : 'edits'} would be overwritten`}
          </h2>
          <p className="mt-1 text-xs text-neutral-600">
            From {sourceLabel}.{' '}
            {review.silentChangeCount > 0 && (
              <>
                {review.silentChangeCount} other{' '}
                {review.silentChangeCount === 1 ? 'value updates' : 'values update'} automatically — you
                never edited {review.silentChangeCount === 1 ? 'it' : 'those'}, so there is nothing to decide.
              </>
            )}
          </p>
        </div>

        <div className="min-h-0 flex-1 overflow-y-auto px-5 py-3">
          {review.conflicts.length > 0 && (
            <>
              <div className="mb-2 flex items-center gap-2">
                <button
                  type="button"
                  className={`rounded border px-2 py-1 text-xs font-semibold ${
                    taken.size === 0
                      ? 'border-blue-500 bg-blue-50 text-blue-800'
                      : 'border-neutral-300 bg-white text-neutral-700 hover:bg-neutral-50'
                  }`}
                  onClick={() => setTaken(new Set())}
                >
                  Keep all my edits
                </button>
                <button
                  type="button"
                  className={`rounded border px-2 py-1 text-xs font-semibold ${
                    allTaken
                      ? 'border-blue-500 bg-blue-50 text-blue-800'
                      : 'border-neutral-300 bg-white text-neutral-700 hover:bg-neutral-50'
                  }`}
                  onClick={() => setTaken(new Set(review.conflicts.map(conflictId)))}
                >
                  Take all new values
                </button>
                <span className="text-[11px] text-neutral-500">or choose per row below</span>
              </div>

              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr className="bg-neutral-100 text-left">
                    <th className="border border-neutral-300 px-2 py-1">Row</th>
                    <th className="border border-neutral-300 px-2 py-1">Field</th>
                    <th className="border border-neutral-300 px-2 py-1 text-right">Your edit</th>
                    <th className="border border-neutral-300 px-2 py-1 text-right">Was</th>
                    <th className="border border-neutral-300 px-2 py-1 text-right">New file</th>
                    <th className="border border-neutral-300 px-2 py-1">Use</th>
                  </tr>
                </thead>
                <tbody>
                  {review.conflicts.map((c) => {
                    const take = taken.has(conflictId(c));
                    return (
                      <tr key={conflictId(c)} className={take ? 'bg-amber-50' : ''}>
                        <td className="border border-neutral-300 px-2 py-1">{c.rowLabel}</td>
                        <td className="border border-neutral-300 px-2 py-1">{c.fieldLabel}</td>
                        <td className="border border-neutral-300 px-2 py-1 text-right font-semibold text-blue-700">
                          {c.manualValue || '(blank)'}
                        </td>
                        <td className="border border-neutral-300 px-2 py-1 text-right text-neutral-500">
                          {c.oldValue || '(blank)'}
                        </td>
                        <td className="border border-neutral-300 px-2 py-1 text-right font-semibold">
                          {c.newValue || '(blank)'}
                        </td>
                        <td className="border border-neutral-300 px-2 py-1">
                          <div className="flex gap-1">
                            <button
                              type="button"
                              className={`rounded px-1.5 py-0.5 text-[11px] font-semibold ${
                                take ? 'bg-neutral-100 text-neutral-600' : 'bg-blue-600 text-white'
                              }`}
                              onClick={() => toggle(c, 'keep')}
                            >
                              mine
                            </button>
                            <button
                              type="button"
                              className={`rounded px-1.5 py-0.5 text-[11px] font-semibold ${
                                take ? 'bg-amber-600 text-white' : 'bg-neutral-100 text-neutral-600'
                              }`}
                              onClick={() => toggle(c, 'take')}
                            >
                              new
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </>
          )}

          {review.orphans.length > 0 && (
            <div className="mt-4">
              <h3 className="text-xs font-bold uppercase tracking-wide text-neutral-700">
                No longer in this file
              </h3>
              <p className="mb-2 mt-0.5 text-[11px] text-neutral-500">
                These rows are not in the new upload, so there is no new value to take. Kept edits are parked,
                not applied — if the row comes back, the edit comes back with it.
              </p>
              <table className="w-full border-collapse text-xs">
                <tbody>
                  {review.orphans.map((o) => {
                    const id = `${o.scope}\u0000${o.editKey}`;
                    const gone = discarded.has(id);
                    return (
                      <tr key={id} className={gone ? 'opacity-50' : ''}>
                        <td className="border border-neutral-300 px-2 py-1">{o.rowLabel}</td>
                        <td className="border border-neutral-300 px-2 py-1 text-neutral-600">{o.summary}</td>
                        <td className="w-24 border border-neutral-300 px-2 py-1">
                          <button
                            type="button"
                            className="rounded bg-neutral-100 px-1.5 py-0.5 text-[11px] font-semibold text-neutral-700 hover:bg-neutral-200"
                            onClick={() =>
                              setDiscarded((prev) => {
                                const s = new Set(prev);
                                if (s.has(id)) s.delete(id);
                                else s.add(id);
                                return s;
                              })
                            }
                          >
                            {gone ? 'undo' : 'discard'}
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>

        <div className="flex justify-end gap-2 border-t border-neutral-200 px-5 py-3">
          <button type="button" className="btn px-4 py-1.5 text-sm" onClick={onCancel}>
            Cancel
          </button>
          <button type="button" className="btn-primary px-4 py-1.5 text-sm" onClick={apply}>
            Update report
          </button>
        </div>
      </div>
    </div>
  );
}
