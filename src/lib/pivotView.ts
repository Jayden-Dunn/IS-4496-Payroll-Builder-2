/**
 * Round 4 §2 — pivot-style filters and sorting for the Pivot by Crew tab.
 * Pure view layer: the underlying data, the other tabs, the payroll math and
 * the exceptions counts never change. The export uses the full model unless
 * the "Export current filtered view" toggle is on.
 */
import { round2 } from './format';
import type { PivotCrew, PivotEmployee, PivotJobRow, PivotModel } from './types';

export type SortDir = 'none' | 'asc' | 'desc';
export type RowField = 'group' | 'username';

export interface PivotViewState {
  /** Unchecked estimator columns (§2.1). */
  hiddenEstimators: string[];
  /** Unchecked crew labels (§2.2, field = group). */
  hiddenGroups: string[];
  /** Unchecked employees (§2.2, field = username). */
  hiddenUsernames: string[];
  /** Column Labels sort; 'none' = persisted Settings order. */
  columnSort: SortDir;
  /** Row Labels sort target field. */
  rowField: RowField;
  /** Row Labels sort; 'none' = crews A→Z with lead-first employees (Round 2). */
  rowSort: SortDir;
}

export const EMPTY_PIVOT_VIEW: PivotViewState = {
  hiddenEstimators: [],
  hiddenGroups: [],
  hiddenUsernames: [],
  columnSort: 'none',
  rowField: 'group',
  rowSort: 'none',
};

export function isPivotViewActive(view: PivotViewState): boolean {
  return (
    view.hiddenEstimators.length > 0 ||
    view.hiddenGroups.length > 0 ||
    view.hiddenUsernames.length > 0 ||
    view.columnSort !== 'none' ||
    view.rowSort !== 'none'
  );
}

const byText = (a: string, b: string): number => a.localeCompare(b, 'en', { sensitivity: 'base' });

/**
 * Apply the view state to a pivot model: hide columns/rows and recompute the
 * Grand Total column + row over what remains — exactly what a real Excel
 * pivot does. Returns the input model untouched when nothing is active.
 */
export function applyPivotView(pivot: PivotModel, view: PivotViewState): PivotModel {
  if (!isPivotViewActive(view)) return pivot;

  const hiddenE = new Set(view.hiddenEstimators);
  let estimators = pivot.estimators.filter((e) => !hiddenE.has(e));
  if (view.columnSort !== 'none') {
    estimators = [...estimators].sort((a, b) => byText(a, b) * (view.columnSort === 'asc' ? 1 : -1));
  }

  const pickRecord = (rec: Record<string, number>): { byEstimator: Record<string, number>; total: number } => {
    const out: Record<string, number> = {};
    let total = 0;
    for (const e of estimators) {
      const v = rec[e];
      if (v !== undefined) {
        out[e] = v;
        total += v;
      }
    }
    return { byEstimator: out, total: round2(total) };
  };

  const hiddenG = new Set(view.hiddenGroups);
  const hiddenU = new Set(view.hiddenUsernames);

  let crews: PivotCrew[] = pivot.crews
    .filter((c) => !hiddenG.has(c.crew))
    .map((c) => {
      let employees = c.employees.filter((e) => !hiddenU.has(e.username));
      if (view.rowField === 'username' && view.rowSort !== 'none') {
        employees = [...employees].sort(
          (a, b) => byText(a.username, b.username) * (view.rowSort === 'asc' ? 1 : -1),
        );
      }
      const mapped: PivotEmployee[] = employees.map((emp) => {
        const jobs: PivotJobRow[] = emp.jobs.map((j) => {
          const picked = pickRecord(j.byEstimator);
          return { ...j, byEstimator: picked.byEstimator, total: picked.total };
        });
        const picked = pickRecord(emp.byEstimator);
        return {
          ...emp,
          jobs,
          byEstimator: picked.byEstimator,
          total: picked.total,
          overtime: picked.total > 40,
        };
      });
      return { crew: c.crew, employees: mapped };
    })
    .filter((c) => c.employees.length > 0);

  if (view.rowField === 'group' && view.rowSort !== 'none') {
    crews = [...crews].sort((a, b) => byText(a.crew, b.crew) * (view.rowSort === 'asc' ? 1 : -1));
  }

  const grandByEstimator: Record<string, number> = {};
  let grandTotal = 0;
  for (const c of crews) {
    for (const emp of c.employees) {
      for (const e of estimators) {
        const v = emp.byEstimator[e];
        if (v !== undefined) {
          grandByEstimator[e] = round2((grandByEstimator[e] ?? 0) + v);
          grandTotal += v;
        }
      }
    }
  }

  return {
    title: pivot.title,
    estimators,
    crews,
    grandByEstimator,
    grandTotal: round2(grandTotal),
  };
}
