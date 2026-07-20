"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.EMPTY_PIVOT_VIEW = void 0;
exports.isPivotViewActive = isPivotViewActive;
exports.applyPivotView = applyPivotView;
/**
 * Round 4 §2 — pivot-style filters and sorting for the Pivot by Crew tab.
 * Pure view layer: the underlying data, the other tabs, the payroll math and
 * the exceptions counts never change. The export uses the full model unless
 * the "Export current filtered view" toggle is on.
 */
const format_1 = require("./format");
exports.EMPTY_PIVOT_VIEW = {
    hiddenEstimators: [],
    hiddenGroups: [],
    hiddenUsernames: [],
    columnSort: 'none',
    rowField: 'group',
    rowSort: 'none',
};
function isPivotViewActive(view) {
    return (view.hiddenEstimators.length > 0 ||
        view.hiddenGroups.length > 0 ||
        view.hiddenUsernames.length > 0 ||
        view.columnSort !== 'none' ||
        view.rowSort !== 'none');
}
const byText = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' });
/**
 * Apply the view state to a pivot model: hide columns/rows and recompute the
 * Grand Total column + row over what remains — exactly what a real Excel
 * pivot does. Returns the input model untouched when nothing is active.
 */
function applyPivotView(pivot, view) {
    if (!isPivotViewActive(view))
        return pivot;
    const hiddenE = new Set(view.hiddenEstimators);
    let estimators = pivot.estimators.filter((e) => !hiddenE.has(e));
    if (view.columnSort !== 'none') {
        estimators = [...estimators].sort((a, b) => byText(a, b) * (view.columnSort === 'asc' ? 1 : -1));
    }
    const pickRecord = (rec) => {
        const out = {};
        let total = 0;
        for (const e of estimators) {
            const v = rec[e];
            if (v !== undefined) {
                out[e] = v;
                total += v;
            }
        }
        return { byEstimator: out, total: (0, format_1.round2)(total) };
    };
    const hiddenG = new Set(view.hiddenGroups);
    const hiddenU = new Set(view.hiddenUsernames);
    let crews = pivot.crews
        .filter((c) => !hiddenG.has(c.crew))
        .map((c) => {
        let employees = c.employees.filter((e) => !hiddenU.has(e.username));
        if (view.rowField === 'username' && view.rowSort !== 'none') {
            employees = [...employees].sort((a, b) => byText(a.username, b.username) * (view.rowSort === 'asc' ? 1 : -1));
        }
        const mapped = employees.map((emp) => {
            const jobs = emp.jobs.map((j) => {
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
    const grandByEstimator = {};
    let grandTotal = 0;
    for (const c of crews) {
        for (const emp of c.employees) {
            for (const e of estimators) {
                const v = emp.byEstimator[e];
                if (v !== undefined) {
                    grandByEstimator[e] = (0, format_1.round2)((grandByEstimator[e] ?? 0) + v);
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
        grandTotal: (0, format_1.round2)(grandTotal),
    };
}
