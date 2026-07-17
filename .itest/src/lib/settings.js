"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEFAULT_PIVOT_COLUMNS = exports.DEFAULT_LEGEND = void 0;
exports.ensureDefaultPivotColumns = ensureDefaultPivotColumns;
exports.defaultSettings = defaultSettings;
exports.loadSettings = loadSettings;
exports.saveSettings = saveSettings;
exports.mergePivotColumnOrder = mergePivotColumnOrder;
/**
 * Shared estimator palette — single source of truth for BOTH tabs (row fills
 * on TimeNavi Export, header/column bands on Pivot by Crew). Colors sampled
 * from the reference workbook; RF is a new pick from the same Excel accent
 * family. Text color is derived from fill luminance (see isDarkColor), so
 * dark fills like ND's 4472C4 automatically get white text.
 * Order drives row sorting in TimeNavi Export.
 */
exports.DEFAULT_LEGEND = [
    { estimator: 'JS', color: 'FABF8F' },
    { estimator: 'MH', color: 'FFFF99' },
    { estimator: 'PG', color: 'B1A0C7' },
    { estimator: 'TE', color: 'C4D79B' },
    { estimator: 'XC', color: '95B3D7' },
    { estimator: 'ND', color: '4472C4' },
    { estimator: 'RF', color: 'DA9694' },
];
/** Default pivot estimator column order (separate from legend row-sort order). */
exports.DEFAULT_PIVOT_COLUMNS = ['MH', 'PG', 'XC', 'ND', 'TE', 'RF'];
// v3: palette changed (ND recolored, RF added) — older stored settings are
// intentionally not migrated so the new defaults take effect.
const STORAGE_KEY = 'pct-completion-builder.settings.v3';
/**
 * Heal a stored pivot column order so every DEFAULT_PIVOT_COLUMNS entry is
 * present, and RF sits immediately after TE (even if a prior auto-append left
 * RF after extras like JS). Other estimators keep their relative order.
 */
function ensureDefaultPivotColumns(order) {
    const out = order.filter((e) => e !== '#N/A');
    for (let i = 0; i < exports.DEFAULT_PIVOT_COLUMNS.length; i++) {
        const def = exports.DEFAULT_PIVOT_COLUMNS[i];
        if (out.includes(def))
            continue;
        let insertAt = out.length;
        for (let j = i - 1; j >= 0; j--) {
            const prevAt = out.indexOf(exports.DEFAULT_PIVOT_COLUMNS[j]);
            if (prevAt >= 0) {
                insertAt = prevAt + 1;
                break;
            }
        }
        out.splice(insertAt, 0, def);
    }
    // RF must sit immediately after TE (Round 2 §4).
    const teAt = out.indexOf('TE');
    const rfAt = out.indexOf('RF');
    if (teAt >= 0 && rfAt >= 0 && rfAt !== teAt + 1) {
        out.splice(rfAt, 1);
        out.splice(out.indexOf('TE') + 1, 0, 'RF');
    }
    return out;
}
function defaultSettings() {
    return {
        legend: exports.DEFAULT_LEGEND.map((l) => ({ ...l })),
        pivotColumnOrder: [...exports.DEFAULT_PIVOT_COLUMNS],
    };
}
function parseLegend(raw) {
    if (typeof raw !== 'object' || raw === null || !Array.isArray(raw.legend)) {
        return null;
    }
    const legend = [];
    for (const entry of raw.legend) {
        if (typeof entry !== 'object' || entry === null)
            continue;
        const e = entry;
        if (typeof e.estimator !== 'string' || typeof e.color !== 'string')
            continue;
        if (!/^[0-9A-Fa-f]{6}$/.test(e.color))
            continue;
        legend.push({ estimator: e.estimator.trim().toUpperCase(), color: e.color.toUpperCase() });
    }
    return legend.length > 0 ? legend : null;
}
function parsePivotOrder(raw) {
    if (typeof raw !== 'object' || raw === null)
        return null;
    const arr = raw.pivotColumnOrder;
    if (!Array.isArray(arr))
        return null;
    const out = [];
    for (const item of arr) {
        if (typeof item !== 'string')
            continue;
        const t = item.trim().toUpperCase();
        if (t && !out.includes(t))
            out.push(t);
    }
    return out.length > 0 ? out : null;
}
function loadSettings() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) {
            const parsed = JSON.parse(raw);
            const legend = parseLegend(parsed);
            if (legend) {
                const pivotColumnOrder = ensureDefaultPivotColumns(parsePivotOrder(parsed) ?? [...exports.DEFAULT_PIVOT_COLUMNS]);
                return { legend, pivotColumnOrder };
            }
        }
    }
    catch {
        // fall through
    }
    return defaultSettings();
}
function saveSettings(settings) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
    }
    catch {
        // localStorage unavailable — settings just won't persist.
    }
}
/**
 * Merge estimators present in data into the stored pivot column order.
 * `#N/A` is never stored — it is appended at render time when hours exist.
 */
function mergePivotColumnOrder(stored, present) {
    const regular = present.filter((e) => e !== '#N/A');
    const merged = [
        ...stored.filter((e) => regular.includes(e)),
        ...regular.filter((e) => !stored.includes(e)),
    ];
    return merged;
}
