import type { AppSettings, LegendEntry } from './types';

/**
 * Shared estimator palette — single source of truth for BOTH tabs (row fills
 * on TimeNavi Export, header/column bands on Pivot by Crew). Colors sampled
 * from the reference workbook; RF is a new pick from the same Excel accent
 * family. Text color is derived from fill luminance (see isDarkColor), so
 * dark fills like ND's 4472C4 automatically get white text.
 * Order drives row sorting in TimeNavi Export.
 */
export const DEFAULT_LEGEND: LegendEntry[] = [
  { estimator: 'JS', color: 'FABF8F' },
  { estimator: 'MH', color: 'FFFF99' },
  { estimator: 'PG', color: 'B1A0C7' },
  { estimator: 'TE', color: 'C4D79B' },
  { estimator: 'XC', color: '95B3D7' },
  { estimator: 'ND', color: '4472C4' },
  { estimator: 'RF', color: 'DA9694' },
];

/** Default pivot estimator column order (separate from legend row-sort order). */
export const DEFAULT_PIVOT_COLUMNS: string[] = ['MH', 'PG', 'XC', 'ND', 'TE', 'RF'];

// v3: palette changed (ND recolored, RF added) — older stored settings are
// intentionally not migrated so the new defaults take effect.
const STORAGE_KEY = 'pct-completion-builder.settings.v3';

/**
 * Heal a stored pivot column order so every DEFAULT_PIVOT_COLUMNS entry is
 * present, and RF sits immediately after TE (even if a prior auto-append left
 * RF after extras like JS). Other estimators keep their relative order.
 */
export function ensureDefaultPivotColumns(order: string[]): string[] {
  const out = order.filter((e) => e !== '#N/A');
  for (let i = 0; i < DEFAULT_PIVOT_COLUMNS.length; i++) {
    const def = DEFAULT_PIVOT_COLUMNS[i];
    if (out.includes(def)) continue;
    let insertAt = out.length;
    for (let j = i - 1; j >= 0; j--) {
      const prevAt = out.indexOf(DEFAULT_PIVOT_COLUMNS[j]);
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

export function defaultSettings(): AppSettings {
  return {
    legend: DEFAULT_LEGEND.map((l) => ({ ...l })),
    pivotColumnOrder: [...DEFAULT_PIVOT_COLUMNS],
  };
}

function parseLegend(raw: unknown): LegendEntry[] | null {
  if (typeof raw !== 'object' || raw === null || !Array.isArray((raw as { legend?: unknown }).legend)) {
    return null;
  }
  const legend: LegendEntry[] = [];
  for (const entry of (raw as { legend: unknown[] }).legend) {
    if (typeof entry !== 'object' || entry === null) continue;
    const e = entry as { estimator?: unknown; color?: unknown };
    if (typeof e.estimator !== 'string' || typeof e.color !== 'string') continue;
    if (!/^[0-9A-Fa-f]{6}$/.test(e.color)) continue;
    legend.push({ estimator: e.estimator.trim().toUpperCase(), color: e.color.toUpperCase() });
  }
  return legend.length > 0 ? legend : null;
}

function parsePivotOrder(raw: unknown): string[] | null {
  if (typeof raw !== 'object' || raw === null) return null;
  const arr = (raw as { pivotColumnOrder?: unknown }).pivotColumnOrder;
  if (!Array.isArray(arr)) return null;
  const out: string[] = [];
  for (const item of arr) {
    if (typeof item !== 'string') continue;
    const t = item.trim().toUpperCase();
    if (t && !out.includes(t)) out.push(t);
  }
  return out.length > 0 ? out : null;
}

export function loadSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed: unknown = JSON.parse(raw);
      const legend = parseLegend(parsed);
      if (legend) {
        const pivotColumnOrder = ensureDefaultPivotColumns(
          parsePivotOrder(parsed) ?? [...DEFAULT_PIVOT_COLUMNS],
        );
        return { legend, pivotColumnOrder };
      }
    }
  } catch {
    // fall through
  }
  return defaultSettings();
}

export function saveSettings(settings: AppSettings): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // localStorage unavailable — settings just won't persist.
  }
}

/**
 * Merge estimators present in data into the stored pivot column order.
 * `#N/A` is never stored — it is appended at render time when hours exist.
 */
export function mergePivotColumnOrder(stored: string[], present: string[]): string[] {
  const regular = present.filter((e) => e !== '#N/A');
  const merged = [
    ...stored.filter((e) => regular.includes(e)),
    ...regular.filter((e) => !stored.includes(e)),
  ];
  return merged;
}
