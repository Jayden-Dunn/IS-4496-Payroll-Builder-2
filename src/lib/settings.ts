import type { AppSettings, LegendEntry } from './types';

/** Exact default fills from the manual workbook. Order drives row sorting in TimeNavi Export. */
export const DEFAULT_LEGEND: LegendEntry[] = [
  { estimator: 'JS', color: 'FABF8F' },
  { estimator: 'MH', color: 'FFFF99' },
  { estimator: 'PG', color: 'B1A0C7' },
  { estimator: 'TE', color: 'C4D79B' },
  { estimator: 'XC', color: '95B3D7' },
  { estimator: 'ND', color: 'D9D9D9' },
];

/** Default pivot estimator column order (separate from legend row-sort order). */
export const DEFAULT_PIVOT_COLUMNS: string[] = ['MH', 'PG', 'XC', 'ND', 'TE', 'JS'];

const STORAGE_KEY = 'pct-completion-builder.settings.v2';
const LEGACY_KEY = 'pct-completion-builder.settings.v1';

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
    for (const key of [STORAGE_KEY, LEGACY_KEY]) {
      const raw = localStorage.getItem(key);
      if (!raw) continue;
      const parsed: unknown = JSON.parse(raw);
      const legend = parseLegend(parsed);
      if (!legend) continue;
      const pivotColumnOrder = parsePivotOrder(parsed) ?? [...DEFAULT_PIVOT_COLUMNS];
      return { legend, pivotColumnOrder };
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
