import type { AppSettings, LegendEntry } from './types';

/** Exact default fills from the manual workbook. Order drives row sorting
 *  in the TimeNavi Export and column order in the pivot. */
export const DEFAULT_LEGEND: LegendEntry[] = [
  { estimator: 'JS', color: 'FABF8F' },
  { estimator: 'MH', color: 'FFFF99' },
  { estimator: 'PG', color: 'B1A0C7' },
  { estimator: 'TE', color: 'C4D79B' },
  { estimator: 'XC', color: '95B3D7' },
  { estimator: 'ND', color: 'D9D9D9' },
];

const STORAGE_KEY = 'pct-completion-builder.settings.v1';

export function defaultSettings(): AppSettings {
  return { legend: DEFAULT_LEGEND.map((l) => ({ ...l })) };
}

export function loadSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultSettings();
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== 'object' || parsed === null || !Array.isArray((parsed as { legend?: unknown }).legend)) {
      return defaultSettings();
    }
    const legend: LegendEntry[] = [];
    for (const entry of (parsed as { legend: unknown[] }).legend) {
      if (typeof entry !== 'object' || entry === null) continue;
      const e = entry as { estimator?: unknown; color?: unknown };
      if (typeof e.estimator !== 'string' || typeof e.color !== 'string') continue;
      if (!/^[0-9A-Fa-f]{6}$/.test(e.color)) continue;
      legend.push({ estimator: e.estimator.trim().toUpperCase(), color: e.color.toUpperCase() });
    }
    return legend.length > 0 ? { legend } : defaultSettings();
  } catch {
    return defaultSettings();
  }
}

export function saveSettings(settings: AppSettings): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // localStorage unavailable (private mode) — settings just won't persist.
  }
}
