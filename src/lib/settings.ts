import type {
  AppSettings,
  LegendEntry,
  ProfitShareEntry,
  SalariedEntry,
  WageRateEntry,
} from './types';

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

/**
 * Round 3 payroll seeds. These are tribal-knowledge DEFAULTS from the
 * reference documents — fully editable in Settings, never logic.
 */
export const DEFAULT_PROFIT_SHARE: ProfitShareEntry[] = [
  { lead: 'Antonio Zuniga', pct: 0.36 },
  { lead: 'Calixto Efrain Climaco', pct: 0.36 },
  { lead: 'Jesus Perez', pct: 0.36 },
  { lead: 'Carlos Ruiz Cruz', pct: 0.38 },
  { lead: 'Damian Vasquez', pct: 0.38 },
  { lead: 'Daniel Gonzales Bonilla', pct: 0.38 },
  { lead: 'Henry Limay Flores', pct: 0.38 },
  { lead: 'Rafael Aguilar', pct: 0.38 },
  { lead: 'Rene Salazar', pct: 0.38 },
  { lead: 'Sergio Viera', pct: 0.38 },
  { lead: 'Tomas Huerta', pct: 0.38 },
  { lead: "Pepe's Painting LLC", pct: 0.5 },
  { lead: 'Color Enterprises LLC', pct: 0.53 },
  { lead: 'Merino Custom Painting', pct: 0.5 },
  { lead: 'Alvaro Reyes', pct: 0.4 },
];

export const DEFAULT_PROFIT_SHARE_PCT = 0.38;

export const DEFAULT_WAGE_RATES: WageRateEntry[] = [
  { name: 'Adriana Lopez Barajas', rate: 25, hourlyRate: 20 },
  { name: 'Ana Bedoya Melgar', rate: 20 },
  { name: 'Antonio Zuniga', rate: 18 },
  { name: 'Arlette Zuniga Lopez', rate: 19.18 },
  { name: 'Armando Alfonso Marin', rate: 15 },
  { name: 'Calixto Efrain Climaco', rate: 19.18 },
  { name: 'Jesus Perez', rate: 19.18 },
  { name: 'Carlos Ruiz Cruz', rate: 18 },
  { name: 'Damian Vasquez', rate: 18 },
  { name: 'Daniel Gonzales Bonilla', rate: 18 },
  { name: 'Gertrudis Campoverde Anguiano', rate: 20 },
  { name: 'Henry Limay Flores', rate: 18 },
  { name: 'Leonardo Martinez-Rincon', rate: 18 },
  { name: 'Manuel Aguilar Campoverde', rate: 22 },
  { name: 'Miguel Borquez Gil', rate: 20 },
  { name: 'Miguel Martin', rate: 20 },
  { name: 'Miguel Rebollar Bocanegra', rate: 23 },
  { name: 'Natividad Ortega', rate: 25 },
  { name: 'Rafael Aguilar', rate: 18, hourlyRate: 25 },
  { name: 'Randolfo Lopez', rate: 18 },
  { name: 'Rene Salazar', rate: 18 },
  { name: 'Samuel Martin', rate: 20 },
  { name: 'Sergio Viera', rate: 18 },
  { name: 'Ted Frank', rate: 21 },
  { name: 'Tomas Huerta', rate: 18 },
  { name: 'Emily Trujillo', rate: 16 },
  { name: 'Jayden Dunn', rate: 18 },
  { name: 'Mikayla Dunn', rate: 17 },
  { name: 'Sandra Hermosillo', rate: 18 },
  { name: 'Tyler Dunn', rate: 15 },
];

export const DEFAULT_WAGE_RATE = 18;

export const DEFAULT_SALARIED: SalariedEntry[] = [
  { name: 'Jodi Dunn', section: 'ADMIN', role: 'Manager Salary', amount: 865 },
  { name: 'Jason Thompson', section: 'ESTIMATORS', role: 'Salary', amount: 1200 },
  { name: 'Paul Good', section: 'ESTIMATORS', role: 'Commissions', amount: 1600 },
  { name: 'Ryan Frey', section: 'ESTIMATORS', role: 'Commercial Estimator', amount: 1442 },
  { name: 'Terry Emry', section: 'ESTIMATORS', role: 'Commissions', amount: 1442 },
  { name: 'Xavier Cuellar', section: 'ESTIMATORS', role: 'Commercial Estimator', amount: 1200 },
  { name: 'Charles Mike Holden', section: 'MANAGER SALARY', role: 'Manager Salary', amount: 2500 },
  { name: 'Maria Donofrio', section: 'MANAGER SALARY', role: 'Manager Salary', amount: 1650 },
  { name: 'Nathan Dunn', section: 'MANAGER SALARY', role: 'Manager Salary', amount: 1000 },
  { name: 'Kristina Benson', section: 'MANAGER SALARY', role: 'Manager Salary', amount: 150 },
];

export const DEFAULT_MILEAGE_RATE = 0.625;

export const DEFAULT_SUBCONTRACTORS: string[] = [
  "Pepe's Painting LLC",
  'Color Enterprises LLC',
  'Merino Custom Painting',
];

// v4: Round 3 payroll settings added. v3 legend/pivot values are migrated;
// missing payroll sections get the seeds.
const STORAGE_KEY = 'pct-completion-builder.settings.v4';
const STORAGE_KEY_V3 = 'pct-completion-builder.settings.v3';

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
    profitShare: DEFAULT_PROFIT_SHARE.map((e) => ({ ...e })),
    defaultProfitSharePct: DEFAULT_PROFIT_SHARE_PCT,
    wageRates: DEFAULT_WAGE_RATES.map((e) => ({ ...e })),
    defaultWageRate: DEFAULT_WAGE_RATE,
    salariedRoster: DEFAULT_SALARIED.map((e) => ({ ...e })),
    mileageRate: DEFAULT_MILEAGE_RATE,
    subcontractors: [...DEFAULT_SUBCONTRACTORS],
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

function parsePayroll(raw: unknown): Partial<AppSettings> {
  const out: Partial<AppSettings> = {};
  if (typeof raw !== 'object' || raw === null) return out;
  const r = raw as Record<string, unknown>;

  if (Array.isArray(r.profitShare)) {
    const list: ProfitShareEntry[] = [];
    for (const e of r.profitShare) {
      const x = e as { lead?: unknown; pct?: unknown };
      if (typeof x.lead === 'string' && typeof x.pct === 'number' && Number.isFinite(x.pct)) {
        list.push({ lead: x.lead, pct: x.pct });
      }
    }
    if (list.length > 0) out.profitShare = list;
  }
  if (typeof r.defaultProfitSharePct === 'number' && Number.isFinite(r.defaultProfitSharePct)) {
    out.defaultProfitSharePct = r.defaultProfitSharePct;
  }
  if (Array.isArray(r.wageRates)) {
    const list: WageRateEntry[] = [];
    for (const e of r.wageRates) {
      const x = e as { name?: unknown; rate?: unknown; hourlyRate?: unknown };
      if (typeof x.name === 'string' && typeof x.rate === 'number' && Number.isFinite(x.rate)) {
        list.push({
          name: x.name,
          rate: x.rate,
          ...(typeof x.hourlyRate === 'number' && Number.isFinite(x.hourlyRate) ? { hourlyRate: x.hourlyRate } : {}),
        });
      }
    }
    if (list.length > 0) out.wageRates = list;
  }
  if (typeof r.defaultWageRate === 'number' && Number.isFinite(r.defaultWageRate)) {
    out.defaultWageRate = r.defaultWageRate;
  }
  if (Array.isArray(r.salariedRoster)) {
    const list: SalariedEntry[] = [];
    for (const e of r.salariedRoster) {
      const x = e as { name?: unknown; section?: unknown; role?: unknown; amount?: unknown };
      if (
        typeof x.name === 'string' &&
        (x.section === 'ADMIN' || x.section === 'ESTIMATORS' || x.section === 'MANAGER SALARY') &&
        typeof x.amount === 'number' &&
        Number.isFinite(x.amount)
      ) {
        list.push({ name: x.name, section: x.section, role: typeof x.role === 'string' ? x.role : '', amount: x.amount });
      }
    }
    if (list.length > 0) out.salariedRoster = list;
  }
  if (typeof r.mileageRate === 'number' && Number.isFinite(r.mileageRate)) out.mileageRate = r.mileageRate;
  if (Array.isArray(r.subcontractors)) {
    const subs = r.subcontractors.filter((s): s is string => typeof s === 'string' && s.trim() !== '');
    if (subs.length > 0) out.subcontractors = subs;
  }
  return out;
}

export function loadSettings(): AppSettings {
  try {
    // v4 first; fall back to v3 (legend/pivot only) so users keep palette edits.
    const raw = localStorage.getItem(STORAGE_KEY) ?? localStorage.getItem(STORAGE_KEY_V3);
    if (raw) {
      const parsed: unknown = JSON.parse(raw);
      const legend = parseLegend(parsed);
      if (legend) {
        const pivotColumnOrder = ensureDefaultPivotColumns(
          parsePivotOrder(parsed) ?? [...DEFAULT_PIVOT_COLUMNS],
        );
        return { ...defaultSettings(), ...parsePayroll(parsed), legend, pivotColumnOrder };
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
