import type {
  AppSettings,
  ConditionalRateRule,
  CrewLeadAlias,
  DirectoryEmployee,
  EntityCode,
  JobRateRule,
  LegendEntry,
  PainterRosterEntry,
  PayrollLineItem,
  PerJobRateOverride,
  PlAccountMapping,
  ProfitShareEntry,
  SalariedEntry,
  SavedDirectory,
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
  { estimator: 'RF', color: 'C4A484' },
  // Round 11 — JT's jobs.
  { estimator: 'JT', color: 'DDD9C4' },
];

/** Round 8 §5 — the RF fill this seed replaced; migrated once on load. */
export const LEGACY_RF_COLOR = 'DA9694';

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
  { lead: 'Damian Vazquez', pct: 0.38 },
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
  // Round 11 — new crew lead, at the default profit-share %.
  { lead: 'Andrew Mazzeo', pct: 0.38 },
];

export const DEFAULT_PROFIT_SHARE_PCT = 0.38;

const BASE_WAGE_RATES: WageRateEntry[] = [
  { name: 'Adriana Lopez Barajas', rate: 25, hourlyRate: 20 },
  { name: 'Ana Bedoya Melgar', rate: 20 },
  { name: 'Antonio Zuniga', rate: 18, hourlyRate: 25 },
  { name: 'Arlette Zuniga Lopez', rate: 19.18 },
  { name: 'Armando Alfonso Marin', rate: 15 },
  { name: 'Calixto Efrain Climaco', rate: 19.18 },
  { name: 'Jesus Perez', rate: 19.18 },
  { name: 'Carlos Ruiz Cruz', rate: 18, hourlyRate: 25 },
  { name: 'Damian Vazquez', rate: 18 },
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

/**
 * Round 8 §9.3 — rate seeds ("hourly" = hourlyRate, "PS" = rate).
 * Merged by name into BASE_WAGE_RATES: matching
 * names are patched field-by-field, new names are appended. SEEDS ONLY —
 * fully editable in Settings, never referenced as constants by rate logic.
 */
export const ROUND8_WAGE_SEEDS: WageRateEntry[] = [
  { name: 'Adriana Lopez Barajas', rate: 25, hourlyRate: 20 },
  { name: 'Antonio Zuniga', rate: 18, hourlyRate: 25 },
  { name: 'Carlos Ruiz Cruz', rate: 18, hourlyRate: 25 },
  { name: 'Damian Vazquez', rate: 18, hourlyRate: 18 },
  { name: 'Daniel Gonzales Bonilla', rate: 18, hourlyRate: 20 },
  { name: 'Henry Limay Flores', rate: 18 },
  { name: 'Jose Santos Hernandez Favela', rate: 18, hourlyRate: 20 },
  { name: 'Martin Ealey', rate: 18, hourlyRate: 18 },
  { name: 'Miguel Martin', rate: 20, hourlyRate: 20 },
  { name: 'Natividad Ortega', rate: 25, hourlyRate: 23.75 },
  { name: 'Rafael Aguilar Echevarria', rate: 18, hourlyRate: 30 },
  { name: 'Rene Salazar', rate: 18, hourlyRate: 16 },
  { name: 'Samuel Brewster', rate: 18 },
  { name: 'Sergio Viera', rate: 18, hourlyRate: 18 },
  { name: 'Thomas Slater', rate: 18, hourlyRate: 20 },
  { name: 'Tomas Huerta', rate: 18, hourlyRate: 25 },
  { name: 'Xavier Green', rate: 18, hourlyRate: 18 },
];

/**
 * Round 9 §12/§17/§18/§21/§22/§27/§35 — corrected and new per-painter rates.
 * `rate` is the profit-share / work-order rate, `hourlyRate` the T&M rate,
 * `commercialHourlyRate` and `wages2Rate` the two Round 9 §18 pay types.
 * SEEDS ONLY — every one of them is editable in Settings.
 */
export const ROUND9_WAGE_SEEDS: WageRateEntry[] = [
  // §22 — was $16/hr.
  { name: 'Emily Trujillo', rate: 18, hourlyRate: 18 },
  // §12 — was running at the $18 default.
  { name: 'Marvin Ealey', rate: 25, hourlyRate: 25, commercialHourlyRate: 25 },
  // §17 — $15 regular, profit-share jobs included.
  { name: 'Rocio Sanchez', rate: 15, hourlyRate: 15 },
  // §18 — Commercial Hourly painter, not a crew lead.
  { name: 'Samuel Brewster', rate: 18, commercialHourlyRate: 30 },
  // §27 — Utah; daily OT never applies (see DEFAULT_DAILY_OT_EXEMPT).
  { name: 'Tyler Dunn', rate: 15, hourlyRate: 15 },
  // §21/§35 — $25 on profit-share / work-order jobs, $18 on hourly jobs.
  { name: 'Xavier Green', rate: 25, hourlyRate: 18 },
];

/**
 * Round 10 — corrected rates. `rate` is the profit-share / work-order rate
 * (crew leads included, now that the separate crew-lead rate is gone) and
 * `hourlyRate` is the T&M rate. SEEDS ONLY — editable in Settings.
 */
export const ROUND10_WAGE_SEEDS: WageRateEntry[] = [
  // $20/hr on profit-share jobs.
  { name: 'Gertrudis Campoverde Anguiano', rate: 20 },
  // $30/hr on every hourly job, whatever the pay item says.
  { name: 'Rafael Aguilar Echevarria', rate: 18, hourlyRate: 30 },
  { name: 'Rafael Aguilar', rate: 18, hourlyRate: 30 },
  // $30/hr on hourly jobs.
  { name: 'Thomas Slater', rate: 18, hourlyRate: 30 },
  // $25/hr on hourly AND profit-share jobs.
  { name: 'Marvin Ealey', rate: 25, hourlyRate: 25, commercialHourlyRate: 25 },
  { name: 'Xavier Green', rate: 25, hourlyRate: 25, commercialHourlyRate: 25 },
  // $15/hr on hourly AND profit-share jobs.
  { name: 'Rocio Sanchez', rate: 15, hourlyRate: 15 },
];

/**
 * Round 11 — corrected rates. `rate` is the profit-share / work-order rate,
 * `hourlyRate` the T&M rate. "All jobs" / "across the board" sets both, plus
 * the Commercial Hourly and Wages 2 slots where the person had one, so no pay
 * item can quietly pay an old figure. SEEDS ONLY — editable in Settings.
 *
 * Names follow Settings, not the request's spelling: "Miguel Revelar" is
 * Miguel Rebollar Bocanegra, "Arlette Zunigo" is Arlette Zuniga Lopez, "Anna
 * Bedoya Melgar" is Ana Bedoya Melgar.
 */
export const ROUND11_WAGE_SEEDS: WageRateEntry[] = [
  { name: 'Emily Trujillo', rate: 16, hourlyRate: 16 },
  { name: 'Arlette Zuniga Lopez', rate: 13, hourlyRate: 13 },
  // $20 on hourly jobs, $18 on profit-share jobs.
  { name: 'Miguel Martin', rate: 18, hourlyRate: 20 },
  { name: 'Miguel Borquez Gil', rate: 20, hourlyRate: 20 },
  { name: 'Miguel Rebollar Bocanegra', rate: 23, hourlyRate: 23 },
  { name: 'Samuel Martin', rate: 20, hourlyRate: 20 },
  { name: 'Ted Frank', rate: 21, hourlyRate: 21 },
  { name: 'Thomas Slater', rate: 35, hourlyRate: 35 },
  { name: 'Ana Bedoya Melgar', rate: 20, hourlyRate: 20 },
  { name: 'Manuel Aguilar Campoverde', rate: 22, hourlyRate: 22 },
  { name: 'Damian Vazquez', rate: 18, hourlyRate: 18 },
  // New crew lead at the default rate.
  { name: 'Andrew Mazzeo', rate: 18 },
];

/**
 * Round 11 — the wage list is always alphabetical by name, however a row got
 * there (seed, migration, a hand-added row in Settings).
 */
export function sortWageRates(list: readonly WageRateEntry[]): WageRateEntry[] {
  return [...list].sort((a, b) => a.name.trim().localeCompare(b.name.trim(), 'en', { sensitivity: 'base' }));
}

/** Name key for wage-map merges (case/whitespace-insensitive). */
function wageKey(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Merge seed rows into a wage list BY NAME: matching rows are patched, unknown
 * names are appended. Never duplicates a row.
 *
 * `mode` decides what happens to a field the matched row already has:
 * - `seedWins` (building the seed list) — the seed value replaces it.
 * - `fillMissing` (migrating someone's saved settings) — only absent fields are
 *   filled, so a hand-edited rate survives.
 */
export function mergeWageRateSeeds(
  current: readonly WageRateEntry[],
  seeds: readonly WageRateEntry[],
  mode: 'seedWins' | 'fillMissing' = 'seedWins',
): WageRateEntry[] {
  const out = current.map((e) => ({ ...e }));
  const indexByName = new Map(out.map((e, i) => [wageKey(e.name), i]));
  for (const seed of seeds) {
    const at = indexByName.get(wageKey(seed.name));
    if (at === undefined) {
      out.push({ ...seed });
      indexByName.set(wageKey(seed.name), out.length - 1);
      continue;
    }
    const row = out[at];
    out[at] =
      mode === 'seedWins'
        ? { ...row, ...seed, name: row.name }
        : {
            ...row,
            ...(row.hourlyRate === undefined && seed.hourlyRate !== undefined
              ? { hourlyRate: seed.hourlyRate }
              : {}),
          };
  }
  return out;
}

export const DEFAULT_WAGE_RATES: WageRateEntry[] = sortWageRates(
  mergeWageRateSeeds(
    mergeWageRateSeeds(
      mergeWageRateSeeds(mergeWageRateSeeds(BASE_WAGE_RATES, ROUND8_WAGE_SEEDS), ROUND9_WAGE_SEEDS),
      ROUND10_WAGE_SEEDS,
    ),
    ROUND11_WAGE_SEEDS,
  ),
);

/**
 * Round 11 — calendar-title nicknames for crew leads. Matched without regard
 * to case, so "(welder)" and "(WELDER)" both resolve. SEED DATA — editable in
 * Settings.
 */
export const DEFAULT_CREW_LEAD_ALIASES: CrewLeadAlias[] = [
  { alias: 'Pepe', lead: "Pepe's Painting LLC" },
  { alias: 'Welder', lead: 'Color Enterprises LLC' },
  { alias: 'Carlito', lead: 'Carlos Ruiz' },
  { alias: 'Santos', lead: 'Jose Santos Hernandez' },
  { alias: 'Miguel', lead: 'Miguel Martin' },
  { alias: 'Andy', lead: 'Andrew Mazzeo' },
];

/**
 * Round 11 — the only people whose Payroll Preview block shows a
 * `NEST [IRA]` column. Token-matched, so "Carlos Ruiz" covers "Carlos Ruiz
 * Cruz". SEED DATA — editable in Settings.
 */
export const DEFAULT_NEST_NAMES: string[] = [
  'Ana Bedoya Melgar',
  'Carlos Ruiz',
  'Manuel Aguilar',
  'Rafael Aguilar',
  'Sergio Viera',
  'Emily Trujillo',
  'Paul Good',
  'Mike Holden',
  'Ryan Frey',
];

/**
 * Round 11 — known misspellings in the source exports, fixed where names
 * enter the app (`parseTimesheet`, `parseEmployeeList`), so every tab, every
 * rate lookup and every export sees one spelling. Each pair must have the same
 * number of name tokens; matching is case-insensitive.
 */
export const NAME_SPELLING_FIXES: { from: string; to: string }[] = [
  { from: 'Damian Vasquez', to: 'Damian Vazquez' },
];

/**
 * Apply `NAME_SPELLING_FIXES` to a name split across first/last fields. A fix
 * applies when every token of its `from` appears in the name; each of those
 * tokens is swapped for its counterpart, wherever it sits.
 */
export function fixNameSpelling(first: string, last: string): { first: string; last: string } {
  let f = first;
  let l = last;
  for (const fix of NAME_SPELLING_FIXES) {
    const fromToks = fix.from.toLowerCase().split(/\s+/);
    const toToks = fix.to.split(/\s+/);
    if (fromToks.length !== toToks.length) continue;
    const have = `${f} ${l}`.toLowerCase().split(/\s+/).filter(Boolean);
    if (!fromToks.every((t) => have.includes(t))) continue;
    const swap = (part: string): string =>
      part
        .split(/(\s+)/)
        .map((tok) => {
          const at = fromToks.indexOf(tok.toLowerCase());
          return at === -1 ? tok : toToks[at];
        })
        .join('');
    f = swap(f);
    l = swap(l);
  }
  return { first: f, last: l };
}

/** Same fix inside free text such as a crew group label ("Damian Vasquez's Crew"). */
export function fixNameSpellingInText(text: string): string {
  let out = text;
  for (const fix of NAME_SPELLING_FIXES) {
    const pattern = fix.from
      .split(/\s+/)
      .map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
      .join('\\s+');
    out = out.replace(new RegExp(`\\b${pattern}\\b`, 'gi'), fix.to);
  }
  return out;
}

export const DEFAULT_WAGE_RATE = 18;

/**
 * Round 8 §9.1 — seeded conditional rate rules.
 *
 * Round 9 §13 — the "Henry Limay Flores on Miguel Martin's job" rule is GONE:
 * when the two work the same job, the crew lead is whoever the Pivot by Job
 * `Crew Lead` column names, and the painter who clocked
 * "Crew Lead - Hourly Wage" is the lead for that job. The rule MECHANISM stays
 * available in Settings for genuinely negotiated cases.
 */
export const DEFAULT_CONDITIONAL_RATE_RULES: ConditionalRateRule[] = [];

/**
 * Round 8 §9.1 — seeded person + AP# rate overrides.
 *
 * Round 9 §18 — Samuel Brewster's $30 is no longer pinned to AP24097: he is
 * paid $30 wherever he clocks the "Commercial Hourly" pay type, which is what
 * his timesheet actually says. The override MECHANISM is untouched.
 */
export const DEFAULT_PER_JOB_RATE_OVERRIDES: PerJobRateOverride[] = [];

/** Round 8 §9.2 — daily-OT thresholds (logic lives in the payroll engine). */
export const DEFAULT_NV_DAILY_OT_RATE_THRESHOLD = 18;
export const DEFAULT_DAILY_OT_HOURS_THRESHOLD = 8;

/**
 * Round 8 §7.1 — per-entity P&L account label mapping. Sundries stays OUT of
 * materials (§8.3 supersedes Round 3 §2).
 */
export const DEFAULT_PL_ACCOUNT_MAPPING: Record<EntityCode, PlAccountMapping> = {
  NV: {
    income: ['Total Income'],
    materials: [
      'Equipment Rental for Jobs',
      'Job Materials Purchased',
      'Other Job Related Costs',
      'Paint',
      'Tools and Small Equipment',
    ],
    laborWagesSubs: ['Employee Labor', 'Subcontractor Labor', 'Wages'],
    sundries: ['Sundries -'],
    travel: ['Travel Expenses'],
    mileage: ['Mileage Expense'],
  },
  UT: {
    income: ['Total Income'],
    materials: ['Equipment Rental for Jobs', 'Job Materials', 'Paint', 'Tools & Small Equipment'],
    laborWagesSubs: ['Employee Labor', 'Subcontractors', 'Wages'],
    sundries: ['Sundries'],
    travel: ['Travel Expense'],
    mileage: [],
  },
};

/**
 * Round 8 §11 — painters seeded INACTIVE. `{Reno}` (and similar) display
 * metadata is stored as `note`, never in `name`. Seed data — fully editable.
 */
export const DEFAULT_PAINTER_ROSTER: PainterRosterEntry[] = [
  'Angel Chavez',
  'Artemio Rios Escobar',
  'Axel Angel Garcia Torres',
  'Calixto Efrain Climaco',
  'Cesar Lopez Molina',
  'Christian Estrada',
  'Danaell Bing',
  'Denuan David Martinez Bonilla',
  'Eddy Huerta Burgos',
  'Eduardo Martinez Guzman',
  'Edward Prieto',
  'Efrain Gutierrez-Salazar',
  'Elvira Barragan Bautista',
  'Enrique Guevara Rivera',
  'Expectacion Garcia',
  'Ezequiel Rodriguez',
  'Gledy Silva Callan',
  'Gustavo Rodriguez',
  'Gustavo Ventura Santos',
  'Isaac Ramirez',
  'Jayton Johnson-Bustamante',
  'Jesus Ochoa',
  'Jesus Perez',
  'Jonathan Rodriguez Gonzalez',
  'Jose Belen Delgado',
  'Jose Garcia Diaz',
  'Jose Guardado',
  'Jose Gustavo Amaya Hernandez',
  'Jose Martinez Sanchez',
  'Jose Rodriguez',
  'Juan Salazar',
  'Justin Fuentes {Reno}',
  'Lismark Prieto Claudio',
  'Luis Perez',
  'Luis Vargas Montor',
  'Marlon Giron',
  'Mauricio Garcia Partida',
  'Miguel Almanza Munoz',
  'Miguel Rebollar Bocanegra',
  'Miguel Rodriguez',
  'Patricio Marcelo Acosta',
  'Ramon Torres Barragan',
  'Randolfo Lopez',
  'Rene Meraz {Reno}',
  'Roman Tamayo',
  'Samuel Hernandez',
  'Steven Adams {Reno}',
  'Ted Frank',
  'Victor Escobar',
  'Welder Padilla Guevara',
  'Yadira Rios Tatara',
  'Zachary Dunn',
].map((raw) => {
  const meta = /\{[^}]*\}|\*[^*]*\*/.exec(raw);
  return {
    name: raw.replace(/\s*(\{[^}]*\}|\*[^*]*\*)\s*/g, ' ').trim(),
    active: false,
    ...(meta ? { note: meta[0] } : {}),
  };
});

const NO_MILEAGE_ROW_NAMES = [
  'Paul Good',
  'Charles Mike Holden',
  'Nathan Dunn',
  'Maria Donofrio',
  'Jeff Sherman',
  'Kristina Benson',
];

const RAW_SALARIED: SalariedEntry[] = [
  // Round 9 §28 — Jodi's mileage row reads "Reimbursements" and is a flat cell.
  {
    name: 'Jodi Dunn',
    section: 'ADMIN',
    role: 'Manager Salary',
    amount: 865,
    entity: 'NV',
    reimbursementsLabel: 'Reimbursements/additional pay',
  },
  { name: 'Jason Thompson', section: 'ESTIMATORS', role: 'Salary', amount: 1200, entity: 'NV' },
  { name: 'Paul Good', section: 'ESTIMATORS', role: 'Commissions', amount: 1600, entity: 'NV' },
  { name: 'Ryan Frey', section: 'ESTIMATORS', role: 'Commercial Estimator', amount: 1442, entity: 'NV' },
  { name: 'Terry Emry', section: 'ESTIMATORS', role: 'Commissions', amount: 1442, entity: 'NV' },
  { name: 'Xavier Cuellar', section: 'ESTIMATORS', role: 'Commercial Estimator', amount: 1200, entity: 'NV' },
  // Round 9 §34 — Jeff Sherman is paid out of the Utah entity.
  {
    name: 'Jeff Sherman',
    section: 'ESTIMATORS',
    role: 'Commercial Estimator',
    amount: 1400,
    entity: 'UT',
    taxes: 238.81,
    net: 1161.19,
  },
  { name: 'Charles Mike Holden', section: 'MANAGER SALARY', role: 'Manager Salary', amount: 2500, entity: 'NV' },
  { name: 'Maria Donofrio', section: 'MANAGER SALARY', role: 'Manager Salary', amount: 1650, entity: 'NV' },
  { name: 'Nathan Dunn', section: 'MANAGER SALARY', role: 'Manager Salary', amount: 1000, entity: 'NV' },
  // Round 9 §25 — once a month, on the month's first pay date.
  {
    name: 'Paul Letourneau',
    section: 'MANAGER SALARY',
    role: 'Monthly',
    amount: 1000,
    entity: 'NV',
    frequency: 'firstPayDateOfMonth',
  },
  /**
   * Round 10 — Kristina Benson is a subcontractor, so her block renders with
   * the subcontractor bands (below Pepe's Painting LLC) rather than in the
   * MANAGER SALARY section.
   */
  {
    name: 'Kristina Benson',
    section: 'SUBCONTRACTORS',
    role: 'Subcontractor',
    amount: 150,
    entity: 'NV',
    sub: true,
  },
];

/** Round 10 — the named people carry no mileage-reimbursement row. */
export const DEFAULT_SALARIED: SalariedEntry[] = RAW_SALARIED.map((e) =>
  NO_MILEAGE_ROW_NAMES.includes(e.name) ? { ...e, noMileageRow: true } : e,
);

/**
 * Round 9 §19/§23/§24/§29 — person-scoped extra rows on the Payroll Preview.
 * `deduction` comes OUT of the person's total, `reimbursement` goes in, and
 * `note` is a labeled blank line for the payroll clerk to fill in.
 * SEED DATA — fully editable in Settings.
 */
export const DEFAULT_PAYROLL_LINE_ITEMS: PayrollLineItem[] = [
  // §19 — labeled row, value intentionally blank.
  { id: 'li-ted-frank-gas', person: 'Ted Frank', label: 'Gas Reimb', amount: null, kind: 'note', frequency: 'weekly' },
  // §23 — weekly Aflac deductions.
  { id: 'li-paul-good-aflac', person: 'Paul Good', label: 'Aflac', amount: 39.9, kind: 'deduction', frequency: 'weekly' },
  {
    id: 'li-mike-holden-aflac',
    person: 'Charles Mike Holden',
    label: 'Aflac',
    amount: 80.26,
    kind: 'deduction',
    frequency: 'weekly',
  },
  // §24 — first pay date of the month only.
  {
    id: 'li-mike-holden-health',
    person: 'Charles Mike Holden',
    label: 'Health Insurance',
    amount: 689.8,
    // Round 11 — paid TO him, not withheld.
    kind: 'reimbursement',
    frequency: 'firstPayDateOfMonth',
  },
  // §29/§34 — a blank "Reimbursements" row in addition to mileage reimb.
  ...[
    'Paul Good',
    'Xavier Cuellar',
    'Charles Mike Holden',
    'Jeff Sherman',
    'Jason Thompson',
    'Ryan Frey',
    'Terry Emry',
    // Round 10 — Jayden Dunn gets one too.
    'Jayden Dunn',
  ].map(
    (person): PayrollLineItem => ({
      id: `li-${person.toLowerCase().replace(/\s+/g, '-')}-reimb`,
      person,
      label: 'Reimbursements/additional pay',
      amount: null,
      kind: 'reimbursement',
      frequency: 'weekly',
    }),
  ),
  // Round 11 — a "Holiday pay" row in the office section, blank until typed.
  ...['Emily Trujillo', 'Sandra Hermosillo'].map(
    (person): PayrollLineItem => ({
      id: `li-${person.toLowerCase().replace(/\s+/g, '-')}-holiday`,
      person,
      label: 'Holiday pay',
      amount: null,
      kind: 'reimbursement',
      frequency: 'weekly',
    }),
  ),
  // Round 10 — Jason Thompson's weekly c/s deduction.
  {
    id: 'li-jason-thompson-cs',
    person: 'Jason Thompson',
    label: 'c/s',
    amount: 134,
    kind: 'deduction',
    frequency: 'weekly',
  },
];

/**
 * Round 10 — salaried people whose block has NO mileage-reimbursement row at
 * all. (Distinct from `reimbursementsLabel`, which swaps the mileage row for a
 * flat reimbursements row.) Seed data — editable in Settings.
 */
export const DEFAULT_NO_MILEAGE_ROW: string[] = [
  'Paul Good',
  'Charles Mike Holden',
  'Nathan Dunn',
  'Maria Donofrio',
  'Jeff Sherman',
  'Kristina Benson',
];

/** Round 9 §32 — leads charged for ALL of a job's materials, not half. */
export const DEFAULT_FULL_MATERIALS_LEADS: string[] = [
  'Color Enterprises LLC',
  // Round 10 — Pepe's pays all of the materials on his jobs too.
  "Pepe's Painting LLC",
];

/** Round 9 §27 — people the NV daily-overtime rule never applies to. */
export const DEFAULT_DAILY_OT_EXEMPT: string[] = ['Tyler Dunn'];

export const DEFAULT_MILEAGE_RATE = 0.625;

export const DEFAULT_SUBCONTRACTORS: string[] = [
  "Pepe's Painting LLC",
  'Color Enterprises LLC',
  'Merino Custom Painting',
];

/**
 * Round 4 §5.8 / Round 8 §9.1 — prevailing-wage (Davis-Bacon) job patterns +
 * minimum rate. Broadened from `SNRHA *SS*` to the bare `SNRHA` substring so
 * every SNRHA job matches regardless of the `*SS*` marker.
 */
export const DEFAULT_PREVAILING_PATTERNS: string[] = ['SNRHA'];
export const DEFAULT_PREVAILING_MIN_RATE = 19.18;

/** The narrow Round 4 pattern replaced by `SNRHA`; migrated once on load. */
export const LEGACY_PREVAILING_PATTERN = 'SNRHA *SS*';

// One-time settings migrations (Round 8). Each id is recorded in
// AppSettings.appliedMigrations so it never runs twice on the same blob.
export const MIGRATION_RF_COLOR = 'r8-rf-color';
export const MIGRATION_DB_PATTERN = 'r8-db-pattern';
export const MIGRATION_WAGE_SEEDS = 'r8-wage-seeds';
export const MIGRATION_CARLOS_HOURLY = 'r8-carlos-hourly-25';
/** Round 9 — corrected wage rates + the new pay-type rates (§12/§17/§18/§21/§22/§35). */
export const MIGRATION_R9_WAGES = 'r9-wage-seeds';
/** Round 9 §19/§23/§24/§29 — seeded payroll line items. */
export const MIGRATION_R9_LINE_ITEMS = 'r9-line-items';
/** Round 9 §25/§26/§28/§34 — salaried roster additions and flags. */
export const MIGRATION_R9_SALARIED = 'r9-salaried';
/** Round 9 §13/§18 — retire the two person-specific rate seeds. */
export const MIGRATION_R9_RATE_RULES = 'r9-retire-rate-seeds';
/** Round 10 — corrected rates + the retired crew-lead rate column. */
export const MIGRATION_R10_WAGES = 'r10-wage-seeds';
/** Round 10 — new line items, Kristina's section, no-mileage rows, Pepe's materials. */
export const MIGRATION_R10_ROSTER = 'r10-roster';
/** Round 10 — subcontractor labor out of the materials mapping. */
export const MIGRATION_R10_PL_MAPPING = 'r10-pl-mapping';
/**
 * Round 10 — heal a roster saved while the loader still rejected the
 * `SUBCONTRACTORS` section. Those saves silently dropped Kristina Benson's
 * row, so re-seed any default person the stored roster is missing.
 */
export const MIGRATION_R10_ROSTER_SUBS = 'r10-roster-subs';
/**
 * Round 11 — rates, the Damian Vazquez spelling, Andrew Mazzeo, Color
 * Enterprises at 53%, the JT color, crew-lead aliases, the NEST list, Mike
 * Holden's health insurance as a reimbursement and the holiday pay rows.
 */
export const MIGRATION_R11 = 'r11-settings';
export const ALL_MIGRATIONS: string[] = [
  MIGRATION_RF_COLOR,
  MIGRATION_DB_PATTERN,
  MIGRATION_WAGE_SEEDS,
  MIGRATION_CARLOS_HOURLY,
  MIGRATION_R9_WAGES,
  MIGRATION_R9_LINE_ITEMS,
  MIGRATION_R9_SALARIED,
  MIGRATION_R9_RATE_RULES,
  MIGRATION_R10_WAGES,
  MIGRATION_R10_ROSTER,
  MIGRATION_R10_PL_MAPPING,
  MIGRATION_R10_ROSTER_SUBS,
  MIGRATION_R11,
];

// v4: Round 3 payroll settings added. v3 legend/pivot values are migrated;
// missing payroll sections get the seeds.
const STORAGE_KEY = 'pct-completion-builder.settings.v4';
const STORAGE_KEY_V3 = 'pct-completion-builder.settings.v3';
const MODE_KEY = 'pct-completion-builder.app-mode.v1';

/** Round 6 §1 — %'s Report vs Payroll process. */
export type AppMode = 'percents' | 'payroll';

export function loadAppMode(): AppMode {
  try {
    const raw = localStorage.getItem(MODE_KEY);
    if (raw === 'payroll' || raw === 'percents') return raw;
  } catch {
    // ignore
  }
  return 'percents';
}

export function saveAppMode(mode: AppMode): void {
  try {
    localStorage.setItem(MODE_KEY, mode);
  } catch {
    // ignore
  }
}

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
    conditionalRateRules: DEFAULT_CONDITIONAL_RATE_RULES.map((r) => ({
      ...r,
      condition: { ...r.condition },
    })),
    perJobRateOverrides: DEFAULT_PER_JOB_RATE_OVERRIDES.map((r) => ({ ...r })),
    nvDailyOtRateThreshold: DEFAULT_NV_DAILY_OT_RATE_THRESHOLD,
    dailyOtHoursThreshold: DEFAULT_DAILY_OT_HOURS_THRESHOLD,
    plAccountMapping: clonePlAccountMapping(DEFAULT_PL_ACCOUNT_MAPPING),
    painterRoster: DEFAULT_PAINTER_ROSTER.map((p) => ({ ...p })),
    salariedRoster: DEFAULT_SALARIED.map((e) => ({ ...e })),
    mileageRate: DEFAULT_MILEAGE_RATE,
    officeMileageRate: 0.5,
    subcontractors: [...DEFAULT_SUBCONTRACTORS],
    prevailingWagePatterns: [...DEFAULT_PREVAILING_PATTERNS],
    prevailingMinRate: DEFAULT_PREVAILING_MIN_RATE,
    jobRateRules: [],
    payrollLineItems: DEFAULT_PAYROLL_LINE_ITEMS.map((e) => ({ ...e })),
    fullMaterialsLeads: [...DEFAULT_FULL_MATERIALS_LEADS],
    dailyOtExemptNames: [...DEFAULT_DAILY_OT_EXEMPT],
    crewLeadAliases: DEFAULT_CREW_LEAD_ALIASES.map((a) => ({ ...a })),
    nestNames: [...DEFAULT_NEST_NAMES],
    // Fresh settings already carry every Round 8 seed.
    appliedMigrations: [...ALL_MIGRATIONS],
  };
}

function clonePlAccountMapping(
  src: Record<EntityCode, PlAccountMapping>,
): Record<EntityCode, PlAccountMapping> {
  const one = (m: PlAccountMapping): PlAccountMapping => ({
    income: [...m.income],
    materials: [...m.materials],
    laborWagesSubs: [...m.laborWagesSubs],
    sundries: [...m.sundries],
    travel: [...m.travel],
    mileage: [...m.mileage],
  });
  return { NV: one(src.NV), UT: one(src.UT) };
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
      const x = e as Record<string, unknown>;
      if (typeof x.name === 'string' && typeof x.rate === 'number' && Number.isFinite(x.rate)) {
        list.push({
          name: x.name,
          rate: x.rate,
          ...(typeof x.hourlyRate === 'number' && Number.isFinite(x.hourlyRate) ? { hourlyRate: x.hourlyRate } : {}),
          // Round 9 §18 — the two per-painter pay-type rates. (Round 10 —
          // a stored `crewLeadRate` is intentionally dropped here.)
          ...(typeof x.commercialHourlyRate === 'number' && Number.isFinite(x.commercialHourlyRate)
            ? { commercialHourlyRate: x.commercialHourlyRate }
            : {}),
          ...(typeof x.wages2Rate === 'number' && Number.isFinite(x.wages2Rate)
            ? { wages2Rate: x.wages2Rate }
            : {}),
        });
      }
    }
    if (list.length > 0) out.wageRates = sortWageRates(list);
  }
  if (typeof r.defaultWageRate === 'number' && Number.isFinite(r.defaultWageRate)) {
    out.defaultWageRate = r.defaultWageRate;
  }
  if (typeof r.nvDailyOtRateThreshold === 'number' && Number.isFinite(r.nvDailyOtRateThreshold)) {
    out.nvDailyOtRateThreshold = r.nvDailyOtRateThreshold;
  }
  if (typeof r.dailyOtHoursThreshold === 'number' && Number.isFinite(r.dailyOtHoursThreshold)) {
    out.dailyOtHoursThreshold = r.dailyOtHoursThreshold;
  }
  if (Array.isArray(r.conditionalRateRules)) {
    const rules: ConditionalRateRule[] = [];
    for (const e of r.conditionalRateRules) {
      if (typeof e !== 'object' || e === null) continue;
      const x = e as Record<string, unknown>;
      const cond = x.condition as { kind?: unknown; leadName?: unknown } | undefined;
      if (typeof x.person !== 'string' || typeof x.rate !== 'number' || !Number.isFinite(x.rate)) continue;
      if (!cond || cond.kind !== 'onSameJobAsLead' || typeof cond.leadName !== 'string') continue;
      rules.push({
        id: typeof x.id === 'string' && x.id ? x.id : `cond-${rules.length}`,
        person: x.person,
        condition: { kind: 'onSameJobAsLead', leadName: cond.leadName },
        rate: x.rate,
        ...(typeof x.note === 'string' ? { note: x.note } : {}),
      });
    }
    out.conditionalRateRules = rules;
  }
  if (Array.isArray(r.perJobRateOverrides)) {
    const rules: PerJobRateOverride[] = [];
    for (const e of r.perJobRateOverrides) {
      if (typeof e !== 'object' || e === null) continue;
      const x = e as Record<string, unknown>;
      if (typeof x.person !== 'string' || typeof x.apKeyPattern !== 'string') continue;
      if (typeof x.rate !== 'number' || !Number.isFinite(x.rate)) continue;
      rules.push({
        id: typeof x.id === 'string' && x.id ? x.id : `perjob-${rules.length}`,
        person: x.person,
        apKeyPattern: x.apKeyPattern,
        rate: x.rate,
        ...(typeof x.note === 'string' ? { note: x.note } : {}),
      });
    }
    out.perJobRateOverrides = rules;
  }
  const mapping = parsePlAccountMapping(r.plAccountMapping);
  if (mapping) out.plAccountMapping = mapping;
  if (Array.isArray(r.painterRoster)) {
    const list: PainterRosterEntry[] = [];
    for (const e of r.painterRoster) {
      if (typeof e !== 'object' || e === null) continue;
      const x = e as { name?: unknown; active?: unknown; note?: unknown };
      if (typeof x.name !== 'string' || x.name.trim() === '') continue;
      list.push({
        name: x.name,
        active: x.active === true,
        ...(typeof x.note === 'string' && x.note !== '' ? { note: x.note } : {}),
      });
    }
    out.painterRoster = list;
  }
  if (Array.isArray(r.salariedRoster)) {
    const list: SalariedEntry[] = [];
    for (const e of r.salariedRoster) {
      const x = e as Record<string, unknown>;
      if (
        typeof x.name === 'string' &&
        // Round 10 — `SUBCONTRACTORS` must be listed here too: a section this
        // guard does not know is dropped on load, which is how Kristina
        // Benson's row vanished the first time settings round-tripped.
        (x.section === 'ADMIN' ||
          x.section === 'ESTIMATORS' ||
          x.section === 'MANAGER SALARY' ||
          x.section === 'SUBCONTRACTORS') &&
        typeof x.amount === 'number' &&
        Number.isFinite(x.amount)
      ) {
        list.push({
          name: x.name,
          section: x.section,
          role: typeof x.role === 'string' ? x.role : '',
          amount: x.amount,
          // Round 8 §7.3 — pre-Round-8 rows have no entity; they are NV.
          entity: x.entity === 'UT' ? 'UT' : 'NV',
          // Round 9 §25/§26/§28/§34.
          ...(x.frequency === 'firstPayDateOfMonth' ? { frequency: 'firstPayDateOfMonth' as const } : {}),
          ...(x.sub === true ? { sub: true } : {}),
          ...(x.noMileageRow === true ? { noMileageRow: true } : {}),
          ...(typeof x.reimbursementsLabel === 'string' && x.reimbursementsLabel.trim() !== ''
            ? { reimbursementsLabel: x.reimbursementsLabel }
            : {}),
          ...(typeof x.taxes === 'number' && Number.isFinite(x.taxes) ? { taxes: x.taxes } : {}),
          ...(typeof x.net === 'number' && Number.isFinite(x.net) ? { net: x.net } : {}),
        });
      }
    }
    if (list.length > 0) out.salariedRoster = list;
  }
  if (typeof r.mileageRate === 'number' && Number.isFinite(r.mileageRate)) out.mileageRate = r.mileageRate;
  if (Array.isArray(r.subcontractors)) {
    const subs = r.subcontractors.filter((s): s is string => typeof s === 'string' && s.trim() !== '');
    if (subs.length > 0) out.subcontractors = subs;
  }
  if (Array.isArray(r.prevailingWagePatterns)) {
    const pats = r.prevailingWagePatterns.filter((s): s is string => typeof s === 'string' && s.trim() !== '');
    if (pats.length > 0) out.prevailingWagePatterns = pats;
  }
  if (typeof r.prevailingMinRate === 'number' && Number.isFinite(r.prevailingMinRate)) {
    out.prevailingMinRate = r.prevailingMinRate;
  }
  if (typeof r.officeMileageRate === 'number' && Number.isFinite(r.officeMileageRate)) {
    out.officeMileageRate = r.officeMileageRate;
  }
  if (Array.isArray(r.jobRateRules)) {
    const rules: JobRateRule[] = [];
    for (const raw of r.jobRateRules) {
      if (typeof raw !== 'object' || raw === null) continue;
      const x = raw as Record<string, unknown>;
      if (typeof x.apKey !== 'string' || typeof x.rate !== 'number') continue;
      rules.push({
        id: typeof x.id === 'string' ? x.id : `rule-${rules.length}`,
        apKey: x.apKey.trim(),
        rate: x.rate,
        employeeNames: Array.isArray(x.employeeNames)
          ? x.employeeNames.filter((n): n is string => typeof n === 'string')
          : [],
        includeSuffixes: x.includeSuffixes !== false,
        note: typeof x.note === 'string' ? x.note : '',
      });
    }
    out.jobRateRules = rules;
  }
  if (Array.isArray(r.payrollLineItems)) {
    const list: PayrollLineItem[] = [];
    for (const e of r.payrollLineItems) {
      if (typeof e !== 'object' || e === null) continue;
      const x = e as Record<string, unknown>;
      if (typeof x.person !== 'string' || typeof x.label !== 'string') continue;
      const kind = x.kind === 'deduction' || x.kind === 'reimbursement' ? x.kind : 'note';
      list.push({
        id: typeof x.id === 'string' && x.id ? x.id : `li-${list.length}`,
        person: x.person,
        label: x.label,
        amount: typeof x.amount === 'number' && Number.isFinite(x.amount) ? x.amount : null,
        kind,
        frequency: x.frequency === 'firstPayDateOfMonth' ? 'firstPayDateOfMonth' : 'weekly',
      });
    }
    out.payrollLineItems = list;
  }
  if (Array.isArray(r.fullMaterialsLeads)) {
    out.fullMaterialsLeads = r.fullMaterialsLeads.filter(
      (n): n is string => typeof n === 'string' && n.trim() !== '',
    );
  }
  if (Array.isArray(r.dailyOtExemptNames)) {
    out.dailyOtExemptNames = r.dailyOtExemptNames.filter(
      (n): n is string => typeof n === 'string' && n.trim() !== '',
    );
  }
  if (Array.isArray(r.crewLeadAliases)) {
    const list: CrewLeadAlias[] = [];
    for (const e of r.crewLeadAliases) {
      const x = e as Record<string, unknown>;
      if (typeof x.alias === 'string' && typeof x.lead === 'string' && x.alias.trim() && x.lead.trim()) {
        list.push({ alias: x.alias.trim(), lead: x.lead.trim() });
      }
    }
    out.crewLeadAliases = list;
  }
  if (Array.isArray(r.nestNames)) {
    out.nestNames = r.nestNames.filter((n): n is string => typeof n === 'string' && n.trim() !== '');
  }
  if (Array.isArray(r.appliedMigrations)) {
    out.appliedMigrations = r.appliedMigrations.filter((s): s is string => typeof s === 'string');
  } else {
    // Pre-Round-8 blob: nothing has been migrated yet.
    out.appliedMigrations = [];
  }
  return out;
}

/** Validate a persisted `plAccountMapping`; unknown metrics/entities are dropped. */
function parsePlAccountMapping(raw: unknown): Record<EntityCode, PlAccountMapping> | null {
  if (typeof raw !== 'object' || raw === null) return null;
  const src = raw as Record<string, unknown>;
  const metrics = ['income', 'materials', 'laborWagesSubs', 'sundries', 'travel', 'mileage'] as const;
  const one = (entity: EntityCode): PlAccountMapping | null => {
    const e = src[entity];
    if (typeof e !== 'object' || e === null) return null;
    const rec = e as Record<string, unknown>;
    const fallback = DEFAULT_PL_ACCOUNT_MAPPING[entity];
    const out = {} as PlAccountMapping;
    for (const m of metrics) {
      const v = rec[m];
      out[m] = Array.isArray(v)
        ? v.filter((s): s is string => typeof s === 'string' && s.trim() !== '')
        : [...fallback[m]];
    }
    return out;
  };
  const nv = one('NV');
  const ut = one('UT');
  if (!nv && !ut) return null;
  return {
    NV: nv ?? clonePlAccountMapping(DEFAULT_PL_ACCOUNT_MAPPING).NV,
    UT: ut ?? clonePlAccountMapping(DEFAULT_PL_ACCOUNT_MAPPING).UT,
  };
}

/**
 * Validate any persisted settings payload (localStorage or the shared store)
 * and fill every missing section from the seeds. Returns null when the payload
 * is not a settings object at all. Round 8 migrations are applied here, so both
 * storage paths heal identically.
 */
export function parseSettings(raw: unknown): AppSettings | null {
  const legend = parseLegend(raw);
  if (!legend) return null;
  const pivotColumnOrder = ensureDefaultPivotColumns(parsePivotOrder(raw) ?? [...DEFAULT_PIVOT_COLUMNS]);
  return applySettingsMigrations({
    ...defaultSettings(),
    ...parsePayroll(raw),
    legend,
    pivotColumnOrder,
  });
}

/**
 * Round 8 one-time data migrations, each gated by an id in `appliedMigrations`:
 * §5 RF `DA9694` → `C4A484`, §9.1 Davis-Bacon `SNRHA *SS*` → `SNRHA`, the
 * §9.3 three-rate wage seeds merged by name, and Carlos Ruiz Cruz hourly $25.
 */
export function applySettingsMigrations(settings: AppSettings): AppSettings {
  const done = new Set(settings.appliedMigrations);
  let next = settings;
  const patch = (p: Partial<AppSettings>): void => {
    next = { ...next, ...p };
  };

  if (!done.has(MIGRATION_RF_COLOR)) {
    // Only the exact old seed moves; a custom RF pick is left alone.
    patch({
      legend: next.legend.map((l) =>
        l.estimator === 'RF' && l.color.toUpperCase() === LEGACY_RF_COLOR ? { ...l, color: 'C4A484' } : l,
      ),
    });
    done.add(MIGRATION_RF_COLOR);
  }

  if (!done.has(MIGRATION_DB_PATTERN)) {
    const broadened: string[] = [];
    for (const p of next.prevailingWagePatterns) {
      const v = p.trim() === LEGACY_PREVAILING_PATTERN ? 'SNRHA' : p;
      if (!broadened.includes(v)) broadened.push(v);
    }
    patch({ prevailingWagePatterns: broadened });
    done.add(MIGRATION_DB_PATTERN);
  }

  if (!done.has(MIGRATION_WAGE_SEEDS)) {
    patch({
      wageRates: mergeWageRateSeeds(next.wageRates, ROUND8_WAGE_SEEDS, 'fillMissing'),
      conditionalRateRules:
        next.conditionalRateRules.length > 0
          ? next.conditionalRateRules
          : DEFAULT_CONDITIONAL_RATE_RULES.map((r) => ({ ...r, condition: { ...r.condition } })),
      perJobRateOverrides:
        next.perJobRateOverrides.length > 0
          ? next.perJobRateOverrides
          : DEFAULT_PER_JOB_RATE_OVERRIDES.map((r) => ({ ...r })),
    });
    done.add(MIGRATION_WAGE_SEEDS);
  }

  // Round 9 §12/§17/§18/§21/§22/§35 — corrected rates WIN over stored values,
  // because they are corrections to numbers the app had wrong.
  if (!done.has(MIGRATION_R9_WAGES)) {
    patch({ wageRates: mergeWageRateSeeds(next.wageRates, ROUND9_WAGE_SEEDS, 'seedWins') });
    done.add(MIGRATION_R9_WAGES);
  }

  // Round 9 §19/§23/§24/§29 — add the seeded line items that are missing by id.
  if (!done.has(MIGRATION_R9_LINE_ITEMS)) {
    const have = new Set((next.payrollLineItems ?? []).map((i) => i.id));
    patch({
      payrollLineItems: [
        ...(next.payrollLineItems ?? []),
        ...DEFAULT_PAYROLL_LINE_ITEMS.filter((i) => !have.has(i.id)).map((i) => ({ ...i })),
      ],
      fullMaterialsLeads:
        next.fullMaterialsLeads && next.fullMaterialsLeads.length > 0
          ? next.fullMaterialsLeads
          : [...DEFAULT_FULL_MATERIALS_LEADS],
      dailyOtExemptNames:
        next.dailyOtExemptNames && next.dailyOtExemptNames.length > 0
          ? next.dailyOtExemptNames
          : [...DEFAULT_DAILY_OT_EXEMPT],
    });
    done.add(MIGRATION_R9_LINE_ITEMS);
  }

  // Round 9 §25/§26/§28/§34 — salaried additions + flags, matched by name.
  if (!done.has(MIGRATION_R9_SALARIED)) {
    const key = (n: string): string => n.trim().toLowerCase().replace(/\s+/g, ' ');
    const bySeed = new Map(DEFAULT_SALARIED.map((e) => [key(e.name), e]));
    const patched = next.salariedRoster.map((row) => {
      const seed = bySeed.get(key(row.name));
      if (!seed) return row;
      return {
        ...row,
        ...(seed.frequency ? { frequency: seed.frequency } : {}),
        ...(seed.sub ? { sub: true } : {}),
        ...(seed.reimbursementsLabel ? { reimbursementsLabel: seed.reimbursementsLabel } : {}),
        ...(seed.taxes !== undefined ? { taxes: seed.taxes } : {}),
        ...(seed.net !== undefined ? { net: seed.net } : {}),
        ...(seed.entity ? { entity: seed.entity } : {}),
      };
    });
    const haveNames = new Set(patched.map((e) => key(e.name)));
    patch({
      salariedRoster: [
        ...patched,
        ...DEFAULT_SALARIED.filter((e) => !haveNames.has(key(e.name))).map((e) => ({ ...e })),
      ],
    });
    done.add(MIGRATION_R9_SALARIED);
  }

  // Round 9 §13/§18 — retire the two person-specific rate seeds (the rules
  // themselves stay available; only these seeded rows go).
  if (!done.has(MIGRATION_R9_RATE_RULES)) {
    patch({
      conditionalRateRules: next.conditionalRateRules.filter(
        (r) => r.id !== 'cond-henry-limay-flores-miguel-martin',
      ),
      perJobRateOverrides: next.perJobRateOverrides.filter(
        (r) => r.id !== 'perjob-samuel-brewster-24097',
      ),
    });
    done.add(MIGRATION_R9_RATE_RULES);
  }

  /**
   * Round 10 — the corrected rates win over stored values (they are fixes to
   * numbers the app had wrong), and the retired crew-lead column is dropped:
   * a crew lead on their own job is on a profit-share job, so the PS rate is
   * their crew-lead rate.
   */
  if (!done.has(MIGRATION_R10_WAGES)) {
    const merged = mergeWageRateSeeds(next.wageRates, ROUND10_WAGE_SEEDS, 'seedWins');
    patch({
      wageRates: merged.map((row) => {
        const { crewLeadRate: _dropped, ...rest } = row as WageRateEntry & { crewLeadRate?: number };
        return rest;
      }),
    });
    done.add(MIGRATION_R10_WAGES);
  }

  // Round 10 — roster/line-item changes, matched by id and by name.
  if (!done.has(MIGRATION_R10_ROSTER)) {
    const key = (n: string): string => n.trim().toLowerCase().replace(/\s+/g, ' ');
    const haveItems = new Set((next.payrollLineItems ?? []).map((i) => i.id));
    const seedByName = new Map(DEFAULT_SALARIED.map((e) => [key(e.name), e]));
    patch({
      payrollLineItems: [
        ...(next.payrollLineItems ?? []),
        ...DEFAULT_PAYROLL_LINE_ITEMS.filter((i) => !haveItems.has(i.id)).map((i) => ({ ...i })),
      ],
      salariedRoster: next.salariedRoster.map((row) => {
        const seed = seedByName.get(key(row.name));
        if (!seed) return row;
        return {
          ...row,
          section: seed.section,
          ...(seed.noMileageRow ? { noMileageRow: true } : {}),
        };
      }),
      fullMaterialsLeads: [
        ...next.fullMaterialsLeads,
        ...DEFAULT_FULL_MATERIALS_LEADS.filter(
          (n) => !next.fullMaterialsLeads.some((h) => key(h) === key(n)),
        ),
      ],
    });
    done.add(MIGRATION_R10_ROSTER);
  }

  /**
   * Round 10 — a saved mapping that listed a labor account (e.g.
   * `Subcontractor Labor`) under materials made a subcontractor's ½ Materials
   * line swallow their own pay. Labor labels are removed from materials, and
   * the default labor labels are re-seeded when they went missing.
   */
  if (!done.has(MIGRATION_R10_PL_MAPPING)) {
    const key = (n: string): string => n.trim().toLowerCase();
    const healed = { ...next.plAccountMapping };
    for (const entity of ['NV', 'UT'] as EntityCode[]) {
      const m = healed[entity];
      const defaults = DEFAULT_PL_ACCOUNT_MAPPING[entity];
      const labor = [
        ...m.laborWagesSubs,
        ...defaults.laborWagesSubs.filter(
          (d) => !m.laborWagesSubs.some((h) => key(h) === key(d)),
        ),
      ];
      const laborKeys = new Set(labor.map(key));
      healed[entity] = {
        ...m,
        laborWagesSubs: labor,
        materials: m.materials.filter((l) => !laborKeys.has(key(l))),
      };
    }
    patch({ plAccountMapping: healed });
    done.add(MIGRATION_R10_PL_MAPPING);
  }

  // Round 10 — put back roster rows an older loader dropped (see the constant).
  if (!done.has(MIGRATION_R10_ROSTER_SUBS)) {
    const key = (n: string): string => n.trim().toLowerCase().replace(/\s+/g, ' ');
    const have = new Set(next.salariedRoster.map((r) => key(r.name)));
    const missing = DEFAULT_SALARIED.filter((e) => !have.has(key(e.name))).map((e) => ({ ...e }));
    if (missing.length > 0) patch({ salariedRoster: [...next.salariedRoster, ...missing] });
    done.add(MIGRATION_R10_ROSTER_SUBS);
  }

  if (!done.has(MIGRATION_CARLOS_HOURLY)) {
    patch({
      wageRates: next.wageRates.map((row) =>
        /^carlos\s+ruiz(\s+cruz)?$/i.test(row.name.trim()) && (row.hourlyRate === undefined || row.hourlyRate === 18)
          ? { ...row, hourlyRate: 25 }
          : row,
      ),
    });
    done.add(MIGRATION_CARLOS_HOURLY);
  }

  /**
   * Round 11. These are explicit corrections, so they overwrite what is saved
   * (`seedWins`) rather than only filling gaps — a stale shared-store value
   * would otherwise keep paying the old rate.
   */
  if (!done.has(MIGRATION_R11)) {
    const key = (n: string): string => n.trim().toLowerCase().replace(/\s+/g, ' ');
    const respell = (n: string): string => fixNameSpellingInText(n);

    const wageRates = sortWageRates(
      mergeWageRateSeeds(
        // Rename first, so the seed for "Damian Vazquez" patches his row
        // instead of appending a duplicate.
        next.wageRates.map((w) => ({ ...w, name: respell(w.name) })),
        ROUND11_WAGE_SEEDS,
        'seedWins',
      ),
    );

    let profitShare = next.profitShare.map((p) => ({ ...p, lead: respell(p.lead) }));
    profitShare = profitShare.map((p) => (key(p.lead) === 'color enterprises llc' ? { ...p, pct: 0.53 } : p));
    if (!profitShare.some((p) => key(p.lead) === 'andrew mazzeo')) {
      profitShare.push({ lead: 'Andrew Mazzeo', pct: next.defaultProfitSharePct });
    }

    const legend = next.legend.some((l) => l.estimator === 'JT')
      ? next.legend.map((l) => (l.estimator === 'JT' ? { ...l, color: 'DDD9C4' } : l))
      : [...next.legend, { estimator: 'JT', color: 'DDD9C4' }];

    const haveItems = new Set(next.payrollLineItems.map((i) => i.id));
    const payrollLineItems = [
      ...next.payrollLineItems.map((i) => ({
        ...i,
        person: respell(i.person),
        // Paid TO Mike Holden, not withheld.
        ...(i.id === 'li-mike-holden-health' ? { kind: 'reimbursement' as const } : {}),
        // The row covers additional pay as well.
        ...(i.label.trim() === 'Reimbursements' ? { label: 'Reimbursements/additional pay' } : {}),
      })),
      ...DEFAULT_PAYROLL_LINE_ITEMS.filter((i) => !haveItems.has(i.id)).map((i) => ({ ...i })),
    ];

    const haveAliases = new Set((next.crewLeadAliases ?? []).map((a) => key(a.alias)));
    const crewLeadAliases = [
      ...(next.crewLeadAliases ?? []),
      ...DEFAULT_CREW_LEAD_ALIASES.filter((a) => !haveAliases.has(key(a.alias))).map((a) => ({ ...a })),
    ];

    patch({
      wageRates,
      profitShare,
      legend,
      payrollLineItems,
      crewLeadAliases,
      nestNames: next.nestNames && next.nestNames.length > 0 ? next.nestNames : [...DEFAULT_NEST_NAMES],
      painterRoster: next.painterRoster.map((r) => ({ ...r, name: respell(r.name) })),
      salariedRoster: next.salariedRoster.map((r) => ({
        ...r,
        name: respell(r.name),
        ...(r.reimbursementsLabel?.trim() === 'Reimbursements'
          ? { reimbursementsLabel: 'Reimbursements/additional pay' }
          : {}),
      })),
      conditionalRateRules: next.conditionalRateRules.map((r) => ({
        ...r,
        person: respell(r.person),
        condition: { ...r.condition, leadName: respell(r.condition.leadName) },
      })),
      perJobRateOverrides: next.perJobRateOverrides.map((r) => ({ ...r, person: respell(r.person) })),
      subcontractors: next.subcontractors.map(respell),
      fullMaterialsLeads: next.fullMaterialsLeads.map(respell),
      dailyOtExemptNames: next.dailyOtExemptNames.map(respell),
    });
    done.add(MIGRATION_R11);
  }

  // Whatever path a row took in, the list stays alphabetical.
  next = { ...next, wageRates: sortWageRates(next.wageRates) };

  return { ...next, appliedMigrations: [...done] };
}

export function loadSettings(): AppSettings {
  try {
    // v4 first; fall back to v3 (legend/pivot only) so users keep palette edits.
    const raw = localStorage.getItem(STORAGE_KEY) ?? localStorage.getItem(STORAGE_KEY_V3);
    if (raw) {
      const parsed = parseSettings(JSON.parse(raw) as unknown);
      if (parsed) return parsed;
    }
  } catch {
    // fall through
  }
  return defaultSettings();
}

/** True when this browser has its own saved settings (drives the §1 push-up). */
export function hasLocalSettings(): boolean {
  try {
    return (
      localStorage.getItem(STORAGE_KEY) !== null || localStorage.getItem(STORAGE_KEY_V3) !== null
    );
  } catch {
    return false;
  }
}

export function saveSettings(settings: AppSettings): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // localStorage unavailable — settings just won't persist.
  }
}

// ---------------------------------------------------------------------------
// Round 8 §1 — shared persisted Settings store (same pattern as the directory)
// ---------------------------------------------------------------------------

/** Whether the shared (cross-device) settings store is wired up on this host. */
export type SettingsStoreStatus = 'ready' | 'empty' | 'unconfigured' | 'error';

export type SettingsPersistResult = { ok: true } | { ok: false; status: number; message: string };

/** Envelope written to the shared store; `savedAt` drives last-write-wins. */
export interface StoredSettings {
  settings: AppSettings;
  savedAt: string; // ISO timestamp
}

export async function probeSettingsStore(): Promise<SettingsStoreStatus> {
  try {
    const res = await fetch('/api/settings', { cache: 'no-store' });
    if (res.status === 503) return 'unconfigured';
    if (res.status === 404) return 'empty';
    if (!res.ok) return 'error';
    // Dev/static hosts answer unknown paths with index.html — that is not a store.
    try {
      await res.json();
    } catch {
      return 'unconfigured';
    }
    return 'ready';
  } catch {
    return 'error';
  }
}

/** Accept either the `{ settings, savedAt }` envelope or a bare AppSettings blob. */
function parseStoredSettings(raw: unknown): AppSettings | null {
  if (typeof raw === 'object' && raw !== null && 'settings' in raw) {
    return parseSettings((raw as { settings: unknown }).settings);
  }
  return parseSettings(raw);
}

/**
 * Load the shared Settings (Vercel Blob via /api). This is what lets a cleared
 * browser or incognito window see the same saved Settings.
 */
export async function fetchAppSettings(): Promise<AppSettings | null> {
  try {
    const res = await fetch('/api/settings', { cache: 'no-store' });
    if (res.ok) return parseStoredSettings((await res.json()) as unknown);
  } catch {
    // fall through
  }
  // Local vite / static fallback (not used for cross-device on Vercel).
  try {
    const res = await fetch('/app-settings.json', { cache: 'no-store' });
    if (res.ok) return parseStoredSettings((await res.json()) as unknown);
  } catch {
    // ignore
  }
  return null;
}

/** Write the current Settings to the shared store (last write wins). */
export async function persistAppSettings(settings: AppSettings): Promise<SettingsPersistResult> {
  const payload: StoredSettings = { settings, savedAt: new Date().toISOString() };
  try {
    const res = await fetch('/api/settings', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (res.ok) return { ok: true };
    let message = `Could not sync Settings (HTTP ${res.status}).`;
    try {
      const body = (await res.json()) as { message?: string };
      if (body.message) message = body.message;
    } catch {
      // ignore
    }
    return { ok: false, status: res.status, message };
  } catch {
    return {
      ok: false,
      status: 0,
      message:
        'Shared sync unavailable — your Settings are still saved in this browser. Cross-device sync is optional.',
    };
  }
}

export async function clearAppSettings(): Promise<void> {
  try {
    await fetch('/api/settings', { method: 'DELETE' });
  } catch {
    // ignore
  }
}

// ---------------------------------------------------------------------------
// Round 4 §4.2 — persisted employee directory (same pattern as the legend)
// ---------------------------------------------------------------------------

const DIRECTORY_KEY = 'pct-completion-builder.employee-directory.v1';

export function loadSavedDirectory(): SavedDirectory | null {
  try {
    const raw = localStorage.getItem(DIRECTORY_KEY);
    if (!raw) return null;
    return parseSavedDirectory(JSON.parse(raw));
  } catch {
    return null;
  }
}

/** Write an already-built SavedDirectory into this browser's localStorage (keeps uploadedAt). */
export function cacheDirectoryLocally(saved: SavedDirectory): void {
  try {
    localStorage.setItem(DIRECTORY_KEY, JSON.stringify(saved));
  } catch {
    // localStorage unavailable
  }
}

export function saveDirectory(employees: DirectoryEmployee[], sourceFilename: string): SavedDirectory {
  const saved: SavedDirectory = {
    employees,
    sourceFilename,
    uploadedAt: new Date().toISOString(),
    activeCount: employees.filter((e) => e.active).length,
    inactiveCount: employees.filter((e) => !e.active).length,
  };
  cacheDirectoryLocally(saved);
  return saved;
}

export function clearSavedDirectory(): void {
  try {
    localStorage.removeItem(DIRECTORY_KEY);
  } catch {
    // ignore
  }
}

/** Parse/validate a SavedDirectory payload (localStorage or app file). */
export function parseSavedDirectory(raw: unknown): SavedDirectory | null {
  if (typeof raw !== 'object' || raw === null) return null;
  const p = raw as Record<string, unknown>;
  if (!Array.isArray(p.employees)) return null;
  const employees: DirectoryEmployee[] = [];
  for (const e of p.employees) {
    const x = e as { fullName?: unknown; firstName?: unknown; lastName?: unknown; active?: unknown };
    if (typeof x.fullName !== 'string' || x.fullName.trim() === '') continue;
    employees.push({
      fullName: x.fullName,
      firstName: typeof x.firstName === 'string' ? x.firstName : '',
      lastName: typeof x.lastName === 'string' ? x.lastName : '',
      active: x.active !== false,
    });
  }
  if (employees.length === 0) return null;
  return {
    employees,
    sourceFilename: typeof p.sourceFilename === 'string' ? p.sourceFilename : '',
    uploadedAt: typeof p.uploadedAt === 'string' ? p.uploadedAt : '',
    activeCount: employees.filter((e) => e.active).length,
    inactiveCount: employees.filter((e) => !e.active).length,
  };
}

/** Prefer the newer of two saved directories by `uploadedAt`. */
export function newerDirectory(a: SavedDirectory | null, b: SavedDirectory | null): SavedDirectory | null {
  if (!a) return b;
  if (!b) return a;
  return (a.uploadedAt || '') >= (b.uploadedAt || '') ? a : b;
}

export type DirectoryPersistResult =
  | { ok: true }
  | { ok: false; status: number; message: string };

/** Whether the shared (cross-device) store is wired up on this host. */
export type DirectoryStoreStatus = 'ready' | 'empty' | 'unconfigured' | 'error';

export async function probeDirectoryStore(): Promise<DirectoryStoreStatus> {
  try {
    const res = await fetch('/api/employee-directory', { cache: 'no-store' });
    if (res.status === 503) return 'unconfigured';
    if (res.status === 404) return 'empty';
    if (res.ok) return 'ready';
    return 'error';
  } catch {
    return 'error';
  }
}

/**
 * Load the shared employee list (Vercel Blob via /api, or local public file).
 * This is what makes phone + desktop see the same list on the Vercel URL.
 */
export async function fetchAppDirectory(): Promise<SavedDirectory | null> {
  try {
    const res = await fetch('/api/employee-directory', { cache: 'no-store' });
    if (res.ok) return parseSavedDirectory(await res.json());
  } catch {
    // fall through
  }
  // Local vite / static fallback (not used for cross-device on Vercel).
  try {
    const res = await fetch('/employee-directory.json', { cache: 'no-store' });
    if (res.ok) return parseSavedDirectory(await res.json());
  } catch {
    // ignore
  }
  return null;
}

/** Write the current list to the shared store (Vercel Blob in production). */
export async function persistAppDirectory(saved: SavedDirectory): Promise<DirectoryPersistResult> {
  try {
    const res = await fetch('/api/employee-directory', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(saved),
    });
    if (res.ok) return { ok: true };
    let message = `Could not sync employee list (HTTP ${res.status}).`;
    try {
      const body = (await res.json()) as { message?: string };
      if (body.message) message = body.message;
    } catch {
      // ignore
    }
    return { ok: false, status: res.status, message };
  } catch {
    return {
      ok: false,
      status: 0,
      message:
        'Shared sync unavailable — your list is still saved in this browser. Cross-device sync is optional.',
    };
  }
}

export async function clearAppDirectory(): Promise<void> {
  try {
    await fetch('/api/employee-directory', { method: 'DELETE' });
  } catch {
    // ignore
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
