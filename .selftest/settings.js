"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ALL_MIGRATIONS = exports.MIGRATION_R9_RATE_RULES = exports.MIGRATION_R9_SALARIED = exports.MIGRATION_R9_LINE_ITEMS = exports.MIGRATION_R9_WAGES = exports.MIGRATION_CARLOS_HOURLY = exports.MIGRATION_WAGE_SEEDS = exports.MIGRATION_DB_PATTERN = exports.MIGRATION_RF_COLOR = exports.LEGACY_PREVAILING_PATTERN = exports.DEFAULT_PREVAILING_MIN_RATE = exports.DEFAULT_PREVAILING_PATTERNS = exports.DEFAULT_SUBCONTRACTORS = exports.DEFAULT_MILEAGE_RATE = exports.DEFAULT_DAILY_OT_EXEMPT = exports.DEFAULT_FULL_MATERIALS_LEADS = exports.DEFAULT_PAYROLL_LINE_ITEMS = exports.DEFAULT_SALARIED = exports.DEFAULT_PAINTER_ROSTER = exports.DEFAULT_PL_ACCOUNT_MAPPING = exports.DEFAULT_DAILY_OT_HOURS_THRESHOLD = exports.DEFAULT_NV_DAILY_OT_RATE_THRESHOLD = exports.DEFAULT_PER_JOB_RATE_OVERRIDES = exports.DEFAULT_CONDITIONAL_RATE_RULES = exports.DEFAULT_CREW_LEAD_RATE = exports.DEFAULT_WAGE_RATE = exports.DEFAULT_WAGE_RATES = exports.ROUND9_WAGE_SEEDS = exports.ROUND8_WAGE_SEEDS = exports.DEFAULT_PROFIT_SHARE_PCT = exports.DEFAULT_PROFIT_SHARE = exports.DEFAULT_PIVOT_COLUMNS = exports.LEGACY_RF_COLOR = exports.DEFAULT_LEGEND = void 0;
exports.mergeWageRateSeeds = mergeWageRateSeeds;
exports.loadAppMode = loadAppMode;
exports.saveAppMode = saveAppMode;
exports.ensureDefaultPivotColumns = ensureDefaultPivotColumns;
exports.defaultSettings = defaultSettings;
exports.parseSettings = parseSettings;
exports.applySettingsMigrations = applySettingsMigrations;
exports.loadSettings = loadSettings;
exports.hasLocalSettings = hasLocalSettings;
exports.saveSettings = saveSettings;
exports.probeSettingsStore = probeSettingsStore;
exports.fetchAppSettings = fetchAppSettings;
exports.persistAppSettings = persistAppSettings;
exports.clearAppSettings = clearAppSettings;
exports.loadSavedDirectory = loadSavedDirectory;
exports.cacheDirectoryLocally = cacheDirectoryLocally;
exports.saveDirectory = saveDirectory;
exports.clearSavedDirectory = clearSavedDirectory;
exports.parseSavedDirectory = parseSavedDirectory;
exports.newerDirectory = newerDirectory;
exports.probeDirectoryStore = probeDirectoryStore;
exports.fetchAppDirectory = fetchAppDirectory;
exports.persistAppDirectory = persistAppDirectory;
exports.clearAppDirectory = clearAppDirectory;
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
    { estimator: 'RF', color: 'C4A484' },
];
/** Round 8 §5 — the RF fill this seed replaced; migrated once on load. */
exports.LEGACY_RF_COLOR = 'DA9694';
/** Default pivot estimator column order (separate from legend row-sort order). */
exports.DEFAULT_PIVOT_COLUMNS = ['MH', 'PG', 'XC', 'ND', 'TE', 'RF'];
/**
 * Round 3 payroll seeds. These are tribal-knowledge DEFAULTS from the
 * reference documents — fully editable in Settings, never logic.
 */
exports.DEFAULT_PROFIT_SHARE = [
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
exports.DEFAULT_PROFIT_SHARE_PCT = 0.38;
const BASE_WAGE_RATES = [
    { name: 'Adriana Lopez Barajas', rate: 25, hourlyRate: 20 },
    { name: 'Ana Bedoya Melgar', rate: 20 },
    { name: 'Antonio Zuniga', rate: 18, hourlyRate: 25 },
    { name: 'Arlette Zuniga Lopez', rate: 19.18 },
    { name: 'Armando Alfonso Marin', rate: 15 },
    { name: 'Calixto Efrain Climaco', rate: 19.18 },
    { name: 'Jesus Perez', rate: 19.18 },
    { name: 'Carlos Ruiz Cruz', rate: 18, hourlyRate: 25 },
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
/**
 * Round 8 §9.1/§9.3 — three-rate seeds ("hourly" = hourlyRate, "PS" = rate,
 * "crew-lead" = crewLeadRate). Merged by name into BASE_WAGE_RATES: matching
 * names are patched field-by-field, new names are appended. SEEDS ONLY —
 * fully editable in Settings, never referenced as constants by rate logic.
 */
exports.ROUND8_WAGE_SEEDS = [
    { name: 'Adriana Lopez Barajas', rate: 25, hourlyRate: 20 },
    { name: 'Antonio Zuniga', rate: 18, hourlyRate: 25, crewLeadRate: 18 },
    { name: 'Carlos Ruiz Cruz', rate: 18, hourlyRate: 25, crewLeadRate: 18 },
    { name: 'Damian Vasquez', rate: 18, hourlyRate: 18 },
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
exports.ROUND9_WAGE_SEEDS = [
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
/** Name key for wage-map merges (case/whitespace-insensitive). */
function wageKey(name) {
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
function mergeWageRateSeeds(current, seeds, mode = 'seedWins') {
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
                    ...(row.crewLeadRate === undefined && seed.crewLeadRate !== undefined
                        ? { crewLeadRate: seed.crewLeadRate }
                        : {}),
                };
    }
    return out;
}
exports.DEFAULT_WAGE_RATES = mergeWageRateSeeds(mergeWageRateSeeds(BASE_WAGE_RATES, exports.ROUND8_WAGE_SEEDS), exports.ROUND9_WAGE_SEEDS);
exports.DEFAULT_WAGE_RATE = 18;
/** Round 8 §9.1 — fallback lead rate on own timesheet hours on PS jobs. */
exports.DEFAULT_CREW_LEAD_RATE = 18;
/**
 * Round 8 §9.1 — seeded conditional rate rules.
 *
 * Round 9 §13 — the "Henry Limay Flores on Miguel Martin's job" rule is GONE:
 * when the two work the same job, the crew lead is whoever the Pivot by Job
 * `Crew Lead` column names, and the painter who clocked
 * "Crew Lead - Hourly Wage" is the lead for that job. The rule MECHANISM stays
 * available in Settings for genuinely negotiated cases.
 */
exports.DEFAULT_CONDITIONAL_RATE_RULES = [];
/**
 * Round 8 §9.1 — seeded person + AP# rate overrides.
 *
 * Round 9 §18 — Samuel Brewster's $30 is no longer pinned to AP24097: he is
 * paid $30 wherever he clocks the "Commercial Hourly" pay type, which is what
 * his timesheet actually says. The override MECHANISM is untouched.
 */
exports.DEFAULT_PER_JOB_RATE_OVERRIDES = [];
/** Round 8 §9.2 — daily-OT thresholds (logic lives in the payroll engine). */
exports.DEFAULT_NV_DAILY_OT_RATE_THRESHOLD = 18;
exports.DEFAULT_DAILY_OT_HOURS_THRESHOLD = 8;
/**
 * Round 8 §7.1 — per-entity P&L account label mapping. Sundries stays OUT of
 * materials (§8.3 supersedes Round 3 §2).
 */
exports.DEFAULT_PL_ACCOUNT_MAPPING = {
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
exports.DEFAULT_PAINTER_ROSTER = [
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
exports.DEFAULT_SALARIED = [
    // Round 9 §28 — Jodi's mileage row reads "Reimbursements" and is a flat cell.
    {
        name: 'Jodi Dunn',
        section: 'ADMIN',
        role: 'Manager Salary',
        amount: 865,
        entity: 'NV',
        reimbursementsLabel: 'Reimbursements',
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
    // Round 9 §26/§33 — subcontractor; renders at the very bottom of the section.
    {
        name: 'Kristina Benson',
        section: 'MANAGER SALARY',
        role: 'Subcontractor',
        amount: 150,
        entity: 'NV',
        sub: true,
    },
];
/**
 * Round 9 §19/§23/§24/§29 — person-scoped extra rows on the Payroll Preview.
 * `deduction` comes OUT of the person's total, `reimbursement` goes in, and
 * `note` is a labeled blank line for the payroll clerk to fill in.
 * SEED DATA — fully editable in Settings.
 */
exports.DEFAULT_PAYROLL_LINE_ITEMS = [
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
        kind: 'deduction',
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
    ].map((person) => ({
        id: `li-${person.toLowerCase().replace(/\s+/g, '-')}-reimb`,
        person,
        label: 'Reimbursements',
        amount: null,
        kind: 'reimbursement',
        frequency: 'weekly',
    })),
];
/** Round 9 §32 — leads charged for ALL of a job's materials, not half. */
exports.DEFAULT_FULL_MATERIALS_LEADS = ['Color Enterprises LLC'];
/** Round 9 §27 — people the NV daily-overtime rule never applies to. */
exports.DEFAULT_DAILY_OT_EXEMPT = ['Tyler Dunn'];
exports.DEFAULT_MILEAGE_RATE = 0.625;
exports.DEFAULT_SUBCONTRACTORS = [
    "Pepe's Painting LLC",
    'Color Enterprises LLC',
    'Merino Custom Painting',
];
/**
 * Round 4 §5.8 / Round 8 §9.1 — prevailing-wage (Davis-Bacon) job patterns +
 * minimum rate. Broadened from `SNRHA *SS*` to the bare `SNRHA` substring so
 * every SNRHA job matches regardless of the `*SS*` marker.
 */
exports.DEFAULT_PREVAILING_PATTERNS = ['SNRHA'];
exports.DEFAULT_PREVAILING_MIN_RATE = 19.18;
/** The narrow Round 4 pattern replaced by `SNRHA`; migrated once on load. */
exports.LEGACY_PREVAILING_PATTERN = 'SNRHA *SS*';
// One-time settings migrations (Round 8). Each id is recorded in
// AppSettings.appliedMigrations so it never runs twice on the same blob.
exports.MIGRATION_RF_COLOR = 'r8-rf-color';
exports.MIGRATION_DB_PATTERN = 'r8-db-pattern';
exports.MIGRATION_WAGE_SEEDS = 'r8-wage-seeds';
exports.MIGRATION_CARLOS_HOURLY = 'r8-carlos-hourly-25';
/** Round 9 — corrected wage rates + the new pay-type rates (§12/§17/§18/§21/§22/§35). */
exports.MIGRATION_R9_WAGES = 'r9-wage-seeds';
/** Round 9 §19/§23/§24/§29 — seeded payroll line items. */
exports.MIGRATION_R9_LINE_ITEMS = 'r9-line-items';
/** Round 9 §25/§26/§28/§34 — salaried roster additions and flags. */
exports.MIGRATION_R9_SALARIED = 'r9-salaried';
/** Round 9 §13/§18 — retire the two person-specific rate seeds. */
exports.MIGRATION_R9_RATE_RULES = 'r9-retire-rate-seeds';
exports.ALL_MIGRATIONS = [
    exports.MIGRATION_RF_COLOR,
    exports.MIGRATION_DB_PATTERN,
    exports.MIGRATION_WAGE_SEEDS,
    exports.MIGRATION_CARLOS_HOURLY,
    exports.MIGRATION_R9_WAGES,
    exports.MIGRATION_R9_LINE_ITEMS,
    exports.MIGRATION_R9_SALARIED,
    exports.MIGRATION_R9_RATE_RULES,
];
// v4: Round 3 payroll settings added. v3 legend/pivot values are migrated;
// missing payroll sections get the seeds.
const STORAGE_KEY = 'pct-completion-builder.settings.v4';
const STORAGE_KEY_V3 = 'pct-completion-builder.settings.v3';
const MODE_KEY = 'pct-completion-builder.app-mode.v1';
function loadAppMode() {
    try {
        const raw = localStorage.getItem(MODE_KEY);
        if (raw === 'payroll' || raw === 'percents')
            return raw;
    }
    catch {
        // ignore
    }
    return 'percents';
}
function saveAppMode(mode) {
    try {
        localStorage.setItem(MODE_KEY, mode);
    }
    catch {
        // ignore
    }
}
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
        profitShare: exports.DEFAULT_PROFIT_SHARE.map((e) => ({ ...e })),
        defaultProfitSharePct: exports.DEFAULT_PROFIT_SHARE_PCT,
        wageRates: exports.DEFAULT_WAGE_RATES.map((e) => ({ ...e })),
        defaultWageRate: exports.DEFAULT_WAGE_RATE,
        defaultCrewLeadRate: exports.DEFAULT_CREW_LEAD_RATE,
        conditionalRateRules: exports.DEFAULT_CONDITIONAL_RATE_RULES.map((r) => ({
            ...r,
            condition: { ...r.condition },
        })),
        perJobRateOverrides: exports.DEFAULT_PER_JOB_RATE_OVERRIDES.map((r) => ({ ...r })),
        nvDailyOtRateThreshold: exports.DEFAULT_NV_DAILY_OT_RATE_THRESHOLD,
        dailyOtHoursThreshold: exports.DEFAULT_DAILY_OT_HOURS_THRESHOLD,
        plAccountMapping: clonePlAccountMapping(exports.DEFAULT_PL_ACCOUNT_MAPPING),
        painterRoster: exports.DEFAULT_PAINTER_ROSTER.map((p) => ({ ...p })),
        salariedRoster: exports.DEFAULT_SALARIED.map((e) => ({ ...e })),
        mileageRate: exports.DEFAULT_MILEAGE_RATE,
        officeMileageRate: 0.5,
        subcontractors: [...exports.DEFAULT_SUBCONTRACTORS],
        prevailingWagePatterns: [...exports.DEFAULT_PREVAILING_PATTERNS],
        prevailingMinRate: exports.DEFAULT_PREVAILING_MIN_RATE,
        jobRateRules: [],
        payrollLineItems: exports.DEFAULT_PAYROLL_LINE_ITEMS.map((e) => ({ ...e })),
        fullMaterialsLeads: [...exports.DEFAULT_FULL_MATERIALS_LEADS],
        dailyOtExemptNames: [...exports.DEFAULT_DAILY_OT_EXEMPT],
        // Fresh settings already carry every Round 8 seed.
        appliedMigrations: [...exports.ALL_MIGRATIONS],
    };
}
function clonePlAccountMapping(src) {
    const one = (m) => ({
        income: [...m.income],
        materials: [...m.materials],
        laborWagesSubs: [...m.laborWagesSubs],
        sundries: [...m.sundries],
        travel: [...m.travel],
        mileage: [...m.mileage],
    });
    return { NV: one(src.NV), UT: one(src.UT) };
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
function parsePayroll(raw) {
    const out = {};
    if (typeof raw !== 'object' || raw === null)
        return out;
    const r = raw;
    if (Array.isArray(r.profitShare)) {
        const list = [];
        for (const e of r.profitShare) {
            const x = e;
            if (typeof x.lead === 'string' && typeof x.pct === 'number' && Number.isFinite(x.pct)) {
                list.push({ lead: x.lead, pct: x.pct });
            }
        }
        if (list.length > 0)
            out.profitShare = list;
    }
    if (typeof r.defaultProfitSharePct === 'number' && Number.isFinite(r.defaultProfitSharePct)) {
        out.defaultProfitSharePct = r.defaultProfitSharePct;
    }
    if (Array.isArray(r.wageRates)) {
        const list = [];
        for (const e of r.wageRates) {
            const x = e;
            if (typeof x.name === 'string' && typeof x.rate === 'number' && Number.isFinite(x.rate)) {
                list.push({
                    name: x.name,
                    rate: x.rate,
                    ...(typeof x.hourlyRate === 'number' && Number.isFinite(x.hourlyRate) ? { hourlyRate: x.hourlyRate } : {}),
                    ...(typeof x.crewLeadRate === 'number' && Number.isFinite(x.crewLeadRate)
                        ? { crewLeadRate: x.crewLeadRate }
                        : {}),
                    // Round 9 §18 — the two new per-painter pay-type rates.
                    ...(typeof x.commercialHourlyRate === 'number' && Number.isFinite(x.commercialHourlyRate)
                        ? { commercialHourlyRate: x.commercialHourlyRate }
                        : {}),
                    ...(typeof x.wages2Rate === 'number' && Number.isFinite(x.wages2Rate)
                        ? { wages2Rate: x.wages2Rate }
                        : {}),
                });
            }
        }
        if (list.length > 0)
            out.wageRates = list;
    }
    if (typeof r.defaultWageRate === 'number' && Number.isFinite(r.defaultWageRate)) {
        out.defaultWageRate = r.defaultWageRate;
    }
    if (typeof r.defaultCrewLeadRate === 'number' && Number.isFinite(r.defaultCrewLeadRate)) {
        out.defaultCrewLeadRate = r.defaultCrewLeadRate;
    }
    if (typeof r.nvDailyOtRateThreshold === 'number' && Number.isFinite(r.nvDailyOtRateThreshold)) {
        out.nvDailyOtRateThreshold = r.nvDailyOtRateThreshold;
    }
    if (typeof r.dailyOtHoursThreshold === 'number' && Number.isFinite(r.dailyOtHoursThreshold)) {
        out.dailyOtHoursThreshold = r.dailyOtHoursThreshold;
    }
    if (Array.isArray(r.conditionalRateRules)) {
        const rules = [];
        for (const e of r.conditionalRateRules) {
            if (typeof e !== 'object' || e === null)
                continue;
            const x = e;
            const cond = x.condition;
            if (typeof x.person !== 'string' || typeof x.rate !== 'number' || !Number.isFinite(x.rate))
                continue;
            if (!cond || cond.kind !== 'onSameJobAsLead' || typeof cond.leadName !== 'string')
                continue;
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
        const rules = [];
        for (const e of r.perJobRateOverrides) {
            if (typeof e !== 'object' || e === null)
                continue;
            const x = e;
            if (typeof x.person !== 'string' || typeof x.apKeyPattern !== 'string')
                continue;
            if (typeof x.rate !== 'number' || !Number.isFinite(x.rate))
                continue;
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
    if (mapping)
        out.plAccountMapping = mapping;
    if (Array.isArray(r.painterRoster)) {
        const list = [];
        for (const e of r.painterRoster) {
            if (typeof e !== 'object' || e === null)
                continue;
            const x = e;
            if (typeof x.name !== 'string' || x.name.trim() === '')
                continue;
            list.push({
                name: x.name,
                active: x.active === true,
                ...(typeof x.note === 'string' && x.note !== '' ? { note: x.note } : {}),
            });
        }
        out.painterRoster = list;
    }
    if (Array.isArray(r.salariedRoster)) {
        const list = [];
        for (const e of r.salariedRoster) {
            const x = e;
            if (typeof x.name === 'string' &&
                (x.section === 'ADMIN' || x.section === 'ESTIMATORS' || x.section === 'MANAGER SALARY') &&
                typeof x.amount === 'number' &&
                Number.isFinite(x.amount)) {
                list.push({
                    name: x.name,
                    section: x.section,
                    role: typeof x.role === 'string' ? x.role : '',
                    amount: x.amount,
                    // Round 8 §7.3 — pre-Round-8 rows have no entity; they are NV.
                    entity: x.entity === 'UT' ? 'UT' : 'NV',
                    // Round 9 §25/§26/§28/§34.
                    ...(x.frequency === 'firstPayDateOfMonth' ? { frequency: 'firstPayDateOfMonth' } : {}),
                    ...(x.sub === true ? { sub: true } : {}),
                    ...(typeof x.reimbursementsLabel === 'string' && x.reimbursementsLabel.trim() !== ''
                        ? { reimbursementsLabel: x.reimbursementsLabel }
                        : {}),
                    ...(typeof x.taxes === 'number' && Number.isFinite(x.taxes) ? { taxes: x.taxes } : {}),
                    ...(typeof x.net === 'number' && Number.isFinite(x.net) ? { net: x.net } : {}),
                });
            }
        }
        if (list.length > 0)
            out.salariedRoster = list;
    }
    if (typeof r.mileageRate === 'number' && Number.isFinite(r.mileageRate))
        out.mileageRate = r.mileageRate;
    if (Array.isArray(r.subcontractors)) {
        const subs = r.subcontractors.filter((s) => typeof s === 'string' && s.trim() !== '');
        if (subs.length > 0)
            out.subcontractors = subs;
    }
    if (Array.isArray(r.prevailingWagePatterns)) {
        const pats = r.prevailingWagePatterns.filter((s) => typeof s === 'string' && s.trim() !== '');
        if (pats.length > 0)
            out.prevailingWagePatterns = pats;
    }
    if (typeof r.prevailingMinRate === 'number' && Number.isFinite(r.prevailingMinRate)) {
        out.prevailingMinRate = r.prevailingMinRate;
    }
    if (typeof r.officeMileageRate === 'number' && Number.isFinite(r.officeMileageRate)) {
        out.officeMileageRate = r.officeMileageRate;
    }
    if (Array.isArray(r.jobRateRules)) {
        const rules = [];
        for (const raw of r.jobRateRules) {
            if (typeof raw !== 'object' || raw === null)
                continue;
            const x = raw;
            if (typeof x.apKey !== 'string' || typeof x.rate !== 'number')
                continue;
            rules.push({
                id: typeof x.id === 'string' ? x.id : `rule-${rules.length}`,
                apKey: x.apKey.trim(),
                rate: x.rate,
                employeeNames: Array.isArray(x.employeeNames)
                    ? x.employeeNames.filter((n) => typeof n === 'string')
                    : [],
                includeSuffixes: x.includeSuffixes !== false,
                note: typeof x.note === 'string' ? x.note : '',
            });
        }
        out.jobRateRules = rules;
    }
    if (Array.isArray(r.payrollLineItems)) {
        const list = [];
        for (const e of r.payrollLineItems) {
            if (typeof e !== 'object' || e === null)
                continue;
            const x = e;
            if (typeof x.person !== 'string' || typeof x.label !== 'string')
                continue;
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
        out.fullMaterialsLeads = r.fullMaterialsLeads.filter((n) => typeof n === 'string' && n.trim() !== '');
    }
    if (Array.isArray(r.dailyOtExemptNames)) {
        out.dailyOtExemptNames = r.dailyOtExemptNames.filter((n) => typeof n === 'string' && n.trim() !== '');
    }
    if (Array.isArray(r.appliedMigrations)) {
        out.appliedMigrations = r.appliedMigrations.filter((s) => typeof s === 'string');
    }
    else {
        // Pre-Round-8 blob: nothing has been migrated yet.
        out.appliedMigrations = [];
    }
    return out;
}
/** Validate a persisted `plAccountMapping`; unknown metrics/entities are dropped. */
function parsePlAccountMapping(raw) {
    if (typeof raw !== 'object' || raw === null)
        return null;
    const src = raw;
    const metrics = ['income', 'materials', 'laborWagesSubs', 'sundries', 'travel', 'mileage'];
    const one = (entity) => {
        const e = src[entity];
        if (typeof e !== 'object' || e === null)
            return null;
        const rec = e;
        const fallback = exports.DEFAULT_PL_ACCOUNT_MAPPING[entity];
        const out = {};
        for (const m of metrics) {
            const v = rec[m];
            out[m] = Array.isArray(v)
                ? v.filter((s) => typeof s === 'string' && s.trim() !== '')
                : [...fallback[m]];
        }
        return out;
    };
    const nv = one('NV');
    const ut = one('UT');
    if (!nv && !ut)
        return null;
    return {
        NV: nv ?? clonePlAccountMapping(exports.DEFAULT_PL_ACCOUNT_MAPPING).NV,
        UT: ut ?? clonePlAccountMapping(exports.DEFAULT_PL_ACCOUNT_MAPPING).UT,
    };
}
/**
 * Validate any persisted settings payload (localStorage or the shared store)
 * and fill every missing section from the seeds. Returns null when the payload
 * is not a settings object at all. Round 8 migrations are applied here, so both
 * storage paths heal identically.
 */
function parseSettings(raw) {
    const legend = parseLegend(raw);
    if (!legend)
        return null;
    const pivotColumnOrder = ensureDefaultPivotColumns(parsePivotOrder(raw) ?? [...exports.DEFAULT_PIVOT_COLUMNS]);
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
function applySettingsMigrations(settings) {
    const done = new Set(settings.appliedMigrations);
    let next = settings;
    const patch = (p) => {
        next = { ...next, ...p };
    };
    if (!done.has(exports.MIGRATION_RF_COLOR)) {
        // Only the exact old seed moves; a custom RF pick is left alone.
        patch({
            legend: next.legend.map((l) => l.estimator === 'RF' && l.color.toUpperCase() === exports.LEGACY_RF_COLOR ? { ...l, color: 'C4A484' } : l),
        });
        done.add(exports.MIGRATION_RF_COLOR);
    }
    if (!done.has(exports.MIGRATION_DB_PATTERN)) {
        const broadened = [];
        for (const p of next.prevailingWagePatterns) {
            const v = p.trim() === exports.LEGACY_PREVAILING_PATTERN ? 'SNRHA' : p;
            if (!broadened.includes(v))
                broadened.push(v);
        }
        patch({ prevailingWagePatterns: broadened });
        done.add(exports.MIGRATION_DB_PATTERN);
    }
    if (!done.has(exports.MIGRATION_WAGE_SEEDS)) {
        patch({
            wageRates: mergeWageRateSeeds(next.wageRates, exports.ROUND8_WAGE_SEEDS, 'fillMissing'),
            conditionalRateRules: next.conditionalRateRules.length > 0
                ? next.conditionalRateRules
                : exports.DEFAULT_CONDITIONAL_RATE_RULES.map((r) => ({ ...r, condition: { ...r.condition } })),
            perJobRateOverrides: next.perJobRateOverrides.length > 0
                ? next.perJobRateOverrides
                : exports.DEFAULT_PER_JOB_RATE_OVERRIDES.map((r) => ({ ...r })),
        });
        done.add(exports.MIGRATION_WAGE_SEEDS);
    }
    // Round 9 §12/§17/§18/§21/§22/§35 — corrected rates WIN over stored values,
    // because they are corrections to numbers the app had wrong.
    if (!done.has(exports.MIGRATION_R9_WAGES)) {
        patch({ wageRates: mergeWageRateSeeds(next.wageRates, exports.ROUND9_WAGE_SEEDS, 'seedWins') });
        done.add(exports.MIGRATION_R9_WAGES);
    }
    // Round 9 §19/§23/§24/§29 — add the seeded line items that are missing by id.
    if (!done.has(exports.MIGRATION_R9_LINE_ITEMS)) {
        const have = new Set((next.payrollLineItems ?? []).map((i) => i.id));
        patch({
            payrollLineItems: [
                ...(next.payrollLineItems ?? []),
                ...exports.DEFAULT_PAYROLL_LINE_ITEMS.filter((i) => !have.has(i.id)).map((i) => ({ ...i })),
            ],
            fullMaterialsLeads: next.fullMaterialsLeads && next.fullMaterialsLeads.length > 0
                ? next.fullMaterialsLeads
                : [...exports.DEFAULT_FULL_MATERIALS_LEADS],
            dailyOtExemptNames: next.dailyOtExemptNames && next.dailyOtExemptNames.length > 0
                ? next.dailyOtExemptNames
                : [...exports.DEFAULT_DAILY_OT_EXEMPT],
        });
        done.add(exports.MIGRATION_R9_LINE_ITEMS);
    }
    // Round 9 §25/§26/§28/§34 — salaried additions + flags, matched by name.
    if (!done.has(exports.MIGRATION_R9_SALARIED)) {
        const key = (n) => n.trim().toLowerCase().replace(/\s+/g, ' ');
        const bySeed = new Map(exports.DEFAULT_SALARIED.map((e) => [key(e.name), e]));
        const patched = next.salariedRoster.map((row) => {
            const seed = bySeed.get(key(row.name));
            if (!seed)
                return row;
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
                ...exports.DEFAULT_SALARIED.filter((e) => !haveNames.has(key(e.name))).map((e) => ({ ...e })),
            ],
        });
        done.add(exports.MIGRATION_R9_SALARIED);
    }
    // Round 9 §13/§18 — retire the two person-specific rate seeds (the rules
    // themselves stay available; only these seeded rows go).
    if (!done.has(exports.MIGRATION_R9_RATE_RULES)) {
        patch({
            conditionalRateRules: next.conditionalRateRules.filter((r) => r.id !== 'cond-henry-limay-flores-miguel-martin'),
            perJobRateOverrides: next.perJobRateOverrides.filter((r) => r.id !== 'perjob-samuel-brewster-24097'),
        });
        done.add(exports.MIGRATION_R9_RATE_RULES);
    }
    if (!done.has(exports.MIGRATION_CARLOS_HOURLY)) {
        patch({
            wageRates: next.wageRates.map((row) => /^carlos\s+ruiz(\s+cruz)?$/i.test(row.name.trim()) && (row.hourlyRate === undefined || row.hourlyRate === 18)
                ? { ...row, hourlyRate: 25 }
                : row),
        });
        done.add(exports.MIGRATION_CARLOS_HOURLY);
    }
    return { ...next, appliedMigrations: [...done] };
}
function loadSettings() {
    try {
        // v4 first; fall back to v3 (legend/pivot only) so users keep palette edits.
        const raw = localStorage.getItem(STORAGE_KEY) ?? localStorage.getItem(STORAGE_KEY_V3);
        if (raw) {
            const parsed = parseSettings(JSON.parse(raw));
            if (parsed)
                return parsed;
        }
    }
    catch {
        // fall through
    }
    return defaultSettings();
}
/** True when this browser has its own saved settings (drives the §1 push-up). */
function hasLocalSettings() {
    try {
        return (localStorage.getItem(STORAGE_KEY) !== null || localStorage.getItem(STORAGE_KEY_V3) !== null);
    }
    catch {
        return false;
    }
}
function saveSettings(settings) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
    }
    catch {
        // localStorage unavailable — settings just won't persist.
    }
}
async function probeSettingsStore() {
    try {
        const res = await fetch('/api/settings', { cache: 'no-store' });
        if (res.status === 503)
            return 'unconfigured';
        if (res.status === 404)
            return 'empty';
        if (!res.ok)
            return 'error';
        // Dev/static hosts answer unknown paths with index.html — that is not a store.
        try {
            await res.json();
        }
        catch {
            return 'unconfigured';
        }
        return 'ready';
    }
    catch {
        return 'error';
    }
}
/** Accept either the `{ settings, savedAt }` envelope or a bare AppSettings blob. */
function parseStoredSettings(raw) {
    if (typeof raw === 'object' && raw !== null && 'settings' in raw) {
        return parseSettings(raw.settings);
    }
    return parseSettings(raw);
}
/**
 * Load the shared Settings (Vercel Blob via /api). This is what lets a cleared
 * browser or incognito window see the same saved Settings.
 */
async function fetchAppSettings() {
    try {
        const res = await fetch('/api/settings', { cache: 'no-store' });
        if (res.ok)
            return parseStoredSettings((await res.json()));
    }
    catch {
        // fall through
    }
    // Local vite / static fallback (not used for cross-device on Vercel).
    try {
        const res = await fetch('/app-settings.json', { cache: 'no-store' });
        if (res.ok)
            return parseStoredSettings((await res.json()));
    }
    catch {
        // ignore
    }
    return null;
}
/** Write the current Settings to the shared store (last write wins). */
async function persistAppSettings(settings) {
    const payload = { settings, savedAt: new Date().toISOString() };
    try {
        const res = await fetch('/api/settings', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        });
        if (res.ok)
            return { ok: true };
        let message = `Could not sync Settings (HTTP ${res.status}).`;
        try {
            const body = (await res.json());
            if (body.message)
                message = body.message;
        }
        catch {
            // ignore
        }
        return { ok: false, status: res.status, message };
    }
    catch {
        return {
            ok: false,
            status: 0,
            message: 'Shared sync unavailable — your Settings are still saved in this browser. Cross-device sync is optional.',
        };
    }
}
async function clearAppSettings() {
    try {
        await fetch('/api/settings', { method: 'DELETE' });
    }
    catch {
        // ignore
    }
}
// ---------------------------------------------------------------------------
// Round 4 §4.2 — persisted employee directory (same pattern as the legend)
// ---------------------------------------------------------------------------
const DIRECTORY_KEY = 'pct-completion-builder.employee-directory.v1';
function loadSavedDirectory() {
    try {
        const raw = localStorage.getItem(DIRECTORY_KEY);
        if (!raw)
            return null;
        return parseSavedDirectory(JSON.parse(raw));
    }
    catch {
        return null;
    }
}
/** Write an already-built SavedDirectory into this browser's localStorage (keeps uploadedAt). */
function cacheDirectoryLocally(saved) {
    try {
        localStorage.setItem(DIRECTORY_KEY, JSON.stringify(saved));
    }
    catch {
        // localStorage unavailable
    }
}
function saveDirectory(employees, sourceFilename) {
    const saved = {
        employees,
        sourceFilename,
        uploadedAt: new Date().toISOString(),
        activeCount: employees.filter((e) => e.active).length,
        inactiveCount: employees.filter((e) => !e.active).length,
    };
    cacheDirectoryLocally(saved);
    return saved;
}
function clearSavedDirectory() {
    try {
        localStorage.removeItem(DIRECTORY_KEY);
    }
    catch {
        // ignore
    }
}
/** Parse/validate a SavedDirectory payload (localStorage or app file). */
function parseSavedDirectory(raw) {
    if (typeof raw !== 'object' || raw === null)
        return null;
    const p = raw;
    if (!Array.isArray(p.employees))
        return null;
    const employees = [];
    for (const e of p.employees) {
        const x = e;
        if (typeof x.fullName !== 'string' || x.fullName.trim() === '')
            continue;
        employees.push({
            fullName: x.fullName,
            firstName: typeof x.firstName === 'string' ? x.firstName : '',
            lastName: typeof x.lastName === 'string' ? x.lastName : '',
            active: x.active !== false,
        });
    }
    if (employees.length === 0)
        return null;
    return {
        employees,
        sourceFilename: typeof p.sourceFilename === 'string' ? p.sourceFilename : '',
        uploadedAt: typeof p.uploadedAt === 'string' ? p.uploadedAt : '',
        activeCount: employees.filter((e) => e.active).length,
        inactiveCount: employees.filter((e) => !e.active).length,
    };
}
/** Prefer the newer of two saved directories by `uploadedAt`. */
function newerDirectory(a, b) {
    if (!a)
        return b;
    if (!b)
        return a;
    return (a.uploadedAt || '') >= (b.uploadedAt || '') ? a : b;
}
async function probeDirectoryStore() {
    try {
        const res = await fetch('/api/employee-directory', { cache: 'no-store' });
        if (res.status === 503)
            return 'unconfigured';
        if (res.status === 404)
            return 'empty';
        if (res.ok)
            return 'ready';
        return 'error';
    }
    catch {
        return 'error';
    }
}
/**
 * Load the shared employee list (Vercel Blob via /api, or local public file).
 * This is what makes phone + desktop see the same list on the Vercel URL.
 */
async function fetchAppDirectory() {
    try {
        const res = await fetch('/api/employee-directory', { cache: 'no-store' });
        if (res.ok)
            return parseSavedDirectory(await res.json());
    }
    catch {
        // fall through
    }
    // Local vite / static fallback (not used for cross-device on Vercel).
    try {
        const res = await fetch('/employee-directory.json', { cache: 'no-store' });
        if (res.ok)
            return parseSavedDirectory(await res.json());
    }
    catch {
        // ignore
    }
    return null;
}
/** Write the current list to the shared store (Vercel Blob in production). */
async function persistAppDirectory(saved) {
    try {
        const res = await fetch('/api/employee-directory', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(saved),
        });
        if (res.ok)
            return { ok: true };
        let message = `Could not sync employee list (HTTP ${res.status}).`;
        try {
            const body = (await res.json());
            if (body.message)
                message = body.message;
        }
        catch {
            // ignore
        }
        return { ok: false, status: res.status, message };
    }
    catch {
        return {
            ok: false,
            status: 0,
            message: 'Shared sync unavailable — your list is still saved in this browser. Cross-device sync is optional.',
        };
    }
}
async function clearAppDirectory() {
    try {
        await fetch('/api/employee-directory', { method: 'DELETE' });
    }
    catch {
        // ignore
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
