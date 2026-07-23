"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.runSelfTests = runSelfTests;
/**
 * Dev-mode self tests for the pure logic pieces. Runs automatically when the
 * dev server starts (see main.tsx) and logs pass/fail to the console.
 */
const apkey_1 = require("./apkey");
const format_1 = require("./format");
const completedPercents_1 = require("./completedPercents");
const parsers_1 = require("./parsers");
const payroll_1 = require("./payroll");
const payrollPreview_1 = require("./payrollPreview");
const prFinalCheck_1 = require("./prFinalCheck");
const pivotByJob_1 = require("./pivotByJob");
const pivot_1 = require("./pivot");
const pivotView_1 = require("./pivotView");
const jobNames_1 = require("./jobNames");
const settings_1 = require("./settings");
const timenavi_1 = require("./timenavi");
function runSelfTests() {
    const results = [];
    const check = (name, actual, expected) => {
        const a = JSON.stringify(actual);
        const e = JSON.stringify(expected);
        results.push({ name, pass: a === e, detail: a === e ? undefined : `expected ${e}, got ${a}` });
    };
    // --- apKey extraction -----------------------------------------------------
    check('apKey: plain', (0, apkey_1.extractApKey)('AP24314 Ext'), '24314');
    check('apKey: with address', (0, apkey_1.extractApKey)('AP24403 5321 Lambrook Dr Int/Ext'), '24403');
    check('apKey: suffixed', (0, apkey_1.extractApKey)('AP21848-A5 Somewhere'), '21848-A5');
    check('apKey: from event title', (0, apkey_1.extractApKey)('♦Int 13k (XC) Regina Conrad Ph1, AP24180-A *PAID* (Daniel)'), '24180-A');
    check('apKey: spaced/dashed', (0, apkey_1.extractApKey)('AP - 24222 job'), '24222');
    check('apKey: none', (0, apkey_1.extractApKey)('Just a random title'), null);
    check('apKey: OFFICE rule', (0, apkey_1.timesheetApKey)('OFFICE [Please put AP# in notes]'), 'OFFICE');
    check('apKey: office lowercase', (0, apkey_1.timesheetApKey)('  office misc'), 'OFFICE');
    check('apKey: normalize bare number', (0, apkey_1.normalizeApInput)('24314'), '24314');
    check('apKey: normalize AP prefix', (0, apkey_1.normalizeApInput)('ap24314 ext'), '24314');
    check('apKey: normalize suffixed lowercase', (0, apkey_1.normalizeApInput)('21848-a5'), '21848-A5');
    check('apKey: normalize empty', (0, apkey_1.normalizeApInput)('  '), null);
    check('apKey: string compare survives suffix', (0, apkey_1.extractApKey)('AP21848-A5') === '21848-A5', true);
    // --- % suggestion parser (pay-date driven) ----------------------------------
    const day = (m, d, y) => new Date(y, m - 1, d);
    const multiWeek = '7/17/26 100% Complete\n7/10/26 30% COMPLETE -SH\n7/03/26 HOURLY';
    check('%: pay date selects its own line (7/10 → 30%)', (0, timenavi_1.suggestPercent)(multiWeek, day(7, 10, 2026)).value, 0.3);
    check('%: same file, pay 7/17 → 100%', (0, timenavi_1.suggestPercent)(multiWeek, day(7, 17, 2026)).value, 1);
    check('%: no exact match → latest ≤ pay date', (0, timenavi_1.suggestPercent)(multiWeek, day(7, 8, 2026)).value, 'HOURLY');
    check('%: all lines after pay date → blank', (0, timenavi_1.suggestPercent)(multiWeek, day(6, 1, 2026)).value, null);
    check('%: exact-match line beats a newer line above it', (0, timenavi_1.suggestPercent)('7/10/26 95% COMPLETE -SH\n7/03/26 90% Complete', day(7, 10, 2026)).value, 0.95);
    check('%: hourly on the selected line', (0, timenavi_1.suggestPercent)('7/10/26 HOURLY -SH', day(7, 10, 2026)).value, 'HOURLY');
    check('%: ♦/whitespace prefix tolerated', (0, timenavi_1.suggestPercent)('  ♦7/10/26 97.5% complete', day(7, 10, 2026)).value, 0.975);
    check('%: 4-digit year status line', (0, timenavi_1.suggestPercent)('7/10/2026 40% Complete', day(7, 10, 2026)).value, 0.4);
    check('%: selected line with no %/HOURLY (e.g. "??") → blank, line kept for tooltip', (0, timenavi_1.suggestPercent)('7/10/26 ??', day(7, 10, 2026)), { value: null, line: '7/10/26 ??' });
    check('%: no status lines at all', (0, timenavi_1.suggestPercent)('Gate code 1234\ncall before arriving', day(7, 10, 2026)).value, null);
    check('%: blank description', (0, timenavi_1.suggestPercent)('', day(7, 10, 2026)).value, null);
    check('%: no pay date → no suggestion', (0, timenavi_1.suggestPercent)(multiWeek, null).value, null);
    check('%: source line kept for tooltip', (0, timenavi_1.suggestPercent)('7/10/26 95% COMPLETE -SH', day(7, 10, 2026)).line, '7/10/26 95% COMPLETE -SH');
    // --- QuickBooks formula-remnant numbers ------------------------------------
    check('num: formula remnant', (0, parsers_1.parseNumberLoose)('=1000'), 1000);
    check('num: subtotal formula', (0, parsers_1.parseNumberLoose)('=(B2)+(B3)'), 'FORMULA');
    check('num: =650', (0, parsers_1.parseNumberLoose)('=650'), 650);
    check('num: currency string', (0, parsers_1.parseNumberLoose)('$1,234.50'), 1234.5);
    check('num: plain', (0, parsers_1.parseNumberLoose)(650), 650);
    check('num: junk', (0, parsers_1.parseNumberLoose)('n/a'), null);
    // --- date parsing -----------------------------------------------------------
    const d1 = (0, parsers_1.parseDate)('6/29/2026');
    check('date: M/D/YYYY', d1 ? [d1.getFullYear(), d1.getMonth() + 1, d1.getDate()] : null, [2026, 6, 29]);
    const d2 = (0, parsers_1.parseDate)('01/01/2024');
    check('date: MM/DD/YYYY', d2 ? [d2.getFullYear(), d2.getMonth() + 1, d2.getDate()] : null, [2024, 1, 1]);
    const d3 = (0, parsers_1.parseDate)('2026-06-23');
    check('date: ISO', d3 ? [d3.getFullYear(), d3.getMonth() + 1, d3.getDate()] : null, [2026, 6, 23]);
    const d4 = (0, parsers_1.parseDate)('6/29/2026 7:00');
    check('date: time component stripped (timezone-safe)', d4 ? [d4.getMonth() + 1, d4.getDate(), d4.getHours(), d4.getMinutes()] : null, [6, 29, 0, 0]);
    const d5 = (0, parsers_1.parseDate)(46213); // Excel serial for 2026-07-10 (1900 system)
    check('date: excel serial', d5 ? [d5.getFullYear(), d5.getMonth() + 1, d5.getDate()] : null, [2026, 7, 10]);
    const d6 = (0, parsers_1.parseDate)(new Date(2026, 5, 23, 7, 0, 0));
    check('date: Date cell stripped to local midnight', d6 ? [d6.getMonth() + 1, d6.getDate(), d6.getHours()] : null, [6, 23, 0]);
    check('date: garbage', (0, parsers_1.parseDate)('not a date'), null);
    // --- date difference ----------------------------------------------------------
    const inv = (0, parsers_1.parseDate)('6/25/2026');
    const start = (0, parsers_1.parseDate)('7/1/2026');
    check('dateDiff: negative ok', (0, format_1.dateDiffDays)(inv, start), -6);
    check('dateDiff: same day', (0, format_1.dateDiffDays)(start, start), 0);
    // --- percent input --------------------------------------------------------------
    check('pct input: 95', (0, format_1.parsePercentInput)('95'), 0.95);
    check('pct input: 95%', (0, format_1.parsePercentInput)('95%'), 0.95);
    check('pct input: 0.95', (0, format_1.parsePercentInput)('0.95'), 0.95);
    check('pct input: hourly', (0, format_1.parsePercentInput)('Hourly'), 'HOURLY');
    check('pct input: dash', (0, format_1.parsePercentInput)('-'), null);
    // --- title initials ----------------------------------------------------------------
    check('initials: (XC)', (0, timenavi_1.extractTitleInitials)('♦Int 13k (XC) Regina Conrad, AP24180-A (Daniel)'), 'XC');
    check('initials: lowercase name ignored', (0, timenavi_1.extractTitleInitials)('Job for (Daniel)'), null);
    // --- estimator fallback from the event title ----------------------------------------
    const known = new Set(['JS', 'MH', 'PG', 'TE', 'XC', 'ND', 'RF']);
    check('title est: first known parenthesized token', (0, timenavi_1.estimatorFromTitle)('♦Int 13k (XC) Regina Conrad Ph1, AP24180-A *PAID* (Daniel)', known), 'XC');
    check('title est: crew names with / or >3 letters never match', (0, timenavi_1.estimatorFromTitle)('♦[Southpoint] 3.5k (MH) Framing/DW/FRP Install, AP24450 (Ted/Rafael)', known), 'MH');
    check('title est: lowercase tag uppercased', (0, timenavi_1.estimatorFromTitle)('Job walk (js) tomorrow', known), 'JS');
    check('title est: known-set guard rejects (Ted)', (0, timenavi_1.estimatorFromTitle)('Fence repair (Ted)', known), null);
    check('title est: skips unknown token to find a later known one', (0, timenavi_1.estimatorFromTitle)('(Ted) plus (PG)', known), 'PG');
    check('title est: no parenthesized tokens', (0, timenavi_1.estimatorFromTitle)('Paint shop maintenance', known), null);
    // --- crew-lead-first ordering ----------------------------------------------------------
    const names = (emps) => emps.map((e) => e.fname);
    const carlosCrew = [
        { fname: 'Ana', lname: 'Bedoya Melgar' },
        { fname: 'Carlos', lname: 'Ruiz Cruz' },
    ];
    check('lead: Carlos first in his crew', names((0, pivot_1.orderLeadFirst)("Carlos Ruiz Cruz's Crew", carlosCrew)), ['Carlos', 'Ana']);
    check('lead: extra surname matches (Daniel Bonilla → Daniel Gonzales Bonilla)', names((0, pivot_1.orderLeadFirst)("Daniel Bonilla's Crew", [
        { fname: 'Armando', lname: 'Lopez' },
        { fname: 'Daniel', lname: 'Gonzales Bonilla' },
    ])), ['Daniel', 'Armando']);
    check('lead: token guard (Miguel Martin ≠ Samuel Martin)', names((0, pivot_1.orderLeadFirst)('Miguel Martin\u2019s Crew', [
        { fname: 'Miguel', lname: 'Martin' },
        { fname: 'Samuel', lname: 'Martin' },
    ])), ['Miguel', 'Samuel']);
    check('lead: hyphenated surname tokenized (Leonardo Martinez → Martinez-Rincon)', names((0, pivot_1.orderLeadFirst)("Leonardo Martinez's Crew", [
        { fname: 'Juan', lname: 'Perez' },
        { fname: 'Leonardo', lname: 'Martinez-Rincon' },
    ])), ['Leonardo', 'Juan']);
    check('lead: tie → fewest extra tokens wins', (0, pivot_1.orderLeadFirst)("Tomas Huerta's Crew", [
        { fname: 'Tomas', lname: 'Huerta Bonilla Garcia' },
        { fname: 'Tomas', lname: 'Huerta Bonilla' },
    ])[0].lname, 'Huerta Bonilla');
    check('lead: no "\'s Crew" suffix → untouched', names((0, pivot_1.orderLeadFirst)('Handyman', carlosCrew)), ['Ana', 'Carlos']);
    check('lead: no matching employee → untouched', names((0, pivot_1.orderLeadFirst)("Rene Salazar's Crew", carlosCrew)), ['Ana', 'Carlos']);
    // --- palette text contrast -----------------------------------------------------------
    check('contrast: ND dark blue → white text', (0, format_1.isDarkColor)('4472C4'), true);
    check('contrast: MH light yellow → black text', (0, format_1.isDarkColor)('FFFF99'), false);
    check('contrast: RF dusty rose → black text', (0, format_1.isDarkColor)('DA9694'), false);
    // --- Round 3 §3.1: job display name -----------------------------------------------------
    check('jobName: AP token stripped', (0, payroll_1.jobDisplayName)('Aaron Ward', 'AP24314 Ext'), 'Aaron Ward Ext');
    check('jobName: suffixed AP + inner text', (0, payroll_1.jobDisplayName)('Tri-State', 'AP24097-A Sams Club Int'), 'Tri-State Sams Club Int');
    check('jobName: single-spaced', (0, payroll_1.jobDisplayName)(' CCSD ', 'AP24403  5321 Lambrook Dr Int/Ext'), 'CCSD 5321 Lambrook Dr Int/Ext');
    // --- Round 3 §3.2: crew-lead extraction ---------------------------------------------------
    check('lead extract: last parenthetical, not the estimator', (0, payroll_1.extractLeadParenthetical)('♦Int 13k (XC) Regina Conrad Ph1, AP24180-A *PAID* (Daniel)', known), 'Daniel');
    check('lead extract: estimator-only title → null', (0, payroll_1.extractLeadParenthetical)('AP24403 Lambrook (PG)', known), null);
    check('lead extract: no parentheses → null', (0, payroll_1.extractLeadParenthetical)('AP24477 Plain title', known), null);
    check('lead extract: business name kept', (0, payroll_1.extractLeadParenthetical)("Ext 8k (MH) Smith, AP24499 (Pepe's Painting)", known), "Pepe's Painting");
    check('lead split: slash', (0, payroll_1.splitLeadCandidates)('Ted/Rafael'), ['Ted', 'Rafael']);
    check('lead split: commas', (0, payroll_1.splitLeadCandidates)('Ted, Antonio, Tomas'), ['Ted', 'Antonio', 'Tomas']);
    check('lead split: hrly_ prefix stripped', (0, payroll_1.splitLeadCandidates)('hrly_Damian/Juan Salazar'), ['Damian', 'Juan Salazar']);
    check('lead split: asterisks stripped', (0, payroll_1.splitLeadCandidates)('*Carlos*'), ['Carlos']);
    const dir = [
        { fullName: 'Carlos Ruiz Cruz', firstName: 'Carlos', lastName: 'Ruiz Cruz', active: true },
        { fullName: 'Miguel Martin', firstName: 'Miguel', lastName: 'Martin', active: true },
        { fullName: 'Miguel Borquez Gil', firstName: 'Miguel', lastName: 'Borquez Gil', active: true },
        { fullName: 'Calixto Efrain Climaco', firstName: 'Calixto Efrain', lastName: 'Climaco', active: false },
        { fullName: 'Ted Frank', firstName: 'Ted', lastName: 'Frank', active: true },
        { fullName: 'Rafael Aguilar', firstName: 'Rafael', lastName: 'Aguilar', active: true },
    ];
    const ctx = {
        directory: dir,
        crewGroups: ["Carlos Ruiz Cruz's Crew", "Miguel Martin's Crew", "Ted Frank's Crew", "Rafael Aguilar's Crew"],
        subcontractors: ["Pepe's Painting LLC", 'Color Enterprises LLC'],
    };
    check('lead resolve: Carlos → CARLOS RUIZ CRUZ', (0, payroll_1.resolveCrewLead)('Carlos', ctx).display, 'CARLOS RUIZ CRUZ');
    check('lead resolve: crew-lead Miguel Martin preferred over other Miguels', (0, payroll_1.resolveCrewLead)('Miguel', ctx).display, 'MIGUEL MARTIN');
    check('lead resolve: sub via token match', [(0, payroll_1.resolveCrewLead)("Pepe's Painting", ctx).display, (0, payroll_1.resolveCrewLead)("Pepe's Painting", ctx).sub], ["PEPE'S PAINTING LLC", true]);
    const welder = (0, payroll_1.resolveCrewLead)('Welder', ctx);
    check('lead resolve: no match anywhere → flag', [welder.display, welder.flags[0]?.kind], ['', 'noMatch']);
    const split = (0, payroll_1.resolveCrewLead)('Ted/Rafael', ctx);
    check('lead resolve: split leads joined + split:2', [split.display, split.split], ['TED FRANK/RAFAEL AGUILAR', 2]);
    const calixto = (0, payroll_1.resolveCrewLead)('Calixto', ctx);
    check('lead resolve: inactive-only match used but flagged', [calixto.display, calixto.flags.some((f) => f.kind === 'inactiveMatch')], ['CALIXTO EFRAIN CLIMACO', true]);
    check('lead resolve: empty raw → empty, no flags', (0, payroll_1.resolveCrewLead)(null, ctx), { display: '', names: [], split: 1, sub: false, raw: '', flags: [] });
    check('tokens: middle initial optional', (0, payroll_1.tokensSubset)('Antonio Zuniga', 'Antonio J Zuniga'), true);
    // --- Round 3 §3.4: overtime split ---------------------------------------------------------
    const day2 = (m, d) => new Date(2026, m - 1, d);
    const ot = (0, payroll_1.splitOvertime)([
        { jobId: 'a', hours: 25, lastDate: day2(6, 30), lastRowIndex: 1 },
        { jobId: 'b', hours: 15.67, lastDate: day2(7, 3), lastRowIndex: 5 },
    ]);
    check('OT: overage on the LAST job', [ot.otJobId, ot.otHours], ['b', 0.67]);
    check('OT: regular reduced so reg+OT = actual', ot.regularByJob.get('b'), 15);
    check('OT: other jobs untouched', ot.regularByJob.get('a'), 25);
    const noOt = (0, payroll_1.splitOvertime)([{ jobId: 'a', hours: 40, lastDate: day2(7, 3), lastRowIndex: 0 }]);
    check('OT: exactly 40 → none', [noOt.otJobId, noOt.otHours], [null, 0]);
    const tie = (0, payroll_1.splitOvertime)([
        { jobId: 'x', hours: 30, lastDate: day2(7, 3), lastRowIndex: 2 },
        { jobId: 'y', hours: 12, lastDate: day2(7, 3), lastRowIndex: 7 },
    ]);
    check('OT: same-day tie → last by file order', tie.otJobId, 'y');
    // --- Round 3 §4.3: WO pre-fills -------------------------------------------------------------
    const jodi = (0, payroll_1.extractWoOverride)('gate 1234\n****PER JODI $4,715.00 is the total of work order he completed****');
    check('WO$: $ before "work order" caught', [jodi?.amount, jodi?.line?.includes('PER JODI')], [4715, true]);
    check('WO$: W.O. before $', (0, payroll_1.extractWoOverride)('W.O. total came to $2,300')?.amount, 2300);
    check('WO$: plain money line ignored', (0, payroll_1.extractWoOverride)('Invoice total $9,000 sent'), null);
    check('WO%: "@ 40% PS" line', (0, payroll_1.extractWoPct)('7/03/26 10% COMPLETE @ 40% PS')?.pct, 0.4);
    check('WO%: absent → null', (0, payroll_1.extractWoPct)('7/03/26 10% COMPLETE'), null);
    // --- Round 3 §7: settings lookups -----------------------------------------------------------
    const psMap = [
        { lead: 'Antonio Zuniga', pct: 0.36 },
        { lead: "Pepe's Painting LLC", pct: 0.5 },
    ];
    check('PS%: token match (ALL-CAPS lead)', (0, payroll_1.profitSharePctFor)('ANTONIO ZUNIGA', psMap, 0.38), 0.36);
    check('PS%: unlisted crew → default', (0, payroll_1.profitSharePctFor)('NEW GUY', psMap, 0.38), 0.38);
    check("PS%: sub matched (PEPE'S PAINTING LLC)", (0, payroll_1.profitSharePctFor)("PEPE'S PAINTING LLC", psMap, 0.38), 0.5);
    const rates = [{ name: 'Rafael Aguilar', rate: 18, hourlyRate: 25 }];
    check('wage: piece-rate job → default rate', (0, payroll_1.wageRateFor)('RAFAEL AGUILAR', rates, 18, false), { rate: 18, defaulted: false });
    check('wage: hourly job → hourly rate', (0, payroll_1.wageRateFor)('RAFAEL AGUILAR', rates, 18, true), { rate: 25, defaulted: false });
    check('wage: unknown employee → fallback + flag', (0, payroll_1.wageRateFor)('Nobody Known', rates, 18, false), { rate: 18, defaulted: true });
    // --- Round 4 §1.1: pink review flags ---------------------------------------------------------
    check('flag %: HOURLY text', (0, timenavi_1.percentIsFlagged)('HOURLY'), true);
    check('flag %: > 100%', (0, timenavi_1.percentIsFlagged)(1.5), true);
    check('flag %: < 0%', (0, timenavi_1.percentIsFlagged)(-0.1), true);
    check('flag %: 95% clean', (0, timenavi_1.percentIsFlagged)(0.95), false);
    check('flag %: blank clean', (0, timenavi_1.percentIsFlagged)(null), false);
    const flagRow = (id, inQboTime, included = true) => ({ id, included, inQboTime });
    const dups = (0, timenavi_1.duplicateInQboValues)([
        flagRow('1', 'AP1 Job'),
        flagRow('2', 'AP1 Job'),
        flagRow('3', 'AP2 Other'),
        flagRow('4', 'NotClockedIn'),
        flagRow('5', 'NotClockedIn'),
        flagRow('6', 'AP3 Excluded', false),
        flagRow('7', 'AP3 Excluded', false),
    ]);
    check('flag M: duplicate included values only', [...dups], ['AP1 Job']);
    check('flag $: nonzero pink', (0, timenavi_1.dollarDiffIsFlagged)(-0.01), true);
    check('flag $: zero clean', (0, timenavi_1.dollarDiffIsFlagged)(0), false);
    check('flag $: blank clean', (0, timenavi_1.dollarDiffIsFlagged)(null), false);
    // --- Round 4 §2: pivot view filters ---------------------------------------------------------
    const pivotModel = {
        title: 't',
        estimators: ['MH', 'PG'],
        crews: [
            {
                crew: "A's Crew",
                employees: [
                    {
                        username: 'ana', fname: 'Ana', lname: 'A',
                        jobs: [{ customer: 'C', job: 'J', byEstimator: { MH: 10, PG: 5 }, total: 15 }],
                        byEstimator: { MH: 10, PG: 5 }, total: 15, overtime: false,
                    },
                ],
            },
            {
                crew: "B's Crew",
                employees: [
                    {
                        username: 'bob', fname: 'Bob', lname: 'B',
                        jobs: [{ customer: 'C', job: 'J2', byEstimator: { PG: 8 }, total: 8 }],
                        byEstimator: { PG: 8 }, total: 8, overtime: false,
                    },
                ],
            },
        ],
        grandByEstimator: { MH: 10, PG: 13 },
        grandTotal: 23,
    };
    check('pivotView: inactive view returns input untouched', (0, pivotView_1.applyPivotView)(pivotModel, pivotView_1.EMPTY_PIVOT_VIEW) === pivotModel, true);
    const noPg = (0, pivotView_1.applyPivotView)(pivotModel, { ...pivotView_1.EMPTY_PIVOT_VIEW, hiddenEstimators: ['PG'] });
    check('pivotView: hiding PG drops its hours from every total', [noPg.grandTotal, noPg.crews[0].employees[0].total], [10, 10]);
    const oneCrew = (0, pivotView_1.applyPivotView)(pivotModel, { ...pivotView_1.EMPTY_PIVOT_VIEW, hiddenGroups: ["B's Crew"] });
    check('pivotView: group filter removes the block + recomputes', [oneCrew.crews.length, oneCrew.grandTotal], [1, 15]);
    const zToA = (0, pivotView_1.applyPivotView)(pivotModel, { ...pivotView_1.EMPTY_PIVOT_VIEW, columnSort: 'desc' });
    check('pivotView: Z→A column sort', zToA.estimators, ['PG', 'MH']);
    const noBob = (0, pivotView_1.applyPivotView)(pivotModel, { ...pivotView_1.EMPTY_PIVOT_VIEW, hiddenUsernames: ['bob'] });
    check('pivotView: username filter drops empty crew', [noBob.crews.length, noBob.grandTotal], [1, 15]);
    // --- Round 4 §3.2-2 / §5.3: display-name helpers ----------------------------------------------
    check('nameStrip: trailing initial', (0, payroll_1.stripSingleLetterTokens)('Rene Salazar M'), 'Rene Salazar');
    check('nameStrip: middle initial', (0, payroll_1.stripSingleLetterTokens)('Adriana G Lopez Barajas'), 'Adriana Lopez Barajas');
    check('nameStrip: dotted initial', (0, payroll_1.stripSingleLetterTokens)('John Q. Public'), 'John Public');
    check('nameStrip: untouched name', (0, payroll_1.stripSingleLetterTokens)('Carlos Ruiz Cruz'), 'Carlos Ruiz Cruz');
    check('titleCase: directory casing wins + initial stripped', (0, payroll_1.titleCaseLeadName)('RENE SALAZAR M', [{ fullName: 'Rene Salazar M', firstName: 'Rene', lastName: 'Salazar M', active: true }], []), 'Rene Salazar');
    check('titleCase: split leads resolved independently', (0, payroll_1.titleCaseLeadName)('TED FRANK/RAFAEL AGUILAR', dir.concat([]), []), 'Ted Frank/Rafael Aguilar');
    check("titleCase: sub keeps its stored text", (0, payroll_1.titleCaseLeadName)("PEPE'S PAINTING LLC", [], ["Pepe's Painting LLC"]), "Pepe's Painting LLC");
    // --- Round 4 §5.8: prevailing-wage rate chain --------------------------------------------------
    const pwSettings = {
        wageRates: [
            { name: 'Antonio Zuniga', rate: 18, hourlyRate: 25 },
            { name: 'Rafael Aguilar', rate: 25 },
        ],
        defaultWageRate: 18,
        prevailingWagePatterns: ['SNRHA *SS*'],
        prevailingMinRate: 19.18,
    };
    check('pw: pattern substring match', (0, payroll_1.isPrevailingWageJob)('AP24393 SNRHA *SS* 210 N 16th St', ['SNRHA *SS*']), true);
    check('pw: no match', (0, payroll_1.isPrevailingWageJob)('AP24314 Ext', ['SNRHA *SS*']), false);
    check('pw: default 18 lifted to 19.18 on SNRHA job', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24393 SNRHA *SS*', false, pwSettings), { rate: 19.18, defaulted: false, prevailing: true });
    check('pw: default 25 stays 25 on the same job', (0, payroll_1.resolveLineRate)('RAFAEL AGUILAR', 'AP24393 SNRHA *SS*', false, pwSettings).rate, 25);
    check('pw: hourly job uses the hourly-jobs rate', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24000 T&M repair', true, pwSettings), { rate: 25, defaulted: false, prevailing: false });
    check('pw: plain job → default rate', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24314 Ext', false, pwSettings).rate, 18);
    check('pw: unknown employee on plain job → fallback + flag', (0, payroll_1.resolveLineRate)('NOBODY', 'AP24314 Ext', false, pwSettings), { rate: 18, defaulted: true, prevailing: false });
    // --- Round 6 §6.2: profit-share chain ----------------------------------------------------------
    check('§6.2 Antonio: PS − labor − ½mat → 200.165', (0, payroll_1.computeTotalProfitShare)({
        profitShare: 432,
        workOrderOverride: null,
        laborWages: 188.11,
        halfMaterials: 43.725,
        percentCompletion: 1,
        mileage: 12.5,
    }), 200.165);
    check('§6.2 % Completion + mileage do not change TPS', (0, payroll_1.computeTotalProfitShare)({
        profitShare: 432,
        workOrderOverride: null,
        laborWages: 188.11,
        halfMaterials: 43.725,
        percentCompletion: 0.5,
        mileage: 999,
    }), 200.165);
    check('§6.2 WO override wins when set', (0, payroll_1.computeTotalProfitShare)({
        profitShare: 432,
        workOrderOverride: 500,
        laborWages: 100,
        halfMaterials: 50,
    }), 350);
    // --- Round 6 §7.1 Branch 1 + §7.2 job rules ----------------------------------------------------
    const ruleSettings = {
        ...pwSettings,
        jobRateRules: [
            {
                id: 'r1',
                apKey: '24314',
                rate: 22,
                employeeNames: [],
                includeSuffixes: true,
                note: 'all painters AP24314',
            },
            {
                id: 'r2',
                apKey: '24314',
                rate: 30,
                employeeNames: ['Leonardo'],
                includeSuffixes: true,
                note: 'Leonardo only',
            },
            {
                id: 'r3',
                apKey: '24393',
                rate: 15,
                employeeNames: [],
                includeSuffixes: true,
                note: 'below DB',
            },
        ],
    };
    check('§7.2 AP24314 @ 22 all painters beats employee default', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24314 Ext', false, { ...ruleSettings, apKey: '24314' }), { rate: 22, defaulted: false, prevailing: false, fromJobRule: true, ruleNote: 'all painters AP24314' });
    check('§7.2 named employee rule beats all-painters', (0, payroll_1.resolveLineRate)('Leonardo Perez', 'AP24314 Ext', false, { ...ruleSettings, apKey: '24314' }).rate, 30);
    check('§7.1 Branch 1: rule at $15 on DB job still pays 19.18', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24393 SNRHA *SS*', false, { ...ruleSettings, apKey: '24393' }).rate, 19.18);
    check('§7.1 forceNonDb skips Branch 1 floor', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24393 SNRHA *SS*', false, {
        ...ruleSettings,
        apKey: '24393',
        forceNonDb: true,
    }).rate, 15);
    check('§7.1 forceNonDb clears prevailing flag', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24393 SNRHA *SS*', false, {
        ...ruleSettings,
        apKey: '24393',
        forceNonDb: true,
    }).prevailing, false);
    // --- Round 6 §7.1 Branch 2: crew-lead Davis-Bacon floor ----------------------------------------
    check('§7.1 Branch 2: 10h DB, chain 100 → 191.80 + floorApplied', (0, payroll_1.applyDavisBaconLeadFloor)({
        isDbJob: true,
        isSubcontractor: false,
        timesheetHoursOnJob: 10,
        profitSharePay: 100,
        dbRate: 19.18,
    }), { pay: 191.8, floorApplied: true, floor: 191.8, priorPay: 100 });
    check('§7.1 Branch 2: non-DB same shape pays 100', (0, payroll_1.applyDavisBaconLeadFloor)({
        isDbJob: false,
        isSubcontractor: false,
        timesheetHoursOnJob: 10,
        profitSharePay: 100,
        dbRate: 19.18,
    }), { pay: 100, floorApplied: false, floor: 0, priorPay: 100 });
    check('§7.1 Branch 2: subcontractor exempt', (0, payroll_1.applyDavisBaconLeadFloor)({
        isDbJob: true,
        isSubcontractor: true,
        timesheetHoursOnJob: 10,
        profitSharePay: 100,
        dbRate: 19.18,
    }).floorApplied, false);
    check('§7.1 Branch 2: zero hours → no-op', (0, payroll_1.applyDavisBaconLeadFloor)({
        isDbJob: true,
        isSubcontractor: false,
        timesheetHoursOnJob: 0,
        profitSharePay: 100,
        dbRate: 19.18,
    }).floorApplied, false);
    // --- Round 6 §5.1: TimeNavi row format ---------------------------------------------------------
    check('TimeNavi row format', (0, pivotByJob_1.formatTimeNaviRow)(19), 'TimeNavi row 19');
    check('TimeNavi row parse labeled', (0, pivotByJob_1.parseTimeNaviRowRef)('TimeNavi row 19'), 19);
    check('TimeNavi row parse bare', (0, pivotByJob_1.parseTimeNaviRowRef)('19'), 19);
    check('TimeNavi row parse junk', (0, pivotByJob_1.parseTimeNaviRowRef)('NOT ON CALENDAR'), null);
    // --- Round 7 §2: PR Final Check rows -----------------------------------------------------------
    const plForPr = {
        byJob: new Map([
            [
                '24314',
                {
                    income: 650,
                    materials: 35.74,
                    laborWagesSubs: 113.04,
                    mileage: 0,
                    sundries: 0,
                    plJobHeader: 'AP24314 Ext',
                    customerName: 'Aaron Ward',
                },
            ],
            [
                '24403',
                {
                    income: 9000,
                    materials: 760.1,
                    laborWagesSubs: 2081.09,
                    mileage: 0,
                    sundries: 0,
                    plJobHeader: 'AP24403 Lambrook',
                    customerName: 'CCSD',
                },
            ],
        ]),
        duplicateColumns: [],
        columnHeaders: ['', 'Aaron Ward', 'AP24314 Ext', 'AP24403 Lambrook'],
        accountRows: [
            { label: 'Total Income', values: [null, null, 650, 9000] },
            { label: 'Employee Labor', values: [null, null, 113.04, 2081.09] },
        ],
    };
    const prRows = (0, prFinalCheck_1.buildPrFinalCheckRows)(plForPr);
    check('buildPrFinalCheckRows length', prRows.length, 2);
    check('buildPrFinalCheckRows first job', [prRows[0].apKey, prRows[0].income, prRows[0].customerName], ['24314', 650, 'Aaron Ward']);
    check('buildPrFinalCheckRows empty P&L', (0, prFinalCheck_1.buildPrFinalCheckRows)({
        byJob: new Map(),
        duplicateColumns: [],
        columnHeaders: [],
        accountRows: [],
    }).length, 0);
    // --- Round 7 §20.1: crew-group owner becomes a lead --------------------------------------------
    const ownerDay = new Date(2026, 5, 30); // 6/30
    const ownerEnriched = [
        {
            username: 'antonio@x.com',
            payrollId: '',
            fname: 'Antonio',
            lname: 'Zuniga',
            number: '',
            group: "Antonio Zuniga's Crew",
            localDate: ownerDay,
            localDay: 'Tue',
            localStartTime: null,
            localEndTime: null,
            tz: '',
            hours: 8,
            jobcode1: 'Cust',
            jobcode2: 'AP25000 Ext',
            billable: '',
            klass: '',
            department: '',
            payrollItem: '',
            serviceItem: '',
            location: '',
            notes: '',
            approvedStatus: '',
            hasFlags: '',
            flagTypes: '',
            sourceEntity: 'NV',
            apKey: '25000',
            estimator: 'MH',
            notOnCalendar: false,
        },
    ];
    const ownerPbj = {
        title: 'NV PAYROLL: 7/10/2026',
        dates: [ownerDay],
        dateIsos: ['2026-06-30'],
        customers: [
            {
                customer: 'Cust',
                byDate: { '2026-06-30': 8 },
                total: 8,
                colorHex: 'FFFF99',
                jobs: [
                    {
                        id: 'Cust\u0000AP25000 Ext',
                        jobcode2: 'AP25000 Ext',
                        byDate: { '2026-06-30': 8 },
                        total: 8,
                        colorHex: 'FFFF99',
                        employees: [{ username: 'antonio@x.com', byDate: { '2026-06-30': 8 }, total: 8, qbHours: null }],
                        analysis: {
                            apKey: '25000',
                            onCalendar: 'NOT ON CALENDAR',
                            estimator: 'MH',
                            crewLead: '',
                            crewLeadResolution: null,
                            woOverride: null,
                            woOverrideSuggestion: null,
                            woOverrideSuggestionSource: null,
                            woPct: null,
                            woPctSuggestion: null,
                            woPctSuggestionSource: null,
                            percent: 0.5,
                            notes: '',
                            hourly: false,
                            sub: false,
                            split: 1,
                            dbPatternMatch: false,
                            isDbJob: false,
                            dbOverride: false,
                        },
                    },
                ],
            },
        ],
        grandByDate: { '2026-06-30': 8 },
        grandTotal: 8,
        balances: [],
        utPayroll: [],
    };
    const ownerSettings = (0, settings_1.defaultSettings)();
    ownerSettings.salariedRoster = [];
    ownerSettings.profitShare = [];
    const ownerPreview = (0, payrollPreview_1.buildPayrollPreview)({
        pbj: ownerPbj,
        enriched: ownerEnriched,
        reportRows: [],
        pl: {
            byJob: new Map([
                [
                    '25000',
                    {
                        income: 1000,
                        materials: 0,
                        laborWagesSubs: 0,
                        mileage: 0,
                        sundries: 0,
                        plJobHeader: 'AP25000 Ext',
                        customerName: 'Cust',
                    },
                ],
            ]),
            duplicateColumns: [],
            columnHeaders: ['', 'AP25000 Ext'],
            accountRows: [{ label: 'Total Income', values: [null, 1000] }],
        },
        directory: [
            { fullName: 'Antonio Zuniga', firstName: 'Antonio', lastName: 'Zuniga', active: true },
        ],
        settings: ownerSettings,
        payDate: new Date(2026, 6, 10),
        edits: {},
        qbHours: {},
        knownEstimators: new Set(['MH']),
        leadCtx: {
            directory: [
                { fullName: 'Antonio Zuniga', firstName: 'Antonio', lastName: 'Zuniga', active: true },
            ],
            crewGroups: ["Antonio Zuniga's Crew"],
            subcontractors: [],
        },
    });
    check('crew-owner lead: employee in "X\'s Crew" becomes a lead', ownerPreview.leads.map((l) => l.leadName), ['ANTONIO ZUNIGA']);
    check('crew-owner lead: not also an hourly painter', ownerPreview.painters.length, 0);
    // --- Round 5 §4: job-naming chain --------------------------------------------------------------
    check('baseApKey: suffixed', (0, jobNames_1.baseApKey)('24180-A'), '24180');
    check('baseApKey: multi-suffix', (0, jobNames_1.baseApKey)('21848-A5'), '21848');
    check('baseApKey: bare → null', (0, jobNames_1.baseApKey)('24314'), null);
    const emptySrc = {
        timesheetByAp: new Map(),
        carryByAp: new Map(),
        salesByAp: new Map(),
        pl: null,
    };
    check('jobNames: timesheet wins', (0, jobNames_1.jobDisplayPair)('24314', {
        ...emptySrc,
        timesheetByAp: new Map([['24314', { customer: 'Aaron Ward', job: 'AP24314 Ext' }]]),
    }), { poLabel: 'AP24314 Ext', customerName: 'Aaron Ward', usedBaseKey: false, unresolved: false });
    check('jobNames: QBO exact', (0, jobNames_1.jobDisplayPair)('24180', {
        ...emptySrc,
        salesByAp: new Map([['24180', { name1: 'Regina Conrad', name2: 'AP24180 Int' }]]),
    }), { poLabel: 'AP24180 Int', customerName: 'Regina Conrad', usedBaseKey: false, unresolved: false });
    check('jobNames: base-key retry', (0, jobNames_1.jobDisplayPair)('24180-A', {
        ...emptySrc,
        salesByAp: new Map([['24180', { name1: 'Regina Conrad', name2: 'AP24180 Int' }]]),
    }), { poLabel: 'AP24180 Int', customerName: 'Regina Conrad', usedBaseKey: true, unresolved: false });
    check('jobNames: carry-forward before QBO', (0, jobNames_1.jobDisplayPair)('21848-A5', {
        ...emptySrc,
        carryByAp: new Map([['21848-A5', { customer: 'Old Customer', job: 'AP21848-A5 Somewhere Int' }]]),
        salesByAp: new Map([['21848', { name1: 'Wrong', name2: 'Wrong' }]]),
    }), { poLabel: 'AP21848-A5 Somewhere Int', customerName: 'Old Customer', usedBaseKey: false, unresolved: false });
    const cleaned = (0, jobNames_1.cleanEventTitle)('♦Int 13k (XC) Regina Conrad Ph1, AP24180-A *PAID* (Daniel)', '24180-A');
    check('jobNames: cleaned title — no ♦ / no raw dump', cleaned?.poLabel.includes('♦') ?? true, false);
    check('jobNames: cleaned title customer', cleaned?.customerName, 'Regina Conrad');
    check('jobNames: cleaned title PO has Int', cleaned?.poLabel.includes('Int'), true);
    check('jobNames: unresolved → AP{key} + blank', (0, jobNames_1.jobDisplayPair)('99999', emptySrc), { poLabel: 'AP99999', customerName: '', usedBaseKey: false, unresolved: true });
    // --- Round 6 §2.2 Slot A: completed %'s workbook ------------------------------
    const dA = new Date(2026, 5, 27); // 6/27
    const dB = new Date(2026, 5, 28);
    // App 2-row header: analysis labels on row 0, weekdays on row 1.
    const twoRow = [
        [
            'NV PAYROLL: 7/10/2026',
            dA,
            dB,
            'Total hrs',
            '',
            'AP#',
            'On Calendar?',
            'Estimator',
            'Crew Lead',
            'WO Override',
            'WO%',
            '% of Completion',
            'Notes',
        ],
        ['', 'Fri', 'Sat', '', '', '', '', '', '', '', '', '', ''],
    ];
    check('pbj header: app 2-row', (0, completedPercents_1.findPbjHeaderRow)(twoRow), 0);
    // Legacy human 4-row: title / blank / sum label / analysis headers.
    const fourRow = [
        ['NV PAYROLL: 7/10/2026'],
        [''],
        ['Sum of hours'],
        [
            'Row Labels',
            dA,
            dB,
            'Total hrs',
            'AP#',
            'On Calendar?',
            'Estimator',
            'Crew Lead',
            'WO Override',
            'WO%',
            '% of Completion',
            'Notes',
        ],
    ];
    check('pbj header: legacy 4-row', (0, completedPercents_1.findPbjHeaderRow)(fourRow), 3);
    const missingTs = {
        sheets: [{ name: 'Pivot by Job', rows: twoRow }],
    };
    const miss = (0, completedPercents_1.parseCompletedPercentsWorkbook)(missingTs);
    check('slot A: missing timesheets sheet', 'error' in miss
        ? miss.error
        : '', "This file has no 'timesheets' sheet — export it from the %'s process first");
    const missingPbj = {
        sheets: [
            {
                name: 'timesheets',
                rows: [
                    ['username', 'group', 'local_date', 'hours', 'jobcode_1', 'jobcode_2'],
                    ['ana.b', "Carlos's Crew", dA, 8, 'Cust', 'AP24314 Ext'],
                ],
            },
        ],
    };
    const missPbj = (0, completedPercents_1.parseCompletedPercentsWorkbook)(missingPbj);
    check('slot A: missing Pivot by Job sheet', 'error' in missPbj ? missPbj.error : '', "This file has no 'Pivot by Job' sheet — export it from the %'s process first");
    const failed = results.filter((r) => !r.pass);
    const label = `[% of Completion Builder self-test] ${results.length - failed.length}/${results.length} passed`;
    if (failed.length === 0) {
        console.log(`%c${label}`, 'color: green; font-weight: bold');
    }
    else {
        console.warn(`%c${label}`, 'color: red; font-weight: bold');
        for (const f of failed)
            console.warn(`  FAIL ${f.name}: ${f.detail ?? ''}`);
    }
    return failed.length === 0;
}
