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
const customerByAp_1 = require("./customerByAp");
const exportXlsx_1 = require("./exportXlsx");
const exportRound6Sheets_1 = require("./exportRound6Sheets");
const fills_1 = require("./fills");
const parsers_1 = require("./parsers");
const payroll_1 = require("./payroll");
const payTypes_1 = require("./payTypes");
const payrollPivot_1 = require("./payrollPivot");
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
    /** Trim float noise so money comparisons survive JSON round-tripping. */
    const round2ish = (n) => n === null ? null : Math.round(n * 1e6) / 1e6;
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
    // --- Round 8 §2.3: % Completion free text ------------------------------------
    check('pct free: 38%', (0, format_1.parsePercentInput)('38%'), 0.38);
    check('pct free: .38', (0, format_1.parsePercentInput)('.38'), 0.38);
    check('pct free: 38', (0, format_1.parsePercentInput)('38'), 0.38);
    check('pct free: 0% stays zero', (0, format_1.parsePercentInput)('0%'), 0);
    check('pct free: hourly any case', (0, format_1.parsePercentInput)('hOuRlY'), 'HOURLY');
    check('pct free: arbitrary text kept verbatim', (0, format_1.parsePercentInput)('waiting on Jose'), {
        kind: 'text',
        text: 'waiting on Jose',
    });
    check('pct free: text predicate', (0, format_1.isPercentText)((0, format_1.parsePercentInput)('see notes')), true);
    check('pct free: text accessor', (0, format_1.percentText)((0, format_1.parsePercentInput)('  see notes  ')), 'see notes');
    check('pct free: number is not text', (0, format_1.isPercentText)(0.38), false);
    check('pct free: blank still null', (0, format_1.parsePercentInput)('   '), null);
    check('NOTES WO note: full amount, PS %, % complete', (0, format_1.formatCrewLeadWoNote)({
        workOrderOverride: null,
        profitShare: 1159.2,
        woPct: 0.4,
        percent: 0.9,
        percentCompletion: 0.9,
        missingPercent: false,
        manualPay: null,
    }).text, 'w.o. $1,159.20 @40% PS & 90% complete');
    check('NOTES WO note: missing % flags ? Complete', (0, format_1.formatCrewLeadWoNote)({
        workOrderOverride: 1159.2,
        profitShare: 0,
        woPct: 0.4,
        percent: null,
        percentCompletion: 0,
        missingPercent: true,
        manualPay: null,
    }), { text: 'w.o. $1,159.20 @40% PS & ? Complete', flagMissing: true });
    // --- Round 8 §2.4: SUBBED / SET RATE forgiving matching ----------------------
    check('subbed: plain', (0, format_1.matchesSubbedOut)('subbed out'), true);
    check('subbed: hyphenated + case', (0, format_1.matchesSubbedOut)('Subbed-out'), true);
    check('subbed: double space + caps', (0, format_1.matchesSubbedOut)('SUBBED  OUT'), true);
    check('subbed: inside a status line', (0, format_1.matchesSubbedOut)('7/10/26 subbed out to Pepe'), true);
    check('subbed: unrelated text', (0, format_1.matchesSubbedOut)('substitute crew'), false);
    check('set rate: trailing period', (0, format_1.matchesSetRate)('set rate.'), true);
    check('set rate: caps + extra spaces', (0, format_1.matchesSetRate)('  SET   RATE  '), true);
    check('set rate: hyphenated', (0, format_1.matchesSetRate)('Set-Rate'), true);
    check('set rate: unrelated text', (0, format_1.matchesSetRate)('rate set by Jodi'), false);
    check('pct input: subbed out → SUBBED', (0, format_1.parsePercentInput)('Subbed-out'), 'SUBBED');
    check('pct input: set rate → SET RATE', (0, format_1.parsePercentInput)('set rate.'), 'SET RATE');
    check('%: description "subbed out" on the selected line', (0, timenavi_1.suggestPercent)('7/10/26 subbed out to Pepe', day(7, 10, 2026)).value, 'SUBBED');
    check('%: description "set rate" on the selected line', (0, timenavi_1.suggestPercent)('7/10/26 Set Rate - no % tracking', day(7, 10, 2026)).value, 'SET RATE');
    check('%: a real % on the line still wins over subbed wording', (0, timenavi_1.suggestPercent)('7/10/26 30% complete, rest subbed out', day(7, 10, 2026)).value, 0.3);
    // --- Round 8 §2.4: invalid-% red predicate ----------------------------------
    check('red %: blank', (0, timenavi_1.percentIsRedFlagged)(null), true);
    check('red %: zero', (0, timenavi_1.percentIsRedFlagged)(0), true);
    check('red %: free text', (0, timenavi_1.percentIsRedFlagged)({ kind: 'text', text: '??' }), true);
    check('red %: valid fraction', (0, timenavi_1.percentIsRedFlagged)(0.38), false);
    check('red %: over 100% is valid (pink, not red)', (0, timenavi_1.percentIsRedFlagged)(1.5), false);
    check('red %: HOURLY', (0, timenavi_1.percentIsRedFlagged)('HOURLY'), false);
    check('red %: SUBBED', (0, timenavi_1.percentIsRedFlagged)('SUBBED'), false);
    check('red %: SET RATE', (0, timenavi_1.percentIsRedFlagged)('SET RATE'), false);
    check('red %: cleared by a manual edit', (0, timenavi_1.percentIsRedFlagged)((0, format_1.parsePercentInput)('38%')), false);
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
    check('titleCase: unmatched ALL CAPS still Title Case', (0, payroll_1.titleCaseLeadName)('CARLOS RUIZ CRUZ', [], []), 'Carlos Ruiz Cruz');
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
    // --- Round 8 §8.1: profit-share chain (supersedes Round 6 §6.2) --------------------------------
    const antonioChain = (percent) => (0, payroll_1.computeTotalProfitShare)({
        profitShare: 432,
        workOrderOverride: null,
        laborWages: 188.11,
        halfMaterials: 43.725,
        percent,
        mileage: 12.5,
    });
    check('§8.1 Antonio @100%: reproduces the Round 6 anchor 200.165', antonioChain(1).total, 200.165);
    check('§8.1 Antonio @50%: base term halves (432*.5 − 188.11 − 43.725)', round2ish(antonioChain(0.5).total), round2ish(432 * 0.5 - 188.11 - 43.725));
    check('§8.1 mileage is displayed, never a term', antonioChain(1).base, 432);
    check('§8.1 WO override multiplies by % too', (0, payroll_1.computeTotalProfitShare)({
        profitShare: 432,
        workOrderOverride: 500,
        laborWages: 100,
        halfMaterials: 50,
        percent: 0.5,
    }), { total: 100, base: 500, percentCompletion: 0.5, manualPay: null, missingPercent: false });
    check('§8.1 WO override @100% = Round 6 behaviour', (0, payroll_1.computeTotalProfitShare)({
        profitShare: 432,
        workOrderOverride: 500,
        laborWages: 100,
        halfMaterials: 50,
        percent: 1,
    }).total, 350);
    for (const state of ['SUBBED', 'SET RATE']) {
        const r = (0, payroll_1.computeTotalProfitShare)({
            profitShare: 432,
            workOrderOverride: null,
            laborWages: 188.11,
            halfMaterials: 43.725,
            percent: state,
        });
        check(`§8.1 ${state}: blank total + manualPay flag`, [r.total, r.manualPay, r.missingPercent], [
            null,
            state,
            false,
        ]);
    }
    const blankPct = antonioChain(null);
    check('§8.1 blank % → pays 0 of the base (never 100%) + missingPercent', [blankPct.percentCompletion, blankPct.missingPercent, round2ish(blankPct.total)], [0, true, round2ish(-188.11 - 43.725)]);
    check('§8.1 free text % → treated as blank', (0, payroll_1.computeTotalProfitShare)({
        profitShare: 432,
        workOrderOverride: null,
        laborWages: 0,
        halfMaterials: 0,
        percent: { kind: 'text', text: 'see notes' },
    }), { total: 0, base: 432, percentCompletion: 0, manualPay: null, missingPercent: true });
    // --- Round 8 §8.2: Total Pay subtracts Deductions/Sundries -------------------------------------
    check('§8.2 deductions are entered positive and SUBTRACTED', (0, payroll_1.computeTotalPay)({ total: 1000, hourlyPay: 200, reimbursements: 50, deductions: 75 }), 1175);
    check('§8.2 no deductions → total + hourly + reimbursements', (0, payroll_1.computeTotalPay)({ total: 1000, hourlyPay: 200, reimbursements: 50, deductions: 0 }), 1250);
    check('§8.2 deductions can push pay below the profit share', (0, payroll_1.computeTotalPay)({ total: 500, hourlyPay: 0, reimbursements: 0, deductions: 120 }), 380);
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
    // --- Round 8 §9.1 / Round 10: hourly vs profit-share rates ------------------
    const r9 = {
        wageRates: [
            { name: 'Antonio Zuniga', rate: 16, hourlyRate: 27 },
            { name: 'Rafael Aguilar', rate: 25 },
            { name: 'Hilda Ruiz', rate: 16 },
        ],
        defaultWageRate: 18,
        prevailingWagePatterns: ['SNRHA'],
        prevailingMinRate: 19.18,
    };
    check('§9.1 lead on an HOURLY job → hourly-jobs rate', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24000 T&M repair', true, { ...r9, isCrewLead: true }).rate, 27);
    // Round 10 — a lead on a job they lead is on a profit-share job, so the PS
    // rate IS their crew-lead rate; the separate rate column is gone.
    check("Round 10 lead's T-SHEETS on a profit-share job → their PS rate", (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24314 Ext', false, { ...r9, isCrewLead: true }), { rate: 16, defaulted: false, prevailing: false, fromCrewLeadRate: true });
    check('Round 10 lead with only a PS rate on file → that rate, not a global lead rate', (0, payroll_1.resolveLineRate)('RAFAEL AGUILAR', 'AP24314 Ext', false, { ...r9, isCrewLead: true }).rate, 25);
    check('§9.1 helper on a profit-share job → their painter rate', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24314 Ext', false, r9).rate, 16);
    check('§9.1 helper with no hourlyRate on an HOURLY job → default', (0, payroll_1.resolveLineRate)('HILDA RUIZ', 'AP24000 T&M repair', true, r9).rate, 18);
    check('§9.1 SNRHA substring floors a $16 painter at 19.18', (0, payroll_1.resolveLineRate)('ANTONIO ZUNIGA', 'AP24393 SNRHA *SS* 210 N 16th', false, r9).rate, 19.18);
    check('§9.1 SNRHA never LOWERS a $25 painter', (0, payroll_1.resolveLineRate)('RAFAEL AGUILAR', 'AP24393 SNRHA *SS*', false, r9).rate, 25);
    check('Carlos Ruiz Cruz HOURLY job → $25 (not the $18 PS/crew-lead rate)', (0, payroll_1.resolveLineRate)('Carlos Ruiz Cruz', 'AP24000 repair', true, {
        wageRates: (0, settings_1.defaultSettings)().wageRates,
        defaultWageRate: 18,
        prevailingWagePatterns: ['SNRHA'],
        prevailingMinRate: 19.18,
        isCrewLead: true,
    }).rate, 25);
    check("Carlos Ruiz Cruz profit-share hours as lead → $18", (0, payroll_1.resolveLineRate)('Carlos Ruiz Cruz', 'AP24314 Ext', false, {
        wageRates: (0, settings_1.defaultSettings)().wageRates,
        defaultWageRate: 18,
        prevailingWagePatterns: ['SNRHA'],
        prevailingMinRate: 19.18,
        isCrewLead: true,
    }).rate, 18);
    check('§9.1 conditional rule beats the generic rates', (0, payroll_1.resolveLineRate)('HILDA RUIZ', 'AP24314 Ext', false, {
        ...r9,
        jobLeadNames: ['ANTONIO ZUNIGA'],
        conditionalRateRules: [
            {
                id: 'c1',
                person: 'Hilda Ruiz',
                rate: 23,
                condition: { kind: 'onSameJobAsLead', leadName: 'Antonio Zuniga' },
                note: "Hilda on Antonio's jobs",
            },
        ],
    }), {
        rate: 23,
        defaulted: false,
        prevailing: false,
        fromConditionalRule: true,
        ruleNote: "Hilda on Antonio's jobs",
    });
    check('§9.1 conditional rule needs the named lead on the job', (0, payroll_1.resolveLineRate)('HILDA RUIZ', 'AP24314 Ext', false, {
        ...r9,
        jobLeadNames: ['TED FRANK'],
        conditionalRateRules: [
            {
                id: 'c1',
                person: 'Hilda Ruiz',
                rate: 23,
                condition: { kind: 'onSameJobAsLead', leadName: 'Antonio Zuniga' },
                note: '',
            },
        ],
    }).rate, 16);
    const perJobSettings = {
        ...r9,
        apKey: '24314',
        jobRateRules: [
            { id: 'j1', apKey: '24314', rate: 22, employeeNames: [], includeSuffixes: true, note: 'job rule' },
        ],
        perJobRateOverrides: [
            { id: 'o1', person: 'Hilda Ruiz', apKeyPattern: '24314', rate: 26, note: 'Hilda AP24314' },
        ],
    };
    check('§9.1 per-job override outranks the job-scoped rule', (0, payroll_1.resolveLineRate)('HILDA RUIZ', 'AP24314 Ext', false, perJobSettings), { rate: 26, defaulted: false, prevailing: false, fromPerJobOverride: true, ruleNote: 'Hilda AP24314' });
    check('§9.1 someone else on the same job still gets the job rule', (0, payroll_1.resolveLineRate)('RAFAEL AGUILAR', 'AP24314 Ext', false, perJobSettings).rate, 22);
    check('§9.1 DB max-wrap sits above the per-job override', (0, payroll_1.resolveLineRate)('HILDA RUIZ', 'AP24393 SNRHA *SS*', false, {
        ...perJobSettings,
        apKey: '24393',
        perJobRateOverrides: [
            { id: 'o1', person: 'Hilda Ruiz', apKeyPattern: '24393', rate: 15, note: 'below DB' },
        ],
    }).rate, 19.18);
    // --- Round 8 §9.2: NV daily overtime -----------------------------------------------------------
    const nvDay = (0, payroll_1.splitDailyOvertime)([{ jobId: 'a', date: day2(7, 1), hours: 10, rowIndex: 0, eligible: true }], 8);
    check('§9.2 10-hr NV day at $16 → 8 straight + 2 @ 1.5×', [nvDay.totalOtHours, nvDay.otByJob.get('a'), nvDay.segments[0]?.dayHours], [2, 2, 10]);
    const utDay = (0, payroll_1.splitDailyOvertime)([{ jobId: 'a', date: day2(7, 1), hours: 10, rowIndex: 0, eligible: false }], 8);
    check('§9.2 same day in UT → no split', [utDay.totalOtHours, utDay.segments.length], [0, 0]);
    const richPainter = (0, payroll_1.splitDailyOvertime)([{ jobId: 'a', date: day2(7, 1), hours: 10, rowIndex: 0, eligible: false }], 8);
    check('§9.2 NV painter at/above $18 → no daily OT', richPainter.totalOtHours, 0);
    const eightHourDay = (0, payroll_1.splitDailyOvertime)([{ jobId: 'a', date: day2(7, 1), hours: 8, rowIndex: 0, eligible: true }], 8);
    check('§9.2 exactly 8 hrs → none', eightHourDay.totalOtHours, 0);
    const twoJobDay = (0, payroll_1.splitDailyOvertime)([
        { jobId: 'a', date: day2(7, 1), hours: 6, rowIndex: 0, eligible: true },
        { jobId: 'b', date: day2(7, 1), hours: 5, rowIndex: 1, eligible: true },
    ], 8);
    check('§9.2 overage lands on the day’s LAST job', [twoJobDay.totalOtHours, twoJobDay.otByJob.get('b'), twoJobDay.otByJob.get('a')], [3, 3, undefined]);
    const mixedDay = (0, payroll_1.splitDailyOvertime)([
        { jobId: 'nv', date: day2(7, 1), hours: 8, rowIndex: 0, eligible: true },
        { jobId: 'ut', date: day2(7, 1), hours: 4, rowIndex: 1, eligible: false },
    ], 8);
    check('§9.2 UT hours consume the overage but never become OT', mixedDay.totalOtHours, 0);
    // 45-hr week with 5 hrs already paid as daily OT → the weekly rule adds none.
    const weekWithDailyOt = (0, payroll_1.splitOvertime)([
        { jobId: 'a', hours: 40, lastDate: day2(7, 1), lastRowIndex: 0 },
        { jobId: 'b', hours: 5, lastDate: day2(7, 3), lastRowIndex: 4 },
    ], { dailyOtHours: 5 });
    check('§9.2 no double-count: daily OT consumes the >40 overage', [weekWithDailyOt.otJobId, weekWithDailyOt.otHours], [null, 0]);
    const weekPartialDailyOt = (0, payroll_1.splitOvertime)([
        { jobId: 'a', hours: 40, lastDate: day2(7, 1), lastRowIndex: 0 },
        { jobId: 'b', hours: 8, lastDate: day2(7, 3), lastRowIndex: 4 },
    ], { dailyOtHours: 5 });
    check('§9.2 remaining >40 hours still pay weekly OT', [weekPartialDailyOt.otJobId, weekPartialDailyOt.otHours], ['b', 3]);
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
                    travel: 0,
                    plJobHeader: 'AP24314 Ext',
                    customerName: 'Aaron Ward',
                    plEntity: 'NV',
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
                    travel: 0,
                    plJobHeader: 'AP24403 Lambrook',
                    customerName: 'CCSD',
                    plEntity: 'NV',
                },
            ],
        ]),
        duplicateColumns: [],
        columnHeaders: ['', 'Aaron Ward', 'AP24314 Ext', 'AP24403 Lambrook'],
        accountRows: [
            { label: 'Total Income', values: [null, null, 650, 9000] },
            { label: 'Employee Labor', values: [null, null, 113.04, 2081.09] },
        ],
        entity: 'NV',
        unmappedCogsRows: [],
    };
    const prRows = (0, prFinalCheck_1.buildPrFinalCheckRows)(plForPr);
    check('buildPrFinalCheckRows length', prRows.length, 2);
    check('buildPrFinalCheckRows first job', [prRows[0].apKey, prRows[0].income, prRows[0].customerName], ['24314', 650, 'Aaron Ward']);
    check('buildPrFinalCheckRows empty P&L', (0, prFinalCheck_1.buildPrFinalCheckRows)({
        byJob: new Map(),
        duplicateColumns: [],
        columnHeaders: [],
        accountRows: [],
        entity: 'NV',
        unmappedCogsRows: [],
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
                            livePercent: 0.5,
                            percentOverride: false,
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
        unresolvedCarriedCustomers: [],
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
                        travel: 0,
                        plJobHeader: 'AP25000 Ext',
                        customerName: 'Cust',
                        plEntity: 'NV',
                    },
                ],
            ]),
            duplicateColumns: [],
            columnHeaders: ['', 'AP25000 Ext'],
            accountRows: [{ label: 'Total Income', values: [null, 1000] }],
            entity: 'NV',
            unmappedCogsRows: [],
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
    // --- Round 10 --------------------------------------------------------------
    check('R10 accounting negatives parse back from parentheses', [(0, format_1.parseMoneyInput)('(39.90)'), (0, format_1.parseMoneyInput)('$1,234.50'), (0, format_1.parseMoneyInput)('(1,000)')], [-39.9, 1234.5, -1000]);
    check('R10 negatives carry the red class, positives do not', [(0, format_1.negativeClass)(-1), (0, format_1.negativeClass)(0), (0, format_1.negativeClass)(12), (0, format_1.negativeClass)(null)], [format_1.NEGATIVE_TEXT_CLASS, '', '', '']);
    check('R10 percents filename leads with the % symbol', (0, exportXlsx_1.percentsExportFilename)(new Date(2026, 7, 14)), '% of Completion 08-14-2026.xlsx');
    check('R10 the office jobcode is not a customer name', [(0, jobNames_1.isOfficeCustomer)('ALLPRO OFFICE'), (0, jobNames_1.isOfficeCustomer)('Allpro  Office'), (0, jobNames_1.isOfficeCustomer)('Allpro Painters')], [true, true, false]);
    check('R10 Sales by Rep Name 1 outranks the timesheet customer', (0, jobNames_1.jobDisplayPair)('24473', {
        timesheetByAp: new Map([['24473', { customer: 'ALLPRO OFFICE', job: 'AP24473 Int' }]]),
        carryByAp: new Map(),
        salesByAp: new Map([['24473', { name1: 'Cerris Systems', name2: 'AP24473 Int' }]]),
        pl: null,
    }), { poLabel: 'AP24473 Int', customerName: 'Cerris Systems', usedBaseKey: false, unresolved: false });
    check('R10 "Regular Pay Overtime" is overtime at the regular rate slot', (0, payTypes_1.classifyPayItem)('Regular Pay Overtime'), { kind: 'overtime', rate: 'regular', reason: 'weekly' });
    check('R10 last % paid takes the LAST non-blank % for an AP#', [
        ...(0, parsers_1.parseLastWeek)({
            sheets: [
                {
                    name: 'TimeNavi Export',
                    rows: [
                        ['AP#', '% Completion'],
                        ['24473', 0.5],
                        ['24500', 0.9],
                        ['24473', 0.75],
                    ],
                },
            ],
        }).entries(),
    ], [
        ['24473', 0.75],
        ['24500', 0.9],
    ]);
    // --- Round 9 §16/§18 — pay items drive the rate slot and the overtime ------
    check('payItem: Regular Pay is straight time on the PS rate', (0, payTypes_1.classifyPayItem)('Regular Pay'), {
        kind: 'regular',
        rate: 'regular',
    });
    check('payItem: Crew Lead - Hourly Wage', (0, payTypes_1.classifyPayItem)('Crew Lead - Hourly Wage'), {
        kind: 'regular',
        rate: 'crewLead',
    });
    check('payItem: Commercial Hourly', (0, payTypes_1.classifyPayItem)('Commercial Hourly'), {
        kind: 'regular',
        rate: 'commercialHourly',
    });
    check('payItem: Wages 2', (0, payTypes_1.classifyPayItem)('Wages 2'), { kind: 'regular', rate: 'wages2' });
    check('payItem: Overtime Pay is weekly OT on the PS rate', (0, payTypes_1.classifyPayItem)('Overtime Pay'), {
        kind: 'overtime',
        rate: 'regular',
        reason: 'weekly',
    });
    check('payItem: Crew Lead - Hourly Wage Overtime is OT on the lead rate', (0, payTypes_1.classifyPayItem)('Crew Lead - Hourly Wage Overtime'), { kind: 'overtime', rate: 'crewLead', reason: 'weekly' });
    check('payItem: Daily Overtime', (0, payTypes_1.classifyPayItem)('Daily Overtime'), {
        kind: 'overtime',
        rate: 'regular',
        reason: 'daily',
    });
    check('payItem: unknown item is never overtime', (0, payTypes_1.classifyPayItem)('Holiday Bonus'), {
        kind: 'regular',
        rate: 'regular',
    });
    check('payItem: blank is never overtime', (0, payTypes_1.classifyPayItem)(''), { kind: 'regular', rate: 'regular' });
    // Round 9 §18 — the per-painter pay-type rates.
    const r9Rates = [
        { name: 'Samuel Brewster', rate: 18, commercialHourlyRate: 30 },
        { name: 'Xavier Green', rate: 25, hourlyRate: 18 },
        { name: 'Damian Vasquez', rate: 18, wages2Rate: 21 },
    ];
    check('§18 Commercial Hourly uses the person\u2019s commercial rate', (0, payroll_1.resolveWageRate)('Samuel Brewster', {
        wageRates: r9Rates,
        defaultWageRate: 18,
        hourlyJob: true,
        payRateKind: 'commercialHourly',
    }).rate, 30);
    check('§35 Commercial Hourly with no commercial rate: hourly job \u2192 hourly rate', (0, payroll_1.resolveWageRate)('Xavier Green', {
        wageRates: r9Rates,
        defaultWageRate: 18,
        hourlyJob: true,
        payRateKind: 'commercialHourly',
    }).rate, 18);
    check('§35 Commercial Hourly with no commercial rate: work-order job \u2192 PS rate', (0, payroll_1.resolveWageRate)('Xavier Green', {
        wageRates: r9Rates,
        defaultWageRate: 18,
        hourlyJob: false,
        payRateKind: 'commercialHourly',
    }).rate, 25);
    check('§18 Wages 2 uses the person\u2019s Wages 2 rate', (0, payroll_1.resolveWageRate)('Damian Vasquez', {
        wageRates: r9Rates,
        defaultWageRate: 18,
        hourlyJob: false,
        payRateKind: 'wages2',
    }).rate, 21);
    check('§13 crew-lead pay item is tagged as crew-lead hours (paid at the PS rate)', (0, payroll_1.resolveWageRate)('Samuel Brewster', {
        wageRates: r9Rates,
        defaultWageRate: 18,
        hourlyJob: false,
        payRateKind: 'crewLead',
    }), { rate: 18, defaulted: false, kind: 'crewLead' });
    check('§9.1 still wins: an HOURLY job pays the hourly rate even on a crew-lead item', (0, payroll_1.resolveWageRate)('Xavier Green', {
        wageRates: r9Rates,
        defaultWageRate: 18,
        hourlyJob: true,
        payRateKind: 'crewLead',
    }).rate, 18);
    // --- Round 9 §6 — the payroll pivot is built from Slot A verbatim ----------
    const slotAJobs = [
        {
            section: 'nv',
            customer: 'Pepe Customer',
            job: 'AP24258 Ext',
            apKey: '24258',
            onCalendar: '-',
            estimator: 'PG',
            crewLead: "PEPE'S PAINTING LLC",
            split: 1,
            woOverride: 4000,
            woPct: 0.53,
            percent: 0.5,
            notes: '',
            hourly: false,
            excludedFromPayroll: false,
            byDate: { '2026-08-03': 8 },
            total: 8,
            employees: [{ username: 'x@y.com', byDate: { '2026-08-03': 8 }, total: 8, qbHours: null }],
        },
        {
            section: 'balances',
            customer: 'Carlos Customer',
            job: 'AP25264 Int',
            apKey: '25264',
            onCalendar: '-',
            estimator: 'MH',
            crewLead: 'CARLOS RUIZ CRUZ',
            split: 1,
            woOverride: null,
            woPct: 0.38,
            percent: 0.9,
            notes: '',
            hourly: false,
            excludedFromPayroll: false,
            byDate: {},
            total: 0,
            employees: [],
        },
        {
            section: 'balances',
            customer: 'Gray Customer',
            job: 'AP20001 Int',
            apKey: '20001',
            onCalendar: '-',
            estimator: 'MH',
            crewLead: 'CARLOS RUIZ CRUZ',
            split: 1,
            woOverride: null,
            woPct: null,
            percent: 0.5,
            notes: '',
            hourly: false,
            excludedFromPayroll: true, // §11 — gray rows stay off payroll
            byDate: {},
            total: 0,
            employees: [],
        },
    ];
    const slotAPivot = (0, payrollPivot_1.buildPayrollPivotByJob)({
        completed: {
            timesheet: { rows: [], lunchRemoved: 0 },
            customerNameByAp: new Map(),
            jobs: slotAJobs,
            dates: [],
            dateIsos: ['2026-08-03'],
            title: 'NV PAYROLL: 8/14/2026',
            lastWeekPbj: [],
        },
        settings: (0, settings_1.defaultSettings)(),
    });
    check('§6 payroll pivot: WO% comes straight off Slot A\u2019s WO% column', slotAPivot.customers[0].jobs[0].analysis.woPct, 0.53);
    check('§8/§20 payroll pivot: WO Override comes straight off Slot A', slotAPivot.customers[0].jobs[0].analysis.woOverride, 4000);
    check('§9/§11 payroll pivot: bottom-of-pivot rows are carried, gray flag intact', slotAPivot.balances.map((r) => [r.apKey, r.excludedFromPayroll]), [
        ['25264', false],
        ['20001', true],
    ]);
    check('§13 payroll pivot: the Crew Lead cell IS the crew lead', slotAPivot.balances[0].crewLead, 'CARLOS RUIZ CRUZ');
    check('§6 payroll pivot: subcontractor leads still resolve as subs', slotAPivot.customers[0].jobs[0].analysis.sub, true);
    const hourlyPbj = {
        ...ownerPbj,
        customers: ownerPbj.customers.map((c) => ({
            ...c,
            jobs: c.jobs.map((j) => ({
                ...j,
                analysis: { ...j.analysis, hourly: true, percent: 'HOURLY', notes: 'punch list' },
            })),
        })),
    };
    const hourlyPreview = (0, payrollPreview_1.buildPayrollPreview)({
        ...{
            pbj: hourlyPbj,
            enriched: ownerEnriched,
            reportRows: [],
            pl: {
                byJob: new Map(),
                duplicateColumns: [],
                columnHeaders: [''],
                accountRows: [],
                entity: 'NV',
                unmappedCogsRows: [],
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
        },
    });
    const hourlyTs = hourlyPreview.leads[0]?.tsheets?.lines[0] ?? hourlyPreview.painters[0]?.lines[0];
    check('hourly job: T-SHEETS uses the hourly-jobs rate, [Hourly] tag, and Pivot notes', [
        hourlyTs?.hourlyJob,
        hourlyTs?.rate,
        hourlyTs?.jobLabel.includes('[Hourly]') === true,
        hourlyTs?.jobLabel.includes('punch list') === true,
    ], [true, 25, true, true]);
    const carriedRow = (over) => ({
        id: `bal:${over.apKey}`,
        customer: 'Carried Cust',
        job: `AP${over.apKey} Ext`,
        estimator: 'MH',
        woOverride: null,
        woPct: 0.38,
        percent: 1,
        livePercent: 1,
        percentOverride: false,
        excludedFromPayroll: false,
        notes: '',
        colorHex: 'FFFF99',
        ...over,
    });
    const carriedOnlyPreview = (0, payrollPreview_1.buildPayrollPreview)({
        pbj: {
            ...ownerPbj,
            customers: [],
            grandTotal: 0,
            grandByDate: {},
            balances: [
                carriedRow({ apKey: '26000', crewLead: 'CARLOS RUIZ CRUZ', percent: 1 }),
                carriedRow({ apKey: '26001', crewLead: 'CARLOS RUIZ CRUZ', excludedFromPayroll: true }),
                carriedRow({ apKey: '26002', crewLead: 'CARLOS RUIZ CRUZ', percent: 'HOURLY' }),
            ],
        },
        enriched: [],
        reportRows: [],
        pl: {
            byJob: new Map([
                [
                    '26000',
                    {
                        income: 4000,
                        materials: 100,
                        laborWagesSubs: 50,
                        mileage: 0,
                        sundries: 25,
                        travel: 10,
                        plJobHeader: 'AP26000 Ext',
                        customerName: 'Carried Cust',
                        plEntity: 'NV',
                    },
                ],
            ]),
            duplicateColumns: [],
            columnHeaders: [],
            accountRows: [],
            entity: 'NV',
            unmappedCogsRows: [],
        },
        directory: [],
        settings: { ...ownerSettings, painterRoster: [] },
        payDate: new Date(2026, 6, 10),
        edits: {},
        qbHours: {},
        knownEstimators: new Set(['MH']),
        leadCtx: { directory: [], crewGroups: [], subcontractors: [] },
    });
    check('§6 lead with only carried jobs still gets a block', carriedOnlyPreview.leads.map((l) => l.leadName), ['CARLOS RUIZ CRUZ']);
    check('§6 non-gray carried job joins the lead band; gray + HOURLY stay off', carriedOnlyPreview.leads[0].jobs.map((j) => j.apKey), ['26000']);
    check('§8.3 sundries default onto Deductions/Sundries', [carriedOnlyPreview.leads[0].deductions, carriedOnlyPreview.leads[0].deductMemo], [25, 'Sundries (P&L)']);
    check('§8.5 travel metric is on the carried job', carriedOnlyPreview.leads[0].jobs[0].travel, 10);
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
    // --- Round 8 §4.1: the A6A6A6 exclusion fill ---------------------------------
    check('fills: theme 0 + tint −0.35 resolves to A6A6A6', (0, fills_1.resolveThemeTintToRgb)(0, -0.3499862666707358), 'A6A6A6');
    check('fills: theme 0, no tint stays white', (0, fills_1.resolveThemeTintToRgb)(0, 0), 'FFFFFF');
    check('fills: literal ARGB gray', (0, fills_1.isExclusionGray)({ rgb: 'FFA6A6A6', patternType: 'solid' }), true);
    check('fills: literal RGB gray (no alpha)', (0, fills_1.isExclusionGray)({ rgb: 'A6A6A6' }), true);
    check('fills: theme-encoded gray', (0, fills_1.isExclusionGray)({ theme: 0, tint: -0.3499862666707358, patternType: 'solid' }), true);
    check('fills: near-miss inside ±2 tolerance', (0, fills_1.isExclusionGray)({ rgb: 'FFA5A7A6' }), true);
    check('fills: a different gray is not the flag', (0, fills_1.isExclusionGray)({ rgb: 'FFD9D9D9' }), false);
    check('fills: yellow legend fill is not the flag', (0, fills_1.isExclusionGray)({ rgb: 'FFFFFF99' }), false);
    check('fills: no fill', (0, fills_1.isExclusionGray)(null), false);
    // Last-week reader: the painted pair round-trips as excludedFromPayroll.
    const grayWb = {
        sheets: [
            {
                name: 'Pivot by Job',
                rows: [
                    ['Row Labels', 'AP#', 'Estimator', 'Crew Lead', '% of Completion', 'Notes'],
                    ['NV PAYROLL (BALANCES)'],
                    ['Aaron Ward'],
                    ['AP24314 Ext', '24314', 'MH', 'CARLOS', 0.5, ''],
                    ['Regina Conrad'],
                    ['AP24180 Int', '24180', 'MH', 'DANIEL', 0.5, ''],
                ],
                fills: [
                    [],
                    [],
                    [{ theme: 0, tint: -0.3499862666707358, patternType: 'solid' }],
                    [{ theme: 0, tint: -0.3499862666707358, patternType: 'solid' }],
                    [],
                    [],
                ],
            },
        ],
    };
    const grayRows = (0, parsers_1.parseLastWeekPivotByJob)(grayWb);
    check('§4.1 last-week reader flags the theme-gray pair', grayRows.map((r) => [r.apKey, r.excludedFromPayroll]), [
        ['24314', true],
        ['24180', false],
    ]);
    check('§4.3 each carried row keeps its own customer label', grayRows.map((r) => r.customer), ['Aaron Ward', 'Regina Conrad']);
    // --- Round 8 §3/§4: pivot-by-job wiring --------------------------------------
    const pbjSettings = (0, settings_1.defaultSettings)();
    const pbjInput = (over) => ({
        enriched: [],
        reportRows: [],
        settings: pbjSettings,
        leadCtx: { directory: [], crewGroups: [], subcontractors: [] },
        knownEstimators: new Set(['MH', 'RF']),
        payDate: new Date(2026, 6, 10),
        edits: {},
        qbHours: {},
        lastWeekPbj: null,
        balancesState: pivotByJob_1.EMPTY_MANUAL_STATE,
        utState: pivotByJob_1.EMPTY_MANUAL_STATE,
        pl: null,
        nvTotals: null,
        utTotals: null,
        nvSales: null,
        utSales: null,
        ...over,
    });
    const carryRow = (over) => ({
        section: 'balances',
        customer: '',
        job: 'AP00000 Int',
        apKey: '00000',
        estimator: 'MH',
        crewLead: '',
        woOverride: null,
        woPct: null,
        percent: 0.5,
        notes: '',
        excludedFromPayroll: false,
        ...over,
    });
    const tnRow = (apKey, percent, over = {}) => ({
        id: `tn-${apKey}`,
        eventIndex: 0,
        included: true,
        event: `AP${apKey} Int`,
        startDate: null,
        endDate: null,
        calendar: 'AllPro Painters',
        attendees: '',
        location: '',
        description: '',
        apKey,
        percent,
        suggestedSourceLine: null,
        estimator: 'MH',
        lastPaid: null,
        qboJobName: '',
        inQboTime: '',
        jobTotal: '#N/A',
        jobTotalFound: false,
        apptivo: null,
        apptivoEdited: false,
        dollarDiff: null,
        invoiceDate: null,
        dateDiff: null,
        titleInitials: null,
        ...over,
    });
    // §4.2 — RF is Reno, NV: estimator outranks the source calendar.
    check('§4.2 carriedSectionFor: RF from UT → balances', (0, pivotByJob_1.carriedSectionFor)('ut', 'RF'), 'balances');
    check('§4.2 carriedSectionFor: lowercase rf', (0, pivotByJob_1.carriedSectionFor)('ut', ' rf '), 'balances');
    check('§4.2 carriedSectionFor: other estimator keeps UT', (0, pivotByJob_1.carriedSectionFor)('ut', 'MH'), 'ut');
    const rfModel = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        lastWeekPbj: [
            carryRow({ section: 'ut', apKey: '24777', job: 'AP24777 Reno Int', estimator: 'RF' }),
            carryRow({ section: 'ut', apKey: '24888', job: 'AP24888 Provo Int', estimator: 'MH' }),
        ],
    }));
    check('§4.2 RF carried job lands in NV BALANCES', [rfModel.balances.map((r) => r.apKey), rfModel.utPayroll.map((r) => r.apKey)], [['24777'], ['24888']]);
    // §4.3 — each carried job resolves its OWN customer by AP key.
    const salesByCust = [
        { text: 'Aaron Ward', value: null, apKey: '', jobTotal: null },
        { text: '   AP24314 Ext', value: 650, apKey: '24314', jobTotal: 650 },
        { text: '   Total Aaron Ward', value: 650, apKey: '', jobTotal: null },
        { text: 'Regina Conrad', value: null, apKey: '', jobTotal: null },
        { text: '   AP24180 Int', value: 900, apKey: '24180', jobTotal: 900 },
        { text: 'CCSD', value: null, apKey: '', jobTotal: null },
        { text: '   AP24403 Int/Ext', value: 1200, apKey: '24403', jobTotal: 1200 },
    ];
    const custMap = (0, customerByAp_1.customerByApFromSalesSummary)(salesByCust);
    check('§4.3 resolver: AP → owning customer header', [custMap.get('24314'), custMap.get('24180'), custMap.get('24403')], ['Aaron Ward', 'Regina Conrad', 'CCSD']);
    check('§4.3 resolver: Total rows never become customers', custMap.has(''), false);
    const carriedCustomers = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        nvSalesByCust: salesByCust,
        // Last week's sheet leaked ONE name onto every carried row — the bug.
        lastWeekPbj: [
            carryRow({ apKey: '24314', job: 'AP24314 Ext', customer: 'Aaron Ward' }),
            carryRow({ apKey: '24180', job: 'AP24180 Int', customer: 'Aaron Ward' }),
            carryRow({ apKey: '24403', job: 'AP24403 Int/Ext', customer: 'Aaron Ward' }),
        ],
    }));
    check('§4.3 three carried APs → three distinct customers', carriedCustomers.balances.map((r) => r.customer), ['Aaron Ward', 'Regina Conrad', 'CCSD']);
    check('§4.3 distinct customer count', new Set(carriedCustomers.balances.map((r) => r.customer)).size, 3);
    const carriedUnresolved = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        nvSalesByCust: salesByCust,
        lastWeekPbj: [carryRow({ apKey: '29999', job: 'AP29999 Int', customer: 'Raw Carried Name' })],
    }));
    check('§4.3 unresolved → keep raw name + raise an exception', [carriedUnresolved.balances[0].customer, carriedUnresolved.unresolvedCarriedCustomers], ['Raw Carried Name', ['29999']]);
    // §3.3 — carried rows show the CURRENT week's TimeNavi %, not last week's.
    const carriedLive = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        reportRows: [tnRow('24222', 0.9)],
        lastWeekPbj: [
            carryRow({ apKey: '24222', job: 'AP24222 Int', percent: 0.6, notes: '*PAID 60% LAST WK*' }),
        ],
    }));
    check('§3.3 carried % = this week TimeNavi (0.9), not last week (0.6)', [carriedLive.balances[0].percent, carriedLive.balances[0].livePercent, carriedLive.balances[0].percentOverride], [0.9, 0.9, false]);
    const carriedSubbed = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        reportRows: [tnRow('24222', 'SUBBED')],
        lastWeekPbj: [carryRow({ apKey: '24222', job: 'AP24222 Int', percent: 0.6 })],
    }));
    check('§3.3 SUBBED flows through the link verbatim', carriedSubbed.balances[0].percent, 'SUBBED');
    const carriedNoLink = (0, pivotByJob_1.buildPivotByJob)(pbjInput({ lastWeekPbj: [carryRow({ apKey: '24222', job: 'AP24222 Int', percent: 0.6 })] }));
    check('§3.3 no TimeNavi row → carried value stays', carriedNoLink.balances[0].percent, 0.6);
    const carriedOverride = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        reportRows: [tnRow('24222', 0.9)],
        lastWeekPbj: [carryRow({ apKey: '24222', job: 'AP24222 Int', percent: 0.6 })],
        balancesState: { added: [], deleted: [], edits: { 'balances:24222': { percent: 0.75 } } },
    }));
    check('§3.1 manual carried edit wins and is marked', [carriedOverride.balances[0].percent, carriedOverride.balances[0].percentOverride], [0.75, true]);
    check('§4.1 carried gray flag round-trips into the model', (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        lastWeekPbj: [carryRow({ apKey: '24222', job: 'AP24222 Int', excludedFromPayroll: true })],
    })).balances[0].excludedFromPayroll, true);
    // §3.1 — main pivot job rows link to TimeNavi too; Slot A is only a seed.
    const liveJobId = (0, pivotByJob_1.pbjJobId)('Cust', 'AP25000 Ext');
    const liveJob = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        enriched: ownerEnriched,
        reportRows: [tnRow('25000', 0.9)],
        edits: { [liveJobId]: { percentSeed: 0.6 } },
    })).customers[0].jobs[0].analysis;
    check('§3.1 pivot % = this week TimeNavi, not the Slot A seed', [liveJob.percent, liveJob.livePercent, liveJob.percentOverride], [0.9, 0.9, false]);
    const overriddenJob = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        enriched: ownerEnriched,
        reportRows: [tnRow('25000', 0.9)],
        edits: { [liveJobId]: { percentSeed: 0.6, percent: 0.75 } },
    })).customers[0].jobs[0].analysis;
    check('§3.1 manual pivot edit wins over the live link', [overriddenJob.percent, overriddenJob.livePercent, overriddenJob.percentOverride], [0.75, 0.9, true]);
    const seededJob = (0, pivotByJob_1.buildPivotByJob)(pbjInput({ enriched: ownerEnriched, edits: { [liveJobId]: { percentSeed: 0.6 } } })).customers[0].jobs[0].analysis;
    check('§3.1 no TimeNavi row → Slot A seed fills in', [seededJob.percent, seededJob.livePercent], [0.6, null]);
    const hourlyFromSeed = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        enriched: ownerEnriched,
        reportRows: [tnRow('25000', 0.9)],
        edits: { [liveJobId]: { percentSeed: 'HOURLY' } },
    })).customers[0].jobs[0].analysis;
    check('Slot A HOURLY seed classifies the job as hourly even when TimeNavi shows a %', [hourlyFromSeed.percent, hourlyFromSeed.hourly], [0.9, true]);
    const psFromSeed = (0, pivotByJob_1.buildPivotByJob)(pbjInput({
        enriched: ownerEnriched,
        reportRows: [tnRow('25000', 'HOURLY')],
        edits: { [liveJobId]: { percentSeed: 0.4 } },
    })).customers[0].jobs[0].analysis;
    check('Slot A percentage seed classifies the job as profit-share even when TimeNavi says HOURLY', [psFromSeed.percent, psFromSeed.hourly], ['HOURLY', false]);
    // §3.2 — the export writes a real cross-sheet formula.
    check('excelColumnLetter', [(0, exportXlsx_1.excelColumnLetter)(1), (0, exportXlsx_1.excelColumnLetter)(9), (0, exportXlsx_1.excelColumnLetter)(26), (0, exportXlsx_1.excelColumnLetter)(27)], ['A', 'I', 'Z', 'AA']);
    check('§3.2 INDEX/MATCH formula (no XLOOKUP)', (0, exportXlsx_1.timeNaviPercentFormula)({
        sheetName: 'TimeNavi Export',
        apColumn: 'H',
        pctColumn: 'I',
        apCellRef: 'F5',
    }), `IFERROR(INDEX('TimeNavi Export'!I:I, MATCH(F5, 'TimeNavi Export'!H:H, 0)), "")`);
    check(
    // Round 10 — the % symbol leads and "Percentages" is gone.
    '§13 percents filename leads with the % symbol', (0, exportXlsx_1.percentsExportFilename)(new Date(2026, 7, 14)), '% of Completion 08-14-2026.xlsx');
    check('§13 payroll filename has no %', (0, exportXlsx_1.payrollExportFilename)(new Date(2026, 7, 14), new Date(2026, 7, 8), new Date(2026, 7, 14)).includes('%'), false);
    check(
    // Round 10 — the ref map also carries Reimbursements, both note cells and
    // the Miles cell, so the Profitshare Output can reference them live.
    "§10 Input #'s cell refs (first two leads)", (0, exportRound6Sheets_1.inputNumbersCellRefs)(2), [
        {
            hourlyCell: 'C10',
            deductCell: 'C11',
            reimbCell: 'C9',
            reimbNoteCell: 'D9',
            deductNoteCell: 'D11',
            milesCell: 'C12',
        },
        {
            hourlyCell: 'C21',
            deductCell: 'C22',
            reimbCell: 'C20',
            reimbNoteCell: 'D20',
            deductNoteCell: 'D22',
            milesCell: 'C23',
        },
    ]);
    check("§10.1 sheet name apostrophe is doubled", (0, exportRound6Sheets_1.inputNumsSheetFormula)('C11'), { formula: `'Input #''s'!C11` });
    check('§5 RF seed color is C4A484', settings_1.DEFAULT_LEGEND.find((l) => l.estimator === 'RF')?.color, 'C4A484');
    const carlosWages = (0, settings_1.defaultSettings)().wageRates.find((w) => w.name === 'Carlos Ruiz Cruz');
    check('Carlos Ruiz Cruz seeds: PS $18 / hourly $25 (Round 10 — no crew-lead column)', [carlosWages?.rate, carlosWages?.hourlyRate, carlosWages?.crewLeadRate], [18, 25, undefined]);
    check('§11 inactive roster seeds 52 unique painters', new Set(settings_1.DEFAULT_PAINTER_ROSTER.map((p) => p.name)).size, 52);
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
