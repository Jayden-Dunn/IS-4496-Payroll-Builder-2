"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.runSelfTests = runSelfTests;
/**
 * Dev-mode self tests for the pure logic pieces. Runs automatically when the
 * dev server starts (see main.tsx) and logs pass/fail to the console.
 */
const apkey_1 = require("./apkey");
const format_1 = require("./format");
const parsers_1 = require("./parsers");
const payroll_1 = require("./payroll");
const pivot_1 = require("./pivot");
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
