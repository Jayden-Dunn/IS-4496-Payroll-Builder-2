/**
 * Dev-mode self tests for the pure logic pieces. Runs automatically when the
 * dev server starts (see main.tsx) and logs pass/fail to the console.
 */
import { extractApKey, normalizeApInput, timesheetApKey } from './apkey';
import { dateDiffDays, isDarkColor, parsePercentInput } from './format';
import { parseDate, parseNumberLoose } from './parsers';
import {
  extractLeadParenthetical,
  extractWoOverride,
  extractWoPct,
  jobDisplayName,
  profitSharePctFor,
  resolveCrewLead,
  splitLeadCandidates,
  splitOvertime,
  tokensSubset,
  wageRateFor,
} from './payroll';
import { orderLeadFirst } from './pivot';
import { estimatorFromTitle, extractTitleInitials, suggestPercent } from './timenavi';
import type { DirectoryEmployee } from './types';

interface TestResult {
  name: string;
  pass: boolean;
  detail?: string;
}

export function runSelfTests(): boolean {
  const results: TestResult[] = [];
  const check = (name: string, actual: unknown, expected: unknown): void => {
    const a = JSON.stringify(actual);
    const e = JSON.stringify(expected);
    results.push({ name, pass: a === e, detail: a === e ? undefined : `expected ${e}, got ${a}` });
  };

  // --- apKey extraction -----------------------------------------------------
  check('apKey: plain', extractApKey('AP24314 Ext'), '24314');
  check('apKey: with address', extractApKey('AP24403 5321 Lambrook Dr Int/Ext'), '24403');
  check('apKey: suffixed', extractApKey('AP21848-A5 Somewhere'), '21848-A5');
  check(
    'apKey: from event title',
    extractApKey('♦Int 13k (XC) Regina Conrad Ph1, AP24180-A *PAID* (Daniel)'),
    '24180-A',
  );
  check('apKey: spaced/dashed', extractApKey('AP - 24222 job'), '24222');
  check('apKey: none', extractApKey('Just a random title'), null);
  check('apKey: OFFICE rule', timesheetApKey('OFFICE [Please put AP# in notes]'), 'OFFICE');
  check('apKey: office lowercase', timesheetApKey('  office misc'), 'OFFICE');
  check('apKey: normalize bare number', normalizeApInput('24314'), '24314');
  check('apKey: normalize AP prefix', normalizeApInput('ap24314 ext'), '24314');
  check('apKey: normalize suffixed lowercase', normalizeApInput('21848-a5'), '21848-A5');
  check('apKey: normalize empty', normalizeApInput('  '), null);
  check('apKey: string compare survives suffix', extractApKey('AP21848-A5') === '21848-A5', true);

  // --- % suggestion parser (pay-date driven) ----------------------------------
  const day = (m: number, d: number, y: number): Date => new Date(y, m - 1, d);
  const multiWeek = '7/17/26 100% Complete\n7/10/26 30% COMPLETE -SH\n7/03/26 HOURLY';
  check('%: pay date selects its own line (7/10 → 30%)', suggestPercent(multiWeek, day(7, 10, 2026)).value, 0.3);
  check('%: same file, pay 7/17 → 100%', suggestPercent(multiWeek, day(7, 17, 2026)).value, 1);
  check('%: no exact match → latest ≤ pay date', suggestPercent(multiWeek, day(7, 8, 2026)).value, 'HOURLY');
  check('%: all lines after pay date → blank', suggestPercent(multiWeek, day(6, 1, 2026)).value, null);
  check(
    '%: exact-match line beats a newer line above it',
    suggestPercent('7/10/26 95% COMPLETE -SH\n7/03/26 90% Complete', day(7, 10, 2026)).value,
    0.95,
  );
  check('%: hourly on the selected line', suggestPercent('7/10/26 HOURLY -SH', day(7, 10, 2026)).value, 'HOURLY');
  check('%: ♦/whitespace prefix tolerated', suggestPercent('  ♦7/10/26 97.5% complete', day(7, 10, 2026)).value, 0.975);
  check('%: 4-digit year status line', suggestPercent('7/10/2026 40% Complete', day(7, 10, 2026)).value, 0.4);
  check(
    '%: selected line with no %/HOURLY (e.g. "??") → blank, line kept for tooltip',
    suggestPercent('7/10/26 ??', day(7, 10, 2026)),
    { value: null, line: '7/10/26 ??' },
  );
  check('%: no status lines at all', suggestPercent('Gate code 1234\ncall before arriving', day(7, 10, 2026)).value, null);
  check('%: blank description', suggestPercent('', day(7, 10, 2026)).value, null);
  check('%: no pay date → no suggestion', suggestPercent(multiWeek, null).value, null);
  check(
    '%: source line kept for tooltip',
    suggestPercent('7/10/26 95% COMPLETE -SH', day(7, 10, 2026)).line,
    '7/10/26 95% COMPLETE -SH',
  );

  // --- QuickBooks formula-remnant numbers ------------------------------------
  check('num: formula remnant', parseNumberLoose('=1000'), 1000);
  check('num: subtotal formula', parseNumberLoose('=(B2)+(B3)'), 'FORMULA');
  check('num: =650', parseNumberLoose('=650'), 650);
  check('num: currency string', parseNumberLoose('$1,234.50'), 1234.5);
  check('num: plain', parseNumberLoose(650), 650);
  check('num: junk', parseNumberLoose('n/a'), null);

  // --- date parsing -----------------------------------------------------------
  const d1 = parseDate('6/29/2026');
  check('date: M/D/YYYY', d1 ? [d1.getFullYear(), d1.getMonth() + 1, d1.getDate()] : null, [2026, 6, 29]);
  const d2 = parseDate('01/01/2024');
  check('date: MM/DD/YYYY', d2 ? [d2.getFullYear(), d2.getMonth() + 1, d2.getDate()] : null, [2024, 1, 1]);
  const d3 = parseDate('2026-06-23');
  check('date: ISO', d3 ? [d3.getFullYear(), d3.getMonth() + 1, d3.getDate()] : null, [2026, 6, 23]);
  const d4 = parseDate('6/29/2026 7:00');
  check(
    'date: time component stripped (timezone-safe)',
    d4 ? [d4.getMonth() + 1, d4.getDate(), d4.getHours(), d4.getMinutes()] : null,
    [6, 29, 0, 0],
  );
  const d5 = parseDate(46213); // Excel serial for 2026-07-10 (1900 system)
  check('date: excel serial', d5 ? [d5.getFullYear(), d5.getMonth() + 1, d5.getDate()] : null, [2026, 7, 10]);
  const d6 = parseDate(new Date(2026, 5, 23, 7, 0, 0));
  check('date: Date cell stripped to local midnight', d6 ? [d6.getMonth() + 1, d6.getDate(), d6.getHours()] : null, [6, 23, 0]);
  check('date: garbage', parseDate('not a date'), null);

  // --- date difference ----------------------------------------------------------
  const inv = parseDate('6/25/2026') as Date;
  const start = parseDate('7/1/2026') as Date;
  check('dateDiff: negative ok', dateDiffDays(inv, start), -6);
  check('dateDiff: same day', dateDiffDays(start, start), 0);

  // --- percent input --------------------------------------------------------------
  check('pct input: 95', parsePercentInput('95'), 0.95);
  check('pct input: 95%', parsePercentInput('95%'), 0.95);
  check('pct input: 0.95', parsePercentInput('0.95'), 0.95);
  check('pct input: hourly', parsePercentInput('Hourly'), 'HOURLY');
  check('pct input: dash', parsePercentInput('-'), null);

  // --- title initials ----------------------------------------------------------------
  check('initials: (XC)', extractTitleInitials('♦Int 13k (XC) Regina Conrad, AP24180-A (Daniel)'), 'XC');
  check('initials: lowercase name ignored', extractTitleInitials('Job for (Daniel)'), null);

  // --- estimator fallback from the event title ----------------------------------------
  const known = new Set(['JS', 'MH', 'PG', 'TE', 'XC', 'ND', 'RF']);
  check(
    'title est: first known parenthesized token',
    estimatorFromTitle('♦Int 13k (XC) Regina Conrad Ph1, AP24180-A *PAID* (Daniel)', known),
    'XC',
  );
  check(
    'title est: crew names with / or >3 letters never match',
    estimatorFromTitle('♦[Southpoint] 3.5k (MH) Framing/DW/FRP Install, AP24450 (Ted/Rafael)', known),
    'MH',
  );
  check('title est: lowercase tag uppercased', estimatorFromTitle('Job walk (js) tomorrow', known), 'JS');
  check('title est: known-set guard rejects (Ted)', estimatorFromTitle('Fence repair (Ted)', known), null);
  check('title est: skips unknown token to find a later known one', estimatorFromTitle('(Ted) plus (PG)', known), 'PG');
  check('title est: no parenthesized tokens', estimatorFromTitle('Paint shop maintenance', known), null);

  // --- crew-lead-first ordering ----------------------------------------------------------
  const names = (emps: { fname: string; lname: string }[]): string[] => emps.map((e) => e.fname);
  const carlosCrew = [
    { fname: 'Ana', lname: 'Bedoya Melgar' },
    { fname: 'Carlos', lname: 'Ruiz Cruz' },
  ];
  check('lead: Carlos first in his crew', names(orderLeadFirst("Carlos Ruiz Cruz's Crew", carlosCrew)), ['Carlos', 'Ana']);
  check(
    'lead: extra surname matches (Daniel Bonilla → Daniel Gonzales Bonilla)',
    names(orderLeadFirst("Daniel Bonilla's Crew", [
      { fname: 'Armando', lname: 'Lopez' },
      { fname: 'Daniel', lname: 'Gonzales Bonilla' },
    ])),
    ['Daniel', 'Armando'],
  );
  check(
    'lead: token guard (Miguel Martin ≠ Samuel Martin)',
    names(orderLeadFirst('Miguel Martin\u2019s Crew', [
      { fname: 'Miguel', lname: 'Martin' },
      { fname: 'Samuel', lname: 'Martin' },
    ])),
    ['Miguel', 'Samuel'],
  );
  check(
    'lead: hyphenated surname tokenized (Leonardo Martinez → Martinez-Rincon)',
    names(orderLeadFirst("Leonardo Martinez's Crew", [
      { fname: 'Juan', lname: 'Perez' },
      { fname: 'Leonardo', lname: 'Martinez-Rincon' },
    ])),
    ['Leonardo', 'Juan'],
  );
  check(
    'lead: tie → fewest extra tokens wins',
    orderLeadFirst("Tomas Huerta's Crew", [
      { fname: 'Tomas', lname: 'Huerta Bonilla Garcia' },
      { fname: 'Tomas', lname: 'Huerta Bonilla' },
    ])[0].lname,
    'Huerta Bonilla',
  );
  check(
    'lead: no "\'s Crew" suffix → untouched',
    names(orderLeadFirst('Handyman', carlosCrew)),
    ['Ana', 'Carlos'],
  );
  check(
    'lead: no matching employee → untouched',
    names(orderLeadFirst("Rene Salazar's Crew", carlosCrew)),
    ['Ana', 'Carlos'],
  );

  // --- palette text contrast -----------------------------------------------------------
  check('contrast: ND dark blue → white text', isDarkColor('4472C4'), true);
  check('contrast: MH light yellow → black text', isDarkColor('FFFF99'), false);
  check('contrast: RF dusty rose → black text', isDarkColor('DA9694'), false);

  // --- Round 3 §3.1: job display name -----------------------------------------------------
  check('jobName: AP token stripped', jobDisplayName('Aaron Ward', 'AP24314 Ext'), 'Aaron Ward Ext');
  check(
    'jobName: suffixed AP + inner text',
    jobDisplayName('Tri-State', 'AP24097-A Sams Club Int'),
    'Tri-State Sams Club Int',
  );
  check('jobName: single-spaced', jobDisplayName(' CCSD ', 'AP24403  5321 Lambrook Dr Int/Ext'), 'CCSD 5321 Lambrook Dr Int/Ext');

  // --- Round 3 §3.2: crew-lead extraction ---------------------------------------------------
  check(
    'lead extract: last parenthetical, not the estimator',
    extractLeadParenthetical('♦Int 13k (XC) Regina Conrad Ph1, AP24180-A *PAID* (Daniel)', known),
    'Daniel',
  );
  check('lead extract: estimator-only title → null', extractLeadParenthetical('AP24403 Lambrook (PG)', known), null);
  check('lead extract: no parentheses → null', extractLeadParenthetical('AP24477 Plain title', known), null);
  check('lead extract: business name kept', extractLeadParenthetical("Ext 8k (MH) Smith, AP24499 (Pepe's Painting)", known), "Pepe's Painting");
  check('lead split: slash', splitLeadCandidates('Ted/Rafael'), ['Ted', 'Rafael']);
  check('lead split: commas', splitLeadCandidates('Ted, Antonio, Tomas'), ['Ted', 'Antonio', 'Tomas']);
  check('lead split: hrly_ prefix stripped', splitLeadCandidates('hrly_Damian/Juan Salazar'), ['Damian', 'Juan Salazar']);
  check('lead split: asterisks stripped', splitLeadCandidates('*Carlos*'), ['Carlos']);

  const dir: DirectoryEmployee[] = [
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
  check('lead resolve: Carlos → CARLOS RUIZ CRUZ', resolveCrewLead('Carlos', ctx).display, 'CARLOS RUIZ CRUZ');
  check(
    'lead resolve: crew-lead Miguel Martin preferred over other Miguels',
    resolveCrewLead('Miguel', ctx).display,
    'MIGUEL MARTIN',
  );
  check('lead resolve: sub via token match', [resolveCrewLead("Pepe's Painting", ctx).display, resolveCrewLead("Pepe's Painting", ctx).sub], ["PEPE'S PAINTING LLC", true]);
  const welder = resolveCrewLead('Welder', ctx);
  check('lead resolve: no match anywhere → flag', [welder.display, welder.flags[0]?.kind], ['', 'noMatch']);
  const split = resolveCrewLead('Ted/Rafael', ctx);
  check('lead resolve: split leads joined + split:2', [split.display, split.split], ['TED FRANK/RAFAEL AGUILAR', 2]);
  const calixto = resolveCrewLead('Calixto', ctx);
  check('lead resolve: inactive-only match used but flagged', [calixto.display, calixto.flags.some((f) => f.kind === 'inactiveMatch')], ['CALIXTO EFRAIN CLIMACO', true]);
  check('lead resolve: empty raw → empty, no flags', resolveCrewLead(null, ctx), { display: '', names: [], split: 1, sub: false, raw: '', flags: [] });
  check('tokens: middle initial optional', tokensSubset('Antonio Zuniga', 'Antonio J Zuniga'), true);

  // --- Round 3 §3.4: overtime split ---------------------------------------------------------
  const day2 = (m: number, d: number): Date => new Date(2026, m - 1, d);
  const ot = splitOvertime([
    { jobId: 'a', hours: 25, lastDate: day2(6, 30), lastRowIndex: 1 },
    { jobId: 'b', hours: 15.67, lastDate: day2(7, 3), lastRowIndex: 5 },
  ]);
  check('OT: overage on the LAST job', [ot.otJobId, ot.otHours], ['b', 0.67]);
  check('OT: regular reduced so reg+OT = actual', ot.regularByJob.get('b'), 15);
  check('OT: other jobs untouched', ot.regularByJob.get('a'), 25);
  const noOt = splitOvertime([{ jobId: 'a', hours: 40, lastDate: day2(7, 3), lastRowIndex: 0 }]);
  check('OT: exactly 40 → none', [noOt.otJobId, noOt.otHours], [null, 0]);
  const tie = splitOvertime([
    { jobId: 'x', hours: 30, lastDate: day2(7, 3), lastRowIndex: 2 },
    { jobId: 'y', hours: 12, lastDate: day2(7, 3), lastRowIndex: 7 },
  ]);
  check('OT: same-day tie → last by file order', tie.otJobId, 'y');

  // --- Round 3 §4.3: WO pre-fills -------------------------------------------------------------
  const jodi = extractWoOverride('gate 1234\n****PER JODI $4,715.00 is the total of work order he completed****');
  check('WO$: $ before "work order" caught', [jodi?.amount, jodi?.line?.includes('PER JODI')], [4715, true]);
  check('WO$: W.O. before $', extractWoOverride('W.O. total came to $2,300')?.amount, 2300);
  check('WO$: plain money line ignored', extractWoOverride('Invoice total $9,000 sent'), null);
  check('WO%: "@ 40% PS" line', extractWoPct('7/03/26 10% COMPLETE @ 40% PS')?.pct, 0.4);
  check('WO%: absent → null', extractWoPct('7/03/26 10% COMPLETE'), null);

  // --- Round 3 §7: settings lookups -----------------------------------------------------------
  const psMap = [
    { lead: 'Antonio Zuniga', pct: 0.36 },
    { lead: "Pepe's Painting LLC", pct: 0.5 },
  ];
  check('PS%: token match (ALL-CAPS lead)', profitSharePctFor('ANTONIO ZUNIGA', psMap, 0.38), 0.36);
  check('PS%: unlisted crew → default', profitSharePctFor('NEW GUY', psMap, 0.38), 0.38);
  check("PS%: sub matched (PEPE'S PAINTING LLC)", profitSharePctFor("PEPE'S PAINTING LLC", psMap, 0.38), 0.5);
  const rates = [{ name: 'Rafael Aguilar', rate: 18, hourlyRate: 25 }];
  check('wage: piece-rate job → default rate', wageRateFor('RAFAEL AGUILAR', rates, 18, false), { rate: 18, defaulted: false });
  check('wage: hourly job → hourly rate', wageRateFor('RAFAEL AGUILAR', rates, 18, true), { rate: 25, defaulted: false });
  check('wage: unknown employee → fallback + flag', wageRateFor('Nobody Known', rates, 18, false), { rate: 18, defaulted: true });

  const failed = results.filter((r) => !r.pass);
  const label = `[% of Completion Builder self-test] ${results.length - failed.length}/${results.length} passed`;
  if (failed.length === 0) {
    console.log(`%c${label}`, 'color: green; font-weight: bold');
  } else {
    console.warn(`%c${label}`, 'color: red; font-weight: bold');
    for (const f of failed) console.warn(`  FAIL ${f.name}: ${f.detail ?? ''}`);
  }
  return failed.length === 0;
}
