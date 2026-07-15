/**
 * Dev-mode self tests for the pure logic pieces. Runs automatically when the
 * dev server starts (see main.tsx) and logs pass/fail to the console.
 */
import { extractApKey, normalizeApInput, timesheetApKey } from './apkey';
import { dateDiffDays, parsePercentInput } from './format';
import { parseDate, parseNumberLoose } from './parsers';
import { extractTitleInitials, suggestPercent } from './timenavi';

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

  // --- % suggestion parser --------------------------------------------------
  check(
    '%: newest-first wins',
    suggestPercent('7/10/26 95% COMPLETE -SH\n7/03/26 90% Complete\nother notes').value,
    0.95,
  );
  check('%: hourly', suggestPercent('7/17/26 HOURLY -SH\n7/03/26 90% Complete').value, 'HOURLY');
  check('%: none', suggestPercent('Gate code 1234\ncall before arriving').value, null);
  check('%: 100 complete', suggestPercent('7/8/26 100% Complete -SH').value, 1);
  check('%: blank description', suggestPercent('').value, null);
  check(
    '%: source line kept for tooltip',
    suggestPercent('7/10/26 95% COMPLETE -SH').line,
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
  check('date: with time', d4 ? [d4.getMonth() + 1, d4.getDate(), d4.getHours()] : null, [6, 29, 7]);
  const d5 = parseDate(46213); // Excel serial for 2026-07-10 (1900 system)
  check('date: excel serial', d5 ? [d5.getFullYear(), d5.getMonth() + 1, d5.getDate()] : null, [2026, 7, 10]);
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
