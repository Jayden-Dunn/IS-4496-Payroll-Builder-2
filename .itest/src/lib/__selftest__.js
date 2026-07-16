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
