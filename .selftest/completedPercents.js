"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.findPbjHeaderRow = findPbjHeaderRow;
exports.parseCompletedPercentsWorkbook = parseCompletedPercentsWorkbook;
/**
 * Round 6 §2.2 Slot A — parse a completed %'s workbook for Payroll mode.
 * Reads `Pivot by Job` + `timesheets` by sheet name; never throws for missing sheets.
 */
const apkey_1 = require("./apkey");
const format_1 = require("./format");
const parsers_1 = require("./parsers");
// ---------------------------------------------------------------------------
// Sheet lookup
// ---------------------------------------------------------------------------
function findSheet(wbd, name) {
    const want = name.trim().toLowerCase();
    return wbd.sheets.find((s) => s.name.trim().toLowerCase() === want);
}
function missingSheetError(name) {
    return {
        error: `This file has no '${name}' sheet — export it from the %'s process first`,
    };
}
function asString(v) {
    if (v == null)
        return '';
    if (v instanceof Date)
        return v.toISOString();
    return String(v).trim();
}
function isEmptyRow(row) {
    return row.every((c) => c == null || (typeof c === 'string' && c.trim() === ''));
}
function cellAt(row, col) {
    if (col === undefined)
        return null;
    return row[col] ?? null;
}
// ---------------------------------------------------------------------------
// findPbjHeaderRow — exported for self-tests
// ---------------------------------------------------------------------------
const PBJ_MARKERS = [
    'ap#',
    'on calendar?',
    'estimator',
    'crew lead',
    'wo override',
    'wo%',
    '% of completion',
    'notes',
];
/**
 * Scan the first 6 rows for the Pivot by Job analysis header row.
 * Accepts the app's 2-row header (labels on row 1) and the legacy human
 * 4-row header (labels lower). Returns -1 when not found.
 */
function findPbjHeaderRow(rows) {
    let best = -1;
    let bestScore = 0;
    for (let i = 0; i < Math.min(6, rows.length); i++) {
        const row = rows[i];
        if (!row)
            continue;
        const keys = new Set();
        for (const c of row) {
            if (typeof c !== 'string')
                continue;
            const k = c.trim().toLowerCase();
            if (k)
                keys.add(k);
        }
        let score = 0;
        for (const m of PBJ_MARKERS) {
            if (keys.has(m)) {
                score++;
                continue;
            }
            if (m === 'on calendar?' && (keys.has('on calendar') || keys.has('on cal?')))
                score++;
            else if (m === 'wo%' && (keys.has('wo %') || keys.has('wo pct')))
                score++;
            else if (m === 'wo override' && keys.has('wo amt'))
                score++;
        }
        // Require AP# plus at least two other analysis labels.
        if (keys.has('ap#') && score >= 3 && score > bestScore) {
            bestScore = score;
            best = i;
        }
    }
    return best;
}
// ---------------------------------------------------------------------------
// timesheets sheet
// ---------------------------------------------------------------------------
/**
 * Truncate to columns A–X and cut after the last row that has a username
 * (helper columns / trailing totals ignored).
 */
function prepareTimesheetRows(rows) {
    const sliced = rows.map((r) => {
        const out = [];
        for (let i = 0; i < 24; i++)
            out.push(r[i] ?? null);
        return out;
    });
    let userCol = 0;
    let dataStart = 0;
    for (let i = 0; i < Math.min(10, sliced.length); i++) {
        const row = sliced[i];
        if (!row)
            continue;
        for (let c = 0; c < row.length; c++) {
            if (asString(row[c]).toLowerCase() === 'username') {
                userCol = c;
                dataStart = i + 1;
                break;
            }
        }
        if (dataStart > 0)
            break;
    }
    let last = dataStart - 1;
    for (let i = dataStart; i < sliced.length; i++) {
        const u = sliced[i]?.[userCol];
        if (u != null && String(u).trim() !== '')
            last = i;
    }
    if (last < dataStart)
        return sliced;
    return sliced.slice(0, last + 1);
}
function parseTimesheetsSheet(sheet) {
    const rows = prepareTimesheetRows(sheet.rows);
    try {
        // Round-trip exports are NV-tagged by default; an optional entity column is rare.
        return (0, parsers_1.parseTimesheet)({ sheets: [{ name: sheet.name, rows }] }, 'NV');
    }
    catch (e) {
        return { error: e instanceof Error ? e.message : String(e) };
    }
}
function buildPbjLayout(rows, headerRow) {
    const hdr = rows[headerRow];
    if (!hdr)
        return null;
    const cols = new Map();
    for (let c = 0; c < hdr.length; c++) {
        const v = hdr[c];
        if (typeof v !== 'string')
            continue;
        const key = v.trim().toLowerCase();
        if (key && !cols.has(key))
            cols.set(key, c);
    }
    const apCol = cols.get('ap#');
    if (apCol === undefined)
        return null;
    const totalCol = cols.get('total hrs') ?? cols.get('total') ?? cols.get('sum of hours');
    const qbCol = cols.get('qb hrs') ?? cols.get('qb hours');
    // Date columns: real dates in the header, left of AP# (and of Total hrs when present).
    const dateLimit = totalCol !== undefined ? Math.min(totalCol, apCol) : apCol;
    const dateCols = [];
    for (let c = 1; c < dateLimit; c++) {
        const d = (0, parsers_1.parseDate)(hdr[c] ?? null);
        if (d)
            dateCols.push({ col: c, date: d, iso: (0, format_1.dateToIso)(d) });
    }
    // Title from column A of the header row (e.g. "NV PAYROLL: 7/10/2026").
    let title = asString(hdr[0]);
    if (!title) {
        for (let i = 0; i < headerRow; i++) {
            const t = asString(rows[i]?.[0] ?? null);
            if (t) {
                title = t;
                break;
            }
        }
    }
    // QB Hrs often has a blank header in the app's 2-row export (empty cell between
    // Total hrs and AP#). Infer it when missing.
    let inferredQb = qbCol;
    if (inferredQb === undefined && totalCol !== undefined && apCol === totalCol + 2) {
        inferredQb = totalCol + 1;
    }
    return {
        headerRow,
        apCol,
        onCalCol: cols.get('on calendar?') ?? cols.get('on calendar') ?? cols.get('on cal?'),
        estCol: cols.get('estimator'),
        leadCol: cols.get('crew lead'),
        woCol: cols.get('wo override') ?? cols.get('wo amt'),
        woPctCol: cols.get('wo%') ?? cols.get('wo %') ?? cols.get('wo pct'),
        pctCol: cols.get('% of completion'),
        notesCol: cols.get('notes'),
        totalCol,
        qbCol: inferredQb,
        dateCols,
        title,
    };
}
function detectSection(label) {
    const t = label.trim();
    if (!t)
        return null;
    if (/balances/i.test(t) && (/nv\s*payroll/i.test(t) || /^balances\b/i.test(t))) {
        return 'balances';
    }
    if (/^balances\b/i.test(t))
        return 'balances';
    if (/^ut\s*payroll/i.test(t))
        return 'ut';
    if (/^nv\s*payroll/i.test(t))
        return 'nv';
    return null;
}
function isSkipLabel(a) {
    return /^(row labels|grand total|sum of hours)$/i.test(a.trim());
}
/** % / WO% cell: fractions, NN% text, HOURLY, `-` → null. */
function parsePctCell(v) {
    if (v == null || v === '')
        return null;
    if (typeof v === 'number')
        return Number.isFinite(v) ? (v > 1 ? v / 100 : v) : null;
    if (typeof v === 'boolean')
        return null;
    if (v instanceof Date)
        return null;
    return (0, format_1.parsePercentInput)(String(v));
}
function parseOnCalendar(v) {
    if (v == null || v === '')
        return '';
    if (typeof v === 'number' && Number.isFinite(v)) {
        // Bare TimeNavi row index (e.g. 19).
        return String(Math.trunc(v));
    }
    return asString(v);
}
function crewLeadSplit(lead) {
    if (!lead)
        return 1;
    const parts = lead.split('/').map((s) => s.trim()).filter(Boolean);
    return Math.max(1, parts.length);
}
function hoursFromRow(row, layout) {
    const byDate = {};
    let sum = 0;
    for (const d of layout.dateCols) {
        const n = (0, parsers_1.parseNumberLoose)(cellAt(row, d.col));
        if (typeof n === 'number' && n !== 0) {
            byDate[d.iso] = (0, format_1.round2)(n);
            sum += n;
        }
    }
    let total = sum;
    if (layout.totalCol !== undefined) {
        const t = (0, parsers_1.parseNumberLoose)(cellAt(row, layout.totalCol));
        if (typeof t === 'number')
            total = t;
    }
    return { byDate, total: (0, format_1.round2)(total) };
}
function parsePivotByJobSheet(sheet) {
    const headerRow = findPbjHeaderRow(sheet.rows);
    if (headerRow < 0) {
        return {
            error: "Could not find the Pivot by Job header row (expected AP#, Estimator, Crew Lead, % of Completion, … in the first 6 rows).",
        };
    }
    const layout = buildPbjLayout(sheet.rows, headerRow);
    if (!layout) {
        return { error: "Pivot by Job header row is missing an 'AP#' column." };
    }
    const jobs = [];
    let section = 'nv';
    let customer = '';
    let collectingEmployees = false;
    let currentJob = null;
    // Skip the header block: app 2-row → +1 weekday row; legacy may have extra
    // label rows above (already skipped by starting after headerRow). Also skip
    // an immediate next row of weekday names / blanks under date cols.
    let start = headerRow + 1;
    if (start < sheet.rows.length) {
        const next = sheet.rows[start];
        const looksLikeWeekdays = next &&
            layout.dateCols.length > 0 &&
            layout.dateCols.every(({ col }) => {
                const v = next[col];
                if (v == null || v === '')
                    return true;
                if (typeof v === 'string' && /^[A-Za-z]{2,9}$/.test(v.trim()))
                    return true;
                return false;
            }) &&
            !asString(cellAt(next, layout.apCol));
        if (looksLikeWeekdays)
            start++;
    }
    for (let i = start; i < sheet.rows.length; i++) {
        const row = sheet.rows[i];
        if (!row || isEmptyRow(row)) {
            collectingEmployees = false;
            currentJob = null;
            continue;
        }
        const a = asString(row[0]);
        const sec = detectSection(a);
        if (sec) {
            section = sec;
            customer = '';
            collectingEmployees = false;
            currentJob = null;
            // Section header rows may repeat the analysis headers — skip this row.
            continue;
        }
        // Repeated header rows inside the sheet (BALANCES / UT black headers).
        const apRaw = cellAt(row, layout.apCol);
        if (typeof apRaw === 'string' && apRaw.trim().toLowerCase() === 'ap#') {
            collectingEmployees = false;
            currentJob = null;
            continue;
        }
        if (isSkipLabel(a)) {
            collectingEmployees = false;
            currentJob = null;
            continue;
        }
        const apKey = apRaw == null || apRaw === '' ? null : (0, apkey_1.normalizeApInput)(String(apRaw));
        if (apKey) {
            const woRaw = cellAt(row, layout.woCol);
            const woPctRaw = cellAt(row, layout.woPctCol);
            const pctRaw = cellAt(row, layout.pctCol);
            const woDash = typeof woRaw === 'string' && woRaw.trim() === '-';
            const woPctDash = typeof woPctRaw === 'string' && woPctRaw.trim() === '-';
            const percent = parsePctCell(pctRaw);
            const woNum = (0, parsers_1.parseNumberLoose)(woRaw);
            const woPct = parsePctCell(woPctRaw);
            const hourly = percent === 'HOURLY' || woDash || woPctDash;
            const lead = asString(cellAt(row, layout.leadCol));
            const hours = hoursFromRow(row, layout);
            const job = {
                section,
                customer,
                job: a,
                apKey,
                onCalendar: parseOnCalendar(cellAt(row, layout.onCalCol)),
                estimator: asString(cellAt(row, layout.estCol)),
                crewLead: lead,
                split: crewLeadSplit(lead),
                woOverride: hourly || typeof woNum !== 'number' ? null : woNum,
                woPct: hourly || typeof woPct !== 'number' ? null : woPct,
                percent,
                notes: asString(cellAt(row, layout.notesCol)),
                hourly,
                byDate: hours.byDate,
                total: hours.total,
                employees: [],
            };
            jobs.push(job);
            currentJob = job;
            collectingEmployees = true;
            continue;
        }
        // Non-AP row: employee under current job, or a new customer label.
        if (collectingEmployees && currentJob && a) {
            const hours = hoursFromRow(row, layout);
            const qb = layout.qbCol !== undefined ? (0, parsers_1.parseNumberLoose)(cellAt(row, layout.qbCol)) : null;
            currentJob.employees.push({
                username: a,
                byDate: hours.byDate,
                total: hours.total,
                qbHours: typeof qb === 'number' ? qb : null,
            });
            continue;
        }
        if (a && !isSkipLabel(a)) {
            customer = a;
            collectingEmployees = false;
            currentJob = null;
        }
    }
    return {
        jobs,
        dates: layout.dateCols.map((d) => d.date),
        dateIsos: layout.dateCols.map((d) => d.iso),
        title: layout.title,
    };
}
function toLastWeekPbj(jobs) {
    return jobs.map((j) => ({
        section: j.section === 'nv' ? 'pivot' : j.section,
        customer: j.customer,
        job: j.job,
        apKey: j.apKey,
        estimator: j.estimator,
        crewLead: j.crewLead,
        woOverride: j.woOverride,
        woPct: j.woPct,
        percent: j.percent,
        notes: j.notes,
    }));
}
// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------
/**
 * Parse a completed %'s workbook (Slot A). Returns `{ error }` for missing
 * sheets or unusable structure — never throws for a missing sheet name.
 */
function parseCompletedPercentsWorkbook(wbd) {
    const tsSheet = findSheet(wbd, 'timesheets');
    if (!tsSheet)
        return missingSheetError('timesheets');
    const pbjSheet = findSheet(wbd, 'Pivot by Job');
    if (!pbjSheet)
        return missingSheetError('Pivot by Job');
    const timesheet = parseTimesheetsSheet(tsSheet);
    if ('error' in timesheet)
        return timesheet;
    const pbj = parsePivotByJobSheet(pbjSheet);
    if ('error' in pbj)
        return pbj;
    return {
        timesheet,
        jobs: pbj.jobs,
        dates: pbj.dates,
        dateIsos: pbj.dateIsos,
        title: pbj.title,
        lastWeekPbj: toLastWeekPbj(pbj.jobs),
    };
}
