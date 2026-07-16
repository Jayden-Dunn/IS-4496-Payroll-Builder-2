"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.normalizeApInput = void 0;
exports.suggestPercent = suggestPercent;
exports.extractTitleInitials = extractTitleInitials;
exports.reportRowId = reportRowId;
exports.buildReportRows = buildReportRows;
exports.sortReportRows = sortReportRows;
const apkey_1 = require("./apkey");
Object.defineProperty(exports, "normalizeApInput", { enumerable: true, get: function () { return apkey_1.normalizeApInput; } });
const format_1 = require("./format");
const joins_1 = require("./joins");
// ---------------------------------------------------------------------------
// Description → % pre-fill parser (pay-date driven)
// ---------------------------------------------------------------------------
const STATUS_DATE_RE = /^\s*[♦*]?\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/;
const PCT_RE = /(\d{1,3}(?:\.\d+)?)\s*%/;
function statusLineDate(month, day, year) {
    let y = Number(year);
    if (y < 100)
        y += 2000;
    return new Date(y, Number(month) - 1, Number(day));
}
function extractStatusLines(description) {
    const out = [];
    for (const raw of description.split(/\r?\n/)) {
        const line = raw.trim();
        if (!line)
            continue;
        const m = STATUS_DATE_RE.exec(line);
        if (!m)
            continue;
        out.push({ date: statusLineDate(m[1], m[2], m[3]), line });
    }
    return out;
}
/**
 * Select the status line for the pay date: exact match first, else the latest
 * line dated on or before the pay date. Parse % or HOURLY from that line only.
 */
function suggestPercent(description, payDate) {
    const statusLines = extractStatusLines(description);
    if (statusLines.length === 0 || !payDate)
        return { value: null, line: null };
    const payMs = (0, format_1.dateOnly)(payDate).getTime();
    let selected = statusLines.find((s) => (0, format_1.dateOnly)(s.date).getTime() === payMs) ?? null;
    if (!selected) {
        const eligible = statusLines.filter((s) => (0, format_1.dateOnly)(s.date).getTime() <= payMs);
        if (eligible.length === 0)
            return { value: null, line: null };
        selected = eligible.reduce((best, cur) => (0, format_1.dateOnly)(cur.date).getTime() > (0, format_1.dateOnly)(best.date).getTime() ? cur : best);
    }
    const text = selected.line;
    const pct = PCT_RE.exec(text);
    if (pct) {
        const n = Number(pct[1]);
        if (n >= 0 && n <= 100)
            return { value: n / 100, line: selected.line };
    }
    if (/hourly/i.test(text))
        return { value: 'HOURLY', line: selected.line };
    return { value: null, line: selected.line };
}
/** First "(XX)" group of 2–3 uppercase letters in an event title. */
function extractTitleInitials(title) {
    const m = /\(([A-Z]{2,3})\)/.exec(title);
    return m ? m[1] : null;
}
function reportRowId(ev) {
    return `ev${ev.index}:${ev.apKey ?? ''}`;
}
function buildReportRows(events, edits, deps, payDate) {
    const nv = (0, joins_1.buildSalesIndex)(deps.nvSales);
    const ut = (0, joins_1.buildSalesIndex)(deps.utSales);
    return events.map((ev) => {
        const id = reportRowId(ev);
        const edit = edits[id] ?? {};
        const apKey = edit.apKey !== undefined ? edit.apKey : ev.apKey;
        const included = edit.included !== undefined ? edit.included : apKey !== null;
        const suggestion = suggestPercent(ev.description, payDate);
        const percent = edit.percent !== undefined ? edit.percent : suggestion.value;
        const nvMatch = apKey ? nv.byApKey.get(apKey) : undefined;
        const utMatch = apKey ? ut.byApKey.get(apKey) : undefined;
        const estimator = edit.estimator !== undefined
            ? edit.estimator
            : utMatch?.salesRep.trim() || nvMatch?.salesRep.trim() || '';
        // K: only from last week's workbook — never from Description.
        const lastPaid = apKey && deps.lastWeekMap ? (deps.lastWeekMap.get(apKey) ?? null) : null;
        const qboJobName = nvMatch?.name2 ?? utMatch?.name2 ?? '';
        const invoiceDate = nvMatch ? nvMatch.date : utMatch ? utMatch.date : null;
        const inQboTime = (apKey && deps.timesheetJobcodeByApKey.get(apKey)) || 'NotClockedIn';
        const nvTotal = apKey ? deps.nvTotals.get(apKey) : undefined;
        const utTotal = apKey ? deps.utTotals.get(apKey) : undefined;
        const jobTotalFound = nvTotal !== undefined || utTotal !== undefined;
        const matchedTotal = nvTotal ?? utTotal ?? 0;
        const jobTotal = jobTotalFound ? matchedTotal : '#N/A';
        const apptivoEdited = edit.apptivo !== undefined;
        const apptivo = apptivoEdited
            ? (edit.apptivo ?? null)
            : jobTotalFound
                ? matchedTotal
                : null;
        const dollarDiff = jobTotalFound && apptivo !== null ? matchedTotal - apptivo : null;
        const dateDiff = invoiceDate && ev.startDate ? (0, format_1.dateDiffDays)(invoiceDate, ev.startDate) : null;
        return {
            id,
            eventIndex: ev.index,
            included,
            event: ev.event,
            startDate: ev.startDate,
            endDate: ev.endDate,
            calendar: ev.calendar,
            attendees: ev.attendees,
            location: ev.location,
            description: ev.description,
            apKey,
            percent,
            suggestedSourceLine: suggestion.line,
            estimator,
            lastPaid,
            qboJobName,
            inQboTime,
            jobTotal,
            jobTotalFound,
            apptivo,
            apptivoEdited,
            dollarDiff,
            invoiceDate,
            dateDiff,
            titleInitials: extractTitleInitials(ev.event),
        };
    });
}
/** Sort by estimator (legend order from Settings), then Start Date; stable. */
function sortReportRows(rows, estimatorOrder) {
    const idx = new Map(estimatorOrder.map((e, i) => [e, i]));
    return rows
        .map((r, i) => ({ r, i }))
        .sort((a, b) => {
        const ea = idx.get(a.r.estimator) ?? estimatorOrder.length;
        const eb = idx.get(b.r.estimator) ?? estimatorOrder.length;
        if (ea !== eb)
            return ea - eb;
        const ta = a.r.startDate ? a.r.startDate.getTime() : Number.POSITIVE_INFINITY;
        const tb = b.r.startDate ? b.r.startDate.getTime() : Number.POSITIVE_INFINITY;
        if (ta !== tb)
            return ta - tb;
        return a.i - b.i;
    })
        .map((x) => x.r);
}
