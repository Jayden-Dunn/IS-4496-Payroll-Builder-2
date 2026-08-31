"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.percentTextValue = exports.percentText = exports.percentIsValidState = exports.percentExcelValue = exports.matchesSubbedOut = exports.matchesSetRate = exports.isPercentText = exports.isPercentState = exports.isPercentNumber = exports.normalizeApInput = void 0;
exports.suggestPercent = suggestPercent;
exports.extractTitleInitials = extractTitleInitials;
exports.estimatorFromTitle = estimatorFromTitle;
exports.reportRowId = reportRowId;
exports.buildReportRows = buildReportRows;
exports.sortReportRows = sortReportRows;
exports.percentIsFlagged = percentIsFlagged;
exports.percentIsRedFlagged = percentIsRedFlagged;
exports.duplicateInQboValues = duplicateInQboValues;
exports.dollarDiffIsFlagged = dollarDiffIsFlagged;
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
 * line dated on or before the pay date. Parse % / HOURLY / SUBBED / SET RATE
 * from that line only (Round 8 §2.4).
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
    // §2.4 — the slot where the % normally sits can instead read "subbed out"
    // or "set rate"; both are valid states, not missing data.
    if ((0, format_1.matchesSubbedOut)(text))
        return { value: 'SUBBED', line: selected.line };
    if ((0, format_1.matchesSetRate)(text))
        return { value: 'SET RATE', line: selected.line };
    return { value: null, line: selected.line };
}
/** First "(XX)" group of 2–3 uppercase letters in an event title. */
function extractTitleInitials(title) {
    const m = /\(([A-Z]{2,3})\)/.exec(title);
    return m ? m[1] : null;
}
const TITLE_TOKEN_RE = /\(([A-Za-z]{2,3})\)/g;
/**
 * Estimator fallback from the event title: scan left→right for parenthesized
 * 2–3 letter tokens (any case) and return the first one that is a known
 * estimator, uppercased. The known-set guard keeps short crew nicknames like
 * "(Ted)" from being mistaken for an estimator; tokens with "/" or >3 letters
 * ("(Daniel)", "(Ted/Rafael)") never match the regex at all.
 */
function estimatorFromTitle(title, known) {
    let first = null;
    for (const m of title.matchAll(TITLE_TOKEN_RE)) {
        const token = m[1].toUpperCase();
        if (first === null)
            first = token;
        if (known.has(token))
            return token;
    }
    return known.size === 0 ? first : null;
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
        // J: books lookup first; when the books are blank, fall back to the first
        // known-estimator "(XX)" token in the event title (see estimatorFromTitle).
        const booksEstimator = utMatch?.salesRep.trim() || nvMatch?.salesRep.trim() || '';
        const estimator = edit.estimator !== undefined
            ? edit.estimator
            : booksEstimator || estimatorFromTitle(ev.event, deps.knownEstimators) || '';
        /**
         * K: the `% Completion` recorded for this AP# in last week's finished
         * workbook (Slot 8) — never from the Description. Round 10 — a typed
         * value on the grid wins, so a wrong carry-forward can be corrected in
         * place before the workbook is downloaded.
         */
        const lastPaid = edit.lastPaid !== undefined
            ? edit.lastPaid
            : apKey && deps.lastWeekMap
                ? (deps.lastWeekMap.get(apKey) ?? null)
                : null;
        const qboJobName = edit.qboJobName !== undefined ? edit.qboJobName : (nvMatch?.name2 ?? utMatch?.name2 ?? '');
        const invoiceDate = edit.invoiceDate !== undefined ? edit.invoiceDate : nvMatch ? nvMatch.date : utMatch ? utMatch.date : null;
        const inQboTime = edit.inQboTime !== undefined
            ? edit.inQboTime
            : (apKey && deps.timesheetJobcodeByApKey.get(apKey)) || 'NotClockedIn';
        const nvTotal = apKey ? deps.nvTotals.get(apKey) : undefined;
        const utTotal = apKey ? deps.utTotals.get(apKey) : undefined;
        const lookedUpFound = nvTotal !== undefined || utTotal !== undefined;
        const jobTotalFound = edit.jobTotal !== undefined ? edit.jobTotal !== '#N/A' : lookedUpFound;
        const matchedTotal = edit.jobTotal !== undefined && edit.jobTotal !== '#N/A'
            ? edit.jobTotal
            : (nvTotal ?? utTotal ?? 0);
        const jobTotal = edit.jobTotal !== undefined ? edit.jobTotal : lookedUpFound ? matchedTotal : '#N/A';
        // Round 8 §2.2 — Apptivo $ is blank until someone types it; no auto-fill
        // from the job-total lookup, so $ Difference stays blank too (and the
        // Round 4 pink flag never fires off a blank pair).
        const apptivoEdited = edit.apptivo !== undefined;
        const apptivo = apptivoEdited ? (edit.apptivo ?? null) : null;
        const dollarDiff = jobTotalFound && apptivo !== null ? matchedTotal - apptivo : null;
        const dateDiff = invoiceDate && ev.startDate ? (0, format_1.dateDiffDays)(invoiceDate, ev.startDate) : null;
        return {
            id,
            eventIndex: ev.index,
            included,
            // Round 10 — every displayed column is overridable on the grid.
            event: edit.event !== undefined ? edit.event : ev.event,
            startDate: ev.startDate,
            endDate: ev.endDate,
            calendar: edit.calendar !== undefined ? edit.calendar : ev.calendar,
            attendees: edit.attendees !== undefined ? edit.attendees : ev.attendees,
            location: edit.location !== undefined ? edit.location : ev.location,
            description: edit.description !== undefined ? edit.description : ev.description,
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
// Round 8 §2.3 — % Completion value helpers live in format.ts; re-exported here
// so grid/export consumers can pull them from either module.
var format_2 = require("./format");
Object.defineProperty(exports, "isPercentNumber", { enumerable: true, get: function () { return format_2.isPercentNumber; } });
Object.defineProperty(exports, "isPercentState", { enumerable: true, get: function () { return format_2.isPercentState; } });
Object.defineProperty(exports, "isPercentText", { enumerable: true, get: function () { return format_2.isPercentText; } });
Object.defineProperty(exports, "matchesSetRate", { enumerable: true, get: function () { return format_2.matchesSetRate; } });
Object.defineProperty(exports, "matchesSubbedOut", { enumerable: true, get: function () { return format_2.matchesSubbedOut; } });
Object.defineProperty(exports, "percentExcelValue", { enumerable: true, get: function () { return format_2.percentExcelValue; } });
Object.defineProperty(exports, "percentIsValidState", { enumerable: true, get: function () { return format_2.percentIsValidState; } });
Object.defineProperty(exports, "percentText", { enumerable: true, get: function () { return format_2.percentText; } });
Object.defineProperty(exports, "percentTextValue", { enumerable: true, get: function () { return format_2.percentTextValue; } });
// ---------------------------------------------------------------------------
// Round 4 §1.1 — pink review flags (shared by the grid and the export)
// ---------------------------------------------------------------------------
/** `% Completion` gets the pink style on HOURLY text or a numeric > 100% / < 0%. */
function percentIsFlagged(v) {
    if (v === 'HOURLY')
        return true;
    return typeof v === 'number' && (v > 1 || v < 0);
}
/**
 * Round 8 §2.4 — `% Completion` turns solid red (white bold text) when it holds
 * no valid state: blank, 0%, or text that isn't HOURLY / SUBBED / SET RATE.
 */
function percentIsRedFlagged(v) {
    return !(0, format_1.percentIsValidState)(v);
}
/**
 * `In QBO Time?` values shared by 2+ included rows (a duplicate timesheet job
 * mapped to multiple events). NotClockedIn is flagged by its own rule.
 */
function duplicateInQboValues(rows) {
    const counts = new Map();
    for (const r of rows) {
        if (!r.included)
            continue;
        const v = r.inQboTime.trim();
        if (!v || v === 'NotClockedIn')
            continue;
        counts.set(v, (counts.get(v) ?? 0) + 1);
    }
    return new Set([...counts.entries()].filter(([, n]) => n >= 2).map(([v]) => v));
}
/** `$ Difference` gets the pink style on any nonzero value. */
function dollarDiffIsFlagged(v) {
    return typeof v === 'number' && v !== 0;
}
