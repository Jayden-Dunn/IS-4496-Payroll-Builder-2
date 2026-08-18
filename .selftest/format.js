"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.fmtDate = fmtDate;
exports.fmtDateShort = fmtDateShort;
exports.fmtAccounting = fmtAccounting;
exports.fmtJobTotal = fmtJobTotal;
exports.fmtMoney = fmtMoney;
exports.isPercentNumber = isPercentNumber;
exports.isPercentState = isPercentState;
exports.isPercentText = isPercentText;
exports.percentText = percentText;
exports.percentTextValue = percentTextValue;
exports.percentIsValidState = percentIsValidState;
exports.percentExcelValue = percentExcelValue;
exports.fmtPercent = fmtPercent;
exports.normalizeStateText = normalizeStateText;
exports.matchesSubbedOut = matchesSubbedOut;
exports.matchesSetRate = matchesSetRate;
exports.fmtHours = fmtHours;
exports.fmtHoursFixed = fmtHoursFixed;
exports.round2 = round2;
exports.isDarkColor = isDarkColor;
exports.dateOnly = dateOnly;
exports.dateDiffDays = dateDiffDays;
exports.addDays = addDays;
exports.isoToDate = isoToDate;
exports.dateToIso = dateToIso;
exports.parsePercentInput = parsePercentInput;
exports.parseMoneyInput = parseMoneyInput;
exports.moneyPlain = moneyPlain;
exports.formatCrewLeadWoNote = formatCrewLeadWoNote;
/** M/D/YYYY */
function fmtDate(d) {
    if (!d)
        return '';
    return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}
/** mm-dd-yy (reference Start/End Date display) */
function fmtDateShort(d) {
    if (!d)
        return '';
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const yy = String(d.getFullYear()).slice(-2);
    return `${mm}-${dd}-${yy}`;
}
/** Accounting-style number without $ (reference Job Total / Apptivo columns). */
function fmtAccounting(v) {
    if (v === null || v === undefined)
        return '';
    const abs = Math.abs(v);
    const s = abs.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return v < 0 ? `(${s})` : s;
}
/** Job Total $ — `#N/A` when unmatched, else accounting style. */
function fmtJobTotal(v) {
    if (v === '#N/A')
        return '#N/A';
    return fmtAccounting(v);
}
/** $14,411 (legacy whole-dollar display) */
function fmtMoney(v) {
    const r = Math.round(v);
    const s = Math.abs(r).toLocaleString('en-US');
    return (r < 0 ? '-$' : '$') + s;
}
// ---------------------------------------------------------------------------
// % Completion values (Round 8 §2.3 / §2.4)
// ---------------------------------------------------------------------------
const PERCENT_STATES = ['HOURLY', 'SUBBED', 'SET RATE'];
/** True for a numeric fraction (0.38 = 38%). */
function isPercentNumber(v) {
    return typeof v === 'number';
}
/** True for one of the literal states: HOURLY, SUBBED, SET RATE. */
function isPercentState(v) {
    return typeof v === 'string' && PERCENT_STATES.includes(v);
}
/** True for a free-text value the user typed (Round 8 §2.3). */
function isPercentText(v) {
    return typeof v === 'object' && v !== null && v.kind === 'text';
}
/** The verbatim text of a free-text value; null for every other variant. */
function percentText(v) {
    return isPercentText(v) ? v.text : null;
}
/**
 * Wrap text verbatim. Use `parsePercentInput` for anything a user typed — it
 * canonicalizes percents and the HOURLY / SUBBED / SET RATE states first.
 */
function percentTextValue(text) {
    return { kind: 'text', text };
}
/**
 * §2.4 — a cell holds a "valid" state when it has a nonzero %, or reads
 * HOURLY / SUBBED / SET RATE. Blank, 0% and free text are invalid (red).
 */
function percentIsValidState(v) {
    if (isPercentState(v))
        return true;
    return isPercentNumber(v) && v !== 0;
}
/** What a % cell writes into a worksheet: the fraction, or its text. */
function percentExcelValue(v) {
    if (v === null)
        return '';
    if (isPercentNumber(v))
        return v;
    if (isPercentState(v))
        return v;
    return v.text;
}
/** 0.95 → "95%", 'HOURLY' → "HOURLY", free text verbatim, null → "" */
function fmtPercent(v) {
    if (v === null)
        return '';
    if (isPercentState(v))
        return v;
    if (isPercentText(v))
        return v.text;
    return `${Math.round(v * 100)}%`;
}
/**
 * Lowercase, drop punctuation/hyphens and collapse whitespace so that
 * "Subbed-out", "SUBBED  OUT" and "set rate." all reduce to a plain phrase.
 */
function normalizeStateText(text) {
    return text
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
}
/** §2.4 — forgiving "subbed out" match (also plain "subbed"). */
function matchesSubbedOut(text) {
    return /\bsubbed\b/.test(normalizeStateText(text));
}
/** §2.4 — forgiving "set rate" match. */
function matchesSetRate(text) {
    return /\bset rate\b/.test(normalizeStateText(text));
}
/** Hours with at most 2 decimals: 4.26, 20.08, 8 */
function fmtHours(h) {
    return String(Math.round(h * 100) / 100);
}
/** Pivot hours always show two decimals: 4.26 → "4.26", 8 → "8.00" */
function fmtHoursFixed(h) {
    return h.toFixed(2);
}
function round2(n) {
    return Math.round(n * 100) / 100;
}
/**
 * Perceived-luminance check for a 6-digit hex fill (no '#'). Dark fills
 * (e.g. ND's 4472C4) get white text; light fills get black — computed, so any
 * palette edit in Settings keeps readable text automatically.
 */
function isDarkColor(hex6) {
    const r = parseInt(hex6.slice(0, 2), 16);
    const g = parseInt(hex6.slice(2, 4), 16);
    const b = parseInt(hex6.slice(4, 6), 16);
    return 0.299 * r + 0.587 * g + 0.114 * b < 140;
}
/** Strip the time-of-day portion (local). */
function dateOnly(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
/** Whole days a − b (may be negative). */
function dateDiffDays(a, b) {
    const ms = dateOnly(a).getTime() - dateOnly(b).getTime();
    return Math.round(ms / 86400000);
}
function addDays(d, n) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}
/** "2026-07-10" → local Date, or null. */
function isoToDate(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s.trim());
    if (!m)
        return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
/** Local Date → "2026-07-10" for <input type="date">. */
function dateToIso(d) {
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${d.getFullYear()}-${mm}-${dd}`;
}
/**
 * Parse user input for a % cell (Round 8 §2.3 — free text allowed).
 *
 * - percent-looking input ("38", "38%", ".38", "0.38") → fraction
 * - "hourly" (any case) → HOURLY; "subbed out" → SUBBED; "set rate" → SET RATE
 * - blank / "-" → null
 * - anything else → the text, stored and exported verbatim
 */
function parsePercentInput(text) {
    const t = text.trim();
    if (!t || t === '-')
        return null;
    const m = /^(-?(?:\d+(?:\.\d+)?|\.\d+))\s*(%?)$/.exec(t);
    if (m) {
        const n = Number(m[1]);
        if (m[2] === '%')
            return n / 100;
        return n > 1 ? n / 100 : n;
    }
    if (/hourly/i.test(t))
        return 'HOURLY';
    if (matchesSubbedOut(t))
        return 'SUBBED';
    if (matchesSetRate(t))
        return 'SET RATE';
    return { kind: 'text', text: t };
}
/** Parse user input for a money cell: "$2,095.00", "2095". Null when invalid/blank. */
function parseMoneyInput(text) {
    const t = text.trim().replace(/[$,\s]/g, '');
    if (!t)
        return null;
    const n = Number(t);
    return Number.isFinite(n) ? n : null;
}
/** Plain 1,234.56 — used in NOTES work-order strings. */
function moneyPlain(n) {
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
/**
 * Crew-lead NOTES column F: `w.o. $1159.20 @40% PS & 90% complete`.
 * Missing % of completion becomes `& ? Complete` (red-flagged by the caller).
 */
function formatCrewLeadWoNote(args) {
    const woBase = args.workOrderOverride !== null ? args.workOrderOverride : args.profitShare;
    const psPct = args.woPct !== null && Number.isFinite(args.woPct) ? Math.round(args.woPct * 100) : null;
    const psPart = `${psPct !== null ? String(psPct) : '?'}% PS`;
    let completePart;
    if (args.manualPay)
        completePart = args.manualPay;
    else if (args.percent === 'HOURLY')
        completePart = 'HOURLY';
    else if (typeof args.percent === 'number' && !args.missingPercent) {
        completePart = `${Math.round(args.percentCompletion * 100)}% complete`;
    }
    else if (args.percent && typeof args.percent === 'object' && args.percent.kind === 'text') {
        completePart = args.percent.text;
    }
    else {
        completePart = '? Complete';
    }
    return {
        text: `w.o. $${moneyPlain(woBase)} @${psPart} & ${completePart}`,
        flagMissing: completePart === '? Complete',
    };
}
