"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.fmtDate = fmtDate;
exports.fmtDateShort = fmtDateShort;
exports.fmtAccounting = fmtAccounting;
exports.fmtJobTotal = fmtJobTotal;
exports.fmtMoney = fmtMoney;
exports.fmtPercent = fmtPercent;
exports.fmtHours = fmtHours;
exports.fmtHoursFixed = fmtHoursFixed;
exports.round2 = round2;
exports.dateOnly = dateOnly;
exports.dateDiffDays = dateDiffDays;
exports.addDays = addDays;
exports.isoToDate = isoToDate;
exports.dateToIso = dateToIso;
exports.parsePercentInput = parsePercentInput;
exports.parseMoneyInput = parseMoneyInput;
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
/** 0.95 → "95%", 'HOURLY' → "HOURLY", null → "" */
function fmtPercent(v) {
    if (v === null)
        return '';
    if (v === 'HOURLY')
        return 'HOURLY';
    return `${Math.round(v * 100)}%`;
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
/** Parse user input for a % cell: "95", "95%", "0.95", "hourly", "-", "". */
function parsePercentInput(text) {
    const t = text.trim();
    if (!t || t === '-')
        return null;
    if (/hourly/i.test(t))
        return 'HOURLY';
    const m = /^(-?\d+(?:\.\d+)?)\s*%?$/.exec(t);
    if (!m)
        return null;
    const n = Number(m[1]);
    return n > 1 ? n / 100 : n;
}
/** Parse user input for a money cell: "$2,095.00", "2095". Null when invalid/blank. */
function parseMoneyInput(text) {
    const t = text.trim().replace(/[$,\s]/g, '');
    if (!t)
        return null;
    const n = Number(t);
    return Number.isFinite(n) ? n : null;
}
