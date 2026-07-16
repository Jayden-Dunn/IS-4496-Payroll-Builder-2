"use strict";
/**
 * The canonical job key. Every join in this app is on the AP number.
 * apKeys are ALWAYS compared as strings (never numbers) so suffixed keys
 * like "21848-A5" survive round trips between sources.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.AP_RE = void 0;
exports.extractApKey = extractApKey;
exports.timesheetApKey = timesheetApKey;
exports.normalizeApInput = normalizeApInput;
exports.AP_RE = /AP\s*-?\s*(\d{4,6}(?:-[A-Za-z0-9]+)*)/i;
/** Extract the AP key from free text (event title, jobcode_2, QBO job name…). */
function extractApKey(text) {
    if (!text)
        return null;
    const m = exports.AP_RE.exec(text);
    return m ? m[1].toUpperCase().trim() : null;
}
/**
 * apKey for a timesheet jobcode_2. Office entries (e.g. "OFFICE [Please put
 * AP# in notes]") map to the special key "OFFICE".
 */
function timesheetApKey(jobcode2) {
    const t = jobcode2.trim();
    if (/^office/i.test(t))
        return 'OFFICE';
    return extractApKey(t);
}
/**
 * Normalize a user-typed AP# ("AP24314", "24314", "21848-a5", "-", "").
 * Returns null when the field should be treated as empty.
 */
function normalizeApInput(raw) {
    const t = raw.trim().toUpperCase();
    if (!t || t === '-')
        return null;
    if (t === 'OFFICE')
        return 'OFFICE';
    const viaRe = extractApKey(t);
    if (viaRe)
        return viaRe;
    if (/^\d{1,6}(?:-[A-Z0-9]+)*$/.test(t))
        return t;
    // Fall back to the raw uppercased text so odd keys still compare equal.
    return t;
}
