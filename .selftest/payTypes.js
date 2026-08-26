"use strict";
/**
 * Round 9 §13/§16/§18 — the timesheet `payroll item` column is the source of
 * truth for BOTH how an hour is paid and whether it is overtime.
 *
 * QuickBooks Time emits one row per (day, job, pay item), so the week already
 * carries its own overtime split:
 *
 *   "Overtime Pay"                     → the person's own rate × 1.5
 *   "Crew Lead - Hourly Wage Overtime" → the crew-lead rate × 1.5
 *   "Daily Overtime"                   → NV >8 hrs/day, already split out
 *
 * Everything else is a straight-time pay item that selects WHICH rate slot the
 * person is paid from (`Commercial Hourly` and `Wages 2` are Round 9 §18
 * additions, saved per painter in Settings).
 *
 * Pure string logic — no dataset-specific knowledge beyond the QuickBooks
 * item names themselves.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.isOvertimePayItem = isOvertimePayItem;
exports.isCrewLeadPayItem = isCrewLeadPayItem;
exports.isCommercialHourlyPayItem = isCommercialHourlyPayItem;
exports.isWages2PayItem = isWages2PayItem;
exports.classifyPayItem = classifyPayItem;
exports.payItemTag = payItemTag;
function norm(payItem) {
    return payItem.trim().toLowerCase().replace(/[\s_]+/g, ' ');
}
/** True for any timesheet pay item that means "these hours are overtime". */
function isOvertimePayItem(payItem) {
    return /overtime|\bot\b/.test(norm(payItem));
}
/** True for the crew-lead hourly pay items (straight time or overtime). */
function isCrewLeadPayItem(payItem) {
    return /crew\s*lead/.test(norm(payItem));
}
/** True for the "Commercial Hourly" pay item (Round 9 §18). */
function isCommercialHourlyPayItem(payItem) {
    return /commercial\s*hourly/.test(norm(payItem));
}
/** True for the "Wages 2" pay item (Round 9 §18). */
function isWages2PayItem(payItem) {
    return /^wages\s*2$/.test(norm(payItem));
}
/**
 * Classify one timesheet row's pay item. Unknown / blank items fall back to
 * straight-time `regular`, so a new QuickBooks item can never silently turn
 * into overtime.
 */
function classifyPayItem(payItem) {
    const t = norm(payItem);
    const overtime = isOvertimePayItem(t);
    const daily = /daily/.test(t);
    let rate = 'regular';
    if (isCrewLeadPayItem(t))
        rate = 'crewLead';
    else if (isCommercialHourlyPayItem(t))
        rate = 'commercialHourly';
    else if (isWages2PayItem(t))
        rate = 'wages2';
    else if (/^salary$/.test(t))
        rate = 'salary';
    return overtime
        ? { kind: 'overtime', rate, reason: daily ? 'daily' : 'weekly' }
        : { kind: 'regular', rate };
}
/** Short tag appended to a payroll line label for a non-plain pay item. */
function payItemTag(payItem) {
    const c = classifyPayItem(payItem);
    if (c.kind === 'overtime') {
        return c.reason === 'daily' ? ' [Daily OT]' : ' [OT]';
    }
    switch (c.rate) {
        case 'crewLead':
            return ' [Crew Lead]';
        case 'commercialHourly':
            return ' [Commercial Hourly]';
        case 'wages2':
            return ' [Wages 2]';
        default:
            return '';
    }
}
