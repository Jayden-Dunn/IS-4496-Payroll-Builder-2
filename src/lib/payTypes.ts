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

/** Straight-time rate slot a pay item selects. */
export type PayRateKind = 'regular' | 'crewLead' | 'commercialHourly' | 'wages2' | 'salary' | 'davisBacon';

/** How an hour is classified once the pay item is read. */
export type PayItemKind =
  | { kind: 'regular'; rate: PayRateKind }
  | { kind: 'overtime'; rate: PayRateKind; reason: 'weekly' | 'daily' };

function norm(payItem: string): string {
  return payItem.trim().toLowerCase().replace(/[\s_]+/g, ' ');
}

/** True for any timesheet pay item that means "these hours are overtime". */
export function isOvertimePayItem(payItem: string): boolean {
  return /overtime|\bot\b/.test(norm(payItem));
}

/** True for the crew-lead hourly pay items (straight time or overtime). */
export function isCrewLeadPayItem(payItem: string): boolean {
  return /crew\s*lead/.test(norm(payItem));
}

/** True for the "Commercial Hourly" pay item (Round 9 §18). */
export function isCommercialHourlyPayItem(payItem: string): boolean {
  return /commercial\s*hourly/.test(norm(payItem));
}

/**
 * Round 11 — the "Davis Bacon" pay item (and its overtime twin). These hours
 * pay the Davis-Bacon rate from Settings, whatever the person's usual rate.
 */
export function isDavisBaconPayItem(payItem: string): boolean {
  return /davis[\s-]*bacon/.test(norm(payItem));
}

/** True for the "Wages 2" pay item (Round 9 §18). */
export function isWages2PayItem(payItem: string): boolean {
  return /^wages\s*2$/.test(norm(payItem));
}

/**
 * Classify one timesheet row's pay item. Unknown / blank items fall back to
 * straight-time `regular`, so a new QuickBooks item can never silently turn
 * into overtime.
 */
export function classifyPayItem(payItem: string): PayItemKind {
  const t = norm(payItem);
  const overtime = isOvertimePayItem(t);
  const daily = /daily/.test(t);

  let rate: PayRateKind = 'regular';
  // Checked first: a Davis-Bacon line is paid by statute, not by role.
  if (isDavisBaconPayItem(t)) rate = 'davisBacon';
  else if (isCrewLeadPayItem(t)) rate = 'crewLead';
  else if (isCommercialHourlyPayItem(t)) rate = 'commercialHourly';
  else if (isWages2PayItem(t)) rate = 'wages2';
  else if (/^salary$/.test(t)) rate = 'salary';

  return overtime
    ? { kind: 'overtime', rate, reason: daily ? 'daily' : 'weekly' }
    : { kind: 'regular', rate };
}

/** Short tag appended to a payroll line label for a non-plain pay item. */
export function payItemTag(payItem: string): string {
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
    case 'davisBacon':
      return ' [Davis Bacon]';
    default:
      return '';
  }
}
