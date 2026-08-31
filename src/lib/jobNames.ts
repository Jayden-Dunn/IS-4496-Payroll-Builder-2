/**
 * Round 5 §4 — shared job/customer display-name chain.
 * Every renderer of a PO/Name or customer/job pair consumes this; the raw
 * TimeNavi event title must never appear as a display name.
 */
import { extractApKey } from './apkey';
import type { LastWeekPbjRow } from './parsers';
import type { EnrichedTimesheetRow, PlParseResult, SalesRow } from './types';

export interface JobDisplayPair {
  poLabel: string;
  customerName: string;
  /** Names came from the base-key QBO retry (suffixed AP# missed exact match). */
  usedBaseKey: boolean;
  /** Nothing resolved — poLabel is AP{key}, customer blank. */
  unresolved: boolean;
}

export interface JobDisplaySources {
  timesheetByAp: ReadonlyMap<string, { customer: string; job: string }>;
  carryByAp: ReadonlyMap<string, { customer: string; job: string }>;
  /** NV-then-UT sales rows keyed by apKey (first hit wins). */
  salesByAp: ReadonlyMap<string, { name1: string; name2: string }>;
  pl: PlParseResult | null;
  /** Event title used only as the cleaned last resort (§4 step 6). */
  eventTitle?: string | null;
}

/** Numeric base of a suffixed AP key: "24180-A" → "24180"; bare keys → null. */
export function baseApKey(apKey: string): string | null {
  const m = /^(\d{4,6})-(.+)$/i.exec(apKey.trim());
  return m ? m[1] : null;
}

/**
 * Round 10 — the office jobcode is not a customer. Someone clocking office
 * time against a real AP# used to make that job's Name read "ALLPRO OFFICE".
 */
export function isOfficeCustomer(name: string): boolean {
  return /^all\s*pro\s+office$/i.test(name.trim());
}

/** Build the timesheet apKey → {customer=jobcode_1, job=jobcode_2} map. */
export function timesheetDisplayMap(
  rows: readonly EnrichedTimesheetRow[],
): Map<string, { customer: string; job: string }> {
  const m = new Map<string, { customer: string; job: string }>();
  for (const r of rows) {
    if (!r.apKey || r.apKey === 'OFFICE' || m.has(r.apKey)) continue;
    // Round 10 — keep the job text, drop an office "customer": the real name
    // comes from Sales by Rep's `Name 1` further down the chain.
    m.set(r.apKey, { customer: isOfficeCustomer(r.jobcode1) ? '' : r.jobcode1, job: r.jobcode2 });
  }
  return m;
}

/** Slot-7 carry-forward pairs keyed by apKey (first occurrence wins). */
export function carryDisplayMap(
  rows: readonly LastWeekPbjRow[] | null | undefined,
): Map<string, { customer: string; job: string }> {
  const m = new Map<string, { customer: string; job: string }>();
  for (const r of rows ?? []) {
    if (!r.apKey || m.has(r.apKey)) continue;
    m.set(r.apKey, { customer: r.customer, job: r.job });
  }
  return m;
}

/**
 * Sales lookup: NV rows first, then UT — mirrors the TimeNavi QBO Job Name
 * column. First name1/name2 for each apKey wins.
 */
export function salesDisplayMap(
  nvSales: readonly SalesRow[] | null | undefined,
  utSales: readonly SalesRow[] | null | undefined,
): Map<string, { name1: string; name2: string }> {
  const m = new Map<string, { name1: string; name2: string }>();
  for (const list of [nvSales, utSales]) {
    if (!list) continue;
    for (const r of list) {
      if (!r.apKey || m.has(r.apKey)) continue;
      if (!r.name1 && !r.name2) continue;
      m.set(r.apKey, { name1: r.name1, name2: r.name2 });
    }
  }
  return m;
}

/**
 * Round 5 §4 step 6 — parse a calendar title into display names without ever
 * dumping the verbatim string. Returns null when nothing usable remains.
 */
export function cleanEventTitle(
  title: string,
  apKey: string | null,
): { poLabel: string; customerName: string } | null {
  let t = title.trim();
  if (!t) return null;
  // Strip leading decoration glyphs (♦, •, etc.).
  t = t.replace(/^[\u2666\u25C6\u2022\u25CF\u25A0♦●•]+\s*/u, '');
  // Drop *[STATUS]* tokens like *PAID*.
  t = t.replace(/\*[A-Za-z][A-Za-z0-9 _-]*\*/g, ' ');
  // Drop $ / k amount tokens: "13k", "3.5k", "$8k", "13,000".
  t = t.replace(/\$?\d+(?:[.,]\d+)?k\b/gi, ' ');
  t = t.replace(/\$[\d,]+(?:\.\d+)?/g, ' ');

  // Collect parentheticals; drop known-estimator (1–3 uppercase letters) and
  // the trailing crew parenthetical (last one that is not an estimator).
  const parenRe = /\(([^)]*)\)/g;
  const parens: { raw: string; inner: string; start: number; end: number }[] = [];
  let pm: RegExpExecArray | null;
  while ((pm = parenRe.exec(t)) !== null) {
    parens.push({ raw: pm[0], inner: pm[1].trim(), start: pm.index, end: pm.index + pm[0].length });
  }
  const isEstimatorParen = (inner: string): boolean => /^[A-Za-z]{1,3}$/.test(inner.trim());
  // Strip all estimator parentheticals.
  for (const p of [...parens].reverse()) {
    if (isEstimatorParen(p.inner)) {
      t = t.slice(0, p.start) + ' ' + t.slice(p.end);
    }
  }
  // Re-scan; drop the last remaining parenthetical as the crew lead.
  const remaining: { start: number; end: number; inner: string }[] = [];
  parenRe.lastIndex = 0;
  while ((pm = parenRe.exec(t)) !== null) {
    remaining.push({ start: pm.index, end: pm.index + pm[0].length, inner: pm[1].trim() });
  }
  if (remaining.length > 0) {
    const last = remaining[remaining.length - 1];
    t = t.slice(0, last.start) + ' ' + t.slice(last.end);
  }

  // Bracketed text → preferred customer name.
  let bracketCustomer = '';
  const br = /\[([^\]]+)\]/.exec(t);
  if (br) {
    bracketCustomer = br[1].trim();
    t = t.replace(br[0], ' ');
  }

  // Pull the AP token (and anything after a comma that introduces it).
  const apMatch = /,\s*AP\s*-?\s*\d{4,6}(?:-[A-Za-z0-9]+)*/i.exec(t);
  let beforeAp = t;
  if (apMatch && apMatch.index !== undefined) {
    beforeAp = t.slice(0, apMatch.index);
  } else {
    // No comma-AP form — strip a bare AP token in place.
    beforeAp = t.replace(/AP\s*-?\s*\d{4,6}(?:-[A-Za-z0-9]+)*/gi, ' ');
  }

  // Short type tokens for the PO label.
  const typeTok = /\b(Int\/Ext|Int|Ext|T\s*&\s*M)\b/i.exec(title);
  const typePart = typeTok ? typeTok[1].replace(/\s+/g, '') : '';
  const key = apKey ?? extractApKey(title);
  const poLabel = key ? `AP${key}${typePart ? ` ${typePart}` : ''}` : '';

  let customerName = bracketCustomer;
  if (!customerName) {
    customerName = beforeAp
      .replace(/\b(Int\/Ext|Int|Ext|T\s*&\s*M)\b/gi, ' ')
      .replace(/\bPh\s*\d+\b/gi, ' ')
      .replace(/\bPart\b[\s\w-]*/gi, ' ')
      .replace(/[^\w\s'/.-]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    // Drop leading connector junk and very short leftovers.
    customerName = customerName.replace(/^(and|or|for|the)\s+/i, '').trim();
  }

  if (!poLabel && !customerName) return null;
  return { poLabel: poLabel || (key ? `AP${key}` : ''), customerName };
}

/**
 * §4 chain — first hit wins per field. `customerName` is never a copy of the
 * job string; unresolved jobs get `AP{key}` / blank + the caller's exception.
 */
export function jobDisplayPair(apKey: string | null | undefined, sources: JobDisplaySources): JobDisplayPair {
  const empty: JobDisplayPair = { poLabel: '', customerName: '', usedBaseKey: false, unresolved: true };
  if (!apKey || apKey === 'OFFICE') {
    return apKey === 'OFFICE'
      ? { poLabel: 'OFFICE', customerName: 'ALLPRO OFFICE', usedBaseKey: false, unresolved: false }
      : empty;
  }

  const hit = (po: string, cust: string, usedBaseKey = false): JobDisplayPair => ({
    poLabel: po.trim(),
    customerName: cust.trim(),
    usedBaseKey,
    unresolved: false,
  });

  /**
   * Round 10 — `Name 1` from Sales by Rep is the authoritative customer name.
   * The PO label still prefers the timesheet / carry job text, so only the
   * customer half jumps the queue.
   */
  const salesName1 = sources.salesByAp.get(apKey)?.name1?.trim() ?? '';
  const withCustomer = (po: string, cust: string, usedBaseKey = false): JobDisplayPair =>
    hit(po, salesName1 || cust, usedBaseKey);

  // 1. Timesheet
  const ts = sources.timesheetByAp.get(apKey);
  if (ts && (ts.job || ts.customer)) return withCustomer(ts.job || `AP${apKey}`, ts.customer);

  // 2. Slot-7 carry-forward
  const carry = sources.carryByAp.get(apKey);
  if (carry && (carry.job || carry.customer)) {
    return withCustomer(carry.job || `AP${apKey}`, carry.customer);
  }

  // 3. QBO sales exact
  const sales = sources.salesByAp.get(apKey);
  if (sales && (sales.name2 || sales.name1)) return hit(sales.name2 || `AP${apKey}`, sales.name1);

  // 4. Base-key retry on QBO
  const base = baseApKey(apKey);
  if (base) {
    const baseSales = sources.salesByAp.get(base);
    if (baseSales && (baseSales.name2 || baseSales.name1)) {
      return hit(baseSales.name2 || `AP${base}`, baseSales.name1, true);
    }
  }

  // 5. P&L
  const pl = sources.pl?.byJob.get(apKey) ?? (base ? sources.pl?.byJob.get(base) : undefined);
  if (pl?.plJobHeader) {
    return hit(pl.plJobHeader, pl.customerName ?? '', base !== null && !sources.pl?.byJob.has(apKey));
  }

  // 6. Cleaned event title
  if (sources.eventTitle) {
    const cleaned = cleanEventTitle(sources.eventTitle, apKey);
    if (cleaned && (cleaned.poLabel || cleaned.customerName)) {
      return hit(cleaned.poLabel || `AP${apKey}`, cleaned.customerName);
    }
  }

  // 7. Nothing
  return { poLabel: `AP${apKey}`, customerName: '', usedBaseKey: false, unresolved: true };
}
