/**
 * Round 8 §4.3 — resolve a job's CUSTOMER from its AP#.
 *
 * The Sales by Customer Summary is hierarchical: a non-indented customer row
 * ("Aaron Ward") followed by indented job rows ("   AP24314 Ext"). The owning
 * customer of a job is the nearest non-indented row above it. The P&L header
 * has the same shape sideways (a job column belongs to the nearest non-`Total`
 * customer header to its left) and serves as the secondary source.
 *
 * Rows come from the already-parsed `mirrorTabs` models — no second parser.
 */
import { extractApKey } from './apkey';
import { baseApKey } from './jobNames';
import type { SalesByCustMirrorRow } from './mirrorTabs';
import type { PlParseResult } from './types';

/** Which Sales by Customer Summary a job belongs to. */
export type CustomerEntity = 'NV' | 'UT';

export type CustomerSource = 'salesNv' | 'salesUt' | 'pl' | null;

export interface CustomerResolution {
  customer: string;
  source: CustomerSource;
}

export interface CustomerByApSources {
  nvSalesByCust?: readonly SalesByCustMirrorRow[] | null;
  utSalesByCust?: readonly SalesByCustMirrorRow[] | null;
  pl?: PlParseResult | null;
}

export interface CustomerByApResolver {
  /** Preferred file first (NV/Reno → NV, UT → UT), then the other, then P&L. */
  resolve: (apKey: string | null | undefined, entity: CustomerEntity) => CustomerResolution;
}

const UNRESOLVED: CustomerResolution = { customer: '', source: null };

/** Summary rows indent job lines; customer headers start at column A, flush left. */
function isIndented(text: string): boolean {
  return /^[\s\u00a0]/.test(text);
}

function isTotalRow(text: string): boolean {
  return /^\s*total\b/i.test(text);
}

/**
 * apKey → owning customer for one Sales by Customer Summary. First occurrence
 * wins so a repeated AP# behaves like every other first-match lookup here.
 */
export function customerByApFromSalesSummary(
  rows: readonly SalesByCustMirrorRow[] | null | undefined,
): Map<string, string> {
  const map = new Map<string, string>();
  if (!rows) return map;
  let currentCustomer = '';
  for (const row of rows) {
    const text = row.text ?? '';
    if (!text.trim()) continue;
    if (isTotalRow(text)) continue;
    const apKey = row.apKey || extractApKey(text);
    if (apKey) {
      if (currentCustomer && !map.has(apKey)) map.set(apKey, currentCustomer);
      continue;
    }
    // A flush-left, AP-less row is the customer header the job rows hang under.
    if (!isIndented(text)) currentCustomer = text.trim();
  }
  return map;
}

/** apKey → customer header from the P&L column structure (already parsed). */
export function customerByApFromPl(pl: PlParseResult | null | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!pl) return map;
  for (const [apKey, metrics] of pl.byJob) {
    const name = metrics.customerName?.trim();
    if (name && !map.has(apKey)) map.set(apKey, name);
  }
  return map;
}

export function buildCustomerByApResolver(sources: CustomerByApSources): CustomerByApResolver {
  const nv = customerByApFromSalesSummary(sources.nvSalesByCust);
  const ut = customerByApFromSalesSummary(sources.utSalesByCust);
  const pl = customerByApFromPl(sources.pl);

  const lookup = (map: ReadonlyMap<string, string>, apKey: string): string => {
    const exact = map.get(apKey);
    if (exact) return exact;
    // Change orders ("24097-A") bill under the base job's customer.
    const base = baseApKey(apKey);
    return (base ? map.get(base) : undefined) ?? '';
  };

  return {
    resolve(apKey, entity) {
      if (!apKey || apKey === 'OFFICE') return UNRESOLVED;
      const order: { map: ReadonlyMap<string, string>; source: CustomerSource }[] =
        entity === 'UT'
          ? [
              { map: ut, source: 'salesUt' },
              { map: nv, source: 'salesNv' },
            ]
          : [
              { map: nv, source: 'salesNv' },
              { map: ut, source: 'salesUt' },
            ];
      for (const step of order) {
        const hit = lookup(step.map, apKey);
        if (hit) return { customer: hit, source: step.source };
      }
      const fromPl = lookup(pl, apKey);
      if (fromPl) return { customer: fromPl, source: 'pl' };
      return UNRESOLVED;
    },
  };
}
