import { timesheetApKey } from './apkey';
import type { EnrichedTimesheetRow, SalesRow, TimesheetRow } from './types';

/** First (= earliest, file order) sales row matching an apKey. */
export function firstByApKey(list: SalesRow[], key: string | null): SalesRow | null {
  if (!key) return null;
  for (const row of list) {
    if (row.apKey === key) return row;
  }
  return null;
}

export interface SalesIndex {
  byApKey: Map<string, SalesRow>;
  byName2: Map<string, SalesRow>;
}

/** O(1) lookup maps keeping only the FIRST (earliest) row per key. */
export function buildSalesIndex(list: SalesRow[]): SalesIndex {
  const byApKey = new Map<string, SalesRow>();
  const byName2 = new Map<string, SalesRow>();
  for (const row of list) {
    if (row.apKey && !byApKey.has(row.apKey)) byApKey.set(row.apKey, row);
    const n2 = row.name2.trim();
    if (n2 && !byName2.has(n2)) byName2.set(n2, row);
  }
  return { byApKey, byName2 };
}

/**
 * Mirror of the workbook's helper columns AA–AE on the timesheet:
 * apKey, estimator, and the not-on-calendar flag.
 *
 * Estimator resolution: the RESOLVED estimator from the TimeNavi Export grid
 * (books lookup + title fallback + user edits) is the single source of truth
 * — when the row's apKey matches an included event, that value is used, so
 * grid edits and title-derived estimators flow into the pivot. Jobs not on
 * the calendar fall back to the books lookup (exact jobcode_2 → name2 match
 * first, then apKey match; NV before UT). OFFICE → ND.
 */
export function enrichTimesheet(
  rows: TimesheetRow[],
  nvSales: SalesRow[],
  utSales: SalesRow[],
  includedEventApKeys: ReadonlySet<string>,
  resolvedEstimatorByApKey: ReadonlyMap<string, string>,
): EnrichedTimesheetRow[] {
  const nv = buildSalesIndex(nvSales);
  const ut = buildSalesIndex(utSales);
  return rows.map((row) => {
    const apKey = timesheetApKey(row.jobcode2);
    let estimator = '#N/A';
    if (apKey === 'OFFICE') {
      estimator = 'ND';
    } else {
      const resolved = apKey ? resolvedEstimatorByApKey.get(apKey) : undefined;
      if (resolved) {
        estimator = resolved;
      } else {
        const j2 = row.jobcode2.trim();
        const candidates: (SalesRow | undefined)[] = [
          nv.byName2.get(j2),
          ut.byName2.get(j2),
          apKey ? nv.byApKey.get(apKey) : undefined,
          apKey ? ut.byApKey.get(apKey) : undefined,
        ];
        for (const c of candidates) {
          const rep = c?.salesRep.trim();
          if (rep) {
            estimator = rep;
            break;
          }
        }
      }
    }
    const notOnCalendar = apKey !== null && apKey !== 'OFFICE' && !includedEventApKeys.has(apKey);
    return { ...row, apKey, estimator, notOnCalendar };
  });
}
