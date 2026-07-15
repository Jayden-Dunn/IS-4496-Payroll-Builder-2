import { normalizeApInput } from './apkey';
import { dateDiffDays } from './format';
import { buildSalesIndex } from './joins';
import type {
  EditsMap,
  PercentValue,
  PrevPercent,
  ReportRow,
  SalesRow,
  TimeNaviEvent,
} from './types';

// ---------------------------------------------------------------------------
// Description → % pre-fill parser
// ---------------------------------------------------------------------------

const PCT_RE = /(\d{1,3})\s*%/;

/**
 * Scan Description lines top-down (newest status first). The first line with
 * a percent AND the word "complete" nearby wins; if the first status-looking
 * line says HOURLY instead, return 'HOURLY'. Otherwise leave blank — percent
 * completion is a human judgment, the app only suggests.
 */
export function suggestPercent(description: string): { value: PercentValue; line: string | null } {
  for (const raw of description.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line) continue;
    const m = PCT_RE.exec(line);
    if (m && /complete/i.test(line)) {
      const n = Number(m[1]);
      if (n >= 0 && n <= 100) return { value: n / 100, line };
    }
    if (/hourly/i.test(line)) return { value: 'HOURLY', line };
  }
  return { value: null, line: null };
}

/** First "(XX)" group of 2–3 uppercase letters in an event title. */
export function extractTitleInitials(title: string): string | null {
  const m = /\(([A-Z]{2,3})\)/.exec(title);
  return m ? m[1] : null;
}

// ---------------------------------------------------------------------------
// Building the TimeNavi Export rows (columns A–S)
// ---------------------------------------------------------------------------

export interface ReportDeps {
  nvSales: SalesRow[];
  utSales: SalesRow[];
  nvTotals: ReadonlyMap<string, number>;
  utTotals: ReadonlyMap<string, number>;
  lastWeekMap: ReadonlyMap<string, PrevPercent> | null;
  /** apKey → jobcode_2 of the FIRST timesheet row with that key. */
  timesheetJobcodeByApKey: ReadonlyMap<string, string>;
}

export function reportRowId(ev: TimeNaviEvent): string {
  // Keyed by apKey + row position so re-parsing the same file keeps edits.
  return `ev${ev.index}:${ev.apKey ?? ''}`;
}

export function buildReportRows(events: TimeNaviEvent[], edits: EditsMap, deps: ReportDeps): ReportRow[] {
  const nv = buildSalesIndex(deps.nvSales);
  const ut = buildSalesIndex(deps.utSales);

  return events.map((ev) => {
    const id = reportRowId(ev);
    const edit = edits[id] ?? {};

    const apKey = edit.apKey !== undefined ? edit.apKey : ev.apKey;
    const included = edit.included !== undefined ? edit.included : apKey !== null;
    const percent = edit.percent !== undefined ? edit.percent : ev.suggestedPercent;

    const nvMatch = apKey ? nv.byApKey.get(apKey) : undefined;
    const utMatch = apKey ? ut.byApKey.get(apKey) : undefined;

    // Column J: estimator lookup is UT first, then NV (intentionally the
    // reverse of columns L/R). Blank sales reps fall through.
    const estimator =
      edit.estimator !== undefined
        ? edit.estimator
        : utMatch?.salesRep.trim() || nvMatch?.salesRep.trim() || '';

    const lastPaid = apKey && deps.lastWeekMap ? deps.lastWeekMap.get(apKey) ?? null : null;

    // Columns L and R: NV first, then UT.
    const qboJobName = nvMatch?.name2 ?? utMatch?.name2 ?? '';
    const invoiceDate = nvMatch ? nvMatch.date : utMatch ? utMatch.date : null;

    const inQboTime = (apKey && deps.timesheetJobcodeByApKey.get(apKey)) || 'NotClockedIn';

    const nvTotal = apKey ? deps.nvTotals.get(apKey) : undefined;
    const utTotal = apKey ? deps.utTotals.get(apKey) : undefined;
    const jobTotalFound = nvTotal !== undefined || utTotal !== undefined;
    const jobTotal = nvTotal ?? utTotal ?? 0;

    const apptivoEdited = edit.apptivo !== undefined;
    const apptivo = apptivoEdited ? (edit.apptivo as number) : jobTotal;

    const dateDiff = invoiceDate && ev.startDate ? dateDiffDays(invoiceDate, ev.startDate) : null;

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
      suggestedSourceLine: ev.suggestedSourceLine,
      estimator,
      lastPaid,
      qboJobName,
      inQboTime,
      jobTotal,
      jobTotalFound,
      apptivo,
      apptivoEdited,
      dollarDiff: jobTotal - apptivo,
      invoiceDate,
      dateDiff,
      titleInitials: extractTitleInitials(ev.event),
    };
  });
}

/** Sort by estimator (legend order from Settings), then Start Date; stable. */
export function sortReportRows(rows: ReportRow[], estimatorOrder: string[]): ReportRow[] {
  const idx = new Map<string, number>(estimatorOrder.map((e, i) => [e, i]));
  return rows
    .map((r, i) => ({ r, i }))
    .sort((a, b) => {
      const ea = idx.get(a.r.estimator) ?? estimatorOrder.length;
      const eb = idx.get(b.r.estimator) ?? estimatorOrder.length;
      if (ea !== eb) return ea - eb;
      const ta = a.r.startDate ? a.r.startDate.getTime() : Number.POSITIVE_INFINITY;
      const tb = b.r.startDate ? b.r.startDate.getTime() : Number.POSITIVE_INFINITY;
      if (ta !== tb) return ta - tb;
      return a.i - b.i;
    })
    .map((x) => x.r);
}

/** Re-export so callers editing AP# share the exact same normalization. */
export { normalizeApInput };
