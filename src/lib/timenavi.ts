import { normalizeApInput } from './apkey';
import {
  dateDiffDays,
  dateOnly,
  matchesSetRate,
  matchesSubbedOut,
  percentIsValidState,
} from './format';
import { buildSalesIndex } from './joins';
import type {
  EditsMap,
  JobTotalValue,
  PercentValue,
  PrevPercent,
  ReportRow,
  SalesRow,
  TimeNaviEvent,
} from './types';

// ---------------------------------------------------------------------------
// Description → % pre-fill parser (pay-date driven)
// ---------------------------------------------------------------------------

/**
 * A status line opens with its date. Leading bullets are tolerated — ♦, ◆, •,
 * *, dashes, ">" — because a line the regex skips is not just lost: the line
 * below it gets promoted to "this week", which is how an old 100% ended up in
 * `% Completion` above a newer HOURLY entry.
 */
const STATUS_DATE_RE = /^\s*[♦◆•*>\-–—]*\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/;
/**
 * Round 12 — the same opening with the year left off ("9/04/HOURLY - JT",
 * "10/9 50% COMPLETE"). Only trusted when the line also states a status, so a
 * plain dated note ("10/9 STILL HAVE 2 MORE RMS TO GO") never becomes a week.
 */
const STATUS_DATE_NO_YEAR_RE = /^\s*[♦◆•*>\-–—]*\s*(\d{1,2})\/(\d{1,2})(?![\d])/;
const PCT_RE = /(\d{1,3}(?:\.\d+)?)\s*%/;
/** "no pay", "no-pay", "nopay" — a week the job was not paid; never a "last paid". */
const NO_PAY_RE = /\bno[\s-]*pay\b/i;
/**
 * Round 12 — two entries run together on one line
 * ("11/06/26 HOURLY -SH10/09/26 No Pay -SH"). A date in the middle of a line
 * starts a new entry only when a status follows it, so a date mentioned in
 * passing ("touch-ups due 7/10/26") stays part of its own line.
 */
const INLINE_ENTRY_RE =
  /(^|[^\d/])(\d{1,2}\/\d{1,2}\/\d{2,4}\s*[-–—:]?\s*(?:\d{1,3}(?:\.\d+)?\s*%|hourly|no[\s-]*pay|subbed|set\s*rate))/gi;

/** Cut one text line into its entries; almost always just the line itself. */
function splitInlineEntries(raw: string): string[] {
  const cuts: number[] = [];
  INLINE_ENTRY_RE.lastIndex = 0;
  for (let m = INLINE_ENTRY_RE.exec(raw); m; m = INLINE_ENTRY_RE.exec(raw)) {
    const at = m.index + m[1].length;
    if (at > 0) cuts.push(at);
  }
  if (cuts.length === 0) return [raw];
  const out: string[] = [];
  let from = 0;
  for (const at of cuts) {
    out.push(raw.slice(from, at));
    from = at;
  }
  out.push(raw.slice(from));
  return out;
}

interface StatusLine {
  date: Date;
  line: string;
  /** File order, so lines sharing a date still sort deterministically. */
  order: number;
}

function statusLineDate(month: string, day: string, year: string): Date {
  let y = Number(year);
  if (y < 100) y += 2000;
  return new Date(y, Number(month) - 1, Number(day));
}

function extractStatusLines(description: string, payDate: Date): StatusLine[] {
  const out: StatusLine[] = [];
  for (const raw of description.split(/\r?\n/)) {
    // The first piece keeps whatever leads the line; later pieces each open
    // with their own date.
    for (const piece of splitInlineEntries(raw)) {
      const line = piece.trim();
      if (!line) continue;
      const m = STATUS_DATE_RE.exec(line);
      if (m) {
        out.push({ date: statusLineDate(m[1], m[2], m[3]), line, order: out.length });
        continue;
      }
      const short = STATUS_DATE_NO_YEAR_RE.exec(line);
      if (!short) continue;
      if (readStatusValue(line.slice(short[0].length)) === null && !NO_PAY_RE.test(line)) continue;
      // No year written: the latest such date that is not after the pay date.
      const payYear = payDate.getFullYear();
      let date = statusLineDate(short[1], short[2], String(payYear));
      if (dateOnly(date).getTime() > dateOnly(payDate).getTime()) {
        date = statusLineDate(short[1], short[2], String(payYear - 1));
      }
      out.push({ date, line, order: out.length });
    }
  }
  return out;
}

/** What one status line says about the job's state that week. */
function readStatusValue(text: string): PercentValue {
  const pct = PCT_RE.exec(text);
  if (pct) {
    const n = Number(pct[1]);
    if (n >= 0 && n <= 100) return n / 100;
  }
  if (/hourly/i.test(text)) return 'HOURLY';
  // §2.4 — the slot where the % normally sits can instead read "subbed out"
  // or "set rate"; both are valid states, not missing data.
  if (matchesSubbedOut(text)) return 'SUBBED';
  if (matchesSetRate(text)) return 'SET RATE';
  return null;
}

/**
 * This week's line: the one dated on the pay date, else the latest line dated
 * on or before it. Lines dated after the pay date are next week's notes.
 */
function selectCurrentLine(lines: readonly StatusLine[], payDate: Date): StatusLine | null {
  const payMs = dateOnly(payDate).getTime();
  const exact = lines.find((s) => dateOnly(s.date).getTime() === payMs);
  if (exact) return exact;
  const eligible = lines.filter((s) => dateOnly(s.date).getTime() <= payMs);
  if (eligible.length === 0) return null;
  return eligible.reduce((best, cur) =>
    dateOnly(cur.date).getTime() > dateOnly(best.date).getTime() ? cur : best,
  );
}

export interface DescriptionPercents {
  /** This week's state — `% Completion`. */
  current: { value: PercentValue; line: string | null };
  /**
   * The most recent ACTUAL percentage below this week's line, skipping "no
   * pay" weeks, 0% and any line that is not a percentage (HOURLY, subbed, set
   * rate). Not necessarily last week: a job unpaid for three weeks still
   * reports the last % it was actually paid at.
   *
   * Round 12 — when no earlier line carries a percentage but one says HOURLY,
   * the job was last paid by the hour, and this reads `'HOURLY'`.
   */
  previous: { value: PrevPercent; line: string } | null;
  /**
   * True when this week's line is the only dated line on or before the pay
   * date. That means the job started this pay week, so there IS no last paid
   * percentage — and printing this week's figure again as last week's is the
   * duplication bug this guards against.
   */
  onlyCurrent: boolean;
}

/**
 * Read a calendar description once for both columns. Keeping `% Completion`
 * and `Last % Paid` on one parse guarantees they can never pick the same line.
 */
export function readDescriptionPercents(description: string, payDate: Date | null): DescriptionPercents {
  const none: DescriptionPercents = { current: { value: null, line: null }, previous: null, onlyCurrent: false };
  if (!payDate) return none;
  const lines = extractStatusLines(description, payDate);
  if (lines.length === 0) return none;

  const currentLine = selectCurrentLine(lines, payDate);
  if (!currentLine) return none;
  const current = { value: readStatusValue(currentLine.line), line: currentLine.line };

  const curMs = dateOnly(currentLine.date).getTime();
  const payMs = dateOnly(payDate).getTime();
  const others = lines.filter((s) => s !== currentLine && dateOnly(s.date).getTime() <= payMs);

  // Newest first; a later position in the text is the older entry on a tie.
  const older = others
    .filter((s) => dateOnly(s.date).getTime() < curMs || (dateOnly(s.date).getTime() === curMs && s.order > currentLine.order))
    .sort((a, b) => dateOnly(b.date).getTime() - dateOnly(a.date).getTime() || a.order - b.order);

  let previous: DescriptionPercents['previous'] = null;
  let lastHourly: StatusLine | null = null;
  for (const s of older) {
    if (NO_PAY_RE.test(s.line)) continue;
    const v = readStatusValue(s.line);
    // 0% paid nothing, so it is not a "last paid" either.
    if (typeof v === 'number' && v > 0) {
      previous = { value: v, line: s.line };
      break;
    }
    if (v === 'HOURLY' && !lastHourly) lastHourly = s;
  }
  // Nothing but hourly weeks behind this one: last paid by the hour.
  if (!previous && lastHourly) previous = { value: 'HOURLY', line: lastHourly.line };
  return { current, previous, onlyCurrent: others.length === 0 };
}

/**
 * Select the status line for the pay date: exact match first, else the latest
 * line dated on or before the pay date. Parse % / HOURLY / SUBBED / SET RATE
 * from that line only (Round 8 §2.4).
 */
export function suggestPercent(
  description: string,
  payDate: Date | null,
): { value: PercentValue; line: string | null } {
  return readDescriptionPercents(description, payDate).current;
}

/**
 * `Last % Paid` source #1 — last week's Pivot by Job (input #8), `% of
 * Completion` by AP#. Only real percentages count: a job that was HOURLY last
 * week has no "% paid", so the calendar fallback gets its turn instead.
 */
export function lastPaidMapFromPivot(
  rows: readonly { apKey: string; percent: PercentValue }[] | null,
): Map<string, number> {
  const m = new Map<string, number>();
  for (const r of rows ?? []) {
    const key = normalizeApInput(r.apKey);
    // Round 12 — 0% is a week nothing was paid, not a percentage paid.
    if (!key || typeof r.percent !== 'number' || r.percent <= 0) continue;
    // The last row for an AP# wins: the sheet is in row order, so a later row
    // is the more recent statement of what that job was paid to.
    m.set(key, r.percent);
  }
  return m;
}

/** First "(XX)" group of 2–3 uppercase letters in an event title. */
export function extractTitleInitials(title: string): string | null {
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
export function estimatorFromTitle(title: string, known: ReadonlySet<string>): string | null {
  let first: string | null = null;
  for (const m of title.matchAll(TITLE_TOKEN_RE)) {
    const token = m[1].toUpperCase();
    if (first === null) first = token;
    if (known.has(token)) return token;
  }
  return known.size === 0 ? first : null;
}

// ---------------------------------------------------------------------------
// Building the TimeNavi Export rows (columns A–S)
// ---------------------------------------------------------------------------

export interface ReportDeps {
  nvSales: SalesRow[];
  utSales: SalesRow[];
  nvTotals: ReadonlyMap<string, number>;
  utTotals: ReadonlyMap<string, number>;
  /**
   * `Last % Paid` source #1: last week's Pivot by Job `% of Completion` by
   * AP# (see `lastPaidMapFromPivot`). The calendar description is the fallback.
   */
  lastWeekMap: ReadonlyMap<string, number> | null;
  /** apKey → jobcode_2 of the FIRST timesheet row with that key. */
  timesheetJobcodeByApKey: ReadonlyMap<string, string>;
  /** Known estimator initials (palette + configured lists) for the title fallback. */
  knownEstimators: ReadonlySet<string>;
}

/**
 * A row key that survives a re-upload of the calendar export.
 *
 * It used to be `ev${ev.index}:${apKey}`, where `index` is just the row's
 * position in the file (see `parsers.ts`, `index: events.length`). Re-uploading
 * the calendar with one event added, removed or reordered shifted every later
 * key: edits on AP#-bearing rows were silently orphaned, and edits on rows
 * WITHOUT an AP# landed on whatever row had slid into that position. The drift
 * reached payroll too, because calendar-only job columns key off
 * `cal:${reportRow.id}`.
 *
 * The AP# is the real identity of a row, so it is the key when there is one.
 * AP-less rows fall back to a hash of the content that identifies the event to
 * a human — title, start date, calendar.
 *
 * NOTE: the key must come from the FILE's apKey, never the edited one. Keying
 * off the edit would move an edit out from under itself the moment someone
 * corrects a row's AP#.
 */
export function reportRowId(ev: TimeNaviEvent): string {
  if (ev.apKey) return `ap:${ev.apKey}`;
  return `ev:${stableRowHash(ev)}`;
}

/** FNV-1a over the fields that identify an AP-less event to a human. */
function stableRowHash(ev: TimeNaviEvent): string {
  const text = [
    ev.event.trim().toLowerCase(),
    ev.startDate ? ev.startDate.toISOString().slice(0, 10) : '',
    ev.calendar.trim().toLowerCase(),
  ].join('\u0000');
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(36);
}

/**
 * Two calendar events can legitimately share an AP# (one job, two visits), and
 * two AP-less events can be identical in title, date and calendar. Both would
 * collapse onto one key and make the second row un-editable, so repeats get a
 * deterministic `#2`, `#3`… suffix in file order.
 */
export function reportRowIds(events: readonly TimeNaviEvent[]): string[] {
  const seen = new Map<string, number>();
  return events.map((ev) => {
    const base = reportRowId(ev);
    const n = (seen.get(base) ?? 0) + 1;
    seen.set(base, n);
    return n === 1 ? base : `${base}#${n}`;
  });
}

export function buildReportRows(
  events: TimeNaviEvent[],
  edits: EditsMap,
  deps: ReportDeps,
  payDate: Date | null,
): ReportRow[] {
  const nv = buildSalesIndex(deps.nvSales);
  const ut = buildSalesIndex(deps.utSales);
  const ids = reportRowIds(events);

  return events.map((ev, evIndex) => {
    const id = ids[evIndex];
    const edit = edits[id] ?? {};

    const apKey = edit.apKey !== undefined ? edit.apKey : ev.apKey;
    const included = edit.included !== undefined ? edit.included : apKey !== null;

    const desc = readDescriptionPercents(ev.description, payDate);
    const suggestion = desc.current;
    const percent = edit.percent !== undefined ? edit.percent : suggestion.value;

    const nvMatch = apKey ? nv.byApKey.get(apKey) : undefined;
    const utMatch = apKey ? ut.byApKey.get(apKey) : undefined;

    // J: books lookup first; when the books are blank, fall back to the first
    // known-estimator "(XX)" token in the event title (see estimatorFromTitle).
    const booksEstimator = utMatch?.salesRep.trim() || nvMatch?.salesRep.trim() || '';
    const estimator =
      edit.estimator !== undefined
        ? edit.estimator
        : booksEstimator || estimatorFromTitle(ev.event, deps.knownEstimators) || '';

    /**
     * K: `Last % Paid`.
     *
     *   1. This week's line is the only dated line → blank. The job started
     *      this pay week; showing this week's % again as "last paid" was the
     *      duplication bug.
     *   2. Last week's Pivot by Job (input #8) `% of Completion` for this AP#
     *      — the record of what was actually paid.
     *   3. Otherwise the most recent real % below this week's line in the
     *      description, skipping "no pay" weeks — not necessarily last week.
     *   4. Otherwise blank.
     *
     * A typed value on the grid still wins over all of it.
     */
    const lastPaid: PrevPercent | null =
      edit.lastPaid !== undefined
        ? edit.lastPaid
        : desc.onlyCurrent
          ? null
          : ((apKey ? deps.lastWeekMap?.get(apKey) : undefined) ?? desc.previous?.value ?? null);

    const qboJobName =
      edit.qboJobName !== undefined ? edit.qboJobName : (nvMatch?.name2 ?? utMatch?.name2 ?? '');
    const invoiceDate =
      edit.invoiceDate !== undefined ? edit.invoiceDate : nvMatch ? nvMatch.date : utMatch ? utMatch.date : null;

    const inQboTime =
      edit.inQboTime !== undefined
        ? edit.inQboTime
        : (apKey && deps.timesheetJobcodeByApKey.get(apKey)) || 'NotClockedIn';

    const nvTotal = apKey ? deps.nvTotals.get(apKey) : undefined;
    const utTotal = apKey ? deps.utTotals.get(apKey) : undefined;
    const lookedUpFound = nvTotal !== undefined || utTotal !== undefined;
    const jobTotalFound = edit.jobTotal !== undefined ? edit.jobTotal !== '#N/A' : lookedUpFound;
    const matchedTotal =
      edit.jobTotal !== undefined && edit.jobTotal !== '#N/A'
        ? edit.jobTotal
        : (nvTotal ?? utTotal ?? 0);
    const jobTotal: JobTotalValue =
      edit.jobTotal !== undefined ? edit.jobTotal : lookedUpFound ? matchedTotal : '#N/A';

    // Round 8 §2.2 — Apptivo $ is blank until someone types it; no auto-fill
    // from the job-total lookup, so $ Difference stays blank too (and the
    // Round 4 pink flag never fires off a blank pair).
    const apptivoEdited = edit.apptivo !== undefined;
    const apptivo: number | null = apptivoEdited ? (edit.apptivo ?? null) : null;

    const dollarDiff =
      jobTotalFound && apptivo !== null ? matchedTotal - apptivo : null;

    const effectiveStart = edit.startDate !== undefined ? edit.startDate : ev.startDate;
    const dateDiff = invoiceDate && effectiveStart ? dateDiffDays(invoiceDate, effectiveStart) : null;

    return {
      id,
      eventIndex: ev.index,
      included,
      // Round 10 — every displayed column is overridable on the grid.
      event: edit.event !== undefined ? edit.event : ev.event,
      startDate: edit.startDate !== undefined ? edit.startDate : ev.startDate,
      endDate: edit.endDate !== undefined ? edit.endDate : ev.endDate,
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

export { normalizeApInput };

// Round 8 §2.3 — % Completion value helpers live in format.ts; re-exported here
// so grid/export consumers can pull them from either module.
export {
  isPercentNumber,
  isPercentState,
  isPercentText,
  matchesSetRate,
  matchesSubbedOut,
  percentExcelValue,
  percentIsValidState,
  percentText,
  percentTextValue,
} from './format';

// ---------------------------------------------------------------------------
// Round 4 §1.1 — pink review flags (shared by the grid and the export)
// ---------------------------------------------------------------------------

/** `% Completion` gets the pink style on HOURLY text or a numeric > 100% / < 0%. */
export function percentIsFlagged(v: ReportRow['percent']): boolean {
  if (v === 'HOURLY') return true;
  return typeof v === 'number' && (v > 1 || v < 0);
}

/**
 * Round 8 §2.4 — `% Completion` turns solid red (white bold text) when it holds
 * no valid state: blank, 0%, or text that isn't HOURLY / SUBBED / SET RATE.
 */
export function percentIsRedFlagged(v: ReportRow['percent']): boolean {
  return !percentIsValidState(v);
}

/**
 * `In QBO Time?` values shared by 2+ included rows (a duplicate timesheet job
 * mapped to multiple events). NotClockedIn is flagged by its own rule.
 */
export function duplicateInQboValues(rows: readonly ReportRow[]): Set<string> {
  const counts = new Map<string, number>();
  for (const r of rows) {
    if (!r.included) continue;
    const v = r.inQboTime.trim();
    if (!v || v === 'NotClockedIn') continue;
    counts.set(v, (counts.get(v) ?? 0) + 1);
  }
  return new Set([...counts.entries()].filter(([, n]) => n >= 2).map(([v]) => v));
}

/** `$ Difference` gets the pink style on any nonzero value. */
export function dollarDiffIsFlagged(v: number | null): boolean {
  return typeof v === 'number' && v !== 0;
}
