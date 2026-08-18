import type { JobTotalValue, PercentState, PercentText, PercentValue } from './types';

/** M/D/YYYY */
export function fmtDate(d: Date | null | undefined): string {
  if (!d) return '';
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

/** mm-dd-yy (reference Start/End Date display) */
export function fmtDateShort(d: Date | null | undefined): string {
  if (!d) return '';
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const yy = String(d.getFullYear()).slice(-2);
  return `${mm}-${dd}-${yy}`;
}

/** Accounting-style number without $ (reference Job Total / Apptivo columns). */
export function fmtAccounting(v: number | null | undefined): string {
  if (v === null || v === undefined) return '';
  const abs = Math.abs(v);
  const s = abs.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return v < 0 ? `(${s})` : s;
}

/** Job Total $ — `#N/A` when unmatched, else accounting style. */
export function fmtJobTotal(v: JobTotalValue): string {
  if (v === '#N/A') return '#N/A';
  return fmtAccounting(v);
}

/** $14,411 (legacy whole-dollar display) */
export function fmtMoney(v: number): string {
  const r = Math.round(v);
  const s = Math.abs(r).toLocaleString('en-US');
  return (r < 0 ? '-$' : '$') + s;
}

// ---------------------------------------------------------------------------
// % Completion values (Round 8 §2.3 / §2.4)
// ---------------------------------------------------------------------------

const PERCENT_STATES: readonly PercentState[] = ['HOURLY', 'SUBBED', 'SET RATE'];

/** True for a numeric fraction (0.38 = 38%). */
export function isPercentNumber(v: PercentValue): v is number {
  return typeof v === 'number';
}

/** True for one of the literal states: HOURLY, SUBBED, SET RATE. */
export function isPercentState(v: PercentValue): v is PercentState {
  return typeof v === 'string' && (PERCENT_STATES as readonly string[]).includes(v);
}

/** True for a free-text value the user typed (Round 8 §2.3). */
export function isPercentText(v: PercentValue): v is PercentText {
  return typeof v === 'object' && v !== null && v.kind === 'text';
}

/** The verbatim text of a free-text value; null for every other variant. */
export function percentText(v: PercentValue): string | null {
  return isPercentText(v) ? v.text : null;
}

/**
 * Wrap text verbatim. Use `parsePercentInput` for anything a user typed — it
 * canonicalizes percents and the HOURLY / SUBBED / SET RATE states first.
 */
export function percentTextValue(text: string): PercentText {
  return { kind: 'text', text };
}

/**
 * §2.4 — a cell holds a "valid" state when it has a nonzero %, or reads
 * HOURLY / SUBBED / SET RATE. Blank, 0% and free text are invalid (red).
 */
export function percentIsValidState(v: PercentValue): boolean {
  if (isPercentState(v)) return true;
  return isPercentNumber(v) && v !== 0;
}

/** What a % cell writes into a worksheet: the fraction, or its text. */
export function percentExcelValue(v: PercentValue): string | number {
  if (v === null) return '';
  if (isPercentNumber(v)) return v;
  if (isPercentState(v)) return v;
  return v.text;
}

/** 0.95 → "95%", 'HOURLY' → "HOURLY", free text verbatim, null → "" */
export function fmtPercent(v: PercentValue): string {
  if (v === null) return '';
  if (isPercentState(v)) return v;
  if (isPercentText(v)) return v.text;
  return `${Math.round(v * 100)}%`;
}

/**
 * Lowercase, drop punctuation/hyphens and collapse whitespace so that
 * "Subbed-out", "SUBBED  OUT" and "set rate." all reduce to a plain phrase.
 */
export function normalizeStateText(text: string): string {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

/** §2.4 — forgiving "subbed out" match (also plain "subbed"). */
export function matchesSubbedOut(text: string): boolean {
  return /\bsubbed\b/.test(normalizeStateText(text));
}

/** §2.4 — forgiving "set rate" match. */
export function matchesSetRate(text: string): boolean {
  return /\bset rate\b/.test(normalizeStateText(text));
}

/** Hours with at most 2 decimals: 4.26, 20.08, 8 */
export function fmtHours(h: number): string {
  return String(Math.round(h * 100) / 100);
}

/** Pivot hours always show two decimals: 4.26 → "4.26", 8 → "8.00" */
export function fmtHoursFixed(h: number): string {
  return h.toFixed(2);
}

export function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

/**
 * Perceived-luminance check for a 6-digit hex fill (no '#'). Dark fills
 * (e.g. ND's 4472C4) get white text; light fills get black — computed, so any
 * palette edit in Settings keeps readable text automatically.
 */
export function isDarkColor(hex6: string): boolean {
  const r = parseInt(hex6.slice(0, 2), 16);
  const g = parseInt(hex6.slice(2, 4), 16);
  const b = parseInt(hex6.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b < 140;
}

/** Strip the time-of-day portion (local). */
export function dateOnly(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** Whole days a − b (may be negative). */
export function dateDiffDays(a: Date, b: Date): number {
  const ms = dateOnly(a).getTime() - dateOnly(b).getTime();
  return Math.round(ms / 86_400_000);
}

export function addDays(d: Date, n: number): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

/** "2026-07-10" → local Date, or null. */
export function isoToDate(s: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s.trim());
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** Local Date → "2026-07-10" for <input type="date">. */
export function dateToIso(d: Date): string {
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
export function parsePercentInput(text: string): PercentValue {
  const t = text.trim();
  if (!t || t === '-') return null;

  const m = /^(-?(?:\d+(?:\.\d+)?|\.\d+))\s*(%?)$/.exec(t);
  if (m) {
    const n = Number(m[1]);
    if (m[2] === '%') return n / 100;
    return n > 1 ? n / 100 : n;
  }

  if (/hourly/i.test(t)) return 'HOURLY';
  if (matchesSubbedOut(t)) return 'SUBBED';
  if (matchesSetRate(t)) return 'SET RATE';
  return { kind: 'text', text: t };
}

/** Parse user input for a money cell: "$2,095.00", "2095". Null when invalid/blank. */
export function parseMoneyInput(text: string): number | null {
  const t = text.trim().replace(/[$,\s]/g, '');
  if (!t) return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}
