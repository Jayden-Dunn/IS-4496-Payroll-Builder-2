import type { PercentValue } from './types';

/** M/D/YYYY */
export function fmtDate(d: Date | null | undefined): string {
  if (!d) return '';
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

/** $14,411 (rounded to whole dollars for on-screen display) */
export function fmtMoney(v: number): string {
  const r = Math.round(v);
  const s = Math.abs(r).toLocaleString('en-US');
  return (r < 0 ? '-$' : '$') + s;
}

/** 0.95 → "95%", 'HOURLY' → "HOURLY", null → "" */
export function fmtPercent(v: PercentValue): string {
  if (v === null) return '';
  if (v === 'HOURLY') return 'HOURLY';
  return `${Math.round(v * 100)}%`;
}

/** Hours with at most 2 decimals: 4.26, 20.08, 8 */
export function fmtHours(h: number): string {
  return String(Math.round(h * 100) / 100);
}

export function round2(n: number): number {
  return Math.round(n * 100) / 100;
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

/** Parse user input for a % cell: "95", "95%", "0.95", "hourly", "-", "". */
export function parsePercentInput(text: string): PercentValue {
  const t = text.trim();
  if (!t || t === '-') return null;
  if (/hourly/i.test(t)) return 'HOURLY';
  const m = /^(-?\d+(?:\.\d+)?)\s*%?$/.exec(t);
  if (!m) return null;
  const n = Number(m[1]);
  return n > 1 ? n / 100 : n;
}

/** Parse user input for a money cell: "$2,095.00", "2095". Null when invalid. */
export function parseMoneyInput(text: string): number | null {
  const t = text.trim().replace(/[$,\s]/g, '');
  if (!t) return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}
