/**
 * Round 8 §4.1 — cell-fill helpers for the A6A6A6 "exclude from payroll" gray.
 *
 * A human painting a row in Excel almost never produces a literal ARGB: the
 * swatch they pick ("White, Background 1, Darker 35%") is stored as a THEME
 * reference plus a tint, so a string compare against `FFA6A6A6` misses every
 * hand-colored row. Resolve theme + tint to RGB first, then compare with a
 * small per-channel tolerance.
 */
import type { CellFill } from './types';

/** The exclusion swatch as ExcelJS writes it. */
export const EXCLUSION_GRAY_ARGB = 'FFA6A6A6';

/** Per-channel slack when comparing resolved colors (rounding differs by tool). */
const CHANNEL_TOLERANCE = 2;

/**
 * Office default theme colors in the order Excel indexes them from a style
 * (`theme="0"` = Background 1 / window). A workbook with a custom theme can
 * differ, which is why `isExclusionGray` also accepts a literal RGB match.
 */
const DEFAULT_THEME_RGB: readonly string[] = [
  'FFFFFF', // 0 — Background 1 (window)
  '000000', // 1 — Text 1
  'EEECE1', // 2 — Background 2
  '1F497D', // 3 — Text 2
  '4F81BD', // 4 — Accent 1
  'C0504D', // 5 — Accent 2
  '9BBB59', // 6 — Accent 3
  '8064A2', // 7 — Accent 4
  '4BACC6', // 8 — Accent 5
  'F79646', // 9 — Accent 6
  '0000FF', // 10 — Hyperlink
  '800080', // 11 — Followed hyperlink
];

export interface Rgb {
  r: number;
  g: number;
  b: number;
}

/** "FFA6A6A6" / "A6A6A6" / "#a6a6a6" → channels; null when unparseable. */
export function parseRgbHex(value: string | null | undefined): Rgb | null {
  if (!value) return null;
  const hex = value.trim().replace(/^#/, '');
  const six = hex.length === 8 ? hex.slice(2) : hex;
  if (!/^[0-9a-fA-F]{6}$/.test(six)) return null;
  return {
    r: parseInt(six.slice(0, 2), 16),
    g: parseInt(six.slice(2, 4), 16),
    b: parseInt(six.slice(4, 6), 16),
  };
}

function toHex(rgb: Rgb): string {
  const part = (n: number): string =>
    Math.max(0, Math.min(255, Math.round(n))).toString(16).toUpperCase().padStart(2, '0');
  return `${part(rgb.r)}${part(rgb.g)}${part(rgb.b)}`;
}

function rgbToHsl(rgb: Rgb): { h: number; s: number; l: number } {
  const r = rgb.r / 255;
  const g = rgb.g / 255;
  const b = rgb.b / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
  else if (max === g) h = ((b - r) / d + 2) / 6;
  else h = ((r - g) / d + 4) / 6;
  return { h, s, l };
}

function hueToChannel(p: number, q: number, t: number): number {
  let x = t;
  if (x < 0) x += 1;
  if (x > 1) x -= 1;
  if (x < 1 / 6) return p + (q - p) * 6 * x;
  if (x < 1 / 2) return q;
  if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
  return p;
}

function hslToRgb(h: number, s: number, l: number): Rgb {
  if (s === 0) return { r: l * 255, g: l * 255, b: l * 255 };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: hueToChannel(p, q, h + 1 / 3) * 255,
    g: hueToChannel(p, q, h) * 255,
    b: hueToChannel(p, q, h - 1 / 3) * 255,
  };
}

/**
 * ECMA-376 tint: shift the HSL luminance of a theme color.
 * `tint < 0` darkens (L × (1 + tint)), `tint > 0` lightens.
 */
export function applyTint(base: Rgb, tint: number): Rgb {
  if (!tint) return base;
  const { h, s, l } = rgbToHsl(base);
  const clamped = Math.max(-1, Math.min(1, tint));
  const lum = clamped < 0 ? l * (1 + clamped) : l * (1 - clamped) + clamped;
  return hslToRgb(h, s, Math.max(0, Math.min(1, lum)));
}

/**
 * Theme index + tint → 6-digit hex. Theme 0 (white) at tint −0.35 resolves to
 * A6A6A6 — the reference workbook's exclusion gray. Null for unknown indexes.
 */
export function resolveThemeTintToRgb(
  theme: number | null | undefined,
  tint: number | null | undefined,
): string | null {
  if (theme === null || theme === undefined || !Number.isInteger(theme)) return null;
  const base = parseRgbHex(DEFAULT_THEME_RGB[theme]);
  if (!base) return null;
  return toHex(applyTint(base, tint ?? 0));
}

/** Both encodings of a fill collapsed to 6-digit hex candidates. */
export function fillRgbCandidates(fill: CellFill | null | undefined): string[] {
  if (!fill) return [];
  const out: string[] = [];
  const literal = parseRgbHex(fill.rgb);
  if (literal) out.push(toHex(literal));
  const themed = resolveThemeTintToRgb(fill.theme, fill.tint);
  if (themed) out.push(themed);
  return out;
}

/** True when two 6-digit hex colors match within ±2 per channel. */
export function rgbNearlyEqual(a: string, b: string): boolean {
  const x = parseRgbHex(a);
  const y = parseRgbHex(b);
  if (!x || !y) return false;
  return (
    Math.abs(x.r - y.r) <= CHANNEL_TOLERANCE &&
    Math.abs(x.g - y.g) <= CHANNEL_TOLERANCE &&
    Math.abs(x.b - y.b) <= CHANNEL_TOLERANCE
  );
}

/**
 * Round 8 §4.1 — is this the "do not send to payroll" gray? Accepts the
 * literal `FFA6A6A6` and the theme-0 / tint −0.35 encoding Excel writes when
 * a human picks the swatch.
 */
export function isExclusionGray(fill: CellFill | null | undefined): boolean {
  return fillRgbCandidates(fill).some((hex) => rgbNearlyEqual(hex, EXCLUSION_GRAY_ARGB));
}
