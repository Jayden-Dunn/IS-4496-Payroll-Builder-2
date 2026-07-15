/**
 * The canonical job key. Every join in this app is on the AP number.
 * apKeys are ALWAYS compared as strings (never numbers) so suffixed keys
 * like "21848-A5" survive round trips between sources.
 */

export const AP_RE = /AP\s*-?\s*(\d{4,6}(?:-[A-Za-z0-9]+)*)/i;

/** Extract the AP key from free text (event title, jobcode_2, QBO job name…). */
export function extractApKey(text: string | null | undefined): string | null {
  if (!text) return null;
  const m = AP_RE.exec(text);
  return m ? m[1].toUpperCase().trim() : null;
}

/**
 * apKey for a timesheet jobcode_2. Office entries (e.g. "OFFICE [Please put
 * AP# in notes]") map to the special key "OFFICE".
 */
export function timesheetApKey(jobcode2: string): string | null {
  const t = jobcode2.trim();
  if (/^office/i.test(t)) return 'OFFICE';
  return extractApKey(t);
}

/**
 * Normalize a user-typed AP# ("AP24314", "24314", "21848-a5", "-", "").
 * Returns null when the field should be treated as empty.
 */
export function normalizeApInput(raw: string): string | null {
  const t = raw.trim().toUpperCase();
  if (!t || t === '-') return null;
  if (t === 'OFFICE') return 'OFFICE';
  const viaRe = extractApKey(t);
  if (viaRe) return viaRe;
  if (/^\d{1,6}(?:-[A-Z0-9]+)*$/.test(t)) return t;
  // Fall back to the raw uppercased text so odd keys still compare equal.
  return t;
}
