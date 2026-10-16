/**
 * Crash recovery for a working session.
 *
 * Uploads and manual edits live in React state, so until now an accidental
 * refresh, a crash, a browser restart or Chrome discarding a backgrounded tab
 * took hours of hand-corrections with it. This caches the whole working set in
 * the browser so the next load picks up where the last one left off.
 *
 * Two decisions worth stating:
 *
 * 1. IndexedDB, not localStorage. The parsed uploads are full of `Date`s and
 *    `Map`s (`nvTotals`, `lastWeek`, `pl.byJob`, `customerNameByAp`, every
 *    timesheet row's `localDate`). IndexedDB stores them by structured clone,
 *    so they come back as real `Date`s and `Map`s. A JSON round-trip would need
 *    a hand-written reviver for a dozen nested shapes, and a missed one fails
 *    late and quietly — a `Date` that came back a string renders fine in a cell
 *    and then throws inside a builder. localStorage is also ~5 MB and
 *    synchronous, which a week's timesheets will not fit.
 *
 * 2. Nothing leaves the machine. This is browser-local only — not the shared
 *    `/api/settings` blob, which every client of a deployment would clobber,
 *    and which would put timesheets and wage rates on a server.
 *
 * Three stores, split by how often they are written: `sessions` (the envelope —
 * edit maps and flags, small, written on a debounce), `slots` (one record per
 * upload, so re-uploading one file rewrites one record) and `files` (the raw
 * bytes, so a parser change in a later deploy can re-read the uploads instead
 * of throwing the session away).
 */
import type { SlotChip } from '../components/UploadSlot';

export const SESSION_DB_NAME = 'pct-completion-builder';
export const SESSION_DB_VERSION = 1;
/** Envelope shape. Bump to invalidate stored sessions outright. */
export const SESSION_PAYLOAD_VERSION = 1;
/**
 * Parsed-payload shape. Bump when a parser's OUTPUT changes; a mismatch
 * re-parses from the stored bytes and KEEPS the edits, rather than discarding
 * the session at exactly the moment it is most needed.
 */
export const SESSION_PARSER_VERSION = 1;

const POINTER_KEY = 'pct-completion-builder.session-pointer.v1';
const STORE_SESSIONS = 'sessions';
const STORE_SLOTS = 'slots';
const STORE_FILES = 'files';

export interface SessionSlotMeta {
  slot: string;
  fileName: string;
  fileSize: number;
  fileLastModified: number;
  uploadedAt: string;
  chip: SlotChip;
}

/** Everything the App layer restores that is not a parsed upload. */
export interface SessionEnvelope {
  id: string;
  payloadVersion: number;
  parserVersion: number;
  savedAt: string;
  payDateIso: string;
  createdPercents: boolean;
  createdPayroll: boolean;
  mode: string;
  tab: string;
  slots: SessionSlotMeta[];
  /** The six edit maps, stored as-is (sparse and small). */
  edits: Record<string, unknown>;
}

export interface StoredSlot {
  sessionId: string;
  slot: string;
  meta: SessionSlotMeta;
  /** That slot's `Partial<ParsedData>` patch, stored as live objects. */
  parsed: Record<string, unknown>;
}

export interface SessionPointer {
  id: string;
  savedAt: string;
  label: string;
  slotCount: number;
}

export interface RestoredSession {
  envelope: SessionEnvelope;
  /** Merged patches of every slot that read back cleanly. */
  parsed: Record<string, unknown>;
  chips: Record<string, SlotChip>;
  /** Slots that were stored but could not be read — surfaced, never hidden. */
  degradedSlots: string[];
}

// ---------------------------------------------------------------------------
// Pointer — a synchronous read, so the first paint already knows a session
// exists and can say "restoring…" instead of flashing the empty placeholder.
// ---------------------------------------------------------------------------

export function loadSessionPointer(): SessionPointer | null {
  try {
    const raw = localStorage.getItem(POINTER_KEY);
    if (!raw) return null;
    const v = JSON.parse(raw) as Partial<SessionPointer>;
    if (typeof v?.id !== 'string' || v.id === '') return null;
    return {
      id: v.id,
      savedAt: typeof v.savedAt === 'string' ? v.savedAt : '',
      label: typeof v.label === 'string' ? v.label : '',
      slotCount: typeof v.slotCount === 'number' ? v.slotCount : 0,
    };
  } catch {
    return null;
  }
}

export function saveSessionPointer(p: SessionPointer): void {
  try {
    localStorage.setItem(POINTER_KEY, JSON.stringify(p));
  } catch {
    // Storage can be denied (private windows); the session simply isn't cached.
  }
}

export function clearSessionPointer(): void {
  try {
    localStorage.removeItem(POINTER_KEY);
  } catch {
    // ignore
  }
}

// ---------------------------------------------------------------------------
// IndexedDB plumbing. Nothing here may throw into render.
// ---------------------------------------------------------------------------

export function sessionStorageAvailable(): boolean {
  try {
    return typeof indexedDB !== 'undefined' && indexedDB !== null;
  } catch {
    return false;
  }
}

function openDb(): Promise<IDBDatabase | null> {
  return new Promise((resolve) => {
    if (!sessionStorageAvailable()) {
      resolve(null);
      return;
    }
    let req: IDBOpenDBRequest;
    try {
      req = indexedDB.open(SESSION_DB_NAME, SESSION_DB_VERSION);
    } catch {
      resolve(null);
      return;
    }
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_SESSIONS)) db.createObjectStore(STORE_SESSIONS, { keyPath: 'id' });
      if (!db.objectStoreNames.contains(STORE_SLOTS)) {
        db.createObjectStore(STORE_SLOTS, { keyPath: ['sessionId', 'slot'] });
      }
      if (!db.objectStoreNames.contains(STORE_FILES)) {
        db.createObjectStore(STORE_FILES, { keyPath: ['sessionId', 'slot'] });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => resolve(null);
    req.onblocked = () => resolve(null);
  });
}

function run<T>(store: IDBObjectStore, op: (s: IDBObjectStore) => IDBRequest<T>): Promise<T | null> {
  return new Promise((resolve) => {
    try {
      const req = op(store);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function withStore<T>(
  name: string,
  mode: IDBTransactionMode,
  fn: (store: IDBObjectStore) => Promise<T | null>,
): Promise<T | null> {
  const db = await openDb();
  if (!db) return null;
  try {
    const tx = db.transaction(name, mode);
    const result = await fn(tx.objectStore(name));
    return result;
  } catch {
    return null;
  } finally {
    try {
      db.close();
    } catch {
      // ignore
    }
  }
}

// ---------------------------------------------------------------------------
// Validators — settings.ts convention: parseX(raw) returns null, never throws
// ---------------------------------------------------------------------------

export function parseSessionEnvelope(raw: unknown): SessionEnvelope | null {
  if (typeof raw !== 'object' || raw === null) return null;
  const v = raw as Record<string, unknown>;
  if (typeof v.id !== 'string' || v.id === '') return null;
  if (v.payloadVersion !== SESSION_PAYLOAD_VERSION) return null;
  if (typeof v.edits !== 'object' || v.edits === null) return null;
  if (!Array.isArray(v.slots)) return null;
  return {
    id: v.id,
    payloadVersion: SESSION_PAYLOAD_VERSION,
    parserVersion: typeof v.parserVersion === 'number' ? v.parserVersion : 0,
    savedAt: typeof v.savedAt === 'string' ? v.savedAt : '',
    payDateIso: typeof v.payDateIso === 'string' ? v.payDateIso : '',
    createdPercents: v.createdPercents === true,
    createdPayroll: v.createdPayroll === true,
    mode: v.mode === 'payroll' ? 'payroll' : 'percents',
    tab: typeof v.tab === 'string' ? v.tab : '',
    slots: v.slots.filter((s): s is SessionSlotMeta => typeof s === 'object' && s !== null),
    edits: v.edits as Record<string, unknown>,
  };
}

/**
 * A spot check, not a deep validation. Structured clone preserves kinds, so the
 * realistic failure is a record written by an older build with a different
 * shape — which the first row reveals. Walking thousands of timesheet rows on
 * every boot would cost more than it catches, and `parserVersion` covers the
 * rest. Stated plainly because it is a real gap, not an airtight check.
 */
export function storedSlotLooksSane(slot: StoredSlot | null): slot is StoredSlot {
  if (!slot || typeof slot !== 'object') return false;
  if (typeof slot.slot !== 'string' || typeof slot.parsed !== 'object' || slot.parsed === null) return false;
  for (const value of Object.values(slot.parsed)) {
    if (value === null || value === undefined) continue;
    if (value instanceof Map || Array.isArray(value) || typeof value === 'object') continue;
    return false;
  }
  return true;
}

/** An invalid Date clones as an invalid Date and goes NaN deep inside a builder. */
export function hasBrokenDate(value: unknown, depth = 0): boolean {
  if (depth > 4) return false;
  if (value instanceof Date) return Number.isNaN(value.getTime());
  if (Array.isArray(value)) return value.slice(0, 50).some((v) => hasBrokenDate(v, depth + 1));
  if (value instanceof Map) {
    let n = 0;
    for (const v of value.values()) {
      if (n++ > 50) break;
      if (hasBrokenDate(v, depth + 1)) return true;
    }
    return false;
  }
  if (typeof value === 'object' && value !== null) {
    return Object.values(value as Record<string, unknown>)
      .slice(0, 50)
      .some((v) => hasBrokenDate(v, depth + 1));
  }
  return false;
}

// ---------------------------------------------------------------------------
// Read / write
// ---------------------------------------------------------------------------

export async function loadSession(id: string): Promise<RestoredSession | null> {
  const rawEnvelope = await withStore(STORE_SESSIONS, 'readonly', (s) => run(s, (st) => st.get(id)));
  const envelope = parseSessionEnvelope(rawEnvelope);
  if (!envelope) return null;

  const rawSlots = await withStore(STORE_SLOTS, 'readonly', (s) => run<StoredSlot[]>(s, (st) => st.getAll()));
  const mine = (rawSlots ?? []).filter((r) => r && r.sessionId === id);

  const parsed: Record<string, unknown> = {};
  const chips: Record<string, SlotChip> = {};
  const degradedSlots: string[] = [];
  for (const meta of envelope.slots) {
    const record = mine.find((r) => r.slot === meta.slot) ?? null;
    if (!storedSlotLooksSane(record) || hasBrokenDate(record.parsed)) {
      degradedSlots.push(meta.slot);
      continue;
    }
    Object.assign(parsed, record.parsed);
    chips[meta.slot] = meta.chip;
  }
  return { envelope, parsed, chips, degradedSlots };
}

export async function saveSessionEnvelope(env: SessionEnvelope): Promise<'ok' | 'stale' | 'failed'> {
  const db = await openDb();
  if (!db) return 'failed';
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(STORE_SESSIONS, 'readwrite');
      const store = tx.objectStore(STORE_SESSIONS);
      const get = store.get(env.id);
      get.onsuccess = () => {
        const existing = parseSessionEnvelope(get.result);
        // Compare-then-write: a second tab on the same session must not
        // silently overwrite the one the user is actually typing in.
        if (existing && existing.savedAt > env.savedAt) {
          resolve('stale');
          return;
        }
        store.put(env);
      };
      tx.oncomplete = () => {
        db.close();
        resolve('ok');
      };
      tx.onerror = () => {
        db.close();
        resolve('failed');
      };
      tx.onabort = () => {
        db.close();
        resolve('failed');
      };
    } catch {
      try {
        db.close();
      } catch {
        // ignore
      }
      resolve('failed');
    }
  });
}

export async function saveSessionSlot(
  sessionId: string,
  slot: string,
  parsed: Record<string, unknown>,
  meta: SessionSlotMeta,
  bytes?: ArrayBuffer,
): Promise<boolean> {
  const okSlot = await withStore(STORE_SLOTS, 'readwrite', (s) =>
    run(s, (st) => st.put({ sessionId, slot, meta, parsed } satisfies StoredSlot)),
  );
  if (bytes) {
    // Best effort: losing the raw bytes costs a re-parse path after a deploy,
    // not the session, so a quota failure here is not worth failing the save.
    await withStore(STORE_FILES, 'readwrite', (s) => run(s, (st) => st.put({ sessionId, slot, bytes })));
  }
  return okSlot !== null;
}

export async function clearSessionSlot(sessionId: string, slot: string): Promise<void> {
  await withStore(STORE_SLOTS, 'readwrite', (s) => run(s, (st) => st.delete([sessionId, slot])));
  await withStore(STORE_FILES, 'readwrite', (s) => run(s, (st) => st.delete([sessionId, slot])));
}

export async function deleteSession(id: string): Promise<void> {
  await withStore(STORE_SESSIONS, 'readwrite', (s) => run(s, (st) => st.delete(id)));
  const rawSlots = await withStore(STORE_SLOTS, 'readonly', (s) => run<StoredSlot[]>(s, (st) => st.getAll()));
  for (const r of rawSlots ?? []) {
    if (r?.sessionId === id) await clearSessionSlot(id, r.slot);
  }
  clearSessionPointer();
}

export function newSessionId(): string {
  try {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return crypto.randomUUID();
  } catch {
    // ignore
  }
  return `s${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
}
