/**
 * "Update existing report" — work out which MANUAL EDITS a re-uploaded input
 * would overwrite, and nothing else.
 *
 * The rule the review obeys: ask about a value only when the user hand-edited
 * it AND the new file would have produced something different than the old one
 * did. A value the new file changes that nobody touched is applied silently; an
 * edit sitting on a value the new file did not move is left alone silently.
 *
 * Getting at "what the file would have produced" takes one trick. The builders
 * collapse the computed value and the edit into a single result
 * (`edit.percent !== undefined ? edit.percent : suggestion.value`), so there is
 * no computed layer to read back. But the builders are pure, so running them a
 * second time with an EMPTY edits map re-exposes it. The App layer does exactly
 * that, and hands both shadow models here:
 *
 *     defaultsBefore = collect(build(parsed,                        {}))
 *     defaultsAfter  = collect(build({ ...parsed, ...pending },     {}))
 *
 * A conflict is then `editedKeys ∩ { k : before[k] !== after[k] }`.
 *
 * Values are compared as formatted STRINGS. That is deliberate: it is what the
 * user sees in the cell, it sidesteps float noise (18.000000000000004), and it
 * is what the review has to print anyway.
 */
import { fmtAccounting, round2 } from './format';
import type {
  EditsMap,
  PayrollEditsMap,
  PayrollPreviewModel,
  PbjEditsMap,
  PercentValue,
  PivotByJobModel,
  PrevPercent,
  QbHoursMap,
  ReportRow,
} from './types';

/** `${editKey}\u0000${field}` → the value the INPUTS alone produce, formatted. */
export type DefaultsMap = Map<string, string>;

export interface Conflict {
  /** The key in its edits map — what gets deleted when the user takes the new value. */
  editKey: string;
  /** The field inside that edit object. */
  field: string;
  /** Which map the key belongs to, so the resolver knows where to apply it. */
  scope: ConflictScope;
  /** "AP24379 · CCSD Earl Jenkins ES Int" */
  rowLabel: string;
  /** "% Completion" */
  fieldLabel: string;
  /** What the user typed — kept when they choose "keep mine". */
  manualValue: string;
  /** What the old input produced under that edit. */
  oldValue: string;
  /** What the new input produces — taken when they choose "take the new value". */
  newValue: string;
}

export type ConflictScope = 'timenavi' | 'pivot' | 'qbHours' | 'payroll';

/** Round 11 — the web-app tab each scope's cells live on, as the tab strip names it. */
export const SCOPE_TAB: Record<ConflictScope, string> = {
  timenavi: 'TimeNavi Export',
  pivot: 'Pivot by Job',
  qbHours: 'Pivot by Job',
  payroll: 'Payroll Preview',
};

/**
 * Round 11 — one value the new input moved on a cell nobody had edited. Kept
 * whole (not just counted) so the summary after an update can say exactly
 * what changed and where.
 */
export interface ValueChange {
  scope: ConflictScope;
  editKey: string;
  field: string;
  rowLabel: string;
  fieldLabel: string;
  oldValue: string;
  newValue: string;
}

/** An edit whose row is not in the new file at all — shown, never gated. */
export interface OrphanedEdit {
  editKey: string;
  scope: ConflictScope;
  rowLabel: string;
  /** "% Completion 45%, WO Override $1,200" */
  summary: string;
}

export interface UpdateReview {
  conflicts: Conflict[];
  orphans: OrphanedEdit[];
  /** How many computed values the new input moves that nobody had edited. */
  silentChangeCount: number;
  /** Round 11 — those same values, one entry each (`length === silentChangeCount`). */
  silentChanges: ValueChange[];
}

const SEP = '\u0000';
const keyOf = (editKey: string, field: string): string => `${editKey}${SEP}${field}`;

// ---------------------------------------------------------------------------
// Formatting — one place, so `before` and `after` can never disagree on shape
// ---------------------------------------------------------------------------

export function fmtPercentValue(v: PercentValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return `${round2(v * 100)}%`;
  if (typeof v === 'string') return v;
  return v.text ?? '';
}

function fmtPrevPercent(v: PrevPercent | null): string {
  if (v === null) return '';
  return typeof v === 'number' ? `${round2(v * 100)}%` : String(v);
}

function fmtDate(d: Date | null): string {
  if (!d || Number.isNaN(d.getTime())) return '';
  return d.toISOString().slice(0, 10);
}

function fmtNum(n: number | null | undefined): string {
  if (n === null || n === undefined) return '';
  return String(round2(n));
}

function fmtMoney(n: number | null | undefined): string {
  if (n === null || n === undefined) return '';
  return fmtAccounting(n);
}

// ---------------------------------------------------------------------------
// Collectors — one per model, covering exactly that model's editable fields
// ---------------------------------------------------------------------------

/** Column labels as the TimeNavi grid heads them. */
const TIMENAVI_FIELD_LABELS: Record<string, string> = {
  apKey: 'AP#',
  percent: '% Completion',
  estimator: 'Estimator',
  apptivo: 'Apptivo $',
  included: 'Included',
  event: 'Event',
  startDate: 'Start Date',
  endDate: 'End Date',
  calendar: 'Calendar',
  attendees: 'Attendees',
  location: 'Location',
  description: 'Description',
  lastPaid: 'Last % Paid',
  qboJobName: 'QBO Job Name',
  inQboTime: 'In QBO Time?',
  jobTotal: 'Job Total $',
  invoiceDate: 'Invoice Date',
};

export function collectReportDefaults(rows: readonly ReportRow[]): DefaultsMap {
  const m: DefaultsMap = new Map();
  for (const r of rows) {
    const put = (field: string, value: string): void => {
      m.set(keyOf(r.id, field), value);
    };
    put('apKey', r.apKey ?? '');
    put('percent', fmtPercentValue(r.percent));
    put('estimator', r.estimator);
    put('apptivo', fmtNum(r.apptivo));
    put('included', r.included ? 'yes' : 'no');
    put('event', r.event);
    put('startDate', fmtDate(r.startDate));
    put('endDate', fmtDate(r.endDate));
    put('calendar', r.calendar);
    put('attendees', r.attendees);
    put('location', r.location);
    put('description', r.description);
    put('lastPaid', fmtPrevPercent(r.lastPaid));
    put('qboJobName', r.qboJobName);
    put('inQboTime', r.inQboTime);
    put('jobTotal', typeof r.jobTotal === 'number' ? fmtNum(r.jobTotal) : String(r.jobTotal));
    put('invoiceDate', fmtDate(r.invoiceDate));
  }
  return m;
}

/** Row labels for the review, keyed the same way as the defaults. */
export function collectReportLabels(rows: readonly ReportRow[]): Map<string, string> {
  const m = new Map<string, string>();
  for (const r of rows) m.set(r.id, [r.apKey, r.event].filter(Boolean).join(' · ') || r.event || r.id);
  return m;
}

const PIVOT_FIELD_LABELS: Record<string, string> = {
  apKey: 'AP#',
  crewLead: 'Crew Lead',
  woOverride: 'WO Override',
  woPct: 'WO%',
  percent: '% of Completion',
  notes: 'Notes',
  dbOverride: 'Davis-Bacon override',
  excludedFromPayroll: 'Keep off payroll',
  customer: 'Customer',
  job: 'Job',
  estimator: 'Estimator',
  qbHours: 'QB Hrs',
};

export function collectPivotDefaults(model: PivotByJobModel): DefaultsMap {
  const m: DefaultsMap = new Map();
  const putJob = (
    id: string,
    a: {
      apKey: string | null;
      crewLead: string;
      woOverride: number | null;
      woPct: number | null;
      percent: PercentValue;
      notes: string;
    },
  ): void => {
    m.set(keyOf(id, 'apKey'), a.apKey ?? '');
    m.set(keyOf(id, 'crewLead'), a.crewLead);
    m.set(keyOf(id, 'woOverride'), fmtMoney(a.woOverride));
    m.set(keyOf(id, 'woPct'), a.woPct === null ? '' : `${round2(a.woPct * 100)}%`);
    m.set(keyOf(id, 'percent'), fmtPercentValue(a.percent));
    m.set(keyOf(id, 'notes'), a.notes);
  };
  for (const c of model.customers) {
    for (const j of c.jobs) {
      // `livePercent`, not `percent`: `percent` already folds in the edit on a
      // non-shadow build, and on a shadow build the two are equal anyway.
      putJob(j.id, { ...j.analysis, percent: j.analysis.livePercent });
      m.set(keyOf(j.id, 'dbOverride'), j.analysis.dbPatternMatch ? 'DB' : 'not DB');
      for (const e of j.employees) {
        m.set(keyOf(`${j.id}${SEP}${e.username}`, 'qbHours'), fmtNum(e.qbHours ?? e.total));
      }
    }
  }
  for (const r of [...model.balances, ...model.utPayroll]) {
    putJob(r.id, { ...r, percent: r.livePercent });
    m.set(keyOf(r.id, 'customer'), r.customer);
    m.set(keyOf(r.id, 'job'), r.job);
    m.set(keyOf(r.id, 'estimator'), r.estimator);
    m.set(keyOf(r.id, 'excludedFromPayroll'), r.excludedFromPayroll ? 'off payroll' : 'on payroll');
  }
  return m;
}

export function collectPivotLabels(model: PivotByJobModel): Map<string, string> {
  const m = new Map<string, string>();
  for (const c of model.customers) {
    for (const j of c.jobs) {
      const label = [j.analysis.apKey, `${c.customer} ${j.jobcode2}`.trim()].filter(Boolean).join(' · ');
      m.set(j.id, label || j.id);
      for (const e of j.employees) m.set(`${j.id}${SEP}${e.username}`, `${label} · ${e.username}`);
    }
  }
  for (const r of [...model.balances, ...model.utPayroll]) {
    m.set(r.id, [r.apKey, `${r.customer} ${r.job}`.trim()].filter(Boolean).join(' · ') || r.id);
  }
  return m;
}

const PAYROLL_FIELD_LABELS: Record<string, string> = {
  hours: 'hrs',
  rate: 'rate',
  taxes: '(-) taxes',
  nest: 'NEST [IRA]',
  net: 'net',
  woPct: 'profit share %',
  workOrderOverride: 'Work Order Override',
  percentCompletion: '% Completion',
  laborWages: '- Labor + Wages',
  mileage: '-Mileage',
  halfMaterials: 'Materials',
  income: 'Total Income',
  materials: 'Materials (books)',
  mileageNew: 'Mileage - New',
  hourlyPay: 'Hourly Pay',
  reimbursements: 'Reimbursements/additional pay',
  deductions: 'Deductions/Sundries',
  miles: 'miles',
  amount: 'amount',
};

export function collectPayrollDefaults(model: PayrollPreviewModel): DefaultsMap {
  const m: DefaultsMap = new Map();
  const hourlyBlock = (b: { key: string; lines: readonly { id: string; hours: number; rate: number }[] }): void => {
    for (const l of b.lines) {
      m.set(keyOf(`hl:${l.id}`, 'hours'), fmtNum(l.hours));
      m.set(keyOf(`hl:${l.id}`, 'rate'), fmtNum(l.rate));
    }
  };
  for (const b of model.painters) hourlyBlock(b);
  for (const b of model.office) hourlyBlock(b);
  for (const entry of model.painterSequence) {
    if (entry.kind === 'hourly') hourlyBlock(entry.block);
    else if (entry.kind === 'ps') {
      const b = entry.block;
      if (b.tsheets) hourlyBlock(b.tsheets);
      for (const j of b.jobs) {
        const k = `ps:${b.key}:${j.jobRowId}`;
        m.set(keyOf(k, 'woPct'), j.woPct === null ? '' : `${round2(j.woPct * 100)}%`);
        m.set(keyOf(k, 'workOrderOverride'), fmtMoney(j.workOrderOverride));
        m.set(keyOf(k, 'percentCompletion'), fmtPercentValue(j.percent));
        m.set(keyOf(k, 'laborWages'), fmtMoney(j.laborWages));
        m.set(keyOf(k, 'mileage'), fmtMoney(j.mileage));
        m.set(keyOf(k, 'halfMaterials'), fmtMoney(j.halfMaterials));
        m.set(keyOf(k, 'income'), fmtMoney(j.income));
        m.set(keyOf(k, 'materials'), fmtMoney(j.materials));
      }
      const bk = `psb:${b.key}`;
      m.set(keyOf(bk, 'mileageNew'), fmtMoney(b.mileageNew));
      m.set(keyOf(bk, 'miles'), fmtNum(b.mileageMiles));
      m.set(keyOf(bk, 'hourlyPay'), fmtMoney(b.hourlyPay));
      for (const x of b.extras) m.set(keyOf(x.key, 'amount'), fmtMoney(x.amount));
    } else {
      const b = entry.block;
      m.set(keyOf(`sal:${b.name}`, 'miles'), fmtNum(b.miles));
      m.set(keyOf(`sal:${b.name}`, 'reimbursements'), fmtMoney(b.reimbursements));
      for (const x of b.extras) m.set(keyOf(x.key, 'amount'), fmtMoney(x.amount));
    }
  }
  for (const b of [...model.admin, ...model.estimators, ...model.managers]) {
    m.set(keyOf(`sal:${b.name}`, 'miles'), fmtNum(b.miles));
    m.set(keyOf(`sal:${b.name}`, 'reimbursements'), fmtMoney(b.reimbursements));
    for (const x of b.extras) m.set(keyOf(x.key, 'amount'), fmtMoney(x.amount));
  }
  return m;
}

export function collectPayrollLabels(model: PayrollPreviewModel): Map<string, string> {
  const m = new Map<string, string>();
  const hourlyBlock = (b: {
    key: string;
    employeeName: string;
    lines: readonly { id: string; jobLabel: string }[];
  }): void => {
    for (const l of b.lines) m.set(`hl:${l.id}`, `${b.employeeName} · ${l.jobLabel}`);
  };
  for (const b of model.painters) hourlyBlock(b);
  for (const b of model.office) hourlyBlock(b);
  for (const entry of model.painterSequence) {
    if (entry.kind === 'hourly') hourlyBlock(entry.block);
    else if (entry.kind === 'ps') {
      const b = entry.block;
      if (b.tsheets) hourlyBlock(b.tsheets);
      for (const j of b.jobs) m.set(`ps:${b.key}:${j.jobRowId}`, `${b.displayName} · ${j.jobShort}`);
      m.set(`psb:${b.key}`, b.displayName);
      for (const x of b.extras) m.set(x.key, `${b.displayName} · ${x.label}`);
    } else {
      const b = entry.block;
      m.set(`sal:${b.name}`, b.name);
      for (const x of b.extras) m.set(x.key, `${b.name} · ${x.label}`);
    }
  }
  for (const b of [...model.admin, ...model.estimators, ...model.managers]) {
    m.set(`sal:${b.name}`, b.name);
    for (const x of b.extras) m.set(x.key, `${b.name} · ${x.label}`);
  }
  return m;
}

// ---------------------------------------------------------------------------
// The comparison
// ---------------------------------------------------------------------------

/** One edits map flattened to the (key, field, printed value) triples it sets. */
export interface EditedField {
  editKey: string;
  field: string;
  manualValue: string;
  /**
   * Round 11 — the map this edit lives in, when it differs from its scope's.
   * QB Hrs edits sit in their own map but on the Pivot by Job's cells, so they
   * share that scope's before/after defaults and are tagged instead of being
   * a second scope over the same cells (which counted every change twice).
   */
  scope?: ConflictScope;
}

export function flattenTimenaviEdits(edits: EditsMap): EditedField[] {
  const out: EditedField[] = [];
  for (const [editKey, e] of Object.entries(edits)) {
    for (const [field, v] of Object.entries(e)) {
      if (v === undefined) continue;
      out.push({ editKey, field, manualValue: printEditValue(field, v) });
    }
  }
  return out;
}

export function flattenPivotEdits(edits: PbjEditsMap): EditedField[] {
  const out: EditedField[] = [];
  for (const [editKey, e] of Object.entries(edits)) {
    for (const [field, v] of Object.entries(e)) {
      // `percentSeed` is a carry-forward seed, never something the user typed.
      if (v === undefined || field === 'percentSeed') continue;
      out.push({ editKey, field, manualValue: printEditValue(field, v) });
    }
  }
  return out;
}

export function flattenQbHourEdits(qbHours: QbHoursMap): EditedField[] {
  return Object.entries(qbHours).map(([editKey, v]) => ({
    editKey,
    field: 'qbHours',
    manualValue: fmtNum(v),
    scope: 'qbHours' as const,
  }));
}

export function flattenPayrollEdits(edits: PayrollEditsMap): EditedField[] {
  const out: EditedField[] = [];
  for (const [editKey, e] of Object.entries(edits)) {
    for (const [field, v] of Object.entries(e)) {
      if (v === undefined) continue;
      // Memo text has no computed counterpart to conflict with.
      if (MEMO_FIELDS.has(field)) continue;
      out.push({ editKey, field, manualValue: printEditValue(field, v) });
    }
  }
  return out;
}

/** Free-text notes typed beside payroll values — never computed, never in conflict. */
const MEMO_FIELDS = new Set([
  'reimbMemo',
  'deductMemo',
  'memo',
  'hourlyPayMemo',
  'mileageNewMemo',
  'amountMemo',
  'mileageMemo',
]);

function printEditValue(field: string, v: unknown): string {
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'yes' : 'no';
  if (v instanceof Date) return fmtDate(v);
  if (typeof v === 'number') {
    if (field === 'percent' || field === 'percentCompletion' || field === 'woPct') {
      return `${round2(v * 100)}%`;
    }
    return fmtNum(v);
  }
  if (typeof v === 'object') return fmtPercentValue(v as PercentValue);
  return String(v);
}

function labelFor(scope: ConflictScope, field: string): string {
  const table =
    scope === 'timenavi'
      ? TIMENAVI_FIELD_LABELS
      : scope === 'payroll'
        ? PAYROLL_FIELD_LABELS
        : PIVOT_FIELD_LABELS;
  return table[field] ?? (scope === 'qbHours' ? 'QB Hrs' : field);
}

export interface ScopeInput {
  scope: ConflictScope;
  before: DefaultsMap;
  after: DefaultsMap;
  labels: Map<string, string>;
  edited: EditedField[];
}

/**
 * The whole decision, in one place.
 *
 * - CONFLICT: this field is hand-edited and the inputs now produce something
 *   different than they did. Only these are put to the user.
 * - ORPHAN: the row the edit belongs to is gone from the new inputs. There is
 *   no new value to accept, so it is reported, not gated. Keeping it is safe —
 *   builders only look up ids they build, so an orphan is inert and re-attaches
 *   if the row comes back.
 * - SILENT: the value moved but nobody had edited it. Counted, never asked
 *   about. This is the rule that keeps the review short enough to read.
 */
export function reviewScopes(scopes: readonly ScopeInput[]): UpdateReview {
  const conflicts: Conflict[] = [];
  const orphans: OrphanedEdit[] = [];
  const silentChanges: ValueChange[] = [];

  for (const s of scopes) {
    const editedFlatKeys = new Set(s.edited.map((e) => keyOf(e.editKey, e.field)));

    // Everything the new input moves that nobody had edited: never asked
    // about, but recorded so the summary can list it.
    for (const [flatKey, beforeValue] of s.before) {
      const afterValue = s.after.get(flatKey);
      if (afterValue === undefined || afterValue === beforeValue) continue;
      if (editedFlatKeys.has(flatKey)) continue;
      const cut = flatKey.lastIndexOf(SEP);
      const editKey = flatKey.slice(0, cut);
      const field = flatKey.slice(cut + 1);
      const scope: ConflictScope = field === 'qbHours' ? 'qbHours' : s.scope;
      silentChanges.push({
        scope,
        editKey,
        field,
        rowLabel: s.labels.get(editKey) ?? editKey,
        fieldLabel: labelFor(scope, field),
        oldValue: beforeValue,
        newValue: afterValue,
      });
    }

    // Edited fields whose computed value moved: the only thing we ask about.
    for (const e of s.edited) {
      const flatKey = keyOf(e.editKey, e.field);
      const beforeValue = s.before.get(flatKey);
      const afterValue = s.after.get(flatKey);
      if (beforeValue === undefined || afterValue === undefined) continue;
      if (beforeValue === afterValue) continue;
      const scope = e.scope ?? s.scope;
      conflicts.push({
        editKey: e.editKey,
        field: e.field,
        scope,
        rowLabel: s.labels.get(e.editKey) ?? e.editKey,
        fieldLabel: labelFor(scope, e.field),
        manualValue: e.manualValue,
        oldValue: beforeValue,
        newValue: afterValue,
      });
    }

    // Edits whose row left the inputs entirely — reported per row, not per field.
    const byKey = new Map<string, EditedField[]>();
    for (const e of s.edited) {
      const list = byKey.get(e.editKey);
      if (list) list.push(e);
      else byKey.set(e.editKey, [e]);
    }
    for (const [editKey, fields] of byKey) {
      const goneNow = fields.every((f) => s.after.get(keyOf(editKey, f.field)) === undefined);
      const wasThere = fields.some((f) => s.before.get(keyOf(editKey, f.field)) !== undefined);
      if (!goneNow || !wasThere) continue;
      const scope = fields[0].scope ?? s.scope;
      orphans.push({
        editKey,
        scope,
        rowLabel: s.labels.get(editKey) ?? editKey,
        summary: fields.map((f) => `${labelFor(scope, f.field)} ${f.manualValue || '(blank)'}`).join(', '),
      });
    }
  }

  return { conflicts, orphans, silentChangeCount: silentChanges.length, silentChanges };
}

// ---------------------------------------------------------------------------
// Round 11 — the details behind the "N values changed" summary
// ---------------------------------------------------------------------------

export type UpdateDetailGroup = 'replaced' | 'kept' | 'orphanDiscarded' | 'orphanKept' | 'changed';

/** One line of the post-update details: where, and what happened there. */
export interface UpdateDetail {
  group: UpdateDetailGroup;
  /** The web-app tab, as the tab strip names it ("TimeNavi Export"). */
  tab: string;
  /** "AP24379 · CCSD Earl Jenkins ES Int" */
  row: string;
  /** The column heading ("Last % Paid"); blank for a whole-row orphan. */
  column: string;
  /** "40% changed to 38%" */
  change: string;
}

const shown = (v: string): string => v || '(blank)';

/** "40% changed to 38%" — the sentence every changed cell is described with. */
export function describeChange(from: string, to: string): string {
  return `${shown(from)} changed to ${shown(to)}`;
}

/**
 * Everything an applied update did, one line per cell, in the order the
 * summary lists them: manual edits the user let the new file replace, manual
 * edits kept over a new value, edits on rows that left the file, then every
 * value that moved on a cell nobody had edited.
 */
export function buildUpdateDetails(
  review: UpdateReview,
  taken: readonly Conflict[],
  discarded: readonly OrphanedEdit[],
): UpdateDetail[] {
  const conflictId = (c: Conflict): string => `${c.scope}${SEP}${c.editKey}${SEP}${c.field}`;
  const orphanId = (o: OrphanedEdit): string => `${o.scope}${SEP}${o.editKey}`;
  const takenIds = new Set(taken.map(conflictId));
  const discardedIds = new Set(discarded.map(orphanId));
  const out: UpdateDetail[] = [];

  for (const c of review.conflicts) {
    if (!takenIds.has(conflictId(c))) continue;
    out.push({
      group: 'replaced',
      tab: SCOPE_TAB[c.scope],
      row: c.rowLabel,
      column: c.fieldLabel,
      change: `${describeChange(c.manualValue, c.newValue)} (your edit replaced by the new file)`,
    });
  }
  for (const c of review.conflicts) {
    if (takenIds.has(conflictId(c))) continue;
    out.push({
      group: 'kept',
      tab: SCOPE_TAB[c.scope],
      row: c.rowLabel,
      column: c.fieldLabel,
      change: `kept your ${shown(c.manualValue)} — the new file has ${shown(c.newValue)} (was ${shown(c.oldValue)})`,
    });
  }
  for (const o of review.orphans) {
    const gone = discardedIds.has(orphanId(o));
    out.push({
      group: gone ? 'orphanDiscarded' : 'orphanKept',
      tab: SCOPE_TAB[o.scope],
      row: o.rowLabel,
      column: '',
      change: `${gone ? 'discarded' : 'kept (parked)'}: ${o.summary} — the row is not in the new file`,
    });
  }
  for (const v of review.silentChanges) {
    out.push({
      group: 'changed',
      tab: SCOPE_TAB[v.scope],
      row: v.rowLabel,
      column: v.fieldLabel,
      change: describeChange(v.oldValue, v.newValue),
    });
  }
  return out;
}
