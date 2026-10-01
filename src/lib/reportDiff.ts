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
  reimbursements: 'Reimbursements',
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
  }));
}

export function flattenPayrollEdits(edits: PayrollEditsMap): EditedField[] {
  const out: EditedField[] = [];
  for (const [editKey, e] of Object.entries(edits)) {
    for (const [field, v] of Object.entries(e)) {
      if (v === undefined) continue;
      // Memo text has no computed counterpart to conflict with.
      if (field === 'reimbMemo' || field === 'deductMemo') continue;
      out.push({ editKey, field, manualValue: printEditValue(field, v) });
    }
  }
  return out;
}

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
  let silent = 0;

  for (const s of scopes) {
    const editedFlatKeys = new Set(s.edited.map((e) => keyOf(e.editKey, e.field)));

    // Everything the new input moves that nobody had edited: counted only.
    for (const [flatKey, beforeValue] of s.before) {
      const afterValue = s.after.get(flatKey);
      if (afterValue === undefined || afterValue === beforeValue) continue;
      if (!editedFlatKeys.has(flatKey)) silent++;
    }

    // Edited fields whose computed value moved: the only thing we ask about.
    for (const e of s.edited) {
      const flatKey = keyOf(e.editKey, e.field);
      const beforeValue = s.before.get(flatKey);
      const afterValue = s.after.get(flatKey);
      if (beforeValue === undefined || afterValue === undefined) continue;
      if (beforeValue === afterValue) continue;
      conflicts.push({
        editKey: e.editKey,
        field: e.field,
        scope: s.scope,
        rowLabel: s.labels.get(e.editKey) ?? e.editKey,
        fieldLabel: labelFor(s.scope, e.field),
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
      orphans.push({
        editKey,
        scope: s.scope,
        rowLabel: s.labels.get(editKey) ?? editKey,
        summary: fields.map((f) => `${labelFor(s.scope, f.field)} ${f.manualValue || '(blank)'}`).join(', '),
      });
    }
  }

  return { conflicts, orphans, silentChangeCount: silent };
}
