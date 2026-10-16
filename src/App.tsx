import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ExceptionsPanel } from './components/ExceptionsPanel';
import { JobsByCrewTable } from './components/JobsByCrewTable';
import {
  SalesByCustTable,
  SalesByRepTable,
  TimesheetMirrorTable,
} from './components/MirrorTable';
import {
  FinalPayrollTab,
  InactiveEmployeesTab,
  InputNumsTab,
  JobsListTab,
  JournalEntryTab,
  ManagementReportTab,
  NotesTab,
  PrFinalCheckTab,
  ProfitshareOutputTab,
} from './components/PayrollExtraTabs';
import { PayrollPreview } from './components/PayrollPreview';
import { PivotByJobTable } from './components/PivotByJobTable';
import { PivotTable } from './components/PivotTable';
import { SettingsModal } from './components/SettingsModal';
import { TimeNaviTable } from './components/TimeNaviTable';
import { TopBar } from './components/TopBar';
import { UpdateReviewModal } from './components/UpdateReviewModal';
import { UploadSlot, type SlotChip } from './components/UploadSlot';
import { timesheetApKey } from './lib/apkey';
import {
  parseCompletedPercentsWorkbook,
  type CompletedPbjJob,
  type CompletedPercentsResult,
} from './lib/completedPercents';
import { buildExceptions, countExceptions } from './lib/exceptions';
import {
  exportFilename,
  exportPayrollWorkbook,
  exportPercentsWorkbook,
  payrollExportFilename,
} from './lib/exportXlsx';
import { addDays, dateToIso, isoToDate } from './lib/format';
import { enrichTimesheet } from './lib/joins';
import { buildJobsByCrew } from './lib/jobsByCrew';
import {
  buildSalesByCustMirror,
  buildSalesByRepMirrorFromParsed,
  buildSalesByRepMirrorFromSheet,
  buildTimesheetMirror,
} from './lib/mirrorTabs';
import {
  parseCustomerSummary,
  parseEmployeeList,
  parseLastWeek,
  parseLastWeekPivotByJob,
  parsePl,
  parseSalesDetail,
  parseTimeNavi,
  parseTimesheet,
  readWorkbook,
  type LastWeekPbjRow,
} from './lib/parsers';
import { directoryFromTimesheet, type LeadResolverContext } from './lib/payroll';
import { buildPayrollExceptions } from './lib/payrollExceptions';
import { buildPayrollPivotByJob } from './lib/payrollPivot';
import { buildPayrollPreview, hasUtPayrollData, mergePayrollPreviews } from './lib/payrollPreview';
import { buildPivot } from './lib/pivot';
import {
  clearSessionSlot,
  deleteSession,
  loadSession,
  loadSessionPointer,
  newSessionId,
  saveSessionEnvelope,
  saveSessionPointer,
  saveSessionSlot,
  sessionStorageAvailable,
  SESSION_PARSER_VERSION,
  SESSION_PAYLOAD_VERSION,
  type SessionSlotMeta,
} from './lib/session';
import {
  collectPayrollDefaults,
  collectPayrollLabels,
  collectPivotDefaults,
  collectPivotLabels,
  collectReportDefaults,
  collectReportLabels,
  flattenPayrollEdits,
  flattenPivotEdits,
  flattenQbHourEdits,
  flattenTimenaviEdits,
  reviewScopes,
  type Conflict,
  type OrphanedEdit,
  type ScopeInput,
  type UpdateReview,
} from './lib/reportDiff';
import { buildShadowPayroll, buildShadowPercents } from './lib/shadowBuild';
import {
  buildPivotByJob,
  EMPTY_MANUAL_STATE,
  pbjJobId,
  type ManualSectionState,
} from './lib/pivotByJob';
import { applyPivotView, EMPTY_PIVOT_VIEW, isPivotViewActive, type PivotViewState } from './lib/pivotView';
import {
  cacheDirectoryLocally,
  clearAppDirectory,
  clearSavedDirectory,
  ensureDefaultPivotColumns,
  fetchAppDirectory,
  fetchAppSettings,
  hasLocalSettings,
  loadAppMode,
  loadSavedDirectory,
  loadSettings,
  newerDirectory,
  persistAppDirectory,
  persistAppSettings,
  probeDirectoryStore,
  probeSettingsStore,
  saveAppMode,
  saveDirectory,
  saveSettings,
  type AppMode,
  type DirectoryStoreStatus,
  type SettingsStoreStatus,
} from './lib/settings';
import { buildReportRows, lastPaidMapFromPivot, sortReportRows } from './lib/timenavi';
import type {
  AppSettings,
  DirectoryEmployee,
  EditsMap,
  ManualSectionRow,
  PayrollEdit,
  PayrollEditsMap,
  PbjEditsMap,
  PbjEmployeeRow,
  PbjJobEdit,
  PivotByJobModel,
  PlParseResult,
  PrevPercent,
  QbHoursMap,
  RowEdit,
  SalesRow,
  SavedDirectory,
  SheetData,
  TimeNaviEvent,
  TimesheetParseResult,
} from './lib/types';

// ---------------------------------------------------------------------------
// Upload slots — Round 6 dual rails
// ---------------------------------------------------------------------------

type SlotId =
  | 'timesheetNv'
  | 'timesheetUt'
  | 'timenavi'
  | 'nvSales'
  | 'nvSummary'
  | 'utSales'
  | 'utSummary'
  | 'lastWeek'
  | 'completedPercents'
  | 'pl'
  | 'plUt'
  | 'employees';

interface SlotDef {
  id: SlotId;
  label: string;
  hint: string;
  optional?: boolean;
}

/** %'s Report rail — Round 6 §1.1. lastWeek + timesheetUt optional; pl is payroll-only. */
const PERCENTS_RAIL: SlotDef[] = [
  { id: 'timesheetNv', label: '1 · NV QuickBooks Time timesheet', hint: 'TSheets weekly payroll export (Nevada)' },
  {
    id: 'timesheetUt',
    label: '2 · UT QuickBooks Time timesheet',
    hint: 'TSheets weekly payroll export (Utah)',
    optional: true,
  },
  { id: 'timenavi', label: '3 · TimeNavi calendar export', hint: 'Reads the "Events" sheet' },
  { id: 'nvSales', label: '4 · NV Sales by Rep Detail', hint: 'QuickBooks Online (Nevada)' },
  { id: 'nvSummary', label: '5 · NV Sales by Customer Summary', hint: 'Job totals (Nevada)' },
  { id: 'utSales', label: '6 · UT Transaction List by Date', hint: 'AllPro Painters LLC (Utah)' },
  { id: 'utSummary', label: '7 · UT Sales by Customer Summary', hint: 'Job totals (Utah)' },
  {
    id: 'lastWeek',
    label: '8 · Last week\u2019s % of Completion',
    hint: 'Last % Paid + Balances/UT carry-forward',
    optional: true,
  },
];

/** Payroll rail — Slot A + Slot B. */
const PAYROLL_RAIL: SlotDef[] = [
  {
    id: 'completedPercents',
    label: 'A · This week\u2019s completed %\u2019s',
    hint: 'Exported % of Completion workbook (Pivot by Job + timesheets)',
  },
  { id: 'pl', label: 'B · NV P&L by Customer (QBO)', hint: 'Required for the payroll tabs' },
  {
    id: 'plUt',
    label: 'C · UT P&L by Customer (QBO)',
    hint: 'Builds Payroll Preview - UT; without it UT jobs keep the Job Total $ fallback',
    optional: true,
  },
];

const SLOT_LABELS: Record<Exclude<SlotId, 'employees'>, string> = {
  timesheetNv: '1 · NV QuickBooks Time timesheet',
  timesheetUt: '2 · UT QuickBooks Time timesheet',
  timenavi: '3 · TimeNavi calendar export',
  nvSales: '4 · NV Sales by Rep Detail',
  nvSummary: '5 · NV Sales by Customer Summary',
  utSales: '6 · UT Transaction List by Date',
  utSummary: '7 · UT Sales by Customer Summary',
  lastWeek: '8 · Last week\u2019s % of Completion',
  completedPercents: 'A · This week\u2019s completed %\u2019s',
  pl: 'B · NV P&L by Customer (QBO)',
  plUt: 'C · UT P&L by Customer (QBO)',
};

interface ParsedData {
  timesheetNv: TimesheetParseResult | null;
  timesheetUt: TimesheetParseResult | null;
  completedPercents: CompletedPercentsResult | null;
  events: TimeNaviEvent[] | null;
  nvSales: SalesRow[] | null;
  nvTotals: Map<string, number> | null;
  utSales: SalesRow[] | null;
  utTotals: Map<string, number> | null;
  lastWeek: Map<string, PrevPercent> | null;
  lastWeekPbj: LastWeekPbjRow[] | null;
  pl: PlParseResult | null;
  /** Round 8 §7.2 — the optional UT P&L (Slot C). */
  plUt: PlParseResult | null;
  directory: DirectoryEmployee[] | null;
  /** Raw first sheets for Round 6 mirror tabs. */
  nvSalesSheet: SheetData | null;
  nvSummarySheet: SheetData | null;
  utSalesSheet: SheetData | null;
  utSummarySheet: SheetData | null;
}

const EMPTY_PARSED: ParsedData = {
  timesheetNv: null,
  timesheetUt: null,
  completedPercents: null,
  events: null,
  nvSales: null,
  nvTotals: null,
  utSales: null,
  utTotals: null,
  lastWeek: null,
  lastWeekPbj: null,
  pl: null,
  plUt: null,
  directory: null,
  nvSalesSheet: null,
  nvSummarySheet: null,
  utSalesSheet: null,
  utSummarySheet: null,
};

const EMPTY_CHIP: SlotChip = { status: 'empty', fileName: '', message: '' };

type PercentsTab =
  | 'salesByRep'
  | 'salesByCust'
  | 'utSalesByRep'
  | 'utSalesByCust'
  | 'timesheets'
  | 'timenavi'
  | 'pivot'
  | 'jobsByCrew'
  | 'pivotByJob'
  | 'exceptions';
type PayrollTab =
  | 'plImport'
  | 'plImportUt'
  | 'mgmtReport'
  | 'payroll'
  | 'payrollUt'
  | 'profitshare'
  | 'inputNums'
  | 'notes'
  | 'final'
  | 'journal'
  | 'jobs'
  | 'inactive'
  | 'payrollExceptions'
  | 'prFinalCheck';
type Tab = PercentsTab | PayrollTab;

export default function App(): JSX.Element {
  const [chips, setChips] = useState<Record<SlotId, SlotChip>>({
    timesheetNv: EMPTY_CHIP,
    timesheetUt: EMPTY_CHIP,
    timenavi: EMPTY_CHIP,
    nvSales: EMPTY_CHIP,
    nvSummary: EMPTY_CHIP,
    utSales: EMPTY_CHIP,
    utSummary: EMPTY_CHIP,
    lastWeek: EMPTY_CHIP,
    completedPercents: EMPTY_CHIP,
    pl: EMPTY_CHIP,
    plUt: EMPTY_CHIP,
    employees: EMPTY_CHIP,
  });
  const [parsed, setParsed] = useState<ParsedData>(EMPTY_PARSED);
  const [settings, setSettings] = useState<AppSettings>(() => loadSettings());
  const [payDateIso, setPayDateIso] = useState<string>('');
  const [edits, setEdits] = useState<EditsMap>({});
  const [pbjEdits, setPbjEdits] = useState<PbjEditsMap>({});
  const [qbHours, setQbHours] = useState<QbHoursMap>({});
  const [balancesState, setBalancesState] = useState<ManualSectionState>(EMPTY_MANUAL_STATE);
  const [utState, setUtState] = useState<ManualSectionState>(EMPTY_MANUAL_STATE);
  const [payrollEdits, setPayrollEdits] = useState<PayrollEditsMap>({});
  const [createdPercents, setCreatedPercents] = useState(false);
  const [createdPayroll, setCreatedPayroll] = useState(false);
  /**
   * A re-upload after a report exists is STAGED, not applied. The report keeps
   * showing the numbers the user has been editing until they press "Update
   * existing report" — which is also what makes a before/after comparison
   * possible, so the review can name exactly which manual edits the new file
   * would overwrite. Only the re-uploaded slot goes pending; every other slot
   * keeps the file it already has, so one changed input never means
   * re-uploading the rest.
   */
  const [pendingParsed, setPendingParsed] = useState<Partial<ParsedData>>({});
  const [pendingChips, setPendingChips] = useState<Partial<Record<SlotId, SlotChip>>>({});
  /** Open review for a staged update; null when there is nothing to decide. */
  const [pendingReview, setPendingReview] = useState<UpdateReview | null>(null);
  /** "Create new report" is destructive once edits exist, so it confirms first. */
  const [confirmRestart, setConfirmRestart] = useState(false);
  /** Passive "here is what the update did" line; dismissible, never a prompt. */
  const [updateSummary, setUpdateSummary] = useState<{
    slots: string[];
    changed: number;
    kept: number;
  } | null>(null);
  /**
   * Crash recovery. The pointer is a SYNCHRONOUS localStorage read, so the very
   * first paint already knows whether a session is coming back and can say so
   * instead of flashing the empty placeholder at someone who has not lost
   * anything. IndexedDB itself is async and arrives a moment later.
   */
  const [sessionId, setSessionId] = useState<string | null>(() => loadSessionPointer()?.id ?? null);
  const [hydrating, setHydrating] = useState<boolean>(() => loadSessionPointer() !== null);
  const [sessionWarning, setSessionWarning] = useState<string | null>(null);
  const [restoredAt, setRestoredAt] = useState<string | null>(null);
  const [tab, setTab] = useState<Tab>('timenavi');
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [mode, setMode] = useState<AppMode>(() => loadAppMode());
  // Round 4 §2 — Pivot by Crew view-level filters (start cleared each session).
  const [pivotView, setPivotView] = useState<PivotViewState>({ ...EMPTY_PIVOT_VIEW });
  const [exportFiltered, setExportFiltered] = useState(false);
  // Employee directory: localStorage (this browser) + shared Vercel Blob store
  // (all devices on the same deployment). Phone never sees another browser's
  // localStorage — only the shared API.
  const [savedDirectory, setSavedDirectory] = useState<SavedDirectory | null>(() => loadSavedDirectory());
  const [directoryStoreStatus, setDirectoryStoreStatus] = useState<DirectoryStoreStatus | null>(null);
  const [directorySyncWarning, setDirectorySyncWarning] = useState<string | null>(null);
  const [missingInputWarning, setMissingInputWarning] = useState<string[] | null>(null);
  // Round 8 §1 — Settings live in a shared store; localStorage is the offline
  // cache. Blob setup stays optional: without it we silently stay local.
  const [settingsStoreStatus, setSettingsStoreStatus] = useState<SettingsStoreStatus | null>(null);
  const [settingsSyncWarning, setSettingsSyncWarning] = useState<string | null>(null);
  const [settingsToast, setSettingsToast] = useState<string | null>(null);
  /** Bumped whenever Settings are adopted from the shared store; remounts the modal. */
  const [settingsRevision, setSettingsRevision] = useState(0);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      const status = await probeDirectoryStore();
      if (cancelled) return;
      setDirectoryStoreStatus(status);
      const remote = await fetchAppDirectory();
      if (cancelled) return;
      const local = loadSavedDirectory();
      // Shared store wins when present; otherwise keep this browser's cache.
      const best = remote ?? local;
      if (!best) return;
      // If both exist, keep the newer uploadedAt.
      const chosen = remote && local ? newerDirectory(local, remote)! : best;
      setSavedDirectory(chosen);
      cacheDirectoryLocally(chosen);
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  // Round 8 §1 — adopt the shared Settings on load; if the store is empty and
  // this browser has its own, push them up once (one-time migration).
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      const status = await probeSettingsStore();
      if (cancelled) return;
      setSettingsStoreStatus(status);
      if (status === 'ready') {
        const remote = await fetchAppSettings();
        if (cancelled || !remote) return;
        setSettings(remote);
        saveSettings(remote);
        setSettingsRevision((n) => n + 1);
        return;
      }
      if (status === 'empty' && hasLocalSettings()) {
        const sync = await persistAppSettings(loadSettings());
        if (cancelled) return;
        if (sync.ok) setSettingsStoreStatus('ready');
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  /**
   * Restore the last session. Runs once, before anything can be uploaded —
   * the slots and both buttons are disabled while this is in flight, so a fast
   * upload cannot be overwritten by the restore landing on top of it.
   *
   * Every setter fires in one block so React batches them into a single
   * commit: there is never a frame showing restored uploads with the edits
   * still missing, which would flash un-edited numbers.
   */
  useEffect(() => {
    const pointer = loadSessionPointer();
    if (!pointer) return;
    let cancelled = false;
    void (async () => {
      const restored = await loadSession(pointer.id);
      if (cancelled) return;
      if (!restored) {
        setHydrating(false);
        setSessionId(null);
        return;
      }
      const env = restored.envelope;
      const e = env.edits as {
        edits?: EditsMap;
        pbjEdits?: PbjEditsMap;
        qbHours?: QbHoursMap;
        balancesState?: ManualSectionState;
        utState?: ManualSectionState;
        payrollEdits?: PayrollEditsMap;
      };
      const parsedBack = restored.parsed as Partial<ParsedData>;
      const nextParsed = { ...EMPTY_PARSED, ...parsedBack };

      setParsed(nextParsed);
      setChips((c) => ({ ...c, ...(restored.chips as Partial<Record<SlotId, SlotChip>>) }));
      setPayDateIso(env.payDateIso);
      setEdits(e.edits ?? {});
      setPbjEdits(e.pbjEdits ?? {});
      setQbHours(e.qbHours ?? {});
      setBalancesState(e.balancesState ?? EMPTY_MANUAL_STATE);
      setUtState(e.utState ?? EMPTY_MANUAL_STATE);
      setPayrollEdits(e.payrollEdits ?? {});
      // Only claim a report the restored inputs can actually build. The model
      // memos guard on these flags, so setting them over a half-restored
      // `parsed` would build a report from partial data.
      setCreatedPercents(env.createdPercents && missingForPercents(nextParsed).length === 0);
      setCreatedPayroll(env.createdPayroll && missingForPayroll(nextParsed).length === 0);
      if (env.mode === 'payroll' || env.mode === 'percents') setMode(env.mode);
      if (env.tab) setTab(env.tab as Tab);
      setSessionId(env.id);
      setRestoredAt(env.savedAt);
      if (restored.degradedSlots.length > 0) {
        setSessionWarning(
          `Restored your session, but ${restored.degradedSlots.length} upload(s) could not be read back. Re-upload them to refresh those numbers.`,
        );
      }
      setHydrating(false);
    })();
    return () => {
      cancelled = true;
    };
    // Mount-only: this restores INTO the initial state, and re-running it would
    // overwrite whatever the user has done since.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!settingsToast) return;
    const t = window.setTimeout(() => setSettingsToast(null), 3500);
    return () => window.clearTimeout(t);
  }, [settingsToast]);

  /**
   * Cache the editing state. Cells commit on blur/Enter rather than per
   * keystroke, so one write per edit is already modest; the debounce just
   * coalesces bursts. The short interval is the real protection — an
   * IndexedDB write started during `pagehide` is not guaranteed to finish.
   */
  useEffect(() => {
    if (hydrating || !sessionId || !sessionStorageAvailable()) return;
    const t = window.setTimeout(() => {
      const envelope = {
        id: sessionId,
        payloadVersion: SESSION_PAYLOAD_VERSION,
        parserVersion: SESSION_PARSER_VERSION,
        savedAt: new Date().toISOString(),
        payDateIso,
        createdPercents,
        createdPayroll,
        mode,
        tab,
        // Which slots a restore should read back. Without this the cached
        // slot records exist but nothing knows to look for them.
        slots: (Object.keys(chips) as SlotId[])
          .filter((k) => k !== 'employees' && chips[k].status === 'ok')
          .map((k) => ({
            slot: k,
            fileName: chips[k].fileName,
            fileSize: 0,
            fileLastModified: 0,
            uploadedAt: new Date().toISOString(),
            chip: chips[k],
          })),
        edits: { edits, pbjEdits, qbHours, balancesState, utState, payrollEdits },
      };
      void saveSessionEnvelope(envelope).then((result) => {
        if (result === 'stale') {
          setSessionWarning('Another tab is editing this session — changes here are not being saved.');
          return;
        }
        if (result === 'failed') {
          setSessionWarning('Could not save this session locally; your work is only in this tab.');
          return;
        }
        saveSessionPointer({
          id: sessionId,
          savedAt: envelope.savedAt,
          label: payDateIso,
          slotCount: Object.values(chips).filter((c) => c.status === 'ok').length,
        });
      });
    }, 600);
    return () => window.clearTimeout(t);
  }, [
    hydrating,
    sessionId,
    payDateIso,
    createdPercents,
    createdPayroll,
    mode,
    tab,
    edits,
    pbjEdits,
    qbHours,
    balancesState,
    utState,
    payrollEdits,
    chips,
  ]);

  /**
   * A stray refresh used to take hours of hand-corrections with no warning.
   * Registered only while there is something to lose.
   */
  const unsavedEditCount =
    countEdits(edits) + countEdits(pbjEdits) + Object.keys(qbHours).length + countEdits(payrollEdits);
  useEffect(() => {
    if (unsavedEditCount === 0) return;
    const onBeforeUnload = (e: BeforeUnloadEvent): void => {
      e.preventDefault();
      // Browsers show their own wording; assigning returnValue is what arms it.
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [unsavedEditCount]);

  // ------------------------------------------------------------------ uploads

  // Round 9 §6 — Slot A feeds the PAYROLL side only. It used to seed the %'s
  // report's Pivot by Job edits, which let a payroll upload move the %'s
  // report's numbers; the payroll process now builds its own model from this
  // upload (see `payrollPivotByJob`) and the two sides never cross.

  /** Mint a session id on first use and publish the synchronous pointer. */
  const sessionIdRef = useRef<string | null>(sessionId);
  sessionIdRef.current = sessionId;
  const ensureSessionId = useCallback((): string => {
    const existing = sessionIdRef.current;
    if (existing) return existing;
    const id = newSessionId();
    sessionIdRef.current = id;
    setSessionId(id);
    return id;
  }, []);

  /**
   * Files staged but not yet applied. Held so that committing an update can
   * cache them the same way a first upload is cached — without re-reading the
   * file off disk, which the browser may no longer allow.
   */
  const pendingUploadsRef = useRef(
    new Map<SlotId, { patch: Partial<ParsedData>; file: File; message: string; bytes: ArrayBuffer }>(),
  );

  /** Cache one slot's parsed patch (and its bytes) for crash recovery. */
  const persistSlot = useCallback(
    (slot: SlotId, patch: Partial<ParsedData>, file: File, message: string, bytes?: ArrayBuffer): void => {
      if (!sessionStorageAvailable()) return;
      const id = ensureSessionId();
      const meta: SessionSlotMeta = {
        slot,
        fileName: file.name,
        fileSize: file.size,
        fileLastModified: file.lastModified,
        uploadedAt: new Date().toISOString(),
        chip: { status: 'ok', fileName: file.name, message },
      };
      void saveSessionSlot(id, slot, strippedForStorage(patch), meta, bytes).then((ok) => {
        if (!ok) setSessionWarning('Could not cache this upload locally; your work is only in this tab.');
      });
    },
    [ensureSessionId],
  );

  const handleFile = useCallback(
    async (id: SlotId, file: File): Promise<void> => {
    // Once a rail has a report on screen, a new file for one of its slots is
    // STAGED rather than applied — the numbers the user has been editing must
    // not move under them until they ask for it. Before that first report,
    // this is the original straight-to-`parsed` path.
    const rail = railForSlot(id);
    const staged = rail === 'percents' ? createdPercents : rail === 'payroll' ? createdPayroll : false;
    const fail = (message: string): void => {
      if (staged) {
        // A bad re-upload must never damage the working report: drop the
        // pending patch and leave the committed data alone.
        setPendingChips((c) => ({ ...c, [id]: { status: 'error', fileName: file.name, message } }));
        setPendingParsed((p) => {
          const next = { ...p };
          for (const k of Object.keys(clearForSlot(id)) as (keyof ParsedData)[]) delete next[k];
          return next;
        });
        return;
      }
      setChips((c) => ({ ...c, [id]: { status: 'error', fileName: file.name, message } }));
      setParsed((p) => ({ ...p, ...clearForSlot(id) }));
      const sid = sessionIdRef.current;
      if (sid) void clearSessionSlot(sid, id);
    };
    try {
      const bytes = await file.arrayBuffer();
      const wbd = readWorkbook(bytes);
      const ok = (message: string, patch: Partial<ParsedData>): void => {
        if (staged) {
          setPendingParsed((p) => ({ ...p, ...patch }));
          setPendingChips((c) => ({ ...c, [id]: { status: 'ok', fileName: file.name, message } }));
          pendingUploadsRef.current.set(id, { patch, file, message, bytes });
          return;
        }
        setParsed((p) => ({ ...p, ...patch }));
        setChips((c) => ({ ...c, [id]: { status: 'ok', fileName: file.name, message } }));
        persistSlot(id, patch, file, message, bytes);
      };
      switch (id) {
        case 'timesheetNv': {
          const res = parseTimesheet(wbd, 'NV');
          ok(
            `✓ ${res.rows.length} rows${res.lunchRemoved > 0 ? ` · ${res.lunchRemoved} lunch removed` : ''}`,
            { timesheetNv: res },
          );
          // Default pay date = week-ending Friday (max local_date) + 7 days.
          let max: Date | null = null;
          for (const r of res.rows) {
            if (r.localDate && (!max || r.localDate.getTime() > max.getTime())) max = r.localDate;
          }
          if (max) setPayDateIso(dateToIso(addDays(max, 7)));
          break;
        }
        case 'timesheetUt': {
          const res = parseTimesheet(wbd, 'UT');
          ok(
            `✓ ${res.rows.length} rows${res.lunchRemoved > 0 ? ` · ${res.lunchRemoved} lunch removed` : ''}`,
            { timesheetUt: res },
          );
          break;
        }
        case 'completedPercents': {
          const res = parseCompletedPercentsWorkbook(wbd);
          if ('error' in res) {
            fail(res.error);
            break;
          }
          ok(
            `✓ ${res.jobs.length} jobs · ${res.timesheet.rows.length} timesheet rows` +
              (res.timesheet.lunchRemoved > 0 ? ` · ${res.timesheet.lunchRemoved} lunch removed` : ''),
            { completedPercents: res },
          );
          // Pay date from completed timesheet when not already set from NV upload.
          if (!payDateIso) {
            let max: Date | null = null;
            for (const r of res.timesheet.rows) {
              if (r.localDate && (!max || r.localDate.getTime() > max.getTime())) max = r.localDate;
            }
            if (max) setPayDateIso(dateToIso(addDays(max, 7)));
          }
          break;
        }
        case 'timenavi': {
          const events = parseTimeNavi(wbd);
          ok(`✓ ${events.length} events`, { events });
          break;
        }
        case 'nvSales': {
          const rows = parseSalesDetail(wbd, 'customer');
          ok(`✓ ${rows.length} sales rows`, { nvSales: rows, nvSalesSheet: wbd.sheets[0] ?? null });
          break;
        }
        case 'nvSummary': {
          const totals = parseCustomerSummary(wbd);
          ok(`✓ ${totals.size} job totals`, { nvTotals: totals, nvSummarySheet: wbd.sheets[0] ?? null });
          break;
        }
        case 'utSales': {
          const rows = parseSalesDetail(wbd, 'name');
          ok(`✓ ${rows.length} transaction rows`, { utSales: rows, utSalesSheet: wbd.sheets[0] ?? null });
          break;
        }
        case 'utSummary': {
          const totals = parseCustomerSummary(wbd);
          ok(`✓ ${totals.size} job totals`, { utTotals: totals, utSummarySheet: wbd.sheets[0] ?? null });
          break;
        }
        case 'lastWeek': {
          const map = parseLastWeek(wbd);
          const pbjRows = parseLastWeekPivotByJob(wbd);
          ok(
            `✓ ${map.size} prior % values${pbjRows.length > 0 ? ` · ${pbjRows.length} Pivot by Job rows` : ''}`,
            { lastWeek: map, lastWeekPbj: pbjRows },
          );
          break;
        }
        case 'pl': {
          const pl = parsePl(wbd, { entity: 'NV', mapping: settings.plAccountMapping.NV });
          ok(
            `✓ ${pl.byJob.size} job columns${pl.duplicateColumns.length > 0 ? ` · ${pl.duplicateColumns.length} duplicated` : ''}`,
            { pl },
          );
          break;
        }
        case 'plUt': {
          const plUt = parsePl(wbd, { entity: 'UT', mapping: settings.plAccountMapping.UT });
          ok(
            `✓ ${plUt.byJob.size} job columns${plUt.duplicateColumns.length > 0 ? ` · ${plUt.duplicateColumns.length} duplicated` : ''}`,
            { plUt },
          );
          break;
        }
        case 'employees': {
          // Round 5 §1 — rail no longer hosts this slot; Settings Replace still
          // routes through handleFile when needed. Prefer onDirectoryReplace.
          const directory = parseEmployeeList(wbd);
          const active = directory.filter((d) => d.active).length;
          ok(`✓ ${active} active / ${directory.length - active} inactive`, { directory });
          const saved = saveDirectory(directory, file.name);
          setSavedDirectory(saved);
          void persistAppDirectory(saved).then((sync) => {
            if (sync.ok) {
              setDirectorySyncWarning(null);
              setDirectoryStoreStatus('ready');
            } else {
              setDirectorySyncWarning(sync.message);
            }
          });
          break;
        }
      }
    } catch (err) {
      fail(err instanceof Error ? err.message : 'Could not read this file.');
    }
    },
    [payDateIso, settings.plAccountMapping, createdPercents, createdPayroll, persistSlot],
  );

  /**
   * Combined NV+UT timesheet for the %'s REPORT side (NV first).
   * Round 9 §6 — this no longer falls back to Slot A: the %'s report reads
   * only the %'s uploads, and the payroll process reads only its own.
   */
  const combinedTimesheet = useMemo((): TimesheetParseResult | null => {
    const nv = parsed.timesheetNv;
    const ut = parsed.timesheetUt;
    if (!nv && !ut) return null;
    return {
      rows: [...(nv?.rows ?? []), ...(ut?.rows ?? [])],
      lunchRemoved: (nv?.lunchRemoved ?? 0) + (ut?.lunchRemoved ?? 0),
    };
  }, [parsed.timesheetNv, parsed.timesheetUt]);

  const percentsMissingInputs = useMemo((): string[] => missingForPercents(parsed), [parsed]);
  const payrollMissingInputs = useMemo((): string[] => missingForPayroll(parsed), [parsed]);

  const percentsReady = percentsMissingInputs.length === 0;
  const payrollReady = payrollMissingInputs.length === 0;
  const ready = mode === 'percents' ? percentsReady : payrollReady;
  const modeCreated = mode === 'percents' ? createdPercents : createdPayroll;
  const canDownload =
    mode === 'percents'
      ? createdPercents && percentsReady
      : createdPayroll && payrollReady;

  // ------------------------------------------------------------------ derived

  const legendOrder = useMemo(() => settings.legend.map((l) => l.estimator), [settings.legend]);
  const colorByEstimator = useMemo(
    () => new Map(settings.legend.map((l) => [l.estimator, l.color])),
    [settings.legend],
  );
  /** Known initials (palette + pivot config) guarding the title fallback. */
  const knownEstimators = useMemo(
    () => new Set<string>([...settings.legend.map((l) => l.estimator), ...settings.pivotColumnOrder]),
    [settings.legend, settings.pivotColumnOrder],
  );
  const payDate = useMemo(() => (payDateIso ? isoToDate(payDateIso) : null), [payDateIso]);

  /**
   * Round 9 §1/§6 — the payroll process's hours come from the `timesheets`
   * tab of Slot A ("This week's completed %'s") and from nowhere else.
   */
  const payrollTimesheet = useMemo(
    (): TimesheetParseResult | null => parsed.completedPercents?.timesheet ?? null,
    [parsed.completedPercents],
  );

  /**
   * Round 9 §6 — the %'s report carries forward from ITS own Slot 8 upload
   * ("Last week's % of Completion"). It no longer borrows Slot A, which
   * belongs to the payroll process.
   */
  const effectiveLastWeekPbj = useMemo(() => parsed.lastWeekPbj, [parsed.lastWeekPbj]);

  /** apKey → jobcode_2 of the first timesheet row with that key (column M). */
  const timesheetApKeyMap = useMemo(() => {
    const m = new Map<string, string>();
    const src = combinedTimesheet;
    if (src) {
      for (const r of src.rows) {
        const key = timesheetApKey(r.jobcode2);
        if (key && !m.has(key)) m.set(key, r.jobcode2);
      }
    }
    return m;
  }, [combinedTimesheet]);

  /** `Last % Paid` source #1 — input #8's Pivot by Job `% of Completion` by AP#. */
  const lastPaidByAp = useMemo(() => lastPaidMapFromPivot(parsed.lastWeekPbj), [parsed.lastWeekPbj]);

  const reportRows = useMemo(() => {
    if (
      !createdPercents ||
      !percentsReady ||
      !parsed.events ||
      !parsed.nvSales ||
      !parsed.utSales ||
      !parsed.nvTotals ||
      !parsed.utTotals
    ) {
      return [];
    }
    const rows = buildReportRows(
      parsed.events,
      edits,
      {
        nvSales: parsed.nvSales,
        utSales: parsed.utSales,
        nvTotals: parsed.nvTotals,
        utTotals: parsed.utTotals,
        lastWeekMap: lastPaidByAp,
        timesheetJobcodeByApKey: timesheetApKeyMap,
        knownEstimators,
      },
      payDate,
    );
    return sortReportRows(rows, legendOrder);
  }, [
    createdPercents,
    percentsReady,
    parsed.events,
    parsed.nvSales,
    parsed.utSales,
    parsed.nvTotals,
    parsed.utTotals,
    lastPaidByAp,
    edits,
    timesheetApKeyMap,
    knownEstimators,
    legendOrder,
    payDate,
  ]);

  const includedApKeys = useMemo(() => {
    const s = new Set<string>();
    for (const r of reportRows) {
      if (r.included && r.apKey) s.add(r.apKey);
    }
    return s;
  }, [reportRows]);

  /** apKey → RESOLVED estimator from the grid (books + fallback + edits). */
  const estimatorByApKey = useMemo(() => {
    const m = new Map<string, string>();
    for (const r of reportRows) {
      if (r.included && r.apKey && r.estimator && !m.has(r.apKey)) m.set(r.apKey, r.estimator);
    }
    return m;
  }, [reportRows]);

  /** %'s report enrichment — built from the %'s uploads only (Round 9 §6). */
  const enriched = useMemo(() => {
    if (!createdPercents || !combinedTimesheet) return [];
    return enrichTimesheet(
      combinedTimesheet.rows,
      parsed.nvSales ?? [],
      parsed.utSales ?? [],
      includedApKeys,
      estimatorByApKey,
    );
  }, [createdPercents, combinedTimesheet, parsed.nvSales, parsed.utSales, includedApKeys, estimatorByApKey]);

  /**
   * Round 9 §6 — the payroll side's enriched timesheet: Slot A's `timesheets`
   * tab, with estimators resolved from Slot A's own Pivot by Job rows.
   */
  const payrollEstimatorByApKey = useMemo(() => {
    const m = new Map<string, string>();
    for (const j of parsed.completedPercents?.jobs ?? []) {
      if (j.apKey && j.estimator && !m.has(j.apKey)) m.set(j.apKey, j.estimator);
    }
    return m;
  }, [parsed.completedPercents]);

  /**
   * Round 10 — AP# → estimator fill hex on the payroll side, from Slot A's own
   * Estimator column and the Settings palette. Colors the Input #'s PO cells.
   */
  const payrollEstimatorColorByAp = useMemo(() => {
    const m = new Map<string, string>();
    for (const [ap, est] of payrollEstimatorByApKey) {
      const hex = colorByEstimator.get(est);
      if (hex) m.set(ap, hex);
    }
    return m;
  }, [payrollEstimatorByApKey, colorByEstimator]);

  const payrollApKeys = useMemo(() => {
    const set = new Set<string>();
    for (const j of parsed.completedPercents?.jobs ?? []) if (j.apKey) set.add(j.apKey);
    return set;
  }, [parsed.completedPercents]);

  const payrollEnriched = useMemo(() => {
    if (!createdPayroll || !payrollTimesheet) return [];
    return enrichTimesheet(payrollTimesheet.rows, [], [], payrollApKeys, payrollEstimatorByApKey);
  }, [createdPayroll, payrollTimesheet, payrollApKeys, payrollEstimatorByApKey]);

  // Round 6 — mirror models for %'s tabs + export
  const salesByRepMirror = useMemo(() => {
    if (parsed.nvSalesSheet) return buildSalesByRepMirrorFromSheet(parsed.nvSalesSheet, 'nv');
    return buildSalesByRepMirrorFromParsed(parsed.nvSales ?? []);
  }, [parsed.nvSalesSheet, parsed.nvSales]);

  const utSalesByRepMirror = useMemo(() => {
    if (parsed.utSalesSheet) return buildSalesByRepMirrorFromSheet(parsed.utSalesSheet, 'ut');
    return buildSalesByRepMirrorFromParsed(parsed.utSales ?? []);
  }, [parsed.utSalesSheet, parsed.utSales]);

  const salesByCustMirror = useMemo(
    () => buildSalesByCustMirror(parsed.nvSummarySheet),
    [parsed.nvSummarySheet],
  );
  const utSalesByCustMirror = useMemo(
    () => buildSalesByCustMirror(parsed.utSummarySheet),
    [parsed.utSummarySheet],
  );

  const timeNaviRowByAp = useMemo(() => {
    const m = new Map<string, number>();
    let sheetRow = 2; // header = 1
    for (const r of reportRows) {
      if (!r.included) continue;
      if (r.apKey && !m.has(r.apKey)) m.set(r.apKey, sheetRow);
      sheetRow++;
    }
    return m;
  }, [reportRows]);

  const timesheetMirror = useMemo(
    () => buildTimesheetMirror(enriched, timeNaviRowByAp),
    [enriched, timeNaviRowByAp],
  );

  const jobsByCrew = useMemo(() => buildJobsByCrew(enriched), [enriched]);

  /** Round 9 §6 — payroll filenames/date ranges follow the PAYROLL timesheet. */
  const payrollWeekBounds = useMemo(() => {
    let weekStart: Date | null = null;
    let weekEnd: Date | null = null;
    for (const r of payrollTimesheet?.rows ?? []) {
      if (!r.localDate) continue;
      if (!weekStart || r.localDate.getTime() < weekStart.getTime()) weekStart = r.localDate;
      if (!weekEnd || r.localDate.getTime() > weekEnd.getTime()) weekEnd = r.localDate;
    }
    return { weekStart, weekEnd };
  }, [payrollTimesheet]);

  const weekBounds = useMemo(() => {
    let weekStart: Date | null = null;
    let weekEnd: Date | null = null;
    for (const r of enriched) {
      if (!r.localDate) continue;
      if (!weekStart || r.localDate.getTime() < weekStart.getTime()) weekStart = r.localDate;
      if (!weekEnd || r.localDate.getTime() > weekEnd.getTime()) weekEnd = r.localDate;
    }
    return { weekStart, weekEnd };
  }, [enriched]);

  const pivot = useMemo(
    () => buildPivot(enriched, settings.pivotColumnOrder, payDate),
    [enriched, settings.pivotColumnOrder, payDate],
  );

  // Keep the fixed MH…RF block intact (heals stale localStorage missing RF)
  // and append any other estimators seen in the data (never '#N/A').
  useEffect(() => {
    const base = ensureDefaultPivotColumns(settings.pivotColumnOrder);
    const extras = pivot.estimators.filter((e) => e !== '#N/A' && !base.includes(e));
    const nextOrder = [...base, ...extras];
    const unchanged =
      nextOrder.length === settings.pivotColumnOrder.length &&
      nextOrder.every((e, i) => e === settings.pivotColumnOrder[i]);
    if (unchanged) return;
    const next: AppSettings = { ...settings, pivotColumnOrder: nextOrder };
    setSettings(next);
    saveSettings(next);
  }, [pivot.estimators, settings]);

  // ------------------------------------------------ Round 3: payroll models

  /** §4.2 — fresh upload wins, else the persisted directory, else null. */
  const effectiveDirectory = useMemo(
    () => parsed.directory ?? savedDirectory?.employees ?? null,
    [parsed.directory, savedDirectory],
  );

  const leadCtx = useMemo<LeadResolverContext>(() => {
    const rows = combinedTimesheet?.rows ?? [];
    const crewGroups = [...new Set(rows.map((r) => r.group).filter(Boolean))];
    return {
      directory: effectiveDirectory ?? directoryFromTimesheet(rows),
      crewGroups,
      subcontractors: settings.subcontractors,
      aliases: settings.crewLeadAliases,
    };
  }, [effectiveDirectory, combinedTimesheet, settings.subcontractors, settings.crewLeadAliases]);

  /** Round 9 §6 — the payroll side resolves leads against ITS timesheet. */
  const payrollLeadCtx = useMemo<LeadResolverContext>(() => {
    const rows = payrollTimesheet?.rows ?? [];
    const crewGroups = [...new Set(rows.map((r) => r.group).filter(Boolean))];
    return {
      directory: effectiveDirectory ?? directoryFromTimesheet(rows),
      crewGroups,
      subcontractors: settings.subcontractors,
      aliases: settings.crewLeadAliases,
    };
  }, [effectiveDirectory, payrollTimesheet, settings.subcontractors, settings.crewLeadAliases]);

  /**
   * Round 9 §6 — the payroll process's Pivot by Job, built straight from Slot
   * A. Every reference the payroll tabs used to take from the %'s report
   * output (WO%, WO Override, Crew Lead, % of Completion, the gray exclusion
   * flag, the bottom-of-pivot jobs) now comes from THIS model, i.e. from the
   * "This week's completed %'s" upload on the payroll rail.
   */
  const payrollPivotByJob = useMemo((): PivotByJobModel | null => {
    if (!parsed.completedPercents) return null;
    return buildPayrollPivotByJob({ completed: parsed.completedPercents, settings });
  }, [parsed.completedPercents, settings]);

  /**
   * Round 9 §6 — QB Hrs overrides for the payroll side come from Slot A's own
   * `QB Hrs` column, not from the %'s report's Pivot by Job tab (that map is
   * the other rail's editing state and must not reach these numbers).
   */
  const payrollQbHours = useMemo((): QbHoursMap => {
    const out: QbHoursMap = {};
    for (const j of parsed.completedPercents?.jobs ?? []) {
      const id = pbjJobId(j.customer || '(blank)', j.job);
      for (const e of j.employees) {
        if (e.qbHours != null) out[`${id}\u0000${e.username}`] = e.qbHours;
      }
    }
    return out;
  }, [parsed.completedPercents]);

  const pivotByJob = useMemo(
    () =>
      buildPivotByJob({
        enriched,
        reportRows,
        settings,
        leadCtx,
        knownEstimators,
        payDate,
        edits: pbjEdits,
        qbHours,
        lastWeekPbj: effectiveLastWeekPbj,
        balancesState,
        utState,
        // Round 9 §6 — the P&L lives on the PAYROLL rail; the %'s report
        // resolves its names from its own sales/summary uploads.
        pl: null,
        nvTotals: parsed.nvTotals,
        utTotals: parsed.utTotals,
        nvSales: parsed.nvSales,
        utSales: parsed.utSales,
        // Round 8 §4.3 — carried-job customers resolve by AP# against these.
        nvSalesByCust: salesByCustMirror,
        utSalesByCust: utSalesByCustMirror,
      }),
    [
      enriched,
      reportRows,
      settings,
      leadCtx,
      knownEstimators,
      payDate,
      pbjEdits,
      qbHours,
      effectiveLastWeekPbj,
      parsed.nvTotals,
      parsed.utTotals,
      parsed.nvSales,
      parsed.utSales,
      salesByCustMirror,
      utSalesByCustMirror,
      balancesState,
      utState,
    ],
  );

  const exceptions = useMemo(
    () =>
      buildExceptions({
        reportRows,
        enriched,
        lunchRemoved: combinedTimesheet?.lunchRemoved ?? 0,
        payDate,
        pivotByJob,
      }),
    [reportRows, enriched, combinedTimesheet, payDate, pivotByJob],
  );
  const exceptionCount = useMemo(() => countExceptions(exceptions), [exceptions]);

  /**
   * Round 9 §6 — every payroll input is the payroll rail's own: Slot A's
   * Pivot by Job + timesheets, and the P&L uploads. `reportRows` is empty
   * here on purpose — the %'s report's TimeNavi grid is no longer a payroll
   * input, so nothing on this tab can move when the %'s report changes.
   */
  const payrollPreview = useMemo(() => {
    if (!createdPayroll || !parsed.pl || !payrollPivotByJob) return null;
    return buildPayrollPreview({
      pbj: payrollPivotByJob,
      enriched: payrollEnriched,
      reportRows: [],
      pl: parsed.pl,
      plUt: parsed.plUt,
      entity: 'NV',
      directory: effectiveDirectory,
      settings,
      payDate,
      edits: payrollEdits,
      qbHours: payrollQbHours,
      knownEstimators,
      leadCtx: payrollLeadCtx,
      nvSales: null,
      utSales: null,
      lastWeekPbj: parsed.completedPercents?.lastWeekPbj ?? null,
      // Round 10 — customer names come from the payroll input's Sales by Rep.
      customerNameByAp: parsed.completedPercents?.customerNameByAp ?? null,
    });
  }, [
    createdPayroll,
    parsed.pl,
    parsed.plUt,
    parsed.completedPercents,
    effectiveDirectory,
    payrollPivotByJob,
    payrollEnriched,
    settings,
    payDate,
    payrollEdits,
    payrollQbHours,
    knownEstimators,
    payrollLeadCtx,
  ]);

  const showUtPayroll = useMemo(
    () =>
      payrollPivotByJob !== null &&
      hasUtPayrollData({ enriched: payrollEnriched, pbj: payrollPivotByJob, plUt: parsed.plUt }),
    [payrollEnriched, payrollPivotByJob, parsed.plUt],
  );

  const payrollPreviewUt = useMemo(() => {
    if (!createdPayroll || !showUtPayroll || !payrollPivotByJob) return null;
    return buildPayrollPreview({
      pbj: payrollPivotByJob,
      enriched: payrollEnriched,
      reportRows: [],
      pl: parsed.pl,
      plUt: parsed.plUt,
      entity: 'UT',
      directory: effectiveDirectory,
      settings,
      payDate,
      edits: payrollEdits,
      qbHours: payrollQbHours,
      knownEstimators,
      leadCtx: payrollLeadCtx,
      nvSales: null,
      utSales: null,
      lastWeekPbj: parsed.completedPercents?.lastWeekPbj ?? null,
      // Round 10 — customer names come from the payroll input's Sales by Rep.
      customerNameByAp: parsed.completedPercents?.customerNameByAp ?? null,
    });
  }, [
    createdPayroll,
    showUtPayroll,
    parsed.pl,
    parsed.plUt,
    parsed.completedPercents,
    effectiveDirectory,
    payrollPivotByJob,
    payrollEnriched,
    settings,
    payDate,
    payrollEdits,
    payrollQbHours,
    knownEstimators,
    payrollLeadCtx,
  ]);

  const payrollPreviewAll = useMemo(
    () => (payrollPreview ? mergePayrollPreviews(payrollPreview, payrollPreviewUt) : null),
    [payrollPreview, payrollPreviewUt],
  );

  const payrollExceptions = useMemo(() => {
    if (!createdPayroll || !payrollPivotByJob) return [];
    return buildPayrollExceptions({
      pbj: payrollPivotByJob,
      preview:
        payrollPreview ?? {
          entity: 'NV',
          title: '',
          weekRange: '',
          painters: [],
          leads: [],
          painterSequence: [],
          office: [],
          admin: [],
          estimators: [],
          managers: [],
          inactiveEmployees: [],
          inactiveZeroHours: [],
          newLeadsToAdd: [],
          activatedPainters: [],
          dailyOtNotes: [],
        },
      previewUt: payrollPreviewUt,
      reportRows: [],
      pl: parsed.pl,
      plUt: parsed.plUt,
      hasDirectory: effectiveDirectory !== null,
      enriched: payrollEnriched,
      nvSales: null,
      utSales: null,
      lastWeekPbj: parsed.completedPercents?.lastWeekPbj ?? null,
    });
  }, [
    createdPayroll,
    payrollPivotByJob,
    payrollPreview,
    payrollPreviewUt,
    parsed.pl,
    parsed.plUt,
    parsed.completedPercents,
    effectiveDirectory,
    payrollEnriched,
  ]);
  const payrollExceptionCount = useMemo(() => countExceptions(payrollExceptions), [payrollExceptions]);

  // Round 8 §9.4 — a crew lead that resolved from the input data but has no
  // entry in the crew profit-share % map is auto-added at the seeded default.
  // buildPayrollPreview stays pure and just reports them; writing Settings is
  // the App layer's job, so no compute path mutates the store.
  const newLeadsToAdd = payrollPreviewAll?.newLeadsToAdd;
  useEffect(() => {
    if (!newLeadsToAdd || newLeadsToAdd.length === 0) return;
    const key = (s: string): string => s.trim().toLowerCase();
    const additions = newLeadsToAdd.filter(
      (n) => !settings.profitShare.some((p) => key(p.lead) === key(n.lead)),
    );
    if (additions.length === 0) return;
    const next: AppSettings = {
      ...settings,
      profitShare: [...settings.profitShare, ...additions.map((n) => ({ lead: n.lead, pct: n.pct }))],
    };
    setSettings(next);
    saveSettings(next);
    void persistAppSettings(next);
    const names = additions.map((n) => n.lead).join(', ');
    setSettingsToast(
      additions.length === 1
        ? `New crew lead ${names} added at default 38%`
        : `New crew leads added at default 38%: ${names}`,
    );
  }, [newLeadsToAdd, settings]);

  /**
   * Round 10 — a painter with hours is active. The preview reports who that
   * applies to and the App layer flips the persisted roster flag, so the
   * "inactive (*)" marker does not come back next week.
   */
  const activatedPainters = payrollPreviewAll?.activatedPainters;
  useEffect(() => {
    if (!activatedPainters || activatedPainters.length === 0) return;
    const names = new Set(activatedPainters.map((n) => n.trim().toLowerCase()));
    if (!settings.painterRoster.some((p) => !p.active && names.has(p.name.trim().toLowerCase()))) {
      return;
    }
    const next: AppSettings = {
      ...settings,
      painterRoster: settings.painterRoster.map((p) =>
        names.has(p.name.trim().toLowerCase()) ? { ...p, active: true } : p,
      ),
    };
    setSettings(next);
    saveSettings(next);
    void persistAppSettings(next);
    setSettingsToast(
      activatedPainters.length === 1
        ? `${activatedPainters[0]} has hours — marked active`
        : `${activatedPainters.length} painters have hours — marked active`,
    );
  }, [activatedPainters, settings]);

  const estimatorOptions = useMemo(() => {
    const set = new Set<string>(legendOrder);
    for (const r of reportRows) if (r.estimator) set.add(r.estimator);
    return [...set];
  }, [legendOrder, reportRows]);

  // ------------------------------------------------------------------ actions

  const onEdit = useCallback((rowId: string, patch: RowEdit): void => {
    setEdits((prev) => {
      const merged: RowEdit = { ...prev[rowId], ...patch };
      // Keys explicitly set to undefined mean "reset to the computed default".
      for (const k of Object.keys(patch) as (keyof RowEdit)[]) {
        if (patch[k] === undefined) delete merged[k];
      }
      const next = { ...prev };
      if (Object.keys(merged).length === 0) delete next[rowId];
      else next[rowId] = merged;
      return next;
    });
  }, []);

  const navigateToRow = useCallback((rowId: string): void => {
    setTab('timenavi');
    window.setTimeout(() => {
      const el = document.getElementById(`tn-${rowId}`);
      if (!el) return;
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      el.classList.add('flash-row');
      window.setTimeout(() => el.classList.remove('flash-row'), 2500);
    }, 80);
  }, []);

  // ---------------------------------------------- Round 3 edit callbacks

  const onEditPbjJob = useCallback((jobId: string, patch: PbjJobEdit): void => {
    setPbjEdits((prev) => ({ ...prev, [jobId]: { ...prev[jobId], ...patch } }));
  }, []);

  const onEditQbHours = useCallback((jobId: string, username: string, value: number | null): void => {
    const key = `${jobId}\u0000${username}`;
    setQbHours((prev) => {
      const next = { ...prev };
      if (value === null) delete next[key];
      else next[key] = value;
      return next;
    });
  }, []);

  const manualStateSetter = (section: 'balances' | 'ut') =>
    section === 'balances' ? setBalancesState : setUtState;

  const onManualEdit = useCallback(
    (section: 'balances' | 'ut', rowId: string, patch: Partial<ManualSectionRow>): void => {
      manualStateSetter(section)((prev) => ({
        ...prev,
        edits: { ...prev.edits, [rowId]: { ...prev.edits[rowId], ...patch } },
      }));
    },
    [],
  );

  const onManualAdd = useCallback((section: 'balances' | 'ut'): void => {
    const row: ManualSectionRow = {
      id: `manual:${section}:${Date.now()}`,
      customer: '',
      job: '',
      apKey: null,
      estimator: '',
      crewLead: '',
      woOverride: null,
      woPct: null,
      percent: null,
      livePercent: null,
      percentOverride: false,
      excludedFromPayroll: false,
      notes: '',
      colorHex: null,
    };
    manualStateSetter(section)((prev) => ({ ...prev, added: [...prev.added, row] }));
  }, []);

  const onManualDelete = useCallback((section: 'balances' | 'ut', rowId: string): void => {
    manualStateSetter(section)((prev) => ({ ...prev, deleted: [...prev.deleted, rowId] }));
  }, []);

  const onPayrollEdit = useCallback((key: string, patch: PayrollEdit): void => {
    setPayrollEdits((prev) => ({ ...prev, [key]: { ...prev[key], ...patch } }));
  }, []);

  // §2 — the filtered view model renders on screen; the underlying data,
  // the other tabs, the payroll math and the exceptions never change.
  const viewPivot = useMemo(() => applyPivotView(pivot, pivotView), [pivot, pivotView]);
  const filtersActive = isPivotViewActive(pivotView);

  const onDownload = useCallback((): void => {
    if (mode === 'payroll') {
      void exportPayrollWorkbook({
        payDate,
        weekStart: payrollWeekBounds.weekStart,
        weekEnd: payrollWeekBounds.weekEnd,
        pl: parsed.pl,
        plUt: parsed.plUt,
        plSourceFileName: chips.pl.fileName || null,
        payrollPreview,
        payrollPreviewUt,
        payrollPreviewAll,
        payrollExceptions,
        mileageRate: settings.mileageRate,
        defaultProfitSharePct: settings.defaultProfitSharePct,
        estimatorByApKey: payrollEstimatorByApKey,
        legend: settings.legend,
      });
      return;
    }
    void exportPercentsWorkbook({
      rows: reportRows,
      pivot: exportFiltered && filtersActive ? viewPivot : pivot,
      legend: settings.legend,
      payDate,
      pivotByJob,
      percentExceptions: exceptions,
      salesByRep: salesByRepMirror,
      salesByCust: salesByCustMirror,
      utSalesByRep: utSalesByRepMirror,
      utSalesByCust: utSalesByCustMirror,
      timesheetMirror,
      jobsByCrew,
    });
  }, [
    mode,
    payDate,
    weekBounds,
    payrollWeekBounds,
    payrollEstimatorByApKey,
    parsed.pl,
    parsed.plUt,
    chips.pl.fileName,
    payrollPreview,
    payrollPreviewUt,
    payrollPreviewAll,
    payrollExceptions,
    settings.mileageRate,
    settings.defaultProfitSharePct,
    settings.legend,
    estimatorByApKey,
    reportRows,
    pivot,
    viewPivot,
    exportFiltered,
    filtersActive,
    pivotByJob,
    exceptions,
    salesByRepMirror,
    salesByCustMirror,
    utSalesByRepMirror,
    utSalesByCustMirror,
    timesheetMirror,
    jobsByCrew,
  ]);

  const onSaveSettings = useCallback((next: AppSettings): void => {
    setSettings(next);
    saveSettings(next);
    setSettingsOpen(false);
    // Last write wins on the shared store; local cache is already updated.
    void persistAppSettings(next).then((sync) => {
      if (sync.ok) {
        setSettingsSyncWarning(null);
        setSettingsStoreStatus('ready');
        setSettingsToast('Settings saved ✓ (shared)');
      } else {
        setSettingsSyncWarning(sync.message);
        setSettingsToast('Settings saved ✓ (this browser)');
      }
    });
  }, []);

  /** Round 8 §1 — re-pull the shared Settings (another browser may have saved). */
  const onReloadSettings = useCallback(async (): Promise<void> => {
    const status = await probeSettingsStore();
    setSettingsStoreStatus(status);
    const remote = await fetchAppSettings();
    if (!remote) {
      setSettingsToast('No shared settings to load');
      return;
    }
    setSettings(remote);
    saveSettings(remote);
    setSettingsRevision((n) => n + 1);
    setSettingsSyncWarning(null);
    setSettingsToast('Settings reloaded ✓ (shared)');
  }, []);

  // Replace / clear: always cache locally; also write the shared store so phone
  // and desktop both see the same list on the Vercel URL.
  const onDirectoryReplace = useCallback(async (file: File): Promise<string | null> => {
    try {
      const wbd = readWorkbook(await file.arrayBuffer());
      const directory = parseEmployeeList(wbd);
      const saved = saveDirectory(directory, file.name);
      setSavedDirectory(saved);
      const sync = await persistAppDirectory(saved);
      if (sync.ok) {
        setDirectorySyncWarning(null);
        setDirectoryStoreStatus('ready');
      } else {
        setDirectorySyncWarning(sync.message);
      }
      setParsed((p) => ({ ...p, directory }));
      const active = directory.filter((d) => d.active).length;
      setChips((c) => ({
        ...c,
        employees: { status: 'ok', fileName: file.name, message: `✓ ${active} active / ${directory.length - active} inactive` },
      }));
      return null;
    } catch (err) {
      return err instanceof Error ? err.message : 'Could not read this file.';
    }
  }, []);

  const onDirectoryClear = useCallback((): void => {
    clearSavedDirectory();
    void clearAppDirectory();
    setSavedDirectory(null);
    setDirectorySyncWarning(null);
    setParsed((p) => ({ ...p, directory: null }));
    setChips((c) => ({ ...c, employees: EMPTY_CHIP }));
  }, []);

  /** Manual edits on the %'s rail, counted for the "start over" confirmation. */
  const percentsEditCount = useMemo(
    () =>
      countEdits(edits) +
      countEdits(pbjEdits) +
      Object.keys(qbHours).length +
      countEdits(balancesState.edits) +
      countEdits(utState.edits),
    [edits, pbjEdits, qbHours, balancesState, utState],
  );
  const payrollEditCount = useMemo(() => countEdits(payrollEdits), [payrollEdits]);
  const activeEditCount = mode === 'percents' ? percentsEditCount : payrollEditCount;
  /** Does the rail on screen already have a report built? */
  const hasReport = mode === 'percents' ? createdPercents : createdPayroll;
  const pendingSlots = useMemo(
    () => (Object.keys(pendingChips) as SlotId[]).filter((k) => pendingChips[k]?.status === 'ok'),
    [pendingChips],
  );

  /** Drop every manual edit on one rail — only ever from "Create new report". */
  const clearRailEdits = useCallback((which: AppMode): void => {
    if (which === 'percents') {
      setEdits({});
      setPbjEdits({});
      setQbHours({});
      setBalancesState(EMPTY_MANUAL_STATE);
      setUtState(EMPTY_MANUAL_STATE);
    } else {
      setPayrollEdits({});
    }
  }, []);

  /** Commit the staged uploads; nothing else about the session changes. */
  const commitPending = useCallback((): void => {
    setParsed((p) => ({ ...p, ...pendingParsed }));
    for (const [slot, u] of pendingUploadsRef.current) {
      persistSlot(slot, u.patch, u.file, u.message, u.bytes);
    }
    pendingUploadsRef.current.clear();
    setPendingParsed({});
    setPendingChips({});
    setChips((c) => {
      const next = { ...c };
      for (const [slot, chip] of Object.entries(pendingChips)) {
        if (chip && chip.status === 'ok') next[slot as SlotId] = chip;
      }
      return next;
    });
  }, [pendingParsed, pendingChips, persistSlot]);

  const createReport = useCallback(
    (which: AppMode): void => {
      const missing = which === 'percents' ? percentsMissingInputs : payrollMissingInputs;
      if (missing.length > 0) {
        setMissingInputWarning(missing);
        return;
      }
      setMissingInputWarning(null);
      commitPending();
      clearRailEdits(which);
      if (which === 'percents') {
        setCreatedPercents(true);
        setTab('timenavi');
      } else {
        setCreatedPayroll(true);
        setTab('payroll');
      }
    },
    [percentsMissingInputs, payrollMissingInputs, commitPending, clearRailEdits],
  );

  /**
   * "Create new report" starts over, so it throws away hand work. Ask first
   * whenever there is any to lose; go straight through when there is not.
   */
  const onCreateReport = useCallback((): void => {
    if (activeEditCount > 0 && (mode === 'percents' ? createdPercents : createdPayroll)) {
      setConfirmRestart(true);
      return;
    }
    createReport(mode);
  }, [activeEditCount, mode, createdPercents, createdPayroll, createReport]);

  /**
   * "Update existing report" — re-derive from the staged upload while keeping
   * the manual edits, and ask about nothing except the edits the new file would
   * actually overwrite.
   *
   * Both shadow builds run with EMPTY edit maps, which is what exposes the
   * computed layer the builders normally fold the edits into. Comparing the two
   * gives "what the file changed"; intersecting that with the edit maps gives
   * the only thing worth a decision.
   */
  const shadowDeps = useMemo(
    () => ({
      settings,
      directory: effectiveDirectory,
      payDate,
      knownEstimators,
      legendOrder,
      balancesState,
      utState,
    }),
    [settings, effectiveDirectory, payDate, knownEstimators, legendOrder, balancesState, utState],
  );

  /**
   * Apply an update after the review. A conflict the user resolved as "take the
   * new value" is applied by DELETING that one field from its edit map, which
   * relinks the cell to the computed value — the same mechanism as the existing
   * % relink. Everything not named here is left exactly as it was.
   */
  const applyUpdate = useCallback(
    (review: UpdateReview, choices: { take: Conflict[]; discardOrphans: OrphanedEdit[] }): void => {
      /**
       * Remove one overridden field. When that was the row's last field the
       * whole key goes too — an empty `{}` left behind would still be counted
       * as a manual edit, which is what the "clears N edits" confirmation
       * reads, and it would never be true again.
       */
      const withoutField = <M extends Record<string, object>>(prev: M, key: string, field: string): M => {
        const row = { ...(prev[key] ?? {}) } as Record<string, unknown>;
        delete row[field];
        const next = { ...prev } as Record<string, unknown>;
        if (Object.keys(row).length === 0) delete next[key];
        else next[key] = row;
        return next as M;
      };
      for (const c of choices.take) {
        if (c.scope === 'timenavi') {
          setEdits((prev) => withoutField(prev, c.editKey, c.field));
        } else if (c.scope === 'pivot') {
          setPbjEdits((prev) => withoutField(prev, c.editKey, c.field));
        } else if (c.scope === 'qbHours') {
          setQbHours((prev) => {
            const next = { ...prev };
            delete next[c.editKey];
            return next;
          });
        } else {
          setPayrollEdits((prev) => withoutField(prev, c.editKey, c.field));
        }
      }
      for (const o of choices.discardOrphans) {
        if (o.scope === 'timenavi') {
          setEdits((prev) => {
            const next = { ...prev };
            delete next[o.editKey];
            return next;
          });
        } else if (o.scope === 'pivot') {
          setPbjEdits((prev) => {
            const next = { ...prev };
            delete next[o.editKey];
            return next;
          });
        } else if (o.scope === 'qbHours') {
          setQbHours((prev) => {
            const next = { ...prev };
            delete next[o.editKey];
            return next;
          });
        } else {
          setPayrollEdits((prev) => {
            const next = { ...prev };
            delete next[o.editKey];
            return next;
          });
        }
      }
      const slotNames = (Object.keys(pendingChips) as SlotId[])
        .filter((k) => pendingChips[k]?.status === 'ok')
        .map((k) => SLOT_LABELS[k as Exclude<SlotId, 'employees'>] ?? k);
      commitPending();
      setPendingReview(null);
      setUpdateSummary({
        slots: slotNames,
        changed: review.silentChangeCount + choices.take.length,
        kept: review.conflicts.length - choices.take.length,
      });
    },
    [pendingChips, commitPending],
  );

  const onUpdateReport = useCallback((): void => {
    if (pendingSlots.length === 0) return;
    const after = { ...parsed, ...pendingParsed };
    const scopes: ScopeInput[] = [];

    if (mode === 'percents') {
      const b = buildShadowPercents(parsed, shadowDeps);
      const a = buildShadowPercents(after, shadowDeps);
      scopes.push({
        scope: 'timenavi',
        before: collectReportDefaults(b.reportRows),
        after: collectReportDefaults(a.reportRows),
        labels: collectReportLabels(a.reportRows.length > 0 ? a.reportRows : b.reportRows),
        edited: flattenTimenaviEdits(edits),
      });
      if (b.pivotByJob && a.pivotByJob) {
        scopes.push({
          scope: 'pivot',
          before: collectPivotDefaults(b.pivotByJob),
          after: collectPivotDefaults(a.pivotByJob),
          labels: collectPivotLabels(a.pivotByJob),
          edited: flattenPivotEdits(pbjEdits),
        });
        scopes.push({
          scope: 'qbHours',
          before: collectPivotDefaults(b.pivotByJob),
          after: collectPivotDefaults(a.pivotByJob),
          labels: collectPivotLabels(a.pivotByJob),
          edited: flattenQbHourEdits(qbHours),
        });
      }
    } else {
      const b = buildShadowPayroll(parsed, shadowDeps);
      const a = buildShadowPayroll(after, shadowDeps);
      if (b.payroll && a.payroll) {
        scopes.push({
          scope: 'payroll',
          before: collectPayrollDefaults(b.payroll),
          after: collectPayrollDefaults(a.payroll),
          labels: collectPayrollLabels(a.payroll),
          edited: flattenPayrollEdits(payrollEdits),
        });
      }
    }

    const review = reviewScopes(scopes);
    if (review.conflicts.length === 0 && review.orphans.length === 0) {
      // Nothing the user typed is affected — just apply it.
      applyUpdate(review, { take: [], discardOrphans: [] });
      return;
    }
    setPendingReview(review);
  }, [
    pendingSlots,
    parsed,
    pendingParsed,
    mode,
    shadowDeps,
    edits,
    pbjEdits,
    qbHours,
    payrollEdits,
    applyUpdate,
  ]);

  /** Fill Slot A from this session's %'s report (no re-upload). */
  const onUseSessionPercents = useCallback((): void => {
    if (!createdPercents || !combinedTimesheet) return;
    const fromSession = completedFromSession(combinedTimesheet, pivotByJob, parsed.lastWeekPbj);
    setParsed((p) => ({ ...p, completedPercents: fromSession }));
    setChips((c) => ({
      ...c,
      completedPercents: {
        status: 'ok',
        fileName: 'this session',
        message: `✓ from this session · ${fromSession.jobs.length} jobs · ${fromSession.timesheet.rows.length} rows`,
      },
    }));
  }, [createdPercents, combinedTimesheet, pivotByJob, parsed.lastWeekPbj]);

  // ------------------------------------------------------------------ render

  const includedCount = reportRows.filter((r) => r.included).length;
  const activeRail = mode === 'percents' ? PERCENTS_RAIL : PAYROLL_RAIL;

  return (
    <div className="flex h-full flex-col">
      <TopBar
        canDownload={canDownload}
        onDownload={onDownload}
        onOpenSettings={() => setSettingsOpen(true)}
        downloadName={
          mode === 'payroll'
            ? payrollExportFilename(payDate, payrollWeekBounds.weekStart, payrollWeekBounds.weekEnd)
            : exportFilename(payDate)
        }
        exportFiltered={exportFiltered}
        onExportFilteredChange={setExportFiltered}
        filtersActive={filtersActive}
        mode={mode}
        onModeChange={(m) => {
          setMode(m);
          saveAppMode(m);
          setTab(m === 'percents' ? 'timenavi' : 'payroll');
        }}
      />
      <div className="flex min-h-0 flex-1">
        {/* Left rail — uploads (mode-specific; parsed data persists across switches) */}
        <aside className="flex w-[320px] shrink-0 flex-col gap-2 overflow-y-auto border-r border-neutral-300 bg-neutral-50 p-3">
          <p className="text-[11px] leading-snug text-neutral-500">
            {mode === 'percents'
              ? 'Upload the weekly %\u2019s inputs (.xlsx or .csv). Everything is parsed locally — no data leaves this computer.'
              : 'Upload this week\u2019s completed %\u2019s workbook and the P&L. Switching modes keeps your uploads.'}
          </p>
          {activeRail.map((def) => (
            <UploadSlot
              key={def.id}
              label={def.label}
              hint={def.hint}
              optional={def.optional}
              chip={chips[def.id]}
              pendingChip={pendingChips[def.id]}
              disabled={hydrating}
              onFile={(file) => void handleFile(def.id, file)}
            />
          ))}
          {mode === 'payroll' && createdPercents && (
            <button
              type="button"
              className="rounded border border-emerald-300 bg-emerald-50 px-2 py-1.5 text-left text-[11px] font-semibold text-emerald-800 hover:bg-emerald-100"
              onClick={onUseSessionPercents}
              title="Copy this session's Pivot by Job + timesheet into Slot A"
            >
              Use this session&apos;s %&apos;s report
            </button>
          )}
          {/* Round 5 §1 — Employee Directory status (Settings is the only upload home). */}
          <button
            type="button"
            className={`flex items-center gap-1.5 rounded border px-2 py-1.5 text-left text-[11px] ${
              savedDirectory
                ? 'border-emerald-300 bg-emerald-50 text-emerald-800 hover:bg-emerald-100'
                : 'border-neutral-300 bg-white text-neutral-600 hover:bg-neutral-50'
            }`}
            onClick={() => setSettingsOpen(true)}
            title="Manage the employee directory in Settings → Employee Directory"
          >
            {savedDirectory ? (
              <>
                <span className="font-semibold">Employee list:</span>
                <span>
                  ✓ {savedDirectory.activeCount} active / {savedDirectory.inactiveCount} inactive · updated{' '}
                  {savedDirectory.uploadedAt
                    ? new Date(savedDirectory.uploadedAt).toLocaleString()
                    : '—'}
                </span>
              </>
            ) : (
              <span>
                Employee list: not saved — upload in <span className="font-semibold">Settings</span>
              </span>
            )}
          </button>
          {/* The cached session is never invisible: it says it is there, when
              it was saved, and offers the way out. */}
          {(restoredAt || sessionWarning) && (
            <div className="mt-1 rounded border border-sky-300 bg-sky-50 px-2 py-1.5 text-[11px] text-sky-900">
              {restoredAt && (
                <div>
                  Session restored · saved{' '}
                  {new Date(restoredAt).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}
                </div>
              )}
              {sessionWarning && <div className="mt-0.5 font-semibold text-amber-800">{sessionWarning}</div>}
              <button
                type="button"
                className="mt-1 rounded bg-sky-200 px-1.5 py-0.5 font-semibold hover:bg-sky-300"
                onClick={() => {
                  if (sessionId) void deleteSession(sessionId);
                  window.location.reload();
                }}
                title="Forget the cached uploads and edits and start from an empty session"
              >
                Start fresh
              </button>
            </div>
          )}
          <label className="mt-1 flex flex-col gap-1 text-xs font-semibold">
            Pay date
            <input
              type="date"
              className="rounded border border-neutral-300 px-2 py-1.5 font-normal"
              value={payDateIso}
              onChange={(e) => setPayDateIso(e.target.value)}
            />
            <span className="text-[10px] font-normal text-neutral-400">
              Defaults to the NV timesheet week-ending Friday + 7 days. Drives the pivot title and the filename.
            </span>
          </label>
          {/* Once a report exists, "Update" is the everyday button and "Create
              new" is the start-over — so Update leads and carries the accent. */}
          {hasReport && (
            <button
              type="button"
              className={`mt-1 py-2.5 text-base ${pendingSlots.length > 0 ? 'btn-primary' : 'btn'}`}
              disabled={pendingSlots.length === 0 || hydrating}
              title={
                pendingSlots.length > 0
                  ? 'Re-derive from the new file, keeping your manual edits'
                  : 'Upload a changed input first — the rest stay as they are'
              }
              onClick={onUpdateReport}
            >
              Update existing report
              {pendingSlots.length > 0 && (
                <span className="ml-1 text-xs font-normal">({pendingSlots.length} pending)</span>
              )}
            </button>
          )}
          <button
            type="button"
            className={hasReport ? 'btn mt-1 py-2 text-sm' : 'btn-primary mt-1 py-2.5 text-base'}
            disabled={hydrating}
            title={
              hasReport
                ? 'Start over from the uploads — clears every manual edit'
                : mode === 'percents'
                  ? percentsReady
                    ? 'Build the %\u2019s report tables'
                    : 'Upload required inputs 1 and 3–7 first'
                  : payrollReady
                    ? 'Build the payroll preview'
                    : 'Upload Slot A and Slot B first'
            }
            onClick={onCreateReport}
          >
            {hasReport
              ? 'Create new report'
              : mode === 'percents'
                ? "Create %'s Report"
                : 'Create Payroll Preview'}
          </button>
        </aside>

        {pendingReview && (
          <UpdateReviewModal
            review={pendingReview}
            sourceLabel={pendingSlots
              .map((k) => SLOT_LABELS[k as Exclude<SlotId, 'employees'>] ?? k)
              .join(', ')}
            onCancel={() => setPendingReview(null)}
            onApply={(choices) => applyUpdate(pendingReview, choices)}
          />
        )}

        {confirmRestart && (
          <div
            className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
            role="dialog"
            aria-modal="true"
            aria-labelledby="restart-title"
          >
            <div className="w-full max-w-md rounded-lg border border-neutral-300 bg-white p-5 shadow-xl">
              <h2 id="restart-title" className="mb-2 text-base font-bold text-neutral-900">
                Start over?
              </h2>
              <p className="mb-4 text-xs text-neutral-600">
                This clears {activeEditCount} manual{' '}
                {activeEditCount === 1 ? 'edit' : 'edits'} on the{' '}
                {mode === 'percents' ? "%’s report" : 'payroll preview'} and rebuilds from the uploads.
                To keep {activeEditCount === 1 ? 'it' : 'them'}, use{' '}
                <span className="font-semibold">Update existing report</span> instead.
              </p>
              <div className="flex justify-end gap-2">
                <button type="button" className="btn px-4 py-1.5 text-sm" onClick={() => setConfirmRestart(false)}>
                  Cancel
                </button>
                <button
                  type="button"
                  className="btn-primary px-4 py-1.5 text-sm"
                  onClick={() => {
                    setConfirmRestart(false);
                    createReport(mode);
                  }}
                >
                  Clear {activeEditCount} and rebuild
                </button>
              </div>
            </div>
          </div>
        )}

        {missingInputWarning && (
          <div
            className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
            role="dialog"
            aria-modal="true"
            aria-labelledby="missing-inputs-title"
          >
            <div className="w-full max-w-md rounded-lg border border-neutral-300 bg-white p-5 shadow-xl">
              <h2 id="missing-inputs-title" className="mb-2 text-base font-bold text-neutral-900">
                Missing inputs
              </h2>
              <p className="mb-3 text-xs text-neutral-600">
                {mode === 'percents'
                  ? 'Upload the required %\u2019s files before creating the report. Slots 2 and 8 are optional.'
                  : 'Upload Slot A (completed %\u2019s) and Slot B (P&L) before creating the payroll preview.'}
              </p>
              <ul className="mb-4 list-inside list-disc space-y-1 text-xs text-red-700">
                {missingInputWarning.map((label) => (
                  <li key={label}>{label}</li>
                ))}
              </ul>
              <div className="flex justify-end">
                <button type="button" className="btn-primary px-4 py-1.5 text-sm" onClick={() => setMissingInputWarning(null)}>
                  OK
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Main area — tabs */}
        <main className="flex min-w-0 flex-1 flex-col">
          <nav className="flex shrink-0 gap-1 overflow-x-auto border-b border-neutral-300 bg-white px-3 pt-2">
            {(
              mode === 'percents'
                ? ([
                    ['salesByRep', 'Sales by Rep'],
                    ['salesByCust', 'Sales by Cust'],
                    ['utSalesByRep', 'UT SalesbyRep'],
                    ['utSalesByCust', 'UT SalesbyCust'],
                    ['timesheets', 'timesheets'],
                    ['timenavi', `TimeNavi Export${createdPercents ? ` (${includedCount})` : ''}`],
                    ['pivot', 'Pivot by Crew'],
                    ['jobsByCrew', 'Jobs by Crew'],
                    ['pivotByJob', 'Pivot by Job'],
                    ['exceptions', `%'s Exceptions${createdPercents ? ` (${exceptionCount})` : ''}`],
                  ] as [Tab, string][])
                : ([
                    ['plImport', 'P&L by Job Import'],
                    ...(parsed.plUt ? ([['plImportUt', 'P&L by Job Import - UT']] as [Tab, string][]) : []),
                    ['mgmtReport', 'Management Report Export'],
                    ['payroll', 'Payroll Preview'],
                    ...(showUtPayroll
                      ? ([['payrollUt', 'Payroll Preview - UT']] as [Tab, string][])
                      : []),
                    ['profitshare', 'Profitshare Output'],
                    ['inputNums', "Input #'s"],
                    ['notes', 'Notes'],
                    ['final', 'Final Payroll Breakdown'],
                    ['journal', 'Journal Entry'],
                    ['jobs', 'Jobs'],
                    ['inactive', 'Inactive Employees'],
                    [
                      'payrollExceptions',
                      `Payroll Exceptions${createdPayroll ? ` (${payrollExceptionCount})` : ''}`,
                    ],
                    ['prFinalCheck', 'PR Final Check'],
                  ] as [Tab, string][])
            ).map(([t, label]) => (
              <button
                key={t}
                type="button"
                className={`shrink-0 rounded-t-md border border-b-0 px-3 py-1.5 text-sm font-medium ${
                  tab === t
                    ? 'border-neutral-300 bg-neutral-100 text-neutral-900'
                    : 'border-transparent text-neutral-500 hover:text-neutral-800'
                }`}
                onClick={() => setTab(t)}
              >
                {label}
              </button>
            ))}
          </nav>
          <div className="min-h-0 flex-1 overflow-auto bg-neutral-100">
            {hydrating ? (
              <div className="flex h-full items-center justify-center p-10">
                <div className="rounded-lg border border-sky-300 bg-sky-50 px-6 py-5 text-center">
                  <div className="text-sm font-bold text-sky-900">Restoring your last session…</div>
                  <div className="mt-1 text-xs text-sky-800">
                    Reading the uploads and manual edits cached in this browser.
                  </div>
                </div>
              </div>
            ) : !modeCreated ? (
              <Placeholder mode={mode} ready={ready} />
            ) : tab === 'salesByRep' ? (
              <SalesByRepTable headers={salesByRepMirror.headers} rows={salesByRepMirror.rows} />
            ) : tab === 'salesByCust' ? (
              <SalesByCustTable rows={salesByCustMirror} />
            ) : tab === 'utSalesByRep' ? (
              <SalesByRepTable headers={utSalesByRepMirror.headers} rows={utSalesByRepMirror.rows} />
            ) : tab === 'utSalesByCust' ? (
              <SalesByCustTable rows={utSalesByCustMirror} />
            ) : tab === 'timesheets' ? (
              <TimesheetMirrorTable rows={timesheetMirror} />
            ) : tab === 'timenavi' ? (
              <TimeNaviTable
                rows={reportRows}
                colorByEstimator={colorByEstimator}
                estimatorOptions={estimatorOptions}
                onEdit={onEdit}
              />
            ) : tab === 'pivot' ? (
              <PivotTable
                pivot={viewPivot}
                fullPivot={pivot}
                colorByEstimator={colorByEstimator}
                view={pivotView}
                onViewChange={setPivotView}
              />
            ) : tab === 'jobsByCrew' ? (
              <JobsByCrewTable model={jobsByCrew} />
            ) : tab === 'pivotByJob' ? (
              <PivotByJobTable
                model={pivotByJob}
                onEditJob={onEditPbjJob}
                onEditQbHours={onEditQbHours}
                onManualEdit={onManualEdit}
                onManualAdd={onManualAdd}
                onManualDelete={onManualDelete}
              />
            ) : tab === 'exceptions' ? (
              <ExceptionsPanel groups={exceptions} onNavigate={navigateToRow} />
            ) : tab === 'payroll' ? (
              <PayrollPreview model={payrollPreview} mileageRate={settings.mileageRate} onEdit={onPayrollEdit} />
            ) : tab === 'payrollUt' ? (
              <PayrollPreview model={payrollPreviewUt} mileageRate={settings.mileageRate} onEdit={onPayrollEdit} />
            ) : tab === 'profitshare' ? (
              <ProfitshareOutputTab model={payrollPreviewAll} onEdit={onPayrollEdit} />
            ) : tab === 'mgmtReport' ? (
              <ManagementReportTab
                model={payrollPreviewAll}
                defaultPsPct={settings.defaultProfitSharePct}
                onEdit={onPayrollEdit}
              />
            ) : tab === 'inputNums' ? (
              <InputNumsTab
                model={payrollPreviewAll}
                onEdit={onPayrollEdit}
                mileageRate={settings.mileageRate}
                estimatorColorByAp={payrollEstimatorColorByAp}
              />
            ) : tab === 'notes' ? (
              <NotesTab model={payrollPreviewAll} mileageRate={settings.mileageRate} onEdit={onPayrollEdit} />
            ) : tab === 'final' ? (
              <FinalPayrollTab
                model={payrollPreviewAll}
                mileageRate={settings.mileageRate}
                onEdit={onPayrollEdit}
                defaultPsPct={settings.defaultProfitSharePct}
              />
            ) : tab === 'journal' ? (
              <JournalEntryTab model={payrollPreviewAll} />
            ) : tab === 'jobs' ? (
              <JobsListTab model={payrollPreviewAll} pl={parsed.pl} />
            ) : tab === 'payrollExceptions' ? (
              <ExceptionsPanel groups={payrollExceptions} onNavigate={navigateToRow} />
            ) : tab === 'prFinalCheck' ? (
              <PrFinalCheckTab pl={parsed.pl} sourceFileName={chips.pl.fileName} />
            ) : tab === 'inactive' ? (
              <InactiveEmployeesTab model={payrollPreviewAll} />
            ) : tab === 'plImport' ? (
              <PlImportTable pl={parsed.pl} />
            ) : tab === 'plImportUt' ? (
              <PlImportTable pl={parsed.plUt} />
            ) : (
              <PayrollPreview model={payrollPreview} mileageRate={settings.mileageRate} onEdit={onPayrollEdit} />
            )}
          </div>
        </main>
      </div>

      {settingsOpen && (
        <SettingsModal
          key={settingsRevision}
          settings={settings}
          directory={savedDirectory}
          directoryStoreStatus={directoryStoreStatus}
          directorySyncWarning={directorySyncWarning}
          settingsStoreStatus={settingsStoreStatus}
          settingsSyncWarning={settingsSyncWarning}
          onReloadSettings={onReloadSettings}
          onDirectoryReplace={onDirectoryReplace}
          onDirectoryClear={onDirectoryClear}
          onSave={onSaveSettings}
          onClose={() => setSettingsOpen(false)}
        />
      )}

      {settingsToast && (
        <div
          className="fixed bottom-4 right-4 z-[60] rounded bg-neutral-900 px-3 py-2 text-xs font-semibold text-white shadow-lg"
          role="status"
        >
          {settingsToast}
        </div>
      )}

      {/* What the last update actually did. Information, never a prompt — it
          stays until dismissed so it survives a scroll through the tabs. */}
      {updateSummary && (
        <div
          className="fixed bottom-4 left-1/2 z-[60] flex -translate-x-1/2 items-center gap-3 rounded border border-amber-300 bg-amber-50 px-3 py-2 text-xs font-semibold text-amber-900 shadow-lg"
          role="status"
        >
          <span>
            Updated from {updateSummary.slots.join(', ') || 'the new upload'} ·{' '}
            {updateSummary.changed} {updateSummary.changed === 1 ? 'value' : 'values'} changed
            {updateSummary.kept > 0 && (
              <> · {updateSummary.kept} manual {updateSummary.kept === 1 ? 'edit' : 'edits'} kept</>
            )}
          </span>
          <button
            type="button"
            className="rounded bg-amber-200 px-1.5 py-0.5 text-[11px] hover:bg-amber-300"
            onClick={() => setUpdateSummary(null)}
          >
            dismiss
          </button>
        </div>
      )}
    </div>
  );
}

/**
 * Which report a slot feeds. Round 9 §6 keeps the two rails apart, so a
 * re-upload is staged against the rail that owns it: re-reading Slot A can
 * never stage against the %'s report, and vice versa. The employee directory
 * belongs to neither — it persists on its own and is never staged.
 */
/**
 * Which required uploads a rail is still missing. Pure, because the restore
 * path has to ask the same question of a freshly-read `ParsedData` before it
 * can claim a report is ready to show.
 */
/**
 * Trim a slot's patch before caching it. `fills` is a dense per-cell grid (up
 * to 400k cells) that only `rowHasExclusionGray` reads, during parsing — dead
 * weight once the patch exists, and the bulkiest thing in it. The employee
 * directory is skipped too: it already persists through `settings.ts`.
 */
function strippedForStorage(patch: Partial<ParsedData>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(patch)) {
    if (k === 'directory') continue;
    const asRecord = v as unknown as Record<string, unknown> | null;
    if (asRecord && typeof asRecord === 'object' && 'rows' in asRecord && 'fills' in asRecord) {
      const { fills: _dropped, ...rest } = asRecord;
      out[k] = rest;
      continue;
    }
    out[k] = v;
  }
  return out;
}

export function missingForPercents(parsed: ParsedData): string[] {
  const checks: [Exclude<SlotId, 'employees'>, boolean][] = [
    ['timesheetNv', parsed.timesheetNv !== null],
    ['timenavi', parsed.events !== null],
    ['nvSales', parsed.nvSales !== null],
    ['nvSummary', parsed.nvTotals !== null],
    ['utSales', parsed.utSales !== null],
    ['utSummary', parsed.utTotals !== null],
    // lastWeek + timesheetUt optional; pl is payroll-only
  ];
  return checks.filter(([, ok]) => !ok).map(([id]) => SLOT_LABELS[id]);
}

export function missingForPayroll(parsed: ParsedData): string[] {
  const checks: [Exclude<SlotId, 'employees'>, boolean][] = [
    ['completedPercents', parsed.completedPercents !== null],
    ['pl', parsed.pl !== null],
  ];
  return checks.filter(([, ok]) => !ok).map(([id]) => SLOT_LABELS[id]);
}

/** Rows that still carry a field; an emptied-out edit object is not an edit. */
function countEdits(map: Record<string, object>): number {
  return Object.values(map).filter((row) => Object.keys(row).length > 0).length;
}

function railForSlot(id: SlotId): 'percents' | 'payroll' | null {
  switch (id) {
    case 'completedPercents':
    case 'pl':
    case 'plUt':
      return 'payroll';
    case 'employees':
      return null;
    default:
      return 'percents';
  }
}

function clearForSlot(id: SlotId): Partial<ParsedData> {
  switch (id) {
    case 'timesheetNv':
      return { timesheetNv: null };
    case 'timesheetUt':
      return { timesheetUt: null };
    case 'completedPercents':
      return { completedPercents: null };
    case 'timenavi':
      return { events: null };
    case 'nvSales':
      return { nvSales: null, nvSalesSheet: null };
    case 'nvSummary':
      return { nvTotals: null, nvSummarySheet: null };
    case 'utSales':
      return { utSales: null, utSalesSheet: null };
    case 'utSummary':
      return { utTotals: null, utSummarySheet: null };
    case 'lastWeek':
      return { lastWeek: null, lastWeekPbj: null };
    case 'pl':
      return { pl: null };
    case 'plUt':
      return { plUt: null };
    case 'employees':
      return { directory: null };
  }
}

/** Build a CompletedPercentsResult from the in-session %'s models. */
function completedFromSession(
  timesheet: TimesheetParseResult,
  pbj: PivotByJobModel,
  lastWeekPbj: LastWeekPbjRow[] | null,
): CompletedPercentsResult {
  const jobs: CompletedPbjJob[] = [];
  for (const c of pbj.customers) {
    for (const j of c.jobs) {
      jobs.push({
        section: 'nv',
        customer: c.customer,
        job: j.jobcode2,
        apKey: j.analysis.apKey ?? '',
        onCalendar: j.analysis.onCalendar,
        estimator: j.analysis.estimator,
        crewLead: j.analysis.crewLead,
        split: j.analysis.split,
        woOverride: j.analysis.woOverride,
        woPct: j.analysis.woPct,
        percent: j.analysis.percent,
        notes: j.analysis.notes,
        hourly: j.analysis.hourly,
        excludedFromPayroll: false,
        byDate: j.byDate,
        total: j.total,
        employees: j.employees.map((e: PbjEmployeeRow) => ({
          username: e.username,
          byDate: e.byDate,
          total: e.total,
          qbHours: e.qbHours,
        })),
      });
    }
  }
  for (const row of pbj.balances) {
    jobs.push(manualToCompleted('balances', row));
  }
  for (const row of pbj.utPayroll) {
    jobs.push(manualToCompleted('ut', row));
  }
  return {
    timesheet,
    // Round 10 — the session's own resolved customer names, keyed by AP#.
    customerNameByAp: new Map(
      jobs.filter((j) => j.apKey && j.customer).map((j) => [j.apKey, j.customer] as const),
    ),
    jobs,
    dates: pbj.dates,
    dateIsos: pbj.dateIsos,
    title: pbj.title,
    lastWeekPbj:
      lastWeekPbj ??
      jobs.map((j) => ({
        section: j.section === 'nv' ? 'pivot' : j.section,
        customer: j.customer,
        job: j.job,
        apKey: j.apKey,
        estimator: j.estimator,
        crewLead: j.crewLead,
        woOverride: j.woOverride,
        woPct: j.woPct,
        percent: j.percent,
        notes: j.notes,
        excludedFromPayroll: j.excludedFromPayroll,
      })),
  };
}

function manualToCompleted(
  section: 'balances' | 'ut',
  row: ManualSectionRow,
): CompletedPbjJob {
  return {
    section,
    customer: row.customer,
    job: row.job,
    apKey: row.apKey ?? '',
    onCalendar: '-',
    estimator: row.estimator,
    crewLead: row.crewLead,
    split: 1,
    woOverride: row.woOverride,
    woPct: row.woPct,
    percent: row.percent,
    notes: row.notes,
    hourly: row.percent === 'HOURLY' || /^hourly$/i.test(row.crewLead),
    excludedFromPayroll: row.excludedFromPayroll,
    byDate: {},
    total: 0,
    employees: [],
  };
}

function Placeholder({ mode, ready }: { mode: AppMode; ready: boolean }): JSX.Element {
  const isPayroll = mode === 'payroll';
  return (
    <div className="flex h-full items-center justify-center">
      <div className="max-w-md rounded-lg border border-neutral-300 bg-white p-6 text-center shadow-sm">
        <div className="mb-2 text-4xl">{isPayroll ? '💵' : '📋'}</div>
        <h2 className="mb-1 text-base font-bold">
          {isPayroll ? 'Build the weekly payroll preview' : 'Build the weekly % of Completion report'}
        </h2>
        <p className="text-xs leading-relaxed text-neutral-500">
          {isPayroll
            ? ready
              ? 'Press "Create Payroll Preview" to build the payroll tabs.'
              : 'Upload Slot A (completed %\u2019s — or use this session\u2019s report) and Slot B (P&L), check the pay date, then create.'
            : ready
              ? 'Press "Create %\u2019s Report" to build the tables for each item in the navigation menu above.'
              : 'Upload required files 1 and 3–7 on the left (2 and 8 are optional), check the pay date, then press Create %\u2019s Report. Edits recompute live.'}
        </p>
      </div>
    </div>
  );
}

/** Full P&L account tree when available; else byJob summary metrics. */
function PlImportTable({ pl }: { pl: PlParseResult | null }): JSX.Element {
  if (!pl) {
    return (
      <div className="p-4 text-xs text-neutral-500 italic">Upload the P&amp;L by Customer file to populate this tab.</div>
    );
  }

  if (pl.columnHeaders.length > 0 && pl.accountRows.length > 0) {
    const headers = [...pl.columnHeaders];
    if (!headers[0]) headers[0] = 'Account';
    return (
      <div className="p-2 overflow-auto">
        <table className="w-max border-collapse bg-white text-xs">
          <thead>
            <tr>
              {headers.map((h, i) => (
                <th key={`${i}-${h}`} className="xcell2 text-left font-bold whitespace-nowrap">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pl.accountRows.map((ar, ri) => (
              <tr key={`${ri}-${ar.label}`}>
                <td className="xcell font-bold whitespace-nowrap">{ar.label}</td>
                {headers.slice(1).map((_, ci) => {
                  const v = ar.values[ci + 1];
                  return (
                    <td key={ci} className="xcell text-right whitespace-nowrap">
                      {v === null || v === undefined ? '' : v.toFixed(2)}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  const jobs = [...pl.byJob.entries()];
  const headers = ['Account', ...jobs.map(([ap, m]) => m.plJobHeader || `AP${ap}`)];
  const metricRows: { label: string; pick: (m: (typeof jobs)[0][1]) => number }[] = [
    { label: 'Total Income', pick: (m) => m.income },
    { label: 'Materials', pick: (m) => m.materials },
    { label: 'QB Labor-Wages-Subs', pick: (m) => m.laborWagesSubs },
    { label: 'Mileage Expense', pick: (m) => m.mileage },
    { label: 'Sundries -', pick: (m) => m.sundries },
  ];
  return (
    <div className="p-2 overflow-auto">
      <table className="w-max border-collapse bg-white text-xs">
        <thead>
          <tr>
            {headers.map((h) => (
              <th key={h} className="xcell2 text-left font-bold whitespace-nowrap">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {metricRows.map((mr) => (
            <tr key={mr.label}>
              <td className="xcell font-bold whitespace-nowrap">{mr.label}</td>
              {jobs.map(([ap, m]) => (
                <td key={ap} className="xcell text-right whitespace-nowrap">
                  {mr.pick(m).toFixed(2)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
