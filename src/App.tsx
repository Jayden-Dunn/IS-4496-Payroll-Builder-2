import { useCallback, useEffect, useMemo, useState } from 'react';
import { ExceptionsPanel } from './components/ExceptionsPanel';
import { PayrollPreview } from './components/PayrollPreview';
import { PivotByJobTable } from './components/PivotByJobTable';
import { PivotTable } from './components/PivotTable';
import { SettingsModal } from './components/SettingsModal';
import { TimeNaviTable } from './components/TimeNaviTable';
import { TopBar } from './components/TopBar';
import { UploadSlot, type SlotChip } from './components/UploadSlot';
import { timesheetApKey } from './lib/apkey';
import { buildExceptions, countExceptions } from './lib/exceptions';
import { exportFilename, exportWorkbook } from './lib/exportXlsx';
import { addDays, dateToIso, isoToDate } from './lib/format';
import { enrichTimesheet } from './lib/joins';
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
import { buildPayrollPreview } from './lib/payrollPreview';
import { buildPivot } from './lib/pivot';
import { buildPivotByJob, EMPTY_MANUAL_STATE, type ManualSectionState } from './lib/pivotByJob';
import { ensureDefaultPivotColumns, loadSettings, saveSettings } from './lib/settings';
import { buildReportRows, sortReportRows } from './lib/timenavi';
import type {
  AppSettings,
  DirectoryEmployee,
  EditsMap,
  ManualSectionRow,
  PayrollEdit,
  PayrollEditsMap,
  PbjEditsMap,
  PbjJobEdit,
  PlParseResult,
  PrevPercent,
  QbHoursMap,
  RowEdit,
  SalesRow,
  TimeNaviEvent,
  TimesheetParseResult,
} from './lib/types';

// ---------------------------------------------------------------------------
// Upload slots
// ---------------------------------------------------------------------------

type SlotId =
  | 'timesheet' | 'timenavi' | 'nvSales' | 'nvSummary' | 'utSales' | 'utSummary' | 'lastWeek'
  | 'pl' | 'employees';

interface SlotDef {
  id: SlotId;
  label: string;
  hint: string;
  optional?: boolean;
}

const SLOT_DEFS: SlotDef[] = [
  { id: 'timesheet', label: '1 · QuickBooks Time timesheet', hint: 'TSheets weekly payroll export' },
  { id: 'timenavi', label: '2 · TimeNavi calendar export', hint: 'Reads the "Events" sheet' },
  { id: 'nvSales', label: '3 · NV Sales by Rep Detail', hint: 'QuickBooks Online (Nevada)' },
  { id: 'nvSummary', label: '4 · NV Sales by Customer Summary', hint: 'Job totals (Nevada)' },
  { id: 'utSales', label: '5 · UT Transaction List by Date', hint: 'AllPro Painters LLC (Utah)' },
  { id: 'utSummary', label: '6 · UT Sales by Customer Summary', hint: 'Job totals (Utah)' },
  { id: 'lastWeek', label: '7 · Last week\u2019s % of Completion', hint: 'Last % Paid + Balances/UT carry-forward', optional: true },
  { id: 'pl', label: '8 · P&L by Customer (QBO)', hint: 'Required for the payroll tabs', optional: true },
  { id: 'employees', label: '9 · Employee Name List (QB Time)', hint: 'Crew-lead name resolution', optional: true },
];

interface ParsedData {
  timesheet: TimesheetParseResult | null;
  events: TimeNaviEvent[] | null;
  nvSales: SalesRow[] | null;
  nvTotals: Map<string, number> | null;
  utSales: SalesRow[] | null;
  utTotals: Map<string, number> | null;
  lastWeek: Map<string, PrevPercent> | null;
  lastWeekPbj: LastWeekPbjRow[] | null;
  pl: PlParseResult | null;
  directory: DirectoryEmployee[] | null;
}

const EMPTY_PARSED: ParsedData = {
  timesheet: null,
  events: null,
  nvSales: null,
  nvTotals: null,
  utSales: null,
  utTotals: null,
  lastWeek: null,
  lastWeekPbj: null,
  pl: null,
  directory: null,
};

const EMPTY_CHIP: SlotChip = { status: 'empty', fileName: '', message: '' };

type Tab = 'timenavi' | 'pivot' | 'pivotByJob' | 'payroll' | 'exceptions' | 'payrollExceptions';

export default function App(): JSX.Element {
  const [chips, setChips] = useState<Record<SlotId, SlotChip>>({
    timesheet: EMPTY_CHIP,
    timenavi: EMPTY_CHIP,
    nvSales: EMPTY_CHIP,
    nvSummary: EMPTY_CHIP,
    utSales: EMPTY_CHIP,
    utSummary: EMPTY_CHIP,
    lastWeek: EMPTY_CHIP,
    pl: EMPTY_CHIP,
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
  const [created, setCreated] = useState(false);
  const [tab, setTab] = useState<Tab>('timenavi');
  const [settingsOpen, setSettingsOpen] = useState(false);

  // ------------------------------------------------------------------ uploads

  const handleFile = useCallback(async (id: SlotId, file: File): Promise<void> => {
    const fail = (message: string): void => {
      setChips((c) => ({ ...c, [id]: { status: 'error', fileName: file.name, message } }));
      setParsed((p) => ({ ...p, ...clearForSlot(id) }));
    };
    try {
      const wbd = readWorkbook(await file.arrayBuffer());
      const ok = (message: string, patch: Partial<ParsedData>): void => {
        setParsed((p) => ({ ...p, ...patch }));
        setChips((c) => ({ ...c, [id]: { status: 'ok', fileName: file.name, message } }));
      };
      switch (id) {
        case 'timesheet': {
          const res = parseTimesheet(wbd);
          ok(
            `✓ ${res.rows.length} rows${res.lunchRemoved > 0 ? ` · ${res.lunchRemoved} lunch removed` : ''}`,
            { timesheet: res },
          );
          // Default pay date = week-ending Friday (max local_date) + 7 days.
          let max: Date | null = null;
          for (const r of res.rows) {
            if (r.localDate && (!max || r.localDate.getTime() > max.getTime())) max = r.localDate;
          }
          if (max) setPayDateIso(dateToIso(addDays(max, 7)));
          break;
        }
        case 'timenavi': {
          const events = parseTimeNavi(wbd);
          ok(`✓ ${events.length} events`, { events });
          break;
        }
        case 'nvSales': {
          const rows = parseSalesDetail(wbd, 'customer');
          ok(`✓ ${rows.length} sales rows`, { nvSales: rows });
          break;
        }
        case 'nvSummary': {
          const totals = parseCustomerSummary(wbd);
          ok(`✓ ${totals.size} job totals`, { nvTotals: totals });
          break;
        }
        case 'utSales': {
          const rows = parseSalesDetail(wbd, 'name');
          ok(`✓ ${rows.length} transaction rows`, { utSales: rows });
          break;
        }
        case 'utSummary': {
          const totals = parseCustomerSummary(wbd);
          ok(`✓ ${totals.size} job totals`, { utTotals: totals });
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
          const pl = parsePl(wbd);
          ok(
            `✓ ${pl.byJob.size} job columns${pl.duplicateColumns.length > 0 ? ` · ${pl.duplicateColumns.length} duplicated` : ''}`,
            { pl },
          );
          break;
        }
        case 'employees': {
          const directory = parseEmployeeList(wbd);
          const active = directory.filter((d) => d.active).length;
          ok(`✓ ${active} active / ${directory.length - active} inactive`, { directory });
          break;
        }
      }
    } catch (err) {
      fail(err instanceof Error ? err.message : 'Could not read this file.');
    }
  }, []);

  const ready =
    parsed.timesheet !== null &&
    parsed.events !== null &&
    parsed.nvSales !== null &&
    parsed.nvTotals !== null &&
    parsed.utSales !== null &&
    parsed.utTotals !== null;

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

  /** apKey → jobcode_2 of the first timesheet row with that key (column M). */
  const timesheetJobcodeByApKey = useMemo(() => {
    const m = new Map<string, string>();
    if (parsed.timesheet) {
      for (const r of parsed.timesheet.rows) {
        const key = timesheetApKey(r.jobcode2);
        if (key && !m.has(key)) m.set(key, r.jobcode2);
      }
    }
    return m;
  }, [parsed.timesheet]);

  const reportRows = useMemo(() => {
    if (!created || !ready || !parsed.events || !parsed.nvSales || !parsed.utSales || !parsed.nvTotals || !parsed.utTotals) {
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
        lastWeekMap: parsed.lastWeek,
        timesheetJobcodeByApKey,
        knownEstimators,
      },
      payDate,
    );
    return sortReportRows(rows, legendOrder);
  }, [created, ready, parsed, edits, timesheetJobcodeByApKey, knownEstimators, legendOrder, payDate]);

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

  const enriched = useMemo(() => {
    if (!created || !parsed.timesheet || !parsed.nvSales || !parsed.utSales) return [];
    return enrichTimesheet(parsed.timesheet.rows, parsed.nvSales, parsed.utSales, includedApKeys, estimatorByApKey);
  }, [created, parsed.timesheet, parsed.nvSales, parsed.utSales, includedApKeys, estimatorByApKey]);

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

  const exceptions = useMemo(
    () =>
      buildExceptions({
        reportRows,
        enriched,
        lunchRemoved: parsed.timesheet?.lunchRemoved ?? 0,
        payDate,
      }),
    [reportRows, enriched, parsed.timesheet, payDate],
  );
  const exceptionCount = useMemo(() => countExceptions(exceptions), [exceptions]);

  // ------------------------------------------------ Round 3: payroll models

  const leadCtx = useMemo<LeadResolverContext>(() => {
    const crewGroups = [...new Set((parsed.timesheet?.rows ?? []).map((r) => r.group).filter(Boolean))];
    return {
      directory: parsed.directory ?? directoryFromTimesheet(parsed.timesheet?.rows ?? []),
      crewGroups,
      subcontractors: settings.subcontractors,
    };
  }, [parsed.directory, parsed.timesheet, settings.subcontractors]);

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
        lastWeekPbj: parsed.lastWeekPbj,
        balancesState,
        utState,
      }),
    [enriched, reportRows, settings, leadCtx, knownEstimators, payDate, pbjEdits, qbHours, parsed.lastWeekPbj, balancesState, utState],
  );

  const payrollPreview = useMemo(() => {
    if (!created || !parsed.pl) return null;
    return buildPayrollPreview({
      pbj: pivotByJob,
      enriched,
      reportRows,
      pl: parsed.pl,
      directory: parsed.directory,
      settings,
      payDate,
      edits: payrollEdits,
      qbHours,
      knownEstimators,
      leadCtx,
    });
  }, [created, parsed.pl, parsed.directory, pivotByJob, enriched, reportRows, settings, payDate, payrollEdits, qbHours, knownEstimators, leadCtx]);

  const payrollExceptions = useMemo(() => {
    if (!created) return [];
    return buildPayrollExceptions({
      pbj: pivotByJob,
      preview:
        payrollPreview ?? {
          title: '',
          painters: [],
          leads: [],
          office: [],
          admin: [],
          estimators: [],
          managers: [],
          inactiveZeroHours: [],
        },
      reportRows,
      pl: parsed.pl,
      hasDirectory: parsed.directory !== null,
    });
  }, [created, pivotByJob, payrollPreview, reportRows, parsed.pl, parsed.directory]);
  const payrollExceptionCount = useMemo(() => countExceptions(payrollExceptions), [payrollExceptions]);

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

  const onDownload = useCallback((): void => {
    void exportWorkbook({
      rows: reportRows,
      pivot,
      legend: settings.legend,
      payDate,
      pivotByJob,
      payrollPreview,
    });
  }, [reportRows, pivot, settings.legend, payDate, pivotByJob, payrollPreview]);

  const onSaveSettings = useCallback((next: AppSettings): void => {
    setSettings(next);
    saveSettings(next);
    setSettingsOpen(false);
  }, []);

  // ------------------------------------------------------------------ render

  const includedCount = reportRows.filter((r) => r.included).length;

  return (
    <div className="flex h-full flex-col">
      <TopBar
        canDownload={created && ready}
        onDownload={onDownload}
        onOpenSettings={() => setSettingsOpen(true)}
        downloadName={exportFilename(payDate)}
      />
      <div className="flex min-h-0 flex-1">
        {/* Left rail — uploads */}
        <aside className="flex w-[320px] shrink-0 flex-col gap-2 overflow-y-auto border-r border-neutral-300 bg-neutral-50 p-3">
          <p className="text-[11px] leading-snug text-neutral-500">
            Upload the weekly reports (<b>.xlsx</b> or <b>.csv</b>). Everything is parsed locally in your browser — no
            data leaves this computer.
          </p>
          {SLOT_DEFS.map((def) => (
            <UploadSlot
              key={def.id}
              label={def.label}
              hint={def.hint}
              optional={def.optional}
              chip={chips[def.id]}
              onFile={(file) => void handleFile(def.id, file)}
            />
          ))}
          <label className="mt-1 flex flex-col gap-1 text-xs font-semibold">
            Pay date
            <input
              type="date"
              className="rounded border border-neutral-300 px-2 py-1.5 font-normal"
              value={payDateIso}
              onChange={(e) => setPayDateIso(e.target.value)}
            />
            <span className="text-[10px] font-normal text-neutral-400">
              Defaults to the timesheet week-ending Friday + 7 days. Drives the pivot title and the filename.
            </span>
          </label>
          <button
            type="button"
            className="btn-primary mt-1 py-2.5 text-base"
            disabled={!ready}
            title={ready ? 'Build both output tables' : 'Upload files 1–6 first'}
            onClick={() => {
              setCreated(true);
              setTab('timenavi');
            }}
          >
            Create Report
          </button>
        </aside>

        {/* Main area — tabs */}
        <main className="flex min-w-0 flex-1 flex-col">
          <nav className="flex shrink-0 gap-1 border-b border-neutral-300 bg-white px-3 pt-2">
            {(
              [
                ['timenavi', `TimeNavi Export${created ? ` (${includedCount})` : ''}`],
                ['pivot', 'Pivot by Crew'],
                ['pivotByJob', 'Pivot by Job'],
                ['payroll', 'Payroll Preview'],
                ['exceptions', `%'s Exceptions${created ? ` (${exceptionCount})` : ''}`],
                ['payrollExceptions', `Payroll Exceptions${created ? ` (${payrollExceptionCount})` : ''}`],
              ] as [Tab, string][]
            ).map(([t, label]) => (
              <button
                key={t}
                type="button"
                className={`rounded-t-md border border-b-0 px-4 py-1.5 text-sm font-medium ${
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
            {!created ? (
              <Placeholder ready={ready} />
            ) : tab === 'timenavi' ? (
              <TimeNaviTable
                rows={reportRows}
                colorByEstimator={colorByEstimator}
                estimatorOptions={estimatorOptions}
                onEdit={onEdit}
              />
            ) : tab === 'pivot' ? (
              <PivotTable pivot={pivot} colorByEstimator={colorByEstimator} />
            ) : tab === 'pivotByJob' ? (
              <PivotByJobTable
                model={pivotByJob}
                onEditJob={onEditPbjJob}
                onEditQbHours={onEditQbHours}
                onManualEdit={onManualEdit}
                onManualAdd={onManualAdd}
                onManualDelete={onManualDelete}
              />
            ) : tab === 'payroll' ? (
              <PayrollPreview model={payrollPreview} mileageRate={settings.mileageRate} onEdit={onPayrollEdit} />
            ) : tab === 'exceptions' ? (
              <ExceptionsPanel groups={exceptions} onNavigate={navigateToRow} />
            ) : (
              <ExceptionsPanel groups={payrollExceptions} onNavigate={navigateToRow} />
            )}
          </div>
        </main>
      </div>

      {settingsOpen && (
        <SettingsModal settings={settings} onSave={onSaveSettings} onClose={() => setSettingsOpen(false)} />
      )}
    </div>
  );
}

function clearForSlot(id: SlotId): Partial<ParsedData> {
  switch (id) {
    case 'timesheet':
      return { timesheet: null };
    case 'timenavi':
      return { events: null };
    case 'nvSales':
      return { nvSales: null };
    case 'nvSummary':
      return { nvTotals: null };
    case 'utSales':
      return { utSales: null };
    case 'utSummary':
      return { utTotals: null };
    case 'lastWeek':
      return { lastWeek: null, lastWeekPbj: null };
    case 'pl':
      return { pl: null };
    case 'employees':
      return { directory: null };
  }
}

function Placeholder({ ready }: { ready: boolean }): JSX.Element {
  return (
    <div className="flex h-full items-center justify-center">
      <div className="max-w-md rounded-lg border border-neutral-300 bg-white p-6 text-center shadow-sm">
        <div className="mb-2 text-4xl">📋</div>
        <h2 className="mb-1 text-base font-bold">Build the weekly % of Completion report</h2>
        <p className="text-xs leading-relaxed text-neutral-500">
          {ready
            ? 'All required files parsed. Press Create Report to build the TimeNavi Export and Pivot by Crew tables.'
            : 'Upload files 1–6 on the left (file 7 is optional), check the pay date, then press Create Report. You can edit AP#, % Completion, Estimator, Apptivo $ and include/exclude rows afterwards — both tables and the exceptions list recompute live.'}
        </p>
      </div>
    </div>
  );
}
