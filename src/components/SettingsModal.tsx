import { useRef, useState } from 'react';
import {
  DEFAULT_CONDITIONAL_RATE_RULES,
  DEFAULT_CREW_LEAD_RATE,
  DEFAULT_DAILY_OT_HOURS_THRESHOLD,
  DEFAULT_LEGEND,
  DEFAULT_NV_DAILY_OT_RATE_THRESHOLD,
  DEFAULT_DAILY_OT_EXEMPT,
  DEFAULT_FULL_MATERIALS_LEADS,
  DEFAULT_PAINTER_ROSTER,
  DEFAULT_PAYROLL_LINE_ITEMS,
  DEFAULT_PER_JOB_RATE_OVERRIDES,
  DEFAULT_PIVOT_COLUMNS,
  DEFAULT_PL_ACCOUNT_MAPPING,
  DEFAULT_PREVAILING_MIN_RATE,
  DEFAULT_PREVAILING_PATTERNS,
  DEFAULT_PROFIT_SHARE,
  DEFAULT_SALARIED,
  DEFAULT_SUBCONTRACTORS,
  DEFAULT_WAGE_RATES,
  type DirectoryStoreStatus,
  type SettingsStoreStatus,
} from '../lib/settings';
import type {
  AppSettings,
  ConditionalRateRule,
  EntityCode,
  JobRateRule,
  LegendEntry,
  PainterRosterEntry,
  PayrollLineItem,
  PerJobRateOverride,
  PlAccountMapping,
  ProfitShareEntry,
  SalariedEntry,
  SalarySection,
  SavedDirectory,
  WageRateEntry,
} from '../lib/types';

interface SettingsModalProps {
  settings: AppSettings;
  /** Persisted employee directory (null = never saved). */
  directory: SavedDirectory | null;
  /** Shared store probe (Vercel Blob API) — null while loading. */
  directoryStoreStatus?: DirectoryStoreStatus | null;
  /** Set when local save worked but cross-device sync failed. */
  directorySyncWarning?: string | null;
  /** Round 8 §1 — shared Settings store probe; null while loading. */
  settingsStoreStatus?: SettingsStoreStatus | null;
  /** Round 8 §1 — set when the local save worked but the shared write failed. */
  settingsSyncWarning?: string | null;
  /** Round 8 §1 — pull the shared Settings again, discarding unsaved edits. */
  onReloadSettings?: () => void | Promise<void>;
  /** Parse + persist a replacement list; resolves to an error message or null. */
  onDirectoryReplace: (file: File) => Promise<string | null>;
  onDirectoryClear: () => void;
  onSave: (settings: AppSettings) => void;
  onClose: () => void;
}

const PL_METRICS: { key: keyof PlAccountMapping; label: string }[] = [
  { key: 'income', label: 'Income' },
  { key: 'materials', label: 'Materials' },
  { key: 'laborWagesSubs', label: 'Labor / Wages / Subs' },
  { key: 'sundries', label: 'Sundries' },
  { key: 'travel', label: 'Travel' },
  { key: 'mileage', label: 'Mileage' },
];

const PL_ENTITIES: EntityCode[] = ['NV', 'UT'];

function clonePlMapping(src: Record<EntityCode, PlAccountMapping>): Record<EntityCode, PlAccountMapping> {
  const one = (m: PlAccountMapping): PlAccountMapping => ({
    income: [...m.income],
    materials: [...m.materials],
    laborWagesSubs: [...m.laborWagesSubs],
    sundries: [...m.sundries],
    travel: [...m.travel],
    mileage: [...m.mileage],
  });
  return { NV: one(src.NV), UT: one(src.UT) };
}

const HEX_RE = /^[0-9A-Fa-f]{6}$/;

function moveItem<T>(list: T[], i: number, dir: -1 | 1): T[] {
  const j = i + dir;
  if (j < 0 || j >= list.length) return list;
  const next = [...list];
  const tmp = next[i];
  next[i] = next[j];
  next[j] = tmp;
  return next;
}

function SectionTitle({ children, sub }: { children: string; sub?: string }): JSX.Element {
  return (
    <>
      <h3 className="mb-1 mt-6 text-xs font-bold uppercase tracking-wide text-neutral-600 first:mt-0">{children}</h3>
      {sub && <p className="mb-3 text-[11px] text-neutral-500">{sub}</p>}
    </>
  );
}

export function SettingsModal({
  settings,
  directory,
  directoryStoreStatus: _directoryStoreStatus = null,
  directorySyncWarning = null,
  settingsStoreStatus = null,
  settingsSyncWarning = null,
  onReloadSettings,
  onDirectoryReplace,
  onDirectoryClear,
  onSave,
  onClose,
}: SettingsModalProps): JSX.Element {
  const [entries, setEntries] = useState<LegendEntry[]>(settings.legend.map((l) => ({ ...l })));
  const [pivotCols, setPivotCols] = useState<string[]>([...settings.pivotColumnOrder]);
  const [newPivotCol, setNewPivotCol] = useState('');
  const [ps, setPs] = useState<ProfitShareEntry[]>(settings.profitShare.map((e) => ({ ...e })));
  const [defaultPs, setDefaultPs] = useState(String(Math.round(settings.defaultProfitSharePct * 100)));
  const [wages, setWages] = useState<WageRateEntry[]>(settings.wageRates.map((e) => ({ ...e })));
  // Round 9 §19/§23/§24/§29 — person-scoped extra payroll rows.
  const [lineItems, setLineItems] = useState<PayrollLineItem[]>(
    (settings.payrollLineItems ?? []).map((e) => ({ ...e })),
  );
  /** Round 9 §32 — leads charged for all of a job's materials. */
  const [fullMatLeads, setFullMatLeads] = useState<string[]>([...(settings.fullMaterialsLeads ?? [])]);
  /** Round 9 §27 — people exempt from NV daily overtime. */
  const [dailyOtExempt, setDailyOtExempt] = useState<string[]>([...(settings.dailyOtExemptNames ?? [])]);
  const [defaultWage, setDefaultWage] = useState(String(settings.defaultWageRate));
  const [defaultCrewLead, setDefaultCrewLead] = useState(String(settings.defaultCrewLeadRate));
  const [nvOtRate, setNvOtRate] = useState(String(settings.nvDailyOtRateThreshold));
  const [otHours, setOtHours] = useState(String(settings.dailyOtHoursThreshold));
  const [condRules, setCondRules] = useState<ConditionalRateRule[]>(
    settings.conditionalRateRules.map((r) => ({ ...r, condition: { ...r.condition } })),
  );
  const [perJobRules, setPerJobRules] = useState<PerJobRateOverride[]>(
    settings.perJobRateOverrides.map((r) => ({ ...r })),
  );
  const [plMap, setPlMap] = useState<Record<EntityCode, PlAccountMapping>>(() =>
    clonePlMapping(settings.plAccountMapping),
  );
  const [painters, setPainters] = useState<PainterRosterEntry[]>(settings.painterRoster.map((p) => ({ ...p })));
  const [paintersOpen, setPaintersOpen] = useState(false);
  const [painterQuery, setPainterQuery] = useState('');
  const [newPainter, setNewPainter] = useState('');
  const [roster, setRoster] = useState<SalariedEntry[]>(settings.salariedRoster.map((e) => ({ ...e })));
  const [mileageRate, setMileageRate] = useState(String(settings.mileageRate));
  const [subs, setSubs] = useState<string[]>([...settings.subcontractors]);
  const [pwPatterns, setPwPatterns] = useState<string[]>([...settings.prevailingWagePatterns]);
  const [pwMin, setPwMin] = useState(String(settings.prevailingMinRate));
  const [officeMileage, setOfficeMileage] = useState(String(settings.officeMileageRate));
  const [jobRules, setJobRules] = useState<JobRateRule[]>(settings.jobRateRules.map((r) => ({ ...r, employeeNames: [...r.employeeNames] })));
  const [dirError, setDirError] = useState<string | null>(null);
  const [dirBusy, setDirBusy] = useState(false);
  const dirInputRef = useRef<HTMLInputElement>(null);

  const replaceDirectory = async (file: File): Promise<void> => {
    setDirBusy(true);
    setDirError(null);
    const err = await onDirectoryReplace(file);
    setDirError(err);
    setDirBusy(false);
  };

  const updateEntry = (i: number, patch: Partial<LegendEntry>): void => {
    setEntries((prev) => prev.map((e, j) => (j === i ? { ...e, ...patch } : e)));
  };

  const cleanedLegend = entries
    .map((e) => ({ estimator: e.estimator.trim().toUpperCase(), color: e.color.trim().toUpperCase() }))
    .filter((e) => e.estimator !== '');
  const legendValid =
    cleanedLegend.length > 0 &&
    cleanedLegend.every((e) => HEX_RE.test(e.color)) &&
    new Set(cleanedLegend.map((e) => e.estimator)).size === cleanedLegend.length;

  const cleanedPivot = pivotCols.map((e) => e.trim().toUpperCase()).filter((e) => e !== '');
  const pivotValid = cleanedPivot.length > 0 && new Set(cleanedPivot).size === cleanedPivot.length;

  const defaultPsNum = Number(defaultPs);
  const defaultWageNum = Number(defaultWage);
  const defaultCrewLeadNum = Number(defaultCrewLead);
  const nvOtRateNum = Number(nvOtRate);
  const otHoursNum = Number(otHours);
  const mileageNum = Number(mileageRate);
  const officeMileageNum = Number(officeMileage);
  const pwMinNum = Number(pwMin);
  const numbersValid =
    Number.isFinite(defaultPsNum) &&
    Number.isFinite(defaultWageNum) &&
    Number.isFinite(defaultCrewLeadNum) &&
    Number.isFinite(nvOtRateNum) &&
    Number.isFinite(otHoursNum) &&
    Number.isFinite(mileageNum) &&
    Number.isFinite(officeMileageNum) &&
    Number.isFinite(pwMinNum);

  const painterMatches = painters
    .map((p, i) => ({ p, i }))
    .filter(({ p }) => {
      const q = painterQuery.trim().toLowerCase();
      if (!q) return true;
      return p.name.toLowerCase().includes(q) || (p.note ?? '').toLowerCase().includes(q);
    });
  const activePainterCount = painters.filter((p) => p.active).length;

  const setPlList = (entity: EntityCode, metric: keyof PlAccountMapping, raw: string): void => {
    const list = raw
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
    setPlMap((prev) => ({ ...prev, [entity]: { ...prev[entity], [metric]: list } }));
  };

  const addPivotCol = (): void => {
    const t = newPivotCol.trim().toUpperCase();
    if (t && !pivotCols.includes(t)) setPivotCols((prev) => [...prev, t]);
    setNewPivotCol('');
  };

  const save = (): void => {
    onSave({
      legend: cleanedLegend,
      pivotColumnOrder: cleanedPivot,
      profitShare: ps.filter((e) => e.lead.trim() !== ''),
      defaultProfitSharePct: defaultPsNum > 1 ? defaultPsNum / 100 : defaultPsNum,
      wageRates: wages.filter((e) => e.name.trim() !== ''),
      defaultWageRate: defaultWageNum,
      defaultCrewLeadRate: defaultCrewLeadNum,
      nvDailyOtRateThreshold: nvOtRateNum,
      dailyOtHoursThreshold: otHoursNum,
      conditionalRateRules: condRules
        .map((r) => ({
          ...r,
          person: r.person.trim(),
          condition: { ...r.condition, leadName: r.condition.leadName.trim() },
        }))
        .filter((r) => r.person !== '' && r.condition.leadName !== '' && Number.isFinite(r.rate)),
      perJobRateOverrides: perJobRules
        .map((r) => ({ ...r, person: r.person.trim(), apKeyPattern: r.apKeyPattern.trim() }))
        .filter((r) => r.person !== '' && r.apKeyPattern !== '' && Number.isFinite(r.rate)),
      plAccountMapping: clonePlMapping(plMap),
      painterRoster: painters
        .map((p) => ({ ...p, name: p.name.trim() }))
        .filter((p) => p.name !== ''),
      salariedRoster: roster.filter((e) => e.name.trim() !== ''),
      mileageRate: mileageNum,
      officeMileageRate: officeMileageNum,
      subcontractors: subs.map((s) => s.trim()).filter(Boolean),
      prevailingWagePatterns: pwPatterns.map((s) => s.trim()).filter(Boolean),
      prevailingMinRate: pwMinNum,
      jobRateRules: jobRules
        .map((r) => ({
          ...r,
          apKey: r.apKey.trim(),
          note: r.note.trim(),
          employeeNames: r.employeeNames.map((n) => n.trim()).filter(Boolean),
        }))
        .filter((r) => r.apKey !== '' && Number.isFinite(r.rate)),
      // Round 9 §19/§23/§24/§29/§32/§27.
      payrollLineItems: lineItems
        .map((r) => ({ ...r, person: r.person.trim(), label: r.label.trim() }))
        .filter((r) => r.person !== '' && r.label !== ''),
      fullMaterialsLeads: fullMatLeads.map((n) => n.trim()).filter(Boolean),
      dailyOtExemptNames: dailyOtExempt.map((n) => n.trim()).filter(Boolean),
      appliedMigrations: [...settings.appliedMigrations],
    });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div
        className="w-full max-w-2xl rounded-lg bg-white shadow-xl"
        onClick={(e) => e.stopPropagation()}
        role="dialog"
        aria-label="Settings"
      >
        <header className="flex items-center justify-between border-b border-neutral-200 px-4 py-3">
          <h2 className="text-sm font-bold">Settings</h2>
          <button type="button" className="text-neutral-400 hover:text-neutral-700" onClick={onClose} aria-label="Close">
            ✕
          </button>
        </header>
        <div className="max-h-[70vh] overflow-y-auto p-4">
          <div className="mb-3 flex items-center gap-2 rounded border border-neutral-200 bg-neutral-50 px-2 py-1.5 text-[11px]">
            <span className="text-neutral-600">
              {settingsStoreStatus === 'ready'
                ? 'Settings are shared — saving here updates every browser on this deployment.'
                : settingsStoreStatus === 'empty'
                  ? 'Shared store is empty — your next save becomes the shared copy.'
                  : settingsStoreStatus === null
                    ? 'Checking the shared Settings store…'
                    : 'Shared Settings unavailable — saving to this browser only (optional feature).'}
            </span>
            {onReloadSettings && (
              <button type="button" className="btn ml-auto text-[11px]" onClick={() => void onReloadSettings()}>
                Reload settings
              </button>
            )}
          </div>
          {settingsSyncWarning && (
            <div className="mb-3 rounded border border-amber-300 bg-amber-50 px-2 py-1.5 text-[11px] text-amber-900">
              {settingsSyncWarning}
            </div>
          )}

          <SectionTitle sub="This order drives row sorting on the TimeNavi Export. Colors are the fills on all tabs.">
            Estimator color legend &amp; row-sort order
          </SectionTitle>
          <div className="flex flex-col gap-1.5">
            {entries.map((e, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <div className="flex flex-col">
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === 0} onClick={() => setEntries((prev) => moveItem(prev, i, -1))} aria-label="Move up">▲</button>
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === entries.length - 1} onClick={() => setEntries((prev) => moveItem(prev, i, 1))} aria-label="Move down">▼</button>
                </div>
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-xs font-semibold uppercase" value={e.estimator} maxLength={6} placeholder="Init." onChange={(ev) => updateEntry(i, { estimator: ev.target.value.toUpperCase() })} />
                <input type="color" className="h-7 w-9 cursor-pointer rounded border border-neutral-300" value={HEX_RE.test(e.color) ? `#${e.color}` : '#cccccc'} onChange={(ev) => updateEntry(i, { color: ev.target.value.slice(1).toUpperCase() })} />
                <input className={`w-20 rounded border px-1.5 py-1 font-mono text-xs uppercase ${HEX_RE.test(e.color) ? 'border-neutral-300' : 'border-red-400 bg-red-50'}`} value={e.color} maxLength={6} onChange={(ev) => updateEntry(i, { color: ev.target.value })} />
                <div className="h-6 flex-1 rounded border border-neutral-200" style={{ backgroundColor: HEX_RE.test(e.color) ? `#${e.color}` : undefined }} />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setEntries((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove estimator">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-3 flex gap-2">
            <button type="button" className="btn text-xs" onClick={() => setEntries((prev) => [...prev, { estimator: '', color: 'CCCCCC' }])}>+ Add estimator</button>
            <button type="button" className="btn text-xs" onClick={() => setEntries(DEFAULT_LEGEND.map((l) => ({ ...l })))}>Reset to defaults</button>
          </div>

          <SectionTitle sub="Column order on the Pivot by Crew tab and sheet (separate from the row-sort order above).">
            Pivot estimator column order
          </SectionTitle>
          <div className="flex flex-col gap-1.5">
            {pivotCols.map((col, i) => (
              <div key={`${col}-${i}`} className="flex items-center gap-1.5">
                <div className="flex flex-col">
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === 0} onClick={() => setPivotCols((prev) => moveItem(prev, i, -1))} aria-label="Move up">▲</button>
                  <button type="button" className="text-[10px] leading-none text-neutral-400 hover:text-neutral-800 disabled:opacity-30" disabled={i === pivotCols.length - 1} onClick={() => setPivotCols((prev) => moveItem(prev, i, 1))} aria-label="Move down">▼</button>
                </div>
                <span className="w-16 rounded border border-neutral-200 bg-neutral-50 px-1.5 py-1 text-xs font-semibold">{col}</span>
                <span className="flex-1" />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setPivotCols((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove column">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-3 flex items-center gap-2">
            <input className="w-20 rounded border border-neutral-300 px-1.5 py-1 text-xs font-semibold uppercase" value={newPivotCol} maxLength={6} placeholder="Init." onChange={(e) => setNewPivotCol(e.target.value.toUpperCase())} onKeyDown={(e) => { if (e.key === 'Enter') addPivotCol(); }} />
            <button type="button" className="btn text-xs" onClick={addPivotCol}>+ Add</button>
            <button type="button" className="btn text-xs" onClick={() => setPivotCols([...DEFAULT_PIVOT_COLUMNS])}>Reset to defaults</button>
          </div>

          <SectionTitle sub="WO% defaults per crew lead on Pivot by Job and Payroll Preview. A description line like '@ 40% PS' still overrides per job.">
            Crew profit-share %
          </SectionTitle>
          <div className="flex flex-col gap-1">
            {ps.map((e, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={e.lead} placeholder="Crew lead / sub name" onChange={(ev) => setPs((prev) => prev.map((x, j) => (j === i ? { ...x, lead: ev.target.value } : x)))} />
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={Math.round(e.pct * 100)} onChange={(ev) => { const n = Number(ev.target.value); if (Number.isFinite(n)) setPs((prev) => prev.map((x, j) => (j === i ? { ...x, pct: n / 100 } : x))); }} />
                <span className="text-xs text-neutral-400">%</span>
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setPs((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex items-center gap-2">
            <button type="button" className="btn text-xs" onClick={() => setPs((prev) => [...prev, { lead: '', pct: 0.38 }])}>+ Add crew</button>
            <button type="button" className="btn text-xs" onClick={() => setPs(DEFAULT_PROFIT_SHARE.map((e) => ({ ...e })))}>Reset</button>
            <span className="ml-auto text-xs">Default for unlisted crews:</span>
            <input className="w-14 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={defaultPs} onChange={(e) => setDefaultPs(e.target.value)} />
            <span className="text-xs text-neutral-400">%</span>
          </div>

          <SectionTitle sub="Five rates per person: PS = profit-share/work-order $/hr, hrly = hourly (T&M) jobs, lead = a crew lead's own hours on profit-share jobs, comm = the &quot;Commercial Hourly&quot; pay type, wg2 = the &quot;Wages 2&quot; pay type. The timesheet's pay item picks the column; blank falls back to PS, then to the defaults below. OT lines are always 1.5× the line rate.">
            Wage rates
          </SectionTitle>
          <div className="mb-1 flex items-center gap-1.5 text-[10px] font-semibold uppercase tracking-wide text-neutral-500">
            <span className="flex-1">Employee</span>
            <span className="w-16 text-right">PS</span>
            <span className="w-16 text-right">hrly</span>
            <span className="w-16 text-right">lead</span>
            <span className="w-16 text-right">comm</span>
            <span className="w-16 text-right">wg2</span>
            <span className="w-4" />
          </div>
          <div className="flex max-h-64 flex-col gap-1 overflow-y-auto rounded border border-neutral-100 p-1">
            {wages.map((e, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={e.name} placeholder="Employee" onChange={(ev) => setWages((prev) => prev.map((x, j) => (j === i ? { ...x, name: ev.target.value } : x)))} />
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.rate} onChange={(ev) => { const n = Number(ev.target.value); if (Number.isFinite(n)) setWages((prev) => prev.map((x, j) => (j === i ? { ...x, rate: n } : x))); }} />
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.hourlyRate ?? ''} placeholder="hrly" onChange={(ev) => { const t = ev.target.value.trim(); const n = Number(t); setWages((prev) => prev.map((x, j) => (j === i ? { ...x, ...(t === '' || !Number.isFinite(n) ? { hourlyRate: undefined } : { hourlyRate: n }) } : x))); }} />
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.crewLeadRate ?? ''} placeholder="lead" title="Crew-lead rate on their own timesheet hours on profit-share jobs" onChange={(ev) => { const t = ev.target.value.trim(); const n = Number(t); setWages((prev) => prev.map((x, j) => (j === i ? { ...x, ...(t === '' || !Number.isFinite(n) ? { crewLeadRate: undefined } : { crewLeadRate: n }) } : x))); }} />
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.commercialHourlyRate ?? ''} placeholder="comm" title="Rate on timesheet lines whose pay type is Commercial Hourly" onChange={(ev) => { const t = ev.target.value.trim(); const n = Number(t); setWages((prev) => prev.map((x, j) => (j === i ? { ...x, ...(t === '' || !Number.isFinite(n) ? { commercialHourlyRate: undefined } : { commercialHourlyRate: n }) } : x))); }} />
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.wages2Rate ?? ''} placeholder="wg2" title="Rate on timesheet lines whose pay type is Wages 2" onChange={(ev) => { const t = ev.target.value.trim(); const n = Number(t); setWages((prev) => prev.map((x, j) => (j === i ? { ...x, ...(t === '' || !Number.isFinite(n) ? { wages2Rate: undefined } : { wages2Rate: n }) } : x))); }} />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setWages((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex flex-wrap items-center gap-2">
            <button type="button" className="btn text-xs" onClick={() => setWages((prev) => [...prev, { name: '', rate: 18 }])}>+ Add employee</button>
            <button type="button" className="btn text-xs" onClick={() => setWages(DEFAULT_WAGE_RATES.map((e) => ({ ...e })))}>Reset</button>
            <span className="ml-auto text-xs">Fallback rate $:</span>
            <input className="w-14 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={defaultWage} onChange={(e) => setDefaultWage(e.target.value)} />
            <span className="text-xs">Fallback crew-lead $:</span>
            <input className="w-14 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={defaultCrewLead} title="Used when a lead has no 'lead' rate of their own" onChange={(e) => setDefaultCrewLead(e.target.value)} />
            <button type="button" className="btn text-xs" onClick={() => setDefaultCrewLead(String(DEFAULT_CREW_LEAD_RATE))}>Reset</button>
          </div>

          <SectionTitle sub="Daily overtime for NV: hours past the daily threshold pay 1.5× — but only for people at or below the base-rate threshold. Weekly 40-hour OT is unaffected.">
            NV daily overtime
          </SectionTitle>
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-xs">Daily OT after (hrs):</span>
            <input className="w-14 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={otHours} onChange={(e) => setOtHours(e.target.value)} />
            <span className="text-xs">Applies at base rate ≤ $:</span>
            <input className="w-14 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={nvOtRate} onChange={(e) => setNvOtRate(e.target.value)} />
            <button
              type="button"
              className="btn text-xs"
              onClick={() => {
                setOtHours(String(DEFAULT_DAILY_OT_HOURS_THRESHOLD));
                setNvOtRate(String(DEFAULT_NV_DAILY_OT_RATE_THRESHOLD));
              }}
            >
              Reset
            </button>
          </div>
          <div className="mt-2 flex flex-wrap items-start gap-2">
            <span className="mt-1 text-xs">Never gets daily OT:</span>
            <div className="flex flex-1 flex-col gap-1">
              {dailyOtExempt.map((n, i) => (
                <div key={i} className="flex items-center gap-1.5">
                  <input
                    className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                    value={n}
                    placeholder="Employee"
                    onChange={(ev) => setDailyOtExempt((prev) => prev.map((x, j) => (j === i ? ev.target.value : x)))}
                  />
                  <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setDailyOtExempt((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
                </div>
              ))}
              <div className="flex gap-2">
                <button type="button" className="btn text-xs" onClick={() => setDailyOtExempt((prev) => [...prev, ''])}>+ Add person</button>
                <button type="button" className="btn text-xs" onClick={() => setDailyOtExempt([...DEFAULT_DAILY_OT_EXEMPT])}>Reset</button>
              </div>
            </div>
          </div>

          <SectionTitle sub="Extra rows on a person's Payroll Preview block. Deductions come out of their total, reimbursements go in, and a note is just a labeled blank line. &quot;1st pay date&quot; rows appear only when the pay date falls on the 1st–7th of the month.">
            Payroll line items
          </SectionTitle>
          <div className="flex max-h-64 flex-col gap-1 overflow-y-auto rounded border border-neutral-100 p-1">
            {lineItems.map((r, i) => (
              <div key={r.id} className="flex flex-wrap items-center gap-1.5">
                <input
                  className="min-w-[8rem] flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                  value={r.person}
                  placeholder="Person"
                  onChange={(ev) => setLineItems((prev) => prev.map((x, j) => (j === i ? { ...x, person: ev.target.value } : x)))}
                />
                <input
                  className="min-w-[7rem] flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                  value={r.label}
                  placeholder="Label"
                  onChange={(ev) => setLineItems((prev) => prev.map((x, j) => (j === i ? { ...x, label: ev.target.value } : x)))}
                />
                <input
                  className="w-20 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs"
                  value={r.amount ?? ''}
                  placeholder="blank"
                  title="Leave blank for an empty, editable cell on the preview"
                  onChange={(ev) => {
                    const t = ev.target.value.trim();
                    const n = Number(t);
                    setLineItems((prev) =>
                      prev.map((x, j) => (j === i ? { ...x, amount: t === '' || !Number.isFinite(n) ? null : n } : x)),
                    );
                  }}
                />
                <select
                  className="rounded border border-neutral-300 px-1 py-1 text-xs"
                  value={r.kind}
                  onChange={(ev) => setLineItems((prev) => prev.map((x, j) => (j === i ? { ...x, kind: ev.target.value as PayrollLineItem['kind'] } : x)))}
                >
                  <option value="deduction">deduction</option>
                  <option value="reimbursement">reimbursement</option>
                  <option value="note">note</option>
                </select>
                <select
                  className="rounded border border-neutral-300 px-1 py-1 text-xs"
                  value={r.frequency}
                  onChange={(ev) => setLineItems((prev) => prev.map((x, j) => (j === i ? { ...x, frequency: ev.target.value as PayrollLineItem['frequency'] } : x)))}
                >
                  <option value="weekly">every week</option>
                  <option value="firstPayDateOfMonth">1st pay date</option>
                </select>
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setLineItems((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex gap-2">
            <button
              type="button"
              className="btn text-xs"
              onClick={() =>
                setLineItems((prev) => [
                  ...prev,
                  { id: `li-${Date.now()}`, person: '', label: '', amount: null, kind: 'note', frequency: 'weekly' },
                ])
              }
            >
              + Add line item
            </button>
            <button type="button" className="btn text-xs" onClick={() => setLineItems(DEFAULT_PAYROLL_LINE_ITEMS.map((e) => ({ ...e })))}>Reset</button>
          </div>

          <SectionTitle sub="These crew leads / subcontractors are charged for ALL of a job's materials instead of half — the band row reads &quot;- Materials&quot;.">
            Full-materials leads
          </SectionTitle>
          <div className="flex flex-col gap-1">
            {fullMatLeads.map((n, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input
                  className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                  value={n}
                  placeholder="Crew lead / subcontractor"
                  onChange={(ev) => setFullMatLeads((prev) => prev.map((x, j) => (j === i ? ev.target.value : x)))}
                />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setFullMatLeads((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex gap-2">
            <button type="button" className="btn text-xs" onClick={() => setFullMatLeads((prev) => [...prev, ''])}>+ Add lead</button>
            <button type="button" className="btn text-xs" onClick={() => setFullMatLeads([...DEFAULT_FULL_MATERIALS_LEADS])}>Reset</button>
          </div>

          <SectionTitle sub="A person's rate while working a job led by someone specific. Note: the crew lead for a job comes from the Pivot by Job Crew Lead column and from who clocked the &quot;Crew Lead - Hourly Wage&quot; pay type — use these rules only for genuinely negotiated rates.">
            Conditional rate rules
          </SectionTitle>
          <div className="flex flex-col gap-1">
            {condRules.map((r, i) => (
              <div key={r.id} className="flex flex-wrap items-center gap-1.5">
                <input
                  className="min-w-[9rem] flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                  value={r.person}
                  placeholder="Person"
                  onChange={(ev) => setCondRules((prev) => prev.map((x, j) => (j === i ? { ...x, person: ev.target.value } : x)))}
                />
                <span className="text-[11px] text-neutral-500">on jobs led by</span>
                <input
                  className="min-w-[9rem] flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                  value={r.condition.leadName}
                  placeholder="Crew lead"
                  onChange={(ev) =>
                    setCondRules((prev) =>
                      prev.map((x, j) => (j === i ? { ...x, condition: { ...x.condition, leadName: ev.target.value } } : x)),
                    )
                  }
                />
                <span className="text-[11px] text-neutral-500">$/hr</span>
                <input
                  className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs"
                  value={r.rate}
                  onChange={(ev) => {
                    const n = Number(ev.target.value);
                    if (Number.isFinite(n)) setCondRules((prev) => prev.map((x, j) => (j === i ? { ...x, rate: n } : x)));
                  }}
                />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setCondRules((prev) => prev.filter((_, j) => j !== i))} aria-label="Delete rule">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex gap-2">
            <button
              type="button"
              className="btn text-xs"
              onClick={() =>
                setCondRules((prev) => [
                  ...prev,
                  {
                    id: `cond-${Date.now()}-${prev.length}`,
                    person: '',
                    condition: { kind: 'onSameJobAsLead', leadName: '' },
                    rate: 0,
                  },
                ])
              }
            >
              + Add rule
            </button>
            <button type="button" className="btn text-xs" onClick={() => setCondRules(DEFAULT_CONDITIONAL_RATE_RULES.map((r) => ({ ...r, condition: { ...r.condition } })))}>Reset</button>
          </div>

          <SectionTitle sub="One person's $/hr on one job. The AP# is matched as a substring, so 24097 also covers its change-order suffixes.">
            Per-job rate overrides (per person)
          </SectionTitle>
          <div className="flex flex-col gap-1">
            {perJobRules.map((r, i) => (
              <div key={r.id} className="flex flex-wrap items-center gap-1.5">
                <input
                  className="min-w-[9rem] flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                  value={r.person}
                  placeholder="Person"
                  onChange={(ev) => setPerJobRules((prev) => prev.map((x, j) => (j === i ? { ...x, person: ev.target.value } : x)))}
                />
                <span className="text-[11px] text-neutral-500">AP#</span>
                <input
                  className="w-24 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                  value={r.apKeyPattern}
                  placeholder="24097"
                  onChange={(ev) => setPerJobRules((prev) => prev.map((x, j) => (j === i ? { ...x, apKeyPattern: ev.target.value } : x)))}
                />
                <span className="text-[11px] text-neutral-500">$/hr</span>
                <input
                  className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs"
                  value={r.rate}
                  onChange={(ev) => {
                    const n = Number(ev.target.value);
                    if (Number.isFinite(n)) setPerJobRules((prev) => prev.map((x, j) => (j === i ? { ...x, rate: n } : x)));
                  }}
                />
                <input
                  className="min-w-[7rem] flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                  value={r.note ?? ''}
                  placeholder="Optional note"
                  onChange={(ev) => setPerJobRules((prev) => prev.map((x, j) => (j === i ? { ...x, note: ev.target.value } : x)))}
                />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setPerJobRules((prev) => prev.filter((_, j) => j !== i))} aria-label="Delete override">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex gap-2">
            <button
              type="button"
              className="btn text-xs"
              onClick={() =>
                setPerJobRules((prev) => [
                  ...prev,
                  { id: `perjob-${Date.now()}-${prev.length}`, person: '', apKeyPattern: '', rate: 0 },
                ])
              }
            >
              + Add override
            </button>
            <button type="button" className="btn text-xs" onClick={() => setPerJobRules(DEFAULT_PER_JOB_RATE_OVERRIDES.map((r) => ({ ...r })))}>Reset</button>
          </div>

          <SectionTitle sub="Which P&L account rows feed each metric, per entity. Comma-separated labels. Sundries is intentionally NOT part of Materials.">
            P&amp;L account mapping
          </SectionTitle>
          <div className="flex flex-col gap-2">
            {PL_ENTITIES.map((entity) => (
              <div key={entity} className="rounded border border-neutral-200 p-2">
                <div className="mb-1 text-[11px] font-bold uppercase tracking-wide text-neutral-600">{entity}</div>
                <div className="flex flex-col gap-1">
                  {PL_METRICS.map((m) => (
                    <div key={m.key} className="flex items-center gap-1.5">
                      <span className="w-32 shrink-0 text-[11px] text-neutral-500">{m.label}</span>
                      <input
                        className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                        value={plMap[entity][m.key].join(', ')}
                        placeholder="Account label, Another label"
                        onChange={(ev) => setPlList(entity, m.key, ev.target.value)}
                      />
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
          <div className="mt-2 flex gap-2">
            <button type="button" className="btn text-xs" onClick={() => setPlMap(clonePlMapping(DEFAULT_PL_ACCOUNT_MAPPING))}>Reset</button>
          </div>

          <SectionTitle sub="Salary blocks by section on the Payroll Preview. wk = every week, mo = once a month on the month's first pay date (pay date on the 1st–7th). &quot;sub&quot; rows render at the very bottom of their section.">Salaried roster</SectionTitle>
          <div className="flex flex-col gap-1">
            {roster.map((e, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={e.name} placeholder="Name" onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, name: ev.target.value } : x)))} />
                <select className="rounded border border-neutral-300 px-1 py-1 text-xs" value={e.section} onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, section: ev.target.value as SalarySection } : x)))}>
                  <option>ADMIN</option>
                  <option>ESTIMATORS</option>
                  <option>MANAGER SALARY</option>
                </select>
                <select className="rounded border border-neutral-300 px-1 py-1 text-xs" value={e.entity} title="Entity" onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, entity: ev.target.value as EntityCode } : x)))}>
                  <option>NV</option>
                  <option>UT</option>
                </select>
                <input className="w-28 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={e.role} placeholder="Role text" onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, role: ev.target.value } : x)))} />
                <input className="w-20 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.amount} onChange={(ev) => { const n = Number(ev.target.value); if (Number.isFinite(n)) setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, amount: n } : x))); }} />
                {/* Round 9 §25 — monthly people show only on the month's first pay date. */}
                <select className="rounded border border-neutral-300 px-1 py-1 text-xs" value={e.frequency ?? 'weekly'} title="How often this person is paid" onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, frequency: ev.target.value as SalariedEntry['frequency'] } : x)))}>
                  <option value="weekly">wk</option>
                  <option value="firstPayDateOfMonth">mo</option>
                </select>
                {/* Round 9 §26 — subcontractor rows render at the bottom of the section. */}
                <label className="flex items-center gap-1 text-[10px] text-neutral-600" title="Subcontractor — renders at the very bottom of the section">
                  <input type="checkbox" checked={e.sub === true} onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, sub: ev.target.checked } : x)))} />
                  sub
                </label>
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setRoster((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex items-center gap-2">
            <button type="button" className="btn text-xs" onClick={() => setRoster((prev) => [...prev, { name: '', section: 'ESTIMATORS', role: '', amount: 0, entity: 'NV' }])}>+ Add person</button>
            <button type="button" className="btn text-xs" onClick={() => setRoster(DEFAULT_SALARIED.map((e) => ({ ...e })))}>Reset</button>
            <span className="ml-auto text-xs">Mileage reimb $/mile:</span>
            <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={mileageRate} onChange={(e) => setMileageRate(e.target.value)} />
            <span className="text-xs">Office mileage $/mile:</span>
            <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={officeMileage} onChange={(e) => setOfficeMileage(e.target.value)} />
          </div>

          <SectionTitle sub="Business names treated as valid crew leads without an employee-directory match.">Subcontractors</SectionTitle>
          <div className="flex flex-col gap-1">
            {subs.map((s, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={s} onChange={(ev) => setSubs((prev) => prev.map((x, j) => (j === i ? ev.target.value : x)))} />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setSubs((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex gap-2">
            <button type="button" className="btn text-xs" onClick={() => setSubs((prev) => [...prev, ''])}>+ Add subcontractor</button>
            <button type="button" className="btn text-xs" onClick={() => setSubs([...DEFAULT_SUBCONTRACTORS])}>Reset</button>
          </div>

          <SectionTitle sub="Jobs whose jobcode / event title contains one of these patterns are Davis-Bacon (DB) jobs. Hourly lines pay max(resolved rate, Davis-Bacon rate); crew-lead work-order pay is floored at hours × Davis-Bacon rate.">
            Davis-Bacon job patterns
          </SectionTitle>
          <div className="flex flex-col gap-1">
            {pwPatterns.map((s, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={s} placeholder="Substring pattern, e.g. SNRHA *SS*" onChange={(ev) => setPwPatterns((prev) => prev.map((x, j) => (j === i ? ev.target.value : x)))} />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setPwPatterns((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex items-center gap-2">
            <button type="button" className="btn text-xs" onClick={() => setPwPatterns((prev) => [...prev, ''])}>+ Add pattern</button>
            <button type="button" className="btn text-xs" onClick={() => { setPwPatterns([...DEFAULT_PREVAILING_PATTERNS]); setPwMin(String(DEFAULT_PREVAILING_MIN_RATE)); }}>Reset</button>
            <span className="ml-auto text-xs">Davis-Bacon rate $/hr:</span>
            <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={pwMin} onChange={(e) => setPwMin(e.target.value)} />
          </div>

          <SectionTitle sub="Negotiated $/hr for a specific AP# (optionally including change-order suffixes). Exact AP# + named employee beats all-painters; base-key matches are lower priority. Davis-Bacon still floors up when higher.">
            Job-specific pay rules
          </SectionTitle>
          <div className="flex flex-col gap-2">
            {jobRules.map((r, i) => (
              <div key={r.id} className="rounded border border-neutral-200 p-2">
                <div className="flex flex-wrap items-center gap-1.5">
                  <span className="text-[11px] text-neutral-500">AP#</span>
                  <input
                    className="w-24 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                    value={r.apKey}
                    placeholder="24314"
                    onChange={(ev) =>
                      setJobRules((prev) => prev.map((x, j) => (j === i ? { ...x, apKey: ev.target.value } : x)))
                    }
                  />
                  <span className="text-[11px] text-neutral-500">$/hr</span>
                  <input
                    className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs"
                    value={r.rate}
                    onChange={(ev) => {
                      const n = Number(ev.target.value);
                      if (Number.isFinite(n)) setJobRules((prev) => prev.map((x, j) => (j === i ? { ...x, rate: n } : x)));
                    }}
                  />
                  <label className="ml-1 flex items-center gap-1 text-[11px] text-neutral-600">
                    <input
                      type="checkbox"
                      checked={r.includeSuffixes}
                      onChange={(ev) =>
                        setJobRules((prev) =>
                          prev.map((x, j) => (j === i ? { ...x, includeSuffixes: ev.target.checked } : x)),
                        )
                      }
                    />
                    include suffixes
                  </label>
                  <button
                    type="button"
                    className="ml-auto text-neutral-400 hover:text-red-600"
                    onClick={() => setJobRules((prev) => prev.filter((_, j) => j !== i))}
                    aria-label="Delete rule"
                  >
                    🗑
                  </button>
                </div>
                <div className="mt-1.5 flex flex-wrap items-center gap-1.5">
                  <span className="text-[11px] text-neutral-500">Applies to</span>
                  <input
                    className="min-w-[12rem] flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                    value={r.employeeNames.length === 0 ? '' : r.employeeNames.join(', ')}
                    placeholder="All painters (or comma-separated names)"
                    onChange={(ev) => {
                      const raw = ev.target.value;
                      const names = raw
                        .split(',')
                        .map((s) => s.trim())
                        .filter(Boolean);
                      setJobRules((prev) => prev.map((x, j) => (j === i ? { ...x, employeeNames: names } : x)));
                    }}
                  />
                </div>
                <div className="mt-1.5 flex items-center gap-1.5">
                  <span className="text-[11px] text-neutral-500">Note</span>
                  <input
                    className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                    value={r.note}
                    placeholder="Optional note"
                    onChange={(ev) =>
                      setJobRules((prev) => prev.map((x, j) => (j === i ? { ...x, note: ev.target.value } : x)))
                    }
                  />
                </div>
              </div>
            ))}
          </div>
          <div className="mt-2 flex gap-2">
            <button
              type="button"
              className="btn text-xs"
              onClick={() =>
                setJobRules((prev) => [
                  ...prev,
                  {
                    id: `rule-${Date.now()}-${prev.length}`,
                    apKey: '',
                    rate: 0,
                    employeeNames: [],
                    includeSuffixes: true,
                    note: '',
                  },
                ])
              }
            >
              + Add rule
            </button>
            <button
              type="button"
              className="btn text-xs"
              onClick={() => {
                if (jobRules.length === 0) return;
                if (window.confirm('Clear all job-specific pay rules?')) setJobRules([]);
              }}
            >
              Clear all
            </button>
          </div>

          <SectionTitle sub="The employee list persists in this browser (localStorage) long-term — upload once until you need to replace it. Cross-device sync on the hosted URL is optional/advanced when Blob is configured.">
            Employee Directory
          </SectionTitle>
          {directorySyncWarning && (
            <div className="mb-2 rounded border border-amber-300 bg-amber-50 px-2 py-1.5 text-[11px] text-amber-900">
              {directorySyncWarning}
            </div>
          )}
          {directory ? (
            <div className="rounded border border-neutral-200 bg-neutral-50 p-2 text-xs">
              <div className="font-semibold">
                ✓ {directory.activeCount} active / {directory.inactiveCount} inactive
              </div>
              <div className="mt-0.5 text-neutral-500">
                {directory.sourceFilename} · updated{' '}
                {directory.uploadedAt
                  ? new Date(directory.uploadedAt).toLocaleString()
                  : '—'}
              </div>
            </div>
          ) : (
            <p className="text-xs text-neutral-500">
              No employee list saved yet — upload one here. The left rail no longer hosts this upload.
            </p>
          )}
          <div className="mt-2 flex items-center gap-2">
            <input
              ref={dirInputRef}
              type="file"
              accept=".xls,.xlsx,.csv"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) void replaceDirectory(file);
                e.target.value = '';
              }}
            />
            <button type="button" className="btn text-xs" disabled={dirBusy} onClick={() => dirInputRef.current?.click()}>
              {dirBusy ? 'Parsing…' : directory ? 'Replace list…' : 'Upload list…'}
            </button>
            {directory && (
              <button type="button" className="btn text-xs" onClick={onDirectoryClear}>
                Clear
              </button>
            )}
            {dirError && <span className="text-xs text-red-600">{dirError}</span>}
          </div>

          <div className="mt-6 border-t border-neutral-200 pt-3">
            <button
              type="button"
              className="flex w-full items-center gap-1.5 text-left text-xs font-bold uppercase tracking-wide text-neutral-600"
              onClick={() => setPaintersOpen((v) => !v)}
              aria-expanded={paintersOpen}
            >
              <span className="text-[10px] text-neutral-400">{paintersOpen ? '▼' : '▶'}</span>
              Painters — Active/Inactive
              <span className="ml-1 font-normal normal-case tracking-normal text-neutral-400">
                {activePainterCount} active / {painters.length - activePainterCount} inactive
              </span>
            </button>
            {paintersOpen && (
              <>
                <p className="mb-2 mt-1 text-[11px] text-neutral-500">
                  Inactive painters stay in the roster but are treated as not currently working. Notes hold crew /
                  *split* / {'{Reno}'} display metadata.
                </p>
                <div className="mb-2 flex items-center gap-2">
                  <input
                    className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                    value={painterQuery}
                    placeholder="Search painters…"
                    onChange={(e) => setPainterQuery(e.target.value)}
                  />
                  {painterQuery && (
                    <button type="button" className="btn text-xs" onClick={() => setPainterQuery('')}>
                      Clear
                    </button>
                  )}
                </div>
                <div className="flex max-h-64 flex-col gap-1 overflow-y-auto rounded border border-neutral-100 p-1">
                  {painterMatches.length === 0 && (
                    <p className="px-1 py-2 text-[11px] text-neutral-400">No painters match “{painterQuery}”.</p>
                  )}
                  {painterMatches.map(({ p, i }) => (
                    <div key={`${p.name}-${i}`} className="flex items-center gap-1.5">
                      <button
                        type="button"
                        className={`w-20 shrink-0 rounded border px-1.5 py-1 text-[11px] font-semibold ${
                          p.active
                            ? 'border-emerald-300 bg-emerald-50 text-emerald-700'
                            : 'border-neutral-300 bg-neutral-50 text-neutral-500'
                        }`}
                        onClick={() =>
                          setPainters((prev) => prev.map((x, j) => (j === i ? { ...x, active: !x.active } : x)))
                        }
                        title="Toggle Active/Inactive"
                      >
                        {p.active ? 'Active' : 'Inactive'}
                      </button>
                      <input
                        className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                        value={p.name}
                        placeholder="Painter name"
                        onChange={(ev) =>
                          setPainters((prev) => prev.map((x, j) => (j === i ? { ...x, name: ev.target.value } : x)))
                        }
                      />
                      <input
                        className="w-28 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                        value={p.note ?? ''}
                        placeholder="note"
                        onChange={(ev) =>
                          setPainters((prev) => prev.map((x, j) => (j === i ? { ...x, note: ev.target.value } : x)))
                        }
                      />
                      <button
                        type="button"
                        className="text-neutral-400 hover:text-red-600"
                        onClick={() => setPainters((prev) => prev.filter((_, j) => j !== i))}
                        aria-label="Remove painter"
                      >
                        🗑
                      </button>
                    </div>
                  ))}
                </div>
                <div className="mt-2 flex flex-wrap items-center gap-2">
                  <input
                    className="w-48 rounded border border-neutral-300 px-1.5 py-1 text-xs"
                    value={newPainter}
                    placeholder="New painter name"
                    onChange={(e) => setNewPainter(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key !== 'Enter') return;
                      const name = newPainter.trim();
                      if (!name) return;
                      setPainters((prev) => [...prev, { name, active: true }]);
                      setNewPainter('');
                    }}
                  />
                  <button
                    type="button"
                    className="btn text-xs"
                    onClick={() => {
                      const name = newPainter.trim();
                      if (!name) return;
                      setPainters((prev) => [...prev, { name, active: true }]);
                      setNewPainter('');
                    }}
                  >
                    + Add painter
                  </button>
                  <button
                    type="button"
                    className="btn text-xs"
                    onClick={() => setPainters(DEFAULT_PAINTER_ROSTER.map((p) => ({ ...p })))}
                  >
                    Reset
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
        <footer className="flex justify-end gap-2 border-t border-neutral-200 px-4 py-3">
          <button type="button" className="btn" onClick={onClose}>Cancel</button>
          <button
            type="button"
            className="btn-primary"
            disabled={!legendValid || !pivotValid || !numbersValid}
            title={legendValid && pivotValid && numbersValid ? undefined : 'Check the legend colors, pivot columns and numeric fields'}
            onClick={save}
          >
            Save
          </button>
        </footer>
      </div>
    </div>
  );
}
