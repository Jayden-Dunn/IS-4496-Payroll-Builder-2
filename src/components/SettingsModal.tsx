import { useRef, useState } from 'react';
import {
  DEFAULT_LEGEND,
  DEFAULT_PIVOT_COLUMNS,
  DEFAULT_PREVAILING_MIN_RATE,
  DEFAULT_PREVAILING_PATTERNS,
  DEFAULT_PROFIT_SHARE,
  DEFAULT_SALARIED,
  DEFAULT_SUBCONTRACTORS,
  DEFAULT_WAGE_RATES,
} from '../lib/settings';
import type {
  AppSettings,
  LegendEntry,
  ProfitShareEntry,
  SalariedEntry,
  SalarySection,
  SavedDirectory,
  WageRateEntry,
} from '../lib/types';

interface SettingsModalProps {
  settings: AppSettings;
  /** Round 4 §4.2 — the persisted employee directory (null = never saved). */
  directory: SavedDirectory | null;
  /** Parse + persist a replacement list; resolves to an error message or null. */
  onDirectoryReplace: (file: File) => Promise<string | null>;
  onDirectoryClear: () => void;
  onSave: (settings: AppSettings) => void;
  onClose: () => void;
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
  const [defaultWage, setDefaultWage] = useState(String(settings.defaultWageRate));
  const [roster, setRoster] = useState<SalariedEntry[]>(settings.salariedRoster.map((e) => ({ ...e })));
  const [mileageRate, setMileageRate] = useState(String(settings.mileageRate));
  const [subs, setSubs] = useState<string[]>([...settings.subcontractors]);
  const [pwPatterns, setPwPatterns] = useState<string[]>([...settings.prevailingWagePatterns]);
  const [pwMin, setPwMin] = useState(String(settings.prevailingMinRate));
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
  const mileageNum = Number(mileageRate);
  const pwMinNum = Number(pwMin);
  const numbersValid =
    Number.isFinite(defaultPsNum) &&
    Number.isFinite(defaultWageNum) &&
    Number.isFinite(mileageNum) &&
    Number.isFinite(pwMinNum);

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
      salariedRoster: roster.filter((e) => e.name.trim() !== ''),
      mileageRate: mileageNum,
      subcontractors: subs.map((s) => s.trim()).filter(Boolean),
      prevailingWagePatterns: pwPatterns.map((s) => s.trim()).filter(Boolean),
      prevailingMinRate: pwMinNum,
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

          <SectionTitle sub="Default $/hr per employee; the optional second rate applies on hourly (T&M) jobs. OT is always 1.5× the line rate.">
            Wage rates
          </SectionTitle>
          <div className="flex max-h-64 flex-col gap-1 overflow-y-auto rounded border border-neutral-100 p-1">
            {wages.map((e, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={e.name} placeholder="Employee" onChange={(ev) => setWages((prev) => prev.map((x, j) => (j === i ? { ...x, name: ev.target.value } : x)))} />
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.rate} onChange={(ev) => { const n = Number(ev.target.value); if (Number.isFinite(n)) setWages((prev) => prev.map((x, j) => (j === i ? { ...x, rate: n } : x))); }} />
                <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.hourlyRate ?? ''} placeholder="hrly" onChange={(ev) => { const t = ev.target.value.trim(); const n = Number(t); setWages((prev) => prev.map((x, j) => (j === i ? { ...x, ...(t === '' || !Number.isFinite(n) ? { hourlyRate: undefined } : { hourlyRate: n }) } : x))); }} />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setWages((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex items-center gap-2">
            <button type="button" className="btn text-xs" onClick={() => setWages((prev) => [...prev, { name: '', rate: 18 }])}>+ Add employee</button>
            <button type="button" className="btn text-xs" onClick={() => setWages(DEFAULT_WAGE_RATES.map((e) => ({ ...e })))}>Reset</button>
            <span className="ml-auto text-xs">Fallback rate $:</span>
            <input className="w-14 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={defaultWage} onChange={(e) => setDefaultWage(e.target.value)} />
          </div>

          <SectionTitle sub="Weekly salary blocks by section on the Payroll Preview.">Salaried roster</SectionTitle>
          <div className="flex flex-col gap-1">
            {roster.map((e, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input className="flex-1 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={e.name} placeholder="Name" onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, name: ev.target.value } : x)))} />
                <select className="rounded border border-neutral-300 px-1 py-1 text-xs" value={e.section} onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, section: ev.target.value as SalarySection } : x)))}>
                  <option>ADMIN</option>
                  <option>ESTIMATORS</option>
                  <option>MANAGER SALARY</option>
                </select>
                <input className="w-32 rounded border border-neutral-300 px-1.5 py-1 text-xs" value={e.role} placeholder="Role text" onChange={(ev) => setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, role: ev.target.value } : x)))} />
                <input className="w-20 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={e.amount} onChange={(ev) => { const n = Number(ev.target.value); if (Number.isFinite(n)) setRoster((prev) => prev.map((x, j) => (j === i ? { ...x, amount: n } : x))); }} />
                <button type="button" className="text-neutral-400 hover:text-red-600" onClick={() => setRoster((prev) => prev.filter((_, j) => j !== i))} aria-label="Remove">🗑</button>
              </div>
            ))}
          </div>
          <div className="mt-2 flex items-center gap-2">
            <button type="button" className="btn text-xs" onClick={() => setRoster((prev) => [...prev, { name: '', section: 'ESTIMATORS', role: '', amount: 0 }])}>+ Add person</button>
            <button type="button" className="btn text-xs" onClick={() => setRoster(DEFAULT_SALARIED.map((e) => ({ ...e })))}>Reset</button>
            <span className="ml-auto text-xs">Mileage reimb $/mile:</span>
            <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={mileageRate} onChange={(e) => setMileageRate(e.target.value)} />
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

          <SectionTitle sub="Round 4 §5.8 — jobs whose jobcode / event title contains one of these patterns pay max(default rate, prevailing minimum).">
            Prevailing-wage (Davis-Bacon) jobs
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
            <span className="ml-auto text-xs">Prevailing minimum $/hr:</span>
            <input className="w-16 rounded border border-neutral-300 px-1.5 py-1 text-right text-xs" value={pwMin} onChange={(e) => setPwMin(e.target.value)} />
          </div>

          <SectionTitle sub="The current Employee Name List is saved into the app (shared by every device on this server) and kept in this browser. Replacing it updates the saved date and re-runs crew-lead resolution.">
            Employee Directory
          </SectionTitle>
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
