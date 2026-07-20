import { useState, type CSSProperties } from 'react';
import { fmtAccounting, isDarkColor, parseMoneyInput, parsePercentInput } from '../lib/format';
import type { HourlyBlock, PayrollEdit, PayrollPreviewModel, PsBlock, SalaryBlock } from '../lib/types';

interface PayrollPreviewProps {
  model: PayrollPreviewModel | null; // null → P&L not uploaded
  mileageRate: number;
  onEdit: (key: string, patch: PayrollEdit) => void;
}

const BLACK: CSSProperties = { backgroundColor: '#000', color: '#fff' };
const TAN: CSSProperties = { backgroundColor: '#D6D4CA', fontWeight: 700 };
const NEARWHITE: CSSProperties = { backgroundColor: '#F9F9F9', fontWeight: 700 };
const MANAGER: CSSProperties = { backgroundColor: '#D99694', fontWeight: 700 };

function NumCell({
  value,
  onCommit,
  pct,
  className,
}: {
  value: number | null;
  onCommit: (n: number | null) => void;
  pct?: boolean;
  className?: string;
}): JSX.Element {
  const [text, setText] = useState<string | null>(null);
  const display = value === null ? '' : pct ? `${Math.round(value * 100)}%` : value.toFixed(2);
  return (
    <input
      className={`cell-input text-right ${className ?? ''}`}
      value={text ?? display}
      onChange={(e) => setText(e.target.value)}
      onFocus={() => setText(display)}
      onBlur={() => {
        if (text !== null && text !== display) {
          if (pct) {
            const p = parsePercentInput(text);
            onCommit(typeof p === 'number' ? p : null);
          } else {
            onCommit(parseMoneyInput(text));
          }
        }
        setText(null);
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
        if (e.key === 'Escape') setText(null);
      }}
    />
  );
}

function jobCellStyle(hex: string | null): CSSProperties | undefined {
  if (!hex) return undefined;
  return { backgroundColor: `#${hex}`, color: isDarkColor(hex) ? '#fff' : '#000' };
}

function HourlyBlockView({
  block,
  onEdit,
  headerStyle = TAN,
}: {
  block: HourlyBlock;
  onEdit: PayrollPreviewProps['onEdit'];
  headerStyle?: CSSProperties;
}): JSX.Element {
  const taxes = block.net !== null ? block.totalPay - block.net : null;
  return (
    <table className="mb-3 w-full max-w-4xl border-collapse bg-white text-xs" id={`pp-${block.key}`}>
      <thead>
        <tr>
          <th className="xcell text-left" style={headerStyle} colSpan={7}>
            {block.title}
            {block.inactive && <span className="ml-2 rounded bg-red-100 px-1 text-[10px] text-red-700">inactive (*)</span>}
          </th>
        </tr>
        <tr className="text-neutral-500">
          <th className="xcell w-24 text-left font-normal">PO#</th>
          <th className="xcell text-left font-normal">Job</th>
          <th className="xcell w-16 text-right font-normal">hrs</th>
          <th className="xcell w-16 text-right font-normal">rate</th>
          <th className="xcell w-24 text-right font-normal">total</th>
          <th className="xcell w-24 text-right font-normal">(-) taxes</th>
          <th className="xcell w-24 text-right font-normal">net</th>
        </tr>
      </thead>
      <tbody>
        {block.lines.map((l) => (
          <tr key={l.id}>
            <td className="xcell" style={jobCellStyle(l.colorHex)}>{l.apKey ? `AP${l.apKey}` : ''}</td>
            <td className="xcell" style={jobCellStyle(l.colorHex)}>{l.jobLabel}</td>
            <td className={`xcell text-right ${l.ot ? 'bg-yellow-100' : ''}`}>
              <NumCell value={l.hours} onCommit={(n) => onEdit(`hl:${l.id}`, { hours: n ?? 0 })} />
            </td>
            <td className={`xcell text-right ${l.rateDefaulted ? 'bg-amber-50' : ''}`} title={l.rateDefaulted ? 'No configured wage rate — default applied' : undefined}>
              <NumCell value={l.rate} onCommit={(n) => onEdit(`hl:${l.id}`, { rate: n ?? 0 })} />
            </td>
            <td className="xcell text-right">{fmtAccounting(l.hours * l.rate)}</td>
            <td className="xcell" />
            <td className="xcell" />
          </tr>
        ))}
        <tr className="font-semibold">
          <td className="xcell" />
          <td className="xcell" />
          <td className="xcell text-right">{block.totalHours.toFixed(2)}</td>
          <td className="xcell" />
          <td className="xcell text-right">{fmtAccounting(block.totalPay)}</td>
          <td className="xcell text-right">{taxes !== null ? fmtAccounting(taxes) : ''}</td>
          <td className="xcell text-right">
            <NumCell value={block.net} onCommit={(n) => onEdit(`hb:${block.key}`, { net: n })} />
          </td>
        </tr>
      </tbody>
    </table>
  );
}

function PsBlockView({ block, onEdit }: { block: PsBlock; onEdit: PayrollPreviewProps['onEdit'] }): JSX.Element {
  const label = (t: string): JSX.Element => <td className="xcell w-44 font-medium">{t}</td>;
  const psKey = (jobRowId: string): string => `ps:${block.key}:${jobRowId}`;
  const anySplit = block.jobs.some((j) => j.splitWith > 1);
  return (
    <table className="mb-4 w-max min-w-[560px] border-collapse bg-white text-xs" id={`pp-${block.key}`}>
      <thead>
        <tr>
          <th className="xcell text-left" style={TAN}>
            {block.leadName}
            {block.sub && <span className="ml-2 rounded bg-blue-100 px-1 text-[10px] font-normal text-blue-700">sub</span>}
          </th>
          {block.jobs.map((j) => (
            <th key={j.jobRowId} className="xcell min-w-[130px] text-right" style={TAN}>
              {j.woPct !== null ? `${Math.round(j.woPct * 100)}%` : '-'}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr>
          {label('PO')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className={`xcell text-right ${j.missingFromPl ? 'text-red-700' : ''}`} title={j.missingFromPl ? 'No P&L column for this job — values default to 0' : undefined}>
              {j.jobShort}
            </td>
          ))}
        </tr>
        <tr>
          {label('Name')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">{j.customer}</td>
          ))}
        </tr>
        <tr>
          {label('Profit Share')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">{fmtAccounting(j.profitShare)}</td>
          ))}
        </tr>
        <tr>
          {label('Work Order Override')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              <NumCell value={j.workOrderOverride} onCommit={(n) => onEdit(psKey(j.jobRowId), { workOrderOverride: n ?? 0 })} />
            </td>
          ))}
        </tr>
        <tr>
          {label('% Completion')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              <NumCell pct value={j.percentCompletion} onCommit={(n) => onEdit(psKey(j.jobRowId), { percentCompletion: n ?? 0 })} />
            </td>
          ))}
        </tr>
        <tr>
          {label('- Labor + Wages')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              <NumCell value={j.laborWages} onCommit={(n) => onEdit(psKey(j.jobRowId), { laborWages: n ?? 0 })} />
            </td>
          ))}
        </tr>
        <tr>
          {label('-Mileage')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              <NumCell value={j.mileage} onCommit={(n) => onEdit(psKey(j.jobRowId), { mileage: n ?? 0 })} />
            </td>
          ))}
        </tr>
        <tr>
          {label('-½ Materials')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              <NumCell value={j.halfMaterials} onCommit={(n) => onEdit(psKey(j.jobRowId), { halfMaterials: n ?? 0 })} />
            </td>
          ))}
        </tr>
        <tr className="font-semibold">
          {label('Total Profit Share')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className={`xcell text-right ${j.totalProfitShare < 25 ? 'bg-amber-100' : ''}`}>
              {fmtAccounting(j.totalProfitShare)}
            </td>
          ))}
        </tr>
        {anySplit && (
          <tr>
            {label('split w/')}
            {block.jobs.map((j) => (
              <td key={j.jobRowId} className="xcell text-right">
                {j.splitWith > 1 ? `÷${j.splitWith} = ${fmtAccounting(j.totalAfterSplit)}` : ''}
              </td>
            ))}
          </tr>
        )}
        <SummaryRow label="Total" value={fmtAccounting(block.total)} />
        <tr>
          {label('Mileage - New')}
          <td className="xcell text-right" colSpan={block.jobs.length}>
            <NumCell value={block.mileageNew} onCommit={(n) => onEdit(`psb:${block.key}`, { mileageNew: n ?? 0 })} />
          </td>
        </tr>
        <SummaryRow label="Hourly Pay" value={fmtAccounting(block.hourlyPay)} span={block.jobs.length} />
        <tr>
          {label('Reimbursements')}
          <td className="xcell text-right">
            <NumCell value={block.reimbursements} onCommit={(n) => onEdit(`psb:${block.key}`, { reimbursements: n ?? 0 })} />
          </td>
          {block.jobs.length > 1 && (
            <td className="xcell" colSpan={block.jobs.length - 1}>
              <MemoCell value={block.reimbMemo} onCommit={(t) => onEdit(`psb:${block.key}`, { reimbMemo: t })} />
            </td>
          )}
        </tr>
        <tr>
          {label('Deductions/Sundries')}
          <td className="xcell text-right">
            <NumCell value={block.deductions} onCommit={(n) => onEdit(`psb:${block.key}`, { deductions: n ?? 0 })} />
          </td>
          {block.jobs.length > 1 && (
            <td className="xcell" colSpan={block.jobs.length - 1}>
              <MemoCell value={block.deductMemo} onCommit={(t) => onEdit(`psb:${block.key}`, { deductMemo: t })} />
            </td>
          )}
        </tr>
        <SummaryRow label="Total Pay" value={fmtAccounting(block.totalPay)} span={block.jobs.length} bold />
        <tr>
          <td className="xcell" colSpan={1 + block.jobs.length} />
        </tr>
        <tr>
          {label('Total Income')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              <NumCell value={j.income} onCommit={(n) => onEdit(psKey(j.jobRowId), { income: n ?? 0 })} />
            </td>
          ))}
        </tr>
        <tr>
          {label('Materials')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              <NumCell value={j.materials} onCommit={(n) => onEdit(psKey(j.jobRowId), { materials: n ?? 0 })} />
            </td>
          ))}
        </tr>
        <tr>
          {label('QB Labor-Wages-Subs')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">{fmtAccounting(j.laborWages)}</td>
          ))}
        </tr>
        <tr>
          {label('Mileage in Books')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">{fmtAccounting(j.mileage)}</td>
          ))}
        </tr>
        <tr>
          {label('% Materials')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              {j.income !== 0 ? `${Math.round((j.materials / j.income) * 100)}%` : '—'}
            </td>
          ))}
        </tr>
      </tbody>
    </table>
  );
}

function SummaryRow({ label, value, span = 1, bold }: { label: string; value: string; span?: number; bold?: boolean }): JSX.Element {
  return (
    <tr className={bold ? 'font-bold' : 'font-semibold'}>
      <td className="xcell w-44">{label}</td>
      <td className="xcell text-right" colSpan={span}>
        {value}
      </td>
    </tr>
  );
}

function MemoCell({ value, onCommit }: { value: string; onCommit: (t: string) => void }): JSX.Element {
  const [text, setText] = useState<string | null>(null);
  return (
    <input
      className="cell-input italic text-neutral-600"
      placeholder="memo…"
      value={text ?? value}
      onChange={(e) => setText(e.target.value)}
      onBlur={() => {
        if (text !== null && text !== value) onCommit(text);
        setText(null);
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
      }}
    />
  );
}

function SalaryBlockView({
  block,
  mileageRate,
  onEdit,
  headerStyle,
}: {
  block: SalaryBlock;
  mileageRate: number;
  onEdit: PayrollPreviewProps['onEdit'];
  headerStyle: CSSProperties;
}): JSX.Element {
  return (
    <table className="mb-3 w-full max-w-2xl border-collapse bg-white text-xs" id={`pp-${block.key}`}>
      <thead>
        <tr>
          <th className="xcell text-left" style={headerStyle} colSpan={3}>
            {block.name}
            {block.timesheetHours > 0 && <span className="ml-1 font-normal">({block.timesheetHours}hrs)</span>}
          </th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td className="xcell w-44">{block.role}</td>
          <td className="xcell" />
          <td className="xcell w-28 text-right">{fmtAccounting(block.amount)}</td>
        </tr>
        <tr>
          <td className="xcell">mileage reimb ({mileageRate.toFixed(3)}/mi)</td>
          <td className="xcell text-right">
            <NumCell value={block.miles} onCommit={(n) => onEdit(block.key, { miles: n ?? 0 })} />
          </td>
          <td className="xcell text-right">{fmtAccounting(block.mileageReimb)}</td>
        </tr>
        <tr className="font-semibold">
          <td className="xcell">Total</td>
          <td className="xcell" />
          <td className="xcell text-right">{fmtAccounting(block.amount + block.mileageReimb)}</td>
        </tr>
      </tbody>
    </table>
  );
}

function SectionHeader({ title }: { title: string }): JSX.Element {
  return (
    <div className="mb-2 mt-4 max-w-4xl px-2 py-1 text-sm font-bold" style={BLACK}>
      {title}
    </div>
  );
}

export function PayrollPreview({ model, mileageRate, onEdit }: PayrollPreviewProps): JSX.Element {
  const [showInactive, setShowInactive] = useState(false);
  if (!model) {
    return (
      <div className="flex h-full items-center justify-center">
        <div className="max-w-md rounded-lg border border-neutral-300 bg-white p-6 text-center shadow-sm">
          <div className="mb-2 text-4xl">🧾</div>
          <h2 className="mb-1 text-base font-bold">Payroll Preview needs the P&amp;L</h2>
          <p className="text-xs leading-relaxed text-neutral-500">
            Upload slot 8 — the QuickBooks <b>Profit and Loss by Customer</b> export — to build the payroll blocks.
            (Slot 9, the Employee Name List, is recommended too for crew-lead name resolution.)
          </p>
        </div>
      </div>
    );
  }
  return (
    <div className="p-3">
      <div className="mb-3 max-w-4xl px-2 py-1.5 text-sm font-bold" style={BLACK}>
        {model.title}
      </div>

      {model.leads.map((b) => (
        <div key={b.key}>
          <PsBlockView block={b} onEdit={onEdit} />
          {b.tsheets && (
            <div className="mb-4 ml-4">
              <div className="mb-1 max-w-4xl px-2 py-1 text-xs" style={NEARWHITE}>
                {b.tsheets.employeeName.split(' ')[0]}&apos;s T-SHEETS BREAKDOWN
              </div>
              <HourlyBlockView block={b.tsheets} onEdit={onEdit} headerStyle={NEARWHITE} />
            </div>
          )}
        </div>
      ))}

      {model.painters.map((b) => (
        <HourlyBlockView key={b.key} block={b} onEdit={onEdit} />
      ))}

      <SectionHeader title="OFFICE" />
      {model.office.map((b) => (
        <HourlyBlockView key={b.key} block={b} onEdit={onEdit} />
      ))}

      <SectionHeader title="ADMIN" />
      {model.admin.map((b) => (
        <SalaryBlockView key={b.key} block={b} mileageRate={mileageRate} onEdit={onEdit} headerStyle={TAN} />
      ))}

      <SectionHeader title="ESTIMATORS" />
      {model.estimators.map((b) => (
        <SalaryBlockView key={b.key} block={b} mileageRate={mileageRate} onEdit={onEdit} headerStyle={TAN} />
      ))}

      <SectionHeader title="MANAGER SALARY" />
      {model.managers.map((b) => (
        <SalaryBlockView key={b.key} block={b} mileageRate={mileageRate} onEdit={onEdit} headerStyle={MANAGER} />
      ))}

      <div className="mt-6 max-w-4xl rounded border border-neutral-300 bg-white">
        <button
          type="button"
          className="flex w-full items-center justify-between px-3 py-2 text-left text-sm font-bold"
          onClick={() => setShowInactive((v) => !v)}
        >
          Inactive Employees ({model.inactiveZeroHours.length})
          <span>{showInactive ? '▾' : '▸'}</span>
        </button>
        {showInactive && (
          <ul className="divide-y divide-neutral-100 border-t border-neutral-200">
            {model.inactiveZeroHours.length === 0 && <li className="px-3 py-1.5 text-xs text-neutral-400">None.</li>}
            {model.inactiveZeroHours.map((n) => (
              <li key={n} className="px-3 py-1.5 text-xs">
                {n}
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}
