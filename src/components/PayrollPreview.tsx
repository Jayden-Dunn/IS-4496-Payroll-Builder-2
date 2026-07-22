import { useState, type CSSProperties } from 'react';
import { fmtAccounting, isDarkColor, parseMoneyInput, parsePercentInput } from '../lib/format';
import type { HourlyBlock, PayrollEdit, PayrollPreviewModel, PsBlock, SalaryBlock } from '../lib/types';

interface PayrollPreviewProps {
  model: PayrollPreviewModel | null; // null → P&L not uploaded
  mileageRate: number;
  onEdit: (key: string, patch: PayrollEdit) => void;
}

// Round 4 §5.1 — the whole tab renders in Garamond.
const GARAMOND: CSSProperties = { fontFamily: 'Garamond, "EB Garamond", "Times New Roman", serif' };

const BLACK: CSSProperties = { backgroundColor: '#000', color: '#fff' };
// Round 4 supersession — name bands are DDD9C4.
const TAN: CSSProperties = { backgroundColor: '#DDD9C4', fontWeight: 700 };
const NEARWHITE: CSSProperties = { backgroundColor: '#F9F9F9', fontWeight: 700 };
const MANAGER: CSSProperties = { backgroundColor: '#D99694', fontWeight: 700 };
const OFFICE_BLUE: CSSProperties = { backgroundColor: '#4F81BD', color: '#fff' };
const HOURLY_PAY_YELLOW: CSSProperties = { backgroundColor: '#FFFF99' };

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

/**
 * §5.5/§5.7 — one hourly block with the labeled `PO# | Job | hrs | rate |
 * total` columns, `(-) taxes`/`net` labels on the last job line, and (in the
 * T-SHEETS variant) the boxed near-white table with dotted separators.
 */
function HourlyBlockView({
  block,
  onEdit,
  tsheets = false,
}: {
  block: HourlyBlock;
  onEdit: PayrollPreviewProps['onEdit'];
  tsheets?: boolean;
}): JSX.Element {
  // Round 6 §6.1 — taxes / nest / net are independent manual cells (never total − net).
  const bodyCls = tsheets ? 'bg-[#F9F9F9]' : '';
  const dotted = tsheets ? 'border-b border-dotted border-neutral-400' : '';
  return (
    <table
      className={`mb-3 w-full max-w-4xl border-collapse bg-white text-xs ${tsheets ? 'border border-neutral-500' : ''}`}
      id={`pp-${tsheets ? 'ts-' : ''}${block.key}`}
    >
      <thead>
        {tsheets ? (
          <tr>
            <th className="xcell border border-neutral-500 text-center" style={NEARWHITE} colSpan={5}>
              {block.employeeName.split(' ')[0]}&apos;s T-SHEETS BREAKDOWN
            </th>
            <th className="xcell" colSpan={2} />
          </tr>
        ) : (
          <tr>
            {/* §5.4 — merged tan name band: name bold black + crew tag bold white. */}
            <th className="xcell py-1 text-center" style={TAN} colSpan={7}>
              <span className="text-[13px] font-bold text-black">{block.employeeName}</span>
              {block.crewTag && <span className="ml-1 text-[11px] font-bold text-white">({block.crewTag})</span>}
              {block.inactive && <span className="ml-2 rounded bg-red-100 px-1 text-[10px] text-red-700">inactive (*)</span>}
            </th>
          </tr>
        )}
        <tr>
          <th className={`xcell w-24 text-center font-bold underline ${bodyCls}`}>PO#</th>
          <th className={`xcell text-center font-bold underline ${bodyCls}`}>Job</th>
          <th className={`xcell w-16 text-center font-bold underline ${bodyCls}`}>hrs</th>
          <th className={`xcell w-16 text-center font-bold underline ${bodyCls}`}>rate</th>
          <th className={`xcell w-24 text-center font-bold underline ${bodyCls}`}>total</th>
          <th className="xcell w-24" />
          <th className="xcell w-24" />
        </tr>
      </thead>
      <tbody>
        {block.lines.map((l, i) => {
          const isLast = i === block.lines.length - 1;
          return (
            <tr key={l.id}>
              <td
                className={`xcell ${l.poText !== null ? 'text-center text-[10px]' : ''} ${dotted} ${bodyCls}`}
                style={l.poText !== null ? OFFICE_BLUE : jobCellStyle(l.colorHex)}
              >
                {l.poText ?? (l.apKey ? `AP${l.apKey}` : '')}
              </td>
              <td className={`xcell text-[11px] ${dotted} ${bodyCls}`}>{l.jobLabel}</td>
              <td className={`xcell text-center ${l.ot ? 'bg-yellow-100' : bodyCls} ${dotted}`}>
                <NumCell value={l.hours} onCommit={(n) => onEdit(`hl:${l.id}`, { hours: n ?? 0 })} className="text-center" />
              </td>
              <td
                className={`xcell text-center ${l.rateDefaulted ? 'bg-amber-50' : bodyCls} ${l.belowPrevailing ? 'bg-red-100' : ''} ${l.fromJobRule ? 'ring-1 ring-inset ring-sky-300' : ''} ${dotted}`}
                title={
                  l.belowPrevailing
                    ? 'Davis-Bacon job paid below the Davis-Bacon rate'
                    : l.fromJobRule
                      ? l.ruleNote
                        ? `Job-rate rule: ${l.ruleNote}`
                        : 'Job-rate rule applied'
                      : l.prevailing
                        ? 'Davis-Bacon job — rate = max(resolved, Davis-Bacon rate)'
                        : l.rateDefaulted
                          ? 'No configured wage rate — default applied'
                          : undefined
                }
              >
                <NumCell value={l.rate} onCommit={(n) => onEdit(`hl:${l.id}`, { rate: n ?? 0 })} className="text-center" />
              </td>
              <td className={`xcell text-center ${dotted} ${bodyCls}`}>{fmtAccounting(l.hours * l.rate)}</td>
              <td className="xcell text-center font-bold">{isLast ? '(-) taxes' : ''}</td>
              <td className="xcell text-center font-bold">{isLast ? 'NEST [IRA]' : ''}</td>
              <td className="xcell text-center font-bold">{isLast ? 'net' : ''}</td>
            </tr>
          );
        })}
        <tr className="font-semibold">
          <td className={`xcell border-t border-t-neutral-500 ${bodyCls}`} />
          <td className={`xcell border-t border-t-neutral-500 ${bodyCls}`} />
          <td className={`xcell border-t border-t-neutral-500 text-center ${bodyCls}`}>{block.totalHours.toFixed(2)}</td>
          <td className={`xcell border-t border-t-neutral-500 ${bodyCls}`} />
          <td className={`xcell border-t border-t-neutral-500 text-center ${bodyCls}`}>{fmtAccounting(block.totalPay)}</td>
          <td className="xcell text-center">
            <NumCell value={block.taxes} onCommit={(n) => onEdit(`hb:${block.key}`, { taxes: n })} className="text-center" />
          </td>
          <td className="xcell text-center">
            <NumCell value={block.nest} onCommit={(n) => onEdit(`hb:${block.key}`, { nest: n })} className="text-center" />
          </td>
          <td className="xcell text-center">
            <NumCell value={block.net} onCommit={(n) => onEdit(`hb:${block.key}`, { net: n })} className="text-center" />
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
  // §5.6 — PO/Name value cells: right-aligned, bold, estimator fill.
  const poNameCell = (j: PsBlock['jobs'][number], text: string, title?: string): JSX.Element => (
    <td
      key={j.jobRowId}
      className={`xcell text-right text-[11px] font-bold ${j.missingFromPl ? 'text-red-700' : ''}`}
      style={jobCellStyle(j.colorHex)}
      title={title}
    >
      {text}
    </td>
  );
  return (
    <table className="mb-4 w-max min-w-[560px] border-collapse bg-white text-xs" id={`pp-${block.key}`}>
      <thead>
        <tr>
          {/* Round 5 §5 — band fills A through H regardless of job count. */}
          <th className="xcell text-left" style={TAN}>
            {block.displayName}
            {block.sub && <span className="ml-2 rounded bg-blue-100 px-1 text-[10px] font-normal text-blue-700">sub</span>}
          </th>
          {block.jobs.map((j) => (
            <th key={j.jobRowId} className="xcell min-w-[130px] text-right" style={TAN}>
              {j.woPct !== null ? `${Math.round(j.woPct * 100)}%` : '-'}
            </th>
          ))}
          {Array.from({ length: Math.max(0, 7 - block.jobs.length) }, (_, i) => (
            <th key={`pad-${i}`} className="xcell min-w-[40px]" style={TAN} />
          ))}
        </tr>
      </thead>
      <tbody>
        <tr>
          {label('PO')}
          {block.jobs.map((j) =>
            poNameCell(j, j.jobShort, j.missingFromPl ? 'No P&L column for this job — values default to 0' : undefined),
          )}
        </tr>
        <tr>
          {label('Name')}
          {block.jobs.map((j) => poNameCell(j, j.customer))}
        </tr>
        <tr>
          {label('Profit Share')}
          {block.jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">{`$${fmtAccounting(j.profitShare)}`}</td>
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
        {/* §5.6 — thin line between -½ Materials and Total Profit Share. */}
        <tr className="border-b border-b-neutral-500">
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
            <td
              key={j.jobRowId}
              className={`xcell text-right ${j.totalProfitShare < 25 ? 'bg-amber-100' : ''} ${j.dbFloorApplied ? 'bg-sky-50 ring-1 ring-inset ring-sky-400' : ''}`}
              title={
                j.dbFloorApplied
                  ? `Davis-Bacon floor: profit share ${j.dbFloorPriorPay?.toFixed(2) ?? '?'} < ${j.dbFloorHours ?? 0}h × floor ${j.dbFloor?.toFixed(2) ?? '?'}`
                  : undefined
              }
            >
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
        {/* §5.6 — Hourly Pay value cell highlighted yellow. */}
        <tr className="font-semibold">
          <td className="xcell w-44">Hourly Pay</td>
          <td className="xcell text-right" colSpan={block.jobs.length} style={HOURLY_PAY_YELLOW}>
            {fmtAccounting(block.hourlyPay)}
          </td>
        </tr>
        <tr>
          {label('Reimbursements')}
          <td className="xcell text-right text-[#0070C0]">
            <NumCell value={block.reimbursements} onCommit={(n) => onEdit(`psb:${block.key}`, { reimbursements: n ?? 0 })} className="text-[#0070C0]" />
          </td>
          {block.jobs.length > 1 && (
            <td className="xcell" colSpan={block.jobs.length - 1}>
              <MemoCell value={block.reimbMemo} onCommit={(t) => onEdit(`psb:${block.key}`, { reimbMemo: t })} />
            </td>
          )}
        </tr>
        {/* §5.6 — double line between Deductions/Sundries and Total Pay. */}
        <tr className="border-b-4 border-double border-b-neutral-600">
          {label('Deductions/Sundries')}
          <td className="xcell text-right text-[#FF0000]">
            <NumCell value={block.deductions} onCommit={(n) => onEdit(`psb:${block.key}`, { deductions: n ?? 0 })} className="text-[#FF0000]" />
          </td>
          {block.jobs.length > 1 && (
            <td className="xcell" colSpan={block.jobs.length - 1}>
              <MemoCell value={block.deductMemo} onCommit={(t) => onEdit(`psb:${block.key}`, { deductMemo: t })} />
            </td>
          )}
        </tr>
        <SummaryRow label="Total Pay" value={`$${fmtAccounting(block.totalPay)}`} span={block.jobs.length} bold />
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
  weekRange,
  onEdit,
  headerStyle,
}: {
  block: SalaryBlock;
  mileageRate: number;
  weekRange: string;
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
          <td className="xcell w-44 text-[10px]">{weekRange ? `(${weekRange})` : ''}</td>
          <td className="xcell">{block.role}</td>
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
  if (!model) {
    return (
      <div className="flex h-full items-center justify-center">
        <div className="max-w-md rounded-lg border border-neutral-300 bg-white p-6 text-center shadow-sm">
          <div className="mb-2 text-4xl">🧾</div>
          <h2 className="mb-1 text-base font-bold">Payroll Preview needs the P&amp;L</h2>
          <p className="text-xs leading-relaxed text-neutral-500">
            Upload slot 8 — the QuickBooks <b>Profit and Loss by Customer</b> export — to build the payroll blocks.
            (The Employee Name List is recommended too for crew-lead name resolution.)
          </p>
        </div>
      </div>
    );
  }
  return (
    <div className="p-3" style={GARAMOND}>
      {/* §5.4 — title band: centered, bold white on black. */}
      <div className="mb-3 max-w-4xl px-2 py-1.5 text-center text-sm font-bold" style={BLACK}>
        {model.title}
      </div>

      {/* §5.2 — one A→Z sequence: painters + crew-lead blocks; subs last. */}
      {model.painterSequence.map((entry) =>
        entry.kind === 'hourly' ? (
          <HourlyBlockView key={entry.block.key} block={entry.block} onEdit={onEdit} />
        ) : (
          <div key={entry.block.key}>
            <PsBlockView block={entry.block} onEdit={onEdit} />
            {entry.block.tsheets && (
              <div className="mb-4 ml-4 max-w-4xl">
                <HourlyBlockView block={entry.block.tsheets} onEdit={onEdit} tsheets />
              </div>
            )}
          </div>
        ),
      )}

      <SectionHeader title="OFFICE" />
      {model.office.map((b) => (
        <HourlyBlockView key={b.key} block={b} onEdit={onEdit} />
      ))}

      <SectionHeader title="ADMIN" />
      {model.admin.map((b) => (
        <SalaryBlockView key={b.key} block={b} mileageRate={mileageRate} weekRange={model.weekRange} onEdit={onEdit} headerStyle={TAN} />
      ))}

      <SectionHeader title="ESTIMATORS" />
      {model.estimators.map((b) => (
        <SalaryBlockView key={b.key} block={b} mileageRate={mileageRate} weekRange={model.weekRange} onEdit={onEdit} headerStyle={TAN} />
      ))}

      <SectionHeader title="MANAGER SALARY" />
      {model.managers.map((b) => (
        <SalaryBlockView key={b.key} block={b} mileageRate={mileageRate} weekRange={model.weekRange} onEdit={onEdit} headerStyle={MANAGER} />
      ))}
    </div>
  );
}
