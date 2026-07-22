/**
 * Round 6 payroll derived tabs — condensed layouts fed from PayrollPreviewModel
 * (same shared model as Payroll Preview). Core = data content + recognizable sketch.
 */
import { useState, type CSSProperties, type ReactNode } from 'react';
import { fmtAccounting, isDarkColor, parseMoneyInput, parsePercentInput } from '../lib/format';
import type {
  HourlyBlock,
  PayrollEdit,
  PayrollPreviewModel,
  PlParseResult,
  PsBlock,
  SalaryBlock,
} from '../lib/types';

type OnEdit = (key: string, patch: PayrollEdit) => void;

const GARAMOND: CSSProperties = { fontFamily: 'Garamond, "EB Garamond", "Times New Roman", serif' };
const BLACK: CSSProperties = { backgroundColor: '#000', color: '#fff' };
const TAN: CSSProperties = { backgroundColor: '#DDD9C4', fontWeight: 700 };
const NEARWHITE: CSSProperties = { backgroundColor: '#F9F9F9', fontWeight: 700 };
const MANAGER: CSSProperties = { backgroundColor: '#D99694', fontWeight: 700 };
const HOURLY_PAY_YELLOW: CSSProperties = { backgroundColor: '#FFFF99' };
const OFFICE_BLUE: CSSProperties = { backgroundColor: '#4F81BD', color: '#fff' };

function EmptyPayroll({ label }: { label: string }): JSX.Element {
  return (
    <div className="flex h-full items-center justify-center p-6">
      <div className="max-w-md rounded-lg border border-neutral-300 bg-white p-6 text-center shadow-sm">
        <h2 className="mb-1 text-base font-bold">{label}</h2>
        <p className="text-xs leading-relaxed text-neutral-500">
          Upload the P&amp;L and press Create Payroll Preview to populate this tab.
        </p>
      </div>
    </div>
  );
}

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

function jobCellStyle(hex: string | null): CSSProperties | undefined {
  if (!hex) return undefined;
  return { backgroundColor: `#${hex}`, color: isDarkColor(hex) ? '#fff' : '#000' };
}

function money(n: number): string {
  return fmtAccounting(n);
}

function pctLabel(p: number | null): string {
  return p !== null ? `${Math.round(p * 100)}%` : '-';
}

/** Shared PS chain rows (WO% band optional). */
function PsChainBody({
  block,
  skipWoBand,
  defaultPsPct,
  onEdit,
  showGrossWages,
}: {
  block: PsBlock;
  skipWoBand?: boolean;
  defaultPsPct?: number;
  onEdit?: OnEdit;
  showGrossWages?: boolean;
}): JSX.Element {
  const jobs = block.jobs.slice(0, 6);
  const psKey = (jobRowId: string): string => `ps:${block.key}:${jobRowId}`;
  const label = (t: string): JSX.Element => <td className="xcell w-40 font-medium">{t}</td>;
  const editable = Boolean(onEdit);

  return (
    <table className="mb-2 w-max min-w-[420px] border-collapse bg-white text-xs">
      <thead>
        <tr>
          <th className="xcell text-left" style={TAN}>
            {block.displayName}
            {block.sub && <span className="ml-2 rounded bg-blue-100 px-1 text-[10px] font-normal text-blue-700">sub</span>}
          </th>
          {!skipWoBand &&
            jobs.map((j) => (
              <th key={j.jobRowId} className="xcell min-w-[100px] text-right" style={TAN}>
                {pctLabel(j.woPct)}
              </th>
            ))}
          {skipWoBand &&
            jobs.map((j) => (
              <th key={j.jobRowId} className="xcell min-w-[100px]" style={TAN} />
            ))}
        </tr>
      </thead>
      <tbody>
        <tr>
          {label('PO')}
          {jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right text-[11px] font-bold" style={jobCellStyle(j.colorHex)}>
              {j.jobShort}
            </td>
          ))}
        </tr>
        <tr>
          {label('Name')}
          {jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right text-[11px] font-bold" style={jobCellStyle(j.colorHex)}>
              {j.customer}
            </td>
          ))}
        </tr>
        {(
          [
            ['Profit Share', (j: PsBlock['jobs'][number]) => money(j.profitShare)],
            ['Work Order Override', (j: PsBlock['jobs'][number]) => (j.workOrderOverride !== null ? money(j.workOrderOverride) : '')],
            ['% Completion', (j: PsBlock['jobs'][number]) => pctLabel(j.percentCompletion)],
            ['- Labor + Wages', (j: PsBlock['jobs'][number]) => money(j.laborWages)],
            ['-Mileage', (j: PsBlock['jobs'][number]) => money(j.mileage)],
            ['-½ Materials', (j: PsBlock['jobs'][number]) => money(j.halfMaterials)],
          ] as [string, (j: PsBlock['jobs'][number]) => string][]
        ).map(([lab, pick]) => (
          <tr key={lab}>
            {label(lab)}
            {jobs.map((j) => (
              <td key={j.jobRowId} className="xcell text-right">
                {lab === 'Work Order Override' && editable && onEdit ? (
                  <NumCell
                    value={j.workOrderOverride}
                    onCommit={(n) => onEdit(psKey(j.jobRowId), { workOrderOverride: n ?? 0 })}
                  />
                ) : lab === '% Completion' && editable && onEdit ? (
                  <NumCell
                    pct
                    value={j.percentCompletion}
                    onCommit={(n) => onEdit(psKey(j.jobRowId), { percentCompletion: n ?? 0 })}
                  />
                ) : (
                  pick(j)
                )}
              </td>
            ))}
          </tr>
        ))}
        <tr className="font-semibold border-t border-neutral-400">
          {label('Total Profit Share')}
          {jobs.map((j) => {
            const note =
              skipWoBand &&
              defaultPsPct !== undefined &&
              j.woPct !== null &&
              Math.abs(j.woPct - defaultPsPct) > 1e-9
                ? `@ ${Math.round(j.woPct * 100)}% PS`
                : '';
            return (
              <td key={j.jobRowId} className="xcell text-right">
                {money(j.totalProfitShare)}
                {note && <span className="ml-1 text-[10px] font-normal text-neutral-600">{note}</span>}
              </td>
            );
          })}
        </tr>
        {jobs.some((j) => j.splitWith > 1) && (
          <tr>
            {label('split w/')}
            {jobs.map((j) => (
              <td key={j.jobRowId} className="xcell text-right">
                {j.splitWith > 1 ? `÷${j.splitWith}` : ''}
              </td>
            ))}
          </tr>
        )}
        <tr className="font-semibold">
          {label('Total')}
          <td className="xcell text-right" colSpan={Math.max(jobs.length, 1)}>
            {money(block.total)}
          </td>
        </tr>
        <tr>
          {label('Mileage - New')}
          <td className="xcell text-right" colSpan={Math.max(jobs.length, 1)}>
            {editable && onEdit ? (
              <NumCell value={block.mileageNew} onCommit={(n) => onEdit(`psb:${block.key}`, { mileageNew: n ?? 0 })} />
            ) : (
              money(block.mileageNew)
            )}
          </td>
        </tr>
        <tr className="font-semibold">
          <td className="xcell w-40">Hourly Pay</td>
          <td className="xcell text-right" colSpan={Math.max(jobs.length, 1)} style={HOURLY_PAY_YELLOW}>
            {money(block.hourlyPay)}
          </td>
        </tr>
        <tr>
          {label('Reimbursements')}
          <td className="xcell text-right text-[#0070C0]">
            {editable && onEdit ? (
              <NumCell
                value={block.reimbursements}
                onCommit={(n) => onEdit(`psb:${block.key}`, { reimbursements: n ?? 0 })}
                className="text-[#0070C0]"
              />
            ) : (
              money(block.reimbursements)
            )}
          </td>
          {jobs.length > 1 && (
            <td className="xcell" colSpan={jobs.length - 1}>
              {editable && onEdit ? (
                <MemoCell value={block.reimbMemo} onCommit={(t) => onEdit(`psb:${block.key}`, { reimbMemo: t })} />
              ) : (
                block.reimbMemo
              )}
            </td>
          )}
        </tr>
        <tr className="border-b-4 border-double border-b-neutral-600">
          {label('Deductions/Sundries')}
          <td className="xcell text-right text-[#FF0000]">
            {editable && onEdit ? (
              <NumCell
                value={block.deductions}
                onCommit={(n) => onEdit(`psb:${block.key}`, { deductions: n ?? 0 })}
                className="text-[#FF0000]"
              />
            ) : (
              money(block.deductions)
            )}
          </td>
          {jobs.length > 1 && (
            <td className="xcell" colSpan={jobs.length - 1}>
              {editable && onEdit ? (
                <MemoCell value={block.deductMemo} onCommit={(t) => onEdit(`psb:${block.key}`, { deductMemo: t })} />
              ) : (
                block.deductMemo
              )}
            </td>
          )}
        </tr>
        <tr className="font-bold">
          {label('Total Pay')}
          <td className="xcell text-right" colSpan={Math.max(jobs.length, 1)}>
            ${money(block.totalPay)}
          </td>
        </tr>
        {showGrossWages && (
          <tr className="font-bold">
            {label('Total Gross Wages')}
            <td className="xcell text-right" colSpan={Math.max(jobs.length, 1)}>
              ${money(block.total + block.hourlyPay)}
            </td>
          </tr>
        )}
        <tr>
          <td className="xcell" colSpan={1 + jobs.length} />
        </tr>
        <tr>
          {label('Total Income')}
          {jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              {money(j.income)}
            </td>
          ))}
        </tr>
        <tr>
          {label('Materials')}
          {jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              {money(j.materials)}
            </td>
          ))}
        </tr>
        <tr>
          {label('QB Labor-Wages-Subs')}
          {jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              {money(j.laborWages)}
            </td>
          ))}
        </tr>
        <tr>
          {label('Mileage in Books')}
          {jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              {money(j.mileage)}
            </td>
          ))}
        </tr>
        <tr>
          {label('% Materials')}
          {jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              {j.income !== 0 ? `${Math.round((j.materials / j.income) * 100)}%` : '—'}
            </td>
          ))}
        </tr>
      </tbody>
    </table>
  );
}

function HourlyReadOnly({
  block,
  tsheets = false,
  blankTaxes = false,
}: {
  block: HourlyBlock;
  tsheets?: boolean;
  blankTaxes?: boolean;
}): JSX.Element {
  const bodyCls = tsheets ? 'bg-[#F9F9F9]' : '';
  return (
    <table
      className={`mb-3 w-full max-w-3xl border-collapse bg-white text-xs ${tsheets ? 'border border-neutral-500' : ''}`}
    >
      <thead>
        {tsheets ? (
          <tr>
            <th className="xcell border border-neutral-500 text-center" style={NEARWHITE} colSpan={5}>
              {block.employeeName.split(' ')[0]}&apos;s T-SHEETS BREAKDOWN
            </th>
          </tr>
        ) : (
          <tr>
            <th className="xcell py-1 text-center" style={TAN} colSpan={5}>
              <span className="text-[13px] font-bold text-black">{block.employeeName}</span>
              {block.crewTag && <span className="ml-1 text-[11px] font-bold text-white">({block.crewTag})</span>}
            </th>
          </tr>
        )}
        <tr>
          <th className={`xcell w-24 text-center font-bold underline ${bodyCls}`}>PO#</th>
          <th className={`xcell text-center font-bold underline ${bodyCls}`}>Job</th>
          <th className={`xcell w-16 text-center font-bold underline ${bodyCls}`}>hrs</th>
          <th className={`xcell w-16 text-center font-bold underline ${bodyCls}`}>rate</th>
          <th className={`xcell w-24 text-center font-bold underline ${bodyCls}`}>total</th>
          {!blankTaxes && (
            <>
              <th className="xcell w-20" />
              <th className="xcell w-20" />
              <th className="xcell w-20" />
            </>
          )}
        </tr>
      </thead>
      <tbody>
        {block.lines.map((l, i) => {
          const isLast = i === block.lines.length - 1;
          return (
            <tr key={l.id}>
              <td
                className={`xcell text-center text-[10px] ${bodyCls}`}
                style={l.poText !== null ? OFFICE_BLUE : jobCellStyle(l.colorHex)}
              >
                {l.poText ?? (l.apKey ? `AP${l.apKey}` : '')}
              </td>
              <td className={`xcell text-[11px] ${bodyCls}`}>{l.jobLabel}</td>
              <td className={`xcell text-center ${bodyCls}`}>{l.hours.toFixed(2)}</td>
              <td className={`xcell text-center ${bodyCls}`}>{l.rate.toFixed(2)}</td>
              <td className={`xcell text-center ${bodyCls}`}>{money(l.hours * l.rate)}</td>
              {!blankTaxes && (
                <>
                  <td className="xcell text-center font-bold">{isLast ? '(-) taxes' : ''}</td>
                  <td className="xcell text-center font-bold">{isLast ? 'NEST [IRA]' : ''}</td>
                  <td className="xcell text-center font-bold">{isLast ? 'net' : ''}</td>
                </>
              )}
            </tr>
          );
        })}
        <tr className="font-semibold">
          <td className={`xcell border-t border-t-neutral-500 ${bodyCls}`} />
          <td className={`xcell border-t border-t-neutral-500 ${bodyCls}`} />
          <td className={`xcell border-t border-t-neutral-500 text-center ${bodyCls}`}>{block.totalHours.toFixed(2)}</td>
          <td className={`xcell border-t border-t-neutral-500 ${bodyCls}`} />
          <td className={`xcell border-t border-t-neutral-500 text-center ${bodyCls}`}>{money(block.totalPay)}</td>
          {!blankTaxes && (
            <>
              <td className="xcell text-center" />
              <td className="xcell text-center" />
              <td className="xcell text-center">{block.net !== null ? money(block.net) : ''}</td>
            </>
          )}
        </tr>
      </tbody>
    </table>
  );
}

function SalaryReadOnly({
  block,
  weekRange,
  headerStyle,
}: {
  block: SalaryBlock;
  weekRange: string;
  headerStyle: CSSProperties;
}): JSX.Element {
  return (
    <table className="mb-3 w-full max-w-xl border-collapse bg-white text-xs">
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
          <td className="xcell w-40 text-[10px]">{weekRange ? `(${weekRange})` : ''}</td>
          <td className="xcell">{block.role}</td>
          <td className="xcell w-28 text-right">{money(block.amount)}</td>
        </tr>
        <tr>
          <td className="xcell">mileage reimb</td>
          <td className="xcell text-right">{block.miles} mi</td>
          <td className="xcell text-right">{money(block.mileageReimb)}</td>
        </tr>
        <tr className="font-semibold">
          <td className="xcell">Total</td>
          <td className="xcell" />
          <td className="xcell text-right">{money(block.amount + block.mileageReimb)}</td>
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

function TitleBand({ title }: { title: string }): JSX.Element {
  return (
    <div className="mb-3 max-w-4xl px-2 py-1.5 text-center text-sm font-bold" style={BLACK}>
      {title}
    </div>
  );
}

function PainterSequence({
  model,
  skipWoBand,
  defaultPsPct,
  blankTaxes,
  showGrossWages,
  onEdit,
}: {
  model: PayrollPreviewModel;
  skipWoBand?: boolean;
  defaultPsPct?: number;
  blankTaxes?: boolean;
  showGrossWages?: boolean;
  onEdit?: OnEdit;
}): ReactNode {
  return (
    <>
      {model.painterSequence.map((entry) =>
        entry.kind === 'hourly' ? (
          <HourlyReadOnly key={entry.block.key} block={entry.block} blankTaxes={blankTaxes} />
        ) : (
          <div key={entry.block.key}>
            <PsChainBody
              block={entry.block}
              skipWoBand={skipWoBand}
              defaultPsPct={defaultPsPct}
              onEdit={onEdit}
              showGrossWages={showGrossWages}
            />
            {entry.block.tsheets && (
              <div className="mb-4 ml-4 max-w-3xl">
                <HourlyReadOnly block={entry.block.tsheets} tsheets blankTaxes={blankTaxes} />
              </div>
            )}
          </div>
        ),
      )}
      <SectionHeader title="OFFICE" />
      {model.office.map((b) => (
        <HourlyReadOnly key={b.key} block={b} blankTaxes={blankTaxes} />
      ))}
      <SectionHeader title="ADMIN" />
      {model.admin.map((b) => (
        <SalaryReadOnly key={b.key} block={b} weekRange={model.weekRange} headerStyle={TAN} />
      ))}
      <SectionHeader title="ESTIMATORS" />
      {model.estimators.map((b) => (
        <SalaryReadOnly key={b.key} block={b} weekRange={model.weekRange} headerStyle={TAN} />
      ))}
      <SectionHeader title="MANAGER SALARY" />
      {model.managers.map((b) => (
        <SalaryReadOnly key={b.key} block={b} weekRange={model.weekRange} headerStyle={MANAGER} />
      ))}
    </>
  );
}

// ---------------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------------

export function ProfitshareOutputTab({ model }: { model: PayrollPreviewModel | null }): JSX.Element {
  if (!model) return <EmptyPayroll label="Profitshare Output needs Payroll Preview" />;
  const leads = model.leads;
  return (
    <div className="overflow-x-auto p-3" style={GARAMOND}>
      <TitleBand title="Profitshare Output" />
      {leads.length === 0 ? (
        <p className="text-xs italic text-neutral-500">(no crew leads this week)</p>
      ) : (
        <div className="flex flex-row items-start gap-0">
          {leads.map((block, i) => (
            <div
              key={block.key}
              className="shrink-0"
              style={{ marginLeft: i === 0 ? 0 : '2rem' /* spacer ≈ 8th column between bands */ }}
            >
              <PsChainBody block={block} />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export function ManagementReportTab({
  model,
  defaultPsPct = 0.38,
}: {
  model: PayrollPreviewModel | null;
  defaultPsPct?: number;
}): JSX.Element {
  if (!model) return <EmptyPayroll label="Management Report needs Payroll Preview" />;
  return (
    <div className="p-3" style={GARAMOND}>
      <TitleBand title="Management Report Export" />
      {model.leads.length === 0 ? (
        <p className="text-xs italic text-neutral-500">(no crew leads this week)</p>
      ) : (
        model.leads.map((block) => (
          <div key={block.key} className="mb-6">
            <PsChainBody block={block} defaultPsPct={defaultPsPct} />
            {/* blank spacer between leads */}
            <div className="h-6" />
          </div>
        ))
      )}
    </div>
  );
}

export function InputNumsTab({
  model,
  onEdit,
}: {
  model: PayrollPreviewModel | null;
  onEdit: OnEdit;
}): JSX.Element {
  if (!model) return <EmptyPayroll label="Input #'s needs Payroll Preview" />;
  return (
    <div className="p-3" style={GARAMOND}>
      <TitleBand title="Input #'s" />
      {model.leads.length === 0 ? (
        <p className="text-xs italic text-neutral-500">(no crew leads this week)</p>
      ) : (
        model.leads.map((block, idx) => {
          const jobs = block.jobs.slice(0, 6);
          const psKey = (jobRowId: string): string => `ps:${block.key}:${jobRowId}`;
          return (
            <table key={block.key} className="mb-6 w-max min-w-[480px] border-collapse bg-white text-xs">
              <tbody>
                <tr>
                  <td className="xcell font-bold" style={TAN}>
                    Crew {idx + 1}
                  </td>
                  {jobs.map((_, i) => (
                    <td key={i} className="xcell text-center font-bold" style={TAN}>
                      Job {i + 1}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="xcell font-semibold">{block.displayName}</td>
                  {jobs.map((j) => (
                    <td key={j.jobRowId} className="xcell text-right">
                      {pctLabel(j.woPct)}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="xcell">PO #&apos;s:</td>
                  {jobs.map((j) => (
                    <td key={j.jobRowId} className="xcell text-right text-[11px]" style={jobCellStyle(j.colorHex)}>
                      {j.jobShort}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="xcell">WO Amount/Contract Override:</td>
                  {jobs.map((j) => (
                    <td key={j.jobRowId} className="xcell text-right">
                      <NumCell
                        value={j.workOrderOverride}
                        onCommit={(n) => onEdit(psKey(j.jobRowId), { workOrderOverride: n ?? 0 })}
                      />
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="xcell">% Completion:</td>
                  {jobs.map((j) => (
                    <td key={j.jobRowId} className="xcell text-right">
                      <NumCell
                        pct
                        value={j.percentCompletion}
                        onCommit={(n) => onEdit(psKey(j.jobRowId), { percentCompletion: n ?? 0 })}
                      />
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="xcell">Reported Materials:</td>
                  <td className="xcell text-neutral-500">Amount</td>
                  <td className="xcell text-neutral-500" colSpan={Math.max(jobs.length - 1, 1)}>
                    Notes
                  </td>
                </tr>
                <tr>
                  <td className="xcell">Reimbursements</td>
                  <td className="xcell text-right text-[#0070C0]">
                    <NumCell
                      value={block.reimbursements}
                      onCommit={(n) => onEdit(`psb:${block.key}`, { reimbursements: n ?? 0 })}
                      className="text-[#0070C0]"
                    />
                  </td>
                  <td className="xcell" colSpan={Math.max(jobs.length - 1, 1)}>
                    <MemoCell value={block.reimbMemo} onCommit={(t) => onEdit(`psb:${block.key}`, { reimbMemo: t })} />
                  </td>
                </tr>
                <tr>
                  <td className="xcell">Deductions/Sundries</td>
                  <td className="xcell text-right text-[#FF0000]">
                    <NumCell
                      value={block.deductions}
                      onCommit={(n) => onEdit(`psb:${block.key}`, { deductions: n ?? 0 })}
                      className="text-[#FF0000]"
                    />
                  </td>
                  <td className="xcell" colSpan={Math.max(jobs.length - 1, 1)}>
                    <MemoCell value={block.deductMemo} onCommit={(t) => onEdit(`psb:${block.key}`, { deductMemo: t })} />
                  </td>
                </tr>
                <tr>
                  <td className="xcell">Miles:</td>
                  <td className="xcell text-right">
                    <NumCell value={block.mileageNew} onCommit={(n) => onEdit(`psb:${block.key}`, { mileageNew: n ?? 0 })} />
                  </td>
                  {jobs.slice(1).map((j) => (
                    <td key={j.jobRowId} className="xcell" />
                  ))}
                </tr>
              </tbody>
            </table>
          );
        })
      )}
    </div>
  );
}

export function NotesTab({
  model,
}: {
  model: PayrollPreviewModel | null;
  mileageRate?: number;
  onEdit?: OnEdit;
}): JSX.Element {
  if (!model) return <EmptyPayroll label="Notes needs Payroll Preview" />;
  return (
    <div className="p-3" style={GARAMOND}>
      <TitleBand title={model.title} />
      <PainterSequence model={model} />
    </div>
  );
}

export function FinalPayrollTab({
  model,
  defaultPsPct = 0.38,
}: {
  model: PayrollPreviewModel | null;
  mileageRate?: number;
  onEdit?: OnEdit;
  defaultPsPct?: number;
}): JSX.Element {
  if (!model) return <EmptyPayroll label="Final Payroll Breakdown needs Payroll Preview" />;
  return (
    <div className="p-3" style={GARAMOND}>
      <TitleBand title={model.title} />
      <PainterSequence
        model={model}
        skipWoBand
        defaultPsPct={defaultPsPct}
        blankTaxes
        showGrossWages
      />
    </div>
  );
}

export function JournalEntryTab({ model }: { model: PayrollPreviewModel | null }): JSX.Element {
  if (!model) return <EmptyPayroll label="Journal Entry needs Payroll Preview" />;

  interface JobAgg {
    apKey: string;
    jobShort: string;
    customer: string;
    amount: number;
    parts: { lead: string; amt: number; woPct: number | null; split: number }[];
  }
  const byAp = new Map<string, JobAgg>();
  for (const lead of model.leads) {
    for (const j of lead.jobs) {
      if (!j.apKey) continue;
      let agg = byAp.get(j.apKey);
      if (!agg) {
        agg = { apKey: j.apKey, jobShort: j.jobShort, customer: j.customer, amount: 0, parts: [] };
        byAp.set(j.apKey, agg);
      }
      const amt = j.totalAfterSplit;
      agg.amount += amt;
      agg.parts.push({ lead: lead.displayName, amt, woPct: j.woPct, split: j.splitWith });
    }
  }
  const sorted = [...byAp.values()].sort((a, b) =>
    a.jobShort.localeCompare(b.jobShort, 'en', { sensitivity: 'base' }),
  );

  type Row = { account: string; amount: number; memo: string; po: string; className: string; job: string };
  const rows: Row[] = [];
  let amountSum = 0;

  for (const j of sorted) {
    const pctParts = j.parts.map((p) => `${p.lead} (${p.amt.toFixed(2)})`).join(', ');
    const firstPct = j.parts[0]?.woPct;
    const pctLabelStr = firstPct !== null && firstPct !== undefined ? `${Math.round(firstPct * 100)}%` : '';
    const splitNote = j.parts.some((p) => p.split > 1)
      ? `*split x${Math.max(...j.parts.map((p) => p.split))}*`
      : '';
    const amt = Math.round(j.amount * 100) / 100;
    rows.push({
      account: 'Employee Labor',
      amount: amt,
      memo: `${j.jobShort} [${pctLabelStr}${splitNote}]:  ${pctParts}`,
      po: `AP${j.apKey}`,
      className: 'Painting',
      job: `${j.customer} ${j.jobShort}`.trim(),
    });
    amountSum += amt;
  }

  for (const lead of model.leads) {
    if (lead.mileageNew) {
      rows.push({
        account: 'Mileage',
        amount: lead.mileageNew,
        memo: `mileage reimb:  ${lead.displayName}`,
        po: 'Allpro',
        className: 'Office',
        job: '',
      });
      amountSum += lead.mileageNew;
    }
    if (lead.reimbursements) {
      rows.push({
        account: 'Mileage',
        amount: lead.reimbursements,
        memo: lead.reimbMemo || `reimb: ${lead.displayName}`,
        po: 'Allpro',
        className: 'Office',
        job: '',
      });
      amountSum += lead.reimbursements;
    }
    if (lead.deductions && lead.deductMemo) {
      rows.push({
        account: 'Sundries -',
        amount: lead.deductions,
        memo: `(${lead.displayName}) ${lead.deductMemo}`,
        po: '',
        className: 'Painting',
        job: '',
      });
      amountSum += lead.deductions;
    }
  }
  for (const b of [...model.admin, ...model.estimators, ...model.managers]) {
    if (!b.mileageReimb) continue;
    rows.push({
      account: 'Mileage',
      amount: b.mileageReimb,
      memo: `mileage reimb:  ${b.name} - ${b.miles} miles`,
      po: 'Allpro',
      className: 'Office',
      job: '',
    });
    amountSum += b.mileageReimb;
  }

  return (
    <div className="p-3" style={GARAMOND}>
      <TitleBand title="Journal Entry" />
      <table className="w-max min-w-[720px] border-collapse bg-white text-xs">
        <thead>
          <tr>
            {['Account', 'Amount', 'Memo', 'PO#', 'Class', 'Job'].map((h) => (
              <th key={h} className="xcell text-left font-bold">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}>
              <td className="xcell">{r.account}</td>
              <td className="xcell text-right">{money(r.amount)}</td>
              <td className="xcell text-[11px]">{r.memo}</td>
              <td className="xcell">{r.po}</td>
              <td className="xcell">{r.className}</td>
              <td className="xcell text-[11px]">{r.job}</td>
            </tr>
          ))}
          <tr className="font-bold border-t border-neutral-500">
            <td className="xcell">Total</td>
            <td className="xcell text-right">{money(Math.round(amountSum * 100) / 100)}</td>
            <td className="xcell" colSpan={4} />
          </tr>
        </tbody>
      </table>
    </div>
  );
}

export function JobsListTab({
  model,
  pl: _pl,
}: {
  model: PayrollPreviewModel | null;
  pl?: PlParseResult | null;
}): JSX.Element {
  void _pl;
  if (!model) return <EmptyPayroll label="Jobs needs Payroll Preview" />;

  const map = new Map<string, string>();
  for (const lead of model.leads) {
    for (const j of lead.jobs) {
      if (!j.apKey) continue;
      const po = `AP${j.apKey}`;
      if (!map.has(po)) map.set(po, `${j.customer} ${j.jobShort}`.trim());
    }
  }
  for (const entry of model.painterSequence) {
    const block = entry.kind === 'hourly' ? entry.block : entry.block.tsheets;
    if (!block) continue;
    for (const l of block.lines) {
      if (!l.apKey || l.apKey === 'OFFICE') continue;
      const po = `AP${l.apKey}`;
      if (!map.has(po)) map.set(po, l.jobLabel);
    }
  }
  const rows = [...map.entries()]
    .map(([poLabel, jobLabel]) => ({ poLabel, jobLabel }))
    .sort((a, b) => a.jobLabel.localeCompare(b.jobLabel, 'en', { sensitivity: 'base' }));

  return (
    <div className="p-3" style={GARAMOND}>
      <TitleBand title="Jobs" />
      <table className="w-max min-w-[360px] border-collapse bg-white text-xs">
        <thead>
          <tr>
            <th className="xcell text-left font-bold">PO#</th>
            <th className="xcell text-left font-bold">Job</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.poLabel}>
              <td className="xcell">{r.poLabel}</td>
              <td className="xcell">{r.jobLabel}</td>
            </tr>
          ))}
          {rows.length === 0 && (
            <tr>
              <td className="xcell italic text-neutral-500" colSpan={2}>
                (no jobs)
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

/** Round 6 §4.10 — empty blocks for people without hours this week. */
export function InactiveEmployeesTab({ model }: { model: PayrollPreviewModel | null }): JSX.Element {
  if (!model) return <EmptyPayroll label="Inactive Employees needs Payroll Preview" />;

  const blocks = model.inactiveEmployees;
  if (blocks.length === 0) {
    return (
      <div className="p-4 text-xs text-neutral-500 italic" style={GARAMOND}>
        No inactive or absent employees without hours this week.
      </div>
    );
  }

  return (
    <div className="space-y-4 p-3" style={GARAMOND}>
      {blocks.map((p) => (
        <div key={p.name} className="w-max min-w-[480px] border border-neutral-300 bg-white text-xs">
          <div className="flex items-baseline justify-between px-2 py-1" style={TAN}>
            <span className="font-bold">{p.displayName}</span>
            <span className="text-[11px]">{p.rateNote}</span>
          </div>
          {p.isLead && (
            <table className="w-full border-collapse">
              <thead>
                <tr>
                  {['PO#', 'Job', 'Reported Materials', 'Total Materials', 'DIFFERENCE'].map((h) => (
                    <th key={h} className="xcell text-left font-bold underline">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td className="xcell italic text-neutral-400" colSpan={5}>
                    (empty)
                  </td>
                </tr>
              </tbody>
            </table>
          )}
          <table className="w-full border-collapse">
            <thead>
              <tr>
                {['PO#', 'Job', 'hrs', 'rate', 'total'].map((h) => (
                  <th key={h} className="xcell text-left font-bold underline">
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="xcell" colSpan={2} />
                <td className="xcell text-right">0.00</td>
                <td className="xcell text-right">{p.rateNote.replace('/hr', '')}</td>
                <td className="xcell text-right">0.00</td>
              </tr>
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}
