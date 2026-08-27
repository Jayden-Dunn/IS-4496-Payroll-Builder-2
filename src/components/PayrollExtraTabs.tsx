/**
 * Round 6 payroll derived tabs — condensed layouts fed from PayrollPreviewModel
 * (same shared model as Payroll Preview). Core = data content + recognizable sketch.
 */
import type { CSSProperties, ReactNode } from 'react';
import { fmtAccounting, formatCrewLeadWoNote, isDarkColor, parseMoneyInput, parsePercentInput } from '../lib/format';
import { CellInput } from './CellInput';
import { buildJournalEntryRows } from '../lib/journalEntry';
import { hourSplitRows } from '../lib/payroll';
import { buildPrFinalCheckRows } from '../lib/prFinalCheck';
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
const NAME_RED: CSSProperties = { backgroundColor: '#C0504D', color: '#fff', fontWeight: 700 };
const NAME_GREEN: CSSProperties = { backgroundColor: '#00B050', color: '#fff', fontWeight: 700 };
const HEADER_GRAY: CSSProperties = { backgroundColor: '#808080', color: '#fff', fontWeight: 700 };
const NOTES_FLAG: CSSProperties = { color: '#FF0000', fontWeight: 700 };
const NOTES_LEAD_BAND: CSSProperties = { backgroundColor: '#B8CCE4', fontWeight: 700 };

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
  placeholder,
}: {
  value: number | null;
  onCommit: (n: number | null) => void;
  pct?: boolean;
  className?: string;
  placeholder?: string;
}): JSX.Element {
  const display = value === null ? '' : pct ? `${Math.round(value * 100)}%` : value.toFixed(2);
  return (
    <CellInput
      value={display}
      className={`text-right ${className ?? ''}`}
      placeholder={placeholder}
      onCommit={(text) => {
        if (pct) {
          const p = parsePercentInput(text);
          onCommit(typeof p === 'number' ? p : null);
        } else {
          onCommit(parseMoneyInput(text));
        }
      }}
    />
  );
}

function MemoCell({ value, onCommit }: { value: string; onCommit: (t: string) => void }): JSX.Element {
  return (
    <CellInput
      value={value}
      onCommit={onCommit}
      className="italic text-neutral-600"
      placeholder="memo…"
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
  hideBooksRows,
  nameStyle,
  highlightFirstJob,
}: {
  block: PsBlock;
  skipWoBand?: boolean;
  defaultPsPct?: number;
  onEdit?: OnEdit;
  showGrossWages?: boolean;
  /**
   * Round 9 §30 — the Final Payroll Breakdown drops the reference rows
   * (Total Income, Materials, Travel, QB Labor-Wages-Subs, Mileage in Books,
   * % Materials). They stay on the main Payroll Preview.
   */
  hideBooksRows?: boolean;
  nameStyle?: CSSProperties;
  highlightFirstJob?: boolean;
}): JSX.Element {
  const jobs = block.jobs.slice(0, 6);
  const psKey = (jobRowId: string): string => `ps:${block.key}:${jobRowId}`;
  const label = (t: string): JSX.Element => <td className="xcell w-40 font-medium">{t}</td>;
  const editable = Boolean(onEdit);
  const bandStyle = nameStyle ?? TAN;

  return (
    <table className="mb-2 w-max min-w-[420px] border-collapse bg-white text-xs">
      <thead>
        <tr>
          <th className="xcell text-left" style={bandStyle}>
            {block.displayName}
            {block.sub && <span className="ml-2 rounded bg-blue-100 px-1 text-[10px] font-normal text-blue-700">sub</span>}
          </th>
          {!skipWoBand &&
            jobs.map((j) => (
              <th key={j.jobRowId} className="xcell min-w-[100px] text-right" style={bandStyle}>
                {pctLabel(j.woPct)}
              </th>
            ))}
          {skipWoBand &&
            jobs.map((j) => (
              <th key={j.jobRowId} className="xcell min-w-[100px]" style={bandStyle} />
            ))}
        </tr>
      </thead>
      <tbody>
        <tr>
          {label('PO')}
          {jobs.map((j) => (
            <td
              key={j.jobRowId}
              className="xcell text-right text-[11px] font-bold"
              style={jobCellStyle(j.colorHex)}
            >
              {j.jobShort}
            </td>
          ))}
        </tr>
        <tr>
          {label('Name')}
          {jobs.map((j) => (
            <td
              key={j.jobRowId}
              className="xcell text-right text-[11px] font-bold"
              style={jobCellStyle(j.colorHex)}
            >
              {j.customer}
            </td>
          ))}
        </tr>
        {(
          [
            // Round 8 §8.1 — SUBBED / SET RATE render blank (no auto share).
            ['Profit Share', (j: PsBlock['jobs'][number]) => (j.manualPay ? '' : money(j.profitShare))],
            ['Work Order Override', (j: PsBlock['jobs'][number]) => (j.workOrderOverride !== null ? money(j.workOrderOverride) : '')],
            ['% Completion', (j: PsBlock['jobs'][number]) => (j.manualPay ? j.manualPay : pctLabel(j.percentCompletion))],
            ['- Labor + Wages', (j: PsBlock['jobs'][number]) => money(j.laborWages)],
            ['-Mileage', (j: PsBlock['jobs'][number]) => money(j.mileage)],
            [
              // Round 9 §32 — full-materials leads are charged all of it.
              block.jobs.some((j) => j.fullMaterials) ? '- Materials' : '-½ Materials',
              (j: PsBlock['jobs'][number]) => money(j.halfMaterials),
            ],
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
                    value={j.manualPay ? null : j.percentCompletion}
                    placeholder={j.manualPay ?? undefined}
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
              <td
                key={j.jobRowId}
                className="xcell text-right"
                title={j.manualPay ? `${j.manualPay} — enter pay manually` : undefined}
              >
                {j.manualPay ? '' : money(j.totalProfitShare)}
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
        {/* Round 9 §5 — spacer between Total Profit Share and Total. */}
        <tr>
          <td className="xcell" colSpan={1 + jobs.length} />
        </tr>
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
            {editable && onEdit ? (
              <NumCell
                value={block.hourlyPay}
                onCommit={(n) => onEdit(`psb:${block.key}`, { hourlyPay: n ?? 0 })}
              />
            ) : (
              money(block.hourlyPay)
            )}
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
        {/* Round 9 §19/§23 — configured extra rows. */}
        {block.extras.map((x) => (
          <tr key={x.key}>
            {label(x.label)}
            <td
              className={`xcell text-right ${x.kind === 'deduction' ? 'text-[#FF0000]' : x.kind === 'reimbursement' ? 'text-[#0070C0]' : ''}`}
              colSpan={Math.max(jobs.length, 1)}
            >
              {editable && onEdit ? (
                <NumCell value={x.amount} onCommit={(n) => onEdit(x.key, { amount: n })} />
              ) : x.amount !== null ? (
                money(x.amount)
              ) : (
                ''
              )}
            </td>
          </tr>
        ))}
        <tr className="font-bold">
          {label('Total Pay')}
          <td
            className="xcell text-right"
            colSpan={Math.max(jobs.length, 1)}
            style={highlightFirstJob ? { backgroundColor: '#FAF5F0' } : undefined}
          >
            ${money(block.totalPay)}
          </td>
        </tr>
        {/* Round 9 §31 — hours pay + profit share for this crew lead. */}
        {showGrossWages && (
          <tr className="font-bold">
            {label('Total Gross Wages:')}
            <td className="xcell text-right" colSpan={Math.max(jobs.length, 1)}>
              ${money(block.totalGrossWages)}
            </td>
          </tr>
        )}
        {hideBooksRows ? null : (
        <>
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
        {/* Round 8 §8.5 — informational, directly beneath Materials. */}
        <tr>
          {label('Travel')}
          {jobs.map((j) => (
            <td key={j.jobRowId} className="xcell text-right">
              {money(j.travel)}
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
        </>
        )}
      </tbody>
    </table>
  );
}

function HourlyReadOnly({
  block,
  tsheets = false,
  blankTaxes = false,
  notesLayout = false,
}: {
  block: HourlyBlock;
  tsheets?: boolean;
  blankTaxes?: boolean;
  notesLayout?: boolean;
}): JSX.Element {
  const bodyCls = tsheets ? 'bg-[#F9F9F9]' : '';
  const showTaxCols = notesLayout || !blankTaxes;
  const colCount = showTaxCols ? (notesLayout ? 7 : 8) : 5;
  const nameStyle = notesLayout && !tsheets ? HEADER_GRAY : tsheets ? NEARWHITE : TAN;
  return (
    <table
      className={`mb-3 w-full max-w-3xl border-collapse bg-white text-xs ${tsheets ? 'border border-neutral-500' : ''}`}
    >
      <thead>
        {tsheets ? (
          <tr>
            <th className="xcell border border-neutral-500 text-center" style={NEARWHITE} colSpan={colCount}>
              {block.employeeName.split(' ')[0]}&apos;s T-SHEETS BREAKDOWN
            </th>
          </tr>
        ) : (
          <tr>
            <th className="xcell py-1 text-center" style={nameStyle} colSpan={colCount}>
              <span className={`text-[13px] font-bold ${notesLayout && !tsheets ? 'text-white' : 'text-black'}`}>
                {block.employeeName}
              </span>
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
          {notesLayout ? (
            <>
              <th className={`xcell w-20 text-center font-bold underline ${bodyCls}`}>(-) taxes</th>
              <th className={`xcell w-20 text-center font-bold underline ${bodyCls}`}>net</th>
            </>
          ) : (
            !blankTaxes && (
              <>
                <th className="xcell w-20" />
                <th className="xcell w-20" />
                <th className="xcell w-20" />
              </>
            )
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
              <td className={`xcell text-[11px] ${bodyCls}`} title={l.dailyOtNote}>
                {l.jobLabel}
              </td>
              <td className={`xcell text-center ${bodyCls}`}>{l.hours.toFixed(2)}</td>
              <td className={`xcell text-center ${bodyCls}`}>{l.rate.toFixed(2)}</td>
              <td className={`xcell text-center ${bodyCls}`}>{money(l.hours * l.rate)}</td>
              {notesLayout ? (
                <>
                  <td className={`xcell ${bodyCls}`} />
                  <td className={`xcell ${bodyCls}`} />
                </>
              ) : (
                !blankTaxes && (
                  <>
                    <td className="xcell text-center font-bold">{isLast ? '(-) taxes' : ''}</td>
                    <td className="xcell text-center font-bold">{isLast ? 'NEST [IRA]' : ''}</td>
                    <td className="xcell text-center font-bold">{isLast ? 'net' : ''}</td>
                  </>
                )
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
          {notesLayout ? (
            <>
              <td className={`xcell border-t border-t-neutral-500 text-center ${bodyCls}`}>
                {block.taxes !== null ? money(block.taxes) : ''}
              </td>
              <td className={`xcell border-t border-t-neutral-500 text-center ${bodyCls}`}>
                {block.net !== null ? money(block.net) : ''}
              </td>
            </>
          ) : (
            !blankTaxes && (
              <>
                <td className="xcell text-center" />
                <td className="xcell text-center" />
                <td className="xcell text-center">{block.net !== null ? money(block.net) : ''}</td>
              </>
            )
          )}
        </tr>
        {/* Round 9 §3 — the NOTES log states the hour split explicitly:
            regular vs daily overtime vs weekly overtime, straight off the
            timesheet's own pay items. */}
        {/* Round 9 §19/§23 — configured extra rows (Gas Reimb, Aflac, …) so the
            NOTES log and the Final Payroll Breakdown show the same pay the
            Payroll Preview does. */}
        {block.extras.map((x) => (
          <tr key={x.key}>
            <td className={`xcell font-medium ${bodyCls}`} colSpan={2}>
              {x.label}
            </td>
            <td className={`xcell ${bodyCls}`} />
            <td className={`xcell ${bodyCls}`} />
            <td
              className={`xcell text-center ${bodyCls} ${x.kind === 'deduction' ? 'text-[#FF0000]' : x.kind === 'reimbursement' ? 'text-[#0070C0]' : ''}`}
            >
              {x.amount !== null ? money(x.amount) : ''}
            </td>
            {showTaxCols && <td className={`xcell ${bodyCls}`} />}
            {showTaxCols && <td className={`xcell ${bodyCls}`} />}
            {showTaxCols && !notesLayout && <td className={`xcell ${bodyCls}`} />}
          </tr>
        ))}
        {block.extras.some((x) => x.kind !== 'note') && (
          <tr className="font-bold">
            <td className={`xcell ${bodyCls}`} colSpan={2}>
              Total Pay
            </td>
            <td className={`xcell ${bodyCls}`} />
            <td className={`xcell ${bodyCls}`} />
            <td className={`xcell text-center ${bodyCls}`}>{money(block.netPay)}</td>
            {showTaxCols && <td className={`xcell ${bodyCls}`} />}
            {showTaxCols && <td className={`xcell ${bodyCls}`} />}
            {showTaxCols && !notesLayout && <td className={`xcell ${bodyCls}`} />}
          </tr>
        )}
        {notesLayout &&
          hourSplitRows(block).map((row) => (
            <tr key={row.label}>
              <td className={`xcell ${bodyCls}`} />
              <td className={`xcell text-[11px] italic text-neutral-600 ${bodyCls}`}>{row.label}</td>
              <td className={`xcell text-center ${bodyCls}`}>{row.hours.toFixed(2)}</td>
              <td className={`xcell ${bodyCls}`} />
              <td className={`xcell text-center ${bodyCls}`}>{money(row.pay)}</td>
              <td className={`xcell ${bodyCls}`} />
              <td className={`xcell ${bodyCls}`} />
            </tr>
          ))}
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
            {/* Round 9 §26/§33 — subcontractors in the salary sections. */}
            {block.sub && <span className="ml-2 text-[10px] font-normal">(subcontractor)</span>}
          </th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td className="xcell w-40 text-[10px]">{weekRange ? `(${weekRange})` : ''}</td>
          <td className="xcell">{block.role}</td>
          <td className="xcell w-28 text-right">{money(block.amount)}</td>
        </tr>
        {/* Round 9 §28 — flat "Reimbursements" row instead of mileage reimb. */}
        {block.reimbursementsLabel ? (
          <tr>
            <td className="xcell">{block.reimbursementsLabel}</td>
            <td className="xcell" />
            <td className="xcell text-right">{money(block.reimbursements)}</td>
          </tr>
        ) : (
          <tr>
            <td className="xcell">mileage reimb</td>
            <td className="xcell text-right">{block.miles} mi</td>
            <td className="xcell text-right">{money(block.mileageReimb)}</td>
          </tr>
        )}
        {/* Round 9 §19/§23/§24/§29 — configured extra rows. */}
        {block.extras.map((x) => (
          <tr key={x.key}>
            <td className="xcell">{x.label}</td>
            <td className="xcell" />
            <td className="xcell text-right">{x.amount !== null ? money(x.amount) : ''}</td>
          </tr>
        ))}
        <tr className="font-semibold">
          <td className="xcell">Total</td>
          <td className="xcell" />
          <td className="xcell text-right">{money(block.total)}</td>
        </tr>
        {(block.taxes !== null || block.net !== null) && (
          <>
            <tr>
              <td className="xcell">(-) taxes</td>
              <td className="xcell" />
              <td className="xcell text-right">{block.taxes !== null ? money(block.taxes) : ''}</td>
            </tr>
            <tr className="font-semibold">
              <td className="xcell">net</td>
              <td className="xcell" />
              <td className="xcell text-right">{block.net !== null ? money(block.net) : ''}</td>
            </tr>
          </>
        )}
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

/** NOTES working-log materials table (export parity). */
function LeadMaterialsTable({ block }: { block: PsBlock }): JSX.Element {
  return (
    <table className="mb-2 w-max min-w-[560px] border-collapse bg-white text-xs">
      <thead>
        <tr>
          <th className="xcell text-left" style={NOTES_LEAD_BAND} colSpan={6}>
            {block.displayName}
            {block.sub && (
              <span className="ml-2 rounded bg-blue-100 px-1 text-[10px] font-normal text-blue-700">sub</span>
            )}
          </th>
        </tr>
        <tr>
          {['PO#', 'Job', 'Reported Materials', 'Total Materials', 'DIFFERENCE', ''].map((h, i) => (
            <th key={i} className="xcell text-left font-bold underline">
              {h}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {block.jobs.slice(0, 6).map((j) => {
          const wo = formatCrewLeadWoNote(j);
          const jobText = (() => {
            const base = `${j.customer} ${j.jobShort}`.trim();
            const tagged = `${base} [PS]`;
            const n = j.notes.trim();
            return n ? `${tagged} — ${n}` : tagged;
          })();
          const flag = wo.flagMissing;
          return (
            <tr key={j.jobRowId}>
              <td
                className="xcell whitespace-nowrap"
                style={{ ...jobCellStyle(j.colorHex), ...(flag ? NOTES_FLAG : {}) }}
              >
                {j.apKey ? `AP${j.apKey}` : j.jobShort}
              </td>
              <td className="xcell" style={flag ? NOTES_FLAG : undefined}>
                {jobText}
              </td>
              <td className="xcell" />
              <td className="xcell text-right whitespace-nowrap" style={flag ? NOTES_FLAG : undefined}>
                {flag ? '?Did we finish?' : money(j.materials)}
              </td>
              <td className="xcell" />
              <td
                className="xcell text-[11px] whitespace-nowrap"
                style={flag ? NOTES_FLAG : { color: '#525252' }}
              >
                {wo.text}
              </td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

function PainterSequence({
  model,
  skipWoBand,
  defaultPsPct,
  blankTaxes,
  showGrossWages,
  hideBooksRows,
  notesMode,
  onEdit,
}: {
  model: PayrollPreviewModel;
  skipWoBand?: boolean;
  defaultPsPct?: number;
  blankTaxes?: boolean;
  showGrossWages?: boolean;
  /** Round 9 §30 — Final Payroll Breakdown only. */
  hideBooksRows?: boolean;
  /** NOTES working log: materials table + T-SHEETS; no PS WO% band. */
  notesMode?: boolean;
  onEdit?: OnEdit;
}): ReactNode {
  return (
    <>
      {model.painterSequence.map((entry) =>
        entry.kind === 'hourly' ? (
          <HourlyReadOnly
            key={entry.block.key}
            block={entry.block}
            blankTaxes={blankTaxes}
            notesLayout={notesMode}
          />
        ) : notesMode ? (
          <div key={entry.block.key} className="mb-4">
            <LeadMaterialsTable block={entry.block} />
            {entry.block.tsheets && (
              <div className="mb-2 ml-4 max-w-3xl">
                <HourlyReadOnly
                  block={entry.block.tsheets}
                  tsheets
                  blankTaxes={blankTaxes}
                  notesLayout
                />
              </div>
            )}
          </div>
        ) : (
          <div key={entry.block.key}>
            <PsChainBody
              block={entry.block}
              skipWoBand={skipWoBand}
              defaultPsPct={defaultPsPct}
              onEdit={onEdit}
              showGrossWages={showGrossWages}
              hideBooksRows={hideBooksRows}
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
        <HourlyReadOnly key={b.key} block={b} blankTaxes={blankTaxes} notesLayout={notesMode} />
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

export function ProfitshareOutputTab({
  model,
  onEdit,
}: {
  model: PayrollPreviewModel | null;
  onEdit?: OnEdit;
}): JSX.Element {
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
              <PsChainBody block={block} nameStyle={NAME_RED} highlightFirstJob onEdit={onEdit} />
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
  onEdit,
}: {
  model: PayrollPreviewModel | null;
  defaultPsPct?: number;
  onEdit?: OnEdit;
}): JSX.Element {
  if (!model) return <EmptyPayroll label="Management Report needs Payroll Preview" />;
  return (
    <div className="p-3" style={GARAMOND}>
      <TitleBand title="Management Report Export" />
      {model.leads.length === 0 ? (
        <p className="text-xs italic text-neutral-500">(no crew leads this week)</p>
      ) : (
        <>
          <table className="mb-2 w-max border-collapse text-xs">
            <thead>
              <tr>
                {['Crew', 'Job 1', 'Job 2', 'Job 3', 'Job 4', 'Job 5', 'Job 6'].map((h) => (
                  <th key={h} className="xcell text-left" style={HEADER_GRAY}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
          </table>
          {model.leads.map((block) => (
            <div key={block.key} className="mb-6">
              <PsChainBody
                block={block}
                defaultPsPct={defaultPsPct}
                nameStyle={NAME_GREEN}
                onEdit={onEdit}
              />
              <div className="h-6" />
            </div>
          ))}
        </>
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
      <div className="mb-3 text-[20px] font-bold" style={GARAMOND}>
        Crew Profit Share
      </div>
      {model.leads.length === 0 ? (
        <p className="text-xs italic text-neutral-500">(no crew leads this week)</p>
      ) : (
        model.leads.map((block) => {
          const jobs = block.jobs.slice(0, 6);
          const psKey = (jobRowId: string): string => `ps:${block.key}:${jobRowId}`;
          return (
            <table key={block.key} className="mb-6 w-max min-w-[480px] border-collapse bg-white text-xs">
              <tbody>
                <tr>
                  <td className="xcell" style={NAME_RED}>
                    {block.displayName}
                  </td>
                  {jobs.map((j) => (
                    <td key={j.jobRowId} className="xcell text-right" style={NAME_RED}>
                      {pctLabel(j.woPct)}
                    </td>
                  ))}
                </tr>
                <tr>
                  <td className="xcell" style={HEADER_GRAY}>
                    PO #&apos;s:
                  </td>
                  {jobs.map((j, i) => (
                    <td
                      key={j.jobRowId}
                      className="xcell text-right text-[11px] font-bold"
                      style={i === 0 ? HOURLY_PAY_YELLOW : HEADER_GRAY}
                    >
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
                        value={j.manualPay ? null : j.percentCompletion}
                        placeholder={j.manualPay ?? undefined}
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
                  <td className="xcell">Hourly Pay</td>
                  <td className="xcell text-right" style={HOURLY_PAY_YELLOW}>
                    <NumCell
                      value={block.hourlyPay}
                      onCommit={(n) => onEdit(`psb:${block.key}`, { hourlyPay: n ?? 0 })}
                    />
                  </td>
                  <td className="xcell" colSpan={Math.max(jobs.length - 1, 1)} />
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
  const gross = model.painterSequence.reduce(
    (s, e) => s + (e.kind === 'hourly' ? e.block.totalPay : e.block.totalPay),
    0,
  ) +
    model.office.reduce((s, b) => s + b.totalPay, 0) +
    [...model.admin, ...model.estimators, ...model.managers].reduce((s, b) => s + b.total, 0);
  return (
    <div className="p-3" style={GARAMOND}>
      <div className="mb-3 flex max-w-4xl items-baseline justify-between gap-4">
        <div className="px-2 py-1.5 text-sm font-bold" style={BLACK}>
          «{model.title}
        </div>
        <div className="text-sm font-bold tabular-nums">{money(gross)}</div>
      </div>
      <PainterSequence model={model} notesMode />
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
        hideBooksRows
      />
    </div>
  );
}

export function JournalEntryTab({ model }: { model: PayrollPreviewModel | null }): JSX.Element {
  if (!model) return <EmptyPayroll label="Journal Entry needs Payroll Preview" />;
  const { rows, amountSum } = buildJournalEntryRows(model);

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
            <td className="xcell text-right">{money(amountSum)}</td>
            <td className="xcell" colSpan={4} />
          </tr>
        </tbody>
      </table>
    </div>
  );
}

/** Round 7 §2 — PR Final Check from parsed P&L. */
export function PrFinalCheckTab({
  pl,
  sourceFileName,
}: {
  pl: PlParseResult | null;
  sourceFileName?: string;
}): JSX.Element {
  if (!pl) {
    return (
      <div className="flex h-full items-center justify-center p-6">
        <div className="max-w-md rounded-lg border border-neutral-300 bg-white p-6 text-center shadow-sm">
          <h2 className="mb-1 text-base font-bold">PR Final Check needs the P&amp;L</h2>
          <p className="text-xs leading-relaxed text-neutral-500">
            Upload the P&amp;L by Customer file to populate this tab.
          </p>
        </div>
      </div>
    );
  }
  const rows = buildPrFinalCheckRows(pl);
  return (
    <div className="p-3" style={GARAMOND}>
      <TitleBand title="PR Final Check" />
      {sourceFileName ? (
        <p className="mb-2 text-[11px] text-neutral-600 italic">Source: {sourceFileName}</p>
      ) : null}
      <table className="w-max min-w-[720px] border-collapse bg-white text-xs">
        <thead>
          <tr>
            {['PO#', 'Job', 'Total Income', 'Materials', 'QB Labor-Wages-Subs', 'Mileage Expense', 'Sundries -'].map(
              (h) => (
                <th key={h} className="xcell text-left font-bold whitespace-nowrap">
                  {h}
                </th>
              ),
            )}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => {
            const job =
              r.customerName && r.jobHeader && !r.jobHeader.includes(r.customerName)
                ? `${r.customerName} ${r.jobHeader}`.trim()
                : r.customerName || r.jobHeader || `AP${r.apKey}`;
            return (
              <tr key={r.apKey}>
                <td className="xcell whitespace-nowrap">{`AP${r.apKey}`}</td>
                <td className="xcell">{job}</td>
                <td className="xcell text-right whitespace-nowrap">{money(r.income)}</td>
                <td className="xcell text-right whitespace-nowrap">{money(r.materials)}</td>
                <td className="xcell text-right whitespace-nowrap">{money(r.laborWagesSubs)}</td>
                <td className="xcell text-right whitespace-nowrap">{money(r.mileage)}</td>
                <td className="xcell text-right whitespace-nowrap">{money(r.sundries)}</td>
              </tr>
            );
          })}
          {rows.length === 0 && (
            <tr>
              <td className="xcell italic text-neutral-500" colSpan={7}>
                (no jobs on P&amp;L)
              </td>
            </tr>
          )}
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
