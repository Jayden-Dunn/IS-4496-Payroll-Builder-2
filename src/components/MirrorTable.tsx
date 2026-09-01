import type { CellValue } from '../lib/types';
import type { SalesByCustMirrorRow, SalesByRepMirrorRow, TimesheetMirrorRow } from '../lib/mirrorTabs';
import { fmtHoursFixed } from '../lib/format';

function fmtCell(v: CellValue): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    return `${v.getMonth() + 1}/${v.getDate()}/${v.getFullYear()}`;
  }
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : fmtHoursFixed(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

interface MirrorTableProps {
  headers: string[];
  rows: CellValue[][];
  /** Optional sticky first-column emphasis. */
  boldFirstCol?: boolean;
  /**
   * Round 10 — Sales by Rep: clamp every column so A–M fit on screen without
   * scrolling right. Long text is clipped to the column and kept in the
   * cell's tooltip, matching the widths the workbook exports.
   */
  compact?: boolean;
}

/** Round 6 — generic read-only mirror grid for %'s helper tabs. */
export function MirrorTable({ headers, rows, boldFirstCol, compact }: MirrorTableProps): JSX.Element {
  /** Round 10 — the same width buckets the export uses, in rem. */
  const widthFor = (header: string, i: number): string | undefined => {
    if (!compact) return undefined;
    if (i === 0) return '3.5rem';
    const h = header.trim();
    if (/^(customer|name(\s*[12])?|memo(\/description)?)$/i.test(h)) return '8rem';
    if (/^product\s*\/\s*service$/i.test(h)) return '5.5rem';
    // A date needs room for M/D/YYYY; everything else can stay tight.
    if (/^date$/i.test(h)) return '5rem';
    return '3.5rem';
  };
  return (
    <div className="p-2">
      <table className={`min-w-[320px] border-collapse bg-white text-xs ${compact ? 'table-fixed w-auto' : 'w-max'}`}>
        <thead>
          <tr>
            {headers.map((h, i) => (
              <th
                key={i}
                className="xcell2 truncate text-left font-bold whitespace-nowrap"
                style={compact ? { width: widthFor(h, i), maxWidth: widthFor(h, i) } : undefined}
                title={h}
              >
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 ? (
            <tr>
              <td className="xcell text-neutral-400 italic" colSpan={Math.max(headers.length, 1)}>
                No rows
              </td>
            </tr>
          ) : (
            rows.map((row, ri) => (
              <tr key={ri}>
                {headers.map((h, ci) => {
                  const text = fmtCell(row[ci] ?? null);
                  return (
                    <td
                      key={ci}
                      className={`xcell whitespace-nowrap ${compact ? 'truncate' : ''} ${boldFirstCol && ci === 0 ? 'font-bold' : ''}`}
                      style={compact ? { width: widthFor(h, ci), maxWidth: widthFor(h, ci) } : undefined}
                      title={compact ? text : undefined}
                    >
                      {text}
                    </td>
                  );
                })}
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
}

/** §3.1 / §3.2 Sales by Rep / UT SalesbyRep. */
export function SalesByRepTable({
  headers,
  rows,
}: {
  headers: string[];
  rows: SalesByRepMirrorRow[];
}): JSX.Element {
  const cols = headers.length > 0 ? headers.slice(0, 12) : Array.from({ length: 12 }, (_, i) => `Col ${i + 1}`);
  const allHeaders = ['AP#', ...cols, 'Date', 'QBO Job Name'];
  const grid = rows.map((r) => {
    const body = r.cols.slice(0, 12);
    while (body.length < 12) body.push(null);
    return [r.apKey, ...body, r.dateCopy, r.qboJobName];
  });
  // Round 10 — compact widths so columns A–M are visible without scrolling.
  return <MirrorTable headers={allHeaders} rows={grid} boldFirstCol compact />;
}

/** §3.3 / §3.4 Sales by Cust mirrors. */
export function SalesByCustTable({ rows }: { rows: SalesByCustMirrorRow[] }): JSX.Element {
  return (
    <MirrorTable
      headers={['', 'Total', 'AP#', 'Job total']}
      rows={rows.map((r) => [r.text, r.value, r.apKey, r.jobTotal])}
    />
  );
}

const TS_HEADERS = [
  'username',
  'payroll_id',
  'fname',
  'lname',
  'number',
  'group',
  'local_date',
  'local_day',
  'local_start_time',
  'local_end_time',
  'tz',
  'hours',
  'jobcode_1',
  'jobcode_2',
  'billable',
  'class',
  'department',
  'payroll_item',
  'service_item',
  'location',
  'notes',
  'approved_status',
  'has_flags',
  'flag_types',
  'Apname',
  'AP#',
  'Estimator',
  'QBO Job Name',
  'Not on Calendar',
];

/** §3.5 timesheets mirror. */
export function TimesheetMirrorTable({ rows }: { rows: TimesheetMirrorRow[] }): JSX.Element {
  const grid = rows.map((r) => {
    const raw = r.raw.slice(0, 24);
    while (raw.length < 24) raw.push(null);
    return [...raw, r.apname, r.apKey, r.estimator, r.qboJobName, r.notOnCalendar];
  });
  return <MirrorTable headers={TS_HEADERS} rows={grid} />;
}
