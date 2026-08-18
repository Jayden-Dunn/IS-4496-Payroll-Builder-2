import { addDays, dateOnly, fmtDate, fmtHours } from './format';
import type {
  EnrichedTimesheetRow,
  ExceptionGroup,
  ExceptionItem,
  PivotByJobModel,
  ReportRow,
} from './types';

export interface ExceptionsInput {
  reportRows: ReportRow[];
  enriched: EnrichedTimesheetRow[];
  lunchRemoved: number;
  payDate: Date | null;
  /** Round 8 §4.3 — supplies the carried jobs whose customer never resolved. */
  pivotByJob?: PivotByJobModel | null;
}

/** Total item count across groups (drives the "Exceptions (n)" tab badge). */
export function countExceptions(groups: ExceptionGroup[]): number {
  return groups.reduce((n, g) => n + g.items.length, 0);
}

export function buildExceptions(input: ExceptionsInput): ExceptionGroup[] {
  const { reportRows, enriched, lunchRemoved, payDate } = input;
  const groups: ExceptionGroup[] = [];
  const included = reportRows.filter((r) => r.included);

  // 1. Excluded events (no AP# → default-excluded, or excluded by the user).
  {
    const items: ExceptionItem[] = reportRows
      .filter((r) => !r.included)
      .map((r) => ({
        message: r.apKey
          ? `"${r.event}" — excluded (AP ${r.apKey})`
          : `"${r.event}" — no AP# found in the title (excluded by default)`,
        rowId: r.id,
      }));
    groups.push({
      key: 'excluded',
      title: 'Excluded events',
      note: 'Use the Include toggle on the TimeNavi Export tab to bring an event back in.',
      items,
    });
  }

  // 2. Not clocked in.
  groups.push({
    key: 'notClockedIn',
    title: 'Not clocked in',
    note: 'Calendar jobs with zero timesheet hours. Utah jobs with no Utah crew in the timesheet are expected here.',
    items: included
      .filter((r) => r.inQboTime === 'NotClockedIn')
      .map((r) => ({ message: `AP ${r.apKey ?? '—'} — "${r.event}" has no timesheet hours`, rowId: r.id })),
  });

  // 3. Not on calendar: timesheet apKeys (with hours, excluding OFFICE) that
  // match no included event.
  {
    const agg = new Map<string, { jobcode2: string; hours: number; employees: Set<string> }>();
    for (const r of enriched) {
      if (!r.notOnCalendar || r.hours === 0 || !r.apKey) continue;
      let a = agg.get(r.apKey);
      if (!a) {
        a = { jobcode2: r.jobcode2, hours: 0, employees: new Set() };
        agg.set(r.apKey, a);
      }
      a.hours += r.hours;
      a.employees.add(`${r.fname} ${r.lname}`.trim() || r.username);
    }
    groups.push({
      key: 'notOnCalendar',
      title: 'Not on calendar',
      note: 'Timesheet jobs with hours that match no included TimeNavi event.',
      items: [...agg.entries()].map(([apKey, a]) => ({
        message: `AP ${apKey} — ${a.jobcode2}: ${fmtHours(a.hours)} h (${[...a.employees].join(', ')})`,
      })),
    });
  }

  // 4. Estimator #N/A with hours.
  {
    const agg = new Map<string, { hours: number; employees: Set<string> }>();
    for (const r of enriched) {
      if (r.estimator !== '#N/A' || r.hours === 0) continue;
      const key = r.jobcode2 || '(blank jobcode_2)';
      let a = agg.get(key);
      if (!a) {
        a = { hours: 0, employees: new Set() };
        agg.set(key, a);
      }
      a.hours += r.hours;
      a.employees.add(`${r.fname} ${r.lname}`.trim() || r.username);
    }
    groups.push({
      key: 'estimatorNA',
      title: 'Estimator #N/A with hours',
      note: 'Timesheet rows whose estimator could not be resolved from either sales report.',
      items: [...agg.entries()].map(([job, a]) => ({
        message: `${job} — ${fmtHours(a.hours)} h (${[...a.employees].join(', ')})`,
      })),
    });
  }

  // 5. Estimator mismatch (calendar title initials vs books). Books win.
  groups.push({
    key: 'mismatch',
    title: 'Estimator mismatch',
    note: 'Event-title initials differ from the books lookup. The books win — this is informational.',
    items: included
      .filter((r) => r.titleInitials && r.estimator && r.titleInitials !== r.estimator)
      .map((r) => ({
        message: `AP ${r.apKey ?? '—'}: calendar says (${r.titleInitials}), books say ${r.estimator} — "${r.event}"`,
        rowId: r.id,
      })),
  });

  // 6. Duplicate AP# among included events.
  {
    const byKey = new Map<string, ReportRow[]>();
    for (const r of included) {
      if (!r.apKey) continue;
      const list = byKey.get(r.apKey) ?? [];
      list.push(r);
      byKey.set(r.apKey, list);
    }
    groups.push({
      key: 'duplicates',
      title: 'Duplicate AP#',
      note: 'Lookups use the first match — verify these events are really the same job.',
      items: [...byKey.entries()]
        .filter(([, list]) => list.length > 1)
        .map(([apKey, list]) => ({
          message: `AP ${apKey} appears in ${list.length} included events: ${list.map((r) => `"${r.event}"`).join(' · ')}`,
          rowId: list[0].id,
        })),
    });
  }

  // 7. Date-range warnings.
  {
    const items: ExceptionItem[] = [];
    if (payDate) {
      const weekEnd = addDays(payDate, -7);
      const weekStart = addDays(weekEnd, -6);
      for (const r of included) {
        if (!r.startDate) continue;
        const start = dateOnly(r.startDate);
        const end = dateOnly(r.endDate ?? r.startDate);
        const overlaps = start.getTime() <= weekEnd.getTime() && end.getTime() >= weekStart.getTime();
        if (!overlaps) {
          items.push({
            message: `AP ${r.apKey ?? '—'} — "${r.event}" (${fmtDate(r.startDate)}–${fmtDate(r.endDate)}) does not overlap the pay week ${fmtDate(weekStart)}–${fmtDate(weekEnd)}`,
            rowId: r.id,
          });
        }
      }
    }
    groups.push({
      key: 'dateRange',
      title: 'Date-range warnings',
      note: 'TimeNavi sometimes reports multi-day event end dates one day early — the app shows the file as-is and does not auto-correct.',
      items,
    });
  }

  // 8. Job total not found.
  groups.push({
    key: 'noJobTotal',
    title: 'Job total not found',
    note: 'No match in either Sales by Customer Summary — Job Total $ shows #N/A.',
    items: included
      .filter((r) => r.apKey && !r.jobTotalFound)
      .map((r) => ({ message: `AP ${r.apKey} — "${r.event}"`, rowId: r.id })),
  });

  // 9. Carried jobs whose customer could not be resolved (Round 8 §4.3).
  groups.push({
    key: 'carriedCustomer',
    title: 'Carried job customer not resolved',
    note: 'No AP# match in either Sales by Customer Summary and none in the P&L header structure — the carried name is left as-is.',
    items: (input.pivotByJob?.unresolvedCarriedCustomers ?? []).map((apKey) => ({
      message: `AP ${apKey} — customer not resolved for carried job ${apKey}`,
    })),
  });

  // 10. Lunch rows removed.
  groups.push({
    key: 'lunch',
    title: 'Lunch rows removed',
    items: lunchRemoved > 0 ? [{ message: `${lunchRemoved} lunch row${lunchRemoved === 1 ? '' : 's'} removed from the timesheet` }] : [],
    note: lunchRemoved === 0 ? 'No lunch rows in this timesheet.' : undefined,
  });

  return groups;
}
