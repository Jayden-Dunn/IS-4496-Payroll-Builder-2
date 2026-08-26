/**
 * Round 3 §6 — "Payroll Exceptions": the standing red-flag rules from the
 * payroll process doc plus this round's checks. Same grouped/linked shape as
 * the %'s Exceptions panel.
 */
import { fmtHours } from './format';
import {
  carryDisplayMap,
  jobDisplayPair,
  salesDisplayMap,
  timesheetDisplayMap,
} from './jobNames';
import type { LastWeekPbjRow } from './parsers';
import type {
  EnrichedTimesheetRow,
  ExceptionGroup,
  ExceptionItem,
  PayrollPreviewModel,
  PivotByJobModel,
  PlParseResult,
  ReportRow,
  SalesRow,
} from './types';

export interface PayrollExceptionsInput {
  pbj: PivotByJobModel;
  /** The NV `PAYROLL PREVIEW` model. */
  preview: PayrollPreviewModel;
  /** Round 8 §7.3 — the `Payroll Preview - UT` model, when the week has one. */
  previewUt?: PayrollPreviewModel | null;
  reportRows: ReportRow[];
  /** Slot B — the NV P&L. */
  pl: PlParseResult | null;
  /** Round 8 §7.2 — Slot C, the optional UT P&L. */
  plUt?: PlParseResult | null;
  /** Directory uploaded? When false, lead resolution used timesheet names. */
  hasDirectory: boolean;
  /** Threshold under which Total Profit Share raises a review item (§6.7). */
  psThreshold?: number;
  /** Round 5 §4 — naming-chain inputs for the informational AP#/unresolved flags. */
  enriched?: EnrichedTimesheetRow[];
  nvSales?: readonly SalesRow[] | null;
  utSales?: readonly SalesRow[] | null;
  lastWeekPbj?: LastWeekPbjRow[] | null;
}

const PS_THRESHOLD_DEFAULT = 25;

export function buildPayrollExceptions(input: PayrollExceptionsInput): ExceptionGroup[] {
  const { pbj, preview, reportRows, pl } = input;
  const groups: ExceptionGroup[] = [];
  const jobs = pbj.customers.flatMap((c) => c.jobs.map((j) => ({ job: j, customer: c.customer })));
  const included = reportRows.filter((r) => r.included);

  // Round 8 §7.3 — every band rule runs over BOTH entities' models. Only the
  // "missing from P&L" rule cares which model it came from, because each one
  // resolves against its own upload.
  const previews: { model: PayrollPreviewModel; pl: PlParseResult | null }[] = [
    { model: preview, pl },
    ...(input.previewUt ? [{ model: input.previewUt, pl: input.plUt ?? null }] : []),
  ];
  const allLeads = previews.flatMap((p) => p.model.leads);
  const allPainters = previews.flatMap((p) => p.model.painters);
  const allOffice = previews.flatMap((p) => p.model.office);
  const allBlocks = [
    ...allPainters,
    ...allOffice,
    ...allLeads.flatMap((l) => (l.tsheets ? [l.tsheets] : [])),
  ];

  /**
   * Round 9 §9/§11 — every non-gray job below the Grand Total must reach a
   * Payroll Preview band. A row can only fail to land when nothing names its
   * crew lead, so say so instead of dropping it silently (Carlos Ruiz's
   * AP25264 and Henry Limay Flores's AP25228 both went missing this way).
   */
  {
    const items: ExceptionItem[] = [];
    const bandedAps = new Set<string>();
    for (const p of previews) {
      for (const lead of p.model.leads) {
        for (const j of lead.jobs) if (j.apKey) bandedAps.add(j.apKey);
      }
    }
    for (const row of [...pbj.balances, ...pbj.utPayroll]) {
      if (row.excludedFromPayroll) continue; // gray = deliberately off payroll
      if (row.apKey === 'OFFICE' || !row.apKey) continue;
      if (bandedAps.has(row.apKey)) continue;
      const where = row.job || row.customer || `AP${row.apKey}`;
      items.push({
        message: `AP${row.apKey} ${where} — below Grand Total with no timesheet hours and no crew lead on the Pivot by Job row, so it has no band to pay from`,
      });
    }
    groups.push({
      key: 'carriedJobNotPaid',
      title: 'Job below Grand Total not on any payroll band',
      note: 'Non-gray rows below the Grand Total pay through their crew lead\u2019s band. Fill in the Crew Lead cell on the Pivot by Job tab of the completed %\u2019s workbook.',
      items,
    });
  }

  // 1. Crew-lead resolution problems.
  {
    const items: ExceptionItem[] = [];
    for (const { job } of jobs) {
      if (job.analysis.apKey === 'OFFICE' || job.analysis.hourly) continue;
      const res = job.analysis.crewLeadResolution;
      if (!res) {
        if (job.analysis.onCalendar !== 'NOT ON CALENDAR' && !job.analysis.crewLead) {
          // Round 9 §6 — the payroll process reads the Crew Lead column, so
          // point at the cell the user actually fills in.
          items.push({
            message: `${job.jobcode2} — the Crew Lead cell on Pivot by Job is empty, so this job has no band to pay from`,
          });
        }
        continue;
      }
      if (!res.raw && !job.analysis.crewLead) {
        items.push({ message: `${job.jobcode2} — no crew lead on the Pivot by Job row` });
        continue;
      }
      for (const f of res.flags) {
        const label =
          f.kind === 'noMatch'
            ? `no matching employee or subcontractor for "${f.candidate}"`
            : f.kind === 'ambiguous'
              ? `multiple equally-preferred matches for "${f.candidate}"`
              : f.kind === 'notACrewLead'
                ? `"${f.candidate}" matched an employee who is not a known crew lead`
                : `only match for "${f.candidate}" is marked inactive (*) — used, verify`;
        items.push({ message: `${job.jobcode2} — (${res.raw}): ${label}` });
      }
    }
    groups.push({
      key: 'leadResolution',
      title: 'Crew-lead resolution',
      note: input.hasDirectory
        ? 'Raw parenthetical shown; fix the Crew Lead cell on Pivot by Job.'
        : 'Employee Name List not saved in Settings — resolution used timesheet names only.',
      items,
    });
  }

  // 2. Hours on a job not on the calendar (piece-rate only).
  groups.push({
    key: 'hoursNotOnCalendar',
    title: 'Hours on a job not on the calendar',
    note: 'Red flag: jobs with hours not on the entry (calendar). Verify the hours are on the right job.',
    items: jobs
      .filter(
        ({ job }) =>
          job.analysis.onCalendar === 'NOT ON CALENDAR' && !job.analysis.hourly && job.analysis.apKey !== 'OFFICE',
      )
      .map(({ job }) => ({
        message: `${job.jobcode2} — ${fmtHours(job.total)} h, AP ${job.analysis.apKey ?? '—'} matches no included TimeNavi event`,
      })),
  });

  // 3. Calendar job with no hours (piece-rate, not subbed).
  {
    const timesheetAps = new Set(jobs.map(({ job }) => job.analysis.apKey).filter(Boolean));
    const subAps = new Set<string>();
    for (const b of allLeads) {
      if (!b.sub) continue;
      for (const j of b.jobs) if (j.apKey) subAps.add(j.apKey);
    }
    groups.push({
      key: 'calendarNoHours',
      title: 'Calendar job with no hours',
      note: 'Red flag: job on entry with no hours. Subcontracted jobs are excluded (they never clock in).',
      items: included
        .filter(
          (r) =>
            r.apKey &&
            !timesheetAps.has(r.apKey) &&
            !subAps.has(r.apKey) &&
            r.percent !== 'HOURLY' &&
            !/^utah\b|^reno\b/i.test(r.calendar.trim()),
        )
        .map((r) => ({ message: `AP ${r.apKey} — "${r.event}" has no timesheet hours`, rowId: r.id })),
    });
  }

  // 4. Hours 40+ / under 10.
  {
    const items: ExceptionItem[] = [];
    for (const b of allBlocks) {
      const otLine = b.lines.find((l) => l.ot);
      if (otLine) {
        items.push({
          message: `${b.employeeName} — ${fmtHours(b.totalHours)} h total: ${fmtHours(otLine.hours)} h OT at 1.5× on "${otLine.jobLabel.replace(' [OT >40 hrs/wk]', '')}"`,
        });
      }
    }
    for (const b of allPainters) {
      if (!b.inactive && b.totalHours > 0 && b.totalHours < 10) {
        items.push({ message: `${b.employeeName} — only ${fmtHours(b.totalHours)} h this week (under 10)` });
      }
    }
    groups.push({
      key: 'hours40or10',
      title: 'Hours 40+ / under 10',
      note: 'Overtime is split automatically (1.5× on the last job worked); under-10 painters may be missing punches.',
      items,
    });
  }

  // 5. Missing pay inputs.
  {
    const items: ExceptionItem[] = [];
    for (const { job } of jobs) {
      if (job.analysis.hourly || job.analysis.apKey === 'OFFICE') continue;
      if (job.analysis.woPct === null) {
        items.push({ message: `${job.jobcode2} — no WO% (hourly? set '-' or a % on Pivot by Job)` });
      }
      if (job.analysis.woOverride === null && job.analysis.apKey && input.pl && !input.pl.byJob.get(job.analysis.apKey)?.income) {
        items.push({ message: `${job.jobcode2} — no WO Override and no P&L income (the "????" case)` });
      }
      if (typeof job.analysis.percent !== 'number') {
        items.push({ message: `${job.jobcode2} — piece-rate job with no % of Completion on the TimeNavi grid` });
      }
    }
    for (const b of allBlocks) {
      if (b.lines.some((l) => l.rateDefaulted)) {
        items.push({ message: `${b.employeeName} — no configured wage rate (default $ applied); add them in Settings` });
      }
    }
    groups.push({
      key: 'missingInputs',
      title: 'Missing pay inputs',
      note: 'Piece-rate jobs missing WO% / WO Override + Income, and employees with no configured wage rate.',
      items,
    });
  }

  // 5b (Round 4/6 §7.1). Davis-Bacon job paid below the DB minimum (manual edit).
  {
    const items: ExceptionItem[] = [];
    for (const b of allBlocks) {
      for (const l of b.lines) {
        if (l.belowPrevailing) {
          items.push({
            message: `${b.employeeName} — "${l.jobLabel}" is a Davis-Bacon job but pays ${l.rate.toFixed(2)}/hr (below the Davis-Bacon rate)`,
          });
        }
      }
    }
    groups.push({
      key: 'prevailingWage',
      title: 'Davis-Bacon job paid below DB rate',
      note: 'The job matched a Davis-Bacon pattern in Settings but the rendered rate sits below the Davis-Bacon minimum (usually after a manual edit).',
      items,
    });
  }

  // 5c (Round 6 §7.1 Branch 2). Davis-Bacon floor applied on crew-lead pay.
  {
    const items: ExceptionItem[] = [];
    for (const b of allLeads) {
      for (const j of b.jobs) {
        if (!j.dbFloorApplied) continue;
        const ps = (j.dbFloorPriorPay ?? j.totalAfterSplit).toFixed(2);
        const hrs = j.dbFloorHours ?? 0;
        const rate = j.dbFloor && hrs > 0 ? (j.dbFloor / hrs).toFixed(2) : '?';
        items.push({
          message: `Davis-Bacon floor applied: ${b.leadName} on ${j.apKey ? `AP${j.apKey}` : '—'} — profit share ${ps} < ${hrs}h × ${rate}`,
        });
      }
    }
    groups.push({
      key: 'dbFloorApplied',
      title: 'Davis-Bacon floor applied',
      note: 'Informational: crew-lead per-job pay was raised to hours × Davis-Bacon rate on a DB job.',
      items,
    });
  }

  // 5d (Round 6 §7.2). Job-rate rule applied (informational).
  {
    const items: ExceptionItem[] = [];
    for (const b of allBlocks) {
      for (const l of b.lines) {
        if (!l.fromJobRule) continue;
        const note = l.ruleNote ? ` (${l.ruleNote})` : '';
        items.push({
          message: `${b.employeeName} — "${l.jobLabel}" @ $${l.rate.toFixed(2)}/hr${note}`,
        });
      }
    }
    groups.push({
      key: 'jobRateRule',
      title: 'Job-rate rule applied',
      note: 'Informational: hourly lines paid from a Settings job-specific pay rule this week.',
      items,
    });
  }

  // 5e (Round 6 §6.2 / Round 8 §9.4). No work-order % for lead. With §9.4
  // auto-add in place this only fires when the auto-add was declined.
  {
    const items: ExceptionItem[] = [];
    const seen = new Set<string>();
    for (const b of allLeads) {
      if (b.jobs.some((j) => j.woPct === null) && !seen.has(b.leadName)) {
        seen.add(b.leadName);
        items.push({ message: `no work-order % for ${b.leadName}` });
      }
    }
    groups.push({
      key: 'noWorkOrderPct',
      title: 'No work-order %',
      note: 'Every band resolves its WO% from the pivot cell, Settings, or the §9.4 auto-add — this only fills up if the auto-add was declined.',
      items,
    });
  }

  // 5f (Round 8 §9.4). Crew lead auto-added to the profit-share map.
  {
    const seen = new Set<string>();
    const items: ExceptionItem[] = [];
    for (const p of previews) {
      for (const n of p.model.newLeadsToAdd) {
        const key = n.lead.trim().toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        items.push({ message: `new crew lead ${n.lead} added at default ${Math.round(n.pct * 100)}%` });
      }
    }
    groups.push({
      key: 'newCrewLeadDefault',
      title: 'New crew lead added at the default %',
      note: 'Informational: the lead resolved from the input data with no Work-order % on file, so the seeded default was applied and saved to Settings. Change it there if the crew negotiated a different split.',
      items,
    });
  }

  // 5g (Round 8 §8.1). SUBBED / SET RATE jobs pay nothing automatically.
  {
    const items: ExceptionItem[] = [];
    const missingPct: ExceptionItem[] = [];
    for (const b of allLeads) {
      for (const j of b.jobs) {
        const po = j.apKey ? `AP${j.apKey}` : j.jobShort;
        if (j.manualPay) {
          items.push({ message: `${po} is ${j.manualPay} — enter pay manually` });
        } else if (j.missingPercent) {
          missingPct.push({ message: `missing % completion for ${po} (${b.leadName})` });
        }
      }
    }
    groups.push({
      key: 'manualPayJob',
      title: 'Subbed / set-rate job — enter pay manually',
      note: 'There is no auto profit share on these jobs. The band cells are blank and stay editable for the hand-entered amount.',
      items,
    });
    groups.push({
      key: 'missingPercentCompletion',
      title: 'Missing % completion',
      note: 'The §8.1 chain multiplies the base by % Completion, so a blank or unreadable % pays 0 — a deliberately conservative choice. It never silently pays 100%; fill the % in and the band recomputes.',
      items: missingPct,
    });
  }

  // 5h (Round 8 §9.2). NV daily overtime applied.
  groups.push({
    key: 'nvDailyOt',
    title: 'NV daily overtime (over 8 hrs/day)',
    note: 'Informational: NV hours past 8 in a single day at a rate under $18 pay 1.5×. These hours are excluded from the weekly >40 split, so no hour is paid overtime twice.',
    items: previews.flatMap((p) =>
      p.model.dailyOtNotes.map((n) => ({
        message: `${n.employeeName} — ${n.date}: ${fmtHours(n.hours)} h over 8 @ $${n.rate.toFixed(2)} on "${n.jobLabel}"`,
      })),
    ),
  });

  // 6. Job missing from P&L. Each model is judged against ITS OWN upload, so a
  // week with no UT P&L never accuses every UT job of being missing — those
  // jobs simply keep the Job-Total-$ fallback (Round 8 §7.2).
  groups.push({
    key: 'missingFromPl',
    title: 'Job missing from P&L',
    note: 'Income/Labor default to 0 — the profit share cannot be computed until the P&L includes the job (or you edit the numbers).',
    items: previews.flatMap(({ model, pl: entityPl }) =>
      entityPl
        ? model.leads
            .flatMap((b) => b.jobs.filter((j) => j.missingFromPl).map((j) => ({ b, j })))
            .map(({ b, j }) => ({
              message: `AP ${j.apKey ?? '—'} (${b.leadName})${model.entity === 'UT' ? ' [UT]' : ''} — no P&L column`,
            }))
        : [],
    ),
  });

  // 6b (Round 8 §7.1). COGS rows the Settings mapping never mentions. UT's
  // `Reimbursements` shows up here by design — a new QBO account must never
  // silently drop out of materials/labor without anyone noticing.
  {
    const items: ExceptionItem[] = [];
    for (const source of [pl, input.plUt ?? null]) {
      if (!source) continue;
      for (const label of source.unmappedCogsRows) {
        items.push({ message: `unmapped P&L account row: ${label} (${source.entity})` });
      }
    }
    groups.push({
      key: 'unmappedPlAccount',
      title: 'Unmapped P&L account row',
      note: 'Informational: the P&L carries a Cost of Goods Sold row that no metric in Settings → P&L account mapping claims, so it feeds nothing. Map it there if it belongs in materials, labor, sundries, travel or mileage.',
      items,
    });
  }

  // 7. Profit share ≤ 0 or near-zero.
  {
    const threshold = input.psThreshold ?? PS_THRESHOLD_DEFAULT;
    groups.push({
      key: 'lowProfitShare',
      title: `Profit share below $${threshold}`,
      note: 'Often means the crew went way over on the work order — consider an Addition/adjustment.',
      items: allLeads.flatMap((b) =>
        b.jobs
          // Manual (SUBBED / SET RATE) bands have no computed share to judge.
          .filter((j) => j.totalProfitShare < threshold && !j.missingFromPl && !j.manualPay)
          .map((j) => ({
            message: `${b.leadName} — AP ${j.apKey ?? '—'} Total Profit Share ${j.totalProfitShare.toFixed(2)}`,
          })),
      ),
    });
  }

  // 8 (Round 3 §6-8, reworked by Round 8 §8.3). Sundry double-count hint. The
  // P&L sundries no longer sit inside materials — they ARE the default
  // Deductions/Sundries entry — so the only thing left to flag is a manual
  // edit that walked away from that default.
  {
    const items: ExceptionItem[] = [];
    for (const b of allLeads) {
      if (b.sundriesDefault === 0 && b.deductions === 0) continue;
      const entered = Math.round(b.deductions * 100) / 100;
      const expected = Math.round(b.sundriesDefault * 100) / 100;
      if (entered === expected) continue;
      items.push({
        message:
          expected === 0
            ? `${b.leadName} — Deductions/Sundries reads ${entered.toFixed(2)} but the P&L shows no sundries on these jobs — confirm it is not already inside materials`
            : `${b.leadName} — P&L sundries total ${expected.toFixed(2)} but Deductions/Sundries reads ${entered.toFixed(2)}; the difference is a manual edit, confirm it is not double-counted`,
      });
    }
    groups.push({
      key: 'sundryHint',
      title: 'Sundry double-count hint',
      note: 'Sundries default into Deductions/Sundries (memo "Sundries (P&L)") and are no longer part of ½ Materials. This lists blocks whose entry differs from that default.',
      items,
    });
  }

  // 9. Inactive employee has hours.
  groups.push({
    key: 'inactiveWithHours',
    title: 'Inactive employee has hours',
    items: [...allPainters, ...allOffice]
      .filter((b) => b.inactive && b.totalHours > 0)
      .map((b) => ({ message: `${b.employeeName} — marked inactive (*) in the directory but has ${fmtHours(b.totalHours)} h` })),
  });

  // 10. Duplicate P&L job column.
  groups.push({
    key: 'duplicatePlColumn',
    title: 'Duplicate P&L job column',
    note: 'The same AP number appeared in more than one P&L column — values were summed.',
    items: [pl, input.plUt ?? null].flatMap((source) =>
      (source?.duplicateColumns ?? []).map((ap) => ({
        message: `AP ${ap} appears in multiple P&L columns${source && source.entity === 'UT' ? ' [UT]' : ''}`,
      })),
    ),
  });

  // 11. Balances / UT carry-forward attention.
  {
    const items: ExceptionItem[] = [];
    const thisWeekAps = new Set(included.map((r) => r.apKey).filter(Boolean));
    for (const row of [...pbj.balances, ...pbj.utPayroll]) {
      if (typeof row.percent === 'number' && row.percent >= 1) {
        items.push({ message: `${row.job || row.apKey} — carried forward but % of Completion is already 100% — likely fully paid; delete the row if so` });
      } else if (row.apKey && !thisWeekAps.has(row.apKey) && row.id.startsWith('ut:')) {
        items.push({ message: `${row.job || row.apKey} — UT carry-forward matched nothing this week` });
      }
    }
    groups.push({
      key: 'carryForward',
      title: 'Balances / UT carry-forward attention',
      note: 'Rows carried from last week that look fully paid or matched nothing this week.',
      items,
    });
  }

  // 12 (Round 5 §4). Naming-chain informational flags.
  {
    const baseKeyItems: ExceptionItem[] = [];
    const unresolvedItems: ExceptionItem[] = [];
    const sources = {
      timesheetByAp: timesheetDisplayMap(input.enriched ?? []),
      carryByAp: carryDisplayMap(input.lastWeekPbj),
      salesByAp: salesDisplayMap(input.nvSales, input.utSales),
      pl,
    };
    const seen = new Set<string>();
    const consider = (apKey: string | null | undefined, eventTitle?: string): void => {
      if (!apKey || apKey === 'OFFICE' || seen.has(apKey)) return;
      seen.add(apKey);
      const pair = jobDisplayPair(apKey, { ...sources, eventTitle });
      if (pair.usedBaseKey) {
        baseKeyItems.push({
          message: `AP ${apKey} — suffix not found in QBO; names derived from base AP# (${pair.poLabel} / ${pair.customerName || '—'})`,
        });
      }
      if (pair.unresolved) {
        unresolvedItems.push({ message: `AP ${apKey} — job name unresolved` });
      }
    };
    for (const { job } of jobs) consider(job.analysis.apKey);
    for (const r of [...pbj.balances, ...pbj.utPayroll]) consider(r.apKey);
    for (const r of included) {
      if (r.apKey && !seen.has(r.apKey)) consider(r.apKey, r.event);
    }
    for (const lead of allLeads) {
      for (const j of lead.jobs) consider(j.apKey);
    }
    groups.push({
      key: 'apSuffixBaseKey',
      title: 'AP# suffix not found in QBO; names derived from base AP#',
      note: 'A suffixed calendar AP# missed the sales lookup; display names came from the base number. Edit the AP# cell if the suffix is wrong.',
      items: baseKeyItems,
    });
    groups.push({
      key: 'jobNameUnresolved',
      title: 'Job name unresolved',
      note: 'No timesheet, carry-forward, QBO, or P&L name was found — the PO shows AP{key} with a blank customer.',
      items: unresolvedItems,
    });
  }

  // Round 8 §6.4 — a carried (no-timesheet) job with no resolvable crew lead
  // stays pivot-only and is called out here.
  {
    const items: ExceptionItem[] = [];
    for (const row of [...pbj.balances, ...pbj.utPayroll]) {
      if (row.excludedFromPayroll || row.apKey === 'OFFICE') continue;
      if (row.percent === 'HOURLY') continue;
      const lead = (row.crewLead || '').trim();
      if (lead !== '') continue;
      items.push({
        message: `carried job ${row.apKey ? `AP${row.apKey}` : row.job || row.id} has no crew lead — not added to payroll`,
      });
    }
    groups.push({
      key: 'carriedNoLead',
      title: 'Carried job has no crew lead',
      note: 'The job sits below Grand Total on Pivot by Job but has no Crew Lead, so it was not added to Input #\'s or a preview band.',
      items,
    });
  }

  return groups;
}
