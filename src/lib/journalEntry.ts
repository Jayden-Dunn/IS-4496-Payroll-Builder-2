/**
 * Journal Entry computation (Round 7 §18 — export sheet is a header shell only).
 */
import type { PayrollPreviewModel } from './types';

export interface JournalEntryRow {
  account: string;
  amount: number;
  memo: string;
  po: string;
  className: string;
  job: string;
}

export function buildJournalEntryRows(model: PayrollPreviewModel | null): {
  rows: JournalEntryRow[];
  amountSum: number;
} {
  interface JobAgg {
    apKey: string;
    jobShort: string;
    customer: string;
    amount: number;
    parts: { lead: string; amt: number; woPct: number | null; split: number }[];
  }
  const byAp = new Map<string, JobAgg>();
  for (const lead of model?.leads ?? []) {
    for (const j of lead.jobs) {
      if (!j.apKey) continue;
      let agg = byAp.get(j.apKey);
      if (!agg) {
        agg = { apKey: j.apKey, jobShort: j.jobShort, customer: j.customer, amount: 0, parts: [] };
        byAp.set(j.apKey, agg);
      }
      const amt = j.totalAfterSplit;
      agg.amount += amt;
      agg.parts.push({
        lead: lead.displayName,
        amt,
        woPct: j.woPct,
        split: j.splitWith,
      });
    }
  }
  const sorted = [...byAp.values()].sort((a, b) =>
    a.jobShort.localeCompare(b.jobShort, 'en', { sensitivity: 'base' }),
  );
  const rows: JournalEntryRow[] = [];
  let amountSum = 0;
  for (const j of sorted) {
    const pctParts = j.parts.map((p) => `${p.lead} (${p.amt.toFixed(2)})`).join(', ');
    const firstPct = j.parts[0]?.woPct;
    const pctLabel = firstPct !== null && firstPct !== undefined ? `${Math.round(firstPct * 100)}%` : '';
    const splitNote = j.parts.some((p) => p.split > 1)
      ? `*split x${Math.max(...j.parts.map((p) => p.split))}*`
      : '';
    const amt = Math.round(j.amount * 100) / 100;
    rows.push({
      account: 'Employee Labor',
      amount: amt,
      memo: `${j.jobShort} [${pctLabel}${splitNote}]:  ${pctParts}`,
      po: `AP${j.apKey}`,
      className: 'Painting',
      job: `${j.customer} ${j.jobShort}`.trim(),
    });
    amountSum += amt;
  }
  for (const lead of model?.leads ?? []) {
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
  for (const b of [...(model?.admin ?? []), ...(model?.estimators ?? []), ...(model?.managers ?? [])]) {
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
  return { rows, amountSum: Math.round(amountSum * 100) / 100 };
}
