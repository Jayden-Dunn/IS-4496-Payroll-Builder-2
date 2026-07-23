/**
 * Round 7 §2 — PR Final Check rows from P&L by-job metrics.
 */
import type { PlParseResult } from './types';

export interface PrFinalCheckRow {
  apKey: string;
  jobHeader: string;
  customerName: string;
  income: number;
  materials: number;
  laborWagesSubs: number;
  mileage: number;
  sundries: number;
}

/** One row per job in `pl.byJob` (Total columns already excluded by parsePl). */
export function buildPrFinalCheckRows(pl: PlParseResult): PrFinalCheckRow[] {
  return [...pl.byJob.entries()].map(([apKey, m]) => ({
    apKey,
    jobHeader: m.plJobHeader,
    customerName: m.customerName,
    income: m.income,
    materials: m.materials,
    laborWagesSubs: m.laborWagesSubs,
    mileage: m.mileage,
    sundries: m.sundries,
  }));
}
