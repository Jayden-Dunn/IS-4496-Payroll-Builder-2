"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildPrFinalCheckRows = buildPrFinalCheckRows;
/** One row per job in `pl.byJob` (Total columns already excluded by parsePl). */
function buildPrFinalCheckRows(pl) {
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
