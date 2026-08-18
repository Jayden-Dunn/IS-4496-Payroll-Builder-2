"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.customerByApFromSalesSummary = customerByApFromSalesSummary;
exports.customerByApFromPl = customerByApFromPl;
exports.buildCustomerByApResolver = buildCustomerByApResolver;
/**
 * Round 8 §4.3 — resolve a job's CUSTOMER from its AP#.
 *
 * The Sales by Customer Summary is hierarchical: a non-indented customer row
 * ("Aaron Ward") followed by indented job rows ("   AP24314 Ext"). The owning
 * customer of a job is the nearest non-indented row above it. The P&L header
 * has the same shape sideways (a job column belongs to the nearest non-`Total`
 * customer header to its left) and serves as the secondary source.
 *
 * Rows come from the already-parsed `mirrorTabs` models — no second parser.
 */
const apkey_1 = require("./apkey");
const jobNames_1 = require("./jobNames");
const UNRESOLVED = { customer: '', source: null };
/** Summary rows indent job lines; customer headers start at column A, flush left. */
function isIndented(text) {
    return /^[\s\u00a0]/.test(text);
}
function isTotalRow(text) {
    return /^\s*total\b/i.test(text);
}
/**
 * apKey → owning customer for one Sales by Customer Summary. First occurrence
 * wins so a repeated AP# behaves like every other first-match lookup here.
 */
function customerByApFromSalesSummary(rows) {
    const map = new Map();
    if (!rows)
        return map;
    let currentCustomer = '';
    for (const row of rows) {
        const text = row.text ?? '';
        if (!text.trim())
            continue;
        if (isTotalRow(text))
            continue;
        const apKey = row.apKey || (0, apkey_1.extractApKey)(text);
        if (apKey) {
            if (currentCustomer && !map.has(apKey))
                map.set(apKey, currentCustomer);
            continue;
        }
        // A flush-left, AP-less row is the customer header the job rows hang under.
        if (!isIndented(text))
            currentCustomer = text.trim();
    }
    return map;
}
/** apKey → customer header from the P&L column structure (already parsed). */
function customerByApFromPl(pl) {
    const map = new Map();
    if (!pl)
        return map;
    for (const [apKey, metrics] of pl.byJob) {
        const name = metrics.customerName?.trim();
        if (name && !map.has(apKey))
            map.set(apKey, name);
    }
    return map;
}
function buildCustomerByApResolver(sources) {
    const nv = customerByApFromSalesSummary(sources.nvSalesByCust);
    const ut = customerByApFromSalesSummary(sources.utSalesByCust);
    const pl = customerByApFromPl(sources.pl);
    const lookup = (map, apKey) => {
        const exact = map.get(apKey);
        if (exact)
            return exact;
        // Change orders ("24097-A") bill under the base job's customer.
        const base = (0, jobNames_1.baseApKey)(apKey);
        return (base ? map.get(base) : undefined) ?? '';
    };
    return {
        resolve(apKey, entity) {
            if (!apKey || apKey === 'OFFICE')
                return UNRESOLVED;
            const order = entity === 'UT'
                ? [
                    { map: ut, source: 'salesUt' },
                    { map: nv, source: 'salesNv' },
                ]
                : [
                    { map: nv, source: 'salesNv' },
                    { map: ut, source: 'salesUt' },
                ];
            for (const step of order) {
                const hit = lookup(step.map, apKey);
                if (hit)
                    return { customer: hit, source: step.source };
            }
            const fromPl = lookup(pl, apKey);
            if (fromPl)
                return { customer: fromPl, source: 'pl' };
            return UNRESOLVED;
        },
    };
}
