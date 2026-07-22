"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.baseApKey = baseApKey;
exports.timesheetDisplayMap = timesheetDisplayMap;
exports.carryDisplayMap = carryDisplayMap;
exports.salesDisplayMap = salesDisplayMap;
exports.cleanEventTitle = cleanEventTitle;
exports.jobDisplayPair = jobDisplayPair;
/**
 * Round 5 §4 — shared job/customer display-name chain.
 * Every renderer of a PO/Name or customer/job pair consumes this; the raw
 * TimeNavi event title must never appear as a display name.
 */
const apkey_1 = require("./apkey");
/** Numeric base of a suffixed AP key: "24180-A" → "24180"; bare keys → null. */
function baseApKey(apKey) {
    const m = /^(\d{4,6})-(.+)$/i.exec(apKey.trim());
    return m ? m[1] : null;
}
/** Build the timesheet apKey → {customer=jobcode_1, job=jobcode_2} map. */
function timesheetDisplayMap(rows) {
    const m = new Map();
    for (const r of rows) {
        if (!r.apKey || r.apKey === 'OFFICE' || m.has(r.apKey))
            continue;
        m.set(r.apKey, { customer: r.jobcode1, job: r.jobcode2 });
    }
    return m;
}
/** Slot-7 carry-forward pairs keyed by apKey (first occurrence wins). */
function carryDisplayMap(rows) {
    const m = new Map();
    for (const r of rows ?? []) {
        if (!r.apKey || m.has(r.apKey))
            continue;
        m.set(r.apKey, { customer: r.customer, job: r.job });
    }
    return m;
}
/**
 * Sales lookup: NV rows first, then UT — mirrors the TimeNavi QBO Job Name
 * column. First name1/name2 for each apKey wins.
 */
function salesDisplayMap(nvSales, utSales) {
    const m = new Map();
    for (const list of [nvSales, utSales]) {
        if (!list)
            continue;
        for (const r of list) {
            if (!r.apKey || m.has(r.apKey))
                continue;
            if (!r.name1 && !r.name2)
                continue;
            m.set(r.apKey, { name1: r.name1, name2: r.name2 });
        }
    }
    return m;
}
/**
 * Round 5 §4 step 6 — parse a calendar title into display names without ever
 * dumping the verbatim string. Returns null when nothing usable remains.
 */
function cleanEventTitle(title, apKey) {
    let t = title.trim();
    if (!t)
        return null;
    // Strip leading decoration glyphs (♦, •, etc.).
    t = t.replace(/^[\u2666\u25C6\u2022\u25CF\u25A0♦●•]+\s*/u, '');
    // Drop *[STATUS]* tokens like *PAID*.
    t = t.replace(/\*[A-Za-z][A-Za-z0-9 _-]*\*/g, ' ');
    // Drop $ / k amount tokens: "13k", "3.5k", "$8k", "13,000".
    t = t.replace(/\$?\d+(?:[.,]\d+)?k\b/gi, ' ');
    t = t.replace(/\$[\d,]+(?:\.\d+)?/g, ' ');
    // Collect parentheticals; drop known-estimator (1–3 uppercase letters) and
    // the trailing crew parenthetical (last one that is not an estimator).
    const parenRe = /\(([^)]*)\)/g;
    const parens = [];
    let pm;
    while ((pm = parenRe.exec(t)) !== null) {
        parens.push({ raw: pm[0], inner: pm[1].trim(), start: pm.index, end: pm.index + pm[0].length });
    }
    const isEstimatorParen = (inner) => /^[A-Za-z]{1,3}$/.test(inner.trim());
    // Strip all estimator parentheticals.
    for (const p of [...parens].reverse()) {
        if (isEstimatorParen(p.inner)) {
            t = t.slice(0, p.start) + ' ' + t.slice(p.end);
        }
    }
    // Re-scan; drop the last remaining parenthetical as the crew lead.
    const remaining = [];
    parenRe.lastIndex = 0;
    while ((pm = parenRe.exec(t)) !== null) {
        remaining.push({ start: pm.index, end: pm.index + pm[0].length, inner: pm[1].trim() });
    }
    if (remaining.length > 0) {
        const last = remaining[remaining.length - 1];
        t = t.slice(0, last.start) + ' ' + t.slice(last.end);
    }
    // Bracketed text → preferred customer name.
    let bracketCustomer = '';
    const br = /\[([^\]]+)\]/.exec(t);
    if (br) {
        bracketCustomer = br[1].trim();
        t = t.replace(br[0], ' ');
    }
    // Pull the AP token (and anything after a comma that introduces it).
    const apMatch = /,\s*AP\s*-?\s*\d{4,6}(?:-[A-Za-z0-9]+)*/i.exec(t);
    let beforeAp = t;
    if (apMatch && apMatch.index !== undefined) {
        beforeAp = t.slice(0, apMatch.index);
    }
    else {
        // No comma-AP form — strip a bare AP token in place.
        beforeAp = t.replace(/AP\s*-?\s*\d{4,6}(?:-[A-Za-z0-9]+)*/gi, ' ');
    }
    // Short type tokens for the PO label.
    const typeTok = /\b(Int\/Ext|Int|Ext|T\s*&\s*M)\b/i.exec(title);
    const typePart = typeTok ? typeTok[1].replace(/\s+/g, '') : '';
    const key = apKey ?? (0, apkey_1.extractApKey)(title);
    const poLabel = key ? `AP${key}${typePart ? ` ${typePart}` : ''}` : '';
    let customerName = bracketCustomer;
    if (!customerName) {
        customerName = beforeAp
            .replace(/\b(Int\/Ext|Int|Ext|T\s*&\s*M)\b/gi, ' ')
            .replace(/\bPh\s*\d+\b/gi, ' ')
            .replace(/\bPart\b[\s\w-]*/gi, ' ')
            .replace(/[^\w\s'/.-]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
        // Drop leading connector junk and very short leftovers.
        customerName = customerName.replace(/^(and|or|for|the)\s+/i, '').trim();
    }
    if (!poLabel && !customerName)
        return null;
    return { poLabel: poLabel || (key ? `AP${key}` : ''), customerName };
}
/**
 * §4 chain — first hit wins per field. `customerName` is never a copy of the
 * job string; unresolved jobs get `AP{key}` / blank + the caller's exception.
 */
function jobDisplayPair(apKey, sources) {
    const empty = { poLabel: '', customerName: '', usedBaseKey: false, unresolved: true };
    if (!apKey || apKey === 'OFFICE') {
        return apKey === 'OFFICE'
            ? { poLabel: 'OFFICE', customerName: 'ALLPRO OFFICE', usedBaseKey: false, unresolved: false }
            : empty;
    }
    const hit = (po, cust, usedBaseKey = false) => ({
        poLabel: po.trim(),
        customerName: cust.trim(),
        usedBaseKey,
        unresolved: false,
    });
    // 1. Timesheet
    const ts = sources.timesheetByAp.get(apKey);
    if (ts && (ts.job || ts.customer))
        return hit(ts.job || `AP${apKey}`, ts.customer);
    // 2. Slot-7 carry-forward
    const carry = sources.carryByAp.get(apKey);
    if (carry && (carry.job || carry.customer))
        return hit(carry.job || `AP${apKey}`, carry.customer);
    // 3. QBO sales exact
    const sales = sources.salesByAp.get(apKey);
    if (sales && (sales.name2 || sales.name1))
        return hit(sales.name2 || `AP${apKey}`, sales.name1);
    // 4. Base-key retry on QBO
    const base = baseApKey(apKey);
    if (base) {
        const baseSales = sources.salesByAp.get(base);
        if (baseSales && (baseSales.name2 || baseSales.name1)) {
            return hit(baseSales.name2 || `AP${base}`, baseSales.name1, true);
        }
    }
    // 5. P&L
    const pl = sources.pl?.byJob.get(apKey) ?? (base ? sources.pl?.byJob.get(base) : undefined);
    if (pl?.plJobHeader) {
        return hit(pl.plJobHeader, pl.customerName ?? '', base !== null && !sources.pl?.byJob.has(apKey));
    }
    // 6. Cleaned event title
    if (sources.eventTitle) {
        const cleaned = cleanEventTitle(sources.eventTitle, apKey);
        if (cleaned && (cleaned.poLabel || cleaned.customerName)) {
            return hit(cleaned.poLabel || `AP${apKey}`, cleaned.customerName);
        }
    }
    // 7. Nothing
    return { poLabel: `AP${apKey}`, customerName: '', usedBaseKey: false, unresolved: true };
}
