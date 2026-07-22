"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildSalesByRepMirrorFromSheet = buildSalesByRepMirrorFromSheet;
exports.buildSalesByRepMirrorFromParsed = buildSalesByRepMirrorFromParsed;
exports.buildSalesByCustMirror = buildSalesByCustMirror;
exports.buildTimesheetMirror = buildTimesheetMirror;
/**
 * Round 6 §3.1–§3.5 — input mirror models for the %'s workbook helper tabs.
 * Read-only mirrors of the uploads + the same helper columns the human keeps.
 */
const apkey_1 = require("./apkey");
function sheetCols(row, n) {
    const out = [];
    for (let i = 0; i < n; i++)
        out.push(row[i] ?? null);
    return out;
}
/** §3.1 / §3.2 — Sales by Rep mirrors from already-parsed SalesRow aren't enough
 *  for full column fidelity; build from the raw sheet when available. */
function buildSalesByRepMirrorFromSheet(sheet, kind) {
    if (!sheet || sheet.rows.length === 0)
        return { headers: [], rows: [] };
    // Find header row containing Date + Sales Rep-ish columns.
    let headerIdx = -1;
    for (let i = 0; i < Math.min(sheet.rows.length, 30); i++) {
        const joined = sheet.rows[i].map((c) => String(c ?? '').toLowerCase()).join('|');
        if (joined.includes('date') && (joined.includes('sales rep') || joined.includes('amount'))) {
            headerIdx = i;
            break;
        }
    }
    if (headerIdx < 0)
        headerIdx = 0;
    const headerRow = sheet.rows[headerIdx].map((c) => String(c ?? '').trim());
    const name2Idx = headerRow.findIndex((h) => /^name\s*2$/i.test(h));
    const customerIdx = headerRow.findIndex((h) => /^customer$/i.test(h) || /^name$/i.test(h));
    const dateIdx = headerRow.findIndex((h) => /^date$/i.test(h));
    const colCount = Math.min(12, Math.max(headerRow.length, 12));
    const headers = Array.from({ length: colCount }, (_, i) => headerRow[i] || `Col ${i + 1}`);
    const rows = [];
    for (let r = headerIdx + 1; r < sheet.rows.length; r++) {
        const raw = sheet.rows[r];
        if (!raw || raw.every((c) => c === null || c === ''))
            continue;
        const cols = sheetCols(raw, colCount);
        const name2 = name2Idx >= 0 ? String(cols[name2Idx] ?? '') : '';
        const customer = customerIdx >= 0 ? String(cols[customerIdx] ?? '') : '';
        const apKey = (0, apkey_1.extractApKey)(name2) ?? (0, apkey_1.extractApKey)(customer) ?? '';
        rows.push({
            apKey,
            cols,
            dateCopy: dateIdx >= 0 ? cols[dateIdx] : null,
            qboJobName: name2Idx >= 0 ? cols[name2Idx] : null,
        });
    }
    void kind;
    return { headers, rows };
}
/** Fallback from parsed SalesRow when raw sheet isn't kept. */
function buildSalesByRepMirrorFromParsed(sales) {
    const headers = [
        'Date',
        'Customer',
        'Name 1',
        'Name 2',
        'Transaction Type',
        'Num',
        'Product/Service',
        'Memo/Description',
        'Sales Price',
        'Amount',
        'Balance',
        'Sales Rep',
    ];
    const rows = sales.map((s) => {
        const cols = [
            s.date,
            s.name1,
            s.name1,
            s.name2,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            s.salesRep,
        ];
        return { apKey: s.apKey ?? '', cols, dateCopy: s.date, qboJobName: s.name2 };
    });
    return { headers, rows };
}
function buildSalesByCustMirror(sheet) {
    if (!sheet)
        return [];
    const out = [];
    for (const row of sheet.rows) {
        const text = String(row[0] ?? '');
        if (text === '' && (row[1] === null || row[1] === undefined || row[1] === ''))
            continue;
        let value = null;
        const v = row[1];
        if (typeof v === 'number')
            value = v;
        else if (typeof v === 'string') {
            const m = /^-?\d+(?:\.\d+)?/.exec(v.replace(/[,$=()]/g, '').replace(/^\(/, '-'));
            if (m)
                value = Number(m[0]);
            else {
                const n = Number(String(v).replace(/[,$]/g, '').replace(/^=/, ''));
                if (Number.isFinite(n))
                    value = n;
            }
        }
        const apKey = (0, apkey_1.extractApKey)(text) ?? '';
        out.push({
            text,
            value,
            apKey,
            jobTotal: apKey && value !== null ? value : null,
        });
    }
    return out;
}
/**
 * §3.5 timesheets mirror helpers. `timeNaviRowByAp` maps apKey → 1-based sheet
 * row number on TimeNavi Export (header = 1).
 */
function buildTimesheetMirror(enriched, timeNaviRowByAp) {
    return enriched.map((r) => {
        const apKey = r.apKey === 'OFFICE' || /^office$/i.test(r.jobcode2.trim())
            ? 'OFFICE'
            : (r.apKey ?? (0, apkey_1.timesheetApKey)(r.jobcode2) ?? '');
        const apname = r.jobcode2.trim().split(/\s+/)[0] ?? '';
        const cal = apKey && apKey !== 'OFFICE' && timeNaviRowByAp.has(apKey)
            ? String(timeNaviRowByAp.get(apKey))
            : 'NotOnCalendar';
        const raw = [
            r.username,
            r.payrollId,
            r.fname,
            r.lname,
            r.number,
            r.group,
            r.localDate,
            r.localDay,
            r.localStartTime,
            r.localEndTime,
            r.tz,
            r.hours,
            r.jobcode1,
            r.jobcode2,
            r.billable,
            r.klass,
            r.department,
            r.payrollItem,
            r.serviceItem,
            r.location,
            r.notes,
            r.approvedStatus,
            r.hasFlags,
            r.flagTypes,
        ];
        return {
            raw,
            apname,
            apKey,
            estimator: r.estimator,
            qboJobName: r.jobcode2,
            notOnCalendar: cal,
        };
    });
}
