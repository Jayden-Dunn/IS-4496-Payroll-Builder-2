"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildSalesByRepMirrorSheet = buildSalesByRepMirrorSheet;
exports.buildSalesByCustMirrorSheet = buildSalesByCustMirrorSheet;
exports.buildTimesheetMirrorSheet = buildTimesheetMirrorSheet;
exports.buildJobsByCrewSheet = buildJobsByCrewSheet;
exports.buildPlImportSheet = buildPlImportSheet;
exports.buildProfitshareOutputSheet = buildProfitshareOutputSheet;
exports.buildManagementReportSheet = buildManagementReportSheet;
exports.buildInputNumbersSheet = buildInputNumbersSheet;
exports.buildNotesSheet = buildNotesSheet;
exports.buildFinalPayrollSheet = buildFinalPayrollSheet;
exports.buildJournalEntrySheet = buildJournalEntrySheet;
exports.buildJobsSheet = buildJobsSheet;
exports.buildInactiveEmployeesSheet = buildInactiveEmployeesSheet;
exports.buildPrFinalCheckSheet = buildPrFinalCheckSheet;
const apkey_1 = require("./apkey");
const format_1 = require("./format");
const prFinalCheck_1 = require("./prFinalCheck");
const BLACK = 'FF000000';
const WHITE = 'FFFFFFFF';
const HEADER_GRAY = 'FF808080';
const TAN = 'FFDDD9C4';
const HOURLY_PAY_YELLOW = 'FFFFFF99';
const TOTAL_PAY_CREAM = 'FFFAF5F0';
const NAME_BAND_RED = 'FFC0504D';
const NAME_BAND_GREEN = 'FF00B050';
const REIMB_BLUE = 'FF0070C0';
const DEDUCT_RED = 'FFFF0000';
const FONT = 'Calibri';
const PP_FONT = 'Garamond';
const ACCOUNTING_FMT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const MONEY_FMT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
const PCT_FMT = '0%';
const HOURS_FMT = '0.00';
const DATE_FMT = 'm/d/yyyy';
function solidFill(argb) {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}
function textOn(hex6) {
    return (0, format_1.isDarkColor)(hex6) ? WHITE : BLACK;
}
function excelDate(d) {
    return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
}
function cellValue(v) {
    if (v === null || v === undefined)
        return '';
    if (v instanceof Date)
        return excelDate(v);
    return v;
}
function boldHeader(ws, headers, fontName = FONT) {
    const row = ws.addRow(headers);
    for (let c = 1; c <= headers.length; c++) {
        const cell = row.getCell(c);
        cell.font = { name: fontName, bold: true, size: 11 };
    }
}
const TIMESHEET_HEADERS = [
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
// ---------------------------------------------------------------------------
// %'s mirrors
// ---------------------------------------------------------------------------
const MIRROR_ARIAL = 'Arial';
const WIDE_TEXT_HEADERS = /^(customer|name(\s*[12])?|memo(\/description)?|memo)$/i;
const PRODUCT_SERVICE_HEADER = /^product\s*\/\s*service$/i;
function salesByRepColWidth(header, colIndex0) {
    if (colIndex0 === 0)
        return 10; // AP#
    if (WIDE_TEXT_HEADERS.test(header.trim()))
        return 105.6;
    if (PRODUCT_SERVICE_HEADER.test(header.trim()))
        return 41.9;
    return 14;
}
function buildSalesByRepMirrorSheet(wb, name, data) {
    const isUt = /^ut\s*salesbyrep$/i.test(name);
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    const headers = data?.headers?.length
        ? ['AP#', ...data.headers.slice(0, 12), 'Date', 'QBO Job Name']
        : ['AP#', 'Date', 'Customer', 'Name 1', 'Name 2', 'Transaction Type', 'Num', 'Product/Service', 'Memo/Description', 'Sales Price', 'Amount', 'Balance', 'Sales Rep', 'Date', 'QBO Job Name'];
    ws.columns = headers.map((h, i) => ({ width: salesByRepColWidth(h, i) }));
    const headerRow = ws.addRow(headers);
    for (let c = 1; c <= headers.length; c++) {
        headerRow.getCell(c).font = { name: MIRROR_ARIAL, bold: true, size: isUt ? 9 : 8 };
    }
    for (const r of data?.rows ?? []) {
        const cols = r.cols.slice(0, 12);
        while (cols.length < 12)
            cols.push(null);
        const row = ws.addRow([r.apKey, ...cols.map(cellValue), cellValue(r.dateCopy), cellValue(r.qboJobName)]);
        for (let c = 1; c <= headers.length; c++) {
            row.getCell(c).font = { name: MIRROR_ARIAL, size: 8 };
            const v = row.getCell(c).value;
            if (v instanceof Date)
                row.getCell(c).numFmt = DATE_FMT;
            else if (typeof v === 'number')
                row.getCell(c).numFmt = HOURS_FMT;
        }
    }
}
function looksLikeCustTitleRow(text, index) {
    const t = text.trim();
    if (!t)
        return false;
    if (index === 0 && /allpro/i.test(t))
        return true;
    if (/sales\s+by\s+customer/i.test(t))
        return true;
    // Date-range title line (e.g. "June 1 - June 7, 2026")
    if (index <= 2 && /\d/.test(t) && (/\s-\s/.test(t) || /\b20\d{2}\b/.test(t)))
        return true;
    return false;
}
function buildSalesByCustMirrorSheet(wb, name, rows) {
    const isUt = /^ut\s*salesbycust$/i.test(name);
    const list = rows ?? [];
    const ws = wb.addWorksheet(name, {
        views: [{ state: 'frozen', ySplit: isUt ? 5 : 1 }],
    });
    ws.columns = [{ width: 48 }, { width: 14 }, { width: 10 }, { width: 12 }];
    let dataRows = list;
    let titleCount = 0;
    while (titleCount < 3 && titleCount < list.length && looksLikeCustTitleRow(list[titleCount].text, titleCount)) {
        titleCount++;
    }
    if (isUt && titleCount > 0) {
        // Title block A1–A3; money header "Total" on B5 (not B1).
        for (let t = 0; t < titleCount; t++) {
            const size = t === 2 ? 10 : 14;
            const row = ws.addRow([list[t].text, '']);
            for (let c = 1; c <= 2; c++) {
                row.getCell(c).font = { name: MIRROR_ARIAL, bold: true, size };
                row.getCell(c).alignment = { horizontal: 'center', vertical: 'middle' };
            }
            ws.mergeCells(row.number, 1, row.number, 2);
        }
        // Pad blank rows so the Total header lands on row 5; B1 stays clear of "Total".
        while (ws.rowCount < 4)
            ws.addRow([]);
        const hdr = ws.addRow(['', 'Total', 'AP#', 'Job total']);
        for (let c = 1; c <= 4; c++) {
            hdr.getCell(c).font = { name: MIRROR_ARIAL, bold: true, size: 8 };
        }
        dataRows = list.slice(titleCount);
    }
    else {
        // NV: synthetic header; skip title rows so first customer is on row 2.
        const hdr = ws.addRow(['', 'Total', 'AP#', 'Job total']);
        for (let c = 1; c <= 4; c++) {
            hdr.getCell(c).font = { name: MIRROR_ARIAL, bold: true, size: 8 };
        }
        dataRows = list.slice(titleCount);
    }
    // Drop a leading blank data row after the header (row 2 must be first customer).
    while (dataRows.length > 0 && !dataRows[0].text.trim() && dataRows[0].value === null && !dataRows[0].apKey) {
        dataRows = dataRows.slice(1);
    }
    for (const r of dataRows) {
        const row = ws.addRow([r.text, r.value ?? '', r.apKey, r.jobTotal ?? '']);
        for (let c = 1; c <= 4; c++) {
            row.getCell(c).font = { name: MIRROR_ARIAL, size: 8, bold: c === 1 };
        }
        const aText = String(r.text ?? '');
        if (/^total/i.test(aText.trim())) {
            const thinTop = { top: { style: 'thin' } };
            row.getCell(1).border = { ...row.getCell(1).border, ...thinTop };
            row.getCell(2).border = { ...row.getCell(2).border, ...thinTop };
            row.getCell(2).font = { name: MIRROR_ARIAL, size: 8, bold: true };
            if (typeof row.getCell(2).value === 'number') {
                row.getCell(2).numFmt = '"$"* #,##0.00';
            }
        }
        else if (typeof row.getCell(2).value === 'number') {
            row.getCell(2).numFmt = HOURS_FMT;
        }
        if (typeof row.getCell(4).value === 'number')
            row.getCell(4).numFmt = HOURS_FMT;
    }
}
function buildTimesheetMirrorSheet(wb, rows) {
    const ws = wb.addWorksheet('timesheets', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = TIMESHEET_HEADERS.map((h) => ({
        width: h === 'jobcode_2' || h === 'QBO Job Name' ? 36 : h === 'hours' ? 8 : 12,
    }));
    boldHeader(ws, TIMESHEET_HEADERS);
    for (const r of rows ?? []) {
        const raw = r.raw.slice(0, 24);
        while (raw.length < 24)
            raw.push(null);
        const vals = [
            ...raw.map(cellValue),
            r.apname,
            r.apKey,
            r.estimator,
            r.qboJobName,
            r.notOnCalendar,
        ];
        const row = ws.addRow(vals);
        for (let c = 1; c <= TIMESHEET_HEADERS.length; c++) {
            row.getCell(c).font = { name: FONT, size: 11 };
            const v = row.getCell(c).value;
            if (v instanceof Date)
                row.getCell(c).numFmt = DATE_FMT;
            if (c === 12 && typeof v === 'number')
                row.getCell(c).numFmt = HOURS_FMT;
        }
    }
}
function buildJobsByCrewSheet(wb, model) {
    const ws = wb.addWorksheet('Jobs by Crew');
    // Round 7 §9 — reference column widths (A spacer + B–E data).
    ws.columns = [{ width: 3 }, { width: 25.2 }, { width: 26.2 }, { width: 31.2 }, { width: 12.2 }];
    const bodyFont = { name: FONT, size: 12 };
    const OFFICE_YELLOW = 'FFFFFF99';
    const GRAND_LABEL_GRAY = 'FFBFBFBF';
    // Anchored at B2 per §3.6 — plain Calibri 12, not bold (Round 7 §9).
    ws.getCell('B2').value = 'Row Labels';
    ws.getCell('C2').value = 'jobcode_1';
    ws.getCell('D2').value = 'jobcode_2';
    ws.getCell('E2').value = 'Sum of hours';
    for (const col of ['B', 'C', 'D', 'E']) {
        ws.getCell(`${col}2`).font = { ...bodyFont };
    }
    let excelRow = 3;
    for (const g of model?.groups ?? []) {
        g.jobs.forEach((j, i) => {
            ws.getCell(`B${excelRow}`).value = i === 0 ? g.crew : '';
            ws.getCell(`C${excelRow}`).value = j.showCustomer ? j.customer : '';
            ws.getCell(`D${excelRow}`).value = j.job;
            ws.getCell(`E${excelRow}`).value = j.hours;
            ws.getCell(`E${excelRow}`).numFmt = HOURS_FMT;
            for (const col of ['B', 'C', 'D', 'E']) {
                ws.getCell(`${col}${excelRow}`).font = { ...bodyFont };
            }
            // ALLPRO OFFICE jobcode_1 cells: pale yellow + thin top/bottom border.
            const custCell = ws.getCell(`C${excelRow}`);
            if (typeof custCell.value === 'string' && custCell.value.trim().toLowerCase() === 'allpro office') {
                custCell.fill = solidFill(OFFICE_YELLOW);
                custCell.border = {
                    top: { style: 'thin' },
                    bottom: { style: 'thin' },
                };
            }
            excelRow++;
        });
    }
    ws.getCell(`B${excelRow}`).value = 'Grand Total';
    ws.getCell(`E${excelRow}`).value = model?.grandTotal ?? 0;
    ws.getCell(`E${excelRow}`).numFmt = HOURS_FMT;
    ws.getCell(`B${excelRow}`).font = { ...bodyFont };
    ws.getCell(`E${excelRow}`).font = { ...bodyFont };
    // Grand Total: double top on label + hours; label grey fill.
    ws.getCell(`B${excelRow}`).fill = solidFill(GRAND_LABEL_GRAY);
    ws.getCell(`B${excelRow}`).border = { top: { style: 'double' } };
    ws.getCell(`E${excelRow}`).border = { top: { style: 'double' } };
}
// ---------------------------------------------------------------------------
// Payroll derived sheets (Garamond; values for core)
// ---------------------------------------------------------------------------
function applyGaramond(ws, size = 12) {
    ws.eachRow((row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
            if (!cell.font?.name)
                cell.font = { name: PP_FONT, size, ...(cell.font ?? {}) };
            else if (cell.font.name !== PP_FONT)
                cell.font = { ...cell.font, name: PP_FONT };
        });
    });
}
function buildPlImportSheet(wb, pl, 
// Round 8 §7.3 — the UT tree renders identically onto its own sheet.
sheetName = 'P&L by Job Import') {
    const ws = wb.addWorksheet(sheetName);
    // Round 7 — full uploaded account tree when present; else fall back to byJob summary.
    if (pl && pl.columnHeaders.length > 0 && pl.accountRows.length > 0) {
        const headers = [...pl.columnHeaders];
        if (!headers[0])
            headers[0] = 'Account';
        ws.columns = headers.map((_, i) => ({ width: i === 0 ? 28 : 14 }));
        const hdr = ws.addRow(headers);
        for (let c = 1; c <= headers.length; c++) {
            hdr.getCell(c).fill = solidFill(HEADER_GRAY);
            hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
        }
        for (const ar of pl.accountRows) {
            const cells = [ar.label];
            for (let c = 1; c < headers.length; c++) {
                cells.push(ar.values[c] ?? null);
            }
            const row = ws.addRow(cells);
            row.getCell(1).font = { name: PP_FONT, size: 12 };
            for (let c = 2; c <= headers.length; c++) {
                const cell = row.getCell(c);
                if (typeof cell.value === 'number')
                    cell.numFmt = ACCOUNTING_FMT;
                cell.font = { name: PP_FONT, size: 12 };
            }
        }
        applyGaramond(ws);
        return;
    }
    const jobs = [...(pl?.byJob.entries() ?? [])];
    const headers = ['Account', ...jobs.map(([ap, m]) => m.plJobHeader || `AP${ap}`)];
    ws.columns = headers.map((_, i) => ({ width: i === 0 ? 28 : 14 }));
    const hdr = ws.addRow(headers);
    for (let c = 1; c <= headers.length; c++) {
        hdr.getCell(c).fill = solidFill(HEADER_GRAY);
        hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
    }
    const metricRows = [
        { label: 'Total Income', pick: (m) => m.income },
        { label: 'Materials', pick: (m) => m.materials },
        { label: 'QB Labor-Wages-Subs', pick: (m) => m.laborWagesSubs },
        { label: 'Mileage Expense', pick: (m) => m.mileage },
        { label: 'Sundries -', pick: (m) => m.sundries },
    ];
    for (const mr of metricRows) {
        const row = ws.addRow([mr.label, ...jobs.map(([, m]) => mr.pick(m))]);
        row.getCell(1).font = { name: PP_FONT, size: 12, bold: true };
        for (let c = 2; c <= headers.length; c++) {
            row.getCell(c).numFmt = ACCOUNTING_FMT;
            row.getCell(c).font = { name: PP_FONT, size: 12 };
        }
    }
    applyGaramond(ws);
}
function writePsChainValues(ws, block, startCol, opts) {
    const jobCount = Math.min(block.jobs.length, 6);
    const baseFont = { name: PP_FONT, size: 12 };
    const style = opts.style ?? 'default';
    let r = opts.startRow ?? ws.rowCount + 1;
    const set = (row, col, v, fmt) => {
        const cell = ws.getCell(row, col);
        cell.value = v;
        cell.font = { ...baseFont };
        if (fmt)
            cell.numFmt = fmt;
    };
    const nameFill = style === 'profitshare' ? NAME_BAND_RED : style === 'management' ? NAME_BAND_GREEN : TAN;
    const nameFg = style === 'profitshare' || style === 'management' ? WHITE : BLACK;
    // Name / WO% band
    set(r, startCol, block.displayName);
    ws.getCell(r, startCol).fill = solidFill(nameFill);
    ws.getCell(r, startCol).font = { ...baseFont, bold: true, color: { argb: nameFg } };
    if (!opts.skipWoBand) {
        for (let i = 0; i < jobCount; i++) {
            const pct = block.jobs[i].woPct;
            set(r, startCol + 1 + i, pct !== null ? pct : '-', pct !== null ? PCT_FMT : undefined);
            ws.getCell(r, startCol + 1 + i).fill = solidFill(nameFill);
            ws.getCell(r, startCol + 1 + i).font = { ...baseFont, color: { argb: nameFg } };
        }
    }
    else {
        // FINAL — keep name cell only (no WO% band across job columns).
        for (let i = 0; i < jobCount; i++) {
            ws.getCell(r, startCol + 1 + i).fill = solidFill(nameFill);
        }
    }
    r++;
    set(r, startCol, 'PO');
    for (let i = 0; i < jobCount; i++) {
        const j = block.jobs[i];
        set(r, startCol + 1 + i, j.jobShort);
        if (style === 'profitshare' && i === 0) {
            ws.getCell(r, startCol + 1 + i).fill = solidFill(HOURLY_PAY_YELLOW);
            ws.getCell(r, startCol + 1 + i).font = { ...baseFont, size: 10.5, bold: true };
        }
        else if (j.colorHex) {
            ws.getCell(r, startCol + 1 + i).fill = solidFill('FF' + j.colorHex);
            ws.getCell(r, startCol + 1 + i).font = {
                ...baseFont,
                size: 10.5,
                bold: true,
                color: { argb: textOn(j.colorHex) },
            };
        }
    }
    r++;
    set(r, startCol, 'Name');
    for (let i = 0; i < jobCount; i++) {
        const j = block.jobs[i];
        set(r, startCol + 1 + i, j.customer);
        if (style === 'profitshare' && i === 0) {
            ws.getCell(r, startCol + 1 + i).fill = solidFill(HOURLY_PAY_YELLOW);
            ws.getCell(r, startCol + 1 + i).font = { ...baseFont, size: 10.5, bold: true };
        }
        else if (j.colorHex) {
            ws.getCell(r, startCol + 1 + i).fill = solidFill('FF' + j.colorHex);
            ws.getCell(r, startCol + 1 + i).font = {
                ...baseFont,
                size: 10.5,
                bold: true,
                color: { argb: textOn(j.colorHex) },
            };
        }
    }
    r++;
    // Round 8 §8.1 — SUBBED / SET RATE bands carry no auto profit share: the
    // computed cells export blank and the % cell shows the state verbatim.
    const cut = block.jobs.slice(0, jobCount);
    const rows = [
        { label: 'Profit Share', vals: cut.map((j) => (j.manualPay ? '' : j.profitShare)), fmt: MONEY_FMT },
        {
            label: 'Work Order Override',
            vals: cut.map((j) => (j.workOrderOverride !== null ? j.workOrderOverride : '')),
            fmt: ACCOUNTING_FMT,
        },
        {
            label: '% Completion',
            vals: cut.map((j) => (j.manualPay ? j.manualPay : j.percentCompletion)),
            fmt: PCT_FMT,
        },
        { label: '- Labor + Wages', vals: cut.map((j) => j.laborWages), fmt: MONEY_FMT },
        { label: '-Mileage', vals: cut.map((j) => j.mileage), fmt: MONEY_FMT },
        { label: '-½ Materials', vals: cut.map((j) => j.halfMaterials), fmt: MONEY_FMT },
        {
            label: 'Total Profit Share',
            vals: cut.map((j) => (j.manualPay ? '' : j.totalProfitShare)),
            fmt: ACCOUNTING_FMT,
            bold: true,
        },
    ];
    let halfMatRow = 0;
    for (const rr of rows) {
        if (rr.label === '-½ Materials')
            halfMatRow = r;
        set(r, startCol, rr.label);
        if (rr.bold)
            ws.getCell(r, startCol).font = { ...baseFont, bold: true };
        for (let i = 0; i < jobCount; i++) {
            set(r, startCol + 1 + i, rr.vals[i], rr.fmt);
            if (rr.bold)
                ws.getCell(r, startCol + 1 + i).font = { ...baseFont, bold: true };
            // FINAL — `@ NN% PS` beside each TPS amount when job WO% ≠ default.
            if (opts.skipWoBand && rr.label === 'Total Profit Share') {
                const j = block.jobs[i];
                const def = opts.defaultPsPct;
                if (j.woPct !== null && Math.abs(j.woPct - def) > 1e-9) {
                    set(r, startCol + 1 + jobCount + i, `@ ${Math.round(j.woPct * 100)}% PS`);
                    ws.getCell(r, startCol + 1 + jobCount + i).font = { ...baseFont, size: 10 };
                }
            }
        }
        r++;
    }
    if (halfMatRow > 0) {
        for (let c = 0; c < 5; c++) {
            ws.getCell(halfMatRow, startCol + c).border = { bottom: { style: 'thin' } };
        }
    }
    if (block.jobs.some((j) => j.splitWith > 1)) {
        set(r, startCol, 'split w/');
        for (let i = 0; i < jobCount; i++) {
            const sw = block.jobs[i].splitWith;
            set(r, startCol + 1 + i, sw > 1 ? sw : '');
        }
        r++;
    }
    r++; // blank
    const footer = [
        { label: 'Total', value: block.total, bold: true },
        { label: 'Mileage - New', value: block.mileageNew },
        { label: 'Hourly Pay', value: block.hourlyPay, highlight: true },
        { label: 'Reimbursements', value: block.reimbursements, memo: block.reimbMemo, color: REIMB_BLUE },
        { label: 'Deductions/Sundries', value: block.deductions, memo: block.deductMemo, color: DEDUCT_RED },
        { label: 'Total Pay', value: block.totalPay, bold: true, money: true, cream: style === 'profitshare' },
    ];
    let deductFooterRow = 0;
    for (const f of footer) {
        if (f.label === 'Deductions/Sundries')
            deductFooterRow = r;
        set(r, startCol, f.label);
        set(r, startCol + 1, f.value, f.money ? MONEY_FMT : ACCOUNTING_FMT);
        if (f.memo)
            set(r, startCol + 2, f.memo);
        if (f.bold) {
            ws.getCell(r, startCol).font = { ...baseFont, bold: true };
            ws.getCell(r, startCol + 1).font = { ...baseFont, bold: true };
        }
        if (f.highlight)
            ws.getCell(r, startCol + 1).fill = solidFill(HOURLY_PAY_YELLOW);
        if (f.cream) {
            ws.getCell(r, startCol).fill = solidFill(TOTAL_PAY_CREAM);
            ws.getCell(r, startCol + 1).fill = solidFill(TOTAL_PAY_CREAM);
        }
        if (f.color)
            ws.getCell(r, startCol + 1).font = { ...baseFont, color: { argb: f.color } };
        r++;
    }
    if (deductFooterRow > 0) {
        for (let c = 0; c < 5; c++) {
            ws.getCell(deductFooterRow, startCol + c).border = { bottom: { style: 'double' } };
        }
    }
    if (opts.showGrossWages) {
        const gross = block.total + block.hourlyPay;
        set(r, startCol, 'Total Gross Wages:');
        set(r, startCol + 1, gross, MONEY_FMT);
        ws.getCell(r, startCol).font = { ...baseFont, bold: true };
        ws.getCell(r, startCol + 1).font = { ...baseFont, bold: true };
        r++;
    }
    r++; // blank before income block
    const info = [
        { label: 'Total Income', vals: block.jobs.slice(0, jobCount).map((j) => j.income), fmt: MONEY_FMT },
        { label: 'Materials', vals: block.jobs.slice(0, jobCount).map((j) => j.materials), fmt: MONEY_FMT },
        // Round 8 §8.5 — informational, directly beneath Materials.
        { label: 'Travel', vals: block.jobs.slice(0, jobCount).map((j) => j.travel), fmt: MONEY_FMT },
        { label: 'QB Labor-Wages-Subs', vals: block.jobs.slice(0, jobCount).map((j) => j.laborWages), fmt: MONEY_FMT },
        { label: 'Mileage in Books', vals: block.jobs.slice(0, jobCount).map((j) => j.mileage), fmt: MONEY_FMT },
        {
            label: '% Materials',
            vals: block.jobs.slice(0, jobCount).map((j) => (j.income ? j.materials / j.income : 0)),
            fmt: PCT_FMT,
        },
    ];
    for (const ir of info) {
        set(r, startCol, ir.label);
        for (let i = 0; i < jobCount; i++)
            set(r, startCol + 1 + i, ir.vals[i], ir.fmt);
        r++;
    }
    return r;
}
function moneyPlain(n) {
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
/** NOTES-only materials reconciliation table + WO note in col F. */
function writeLeadMaterialsTable(ws, block) {
    const name = ws.addRow([block.displayName]);
    name.getCell(1).fill = solidFill(TAN);
    name.getCell(1).font = { name: PP_FONT, bold: true, size: 12 };
    const hdr = ws.addRow(['PO#', 'Job', 'Reported Materials', 'Total Materials', 'DIFFERENCE']);
    for (let c = 1; c <= 5; c++) {
        hdr.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 12 };
    }
    for (const j of block.jobs.slice(0, 6)) {
        const woBase = j.workOrderOverride !== null ? j.workOrderOverride : j.profitShare;
        const woNote = `w.o. $${moneyPlain(woBase)} less ½ mats $${moneyPlain(j.halfMaterials)}`;
        const row = ws.addRow([
            j.apKey ? `AP${j.apKey}` : j.jobShort,
            `${j.customer} ${j.jobShort}`.trim(),
            '', // Reported Materials — blank working cell
            j.materials,
            '', // DIFFERENCE — blank until reported is filled
            woNote,
        ]);
        row.getCell(4).numFmt = MONEY_FMT;
        for (let c = 1; c <= 6; c++)
            row.getCell(c).font = { name: PP_FONT, size: 12 };
        if (j.colorHex) {
            row.getCell(1).fill = solidFill('FF' + j.colorHex);
            row.getCell(1).font = {
                name: PP_FONT,
                size: 12,
                color: { argb: textOn(j.colorHex) },
            };
        }
    }
    ws.addRow([]);
}
function buildProfitshareOutputSheet(wb, model, defaultPsPct, mileageRate = 0.625) {
    const ws = wb.addWorksheet('Profitshare Output');
    // 8-col bands: 7 data + 1 spacer → A, I, Q, Y…
    const leads = model?.leads ?? [];
    ws.columns = Array.from({ length: Math.max(8, leads.length * 8) }, (_, i) => ({
        width: i % 8 === 7 ? 3 : 14,
    }));
    let maxEndRow = 1;
    leads.forEach((block, bi) => {
        const startCol = 1 + bi * 8;
        const end = writePsChainValues(ws, block, startCol, {
            defaultPsPct,
            startRow: 1,
            style: 'profitshare',
        });
        if (end > maxEndRow)
            maxEndRow = end;
    });
    if (leads.length === 0) {
        ws.addRow(['(no crew leads this week)']).font = { name: PP_FONT, italic: true };
        maxEndRow = 2;
    }
    // Footer lookups under the tallest band.
    let fr = maxEndRow + 1;
    ws.getCell(fr, 1).value = 'Lookups - DO NOT CHANGE:';
    ws.getCell(fr, 1).font = { name: PP_FONT, bold: true, size: 12 };
    fr++;
    ws.getCell(fr, 1).value = 'Mileage Rate';
    ws.getCell(fr, 2).value = mileageRate;
    ws.getCell(fr, 2).numFmt = '0.000';
    ws.getCell(fr, 1).font = { name: PP_FONT, size: 12 };
    ws.getCell(fr, 2).font = { name: PP_FONT, size: 12 };
    applyGaramond(ws);
}
function buildManagementReportSheet(wb, model, defaultPsPct) {
    const ws = wb.addWorksheet('Management Report Export');
    ws.columns = Array.from({ length: 8 }, () => ({ width: 16 }));
    const leads = model?.leads ?? [];
    const jobCols = 6;
    const hdr = ws.addRow(['Crew', ...Array.from({ length: jobCols }, (_, i) => `Job ${i + 1}`)]);
    for (let c = 1; c <= 1 + jobCols; c++) {
        hdr.getCell(c).fill = solidFill(HEADER_GRAY);
        hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
    }
    for (const block of leads) {
        writePsChainValues(ws, block, 1, { defaultPsPct, style: 'management' });
        ws.addRow([]);
    }
    if (leads.length === 0)
        ws.addRow(['(no crew leads this week)']);
    applyGaramond(ws);
}
function buildInputNumbersSheet(wb, model) {
    const ws = wb.addWorksheet("Input #'s");
    // Col A gutter ~3.4; data starts in B.
    ws.columns = [{ width: 3.4 }, ...Array.from({ length: 8 }, () => ({ width: 16 }))];
    const title = ws.getCell(1, 2);
    title.value = 'Crew Profit Share';
    title.font = { name: PP_FONT, bold: true, size: 20 };
    const leads = model?.leads ?? [];
    let row = 3;
    for (const block of leads) {
        const jobCount = Math.min(block.jobs.length, 6);
        // Red name row in B…
        const nameRow = ws.getRow(row);
        nameRow.getCell(2).value = block.displayName;
        nameRow.getCell(2).fill = solidFill(NAME_BAND_RED);
        nameRow.getCell(2).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
        for (let i = 0; i < jobCount; i++) {
            const pct = block.jobs[i].woPct;
            nameRow.getCell(3 + i).value = pct !== null ? pct : '';
            if (pct !== null)
                nameRow.getCell(3 + i).numFmt = PCT_FMT;
            nameRow.getCell(3 + i).fill = solidFill(NAME_BAND_RED);
            nameRow.getCell(3 + i).font = { name: PP_FONT, size: 12, color: { argb: WHITE } };
        }
        row++;
        // Grey PO labels + first PO yellow.
        const poRow = ws.getRow(row);
        poRow.getCell(2).value = "PO #'s:";
        poRow.getCell(2).fill = solidFill(HEADER_GRAY);
        poRow.getCell(2).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
        for (let i = 0; i < jobCount; i++) {
            poRow.getCell(3 + i).value = block.jobs[i].jobShort;
            poRow.getCell(3 + i).font = { name: PP_FONT, size: 11, bold: true };
            if (i === 0)
                poRow.getCell(3 + i).fill = solidFill(HOURLY_PAY_YELLOW);
            else
                poRow.getCell(3 + i).fill = solidFill(HEADER_GRAY);
        }
        row++;
        const woRow = ws.getRow(row);
        woRow.getCell(2).value = 'WO Amount/Contract Override:';
        woRow.getCell(2).font = { name: PP_FONT, size: 12 };
        for (let i = 0; i < jobCount; i++) {
            const v = block.jobs[i].workOrderOverride;
            woRow.getCell(3 + i).value = v !== null ? v : '';
            if (v !== null)
                woRow.getCell(3 + i).numFmt = ACCOUNTING_FMT;
            woRow.getCell(3 + i).font = { name: PP_FONT, size: 12 };
        }
        row++;
        const pctRow = ws.getRow(row);
        pctRow.getCell(2).value = '% Completion:';
        pctRow.getCell(2).font = { name: PP_FONT, size: 12 };
        for (let i = 0; i < jobCount; i++) {
            const j = block.jobs[i];
            pctRow.getCell(3 + i).value = j.manualPay ? j.manualPay : j.percentCompletion;
            pctRow.getCell(3 + i).numFmt = PCT_FMT;
            pctRow.getCell(3 + i).font = { name: PP_FONT, size: 12 };
        }
        row++;
        const matHdr = ws.getRow(row);
        matHdr.getCell(2).value = 'Reported Materials:';
        matHdr.getCell(3).value = 'Amount';
        matHdr.getCell(4).value = 'Notes';
        for (let c = 2; c <= 4; c++)
            matHdr.getCell(c).font = { name: PP_FONT, size: 12 };
        row++;
        row++; // blank reported-materials entry row
        const reimb = ws.getRow(row);
        reimb.getCell(2).value = 'Reimbursements';
        reimb.getCell(3).value = block.reimbursements;
        reimb.getCell(3).numFmt = ACCOUNTING_FMT;
        reimb.getCell(4).value = block.reimbMemo;
        reimb.getCell(3).font = { name: PP_FONT, size: 12, color: { argb: REIMB_BLUE } };
        row++;
        const deduct = ws.getRow(row);
        deduct.getCell(2).value = 'Deductions/Sundries';
        deduct.getCell(3).value = block.deductions;
        deduct.getCell(3).numFmt = ACCOUNTING_FMT;
        deduct.getCell(4).value = block.deductMemo;
        deduct.getCell(3).font = { name: PP_FONT, size: 12, color: { argb: DEDUCT_RED } };
        row++;
        const miles = ws.getRow(row);
        miles.getCell(2).value = 'Miles:';
        miles.getCell(2).font = { name: PP_FONT, size: 12 };
        row += 2; // spacer between lead blocks
    }
    if (leads.length === 0) {
        ws.getCell(3, 2).value = '(no crew leads this week)';
        ws.getCell(3, 2).font = { name: PP_FONT, italic: true, size: 12 };
    }
    applyGaramond(ws);
}
function writeHourlyBlockValues(ws, block, opts) {
    if (opts?.tsheets) {
        const title = ws.addRow([`${block.employeeName.split(' ')[0]}'s T-SHEETS BREAKDOWN`]);
        for (let c = 1; c <= 5; c++) {
            title.getCell(c).fill = solidFill('FFF9F9F9');
            title.getCell(c).font = { name: PP_FONT, bold: true, size: 12 };
            title.getCell(c).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                ...(c === 1 ? { left: { style: 'thin' } } : {}),
                ...(c === 5 ? { right: { style: 'thin' } } : {}),
            };
        }
    }
    else {
        const nameRow = ws.addRow([
            block.crewTag ? `${block.employeeName} (${block.crewTag})` : block.employeeName,
        ]);
        nameRow.getCell(1).fill = solidFill(TAN);
        nameRow.getCell(1).font = { name: PP_FONT, bold: true, size: 12 };
    }
    const hdr = ws.addRow(['PO#', 'Job', 'hrs', 'rate', 'total']);
    for (let c = 1; c <= 5; c++) {
        hdr.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 12 };
        hdr.getCell(c).alignment = { horizontal: 'center' };
    }
    for (const l of block.lines) {
        const total = Math.round(l.hours * l.rate * 100) / 100;
        const row = ws.addRow([
            l.poText ?? (l.apKey ? `AP${l.apKey}` : ''),
            l.jobLabel,
            l.hours,
            l.rate,
            total,
        ]);
        for (const c of [3, 4, 5]) {
            row.getCell(c).numFmt = HOURS_FMT;
            row.getCell(c).alignment = { horizontal: 'center' };
            row.getCell(c).font = { name: PP_FONT, size: 12 };
        }
        row.getCell(2).font = { name: PP_FONT, size: 10.5 };
    }
    const sums = ws.addRow(['', '', block.totalHours, '', block.totalPay]);
    sums.font = { name: PP_FONT, bold: true, size: 12 };
    sums.getCell(3).numFmt = HOURS_FMT;
    sums.getCell(5).numFmt = HOURS_FMT;
    ws.addRow([]);
}
function writeSalaryBlockValues(ws, block, weekRange, fill) {
    const band = ws.addRow([`${block.name}${block.timesheetHours > 0 ? ` (${block.timesheetHours}hrs)` : ''}`]);
    band.getCell(1).fill = solidFill(fill);
    band.getCell(1).font = { name: PP_FONT, bold: true, size: 12 };
    const r1 = ws.addRow([weekRange ? `(${weekRange})` : '', block.role, block.amount]);
    r1.getCell(3).numFmt = ACCOUNTING_FMT;
    const r2 = ws.addRow([`mileage reimb: ${block.name} - ${block.miles} miles`, '', block.mileageReimb]);
    r2.getCell(3).numFmt = ACCOUNTING_FMT;
    const r3 = ws.addRow(['Total', '', block.amount + block.mileageReimb]);
    r3.getCell(3).numFmt = ACCOUNTING_FMT;
    r3.font = { name: PP_FONT, bold: true, size: 12 };
    ws.addRow([]);
}
/**
 * NOTES vs FINAL share painters / office / admin / estimators / managers,
 * but lead sections differ:
 * - NOTES: materials reconciliation + T-SHEETS (working log; no PS WO% band)
 * - FINAL: PS chain without WO% band, `@ NN% PS` notes, Total Gross Wages; no materials table
 */
function writePreviewSequence(ws, model, opts) {
    let grossSum = 0;
    for (const entry of model.painterSequence) {
        if (entry.kind === 'hourly') {
            writeHourlyBlockValues(ws, entry.block);
            grossSum += entry.block.totalPay;
        }
        else if (opts.mode === 'notes') {
            // Working log: materials table → T-SHEETS (not the full PS WO% band).
            writeLeadMaterialsTable(ws, entry.block);
            if (entry.block.tsheets)
                writeHourlyBlockValues(ws, entry.block.tsheets, { tsheets: true });
            grossSum += entry.block.totalPay;
        }
        else {
            writePsChainValues(ws, entry.block, 1, {
                skipWoBand: true,
                defaultPsPct: opts.defaultPsPct,
                showGrossWages: true,
            });
            ws.addRow([]);
            if (entry.block.tsheets)
                writeHourlyBlockValues(ws, entry.block.tsheets, { tsheets: true });
            grossSum += entry.block.totalPay;
        }
    }
    const section = (title) => {
        const row = ws.addRow([title]);
        row.getCell(1).fill = solidFill(BLACK);
        row.getCell(1).font = { name: PP_FONT, bold: true, size: 12, color: { argb: WHITE } };
    };
    section('OFFICE');
    for (const b of model.office) {
        writeHourlyBlockValues(ws, b);
        grossSum += b.totalPay;
    }
    section('ADMIN');
    for (const b of model.admin) {
        writeSalaryBlockValues(ws, b, model.weekRange, TAN);
        grossSum += b.amount + b.mileageReimb;
    }
    section('ESTIMATORS');
    for (const b of model.estimators) {
        writeSalaryBlockValues(ws, b, model.weekRange, TAN);
        grossSum += b.amount + b.mileageReimb;
    }
    section('MANAGER SALARY');
    for (const b of model.managers) {
        writeSalaryBlockValues(ws, b, model.weekRange, 'FFD99694');
        grossSum += b.amount + b.mileageReimb;
    }
    return grossSum;
}
function payrollBreakdownTitle(payDate, weekStart, weekEnd, withGuillemet) {
    const pay = payDate
        ? `${payDate.getMonth() + 1}/${payDate.getDate()}/${payDate.getFullYear()}`
        : '';
    const range = weekStart && weekEnd
        ? `${weekStart.getMonth() + 1}/${weekStart.getDate()} - ${weekEnd.getMonth() + 1}/${weekEnd.getDate()}`
        : '';
    const core = `PAYROLL BREAKDOWN ${pay}${range ? ` (${range})` : ''}`.trim();
    return withGuillemet ? `«${core}` : core;
}
function buildNotesSheet(wb, input) {
    const ws = wb.addWorksheet('NOTES');
    ws.columns = Array.from({ length: 8 }, () => ({ width: 16 }));
    const title = ws.addRow([
        payrollBreakdownTitle(input.payDate, input.weekStart, input.weekEnd, true),
    ]);
    title.getCell(1).font = { name: PP_FONT, bold: true, size: 13 };
    const model = input.payrollPreview ?? null;
    if (!model) {
        ws.addRow(['(no payroll preview)']);
        applyGaramond(ws);
        return;
    }
    const gross = writePreviewSequence(ws, model, {
        mode: 'notes',
        defaultPsPct: input.defaultProfitSharePct ?? 0.38,
    });
    // Grand-total cell on row 1 (computed value — never a cross-sheet #REF!).
    ws.getCell(1, 10).value = gross;
    ws.getCell(1, 10).numFmt = MONEY_FMT;
    ws.getCell(1, 10).font = { name: PP_FONT, bold: true, size: 12 };
    applyGaramond(ws);
}
function buildFinalPayrollSheet(wb, input) {
    const ws = wb.addWorksheet('FINAL PAYROLL BREAKDOWN');
    ws.columns = Array.from({ length: 10 }, () => ({ width: 16 }));
    const title = ws.addRow([
        payrollBreakdownTitle(input.payDate, input.weekStart, input.weekEnd, false),
    ]);
    title.getCell(1).font = { name: PP_FONT, bold: true, size: 13 };
    const model = input.payrollPreview ?? null;
    if (!model) {
        ws.addRow(['(no payroll preview)']);
        applyGaramond(ws);
        return;
    }
    writePreviewSequence(ws, model, {
        mode: 'final',
        defaultPsPct: input.defaultProfitSharePct ?? 0.38,
    });
    applyGaramond(ws);
}
/** Round 7 §18 — header shell only (PO# | blank | Job). */
function buildJournalEntrySheet(wb) {
    const ws = wb.addWorksheet('►JOURNAL ENTRY◄');
    ws.columns = [{ width: 12 }, { width: 12 }, { width: 48 }];
    const headers = ['PO#', '', 'Job'];
    const row = ws.addRow(headers);
    for (let c = 1; c <= headers.length; c++) {
        row.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 10.5 };
    }
}
function buildJobsSheet(wb, input) {
    const ws = wb.addWorksheet('JOBS');
    ws.columns = [{ width: 12 }, { width: 48 }];
    const font = { name: PP_FONT, size: 10.5 };
    boldHeader(ws, ['PO#', 'Job'], PP_FONT);
    ws.getRow(1).eachCell((c) => {
        c.font = { ...font, bold: true };
    });
    const colorByEstimator = new Map((input.legend ?? []).map((l) => [l.estimator, l.color.replace(/^#/, '').toUpperCase()]));
    let rows = input.jobs ?? null;
    if (!rows) {
        const map = new Map();
        for (const lead of input.payrollPreview?.leads ?? []) {
            for (const j of lead.jobs) {
                if (!j.apKey)
                    continue;
                const po = `AP${j.apKey}`;
                if (!map.has(po))
                    map.set(po, { jobLabel: `${j.customer} ${j.jobShort}`.trim(), apKey: j.apKey });
            }
        }
        for (const entry of input.payrollPreview?.painterSequence ?? []) {
            const block = entry.kind === 'hourly' ? entry.block : entry.block.tsheets;
            if (!block)
                continue;
            for (const l of block.lines) {
                if (!l.apKey)
                    continue;
                const po = `AP${l.apKey}`;
                if (!map.has(po))
                    map.set(po, { jobLabel: l.jobLabel, apKey: l.apKey });
            }
        }
        rows = [...map.entries()]
            .map(([poLabel, v]) => ({ poLabel, jobLabel: v.jobLabel, apKey: v.apKey }))
            .sort((a, b) => a.jobLabel.localeCompare(b.jobLabel, 'en', { sensitivity: 'base' }));
    }
    for (const r of rows) {
        const row = ws.addRow([r.poLabel, r.jobLabel]);
        row.getCell(1).font = { ...font };
        row.getCell(2).font = { ...font };
        const apKey = r.apKey ?? (0, apkey_1.extractApKey)(r.poLabel);
        const estimator = apKey ? input.estimatorByApKey?.get(apKey) : undefined;
        const hex = estimator ? colorByEstimator.get(estimator) : undefined;
        if (hex) {
            const fill = solidFill(hex.length === 8 ? hex : `FF${hex}`);
            const fg = { argb: textOn(hex.replace(/^FF/i, '').slice(0, 6)) };
            row.getCell(1).fill = fill;
            row.getCell(2).fill = fill;
            row.getCell(1).font = { ...font, color: fg };
            row.getCell(2).font = { ...font, color: fg };
        }
    }
    applyGaramond(ws, 10.5);
}
function buildInactiveEmployeesSheet(wb, input) {
    const ws = wb.addWorksheet('Inactive Employees');
    ws.columns = [{ width: 36 }, { width: 16 }];
    const list = input.inactiveEmployees ??
        input.payrollPreview?.inactiveEmployees ??
        (input.payrollPreview?.inactiveZeroHours ?? []).map((name) => ({
            name,
            displayName: name,
            rateNote: '',
            isLead: false,
            directoryInactive: true,
        }));
    const sorted = [...list].sort((a, b) => a.displayName.localeCompare(b.displayName, 'en', { sensitivity: 'base' }));
    for (const p of sorted) {
        const bandLabel = p.rateNote ? `${p.displayName}  ${p.rateNote}` : p.displayName;
        const band = ws.addRow([bandLabel]);
        band.getCell(1).fill = solidFill(TAN);
        band.getCell(1).font = { name: PP_FONT, bold: true, size: 12 };
        if (p.isLead) {
            const matHdr = ws.addRow(['PO#', 'Job', 'Reported Materials', 'Total Materials', 'DIFFERENCE']);
            for (let c = 1; c <= 5; c++) {
                matHdr.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 12 };
            }
            ws.addRow([]);
        }
        const hdr = ws.addRow(['PO#', 'Job', 'hrs', 'rate', 'total']);
        for (let c = 1; c <= 5; c++) {
            hdr.getCell(c).font = { name: PP_FONT, bold: true, underline: true, size: 12 };
        }
        const rateNum = Number.parseFloat(p.rateNote.replace(/[^0-9.]/g, '')) || 0;
        ws.addRow(['', '', 0, rateNum, 0]);
        ws.addRow([]);
    }
    if (sorted.length === 0) {
        ws.addRow(['(no inactive employees without hours)']).font = {
            name: PP_FONT,
            italic: true,
            size: 12,
        };
    }
    applyGaramond(ws);
}
/** Round 7 §2 — PR Final Check (last payroll sheet). Dynamic from P&L byJob. */
function buildPrFinalCheckSheet(wb, pl, sourceFileName) {
    const ws = wb.addWorksheet('PR Final Check', { views: [{ state: 'frozen', ySplit: 1 }] });
    const headers = [
        'PO#',
        'Job',
        'Total Income',
        'Materials',
        'QB Labor-Wages-Subs',
        'Mileage Expense',
        'Sundries -',
    ];
    ws.columns = headers.map((h) => ({ width: h === 'Job' ? 36 : 16 }));
    const hdr = ws.addRow(headers);
    for (let c = 1; c <= headers.length; c++) {
        hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 11 };
    }
    // Legend cell names the P&L input file.
    const legendCol = headers.length + 2;
    ws.getCell(1, legendCol).value = sourceFileName?.trim() ? sourceFileName : '';
    ws.getCell(1, legendCol).font = { name: PP_FONT, italic: true, size: 10 };
    ws.getColumn(legendCol).width = 28;
    const rows = pl ? (0, prFinalCheck_1.buildPrFinalCheckRows)(pl) : [];
    for (const r of rows) {
        const job = r.customerName && r.jobHeader && !r.jobHeader.includes(r.customerName)
            ? `${r.customerName} ${r.jobHeader}`.trim()
            : r.customerName || r.jobHeader || `AP${r.apKey}`;
        const row = ws.addRow([
            `AP${r.apKey}`,
            job,
            r.income,
            r.materials,
            r.laborWagesSubs,
            r.mileage,
            r.sundries,
        ]);
        row.getCell(1).font = { name: PP_FONT, size: 11 };
        row.getCell(2).font = { name: PP_FONT, size: 11 };
        for (let c = 3; c <= 7; c++) {
            row.getCell(c).numFmt = ACCOUNTING_FMT;
            row.getCell(c).font = { name: PP_FONT, size: 11 };
        }
    }
    applyGaramond(ws, 11);
}
