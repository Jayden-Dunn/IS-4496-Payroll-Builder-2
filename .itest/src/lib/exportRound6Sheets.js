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
const format_1 = require("./format");
const BLACK = 'FF000000';
const WHITE = 'FFFFFFFF';
const HEADER_GRAY = 'FF808080';
const TAN = 'FFDDD9C4';
const HOURLY_PAY_YELLOW = 'FFFFFF99';
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
function buildSalesByRepMirrorSheet(wb, name, data) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    const headers = data?.headers?.length
        ? ['AP#', ...data.headers.slice(0, 12), 'Date', 'QBO Job Name']
        : ['AP#', 'Date', 'Customer', 'Name 1', 'Name 2', 'Transaction Type', 'Num', 'Product/Service', 'Memo/Description', 'Sales Price', 'Amount', 'Balance', 'Sales Rep', 'Date', 'QBO Job Name'];
    ws.columns = headers.map((_, i) => ({ width: i === 0 ? 10 : 14 }));
    boldHeader(ws, headers);
    for (const r of data?.rows ?? []) {
        const cols = r.cols.slice(0, 12);
        while (cols.length < 12)
            cols.push(null);
        const row = ws.addRow([r.apKey, ...cols.map(cellValue), cellValue(r.dateCopy), cellValue(r.qboJobName)]);
        for (let c = 1; c <= headers.length; c++) {
            row.getCell(c).font = { name: FONT, size: 11 };
            const v = row.getCell(c).value;
            if (v instanceof Date)
                row.getCell(c).numFmt = DATE_FMT;
            else if (typeof v === 'number')
                row.getCell(c).numFmt = HOURS_FMT;
        }
    }
}
function buildSalesByCustMirrorSheet(wb, name, rows) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = [{ width: 48 }, { width: 14 }, { width: 10 }, { width: 12 }];
    boldHeader(ws, ['', 'Total', 'AP#', 'Job total']);
    for (const r of rows ?? []) {
        const row = ws.addRow([r.text, r.value ?? '', r.apKey, r.jobTotal ?? '']);
        for (let c = 1; c <= 4; c++) {
            row.getCell(c).font = { name: FONT, size: 11 };
            if (c === 2 || c === 4) {
                if (typeof row.getCell(c).value === 'number')
                    row.getCell(c).numFmt = HOURS_FMT;
            }
        }
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
    ws.columns = [{ width: 3 }, { width: 28 }, { width: 28 }, { width: 40 }, { width: 12 }];
    // Anchored at B2 per §3.6
    ws.getCell('B2').value = 'Row Labels';
    ws.getCell('C2').value = 'jobcode_1';
    ws.getCell('D2').value = 'jobcode_2';
    ws.getCell('E2').value = 'Sum of hours';
    for (const col of ['B', 'C', 'D', 'E']) {
        ws.getCell(`${col}2`).font = { name: FONT, bold: true, size: 11 };
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
                ws.getCell(`${col}${excelRow}`).font = {
                    name: FONT,
                    size: 11,
                    bold: i === 0 && col === 'B',
                };
            }
            excelRow++;
        });
    }
    ws.getCell(`B${excelRow}`).value = 'Grand Total';
    ws.getCell(`E${excelRow}`).value = model?.grandTotal ?? 0;
    ws.getCell(`E${excelRow}`).numFmt = HOURS_FMT;
    ws.getCell(`B${excelRow}`).font = { name: FONT, bold: true, size: 11 };
    ws.getCell(`E${excelRow}`).font = { name: FONT, bold: true, size: 11 };
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
function buildPlImportSheet(wb, pl) {
    const ws = wb.addWorksheet('P&L by Job Import');
    const jobs = [...(pl?.byJob.entries() ?? [])];
    const headers = ['Account', ...jobs.map(([ap, m]) => m.plJobHeader || `AP${ap}`)];
    ws.columns = headers.map((_, i) => ({ width: i === 0 ? 28 : 14 }));
    const hdr = ws.addRow(headers);
    for (let c = 1; c <= headers.length; c++) {
        hdr.getCell(c).font = { name: PP_FONT, bold: true, size: 12 };
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
    let r = ws.rowCount + 1;
    const set = (row, col, v, fmt) => {
        const cell = ws.getCell(row, col);
        cell.value = v;
        cell.font = { ...baseFont };
        if (fmt)
            cell.numFmt = fmt;
    };
    // Name / WO% band
    set(r, startCol, block.displayName);
    ws.getCell(r, startCol).fill = solidFill(TAN);
    ws.getCell(r, startCol).font = { ...baseFont, bold: true };
    if (!opts.skipWoBand) {
        for (let i = 0; i < jobCount; i++) {
            const pct = block.jobs[i].woPct;
            set(r, startCol + 1 + i, pct !== null ? pct : '-', pct !== null ? PCT_FMT : undefined);
            ws.getCell(r, startCol + 1 + i).fill = solidFill(TAN);
        }
    }
    r++;
    set(r, startCol, 'PO');
    for (let i = 0; i < jobCount; i++) {
        const j = block.jobs[i];
        set(r, startCol + 1 + i, j.jobShort);
        if (j.colorHex) {
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
        if (j.colorHex) {
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
    const rows = [
        { label: 'Profit Share', vals: block.jobs.slice(0, jobCount).map((j) => j.profitShare), fmt: MONEY_FMT },
        {
            label: 'Work Order Override',
            vals: block.jobs.slice(0, jobCount).map((j) => (j.workOrderOverride !== null ? j.workOrderOverride : '')),
            fmt: ACCOUNTING_FMT,
        },
        {
            label: '% Completion',
            vals: block.jobs.slice(0, jobCount).map((j) => j.percentCompletion),
            fmt: PCT_FMT,
        },
        { label: '- Labor + Wages', vals: block.jobs.slice(0, jobCount).map((j) => j.laborWages), fmt: MONEY_FMT },
        { label: '-Mileage', vals: block.jobs.slice(0, jobCount).map((j) => j.mileage), fmt: MONEY_FMT },
        { label: '-½ Materials', vals: block.jobs.slice(0, jobCount).map((j) => j.halfMaterials), fmt: MONEY_FMT },
        {
            label: 'Total Profit Share',
            vals: block.jobs.slice(0, jobCount).map((j) => j.totalProfitShare),
            fmt: ACCOUNTING_FMT,
            bold: true,
        },
    ];
    for (const rr of rows) {
        set(r, startCol, rr.label);
        if (rr.bold)
            ws.getCell(r, startCol).font = { ...baseFont, bold: true };
        for (let i = 0; i < jobCount; i++) {
            set(r, startCol + 1 + i, rr.vals[i], rr.fmt);
            if (rr.bold)
                ws.getCell(r, startCol + 1 + i).font = { ...baseFont, bold: true };
            if (opts.skipWoBand && rr.label === 'Total Profit Share') {
                const j = block.jobs[i];
                const def = opts.defaultPsPct;
                if (j.woPct !== null && Math.abs(j.woPct - def) > 1e-9) {
                    set(r, startCol + 1 + jobCount, `@ ${Math.round(j.woPct * 100)}% PS`);
                }
            }
        }
        r++;
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
        { label: 'Total Pay', value: block.totalPay, bold: true, money: true },
    ];
    for (const f of footer) {
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
        if (f.color)
            ws.getCell(r, startCol + 1).font = { ...baseFont, color: { argb: f.color } };
        r++;
    }
    r++; // blank before income block
    const info = [
        { label: 'Total Income', vals: block.jobs.slice(0, jobCount).map((j) => j.income), fmt: MONEY_FMT },
        { label: 'Materials', vals: block.jobs.slice(0, jobCount).map((j) => j.materials), fmt: MONEY_FMT },
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
function buildProfitshareOutputSheet(wb, model, defaultPsPct) {
    const ws = wb.addWorksheet('Profitshare Output');
    // 8-col bands: 7 data + 1 spacer → band width 8
    const leads = model?.leads ?? [];
    const maxJobs = Math.max(6, ...leads.map((l) => l.jobs.length), 1);
    ws.columns = Array.from({ length: Math.max(8, leads.length * 8) }, () => ({ width: 14 }));
    leads.forEach((block, bi) => {
        const startCol = 1 + bi * 8;
        // Ensure worksheet has enough rows by seeding
        while (ws.rowCount < 1)
            ws.addRow([]);
        // Write into absolute rows starting at 1 for first band; subsequent bands share rows
        if (bi === 0) {
            writePsChainValues(ws, block, startCol, { defaultPsPct });
        }
        else {
            // Replay into shared row positions — write column-shifted values on existing structure
            writePsBandOverlay(ws, block, startCol, defaultPsPct, maxJobs);
        }
    });
    if (leads.length === 0) {
        ws.addRow(['(no crew leads this week)']).font = { name: PP_FONT, italic: true };
    }
    applyGaramond(ws);
}
/** Overlay a PS band into columns starting at `startCol`, aligning to row 1 layout. */
function writePsBandOverlay(ws, block, startCol, defaultPsPct, _maxJobs) {
    // Build a temp worksheet-like write by resetting row cursor via direct cells.
    // Reconstruct the same row order as writePsChainValues starting at row 1.
    const jobCount = Math.min(block.jobs.length, 6);
    const baseFont = { name: PP_FONT, size: 12 };
    let r = 1;
    const set = (row, col, v, fmt) => {
        const cell = ws.getCell(row, col);
        cell.value = v;
        cell.font = { ...baseFont };
        if (fmt)
            cell.numFmt = fmt;
    };
    set(r, startCol, block.displayName);
    ws.getCell(r, startCol).fill = solidFill(TAN);
    for (let i = 0; i < jobCount; i++) {
        const pct = block.jobs[i].woPct;
        set(r, startCol + 1 + i, pct !== null ? pct : '-', pct !== null ? PCT_FMT : undefined);
        ws.getCell(r, startCol + 1 + i).fill = solidFill(TAN);
    }
    r++;
    set(r, startCol, 'PO');
    for (let i = 0; i < jobCount; i++)
        set(r, startCol + 1 + i, block.jobs[i].jobShort);
    r++;
    set(r, startCol, 'Name');
    for (let i = 0; i < jobCount; i++)
        set(r, startCol + 1 + i, block.jobs[i].customer);
    r++;
    const chain = [
        ['Profit Share', (j) => j.profitShare, MONEY_FMT],
        ['Work Order Override', (j) => j.workOrderOverride ?? '', ACCOUNTING_FMT],
        ['% Completion', (j) => j.percentCompletion, PCT_FMT],
        ['- Labor + Wages', (j) => j.laborWages, MONEY_FMT],
        ['-Mileage', (j) => j.mileage, MONEY_FMT],
        ['-½ Materials', (j) => j.halfMaterials, MONEY_FMT],
        ['Total Profit Share', (j) => j.totalProfitShare, ACCOUNTING_FMT],
    ];
    for (const [label, pick, fmt] of chain) {
        set(r, startCol, label);
        for (let i = 0; i < jobCount; i++)
            set(r, startCol + 1 + i, pick(block.jobs[i]), fmt);
        r++;
    }
    if (block.jobs.some((j) => j.splitWith > 1)) {
        set(r, startCol, 'split w/');
        for (let i = 0; i < jobCount; i++) {
            const sw = block.jobs[i].splitWith;
            set(r, startCol + 1 + i, sw > 1 ? sw : '');
        }
        r++;
    }
    r++;
    for (const [label, val, highlight] of [
        ['Total', block.total, false],
        ['Mileage - New', block.mileageNew, false],
        ['Hourly Pay', block.hourlyPay, true],
        ['Reimbursements', block.reimbursements, false],
        ['Deductions/Sundries', block.deductions, false],
        ['Total Pay', block.totalPay, false],
    ]) {
        set(r, startCol, label);
        set(r, startCol + 1, val, ACCOUNTING_FMT);
        if (highlight)
            ws.getCell(r, startCol + 1).fill = solidFill(HOURLY_PAY_YELLOW);
        r++;
    }
    void defaultPsPct;
    r++;
    for (const [label, pick, fmt] of [
        ['Total Income', (j) => j.income, MONEY_FMT],
        ['Materials', (j) => j.materials, MONEY_FMT],
        ['QB Labor-Wages-Subs', (j) => j.laborWages, MONEY_FMT],
        ['Mileage in Books', (j) => j.mileage, MONEY_FMT],
        ['% Materials', (j) => (j.income ? j.materials / j.income : 0), PCT_FMT],
    ]) {
        set(r, startCol, label);
        for (let i = 0; i < jobCount; i++)
            set(r, startCol + 1 + i, pick(block.jobs[i]), fmt);
        r++;
    }
}
function buildManagementReportSheet(wb, model, defaultPsPct) {
    const ws = wb.addWorksheet('Management Report Export');
    ws.columns = Array.from({ length: 8 }, () => ({ width: 16 }));
    const leads = model?.leads ?? [];
    const maxJobs = Math.max(...leads.map((l) => l.jobs.length), 1);
    const hdr = ['Crew', ...Array.from({ length: Math.min(maxJobs, 6) }, (_, i) => `Job ${i + 1}`)];
    boldHeader(ws, hdr, PP_FONT);
    for (const block of leads) {
        writePsChainValues(ws, block, 1, { defaultPsPct });
        ws.addRow([]);
    }
    if (leads.length === 0)
        ws.addRow(['(no crew leads this week)']);
    applyGaramond(ws);
}
function buildInputNumbersSheet(wb, model) {
    const ws = wb.addWorksheet("Input #'s");
    ws.columns = Array.from({ length: 8 }, () => ({ width: 18 }));
    const leads = model?.leads ?? [];
    leads.forEach((block, idx) => {
        const jobCount = Math.min(block.jobs.length, 6);
        const band = ws.addRow([
            `Crew ${idx + 1}`,
            ...block.jobs.slice(0, jobCount).map((_, i) => `Job ${i + 1}`),
        ]);
        for (let c = 1; c <= 1 + jobCount; c++) {
            band.getCell(c).fill = solidFill(TAN);
            band.getCell(c).font = { name: PP_FONT, bold: true, size: 12 };
        }
        const wo = ws.addRow([
            block.displayName,
            ...block.jobs.slice(0, jobCount).map((j) => (j.woPct !== null ? j.woPct : '')),
        ]);
        for (let i = 0; i < jobCount; i++) {
            if (typeof wo.getCell(2 + i).value === 'number')
                wo.getCell(2 + i).numFmt = PCT_FMT;
        }
        ws.addRow(['PO #\'s:', ...block.jobs.slice(0, jobCount).map((j) => j.jobShort)]);
        ws.addRow([
            'WO Amount/Contract Override:',
            ...block.jobs.slice(0, jobCount).map((j) => (j.workOrderOverride !== null ? j.workOrderOverride : '')),
        ]);
        const pct = ws.addRow([
            '% Completion:',
            ...block.jobs.slice(0, jobCount).map((j) => j.percentCompletion),
        ]);
        for (let i = 0; i < jobCount; i++)
            pct.getCell(2 + i).numFmt = PCT_FMT;
        ws.addRow(['Reported Materials:', 'Amount', 'Notes']);
        ws.addRow(['', '', '']);
        ws.addRow(['Reimbursements', block.reimbursements, block.reimbMemo]);
        ws.addRow(['Deductions/Sundries', block.deductions, block.deductMemo]);
        ws.addRow([
            'Miles:',
            ...block.jobs.slice(0, jobCount).map(() => ''),
        ]);
        ws.addRow([]);
    });
    if (leads.length === 0)
        ws.addRow(['(no crew leads this week)']);
    applyGaramond(ws);
}
function writeHourlyBlockValues(ws, block) {
    const nameRow = ws.addRow([
        block.crewTag ? `${block.employeeName} (${block.crewTag})` : block.employeeName,
    ]);
    nameRow.getCell(1).fill = solidFill(TAN);
    nameRow.getCell(1).font = { name: PP_FONT, bold: true, size: 12 };
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
function writePreviewSequence(ws, model, opts) {
    let grossSum = 0;
    for (const entry of model.painterSequence) {
        if (entry.kind === 'hourly') {
            writeHourlyBlockValues(ws, entry.block);
            grossSum += entry.block.totalPay;
        }
        else {
            writePsChainValues(ws, entry.block, 1, {
                skipWoBand: opts.final === true,
                defaultPsPct: opts.defaultPsPct,
            });
            if (opts.final) {
                const gross = entry.block.total + entry.block.hourlyPay;
                const g = ws.addRow(['Total Gross Wages:', gross]);
                g.getCell(2).numFmt = MONEY_FMT;
                g.font = { name: PP_FONT, bold: true, size: 12 };
                ws.addRow([]);
            }
            grossSum += entry.block.totalPay;
            if (entry.block.tsheets)
                writeHourlyBlockValues(ws, entry.block.tsheets);
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
function buildNotesSheet(wb, input) {
    const ws = wb.addWorksheet('Notes');
    ws.columns = Array.from({ length: 8 }, () => ({ width: 16 }));
    const pay = input.payDate
        ? `${input.payDate.getMonth() + 1}/${input.payDate.getDate()}/${input.payDate.getFullYear()}`
        : '';
    const range = input.weekStart && input.weekEnd
        ? `${input.weekStart.getMonth() + 1}/${input.weekStart.getDate()} - ${input.weekEnd.getMonth() + 1}/${input.weekEnd.getDate()}`
        : '';
    const title = ws.addRow([`PAYROLL BREAKDOWN ${pay}${range ? ` (${range})` : ''}`.trim()]);
    title.getCell(1).font = { name: PP_FONT, bold: true, size: 13 };
    const model = input.payrollPreview ?? null;
    if (!model) {
        ws.addRow(['(no payroll preview)']);
        applyGaramond(ws);
        return;
    }
    const gross = writePreviewSequence(ws, model, {
        defaultPsPct: input.defaultProfitSharePct ?? 0.38,
    });
    ws.getCell(1, 10).value = gross;
    ws.getCell(1, 10).numFmt = MONEY_FMT;
    ws.getCell(1, 10).font = { name: PP_FONT, bold: true, size: 12 };
    applyGaramond(ws);
}
function buildFinalPayrollSheet(wb, input) {
    const ws = wb.addWorksheet('Final Payroll Breakdown');
    ws.columns = Array.from({ length: 8 }, () => ({ width: 16 }));
    const pay = input.payDate
        ? `${input.payDate.getMonth() + 1}/${input.payDate.getDate()}/${input.payDate.getFullYear()}`
        : '';
    const range = input.weekStart && input.weekEnd
        ? `${input.weekStart.getMonth() + 1}/${input.weekStart.getDate()} - ${input.weekEnd.getMonth() + 1}/${input.weekEnd.getDate()}`
        : '';
    const title = ws.addRow([`PAYROLL BREAKDOWN ${pay}${range ? ` (${range})` : ''}`.trim()]);
    title.getCell(1).font = { name: PP_FONT, bold: true, size: 13 };
    const model = input.payrollPreview ?? null;
    if (!model) {
        ws.addRow(['(no payroll preview)']);
        applyGaramond(ws);
        return;
    }
    writePreviewSequence(ws, model, {
        final: true,
        defaultPsPct: input.defaultProfitSharePct ?? 0.38,
    });
    applyGaramond(ws);
}
function buildJournalEntrySheet(wb, model) {
    const ws = wb.addWorksheet('Journal Entry');
    ws.columns = [
        { width: 18 },
        { width: 12 },
        { width: 60 },
        { width: 12 },
        { width: 12 },
        { width: 36 },
    ];
    const font = { name: PP_FONT, size: 10.5 };
    boldHeader(ws, ['Account', 'Amount', 'Memo', 'PO#', 'Class', 'Job'], PP_FONT);
    ws.getRow(1).eachCell((c) => {
        c.font = { ...font, bold: true };
    });
    const byAp = new Map();
    for (const lead of model?.leads ?? []) {
        for (const j of lead.jobs) {
            if (!j.apKey)
                continue;
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
    const sorted = [...byAp.values()].sort((a, b) => a.jobShort.localeCompare(b.jobShort, 'en', { sensitivity: 'base' }));
    let amountSum = 0;
    for (const j of sorted) {
        const pctParts = j.parts
            .map((p) => `${p.lead} (${p.amt.toFixed(2)})`)
            .join(', ');
        const firstPct = j.parts[0]?.woPct;
        const pctLabel = firstPct !== null && firstPct !== undefined ? `${Math.round(firstPct * 100)}%` : '';
        const splitNote = j.parts.some((p) => p.split > 1)
            ? `*split x${Math.max(...j.parts.map((p) => p.split))}*`
            : '';
        const memo = `${j.jobShort} [${pctLabel}${splitNote}]:  ${pctParts}`;
        const row = ws.addRow([
            'Employee Labor',
            Math.round(j.amount * 100) / 100,
            memo,
            `AP${j.apKey}`,
            'Painting',
            `${j.customer} ${j.jobShort}`.trim(),
        ]);
        row.getCell(2).numFmt = ACCOUNTING_FMT;
        for (let c = 1; c <= 6; c++)
            row.getCell(c).font = { ...font };
        amountSum += j.amount;
    }
    // Mileage / reimbursements from leads + salary blocks
    const addMileage = (name, miles, amount, po) => {
        if (!amount && !miles)
            return;
        const row = ws.addRow([
            'Mileage',
            amount,
            `mileage reimb:  ${name} - ${miles} miles`,
            po,
            'Office',
            '',
        ]);
        row.getCell(2).numFmt = ACCOUNTING_FMT;
        for (let c = 1; c <= 6; c++)
            row.getCell(c).font = { ...font };
        amountSum += amount;
    };
    for (const lead of model?.leads ?? []) {
        if (lead.mileageNew)
            addMileage(lead.displayName, 0, lead.mileageNew, 'Allpro');
        if (lead.reimbursements) {
            const row = ws.addRow([
                'Mileage',
                lead.reimbursements,
                lead.reimbMemo || `reimb: ${lead.displayName}`,
                'Allpro',
                'Office',
                '',
            ]);
            row.getCell(2).numFmt = ACCOUNTING_FMT;
            for (let c = 1; c <= 6; c++)
                row.getCell(c).font = { ...font };
            amountSum += lead.reimbursements;
        }
        if (lead.deductions && lead.deductMemo) {
            const row = ws.addRow([
                'Sundries -',
                lead.deductions,
                `(${lead.displayName}) ${lead.deductMemo}`,
                '',
                'Painting',
                '',
            ]);
            row.getCell(2).numFmt = ACCOUNTING_FMT;
            for (let c = 1; c <= 6; c++)
                row.getCell(c).font = { ...font };
            amountSum += lead.deductions;
        }
    }
    for (const b of [...(model?.admin ?? []), ...(model?.estimators ?? []), ...(model?.managers ?? [])]) {
        if (b.mileageReimb)
            addMileage(b.name, b.miles, b.mileageReimb, 'Allpro');
    }
    const total = ws.addRow(['Total', Math.round(amountSum * 100) / 100, '', '', '', '']);
    total.getCell(2).numFmt = ACCOUNTING_FMT;
    total.font = { ...font, bold: true };
    applyGaramond(ws, 10.5);
}
function buildJobsSheet(wb, input) {
    const ws = wb.addWorksheet('Jobs');
    ws.columns = [{ width: 12 }, { width: 48 }];
    const font = { name: PP_FONT, size: 10.5 };
    boldHeader(ws, ['PO#', 'Job'], PP_FONT);
    ws.getRow(1).eachCell((c) => {
        c.font = { ...font, bold: true };
    });
    let rows = input.jobs ?? null;
    if (!rows) {
        const map = new Map();
        for (const lead of input.payrollPreview?.leads ?? []) {
            for (const j of lead.jobs) {
                if (!j.apKey)
                    continue;
                const po = `AP${j.apKey}`;
                if (!map.has(po))
                    map.set(po, `${j.customer} ${j.jobShort}`.trim());
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
                    map.set(po, l.jobLabel);
            }
        }
        rows = [...map.entries()]
            .map(([poLabel, jobLabel]) => ({ poLabel, jobLabel }))
            .sort((a, b) => a.jobLabel.localeCompare(b.jobLabel, 'en', { sensitivity: 'base' }));
    }
    for (const r of rows) {
        const row = ws.addRow([r.poLabel, r.jobLabel]);
        row.getCell(1).font = { ...font };
        row.getCell(2).font = { ...font };
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
