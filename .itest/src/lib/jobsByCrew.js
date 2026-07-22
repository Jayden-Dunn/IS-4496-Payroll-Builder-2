"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildJobsByCrew = buildJobsByCrew;
/**
 * Round 6 §3.6 — "Jobs by Crew": crew → customer → job hour nest.
 * Crew order matches Pivot by Crew (A→Z groups; same cleaning as that pivot).
 */
const format_1 = require("./format");
function buildJobsByCrew(rows) {
    const crews = new Map();
    let grandTotal = 0;
    for (const r of rows) {
        if (r.hours === 0)
            continue;
        const crewName = (r.group || '').trim();
        if (!crewName || /^\(blank\)$/i.test(crewName))
            continue;
        let crew = crews.get(crewName);
        if (!crew) {
            crew = { crew: crewName, jobs: new Map(), total: 0 };
            crews.set(crewName, crew);
        }
        const key = `${r.jobcode1}\u0000${r.jobcode2}`;
        let job = crew.jobs.get(key);
        if (!job) {
            job = { customer: r.jobcode1, job: r.jobcode2, hours: 0 };
            crew.jobs.set(key, job);
        }
        job.hours = (0, format_1.round2)(job.hours + r.hours);
        crew.total = (0, format_1.round2)(crew.total + r.hours);
        grandTotal = (0, format_1.round2)(grandTotal + r.hours);
    }
    const groups = [...crews.values()]
        .sort((a, b) => a.crew.localeCompare(b.crew, 'en', { sensitivity: 'base' }))
        .map((c) => {
        const sorted = [...c.jobs.values()].sort((a, b) => {
            const cust = a.customer.localeCompare(b.customer, 'en', { sensitivity: 'base' });
            return cust !== 0 ? cust : a.job.localeCompare(b.job, 'en', { sensitivity: 'base' });
        });
        let lastCust = '';
        const jobs = sorted.map((j) => {
            const showCustomer = j.customer !== lastCust;
            lastCust = j.customer;
            return { customer: j.customer, job: j.job, hours: j.hours, showCustomer };
        });
        return { crew: c.crew, jobs, total: c.total };
    });
    return { groups, grandTotal };
}
