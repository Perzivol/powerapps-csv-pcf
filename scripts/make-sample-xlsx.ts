/*
 * Generates sample-data/sample-lineitems.xlsx — the canonical fixture for the control.
 * Run with:  npm run make-sample
 *
 * The selftest round-trips this file and asserts: RowCount 20, ErrorRowNumbers "14",
 * AmountSum 1034893.50. Keep those in sync if you change the rows below.
 */
import * as fs from "fs";
import * as path from "path";
import { buildXlsxFromRows } from "../test/xlsxFixture";

/** Excel 1900-system serial for a UTC calendar date (serial 25569 = 1970-01-01). */
function serialFor(y: number, m: number, d: number): number {
    return Math.round(Date.UTC(y, m - 1, d) / 86400000) + 25569;
}

const GL = {
    furniture: "16-00-001 - Furniture and Equipment",
    licenses: "16-01-008 - Software Licenses",
    hardware: "80-01-005 - Hardware Support & Maintenance",
    swMaint: "80-01-009 - Software Maintenance",
    hosted: "80-01-010 - Hosted Services",
    consulting: "80-04-002 - Consulting Services",
    other: "80-07-001 - Other Expense",
    supplies: "80-01-017 - Supplies",
    equipOps: "80-01-006 - Equipment Operating Expense",
    gaSoftware: "80-01-011 GA Software Purchases",
};

const header = ["Line Description", "Project #", "GL Account Description", "Amount", "Business Group", "Budgeted", "Start Date", "End Date"];

// Dates as a mix of native serial numbers (how Excel stores formatted dates) and
// M/D/YYYY text; amounts as a mix of native numbers and thousands-separated text.
const rows: (string | number | null)[][] = [
    header,
    ["#16374 EMTP Basic", "IG0087.76", GL.swMaint, 3420, "NPG", "Yes", serialFor(2026, 7, 29), serialFor(2027, 7, 29)],
    ["Protection Toolbox", "IG0087.76", GL.swMaint, 1260, "NPG", "Yes", serialFor(2026, 7, 29), serialFor(2027, 7, 29)],
    ["Exciters & Governors", "IG0087.76", GL.swMaint, 720, "NPG", "Yes", "7/29/2026", "7/29/2027"],
    ["Server rack refresh", "IG0102.01", GL.hardware, 15000.5, "NPG", "Yes", serialFor(2026, 1, 15), serialFor(2027, 1, 14)],
    ["Enterprise ERP hosting", "IG0090.03", GL.hosted, "952,816.96", "NPG", "Yes", "1/1/2026", "12/31/2026"],
    ["Cable ties & fasteners", "IG0102.01", GL.supplies, 88.25, "NPG", "No", null, null],
    ["CAD seat renewal", "IG0095.10", GL.licenses, 1200, "NPG", "Yes", serialFor(2026, 3, 1), serialFor(2027, 2, 28)],
    ["Office chairs (2)", "IG0102.01", GL.furniture, 450.75, "NPG", "No", null, null],
    ["Network monitoring SaaS", "IG0090.03", GL.hosted, 9800, "NPG", "Yes", "4/1/2026", "3/31/2027"],
    ["Grid study consultants", "IG0110.02", GL.consulting, 12500, "NPG", "Yes", serialFor(2026, 5, 1), serialFor(2026, 10, 31)],
    // A fully blank spacer row — must be skipped without breaking row numbering.
    [null, null, null, null, null, null, null, null],
    ["Misc project expense", "IG0110.02", GL.other, 300, "NPG", "No", null, null],
    ["Diesel generator service", "IG0115.01", GL.equipOps, 7600.4, "NPG", "Yes", "6/15/2026", "6/14/2027"],
    ["Printer toner stock", "IG0102.01", GL.supplies, 215.1, "NPG", "No", null, null],
    ["License count TBD", "IG0095.10", GL.licenses, "TBD", "NPG", "No", "9/1/2026", "8/31/2027"],
    ["Historian database", "IG0090.03", GL.gaSoftware, 5400, "NPG", "Yes", serialFor(2026, 8, 1), serialFor(2027, 7, 31)],
    ["SCADA support contract", "IG0087.76", GL.hardware, "18,801.22", "NPG", "Yes", "3/7/2026", "3/6/2027"],
    ["Adapter cables", "IG0102.01", GL.supplies, 99.99, "NPG", "No", null, null],
    ["Relay test consultants", "IG0110.02", GL.consulting, 2750, "NPG", "Yes", "10/1/2026", "11/30/2026"],
    ["Spare breaker parts", "IG0115.01", GL.equipOps, 640, "NPG", "No", null, null],
    ["Cyber audit retainer", "IG0110.02", GL.consulting, 1830.33, "NPG", "Yes", serialFor(2026, 2, 1), serialFor(2027, 1, 31)],
];

const out = path.join(__dirname, "../../../sample-data/sample-lineitems.xlsx");
fs.writeFileSync(out, buildXlsxFromRows(rows));
console.log(`Wrote ${out} (${fs.statSync(out).size} bytes)`);
