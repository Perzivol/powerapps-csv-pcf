/* eslint-disable */
/*
 * Standalone self-test for csvParser.ts. Pure Node — no PCF/DOM required.
 * Run with:  npm run selftest
 * Exits non-zero if any assertion fails.
 */
import {
    parseCsv,
    tokenize,
    stripBom,
    normalizeNumber,
    parseMdyToIso,
    ParseResult,
} from "../CsvLineItemImport/csvParser";

let failures = 0;

function check(name: string, cond: boolean, detail?: string): void {
    if (cond) {
        console.log(`  PASS  ${name}`);
    } else {
        failures++;
        console.error(`  FAIL  ${name}${detail ? "  ->  " + detail : ""}`);
    }
}

function eq<T>(name: string, actual: T, expected: T): void {
    check(
        name,
        JSON.stringify(actual) === JSON.stringify(expected),
        `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
    );
}

console.log("csvParser self-test\n");

// ---- stripBom ----
eq("stripBom removes BOM", stripBom("﻿abc"), "abc");
eq("stripBom leaves clean text", stripBom("abc"), "abc");

// ---- normalizeNumber ----
eq("number: padded + thousands sep", normalizeNumber(" 18,801.22 "), { value: 18801.22, ok: true });
eq("number: 952,816.96", normalizeNumber(" 952,816.96 "), { value: 952816.96, ok: true });
eq("number: plain 1.00", normalizeNumber("1.00"), { value: 1, ok: true });
eq("number: 56.00", normalizeNumber("56.00"), { value: 56, ok: true });
eq("number: non-numeric -> error/0", normalizeNumber("N/A"), { value: 0, ok: false });
eq("number: blank -> error/0", normalizeNumber("   "), { value: 0, ok: false });

// ---- parseMdyToIso ----
eq("date: 3/7/2026", parseMdyToIso("3/7/2026"), "2026-03-07");
eq("date: 12/31/2027", parseMdyToIso("12/31/2027"), "2027-12-31");
eq("date: blank -> empty", parseMdyToIso(""), "");
eq("date: invalid 2/30/2026 -> empty", parseMdyToIso("2/30/2026"), "");
eq("date: garbage -> empty", parseMdyToIso("not-a-date"), "");

// ---- tokenize: quotes, escapes, embedded commas, CRLF + bare LF ----
const grid = tokenize('a,b,c\r\n"x,1","he said ""hi""",z\nlast,line,here');
eq("tokenize row count (CRLF + LF mix)", grid.length, 3);
eq("tokenize embedded comma preserved", grid[1][0], "x,1");
eq("tokenize escaped quotes unescaped", grid[1][1], 'he said "hi"');
eq("tokenize bare-LF row", grid[2], ["last", "line", "here"]);
eq("tokenize no phantom trailing row", tokenize("a,b\r\nc,d\r\n").length, 2);

// ---- parseCsv: end-to-end with BOM + CRLF + every quirk ----
const csv =
    "﻿product_code,product_description,unit_cost,quantity,coverage_start_date,coverage_end_date,total_cost,source_document,source_page\r\n" +
    'P-1,"Widget, deluxe"," 18,801.22 ",1.00,3/7/2026,3/6/2027," 18,801.22 ",doc1.pdf,12\r\n' +
    'P-2,"Bad cost row",N/A,2.00,1/1/2026,,"2,000.00",doc2.pdf,5\r\n';
const res: ParseResult = parseCsv(csv);
eq("parseCsv rowCount", res.rowCount, 2);
eq("parseCsv errorRowCount", res.errorRowCount, 1);
eq("parseCsv errorRowNumbers", res.errorRowNumbers, [2]);
eq("parseCsv row1 description (embedded comma)", res.rows[0].Description, "Widget, deluxe");
eq("parseCsv row1 UnitCost normalized", res.rows[0].UnitCost, 18801.22);
eq("parseCsv row1 StartDate ISO", res.rows[0].StartDate, "2026-03-07");
eq("parseCsv row1 HasError false", res.rows[0].HasError, false);
eq("parseCsv row2 blank EndDate", res.rows[1].EndDate, "");
eq("parseCsv row2 HasError true (N/A unit_cost)", res.rows[1].HasError, true);
check(
    "parseCsv totalCostSum = 20801.22",
    Math.abs(res.totalCostSum - (18801.22 + 2000)) < 0.001,
    `got ${res.totalCostSum}`
);

// ---- parseCsv: missing column ----
const missing = parseCsv("product_code,unit_cost\r\nX,1\r\n");
check(
    "parseCsv flags missing columns",
    (missing.error ?? "").startsWith("Missing column:"),
    missing.error
);

// ---- parseCsv: empty file ----
eq("parseCsv empty file", parseCsv("").error, "No data: the file is empty.");

// ---- parseCsv: header-only file ----
const headerOnly = parseCsv(
    "﻿product_code,product_description,unit_cost,quantity,coverage_start_date,coverage_end_date,total_cost,source_document,source_page\r\n"
);
eq("parseCsv header-only rowCount", headerOnly.rowCount, 0);
check("parseCsv header-only has no structural error", headerOnly.error === undefined, headerOnly.error);

console.log(failures === 0 ? "\nALL TESTS PASSED" : `\n${failures} TEST(S) FAILED`);
process.exit(failures === 0 ? 0 : 1);
