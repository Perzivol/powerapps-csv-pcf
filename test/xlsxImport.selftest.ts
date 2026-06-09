/* eslint-disable */
/*
 * Standalone self-test for xlsxReader.ts + lineItemParser.ts. Pure Node — no PCF/DOM.
 * Run with:  npm run selftest
 * Exits non-zero if any assertion fails.
 */
import * as fs from "fs";
import * as path from "path";
import {
    parseWorkbookGrid,
    normalizeNumber,
    parseMdyToIso,
    parseBoolean,
    parseDateCell,
    excelSerialToIso,
    ParseResult,
} from "../CsvLineItemImport/lineItemParser";
import { readXlsxGrid, decodeXmlEntities, colRefToIndex } from "../CsvLineItemImport/xlsxReader";
import { buildXlsx, buildXlsxFromRows } from "./xlsxFixture";

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

/** Excel 1900-system serial for a UTC calendar date (serial 25569 = 1970-01-01). */
function serialFor(y: number, m: number, d: number): number {
    return Math.round(Date.UTC(y, m - 1, d) / 86400000) + 25569;
}

console.log("xlsx import self-test\n");

// ---- normalizeNumber ----
eq("number: padded + thousands sep", normalizeNumber(" 18,801.22 "), { value: 18801.22, ok: true });
eq("number: plain 1.00", normalizeNumber("1.00"), { value: 1, ok: true });
eq("number: non-numeric -> error/0", normalizeNumber("N/A"), { value: 0, ok: false });
eq("number: blank -> error/0", normalizeNumber("   "), { value: 0, ok: false });

// ---- parseMdyToIso ----
eq("date: 3/7/2026", parseMdyToIso("3/7/2026"), "2026-03-07");
eq("date: 12/31/2027", parseMdyToIso("12/31/2027"), "2027-12-31");
eq("date: invalid 2/30/2026 -> empty", parseMdyToIso("2/30/2026"), "");
eq("date: garbage -> empty", parseMdyToIso("not-a-date"), "");

// ---- parseBoolean ----
eq("bool: Yes", parseBoolean("Yes"), true);
eq("bool: TRUE", parseBoolean("TRUE"), true);
eq("bool: ' true '", parseBoolean(" true "), true);
eq("bool: 1", parseBoolean("1"), true);
eq("bool: no", parseBoolean("no"), false);
eq("bool: blank", parseBoolean(""), false);
eq("bool: y", parseBoolean("y"), false);

// ---- excelSerialToIso ----
eq("serial: 25569 = 1970-01-01", excelSerialToIso(25569, false), "1970-01-01");
eq("serial: 2026-07-29", excelSerialToIso(serialFor(2026, 7, 29), false), "2026-07-29");
eq("serial: time fraction floored", excelSerialToIso(serialFor(2026, 7, 29) + 0.75, false), "2026-07-29");
// 1904 epoch is 1462 days after the 1900 epoch: serial 24107 = 1970-01-01 in the 1904 system.
eq("serial: 1904 system", excelSerialToIso(25569 - 1462, true), "1970-01-01");

// ---- parseDateCell ----
eq("dateCell: serial text", parseDateCell(String(serialFor(2026, 3, 7)), false), "2026-03-07");
eq("dateCell: ISO passthrough", parseDateCell("2026-03-07", false), "2026-03-07");
eq("dateCell: M/D/YYYY text", parseDateCell("3/7/2026", false), "2026-03-07");
eq("dateCell: blank", parseDateCell("  ", false), "");
eq("dateCell: garbage", parseDateCell("soon", false), "");

// ---- colRefToIndex / decodeXmlEntities ----
eq("colRef: A", colRefToIndex("A"), 0);
eq("colRef: Z", colRefToIndex("Z"), 25);
eq("colRef: AA", colRefToIndex("AA"), 26);
eq("entities: named", decodeXmlEntities("a &amp; b &lt;c&gt; &quot;d&quot; &apos;e&apos;"), `a & b <c> "d" 'e'`);
eq("entities: no double decode", decodeXmlEntities("&amp;lt;"), "&lt;");
eq("entities: numeric", decodeXmlEntities("&#233;&#x41;"), "éA");

// ---- happy-path workbook: hand-written XML exercising every cell type ----
const SHARED = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="8">
<si><r><rPr><b/></rPr><t>#16374 EMTP</t></r><r><t xml:space="preserve"> Basic &amp; Co</t></r></si>
<si><t>IG0087.76</t></si>
<si><t>80-01-009   -   Software Maintenance</t></si>
<si><t>NPG</t></si>
<si><t>Yes</t></si>
<si><t>1,234.56</t></si>
<si><t>N/A</t></si>
<si><t>3/7/2026</t></si>
</sst>`;

const inl = (s: string): string => `t="inlineStr"><is><t>${s}</t></is`;
const HEADER =
    `<row r="1">` +
    `<c r="A1" ${inl("  LINE   Description ")}></c>` +
    `<c r="B1" ${inl("Project #")}></c>` +
    `<c r="C1" ${inl("GL  Account   Description")}></c>` +
    `<c r="D1" ${inl("Amount")}></c>` +
    `<c r="E1" ${inl("Business  Group")}></c>` +
    `<c r="F1" ${inl("BUDGETED")}></c>` +
    `<c r="G1" ${inl("Start Date")}></c>` +
    `<c r="H1" ${inl("End Date")}></c>` +
    `</row>`;

const SHEET = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
${HEADER}
<row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2" t="s"><v>1</v></c><c r="C2" t="s"><v>2</v></c><c r="D2"><v>3420.0000000001</v></c><c r="E2" t="s"><v>3</v></c><c r="F2" t="b"><v>1</v></c><c r="G2"><v>${serialFor(2026, 7, 29)}</v></c><c r="H2" t="s"><v>7</v></c></row>
<row r="3"><c r="A3"/><c r="C3"/></row>
<row r="4"><c r="A4" ${inl("Sparse row")}></c><c r="H4" t="e"><v>#DIV/0!</v></c></row>
<row r="5"><c r="A5" t="str"><v>Calc &amp; co</v></c><c r="B5" t="s"><v>1</v></c><c r="C5" t="s"><v>2</v></c><c r="D5" t="s"><v>5</v></c><c r="E5" t="s"><v>3</v></c><c r="F5" ${inl("no")}></c><c r="G5" ${inl("12/31/2026")}></c><c r="H5"/></row>
<row r="6"><c r="A6" ${inl("Bad amount row")}></c><c r="B6" t="s"><v>1</v></c><c r="C6" t="s"><v>2</v></c><c r="D6" t="s"><v>6</v></c><c r="E6" t="s"><v>3</v></c><c r="F6" ${inl("1")}></c><c r="G6"/><c r="H6"/></row>
<row r="7"/>
</sheetData></worksheet>`;

const happy = readXlsxGrid(buildXlsx({ sharedStringsXml: SHARED, sheetXml: SHEET }));
check("happy: no reader error", happy.error === undefined, happy.error);
eq("happy: date1904 off", happy.date1904, false);

const res: ParseResult = parseWorkbookGrid(happy.grid, { date1904: happy.date1904 });
check("happy: no parse error", res.error === undefined, res.error);
eq("happy: rowCount", res.rowCount, 4);
eq("happy: errorRowNumbers (blank + N/A amounts)", res.errorRowNumbers, [2, 4]);
eq("happy: amountSum", res.amountSum, 4654.56);
eq("happy: rich-text shared string + entity", res.rows[0].Description, "#16374 EMTP Basic & Co");
eq("happy: GLAccount whitespace collapsed", res.rows[0].GLAccount, "80-01-009 - Software Maintenance");
eq("happy: float artifact rounded", res.rows[0].Amount, 3420);
eq("happy: t=\"b\" Budgeted", res.rows[0].Budgeted, true);
eq("happy: serial StartDate", res.rows[0].StartDate, "2026-07-29");
eq("happy: text EndDate", res.rows[0].EndDate, "2026-03-07");
eq("happy: blank row skipped, RowNumber contiguous", res.rows.map((r) => r.RowNumber), [1, 2, 3, 4]);
eq("happy: sparse row description", res.rows[1].Description, "Sparse row");
eq("happy: sparse row blank amount -> HasError", res.rows[1].HasError, true);
eq("happy: t=\"e\" cell -> blank EndDate", res.rows[1].EndDate, "");
eq("happy: t=\"str\" formula cache", res.rows[2].Description, "Calc & co");
eq("happy: text amount with separators", res.rows[2].Amount, 1234.56);
eq("happy: Budgeted no -> false", res.rows[2].Budgeted, false);
eq("happy: Budgeted 1 -> true", res.rows[3].Budgeted, true);
eq("happy: ProjectNumber", res.rows[0].ProjectNumber, "IG0087.76");
eq("happy: BusinessGroup", res.rows[0].BusinessGroup, "NPG");

// ---- missing column ----
const noBudgeted = readXlsxGrid(
    buildXlsxFromRows([
        ["Line Description", "Project #", "GL Account Description", "Amount", "Business Group", "Start Date", "End Date"],
        ["x", "p", "gl", 1, "bg", "7/29/2026", "7/29/2027"],
    ])
);
eq("missing column message", parseWorkbookGrid(noBudgeted.grid).error, "Missing column: Budgeted");

// ---- empty sheet / header-only ----
eq("empty sheet", parseWorkbookGrid(readXlsxGrid(buildXlsx()).grid).error, "No data: the file is empty.");
const headerOnly = parseWorkbookGrid(
    readXlsxGrid(
        buildXlsxFromRows([
            ["Line Description", "Project #", "GL Account Description", "Amount", "Business Group", "Budgeted", "Start Date", "End Date"],
        ])
    ).grid
);
eq("header-only rowCount", headerOnly.rowCount, 0);
check("header-only has no structural error", headerOnly.error === undefined, headerOnly.error);

// ---- date1904 workbook ----
const wb1904 = readXlsxGrid(
    buildXlsxFromRows(
        [
            ["Line Description", "Project #", "GL Account Description", "Amount", "Business Group", "Budgeted", "Start Date", "End Date"],
            ["x", "p", "gl", 10, "bg", "Yes", 25569 - 1462, null],
        ],
        { date1904: true }
    )
);
eq("date1904 flag read", wb1904.date1904, true);
eq("date1904 serial shifted", parseWorkbookGrid(wb1904.grid, { date1904: wb1904.date1904 }).rows[0].StartDate, "1970-01-01");

// ---- rels resolution: first tab is rId2 with an ABSOLUTE target to sheet2 ----
const relsWb = buildXlsx({
    workbookXml: `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Data" sheetId="2" r:id="rId2"/><sheet name="Other" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    relsXml: `<Relationships>
<Relationship Id="rId1" Type="w" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="w" Target="/xl/worksheets/sheet2.xml"/>
</Relationships>`,
    sheetPath: "xl/worksheets/sheet2.xml",
    sheetXml: `<worksheet><sheetData><row r="1"><c r="A1" ${inl("from-sheet2")}></c></row></sheetData></worksheet>`,
});
eq("rels: absolute target resolves to sheet2", readXlsxGrid(relsWb).grid[0][0], "from-sheet2");

// ---- no rels file -> sheet1.xml fallback ----
const noRels = buildXlsx({
    omitRels: true,
    sheetXml: `<worksheet><sheetData><row r="1"><c r="A1" ${inl("fallback")}></c></row></sheetData></worksheet>`,
});
eq("no rels: sheet1 fallback", readXlsxGrid(noRels).grid[0][0], "fallback");

// ---- corrupt input ----
check("corrupt bytes -> error, no throw", readXlsxGrid(new Uint8Array([1, 2, 3])).error !== undefined);
check("empty buffer -> error, no throw", readXlsxGrid(new ArrayBuffer(0)).error !== undefined);

// ---- committed sample round-trip ----
const samplePath = path.join(__dirname, "../../../sample-data/sample-lineitems.xlsx");
if (fs.existsSync(samplePath)) {
    const sample = readXlsxGrid(fs.readFileSync(samplePath));
    check("sample: no reader error", sample.error === undefined, sample.error);
    const sres = parseWorkbookGrid(sample.grid, { date1904: sample.date1904 });
    eq("sample: rowCount", sres.rowCount, 20);
    eq("sample: errorRowNumbers", sres.errorRowNumbers, [14]);
    eq("sample: amountSum", sres.amountSum, 1034893.5);
    eq("sample: row1 GLAccount", sres.rows[0].GLAccount, "80-01-009 - Software Maintenance");
} else {
    console.log("  SKIP  sample round-trip (sample-data/sample-lineitems.xlsx not generated yet)");
}

console.log(failures === 0 ? "\nALL TESTS PASSED" : `\n${failures} TEST(S) FAILED`);
process.exit(failures === 0 ? 0 : 1);
