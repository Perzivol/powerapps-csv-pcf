/*
 * In-memory .xlsx builder for the selftest and the sample-data generator.
 * Hand-written OOXML parts zipped with fflate. Includes [Content_Types].xml and the
 * package rels so generated files also open in real Excel (the reader ignores them).
 */
import { zipSync, strToU8 } from "fflate";

export interface XlsxParts {
    workbookXml?: string;
    relsXml?: string;
    /** null = omit xl/sharedStrings.xml entirely. */
    sharedStringsXml?: string | null;
    sheetXml?: string;
    /** Zip path of the sheet part. Default "xl/worksheets/sheet1.xml". */
    sheetPath?: string;
    /** null = omit xl/_rels/workbook.xml.rels entirely. */
    omitRels?: boolean;
}

const DEFAULT_WORKBOOK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const DEFAULT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

/** Zip the workbook parts (with sensible defaults) into an in-memory .xlsx. */
export function buildXlsx(parts: XlsxParts = {}): Uint8Array {
    const sheetPath = parts.sheetPath ?? "xl/worksheets/sheet1.xml";
    const files: Record<string, Uint8Array> = {
        "[Content_Types].xml": strToU8(CONTENT_TYPES),
        "_rels/.rels": strToU8(PACKAGE_RELS),
        "xl/workbook.xml": strToU8(parts.workbookXml ?? DEFAULT_WORKBOOK),
        [sheetPath]: strToU8(parts.sheetXml ?? "<worksheet><sheetData/></worksheet>"),
    };
    if (!parts.omitRels) {
        files["xl/_rels/workbook.xml.rels"] = strToU8(parts.relsXml ?? DEFAULT_RELS);
    }
    if (parts.sharedStringsXml !== null) {
        files["xl/sharedStrings.xml"] = strToU8(
            parts.sharedStringsXml ??
                `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>`
        );
    }
    return zipSync(files);
}

/** XML-escape a text value for embedding in hand-built sheet/sharedStrings XML. */
export function xmlEscape(s: string): string {
    return s
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

/** Convert a 0-based column index to a column reference ("A", "Z", "AA", ...). */
export function indexToColRef(index: number): string {
    let n = index + 1;
    let ref = "";
    while (n > 0) {
        const rem = (n - 1) % 26;
        ref = String.fromCharCode(65 + rem) + ref;
        n = Math.floor((n - 1) / 26);
    }
    return ref;
}

/**
 * Build a complete workbook from a simple value grid: strings become inline-string
 * cells, numbers numeric cells, null blank (cell omitted -> sparse). Row/cell refs are
 * emitted like real Excel output. Convenient for realistic fixtures; tests that need
 * exact control over shared strings / rich runs hand-write the XML instead.
 */
export function buildXlsxFromRows(rows: (string | number | null)[][], opts: { date1904?: boolean } = {}): Uint8Array {
    const rowsXml = rows
        .map((cells, r) => {
            const cellsXml = cells
                .map((value, c) => {
                    if (value === null) {
                        return "";
                    }
                    const ref = `${indexToColRef(c)}${r + 1}`;
                    return typeof value === "number"
                        ? `<c r="${ref}"><v>${value}</v></c>`
                        : `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`;
                })
                .join("");
            return `<row r="${r + 1}">${cellsXml}</row>`;
        })
        .join("");

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${opts.date1904 ? '<workbookPr date1904="1"/>' : ""}<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

    return buildXlsx({
        workbookXml,
        sheetXml: `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowsXml}</sheetData></worksheet>`,
    });
}
