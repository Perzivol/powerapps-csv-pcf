/*
 * Minimal OOXML (.xlsx) worksheet reader for the CsvLineItemImport PCF control.
 *
 * Scope: extract the FIRST worksheet of a workbook as a grid of raw cell strings.
 * Uses fflate for unzipping only; the machine-generated workbook XML is scanned with
 * targeted regexes rather than a DOM parser so this module stays PURE (no DOM, no PCF
 * types) and runs under plain Node for the selftest.
 *
 * Deliberately NOT parsed: styles.xml. Date cells therefore surface as raw serial-number
 * strings; the row mapper (lineItemParser.ts) converts them because it knows which
 * columns are dates by header name.
 */

import { unzipSync, strFromU8 } from "fflate";

/** Result of extracting the first worksheet of a workbook. */
export interface XlsxGridResult {
    /** Row-major grid of raw cell strings; sparse/missing cells are "". */
    grid: string[][];
    /** true when workbook.xml declares the Mac 1904 date system. */
    date1904: boolean;
    /** Structural failure (corrupt zip / missing parts). When set, grid is []. */
    error?: string;
}

const NOT_A_WORKBOOK = "Could not read file: not a valid .xlsx workbook.";

/**
 * Decode the five named XML entities plus numeric character references, in a single
 * pass so already-decoded text cannot be double-decoded (e.g. "&amp;lt;" -> "&lt;").
 */
export function decodeXmlEntities(s: string): string {
    return s.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g, (whole, ent: string) => {
        switch (ent) {
            case "amp":
                return "&";
            case "lt":
                return "<";
            case "gt":
                return ">";
            case "quot":
                return '"';
            case "apos":
                return "'";
        }
        const code =
            ent[1] === "x" || ent[1] === "X" ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
        return Number.isFinite(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
    });
}

/** Convert a column reference ("A", "Z", "AA", ...) to a 0-based column index. */
export function colRefToIndex(letters: string): number {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
        n = n * 26 + (letters.charCodeAt(i) - 64); // "A" = 65
    }
    return n - 1;
}

/**
 * Concatenate all <t> text runs inside an <si> or <is> body (plain text or rich-text
 * runs). Content is taken verbatim — never trimmed — so xml:space="preserve" spacing
 * survives. <rPh> phonetic (furigana) blocks are stripped first so their runs do not
 * duplicate the visible text.
 */
function extractTextRuns(body: string): string {
    const clean = body.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
    let out = "";
    for (const m of clean.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g)) {
        out += m[1] !== undefined ? decodeXmlEntities(m[1]) : "";
    }
    return out;
}

/**
 * Find the zip path of the first worksheet (first <sheet> tab in workbook.xml, resolved
 * through the workbook rels). Falls back to sheet1.xml, then the lowest-numbered sheet.
 */
function findFirstSheetPath(files: Record<string, Uint8Array>, workbookXml: string): string | undefined {
    let sheetPath: string | undefined;

    const sheetTag = /<sheet\b[^>]*\/?>/.exec(workbookXml)?.[0] ?? "";
    const rid = /\br:id="([^"]+)"/.exec(sheetTag)?.[1];
    if (rid && files["xl/_rels/workbook.xml.rels"]) {
        const relsXml = strFromU8(files["xl/_rels/workbook.xml.rels"]);
        for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
            const tag = m[0];
            const id = /\bId="([^"]+)"/.exec(tag)?.[1];
            if (id !== rid) {
                continue;
            }
            const target = /\bTarget="([^"]+)"/.exec(tag)?.[1];
            if (target) {
                const decoded = decodeXmlEntities(target);
                sheetPath = decoded.startsWith("/")
                    ? decoded.slice(1) // absolute: "/xl/worksheets/sheet1.xml"
                    : "xl/" + decoded.replace(/^\.\//, ""); // relative to xl/
            }
            break;
        }
    }

    if (!sheetPath || !files[sheetPath]) {
        sheetPath = "xl/worksheets/sheet1.xml";
    }
    if (!files[sheetPath]) {
        sheetPath = Object.keys(files)
            .filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k))
            .sort()[0];
    }
    return sheetPath && files[sheetPath] ? sheetPath : undefined;
}

/** Parse xl/sharedStrings.xml (may be absent) into the shared-string table. */
function readSharedStrings(files: Record<string, Uint8Array>): string[] {
    const shared: string[] = [];
    if (files["xl/sharedStrings.xml"]) {
        const sstXml = strFromU8(files["xl/sharedStrings.xml"]);
        for (const m of sstXml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)) {
            shared.push(extractTextRuns(m[1]));
        }
    }
    return shared;
}

/** Extract one cell's raw string value from its attributes + inner XML. */
function readCellValue(attrs: string, inner: string, shared: string[]): string {
    const t = /\bt="([^"]+)"/.exec(attrs)?.[1] ?? "n";
    const v = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(inner)?.[1];

    if (t === "s") {
        // Shared-string index.
        const idx = parseInt(v ?? "", 10);
        return Number.isFinite(idx) ? (shared[idx] ?? "") : "";
    }
    if (t === "str") {
        // Cached formula result (already plain text in <v>).
        return decodeXmlEntities(v ?? "");
    }
    if (t === "inlineStr") {
        const is = /<is(?:\s[^>]*)?>([\s\S]*?)<\/is>/.exec(inner)?.[1] ?? "";
        return extractTextRuns(is);
    }
    if (t === "b") {
        return v === "1" ? "TRUE" : "FALSE";
    }
    if (t === "e") {
        // Cell error (#DIV/0! etc.) — treat as blank.
        return "";
    }
    // t="n" or untyped: raw numeric string (date serials included).
    return decodeXmlEntities(v ?? "");
}

/**
 * Read the first worksheet of an .xlsx workbook into a grid of raw cell strings.
 * Never throws; structural problems are reported via the `error` field.
 */
export function readXlsxGrid(data: ArrayBuffer | Uint8Array): XlsxGridResult {
    try {
        const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
        const files = unzipSync(bytes);

        const workbookXml = files["xl/workbook.xml"] ? strFromU8(files["xl/workbook.xml"]) : "";
        const date1904 = /<workbookPr\b[^>]*\bdate1904="(?:1|true)"/.test(workbookXml);

        const sheetPath = findFirstSheetPath(files, workbookXml);
        if (!sheetPath) {
            return { grid: [], date1904, error: "Could not read file: no worksheet found." };
        }

        const shared = readSharedStrings(files);
        const sheetXml = strFromU8(files[sheetPath]);

        const grid: string[][] = [];
        for (const rowMatch of sheetXml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
            const cells: string[] = [];
            // Cells are SPARSE: place by the r="C5" reference; a <c> without r (legal)
            // lands at the cursor position immediately after the previous cell.
            let cursor = 0;
            for (const cm of rowMatch[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
                const attrs = cm[1];
                const ref = /\br="([A-Za-z]+)\d+"/.exec(attrs);
                const col = ref ? colRefToIndex(ref[1].toUpperCase()) : cursor;
                cursor = col + 1;

                const value = readCellValue(attrs, cm[2] ?? "", shared);
                while (cells.length < col) {
                    cells.push("");
                }
                cells[col] = value;
            }
            grid.push(cells);
        }
        // Note: self-closing <row/> tags don't match above — equivalent to blank rows,
        // which the row mapper skips anyway.

        return { grid, date1904 };
    } catch {
        return { grid: [], date1904: false, error: NOT_A_WORKBOOK };
    }
}
