/*
 * RFC-4180-style CSV parser and line-item normalizers for the CsvLineItemImport PCF control.
 *
 * This module is intentionally PURE: no DOM, no PCF types, no third-party dependencies.
 * That keeps the bundle tiny and lets the parsing logic be unit-tested in isolation
 * (see test/csvParser.selftest.ts).
 */

/** One parsed line-item row. Property names/casing match the ParsedRows output schema. */
export interface ParsedRow {
    ProductCode: string;
    Description: string;
    UnitCost: number;
    Quantity: number;
    /** ISO yyyy-mm-dd, or "" when the source cell is blank/unparseable. */
    StartDate: string;
    /** ISO yyyy-mm-dd, or "" when the source cell is blank/unparseable. */
    EndDate: string;
    TotalCost: number;
    SourceDocument: string;
    SourcePage: string;
    /** 1-based index within the parsed data rows. */
    RowNumber: number;
    /** true if any numeric field (unit_cost / quantity / total_cost) failed to parse. */
    HasError: boolean;
}

/** Result of parsing an entire CSV document. */
export interface ParseResult {
    rows: ParsedRow[];
    rowCount: number;
    errorRowCount: number;
    errorRowNumbers: number[];
    totalCostSum: number;
    /**
     * Set when the file is structurally unusable (empty file, or a required column is
     * missing). When present, `rows` is empty and callers should surface `error` to the user.
     */
    error?: string;
}

/** Exact, lowercase snake_case headers the source file must contain (matched by name, not position). */
export const REQUIRED_COLUMNS: readonly string[] = [
    "product_code",
    "product_description",
    "unit_cost",
    "quantity",
    "coverage_start_date",
    "coverage_end_date",
    "total_cost",
    "source_document",
    "source_page",
];

/** Strip a leading UTF-8 BOM (U+FEFF) if present. */
export function stripBom(text: string): string {
    return text.length > 0 && text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/**
 * Tokenize CSV text into a grid of rows/fields per RFC 4180.
 * Handles: double-quoted fields, "" escaped quotes, commas and newlines inside quotes,
 * CRLF and bare LF (and lone CR) line endings, and a single trailing newline
 * (no phantom empty final row).
 */
export function tokenize(input: string): string[][] {
    const rows: string[][] = [];
    let field = "";
    let row: string[] = [];
    let inQuotes = false;
    let i = 0;
    const n = input.length;

    const endField = (): void => {
        row.push(field);
        field = "";
    };
    const endRow = (): void => {
        endField();
        rows.push(row);
        row = [];
    };

    while (i < n) {
        const c = input[i];

        if (inQuotes) {
            if (c === '"') {
                if (input[i + 1] === '"') {
                    field += '"';
                    i += 2;
                    continue;
                }
                inQuotes = false;
                i++;
                continue;
            }
            field += c;
            i++;
            continue;
        }

        if (c === '"') {
            inQuotes = true;
            i++;
            continue;
        }
        if (c === ",") {
            endField();
            i++;
            continue;
        }
        if (c === "\r") {
            endRow();
            i += input[i + 1] === "\n" ? 2 : 1;
            continue;
        }
        if (c === "\n") {
            endRow();
            i++;
            continue;
        }

        field += c;
        i++;
    }

    // Flush the final field/row, unless the input ended exactly on a row break
    // (in which case `field` is empty and `row` is empty -> nothing to flush).
    if (field.length > 0 || row.length > 0) {
        endRow();
    }

    return rows;
}

/**
 * Normalize a numeric cell: trim -> remove thousands separators -> parseFloat.
 * Returns ok:false (and value 0) when the result is NaN, per spec.
 */
export function normalizeNumber(raw: string): { value: number; ok: boolean } {
    const cleaned = (raw ?? "").trim().replace(/,/g, "");
    if (cleaned === "") {
        return { value: 0, ok: false };
    }
    const v = parseFloat(cleaned);
    return Number.isNaN(v) ? { value: 0, ok: false } : { value: v, ok: true };
}

/**
 * Parse an M/D/YYYY date (single- or double-digit month/day) into ISO yyyy-mm-dd.
 * Blank or unparseable input returns "" (a missing/bad date does NOT, by itself,
 * flag the row as an error).
 */
export function parseMdyToIso(raw: string): string {
    const s = (raw ?? "").trim();
    if (s === "") {
        return "";
    }
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
    if (!m) {
        return "";
    }
    const month = parseInt(m[1], 10);
    const day = parseInt(m[2], 10);
    const year = parseInt(m[3], 10);
    if (month < 1 || month > 12 || day < 1 || day > 31) {
        return "";
    }
    // Reject impossible calendar dates (e.g. 2/30) by round-tripping through Date.
    const d = new Date(year, month - 1, day);
    if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) {
        return "";
    }
    const mm = String(month).padStart(2, "0");
    const dd = String(day).padStart(2, "0");
    return `${year}-${mm}-${dd}`;
}

/**
 * Parse an entire CSV document into typed rows and summary totals.
 * Never throws; structural problems are reported via the `error` field.
 */
export function parseCsv(text: string): ParseResult {
    const empty: ParseResult = {
        rows: [],
        rowCount: 0,
        errorRowCount: 0,
        errorRowNumbers: [],
        totalCostSum: 0,
    };

    if (text == null || stripBom(text).trim() === "") {
        return { ...empty, error: "No data: the file is empty." };
    }

    const grid = tokenize(stripBom(text));
    if (grid.length === 0) {
        return { ...empty, error: "No data: the file is empty." };
    }

    // Header row: trim each name, build a name -> column-index map (first wins on dupes).
    const header = grid[0].map((h) => h.trim());
    const index: Record<string, number> = {};
    header.forEach((h, i) => {
        if (!(h in index)) {
            index[h] = i;
        }
    });

    const missing = REQUIRED_COLUMNS.filter((c) => !(c in index));
    if (missing.length > 0) {
        return { ...empty, error: `Missing column: ${missing.join(", ")}` };
    }

    const cell = (cols: string[], name: string): string => {
        const i = index[name];
        return i < cols.length ? cols[i] : "";
    };

    const rows: ParsedRow[] = [];
    const errorRowNumbers: number[] = [];
    let totalCostSum = 0;

    for (let r = 1; r < grid.length; r++) {
        const cols = grid[r];
        // Ignore a completely blank line so stray blank rows are not counted as data.
        if (cols.length === 1 && cols[0].trim() === "") {
            continue;
        }

        const rowNumber = rows.length + 1;
        const unit = normalizeNumber(cell(cols, "unit_cost"));
        const qty = normalizeNumber(cell(cols, "quantity"));
        const total = normalizeNumber(cell(cols, "total_cost"));
        const hasError = !unit.ok || !qty.ok || !total.ok;

        rows.push({
            ProductCode: cell(cols, "product_code").trim(),
            Description: cell(cols, "product_description").trim(),
            UnitCost: unit.value,
            Quantity: qty.value,
            StartDate: parseMdyToIso(cell(cols, "coverage_start_date")),
            EndDate: parseMdyToIso(cell(cols, "coverage_end_date")),
            TotalCost: total.value,
            SourceDocument: cell(cols, "source_document").trim(),
            SourcePage: cell(cols, "source_page").trim(),
            RowNumber: rowNumber,
            HasError: hasError,
        });

        totalCostSum += total.value;
        if (hasError) {
            errorRowNumbers.push(rowNumber);
        }
    }

    return {
        rows,
        rowCount: rows.length,
        errorRowCount: errorRowNumbers.length,
        errorRowNumbers,
        totalCostSum,
    };
}
