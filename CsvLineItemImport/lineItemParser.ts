/*
 * Line-item row mapper and normalizers for the CsvLineItemImport PCF control.
 *
 * This module is intentionally PURE: no DOM, no PCF types, no third-party dependencies
 * (the fflate-based workbook extraction lives in xlsxReader.ts). It maps a raw cell grid
 * to typed rows and is unit-tested in isolation (see test/xlsxImport.selftest.ts).
 */

/** One parsed line-item row. Property names/casing match the ParsedRows output schema. */
export interface ParsedRow {
    /** "Line Description", trimmed. */
    Description: string;
    /** "Project #", trimmed. */
    ProjectNumber: string;
    /** "GL Account Description", trimmed + internal whitespace collapsed so the text
     *  matches Dataverse choice labels exactly (e.g. "80-01-009 - Software Maintenance"). */
    GLAccount: string;
    /** "Amount", rounded to 2 decimals; 0 (+ HasError) when unparseable. */
    Amount: number;
    /** "Business Group", trimmed. */
    BusinessGroup: string;
    /** "Budgeted": yes/true/1 (case-insensitive) -> true, anything else -> false. */
    Budgeted: boolean;
    /** ISO yyyy-mm-dd, or "" when the source cell is blank/unparseable. */
    StartDate: string;
    /** ISO yyyy-mm-dd, or "" when the source cell is blank/unparseable. */
    EndDate: string;
    /** 1-based index within the parsed data rows. */
    RowNumber: number;
    /** true if the Amount cell failed to parse. */
    HasError: boolean;
}

/** Result of parsing an entire worksheet grid. */
export interface ParseResult {
    rows: ParsedRow[];
    rowCount: number;
    errorRowCount: number;
    errorRowNumbers: number[];
    amountSum: number;
    /**
     * Set when the sheet is structurally unusable (empty, or a required column is
     * missing). When present, `rows` is empty and callers should surface `error` to the user.
     */
    error?: string;
}

/** Display names of the required columns, as they appear in the source file. */
export const REQUIRED_COLUMN_LABELS: readonly string[] = [
    "Line Description",
    "Project #",
    "GL Account Description",
    "Amount",
    "Business Group",
    "Budgeted",
    "Start Date",
    "End Date",
];

/** Lowercase trimmed/collapsed header -> position in REQUIRED_COLUMN_LABELS. */
export const REQUIRED_COLUMNS: readonly string[] = REQUIRED_COLUMN_LABELS.map((l) => l.toLowerCase());

/** Normalize a header cell for matching: trim, collapse whitespace, lowercase. */
export function normalizeHeader(raw: string): string {
    return (raw ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}

/** Trim + collapse internal whitespace (used for GLAccount choice-label matching). */
export function collapseWhitespace(raw: string): string {
    return (raw ?? "").trim().replace(/\s+/g, " ");
}

/**
 * Normalize a numeric cell: trim -> remove thousands separators -> parseFloat.
 * Returns ok:false (and value 0) when the result is NaN.
 */
export function normalizeNumber(raw: string): { value: number; ok: boolean } {
    const cleaned = (raw ?? "").trim().replace(/,/g, "");
    if (cleaned === "") {
        return { value: 0, ok: false };
    }
    const v = parseFloat(cleaned);
    return Number.isNaN(v) ? { value: 0, ok: false } : { value: v, ok: true };
}

/** Round a currency amount to 2 decimals (strips Excel binary-float artifacts). */
function round2(v: number): number {
    return Math.round(v * 100) / 100;
}

/** "Budgeted" cell -> boolean. Covers "Yes", "TRUE" (xlsx t="b" cells), "1". */
export function parseBoolean(raw: string): boolean {
    return /^(yes|true|1)$/i.test((raw ?? "").trim());
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
 * Excel serial day -> ISO yyyy-mm-dd. The 1900 system epoch is 1899-12-30, which
 * absorbs the Lotus leap-year bug for serials >= 61 (pre-1900 dates are out of scope
 * for this control). The 1904 system epoch is 1904-01-01. Any time-of-day fraction
 * is floored away. All math in UTC to avoid local-timezone day shifts.
 */
export function excelSerialToIso(serial: number, date1904: boolean): string {
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    const d = new Date(epoch + Math.floor(serial) * 86400000);
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const dd = String(d.getUTCDate()).padStart(2, "0");
    return `${d.getUTCFullYear()}-${mm}-${dd}`;
}

/**
 * Date-column cell -> ISO text. A purely-numeric value is an Excel date serial
 * (styles.xml is never consulted — the caller knows the column is a date by header);
 * yyyy-mm-dd text passes through; M/D/YYYY text is converted; anything else -> "".
 */
export function parseDateCell(raw: string, date1904: boolean): string {
    const s = (raw ?? "").trim();
    if (s === "") {
        return "";
    }
    if (/^\d+(\.\d+)?$/.test(s)) {
        const n = parseFloat(s);
        return n > 0 ? excelSerialToIso(n, date1904) : "";
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        return s;
    }
    return parseMdyToIso(s);
}

/**
 * Map a worksheet grid (from xlsxReader.readXlsxGrid) into typed rows and summary
 * totals. Never throws; structural problems are reported via the `error` field.
 */
export function parseWorkbookGrid(grid: string[][], opts?: { date1904?: boolean }): ParseResult {
    const empty: ParseResult = {
        rows: [],
        rowCount: 0,
        errorRowCount: 0,
        errorRowNumbers: [],
        amountSum: 0,
    };
    const date1904 = opts?.date1904 ?? false;
    const isBlankRow = (cols: string[]): boolean => cols.every((c) => c.trim() === "");

    // Header = first non-blank row (leading blank rows are tolerated).
    let headerIdx = 0;
    while (headerIdx < grid.length && isBlankRow(grid[headerIdx])) {
        headerIdx++;
    }
    if (headerIdx >= grid.length) {
        return { ...empty, error: "No data: the file is empty." };
    }

    // Build a normalized-name -> column-index map (first wins on dupes).
    const index: Record<string, number> = {};
    grid[headerIdx].forEach((h, i) => {
        const name = normalizeHeader(h);
        if (name !== "" && !(name in index)) {
            index[name] = i;
        }
    });

    const missing = REQUIRED_COLUMN_LABELS.filter((label) => !(label.toLowerCase() in index));
    if (missing.length > 0) {
        return { ...empty, error: `Missing column: ${missing.join(", ")}` };
    }

    const cell = (cols: string[], name: string): string => {
        const i = index[name];
        return i < cols.length ? cols[i] : "";
    };

    const rows: ParsedRow[] = [];
    const errorRowNumbers: number[] = [];
    let amountSum = 0;

    for (let r = headerIdx + 1; r < grid.length; r++) {
        const cols = grid[r];
        // Ignore completely blank rows so stray empty lines are not counted as data.
        if (isBlankRow(cols)) {
            continue;
        }

        const rowNumber = rows.length + 1;
        const amount = normalizeNumber(cell(cols, "amount"));
        const hasError = !amount.ok;
        const amountValue = round2(amount.value);

        rows.push({
            Description: cell(cols, "line description").trim(),
            ProjectNumber: cell(cols, "project #").trim(),
            GLAccount: collapseWhitespace(cell(cols, "gl account description")),
            Amount: amountValue,
            BusinessGroup: cell(cols, "business group").trim(),
            Budgeted: parseBoolean(cell(cols, "budgeted")),
            StartDate: parseDateCell(cell(cols, "start date"), date1904),
            EndDate: parseDateCell(cell(cols, "end date"), date1904),
            RowNumber: rowNumber,
            HasError: hasError,
        });

        amountSum += amountValue;
        if (hasError) {
            errorRowNumbers.push(rowNumber);
        }
    }

    return {
        rows,
        rowCount: rows.length,
        errorRowCount: errorRowNumbers.length,
        errorRowNumbers,
        amountSum: round2(amountSum),
    };
}
