import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { parseWorkbookGrid, ParsedRow, ParseResult } from "./lineItemParser";
import { readXlsxGrid } from "./xlsxReader";

const DEFAULT_BUTTON_TEXT = "Import Excel";

export class CsvLineItemImport implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container!: HTMLDivElement;
    private _notifyOutputChanged!: () => void;

    private _root!: HTMLDivElement;
    private _button!: HTMLButtonElement;
    private _fileInput!: HTMLInputElement;

    private _buttonText = DEFAULT_BUTTON_TEXT;

    // Output state — fully replaced on every successful (re)import.
    private _rows: ParsedRow[] = [];
    private _rowCount = 0;
    private _errorRowCount = 0;
    private _amountSum = 0;
    private _errorMessage = "";
    private _errorRowNumbers = "";

    constructor() {
        // Empty
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;
        this._buttonText = context.parameters.ButtonText.raw || DEFAULT_BUTTON_TEXT;

        this._root = document.createElement("div");
        this._root.className = "csv-import";

        this._button = document.createElement("button");
        this._button.type = "button";
        this._button.className = "csv-import-btn";
        this._button.textContent = this._buttonText;
        this._button.addEventListener("click", this.onButtonClick);

        // Hidden native file picker, filtered to Excel workbooks.
        this._fileInput = document.createElement("input");
        this._fileInput.type = "file";
        this._fileInput.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        this._fileInput.style.display = "none";
        this._fileInput.addEventListener("change", this.onFileChange);

        this._root.appendChild(this._button);
        this._root.appendChild(this._fileInput);
        this._container.appendChild(this._root);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const label = context.parameters.ButtonText.raw || DEFAULT_BUTTON_TEXT;
        if (label !== this._buttonText) {
            this._buttonText = label;
            this._button.textContent = label;
        }

        // Honor the maker-allocated size when the host provides it.
        const w = context.mode.allocatedWidth;
        const h = context.mode.allocatedHeight;
        if (typeof w === "number" && w > 0) {
            this._root.style.width = `${w}px`;
        }
        if (typeof h === "number" && h > 0) {
            this._root.style.minHeight = `${h}px`;
        }
    }

    private onButtonClick = (): void => {
        this._fileInput.click();
    };

    private onFileChange = (): void => {
        const file = this._fileInput.files && this._fileInput.files[0];
        if (!file) {
            return;
        }

        const reader = new FileReader();
        reader.onload = (): void => {
            const buf = reader.result instanceof ArrayBuffer ? reader.result : new ArrayBuffer(0);
            const workbook = readXlsxGrid(buf);
            if (workbook.error) {
                this.applyError(workbook.error);
            } else {
                this.applyResult(parseWorkbookGrid(workbook.grid, { date1904: workbook.date1904 }));
            }
            // Clear the value so picking the SAME filename again re-fires "change".
            this._fileInput.value = "";
        };
        reader.onerror = (): void => {
            this.applyError("Could not read file.");
            this._fileInput.value = "";
        };

        try {
            reader.readAsArrayBuffer(file);
        } catch {
            this.applyError("Could not read file.");
            this._fileInput.value = "";
        }
    };

    private applyResult(result: ParseResult): void {
        if (result.error) {
            // Structural problem (empty / missing column): clear data and surface the message.
            this.resetOutputs();
            this._errorMessage = result.error;
        } else {
            this._rows = result.rows;
            this._rowCount = result.rowCount;
            this._errorRowCount = result.errorRowCount;
            this._amountSum = result.amountSum;
            this._errorMessage = "";
            this._errorRowNumbers = result.errorRowNumbers.join(", ");
        }
        // Fire OnChange exactly once per import (success OR handled error).
        this._notifyOutputChanged();
    }

    private applyError(message: string): void {
        this.resetOutputs();
        this._errorMessage = message;
        this._notifyOutputChanged();
    }

    private resetOutputs(): void {
        this._rows = [];
        this._rowCount = 0;
        this._errorRowCount = 0;
        this._amountSum = 0;
        this._errorMessage = "";
        this._errorRowNumbers = "";
    }

    /**
     * Provides the schema for the object-typed ParsedRows output. Returning a top-level
     * `array` schema makes ParsedRows surface directly as a typed Table in canvas Power Fx.
     * Called by the platform before control initialization.
     */
    public getOutputSchema(context: ComponentFramework.Context<IInputs>): Promise<Record<string, unknown>> {
        return Promise.resolve({
            ParsedRows: {
                $schema: "http://json-schema.org/draft-04/schema#",
                type: "array",
                items: {
                    type: "object",
                    properties: {
                        Description: { type: "string" },
                        ProjectNumber: { type: "string" },
                        GLAccount: { type: "string" },
                        Amount: { type: "number" },
                        BusinessGroup: { type: "string" },
                        Budgeted: { type: "boolean" },
                        StartDate: { type: "string" },
                        EndDate: { type: "string" },
                        RowNumber: { type: "integer" },
                        HasError: { type: "boolean" },
                    },
                },
            },
        });
    }

    public getOutputs(): IOutputs {
        return {
            ParsedRows: this._rows,
            RowCount: this._rowCount,
            ErrorRowCount: this._errorRowCount,
            AmountSum: this._amountSum,
            ErrorMessage: this._errorMessage,
            ErrorRowNumbers: this._errorRowNumbers,
            ParsedRowsJson: JSON.stringify(this._rows),
        };
    }

    public destroy(): void {
        this._button.removeEventListener("click", this.onButtonClick);
        this._fileInput.removeEventListener("change", this.onFileChange);
    }
}
