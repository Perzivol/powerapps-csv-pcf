import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { parseCsv, ParsedRow, ParseResult } from "./csvParser";

const DEFAULT_BUTTON_TEXT = "Import CSV";

export class CsvLineItemImport implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container!: HTMLDivElement;
    private _notifyOutputChanged!: () => void;

    private _root!: HTMLDivElement;
    private _button!: HTMLButtonElement;
    private _fileInput!: HTMLInputElement;
    private _summary!: HTMLDivElement;

    private _buttonText = DEFAULT_BUTTON_TEXT;

    // Output state — fully replaced on every successful (re)import.
    private _rows: ParsedRow[] = [];
    private _rowCount = 0;
    private _errorRowCount = 0;
    private _totalCostSum = 0;

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

        // Hidden native file picker, filtered to CSV.
        this._fileInput = document.createElement("input");
        this._fileInput.type = "file";
        this._fileInput.accept = ".csv,text/csv";
        this._fileInput.style.display = "none";
        this._fileInput.addEventListener("change", this.onFileChange);

        this._summary = document.createElement("div");
        this._summary.className = "csv-summary";
        this._summary.textContent = "No file selected";

        this._root.appendChild(this._button);
        this._root.appendChild(this._fileInput);
        this._root.appendChild(this._summary);
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
            const text = typeof reader.result === "string" ? reader.result : "";
            this.applyResult(parseCsv(text));
            // Clear the value so picking the SAME filename again re-fires "change".
            this._fileInput.value = "";
        };
        reader.onerror = (): void => {
            this.applyError("Could not read file.");
            this._fileInput.value = "";
        };

        try {
            reader.readAsText(file, "UTF-8");
        } catch {
            this.applyError("Could not read file.");
            this._fileInput.value = "";
        }
    };

    private applyResult(result: ParseResult): void {
        if (result.error) {
            // Structural problem (empty / missing column): clear data and show the message.
            this.resetOutputs();
            this.renderError(result.error);
        } else {
            this._rows = result.rows;
            this._rowCount = result.rowCount;
            this._errorRowCount = result.errorRowCount;
            this._totalCostSum = result.totalCostSum;
            this.renderSummary(result);
        }
        // Fire OnChange exactly once per import (success OR handled error).
        this._notifyOutputChanged();
    }

    private applyError(message: string): void {
        this.resetOutputs();
        this.renderError(message);
        this._notifyOutputChanged();
    }

    private resetOutputs(): void {
        this._rows = [];
        this._rowCount = 0;
        this._errorRowCount = 0;
        this._totalCostSum = 0;
    }

    private renderSummary(result: ParseResult): void {
        const money = new Intl.NumberFormat("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
        }).format(result.totalCostSum);

        const errorLine =
            result.errorRowCount > 0
                ? `Error rows: ${result.errorRowCount} (rows: ${result.errorRowNumbers.join(", ")})`
                : "Error rows: 0";

        this._summary.className = "csv-summary";
        this._summary.textContent = "";
        for (const line of [`Rows parsed: ${result.rowCount}`, `Total cost: ${money}`, errorLine]) {
            const div = document.createElement("div");
            div.textContent = line;
            this._summary.appendChild(div);
        }
    }

    private renderError(message: string): void {
        this._summary.className = "csv-summary csv-summary--error";
        this._summary.textContent = message;
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
                        ProductCode: { type: "string" },
                        Description: { type: "string" },
                        UnitCost: { type: "number" },
                        Quantity: { type: "number" },
                        StartDate: { type: "string" },
                        EndDate: { type: "string" },
                        TotalCost: { type: "number" },
                        SourceDocument: { type: "string" },
                        SourcePage: { type: "string" },
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
            TotalCostSum: this._totalCostSum,
            ParsedRowsJson: JSON.stringify(this._rows),
        };
    }

    public destroy(): void {
        this._button.removeEventListener("click", this.onButtonClick);
        this._fileInput.removeEventListener("change", this.onFileChange);
    }
}
