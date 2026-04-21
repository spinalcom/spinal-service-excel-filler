import ExcelJS from "exceljs";
/**
 * A cell value with optional conditional coloring.
 */
export interface CellEntry {
    value: ExcelJS.CellValue;
    /**
     * Either a static hex color (e.g. "4CAF50")
     * or a function that receives the value and returns a hex color.
     */
    color?: string | ((value: ExcelJS.CellValue) => string | undefined);
}
/**
 * A cell value, either raw or wrapped in a CellEntry for conditional coloring.
 */
export type CellValueOrEntry = ExcelJS.CellValue | CellEntry;
/**
 * Map of cell references to values.
 * Keys use the format "SheetName!CellRef" (e.g. "Sheet1!B3").
 */
export type CellMap = Record<string, CellValueOrEntry>;
export interface FillOptions {
    /** If true, returns a Buffer instead of writing to a file. */
    asBuffer?: boolean;
}
export interface SpinalExcelFillerConfig {
    /** Optional default background color (hex without #) to apply to all filled cells. */
    defaultColor?: string;
}
export type RangeDirection = "row" | "column";
export interface SetRangeOptions {
    /** Axis for 1D arrays. Ignored for 2D arrays. Default: "column" (fills downward). */
    direction?: RangeDirection;
}
/**
 * A value assigned to a template variable.
 * Arrays are filled from the variable's anchor cell (downward by default) and
 * only work when the cell contains the token by itself.
 */
export type VariableValue = CellValueOrEntry | CellValueOrEntry[];
export type VariableMap = Record<string, VariableValue>;
export declare class SpinalExcelFiller {
    private workbook;
    private config;
    private templateCells;
    private variableIndex;
    private resolvedScalars;
    constructor(config?: SpinalExcelFillerConfig);
    /**
     * Load an Excel template from a file path.
     */
    loadTemplate(filePath: string): Promise<void>;
    /**
     * Load an Excel template from a Buffer.
     */
    loadTemplateFromBuffer(buffer: Buffer): Promise<void>;
    private scanVariables;
    /**
     * Returns the list of variable names discovered in the template
     * (tokens written as `{{name}}`).
     */
    getVariables(): string[];
    /**
     * Returns the cell references where each variable appears.
     */
    getVariableLocations(): Record<string, string[]>;
    /**
     * Assign values to template variables by name.
     *
     * - If a cell contains the token by itself (`{{name}}`), the raw value
     *   is written, preserving its type (number, Date, etc.) and any
     *   CellEntry coloring.
     * - If the token is embedded in text (`"Hello {{name}}"`), the value
     *   is stringified and substituted.
     * - Array values are supported only on sole-token cells and are filled
     *   downward from the anchor.
     *
     * May be called multiple times; previously assigned scalars are remembered
     * so later calls to a cell with multiple tokens still resolve cleanly.
     */
    setVariables(vars: VariableMap): void;
    /**
     * Fill cells in the loaded template.
     *
     * @example
     * filler.setCells({ "Sheet1!B3": "Hello", "Sheet1!C5": 42 });
     *
     * @example
     * filler.setCells({
     *   "Sheet1!B3": {
     *     value: 85,
     *     color: (v) => (v as number) >= 80 ? "4CAF50" : "F44336",
     *   },
     * });
     */
    setCells(cells: CellMap): void;
    /**
     * Fill a range of cells from an array.
     *
     * @param anchor single-cell reference like "Sheet1!B3"
     * @param values 1D array (uses `options.direction`) or 2D array (row-major block)
     *
     * @example
     * // Column downward from B3
     * filler.setRange("Sheet1!B3", [1, 2, 3]);
     *
     * @example
     * // Row rightward from B3
     * filler.setRange("Sheet1!B3", ["a", "b", "c"], { direction: "row" });
     *
     * @example
     * // 2D block
     * filler.setRange("Sheet1!B3", [
     *   ["name", "score"],
     *   ["Alice", 92],
     *   ["Bob",   71],
     * ]);
     */
    setRange(anchor: string, values: CellValueOrEntry[] | CellValueOrEntry[][], options?: SetRangeOptions): void;
    private writeCell;
    /**
     * Save the filled workbook to a file.
     */
    save(outputPath: string): Promise<void>;
    /**
     * Return the filled workbook as a Buffer.
     */
    toBuffer(): Promise<Buffer>;
}
//# sourceMappingURL=SpinalExcelFiller.d.ts.map