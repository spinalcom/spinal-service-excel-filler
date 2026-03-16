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
 * Map of cell references to values.
 * Keys use the format "SheetName!CellRef" (e.g. "Sheet1!B3").
 * Values can be plain primitives or a CellEntry for conditional coloring.
 */
export type CellMap = Record<string, ExcelJS.CellValue | CellEntry>;
export interface FillOptions {
    /** If true, returns a Buffer instead of writing to a file. */
    asBuffer?: boolean;
}
export interface SpinalExcelFillerConfig {
    /** Optional default background color (hex without #) to apply to all filled cells. */
    defaultColor?: string;
}
export declare class SpinalExcelFiller {
    private workbook;
    private config;
    constructor(config?: SpinalExcelFillerConfig);
    /**
     * Load an Excel template from a file path.
     */
    loadTemplate(filePath: string): Promise<void>;
    /**
     * Load an Excel template from a Buffer.
     */
    loadTemplateFromBuffer(buffer: Buffer): Promise<void>;
    /**
     * Fill cells in the loaded template.
     *
     * @example
     * // Plain values
     * filler.setCells({ "Sheet1!B3": "Hello", "Sheet1!C5": 42 });
     *
     * // With conditional coloring
     * filler.setCells({
     *   "Sheet1!B3": {
     *     value: 85,
     *     color: (v) => (v as number) >= 80 ? "4CAF50" : "F44336",
     *   },
     * });
     */
    setCells(cells: CellMap): void;
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