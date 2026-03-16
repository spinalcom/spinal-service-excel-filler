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

function isCellEntry(v: unknown): v is CellEntry {
  return typeof v === "object" && v !== null && "value" in v;
}

function parseCellRef(ref: string): { sheet: string; cell: string } {
  const idx = ref.indexOf("!");
  if (idx === -1) {
    throw new Error(
      `Invalid cell reference "${ref}". Expected format: "SheetName!CellRef" (e.g. "Sheet1!B3").`
    );
  }
  return { sheet: ref.slice(0, idx), cell: ref.slice(idx + 1) };
}

function applyFill(cell: ExcelJS.Cell, hex: string): void {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: `FF${hex.replace(/^#/, "")}` },
  };
}

export class SpinalExcelFiller {
  private workbook: ExcelJS.Workbook | null = null;
  private config: SpinalExcelFillerConfig;

  constructor(config: SpinalExcelFillerConfig = {}) {
    this.config = config;
  }

  /**
   * Load an Excel template from a file path.
   */
  async loadTemplate(filePath: string): Promise<void> {
    this.workbook = new ExcelJS.Workbook();
    await this.workbook.xlsx.readFile(filePath);
  }

  /**
   * Load an Excel template from a Buffer.
   */
  async loadTemplateFromBuffer(buffer: Buffer): Promise<void> {
    this.workbook = new ExcelJS.Workbook();
    await this.workbook.xlsx.load(
      buffer.buffer.slice(
        buffer.byteOffset,
        buffer.byteOffset + buffer.byteLength
      ) as ArrayBuffer
    );
  }

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
  setCells(cells: CellMap): void {
    if (!this.workbook) {
      throw new Error("No template loaded. Call loadTemplate() first.");
    }

    for (const [ref, entry] of Object.entries(cells)) {
      const { sheet, cell: cellRef } = parseCellRef(ref);
      const ws = this.workbook.getWorksheet(sheet);
      if (!ws) {
        throw new Error(`Worksheet "${sheet}" not found in the template.`);
      }

      const cell = ws.getCell(cellRef);

      if (isCellEntry(entry)) {
        cell.value = entry.value;

        let color: string | undefined;
        if (typeof entry.color === "function") {
          color = entry.color(entry.value);
        } else if (typeof entry.color === "string") {
          color = entry.color;
        }

        if (color) {
          applyFill(cell, color);
        } else if (this.config.defaultColor) {
          applyFill(cell, this.config.defaultColor);
        }
      } else {
        cell.value = entry;
        if (this.config.defaultColor) {
          applyFill(cell, this.config.defaultColor);
        }
      }
    }
  }

  /**
   * Save the filled workbook to a file.
   */
  async save(outputPath: string): Promise<void> {
    if (!this.workbook) {
      throw new Error("No template loaded. Call loadTemplate() first.");
    }
    await this.workbook.xlsx.writeFile(outputPath);
  }

  /**
   * Return the filled workbook as a Buffer.
   */
  async toBuffer(): Promise<Buffer> {
    if (!this.workbook) {
      throw new Error("No template loaded. Call loadTemplate() first.");
    }
    const arrayBuffer = await this.workbook.xlsx.writeBuffer();
    return Buffer.from(arrayBuffer);
  }
}
