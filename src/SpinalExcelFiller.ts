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

const TOKEN_RE = /\{\{\s*([A-Za-z_][\w.]*)\s*\}\}/g;
const SOLE_TOKEN_RE = /^\s*\{\{\s*([A-Za-z_][\w.]*)\s*\}\}\s*$/;

function isCellEntry(v: unknown): v is CellEntry {
  return (
    typeof v === "object" &&
    v !== null &&
    !Array.isArray(v) &&
    "value" in v
  );
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

function a1ToRowCol(a1: string): { row: number; col: number } {
  const m = /^([A-Za-z]+)(\d+)$/.exec(a1);
  if (!m) throw new Error(`Invalid A1 cell reference "${a1}".`);
  const letters = m[1].toUpperCase();
  const rowStr = m[2];
  let col = 0;
  for (const ch of letters) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(rowStr, 10), col };
}

function applyFill(cell: ExcelJS.Cell, hex: string): void {
  // Assign a full new style object to avoid mutating a shared internal style reference,
  // which causes ExcelJS to apply the wrong color to cells that share the same style.
  cell.style = {
    ...cell.style,
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: `FF${hex.replace(/^#/, "")}` },
    },
  };
}

function readPlainString(value: ExcelJS.CellValue): string | null {
  if (typeof value === "string") return value;
  if (value && typeof value === "object" && "richText" in value) {
    const runs = (value as { richText: { text: string }[] }).richText;
    return runs.map((r) => r.text).join("");
  }
  return null;
}

interface TemplateCell {
  sheet: string;
  cell: string;
  template: string;
  tokens: string[];
  soleToken: string | null;
}

export class SpinalExcelFiller {
  private workbook: ExcelJS.Workbook | null = null;
  private config: SpinalExcelFillerConfig;

  private templateCells = new Map<string, TemplateCell>();
  private variableIndex = new Map<string, Set<string>>();
  private resolvedScalars = new Map<string, string>();

  constructor(config: SpinalExcelFillerConfig = {}) {
    this.config = config;
  }

  /**
   * Load an Excel template from a file path.
   */
  async loadTemplate(filePath: string): Promise<void> {
    this.workbook = new ExcelJS.Workbook();
    await this.workbook.xlsx.readFile(filePath);
    this.scanVariables();
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
    this.scanVariables();
  }

  private scanVariables(): void {
    this.templateCells.clear();
    this.variableIndex.clear();
    this.resolvedScalars.clear();
    if (!this.workbook) return;

    this.workbook.eachSheet((ws) => {
      ws.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          const text = readPlainString(cell.value);
          if (!text) return;

          TOKEN_RE.lastIndex = 0;
          const tokens: string[] = [];
          let m: RegExpExecArray | null;
          while ((m = TOKEN_RE.exec(text)) !== null) tokens.push(m[1]);
          if (tokens.length === 0) return;

          const soleMatch = SOLE_TOKEN_RE.exec(text);
          const key = `${ws.name}!${cell.address}`;
          this.templateCells.set(key, {
            sheet: ws.name,
            cell: cell.address,
            template: text,
            tokens,
            soleToken: soleMatch ? soleMatch[1] : null,
          });

          for (const name of tokens) {
            let set = this.variableIndex.get(name);
            if (!set) {
              set = new Set();
              this.variableIndex.set(name, set);
            }
            set.add(key);
          }
        });
      });
    });
  }

  /**
   * Returns the list of variable names discovered in the template
   * (tokens written as `{{name}}`).
   */
  getVariables(): string[] {
    return Array.from(this.variableIndex.keys());
  }

  /**
   * Returns the cell references where each variable appears.
   */
  getVariableLocations(): Record<string, string[]> {
    const out: Record<string, string[]> = {};
    for (const [name, refs] of this.variableIndex) {
      out[name] = Array.from(refs);
    }
    return out;
  }

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
  setVariables(vars: VariableMap): void {
    if (!this.workbook) {
      throw new Error("No template loaded. Call loadTemplate() first.");
    }

    for (const [name, val] of Object.entries(vars)) {
      if (Array.isArray(val)) continue;
      const raw = isCellEntry(val) ? val.value : val;
      this.resolvedScalars.set(name, raw == null ? "" : String(raw));
    }

    const affected = new Set<string>();
    for (const name of Object.keys(vars)) {
      const refs = this.variableIndex.get(name);
      if (refs) for (const k of refs) affected.add(k);
    }

    for (const key of affected) {
      const tpl = this.templateCells.get(key);
      if (!tpl) continue;
      const ws = this.workbook.getWorksheet(tpl.sheet);
      if (!ws) continue;
      const cell = ws.getCell(tpl.cell);

      if (tpl.soleToken && tpl.soleToken in vars) {
        const val = vars[tpl.soleToken];
        if (Array.isArray(val)) {
          this.setRange(`${tpl.sheet}!${tpl.cell}`, val, {
            direction: "column",
          });
          continue;
        }
        this.writeCell(cell, val);
        continue;
      }

      const substituted = tpl.template.replace(TOKEN_RE, (full, name) => {
        const resolved = this.resolvedScalars.get(name);
        return resolved !== undefined ? resolved : full;
      });
      cell.value = substituted;
      if (this.config.defaultColor) applyFill(cell, this.config.defaultColor);
    }
  }

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
      this.writeCell(ws.getCell(cellRef), entry);
    }
  }

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
  setRange(
    anchor: string,
    values: CellValueOrEntry[] | CellValueOrEntry[][],
    options: SetRangeOptions = {}
  ): void {
    if (!this.workbook) {
      throw new Error("No template loaded. Call loadTemplate() first.");
    }

    const { sheet, cell: cellRef } = parseCellRef(anchor);
    const ws = this.workbook.getWorksheet(sheet);
    if (!ws) {
      throw new Error(`Worksheet "${sheet}" not found in the template.`);
    }

    const { row: startRow, col: startCol } = a1ToRowCol(cellRef);
    const is2D = values.length > 0 && Array.isArray(values[0]);

    if (is2D) {
      const rows = values as CellValueOrEntry[][];
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        for (let c = 0; c < row.length; c++) {
          this.writeCell(ws.getCell(startRow + r, startCol + c), row[c]);
        }
      }
    } else {
      const arr = values as CellValueOrEntry[];
      const direction = options.direction ?? "column";
      for (let i = 0; i < arr.length; i++) {
        const target =
          direction === "column"
            ? ws.getCell(startRow + i, startCol)
            : ws.getCell(startRow, startCol + i);
        this.writeCell(target, arr[i]);
      }
    }
  }

  private writeCell(cell: ExcelJS.Cell, entry: CellValueOrEntry): void {
    if (isCellEntry(entry)) {
      cell.value = entry.value;

      let color: string | undefined;
      if (typeof entry.color === "function") {
        color = entry.color(entry.value);
      } else if (typeof entry.color === "string") {
        color = entry.color;
      }

      if (color) applyFill(cell, color);
      else if (this.config.defaultColor) applyFill(cell, this.config.defaultColor);
    } else {
      cell.value = entry as ExcelJS.CellValue;
      if (this.config.defaultColor) applyFill(cell, this.config.defaultColor);
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
