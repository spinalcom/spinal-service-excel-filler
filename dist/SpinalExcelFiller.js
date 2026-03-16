import ExcelJS from "exceljs";
function isCellEntry(v) {
    return typeof v === "object" && v !== null && "value" in v;
}
function parseCellRef(ref) {
    const idx = ref.indexOf("!");
    if (idx === -1) {
        throw new Error(`Invalid cell reference "${ref}". Expected format: "SheetName!CellRef" (e.g. "Sheet1!B3").`);
    }
    return { sheet: ref.slice(0, idx), cell: ref.slice(idx + 1) };
}
function applyFill(cell, hex) {
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
export class SpinalExcelFiller {
    constructor(config = {}) {
        this.workbook = null;
        this.config = config;
    }
    /**
     * Load an Excel template from a file path.
     */
    async loadTemplate(filePath) {
        this.workbook = new ExcelJS.Workbook();
        await this.workbook.xlsx.readFile(filePath);
    }
    /**
     * Load an Excel template from a Buffer.
     */
    async loadTemplateFromBuffer(buffer) {
        this.workbook = new ExcelJS.Workbook();
        await this.workbook.xlsx.load(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength));
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
    setCells(cells) {
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
                let color;
                if (typeof entry.color === "function") {
                    color = entry.color(entry.value);
                }
                else if (typeof entry.color === "string") {
                    color = entry.color;
                }
                if (color) {
                    applyFill(cell, color);
                }
                else if (this.config.defaultColor) {
                    applyFill(cell, this.config.defaultColor);
                }
            }
            else {
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
    async save(outputPath) {
        if (!this.workbook) {
            throw new Error("No template loaded. Call loadTemplate() first.");
        }
        await this.workbook.xlsx.writeFile(outputPath);
    }
    /**
     * Return the filled workbook as a Buffer.
     */
    async toBuffer() {
        if (!this.workbook) {
            throw new Error("No template loaded. Call loadTemplate() first.");
        }
        const arrayBuffer = await this.workbook.xlsx.writeBuffer();
        return Buffer.from(arrayBuffer);
    }
}
//# sourceMappingURL=SpinalExcelFiller.js.map