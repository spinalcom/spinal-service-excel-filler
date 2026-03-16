# spinal-service-excel-filler

Lightweight module to fill an Excel template with values and optional conditional cell coloring. Built on top of [ExcelJS](https://github.com/exceljs/exceljs).

## Installation

```bash
npm install spinal-service-excel-filler
```

## Usage

### Basic — fill cells with plain values

```ts
import { SpinalExcelFiller } from "spinal-service-excel-filler";

const filler = new SpinalExcelFiller();
await filler.loadTemplate("./template.xlsx");

filler.setCells({
  "Sheet1!B3": "John",
  "Sheet1!C5": 42,
  "Sheet1!D2": new Date(),
});

await filler.save("./output.xlsx");
```

Cell references use the format `SheetName!CellRef` (e.g. `Sheet1!B3`).

### Conditional coloring

Pass a `CellEntry` object instead of a plain value to control the background color of the cell:

```ts
filler.setCells({
  // Static color (hex without #)
  "Sheet1!B3": {
    value: "Warning",
    color: "FFC107",
  },

  // Dynamic color based on the value
  "Sheet1!C5": {
    value: 85,
    color: (v) => ((v as number) >= 80 ? "4CAF50" : "F44336"),
  },
});
```

### Default color

You can set a default background color that applies to every filled cell (unless overridden):

```ts
const filler = new SpinalExcelFiller({ defaultColor: "E3F2FD" });
```

### Loading from a Buffer

If you already have the template in memory:

```ts
const templateBuffer: Buffer = /* ... */;
await filler.loadTemplateFromBuffer(templateBuffer);
```

### Getting the output as a Buffer

```ts
const buffer = await filler.toBuffer();
// e.g. send as an HTTP response
```

## API

| Method | Description |
| --- | --- |
| `new SpinalExcelFiller(config?)` | Create an instance. `config.defaultColor` sets a default cell background color. |
| `loadTemplate(filePath)` | Load an `.xlsx` template from disk. |
| `loadTemplateFromBuffer(buffer)` | Load an `.xlsx` template from a Buffer. |
| `setCells(cells)` | Fill cells. Keys are `"SheetName!CellRef"`, values are primitives or `CellEntry` objects. |
| `save(outputPath)` | Write the filled workbook to a file. |
| `toBuffer()` | Return the filled workbook as a Buffer. |

### `CellEntry`

```ts
interface CellEntry {
  value: ExcelJS.CellValue;
  color?: string | ((value: ExcelJS.CellValue) => string | undefined);
}
```

- **`value`** — The cell value (string, number, Date, boolean, etc.).
- **`color`** — A hex color string (e.g. `"4CAF50"`) or a function that receives the value and returns a hex color. If the function returns `undefined`, the default color (if any) is used.

## License

ISC