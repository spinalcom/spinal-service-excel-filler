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

### Fill a range from an array

Use `setRange` to fill a column, a row, or a 2D block from an array, starting at an anchor cell.

```ts
// Column downward from B3 (default direction)
filler.setRange("Sheet1!B3", [10, 20, 30, 40]);

// Row rightward from B3
filler.setRange("Sheet1!B3", ["Jan", "Feb", "Mar"], { direction: "row" });

// 2D block (row-major) — great for filling tables
filler.setRange("Sheet1!B3", [
  ["Name",  "Score"],
  ["Alice", 92],
  ["Bob",   71],
  ["Carol", 88],
]);
```

Each value may also be a `CellEntry`, so per-cell coloring works in ranges too:

```ts
filler.setRange("Sheet1!B3", [
  { value: 92, color: "4CAF50" },
  { value: 71, color: (v) => ((v as number) >= 80 ? "4CAF50" : "F44336") },
  { value: 88, color: "4CAF50" },
]);
```

### Template variables

Instead of hard-coding cell references in your code, write `{{tokens}}` directly in the template cells. The consumer only needs to know the variable names — not where they live.

**In the template (`template.xlsx`):**

| A | B |
| --- | --- |
| Name: | `{{client.name}}` |
| Quote: | `Hello {{client.name}}, your total is {{amount}} €` |
| Items: | `{{items}}` |

**In code:**

```ts
const filler = new SpinalExcelFiller();
await filler.loadTemplate("./template.xlsx");

// Discover what the template needs
console.log(filler.getVariables());
// -> ["client.name", "amount", "items"]

console.log(filler.getVariableLocations());
// -> { "client.name": ["Sheet1!B1", "Sheet1!B2"], "amount": ["Sheet1!B2"], "items": ["Sheet1!B3"] }

filler.setVariables({
  "client.name": "ACME Corp",
  amount: 1250,
  items: ["Widget", "Gadget", "Sprocket"], // array fills downward from {{items}}
});

await filler.save("./output.xlsx");
```

Behavior:

- A cell containing **only** the token (e.g. `{{amount}}`) receives the raw value — types (number, Date, boolean, formula) and `CellEntry` coloring are preserved.
- A cell with an **embedded** token (e.g. `"Hello {{client.name}}"`) gets stringified substitution.
- An **array** value on a sole-token cell fills downward from that cell (delegates to `setRange`).
- `setVariables` can be called multiple times; previously set scalars are remembered for later substitutions.

You can still combine it with `setCells` / `setRange` for cells you don't want to tokenize.

### Quick end-to-end test

A minimal script you can run to sanity-check both features:

```ts
// test.ts — run with: npx tsx test.ts
import { SpinalExcelFiller } from "spinal-service-excel-filler";

const filler = new SpinalExcelFiller({ defaultColor: "E3F2FD" });
await filler.loadTemplate("./template.xlsx");

console.log("Template variables:", filler.getVariables());

filler.setVariables({
  "client.name": "ACME Corp",
  amount: 1250,
  items: ["Widget", "Gadget", "Sprocket"],
});

filler.setRange("Sheet1!D3", [
  ["Q1", 100],
  ["Q2", 140],
  ["Q3", 180],
]);

await filler.save("./output.xlsx");
console.log("Wrote ./output.xlsx");
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
| `setRange(anchor, values, options?)` | Fill a column (default), a row, or a 2D block from an array starting at `anchor`. `options.direction` is `"column"` (default) or `"row"` for 1D arrays. |
| `getVariables()` | Return the list of `{{name}}` tokens found in the loaded template. |
| `getVariableLocations()` | Return `{ name: ["Sheet!Cell", ...] }` — where each variable appears. |
| `setVariables(vars)` | Assign values to variables by name. Arrays on sole-token cells fill downward. Scalars accumulate across calls. |
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