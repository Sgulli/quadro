# @qquadro/core

Fluent, fully-typed Excel workbook builder for Node.js. Built on [excelts](https://github.com/cjnoname/excelts) by [@cjnoname](https://github.com/cjnoname). Created by [@sgulli](https://github.com/Sgulli).

```ts
import { WorkbookBuilder, Styles, F } from "@qquadro/core";

await new WorkbookBuilder({ author: "Acme Corp" })
  .addSheet({ name: "Sales" }, (sheet) => {
    sheet
      .headers([
        { key: "product", header: "Product", width: 24, headerStyle: Styles.header },
        { key: "revenue", header: "Revenue ($)", width: 18, style: Styles.currency, headerStyle: Styles.header },
      ])
      .addRows([
        { product: "Widget Pro", revenue: 120_000 },
        { product: "Gadget X",   revenue: 98_500 },
      ])
      .addRow(
        { product: "Total", revenue: F.sum("B2:B3") },
        { style: Styles.totalRow },
      )
      .autoFitColumns()
      .freeze(1);
  })
  .write("./report.xlsx");
```

## Features

- **Fluent API** — chainable `addSheet` → `headers` → `addRows` → `write`
- **Full styling** — fonts, fills, borders, alignment, number formats with preset styles
- **Formulas** — typed helpers for SUM, AVERAGE, IF, ADD, PCT, and more
- **Data validation** — dropdown lists, number/date ranges, custom rules
- **Conditional formatting** — cell rules, data bars, color scales, icon sets, top N
- **Numeric tuple API** — reference cells by `[col, row]` tuples, not A1 strings
- **Merged cells** — merge ranges with value and style
- **Freeze panes & auto-filter** — header rows stay visible, one-liner filter
- **Headers & footers** — per-sheet odd/even header/footer sections
- **Page setup** — orientation, paper size, fit-to-page, margins
- **Streaming mode** — constant memory for large datasets
- **Dual ESM/CJS** — works with `import` and `require`

## Data Validation

Restrict cell input with dropdown lists, number ranges, dates, or custom formulas. Remove with `sheet.removeDataValidation(address)`.

Columns are inferred from your data — no need to define them explicitly:

```ts
import { WorkbookBuilder } from "@qquadro/core";

await new WorkbookBuilder()
  .addSheet({ name: "Validation" }, (sheet) => {
    sheet
      .addRows([
        { product: "Widget A", quantity: 100, date: new Date("2026-01-15") },
        { product: "Widget B", quantity: 50, date: new Date("2026-02-01") },
      ])
      // Columns "product", "quantity", "date" are auto-inferred from the first row.
      // sheet.columnRange("quantity") → "B2:B3"
      .addListValidation(sheet.columnRange("product"), ["Widget A", "Widget B", "Gadget X"])
      .addRangeValidation(sheet.columnRange("quantity"), {
        type: "whole", operator: "between", formulae: [1, 1000],
        error: "Must be between 1 and 1000",
        errorTitle: "Invalid quantity",
      })
      .addRangeValidation(sheet.columnRange("date"), {
        type: "date", operator: "greaterThanOrEqual", formulae: [new Date("2026-01-01")],
      });
  })
  .write("./validation.xlsx");
```

## Conditional Formatting

Highlight cells, add data bars, color scales, icon sets, and top N rules. Columns are inferred from your data:

```ts
import { WorkbookBuilder } from "@qquadro/core";

await new WorkbookBuilder()
  .addSheet({ name: "Formatting" }, (sheet) => {
    sheet
      .addRows([
        { name: "Alice", score: 85, status: "Pass" },
        { name: "Bob", score: 42, status: "Fail" },
        { name: "Charlie", score: 73, status: "Pass" },
      ])
      .addCellIsRule(sheet.columnRange("score"), "greaterThanOrEqual", [70], {
        font: { bold: true, color: { argb: "FF006100" } },
        fill: { type: "solid", color: { argb: "FFC6EFCE" } },
      })
      .addCellIsRule(sheet.columnRange("score"), "lessThan", [50], {
        font: { color: { argb: "FF9C0006" } },
        fill: { type: "solid", color: { argb: "FFFFC7CE" } },
      })
      .addDataBar(sheet.columnRange("score"), { argb: "FF5B9BD5" })
      .addIconSet(sheet.columnRange("score"), "3TrafficLights1")
      .addContainsTextRule(sheet.columnRange("status"), "Pass", "containsText", {
        font: { color: { argb: "FF006100" } },
      })
      .addContainsTextRule(sheet.columnRange("status"), "Fail", "containsText", {
        font: { color: { argb: "FF9C0006" } },
      });
  })
  .write("./formatting.xlsx");
```

## Numeric Tuple API

Every method that accepts an A1 string also accepts a `[col, row]` or `[c1, r1, c2, r2]` tuple — build entire workbooks with `for` loops, no string concatenation.

```ts
import { WorkbookBuilder, F } from "@qquadro/core";

const data = [
  { name: "Alice", score: 85, status: "Pass" },
  { name: "Bob", score: 42, status: "Fail" },
];
const dataStart = 2;

await new WorkbookBuilder()
  .addSheet({ name: "Grades" }, (sheet) => {
    // Set up headers — columnIndex resolves "score" → 2
    sheet.headers([
      { key: "name", header: "Name" },
      { key: "score", header: "Score" },
      { key: "status", header: "Status" },
    ]);

    const c = { name: sheet.columnIndex("name"), score: sheet.columnIndex("score"), status: sheet.columnIndex("status") };

    // Write data with for loop using tuple addresses
    data.forEach((row, i) => {
      const r = dataStart + i;
      sheet
        .setCell([c.name, r], row.name)
        .setCell([c.score, r], row.score)
        .setCell([c.status, r], row.status);
    });

    // Apply validation and formatting by column index
    const dr = { start: dataStart, end: dataStart + data.length - 1 };
    sheet
      .addDataBar([c.score, dr.start, c.score, dr.end], { argb: "FF5B9BD5" })
      .addCellIsRule([c.score, dr.start, c.score, dr.end], "greaterThanOrEqual", [70], {
        font: { bold: true, color: { argb: "FF006100" } },
      })
      .addIconSet([c.score, dr.start, c.score, dr.end], "3TrafficLights1");

    // Formula: sum in footer row
    sheet.setCell([c.score, dr.end + 1], F.sum(`${c.score}${dr.start}:${c.score}${dr.end}`));
  })
  .write("./grades.xlsx");
```

### Tuple overload reference

All A1 methods also accept numeric tuples. Pass them anywhere you'd pass an A1 string:

| Category | A1 string | Numeric tuple |
|---|---|---|
| Cell | `setCell("D5", val)` | `setCell([4, 5], val)` |
| Merge | `merge("D1:E1")` | `merge([4, 1, 5, 1])` |
| Style | `styleRange("D2:D10", s)` | `styleRange([4, 2, 4, 10], s)` |
| Validation | `addDataValidation("A1", v)` | `addDataValidation([1, 1], v)` |
| List | `addListValidation("B3:B10", [...])` | `addListValidation([2, 3, 2, 10], [...])` |
| Range val. | `addRangeValidation("D2:D10", ...)` | `addRangeValidation([4, 2, 4, 10], ...)` |
| Cell is | `addCellIsRule("D2:D10", ...)` | `addCellIsRule([4, 2, 4, 10], ...)` |
| Expression | `addExpressionRule("D2:D10", ...)` | `addExpressionRule([4, 2, 4, 10], ...)` |
| Data bar | `addDataBar("D2:D10")` | `addDataBar([4, 2, 4, 10])` |
| Color scale | `addColorScale("C1:C10", ...)` | `addColorScale([3, 1, 3, 10], ...)` |
| Icon set | `addIconSet("D1:D10", ...)` | `addIconSet([4, 1, 4, 10], ...)` |
| Top 10 | `addTop10Rule("E1:E10", ...)` | `addTop10Rule([5, 1, 5, 10], ...)` |
| Above avg | `addAboveAverageRule("F1:F10")` | `addAboveAverageRule([6, 1, 6, 10])` |
| Text | `addContainsTextRule("G1:G10", ...)` | `addContainsTextRule([7, 1, 7, 10], ...)` |
| Time | `addTimePeriodRule("H1:H10", ...)` | `addTimePeriodRule([8, 1, 8, 10], ...)` |
| Table | `addTable("T", "A1:C10", ...)` | `addTable("T", [1, 1, 3, 10], ...)` |
| Note | `addNote("B3", text)` | `addNote([2, 3], text)` |
| Hyperlink | `setCellHyperlink("C5", url)` | `setCellHyperlink([3, 5], url)` |
| Rich text | `setCellRichText("A5", rt)` | `setCellRichText([1, 5], rt)` |

### Coordinate helpers

```ts
cellRef(4, 5)              // → "D5"
colRange(4, 2, 10)         // → "D2:D10"
rangeRef(4, 2, 7, 10)      // → "D2:G10"
sheet.columnIndex("score")  // → 2  (1‑based index from headers)
```

## Installation

```bash
npm install @qquadro/core
```

```bash
pnpm add @qquadro/core
```

```bash
yarn add @qquadro/core
```

```bash
bun add @qquadro/core
```

## Documentation

See the [monorepo README](https://github.com/Sgulli/quadro) for full API reference, examples, and style system documentation.

## License

MIT
