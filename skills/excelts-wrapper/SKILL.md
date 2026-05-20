# ExcelTS Wrapper Pattern

Build a typed, type-safe wrapper around [excelts](https://github.com/cj-tech-master/excelts) (`@cj-tech-master/excelts`). This skill documents the patterns used by `@quadro/core` and can be applied to any project wrapping excelts.

## Core Principles

1. **Never leak excelts types to consumers** — wrap `Worksheet`, `Row`, `Cell`, `Workbook` behind your own typed API
2. **Zero `as any` in source** — every type crossing must use discriminated union narrowing, branded types, or mapped transformation
3. **Fluent builders** — method chaining with `return this` for ergonomic construction
4. **Separate concerns** — `WorkbookBuilder` for I/O + metadata; `SheetBuilder` for cell/styling operations
5. **Dual CJS/ESM** — export both formats via Rolldown bundler
6. **Integration tests, not mocks** — test against real excelts objects

## Package Structure

```
packages/core/
├── src/
│   ├── index.ts             # Barrel — re-export all public API
│   ├── workbook-builder.ts  # WorkbookBuilder class (top-level)
│   ├── sheet-builder.ts     # SheetBuilder class (wraps Worksheet)
│   ├── formulas.ts          # F namespace — typed formula helpers
│   ├── types.ts             # Public type exports
│   └── utils.ts             # Style mappers, presets, helpers
├── rolldown.config.mjs      # Dual ESM/CJS build
└── package.json             # exports map: import → esm, require → cjs
```

## Patterns

### Pattern 1: Type-Safe Builder Wrapping

Wrap excelts objects behind a fluent builder. The builder holds a private reference to the underlying excelts object (`_ws`, `_wb`). All type boundaries use discriminated narrowing.

```typescript
import type { Worksheet as ExcelWorksheet } from "@cj-tech-master/excelts";

export class SheetBuilder {
  private readonly _ws: ExcelWorksheet;

  constructor(ws: ExcelWorksheet, opts: SheetOptions) {
    this._ws = ws;
  }

  setCell(address: string, value?: CellValue, style?: CellStyle): this {
    const cell = this._ws.getCell(address);
    this._writeValue(cell, value);
    if (style) applyStyle(cell, style);
    return this;
  }
}
```

**Passing excelts objects** — use static factories that assign private fields directly from same-class scope:

```typescript
export class WorkbookBuilder {
  private _wb: ExcelWorkbook;
  private _sheets: SheetBuilder[] = [];

  constructor(opts: WorkbookOptions = {}) {
    this._wb = new ExcelWorkbook();
  }

  static async load(filePath: string): Promise<WorkbookBuilder> {
    const wb = new ExcelWorkbook();
    await wb.xlsx.readFile(path.resolve(filePath));
    const builder = new WorkbookBuilder();
    builder._wb = wb; // same class scope — no cast needed
    for (let i = 0; i < wb.worksheets.length; i++) {
      const ws = wb.getWorksheet(i + 1);
      if (ws) builder._sheets.push(new SheetBuilder(ws, { name: wb.worksheets[i].name }));
    }
    return builder;
  }
}
```

### Pattern 2: Discriminated Union Narrowing

excelts uses discriminated unions (`WorksheetView`) and overloaded method signatures. Narrow via the discriminant, not `as` casts.

```typescript
freeze(row: number, col = 0): this {
  const prev = this._ws.views?.[0];
  const view: Partial<WorksheetView> = {
    ...(prev ? { showGridLines: prev.showGridLines, zoomScale: prev.zoomScale } : {}),
    state: "frozen",
    xSplit: col,
    ySplit: row,
    topLeftCell: `${colLetter(col + 1)}${row + 1}`,
    activeCell: `${colLetter(col + 1)}${row + 1}`,
  };
  this._ws.views = [view];
  return this;
}
```

### Pattern 3: Overloaded Methods with Fluent API

Use TypeScript overloads so the builder returns `this` (for chaining) when a callback is provided, or the inner builder (for imperative use) when omitted.

```typescript
addSheet(opts: SheetOptions, configure: (sheet: SheetBuilder) => void): this;
addSheet(opts: SheetOptions): SheetBuilder;
addSheet(opts: SheetOptions, configure?: (sheet: SheetBuilder) => void): this | SheetBuilder {
  const ws = this._wb.addWorksheet(opts.name);
  const builder = new SheetBuilder(ws, opts);
  this._sheets.push(builder);
  if (configure) { configure(builder); return this; }
  return builder;
}
```

### Pattern 4: Style Composition

Build cell styles from composable partials. The `style()` function uses typed generic rest parameters.

```typescript
export type StylePart = Partial<CellStyle>;

export function style(...parts: StylePart[]): CellStyle {
  const result: CellStyle = {};
  for (const part of parts) {
    if (part.font) result.font = { ...result.font, ...part.font };
    if (part.fill) result.fill = part.fill;
    if (part.border) result.border = { ...result.border, ...part.border };
    if (part.alignment) result.alignment = { ...result.alignment, ...part.alignment };
    if (part.numberFormat) result.numberFormat = part.numberFormat;
    if (part.protection) result.protection = { ...result.protection, ...part.protection };
  }
  return result;
}
```

### Pattern 5: Typed Bridge Functions

When the library's types are a superset of yours, write bridge functions that explicitly map each field:

```typescript
import type { CellFormulaValue } from "@cj-tech-master/excelts";

function isFormula(val: CellValue): val is FormulaValue {
  return typeof val === "object" && val !== null && !(val instanceof Date) && "formula" in val;
}

function toFormulaValue(v: FormulaValue): CellFormulaValue {
  const fv: CellFormulaValue = { formula: normalizeFormula(v.formula) };
  if (v.result !== undefined) fv.result = v.result;
  return fv;
}

function toExcelValue(val: CellValue): CellPrimitive | CellFormulaValue {
  if (isFormula(val)) return toFormulaValue(val);
  return val;
}
```

### Pattern 6: Integration Tests Without `as any`

Construct real excelts objects and pass them through the wrapper:

```typescript
import { Workbook as ExcelWorkbook } from "@cj-tech-master/excelts";

it("sets column definitions", () => {
  const ws = new ExcelWorkbook().addWorksheet("Test");
  const sheet = new SheetBuilder(ws, { name: "Test" });
  sheet.columns([{ key: "a", header: "A" }]);
  sheet.writeHeaders();
  expect(ws.columns).toBeDefined();
  // Subtype narrowing — compatible with Column[] shape
  const cols = ws.columns as ReadonlyArray<{ key: string }>;
  expect(cols[0].key).toBe("a");
});
```

## Common Mistakes

1. **`as any` on external type boundaries** — type alias instead: `import type { Worksheet as ExcelWorksheet }` and let TypeScript validate.
2. **`as any[]` for array access** — use `as ReadonlyArray<{ key: string }>` — narrows to a compatible subtype.
3. **Non-null assertions (`!`) instead of optional chaining** — `border.thinBlack.border?.top?.color` not `border!`.
4. **Leaking excels types in public API** — never export `ExcelWorkbook`, return your own `SheetBuilder[]`.
5. **Caret range on excels dependency** — pin exact: `"9.5.4"` not `"^9.5.4"`.

## Export Map Configuration

```json
{
  "type": "module",
  "main": "dist/cjs/index.js",
  "types": "dist/cjs/index.d.ts",
  "exports": {
    ".": {
      "import": "./dist/esm/index.js",
      "require": "./dist/cjs/index.js"
    }
  },
  "files": ["dist"]
}
```
