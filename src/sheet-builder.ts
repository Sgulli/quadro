import ExcelJS from "exceljs";
import { applyStyle } from "./utils.js";
import type {
  CellPrimitive,
  CellStyle,
  CellValue,
  ColumnDef,
  MergeRange,
  RowData,
  RowOptions,
  SheetOptions,
} from "./types.js";

// ─── Internal row buffer (streaming-friendly) ────────────────────────────────

interface BufferedRow {
  data: RowData;
  options?: RowOptions;
}

// ─── SheetBuilder ────────────────────────────────────────────────────────────

/**
 * Fluent builder for a single worksheet.
 *
 * Obtain an instance via `WorkbookBuilder.addSheet()`.
 * Call `.done()` to return control to the parent builder.
 */
export class SheetBuilder {
  private readonly _ws: ExcelJS.Worksheet;
  private _columns: ColumnDef[] = [];
  private _merges: MergeRange[] = [];
  private _rows: BufferedRow[] = [];
  private _headerWritten = false;

  /** @internal */
  constructor(
    ws: ExcelJS.Worksheet,
    private readonly _opts: SheetOptions,
    private readonly _done: () => void,
  ) {
    this._ws = ws;
    this._applySheetOptions();
  }

  // ── Column API ─────────────────────────────────────────────────────────────

  /**
   * Define all columns at once. Replaces any previously set columns.
   *
   * @example
   * sheet.columns([
   *   { key: "name",  header: "Name",   width: 25, headerStyle: Styles.header },
   *   { key: "sales", header: "Sales",  width: 15, style: Styles.currency },
   * ])
   */
  columns(defs: ColumnDef[]): this {
    this._columns = defs;
    return this;
  }

  /**
   * Append a single column definition.
   */
  addColumn(def: ColumnDef): this {
    this._columns.push(def);
    return this;
  }

  // ── Header API ─────────────────────────────────────────────────────────────

  /**
   * Write column headers to the current row and advance the row cursor.
   * Optionally provide a global header style that is merged with per-column
   * `headerStyle`.
   *
   * Call this *after* `.columns()`.
   */
  writeHeaders(globalStyle?: CellStyle): this {
    if (this._headerWritten) {
      throw new Error(
        `[SheetBuilder] writeHeaders() already called on sheet "${this._opts.name}".`,
      );
    }
    if (this._columns.length === 0) {
      throw new Error(
        `[SheetBuilder] Call columns() before writeHeaders() on sheet "${this._opts.name}".`,
      );
    }

    const row = this._ws.addRow(this._columns.map((c) => c.header));
    row.eachCell((cell, colNumber) => {
      const col = this._columns[colNumber - 1];
      if (globalStyle) applyStyle(cell, globalStyle);
      if (col?.headerStyle) applyStyle(cell, col.headerStyle);
    });
    row.commit();

    this._headerWritten = true;

    // Register column definitions (width, key, hidden, default style)
    this._ws.columns = this._columns.map((c) => ({
      key: c.key,
      width: c.width ?? 15,
      hidden: c.hidden ?? false,
    }));

    return this;
  }

  // ── Row API ─────────────────────────────────────────────────────────────────

  /**
   * Add a single row. Data can be:
   * - An array of values aligned to column order
   * - A key→value object aligned to `ColumnDef.key`
   *
   * @example
   * sheet.addRow({ name: "Alice", sales: { formula: "=B2*1.1" } })
   * sheet.addRow(["Alice", 42000], { height: 20 })
   */
  addRow(data: RowData, options?: RowOptions): this {
    this._rows.push({ data, options });
    this._flushRow({ data, options });
    return this;
  }

  /**
   * Add multiple rows efficiently.
   */
  addRows(rows: RowData[], options?: RowOptions): this {
    for (const data of rows) this.addRow(data, options);
    return this;
  }

  // ── Cell API ───────────────────────────────────────────────────────────────

  /**
   * Write a value and/or style to an individual cell by address.
   *
   * @param address  A1-style address, e.g. "B3"
   */
  setCell(address: string, value?: CellValue, style?: CellStyle): this {
    const cell = this._ws.getCell(address);
    this._writeValue(cell, value);
    if (style) applyStyle(cell, style);
    return this;
  }

  /**
   * Apply a style to a range of cells.
   *
   * @param range  e.g. "A1:D4"
   */
  styleRange(range: string, style: CellStyle): this {
    const [tl, br] = range.split(":");
    const tlCell = this._ws.getCell(tl!);
    const brCell = this._ws.getCell(br ?? tl!);

    for (let r = tlCell.fullAddress.row; r <= brCell.fullAddress.row; r++) {
      for (let c = tlCell.fullAddress.col; c <= brCell.fullAddress.col; c++) {
        applyStyle(this._ws.getCell(r, c), style);
      }
    }
    return this;
  }

  // ── Merge API ──────────────────────────────────────────────────────────────

  /**
   * Merge a range of cells, optionally writing a value and style to the top-left cell.
   *
   * @example
   * sheet.merge({ range: "A1:E1", value: "Q1 Report", style: Styles.header })
   */
  merge(region: MergeRange): this {
    this._merges.push(region);
    this._ws.mergeCells(region.range);

    const [tl] = region.range.split(":");
    if (!tl) return this;

    const cell = this._ws.getCell(tl);
    if (region.value !== undefined) this._writeValue(cell, region.value);
    if (region.style) applyStyle(cell, region.style);

    return this;
  }

  /**
   * Merge multiple regions at once.
   */
  mergeAll(regions: MergeRange[]): this {
    for (const r of regions) this.merge(r);
    return this;
  }

  // ── Row height / column width shortcuts ────────────────────────────────────

  /**
   * Set the height of a specific row (1-indexed).
   */
  rowHeight(rowNumber: number, height: number): this {
    this._ws.getRow(rowNumber).height = height;
    return this;
  }

  /**
   * Set the width of a specific column (letter or 1-indexed number).
   */
  colWidth(col: string | number, width: number): this {
    const c =
      typeof col === "string"
        ? this._ws.getColumn(col)
        : this._ws.getColumn(col);
    c.width = width;
    return this;
  }

  /**
   * Auto-fit column width based on header length + padding.
   * Call after writing headers.
   */
  autoFitColumns(paddingChars = 4): this {
    this._columns.forEach((def, i) => {
      const col = this._ws.getColumn(i + 1);
      const headerLen = def.header.length;
      if (!def.width) col.width = Math.max(headerLen + paddingChars, 10);
    });
    return this;
  }

  // ── Freeze / auto-filter ───────────────────────────────────────────────────

  /**
   * Freeze panes at the given row and column.
   * E.g. `freeze(1, 0)` freezes the header row.
   */
  freeze(row: number, col = 0): this {
    this._ws.views = [
      {
        state: "frozen",
        xSplit: col,
        ySplit: row,
        topLeftCell: `${colLetter(col + 1)}${row + 1}`,
        activeCell: `${colLetter(col + 1)}${row + 1}`,
      },
    ];
    return this;
  }

  /**
   * Enable Excel's auto-filter on the header row range.
   *
   * @param range  e.g. "A1:F1" — defaults to the full header row.
   */
  autoFilter(range?: string): this {
    const r =
      range ??
      (this._columns.length ? `A1:${colLetter(this._columns.length)}1` : "A1");
    this._ws.autoFilter = r;
    return this;
  }

  // ── Done ───────────────────────────────────────────────────────────────────

  /**
   * Finalise this sheet and return to the parent WorkbookBuilder.
   */
  done(): void {
    this._done();
  }

  // ── Private helpers ────────────────────────────────────────────────────────

  private _applySheetOptions(): void {
    const opts = this._opts;

    if (opts.tabColor) {
      this._ws.properties.tabColor = { argb: opts.tabColor };
    }
    if (opts.defaultRowHeight !== undefined) {
      this._ws.properties.defaultRowHeight = opts.defaultRowHeight;
    }
    if (opts.showGridLines === false) {
      this._ws.views = [{ showGridLines: false }];
    }
    if (opts.zoom !== undefined) {
      this._ws.views = [
        ...(this._ws.views ?? []),
        { zoomScale: opts.zoom } as ExcelJS.WorksheetView,
      ];
    }
    if (opts.freeze) {
      const { row = 0, col = 0 } = opts.freeze;
      this.freeze(row, col);
    }
    if (opts.pageSetup) {
      Object.assign(this._ws.pageSetup, {
        paperSize: opts.pageSetup.paperSize ?? 9,
        orientation: opts.pageSetup.orientation ?? "portrait",
        fitToPage: opts.pageSetup.fitToPage ?? false,
        fitToWidth: opts.pageSetup.fitToWidth ?? 1,
        fitToHeight: opts.pageSetup.fitToHeight ?? 0,
      });
      if (opts.pageSetup.margins) {
        this._ws.pageSetup.margins = opts.pageSetup.margins as ExcelJS.Margins;
      }
    }
    if (opts.headerFooter) {
      const hf = opts.headerFooter;
      const flatSection = (s: NonNullable<typeof hf.oddHeader>): string =>
        [
          s.left && `&L${s.left}`,
          s.center && `&C${s.center}`,
          s.right && `&R${s.right}`,
        ]
          .filter(Boolean)
          .join("");
      if (hf.oddHeader)
        this._ws.headerFooter.oddHeader = flatSection(hf.oddHeader);
      if (hf.oddFooter)
        this._ws.headerFooter.oddFooter = flatSection(hf.oddFooter);
      if (hf.evenHeader)
        this._ws.headerFooter.evenHeader = flatSection(hf.evenHeader);
      if (hf.evenFooter)
        this._ws.headerFooter.evenFooter = flatSection(hf.evenFooter);
    }
    if (opts.protection) {
      this._ws.protect(opts.protection.password ?? "", {});
    }
  }

  private _flushRow(buffered: BufferedRow): void {
    const { data, options } = buffered;

    let excelRow: ExcelJS.Row;

    if (Array.isArray(data)) {
      const rowValues = data.map(toExcelValue);
      excelRow = this._ws.addRow(rowValues);
    } else {
      const objValues: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(data)) {
        objValues[k] = toExcelValue(v);
      }
      excelRow = this._ws.addRow(objValues);
    }

    if (this._columns.length > 0) {
      excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const col = this._columns[colNumber - 1];
        if (col?.style) applyStyle(cell, col.style);
      });
    }

    if (options) {
      if (options.height !== undefined) excelRow.height = options.height;
      if (options.hidden) excelRow.hidden = true;
      if (options.outlineLevel !== undefined)
        excelRow.outlineLevel = options.outlineLevel;
      if (options.style) {
        excelRow.eachCell({ includeEmpty: true }, (cell) =>
          applyStyle(cell, options.style),
        );
      }
    }

    excelRow.commit();
  }

  private _writeValue(cell: ExcelJS.Cell, value: CellValue | undefined): void {
    if (value === undefined || value === null) return;
    if (isFormula(value)) {
      cell.value = toFormulaValue(value);
    } else {
      cell.value = value as ExcelJS.CellValue;
    }
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function isFormula(
  val: CellValue,
): val is { formula: string; result?: CellPrimitive } {
  return (
    typeof val === "object" &&
    val !== null &&
    !(val instanceof Date) &&
    "formula" in val
  );
}

function normalizeFormula(f: string): string {
  return f.replace(/^=+/g, "");
}

function toFormulaValue(v: {
  formula: string;
  result?: CellPrimitive;
}): ExcelJS.CellFormulaValue {
  const fv: ExcelJS.CellFormulaValue = { formula: normalizeFormula(v.formula) };
  if (v.result !== undefined && v.result !== null) fv.result = v.result;
  return fv;
}

function toExcelValue(val: CellValue): unknown {
  if (isFormula(val)) {
    const fv: Record<string, unknown> = { formula: normalizeFormula(val.formula) };
    if (val.result !== undefined && val.result !== null) fv.result = val.result;
    return fv;
  }
  return val;
}

function colLetter(col: number): string {
  let result = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    col = Math.floor((col - 1) / 26);
  }
  return result;
}
