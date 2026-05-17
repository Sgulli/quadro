import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import { SheetBuilder } from "./sheet-builder.js";
import type { SheetOptions, WorkbookOptions, WriteResult } from "./types.js";

/**
 * Top-level fluent builder for an Excel workbook.
 *
 * Supports two write modes:
 * - **Buffer mode** (default): builds in memory, then writes to disk.
 * - **Stream mode**: writes directly to a file via a Node.js write stream —
 *   constant memory regardless of file size, ideal for large datasets.
 *
 * @example
 * ```ts
 * import { WorkbookBuilder, Styles } from "./index";
 *
 * const result = await new WorkbookBuilder({ author: "Acme Corp" })
 *   .addSheet({ name: "Sales" }, (sheet) => {
 *     sheet
 *       .columns([
 *         { key: "region", header: "Region", width: 20, headerStyle: Styles.header },
 *         { key: "revenue", header: "Revenue ($)", width: 18, style: Styles.currency, headerStyle: Styles.header },
 *       ])
 *       .writeHeaders()
 *       .addRows([
 *         { region: "EMEA",  revenue: 120_000 },
 *         { region: "APAC",  revenue: 98_500  },
 *         { region: "AMER",  revenue: 210_000 },
 *       ])
 *       .addRow({ region: "Total", revenue: { formula: "=SUM(B2:B4)" } }, { style: Styles.totalRow })
 *       .autoFitColumns()
 *       .freeze(1);
 *   })
 *   .write("./output/report.xlsx");
 *
 * console.log("Written to", result.filePath, `(${result.sizeBytes} bytes)`);
 * ```
 */
export class WorkbookBuilder {
  private readonly _wb: ExcelJS.Workbook;
  private readonly _opts: WorkbookOptions;
  private _sheetBuilders: SheetBuilder[] = [];

  constructor(opts: WorkbookOptions = {}) {
    this._opts = opts;
    this._wb = new ExcelJS.Workbook();
    this._applyWorkbookMeta();
  }

  // ── Sheet API ──────────────────────────────────────────────────────────────

  /**
   * Add a sheet using a callback-based builder pattern.
   *
   * The callback receives a fully-typed `SheetBuilder`. No need to call
   * `.done()` — the parent regains control automatically after the callback.
   *
   * @example
   * builder.addSheet({ name: "Summary" }, (sheet) => {
   *   sheet.columns([...]).writeHeaders().addRows([...]);
   * })
   */
  addSheet(opts: SheetOptions, configure: (sheet: SheetBuilder) => void): this;

  /**
   * Add a sheet and get back the `SheetBuilder` for imperative use.
   * Call `.done()` on the builder when finished.
   */
  addSheet(opts: SheetOptions): SheetBuilder;

  addSheet(
    opts: SheetOptions,
    configure?: (sheet: SheetBuilder) => void,
  ): this | SheetBuilder {
    const ws = this._wb.addWorksheet(opts.name);
    let resolveBuilder!: () => void;
    const builder = new SheetBuilder(ws, opts, () => resolveBuilder?.());
    this._sheetBuilders.push(builder);

    if (configure) {
      configure(builder);
      return this;
    }

    // Imperative path — return the builder
    resolveBuilder = () => {
      /* no-op; commit happens on write */
    };
    return builder;
  }

  // ── Write API ──────────────────────────────────────────────────────────────

  /**
   * Write the workbook to disk.
   *
   * - Uses **streaming mode** when `opts.streamToFile` was set in the
   *   constructor — O(1) memory for arbitrarily large files.
   * - Otherwise writes via an in-memory buffer.
   *
   * @param outputPath  Destination file path (.xlsx).
   */
  async write(outputPath: string): Promise<WriteResult> {
    const resolved = path.resolve(outputPath);
    const dir = path.dirname(resolved);

    // Ensure the target directory exists
    await fs.promises.mkdir(dir, { recursive: true });

    if (this._opts.streamToFile) {
      // ── Streaming path ─────────────────────────────────────────────────────
      const stream = fs.createWriteStream(resolved);
      await this._wb.xlsx.write(stream);

      await new Promise<void>((ok, err) => {
        stream.on("finish", ok);
        stream.on("error", err);
      });

      const { size } = await fs.promises.stat(resolved);
      return { filePath: resolved, sizeBytes: size };
    }

    // ── Buffer path ────────────────────────────────────────────────────────
    await this._wb.xlsx.writeFile(resolved);
    const { size } = await fs.promises.stat(resolved);
    return { filePath: resolved, sizeBytes: size };
  }

  /**
   * Return the workbook as a `Buffer` (e.g. for HTTP responses).
   */
  async toBuffer(): Promise<Buffer> {
    return this._wb.xlsx.writeBuffer() as unknown as Promise<Buffer>;
  }

  /**
   * Return the underlying ExcelJS `Workbook` for advanced operations not
   * covered by this wrapper.
   */
  get workbook(): ExcelJS.Workbook {
    return this._wb;
  }

  // ── Private ────────────────────────────────────────────────────────────────

  private _applyWorkbookMeta(): void {
    const { author, company, created, useSharedStrings } = this._opts;
    if (author) this._wb.creator = author;
    if (company) this._wb.company = company;
    this._wb.created = created ?? new Date();
    if (useSharedStrings !== undefined)
      (this._wb as unknown as { useSharedStrings: boolean }).useSharedStrings =
        useSharedStrings;
    this._wb.calcProperties.fullCalcOnLoad = true;
  }
}
