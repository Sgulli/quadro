export type * from "./types/options.js";
export type * from "./types/style.js";
export { BUILTIN_FORMATS } from "./types/style.js";
export type * from "./types/validation.js";

// ─── Re-exported excelts type aliases ─────────────────────────────────────────

export type CellHyperlinkValueInput = import("@cj-tech-master/excelts").CellHyperlinkValueInput;
export type CellRichTextValue = import("@cj-tech-master/excelts").CellRichTextValue;
export type ExcelFormCheckbox = import("@cj-tech-master/excelts").FormCheckbox;
export type RichTextRun = import("@cj-tech-master/excelts").RichText;

// ─── SheetBuilder Extension Interface ────────────────────────────────────────

import type {
  AddImageRange,
  CellIsOperators,
  ConditionalFormattingOptions,
  ContainsTextOperators,
  DataValidation,
  DataValidationWithFormulae,
  Style as ExcelStyle,
  FormCheckboxOptions,
  FormControlRange,
  IconSetTypes,
  ThreadedComment,
  TimePeriodTypes,
  WatermarkOptions,
} from "@cj-tech-master/excelts";
import type {
  AddBarChartOptions,
  AddChartExOptions,
  AddChartOptions,
  AddChartRange,
  AddComboChartOptions,
  AddPieChartOptions,
  AddScatterChartOptions,
  AddSurfaceChartOptions,
} from "@cj-tech-master/excelts/chart";
import type {
  AddSparklineGroupOptions,
  CellValue,
  MergeRange,
  NoteConfig,
  RowOptions,
} from "./types/options.js";
import type { CellStyle } from "./types/style.js";
import type { Addr, CellRange, RangeValidationDef } from "./types/validation.js";

export interface SheetBuilderExtension {
  addChart(options: AddChartOptions, range: AddChartRange): unknown;
  addColumnChart(
    options: Omit<AddBarChartOptions, "type" | "barDir">,
    range: AddChartRange,
  ): unknown;
  addBarChart(options: Omit<AddBarChartOptions, "type" | "barDir">, range: AddChartRange): unknown;
  addLineChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addAreaChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addPieChart(options: Omit<AddPieChartOptions, "type">, range: AddChartRange): unknown;
  addDoughnutChart(options: Omit<AddPieChartOptions, "type">, range: AddChartRange): unknown;
  addScatterChart(options: Omit<AddScatterChartOptions, "type">, range: AddChartRange): unknown;
  addBubbleChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addRadarChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addStockChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): unknown;
  addSurfaceChart(options: Omit<AddSurfaceChartOptions, "type">, range: AddChartRange): unknown;
  addHistogramChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addParetoChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addWaterfallChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addFunnelChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addTreemapChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addSunburstChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addBoxWhiskerChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addRegionMapChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): unknown;
  addComboChart(options: AddComboChartOptions, range: AddChartRange): unknown;
  addConditionalFormatting(cf: ConditionalFormattingOptions): unknown;
  removeConditionalFormatting(
    filter?:
      | number
      | ((
          value: ConditionalFormattingOptions,
          index: number,
          array: ConditionalFormattingOptions[],
        ) => boolean),
  ): unknown;
  addCellIsRule(
    range: CellRange,
    operator: CellIsOperators,
    formulae: (string | number)[],
    style?: Partial<ExcelStyle>,
  ): unknown;
  addExpressionRule(range: CellRange, formula: string, style?: Partial<ExcelStyle>): unknown;
  addDataBar(range: CellRange, color?: { argb?: string; theme?: number }): unknown;
  addColorScale(
    range: CellRange,
    cfvo: {
      type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
      value?: number | string;
    }[],
    colors?: { argb?: string; theme?: number }[],
  ): unknown;
  addIconSet(
    range: CellRange,
    iconSet?: IconSetTypes,
    cfvo?: { type: "percent" | "num" | "percentile" | "formula"; value?: number | string }[],
    options?: { showValue?: boolean; reverse?: boolean },
  ): unknown;
  addTop10Rule(
    range: CellRange,
    rank: number,
    options?: { percent?: boolean; bottom?: boolean; style?: Partial<ExcelStyle> },
  ): unknown;
  addAboveAverageRule(
    range: CellRange,
    options?: { aboveAverage?: boolean; style?: Partial<ExcelStyle> },
  ): unknown;
  addContainsTextRule(
    range: CellRange,
    text: string,
    operator?: ContainsTextOperators,
    style?: Partial<ExcelStyle>,
  ): unknown;
  addTimePeriodRule(
    range: CellRange,
    timePeriod: TimePeriodTypes,
    style?: Partial<ExcelStyle>,
  ): unknown;
  addDataValidation(addr: Addr, validation: DataValidation): unknown;
  removeDataValidation(address: string): unknown;
  addListValidation(
    range: CellRange,
    list: (string | number | Date)[],
    options?: Omit<DataValidationWithFormulae, "type" | "formulae" | "operator">,
  ): unknown;
  addRangeValidation(range: CellRange, validation: RangeValidationDef): unknown;
  addNote(addr: Addr, text: string): unknown;
  addNote(addr: Addr, config: NoteConfig): unknown;
  addThreadedComment(ref: string, comment: ThreadedComment): unknown;
  setCellHyperlink(addr: Addr, hyperlink: string, text?: string, tooltip?: string): unknown;
  setCellRichText(addr: Addr, richText: RichTextRun[]): unknown;
  addImage(imageId: string | number, range: AddImageRange): unknown;
  addBackgroundImage(imageId: string | number): unknown;
  addWatermark(options: WatermarkOptions): unknown;
  removeWatermark(): unknown;
  addSparklineGroup(options: AddSparklineGroupOptions): unknown;
  addFormCheckbox(range: FormControlRange, options?: FormCheckboxOptions): unknown;
  getFormCheckboxes(): ExcelFormCheckbox[];
  // ── Format mixin ─────────────────────────────────────────────────────
  freeze: (rowOrOpts: number | { row?: number; col?: number }, maybeCol?: number) => unknown;
  autoFilter: (range?: string) => unknown;
  autoFitColumns: (startCol?: number | string, endCol?: number | string) => unknown;
  rowHeight: (rowNumber: number, height: number) => unknown;
  colWidth: (col: string | number, width: number) => unknown;
  merge: (
    rangeOrRegion: string | MergeRange,
    options?: { value?: CellValue; style?: CellStyle; height?: number },
  ) => unknown;
  mergeAll: (regions: MergeRange[]) => unknown;
  styleRange: (range: CellRange, style: CellStyle) => unknown;
  clear: () => unknown;
  // ── Structure mixin ──────────────────────────────────────────────────
  insertRow: (pos: number, data: unknown[], options?: RowOptions) => unknown;
  duplicateRow: (rowNum: number, count: number, insert?: boolean) => unknown;
  removeRow: (start: number, count: number) => unknown;
  insertColumn: (start: number, count: number, ...inserts: unknown[][]) => unknown;
  removeColumn: (start: number, count: number) => unknown;
  addPageBreak: (rowNum: number) => unknown;
  addColumnPageBreak: (colNum: number) => unknown;
  // ── Data-io mixin ────────────────────────────────────────────────────
  eachRow: (callback: (row: unknown, rowNumber: number) => void) => void;
  toJSON: (opts?: unknown) => unknown;
  toAOA: () => unknown[][];
  addJSON: (data: Record<string, unknown>[], opts?: unknown) => unknown;
  addAOA: (data: unknown[][], opts?: unknown) => unknown;
}
