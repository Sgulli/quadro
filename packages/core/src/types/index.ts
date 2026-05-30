export type * from "./options.js";
export type * from "./style.js";
export { BUILTIN_FORMATS } from "./style.js";
export type * from "./validation.js";

export type CellHyperlinkValueInput = import("@cj-tech-master/excelts").CellHyperlinkValueInput;
export type CellRichTextValue = import("@cj-tech-master/excelts").CellRichTextValue;
export type ExcelFormCheckbox = import("@cj-tech-master/excelts").FormCheckbox;
export type RichTextRun = import("@cj-tech-master/excelts").RichText;

export interface SheetBuilderExtension {
  addChart(
    options: import("@cj-tech-master/excelts/chart").AddChartOptions,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addColumnChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddBarChartOptions, "type" | "barDir">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addBarChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddBarChartOptions, "type" | "barDir">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addLineChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addAreaChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addPieChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddPieChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addDoughnutChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddPieChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addScatterChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddScatterChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addBubbleChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addRadarChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addStockChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addSurfaceChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddSurfaceChartOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addHistogramChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartExOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addParetoChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartExOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addWaterfallChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartExOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addFunnelChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartExOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addTreemapChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartExOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addSunburstChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartExOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addBoxWhiskerChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartExOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addRegionMapChart(
    options: Omit<import("@cj-tech-master/excelts/chart").AddChartExOptions, "type">,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addComboChart(
    options: import("@cj-tech-master/excelts/chart").AddComboChartOptions,
    range: import("@cj-tech-master/excelts/chart").AddChartRange,
  ): unknown;
  addConditionalFormatting(
    cf: import("@cj-tech-master/excelts").ConditionalFormattingOptions,
  ): unknown;
  removeConditionalFormatting(
    filter?:
      | number
      | ((
          value: import("@cj-tech-master/excelts").ConditionalFormattingOptions,
          index: number,
          array: import("@cj-tech-master/excelts").ConditionalFormattingOptions[],
        ) => boolean),
  ): unknown;
  addCellIsRule(
    range: import("./validation.js").CellRange,
    operator: import("@cj-tech-master/excelts").CellIsOperators,
    formulae: (string | number)[],
    style?: Partial<import("@cj-tech-master/excelts").Style>,
  ): unknown;
  addExpressionRule(
    range: import("./validation.js").CellRange,
    formula: string,
    style?: Partial<import("@cj-tech-master/excelts").Style>,
  ): unknown;
  addDataBar(
    range: import("./validation.js").CellRange,
    color?: { argb?: string; theme?: number },
  ): unknown;
  addColorScale(
    range: import("./validation.js").CellRange,
    cfvo: {
      type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
      value?: number | string;
    }[],
    colors?: { argb?: string; theme?: number }[],
  ): unknown;
  addIconSet(
    range: import("./validation.js").CellRange,
    iconSet?: import("@cj-tech-master/excelts").IconSetTypes,
    cfvo?: {
      type: "percent" | "num" | "percentile" | "formula";
      value?: number | string;
    }[],
    options?: { showValue?: boolean; reverse?: boolean },
  ): unknown;
  addTop10Rule(
    range: import("./validation.js").CellRange,
    rank: number,
    options?: {
      percent?: boolean;
      bottom?: boolean;
      style?: Partial<import("@cj-tech-master/excelts").Style>;
    },
  ): unknown;
  addAboveAverageRule(
    range: import("./validation.js").CellRange,
    options?: {
      aboveAverage?: boolean;
      style?: Partial<import("@cj-tech-master/excelts").Style>;
    },
  ): unknown;
  addContainsTextRule(
    range: import("./validation.js").CellRange,
    text: string,
    operator?: import("@cj-tech-master/excelts").ContainsTextOperators,
    style?: Partial<import("@cj-tech-master/excelts").Style>,
  ): unknown;
  addTimePeriodRule(
    range: import("./validation.js").CellRange,
    timePeriod: import("@cj-tech-master/excelts").TimePeriodTypes,
    style?: Partial<import("@cj-tech-master/excelts").Style>,
  ): unknown;
  addDataValidation(
    addr: import("./validation.js").Addr,
    validation: import("@cj-tech-master/excelts").DataValidation,
  ): unknown;
  removeDataValidation(address: string): unknown;
  addListValidation(
    range: import("./validation.js").CellRange,
    list: (string | number | Date)[],
    options?: Omit<
      import("@cj-tech-master/excelts").DataValidationWithFormulae,
      "type" | "formulae" | "operator"
    >,
  ): unknown;
  addRangeValidation(
    range: import("./validation.js").CellRange,
    validation: import("./validation.js").RangeValidationDef,
  ): unknown;
  addNote(addr: import("./validation.js").Addr, text: string): unknown;
  addNote(addr: import("./validation.js").Addr, config: import("./options.js").NoteConfig): unknown;
  addThreadedComment(
    ref: string,
    comment: import("@cj-tech-master/excelts").ThreadedComment,
  ): unknown;
  setCellHyperlink(
    addr: import("./validation.js").Addr,
    hyperlink: string,
    text?: string,
    tooltip?: string,
  ): unknown;
  setCellRichText(addr: import("./validation.js").Addr, richText: RichTextRun[]): unknown;
  addImage(
    imageId: string | number,
    range: import("@cj-tech-master/excelts").AddImageRange,
  ): unknown;
  addBackgroundImage(imageId: string | number): unknown;
  addWatermark(options: import("@cj-tech-master/excelts").WatermarkOptions): unknown;
  removeWatermark(): unknown;
  addSparklineGroup(options: import("./options.js").AddSparklineGroupOptions): unknown;
  addFormCheckbox(
    range: import("@cj-tech-master/excelts").FormControlRange,
    options?: import("@cj-tech-master/excelts").FormCheckboxOptions,
  ): unknown;
  getFormCheckboxes(): ExcelFormCheckbox[];
  freeze: (rowOrOpts: number | { row?: number; col?: number }, maybeCol?: number) => unknown;
  autoFilter: (range?: string) => unknown;
  autoFitColumns: (startCol?: number | string, endCol?: number | string) => unknown;
  rowHeight: (rowNumber: number, height: number) => unknown;
  colWidth: (col: string | number, width: number) => unknown;
  merge: (
    rangeOrRegion: string | import("./options.js").MergeRange,
    options?: {
      value?: import("./options.js").CellValue;
      style?: import("./style.js").CellStyle;
      height?: number;
    },
  ) => unknown;
  mergeAll: (regions: import("./options.js").MergeRange[]) => unknown;
  styleRange: (
    range: import("./validation.js").CellRange,
    style: import("./style.js").CellStyle,
  ) => unknown;
  clear: () => unknown;
  insertRow: (pos: number, data: unknown[], options?: import("./options.js").RowOptions) => unknown;
  duplicateRow: (rowNum: number, count: number, insert?: boolean) => unknown;
  removeRow: (start: number, count: number) => unknown;
  insertColumn: (start: number, count: number, ...inserts: unknown[][]) => unknown;
  removeColumn: (start: number, count: number) => unknown;
  addPageBreak: (rowNum: number) => unknown;
  addColumnPageBreak: (colNum: number) => unknown;
  eachRow: (callback: (row: unknown, rowNumber: number) => void) => void;
  toJSON: (opts?: unknown) => unknown;
  toAOA: () => unknown[][];
  addJSON: (data: Record<string, unknown>[], opts?: unknown) => unknown;
  addAOA: (data: unknown[][], opts?: unknown) => unknown;
}
