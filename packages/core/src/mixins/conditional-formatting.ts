import type {
  AboveAverageRuleType,
  CellIsOperators,
  CellIsRuleType,
  ColorScaleRuleType,
  ConditionalFormattingOptions,
  ContainsTextOperators,
  ContainsTextRuleType,
  DataBarRuleType,
  Style as ExcelStyle,
  ExpressionRuleType,
  IconSetRuleType,
  IconSetTypes,
  TimePeriodRuleType,
  TimePeriodTypes,
  Top10RuleType,
} from "@cj-tech-master/excelts";
import { resolveRange } from "../coords/coords.js";
import type { CellRange, SheetBuilderExtension } from "../types.js";

declare module "../builders/sheet-builder.js" {
  interface SheetBuilder {
    addConditionalFormatting(cf: ConditionalFormattingOptions): this;
    removeConditionalFormatting(
      filter?:
        | number
        | ((
            value: ConditionalFormattingOptions,
            index: number,
            array: ConditionalFormattingOptions[],
          ) => boolean),
    ): this;
    addCellIsRule(
      range: CellRange,
      operator: CellIsOperators,
      formulae: (string | number)[],
      style?: Partial<ExcelStyle>,
    ): this;
    addExpressionRule(range: CellRange, formula: string, style?: Partial<ExcelStyle>): this;
    addDataBar(range: CellRange, color?: { argb?: string; theme?: number }): this;
    addColorScale(
      range: CellRange,
      cfvo: {
        type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
        value?: number | string;
      }[],
      colors?: { argb?: string; theme?: number }[],
    ): this;
    addIconSet(
      range: CellRange,
      iconSet?: IconSetTypes,
      cfvo?: {
        type: "percent" | "num" | "percentile" | "formula";
        value?: number | string;
      }[],
      options?: { showValue?: boolean; reverse?: boolean },
    ): this;
    addTop10Rule(
      range: CellRange,
      rank: number,
      options?: { percent?: boolean; bottom?: boolean; style?: Partial<ExcelStyle> },
    ): this;
    addAboveAverageRule(
      range: CellRange,
      options?: { aboveAverage?: boolean; style?: Partial<ExcelStyle> },
    ): this;
    addContainsTextRule(
      range: CellRange,
      text: string,
      operator?: ContainsTextOperators,
      style?: Partial<ExcelStyle>,
    ): this;
    addTimePeriodRule(
      range: CellRange,
      timePeriod: TimePeriodTypes,
      style?: Partial<ExcelStyle>,
    ): this;
  }
}

interface CfWs {
  addConditionalFormatting(cf: ConditionalFormattingOptions): void;
  removeConditionalFormatting(
    filter?:
      | number
      | ((
          value: ConditionalFormattingOptions,
          index: number,
          array: ConditionalFormattingOptions[],
        ) => boolean),
  ): void;
}

interface CfMethods {
  worksheet: CfWs;
}

export function applyConditionalFormattingMixin(proto: SheetBuilderExtension): void {
  proto.addConditionalFormatting = function (this: CfMethods, cf: ConditionalFormattingOptions) {
    this.worksheet.addConditionalFormatting(cf);
    return this;
  };

  proto.removeConditionalFormatting = function (
    this: CfMethods,
    filter?:
      | number
      | ((
          value: ConditionalFormattingOptions,
          index: number,
          array: ConditionalFormattingOptions[],
        ) => boolean),
  ) {
    this.worksheet.removeConditionalFormatting(filter);
    return this;
  };

  proto.addCellIsRule = function (
    this: CfMethods,
    range: CellRange,
    operator: CellIsOperators,
    formulae: (string | number)[],
    style?: Partial<ExcelStyle>,
  ) {
    const rule: CellIsRuleType = { type: "cellIs", operator, formulae };
    if (style) rule.style = style;
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };

  proto.addExpressionRule = function (
    this: CfMethods,
    range: CellRange,
    formula: string,
    style?: Partial<ExcelStyle>,
  ) {
    const rule: ExpressionRuleType = { type: "expression", formulae: [formula] };
    if (style) rule.style = style;
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };

  proto.addDataBar = function (
    this: CfMethods,
    range: CellRange,
    color?: { argb?: string; theme?: number },
  ) {
    const rule: DataBarRuleType = { type: "dataBar" };
    if (color) rule.color = color;
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };

  proto.addColorScale = function (
    this: CfMethods,
    range: CellRange,
    cfvo: {
      type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
      value?: number | string;
    }[],
    colors?: { argb?: string; theme?: number }[],
  ) {
    const rule: ColorScaleRuleType = { type: "colorScale", cfvo, color: colors };
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };

  proto.addIconSet = function (
    this: CfMethods,
    range: CellRange,
    iconSet?: IconSetTypes,
    cfvo?: {
      type: "percent" | "num" | "percentile" | "formula";
      value?: number | string;
    }[],
    options?: { showValue?: boolean; reverse?: boolean },
  ) {
    const rule: IconSetRuleType = { type: "iconSet", iconSet, cfvo, ...options };
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };

  proto.addTop10Rule = function (
    this: CfMethods,
    range: CellRange,
    rank: number,
    options?: {
      percent?: boolean;
      bottom?: boolean;
      style?: Partial<ExcelStyle>;
    },
  ) {
    const rule: Top10RuleType = {
      type: "top10",
      rank,
      percent: options?.percent ?? false,
      bottom: options?.bottom,
    };
    if (options?.style) rule.style = options.style;
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };

  proto.addAboveAverageRule = function (
    this: CfMethods,
    range: CellRange,
    options?: { aboveAverage?: boolean; style?: Partial<ExcelStyle> },
  ) {
    const rule: AboveAverageRuleType = {
      type: "aboveAverage",
      aboveAverage: options?.aboveAverage,
    };
    if (options?.style) rule.style = options.style;
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };

  proto.addContainsTextRule = function (
    this: CfMethods,
    range: CellRange,
    text: string,
    operator?: ContainsTextOperators,
    style?: Partial<ExcelStyle>,
  ) {
    const rule: ContainsTextRuleType = { type: "containsText", operator, text };
    if (style) rule.style = style;
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };

  proto.addTimePeriodRule = function (
    this: CfMethods,
    range: CellRange,
    timePeriod: TimePeriodTypes,
    style?: Partial<ExcelStyle>,
  ) {
    const rule: TimePeriodRuleType = { type: "timePeriod", timePeriod };
    if (style) rule.style = style;
    this.worksheet.addConditionalFormatting({ ref: resolveRange(range), rules: [rule] });
    return this;
  };
}
