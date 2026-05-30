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

import type { SheetBuilderExtension } from "../types.js";

declare module "../builders/sheet-builder.js" {
  interface SheetBuilder {
    addChart(options: AddChartOptions, range: AddChartRange): this;
    addColumnChart(
      options: Omit<AddBarChartOptions, "type" | "barDir">,
      range: AddChartRange,
    ): this;
    addBarChart(options: Omit<AddBarChartOptions, "type" | "barDir">, range: AddChartRange): this;
    addLineChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): this;
    addAreaChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): this;
    addPieChart(options: Omit<AddPieChartOptions, "type">, range: AddChartRange): this;
    addDoughnutChart(options: Omit<AddPieChartOptions, "type">, range: AddChartRange): this;
    addScatterChart(options: Omit<AddScatterChartOptions, "type">, range: AddChartRange): this;
    addBubbleChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): this;
    addRadarChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): this;
    addStockChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): this;
    addSurfaceChart(options: Omit<AddSurfaceChartOptions, "type">, range: AddChartRange): this;
    addHistogramChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): this;
    addParetoChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): this;
    addWaterfallChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): this;
    addFunnelChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): this;
    addTreemapChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): this;
    addSunburstChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): this;
    addBoxWhiskerChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): this;
    addRegionMapChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): this;
    addComboChart(options: AddComboChartOptions, range: AddChartRange): this;
  }
}

interface ChartWs {
  addChart(options: AddChartOptions, range: AddChartRange): void;
  addColumnChart(options: Omit<AddBarChartOptions, "type" | "barDir">, range: AddChartRange): void;
  addBarChart(options: Omit<AddBarChartOptions, "type" | "barDir">, range: AddChartRange): void;
  addLineChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): void;
  addAreaChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): void;
  addPieChart(options: Omit<AddPieChartOptions, "type">, range: AddChartRange): void;
  addDoughnutChart(options: Omit<AddPieChartOptions, "type">, range: AddChartRange): void;
  addScatterChart(options: Omit<AddScatterChartOptions, "type">, range: AddChartRange): void;
  addBubbleChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): void;
  addRadarChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): void;
  addStockChart(options: Omit<AddChartOptions, "type">, range: AddChartRange): void;
  addSurfaceChart(options: Omit<AddSurfaceChartOptions, "type">, range: AddChartRange): void;
  addHistogramChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): void;
  addParetoChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): void;
  addWaterfallChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): void;
  addFunnelChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): void;
  addTreemapChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): void;
  addSunburstChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): void;
  addBoxWhiskerChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): void;
  addRegionMapChart(options: Omit<AddChartExOptions, "type">, range: AddChartRange): void;
  addComboChart(options: AddComboChartOptions, range: AddChartRange): void;
}

interface ChartMethods {
  worksheet: ChartWs;
}

export function applyChartMixin(proto: SheetBuilderExtension): void {
  proto.addChart = function (this: ChartMethods, options: AddChartOptions, range: AddChartRange) {
    this.worksheet.addChart(options, range);
    return this;
  };

  proto.addColumnChart = function (
    this: ChartMethods,
    options: Omit<AddBarChartOptions, "type" | "barDir">,
    range: AddChartRange,
  ) {
    this.worksheet.addColumnChart(options, range);
    return this;
  };

  proto.addBarChart = function (
    this: ChartMethods,
    options: Omit<AddBarChartOptions, "type" | "barDir">,
    range: AddChartRange,
  ) {
    this.worksheet.addBarChart(options, range);
    return this;
  };

  proto.addLineChart = function (
    this: ChartMethods,
    options: Omit<AddChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addLineChart(options, range);
    return this;
  };

  proto.addAreaChart = function (
    this: ChartMethods,
    options: Omit<AddChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addAreaChart(options, range);
    return this;
  };

  proto.addPieChart = function (
    this: ChartMethods,
    options: Omit<AddPieChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addPieChart(options, range);
    return this;
  };

  proto.addDoughnutChart = function (
    this: ChartMethods,
    options: Omit<AddPieChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addDoughnutChart(options, range);
    return this;
  };

  proto.addScatterChart = function (
    this: ChartMethods,
    options: Omit<AddScatterChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addScatterChart(options, range);
    return this;
  };

  proto.addBubbleChart = function (
    this: ChartMethods,
    options: Omit<AddChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addBubbleChart(options, range);
    return this;
  };

  proto.addRadarChart = function (
    this: ChartMethods,
    options: Omit<AddChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addRadarChart(options, range);
    return this;
  };

  proto.addStockChart = function (
    this: ChartMethods,
    options: Omit<AddChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addStockChart(options, range);
    return this;
  };

  proto.addSurfaceChart = function (
    this: ChartMethods,
    options: Omit<AddSurfaceChartOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addSurfaceChart(options, range);
    return this;
  };

  proto.addHistogramChart = function (
    this: ChartMethods,
    options: Omit<AddChartExOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addHistogramChart(options, range);
    return this;
  };

  proto.addParetoChart = function (
    this: ChartMethods,
    options: Omit<AddChartExOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addParetoChart(options, range);
    return this;
  };

  proto.addWaterfallChart = function (
    this: ChartMethods,
    options: Omit<AddChartExOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addWaterfallChart(options, range);
    return this;
  };

  proto.addFunnelChart = function (
    this: ChartMethods,
    options: Omit<AddChartExOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addFunnelChart(options, range);
    return this;
  };

  proto.addTreemapChart = function (
    this: ChartMethods,
    options: Omit<AddChartExOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addTreemapChart(options, range);
    return this;
  };

  proto.addSunburstChart = function (
    this: ChartMethods,
    options: Omit<AddChartExOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addSunburstChart(options, range);
    return this;
  };

  proto.addBoxWhiskerChart = function (
    this: ChartMethods,
    options: Omit<AddChartExOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addBoxWhiskerChart(options, range);
    return this;
  };

  proto.addRegionMapChart = function (
    this: ChartMethods,
    options: Omit<AddChartExOptions, "type">,
    range: AddChartRange,
  ) {
    this.worksheet.addRegionMapChart(options, range);
    return this;
  };

  proto.addComboChart = function (
    this: ChartMethods,
    options: AddComboChartOptions,
    range: AddChartRange,
  ) {
    this.worksheet.addComboChart(options, range);
    return this;
  };
}
