import { Workbook } from "@cj-tech-master/excelts";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { SheetBuilder } from "../../builders/sheet-builder.js";
import type { SheetOptions } from "../../types.js";

function makeSheet(opts: SheetOptions) {
  const wb = new Workbook();
  const ws = wb.addWorksheet(opts.name);
  const sheet = new SheetBuilder(ws, opts);
  return { ws, sheet };
}

beforeEach(() => {
  vi.restoreAllMocks();
});

describe("SheetBuilder", () => {
  describe("data validation", () => {
    it("addDataValidation delegates to ws.dataValidations.add", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addDataValidation("A1", { type: "any" });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("A1", { type: "any" });
    });

    it("addListValidation creates list type validation", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addListValidation("B2:B10", ["Option A", "Option B"]);
      expect(ws.dataValidations.add).toHaveBeenCalledWith("B2:B10", {
        type: "list",
        formulae: ['"Option A"', '"Option B"'],
      });
    });

    it("addListValidation forwards options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addListValidation("C1", ["X"], { allowBlank: true, prompt: "Pick one" });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("C1", {
        type: "list",
        formulae: ['"X"'],
        allowBlank: true,
        prompt: "Pick one",
      });
    });

    it("addRangeValidation with between operator", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addRangeValidation("D1:D10", {
        type: "whole",
        operator: "between",
        formulae: [1, 100],
      });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("D1:D10", {
        type: "whole",
        operator: "between",
        formulae: [1, 100],
      });
    });

    it("addRangeValidation forwards options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addRangeValidation("E1", {
        type: "decimal",
        operator: "lessThan",
        formulae: [50],
        error: "Too high",
      });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("E1", {
        type: "decimal",
        operator: "lessThan",
        formulae: [50],
        error: "Too high",
      });
    });
  });

  describe("conditional formatting", () => {
    it("addConditionalFormatting delegates to ws.addConditionalFormatting", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addConditionalFormatting({
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "greaterThan", formulae: [100] }],
      });
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "greaterThan", formulae: [100] }],
      });
    });

    it("removeConditionalFormatting delegates to ws", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "removeConditionalFormatting");
      sheet.removeConditionalFormatting(0);
      expect(ws.removeConditionalFormatting).toHaveBeenCalledWith(0);
    });

    it("addCellIsRule adds a cellIs rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addCellIsRule("A1:A10", "between", [10, 100]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:A10",
        rules: [{ type: "cellIs", operator: "between", formulae: [10, 100] }],
      });
    });

    it("addCellIsRule with style", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addCellIsRule("A1", "equal", ["Yes"], { font: { bold: true } });
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1",
        rules: [
          { type: "cellIs", operator: "equal", formulae: ["Yes"], style: { font: { bold: true } } },
        ],
      });
    });

    it("addExpressionRule adds an expression rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addExpressionRule("A1:A10", "A1>100");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:A10",
        rules: [{ type: "expression", formulae: ["A1>100"] }],
      });
    });

    it("addDataBar adds a data bar rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addDataBar("B1:B10");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "B1:B10",
        rules: [{ type: "dataBar" }],
      });
    });

    it("addColorScale adds a color scale rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addColorScale("C1:C10", [{ type: "min" }, { type: "max" }]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "C1:C10",
        rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "max" }] }],
      });
    });

    it("addIconSet adds an icon set rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addIconSet("D1:D10", "3TrafficLights1");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "D1:D10",
        rules: [{ type: "iconSet", iconSet: "3TrafficLights1" }],
      });
    });

    it("addTop10Rule adds a top10 rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addTop10Rule("E1:E10", 5);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "E1:E10",
        rules: [{ type: "top10", rank: 5, percent: false }],
      });
    });

    it("addAboveAverageRule adds an aboveAverage rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addAboveAverageRule("F1:F10");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "F1:F10",
        rules: [{ type: "aboveAverage" }],
      });
    });

    it("addContainsTextRule adds a containsText rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addContainsTextRule("G1:G10", "urgent");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "G1:G10",
        rules: [{ type: "containsText", text: "urgent" }],
      });
    });

    it("addTimePeriodRule adds a timePeriod rule", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addTimePeriodRule("H1:H10", "thisMonth");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "H1:H10",
        rules: [{ type: "timePeriod", timePeriod: "thisMonth" }],
      });
    });

    it("methods are chainable", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      vi.spyOn(ws, "addConditionalFormatting");
      sheet
        .addDataValidation("A1", { type: "any" })
        .addConditionalFormatting({ ref: "B1", rules: [{ type: "dataBar" }] })
        .addCellIsRule("C1", "greaterThan", [0]);
      expect(ws.dataValidations.add).toHaveBeenCalledTimes(1);
      expect(ws.addConditionalFormatting).toHaveBeenCalledTimes(2);
    });
  });

  describe("data validation with tuples", () => {
    it("addDataValidation accepts tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addDataValidation([3, 5], { type: "any" });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("C5", { type: "any" });
    });

    it("addListValidation accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addListValidation([2, 3, 2, 10], ["X", "Y"]);
      expect(ws.dataValidations.add).toHaveBeenCalledWith("B3:B10", {
        type: "list",
        formulae: ['"X"', '"Y"'],
      });
    });

    it("addRangeValidation accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.dataValidations, "add");
      sheet.addRangeValidation([4, 2, 4, 100], {
        type: "whole",
        operator: "between",
        formulae: [1, 100],
      });
      expect(ws.dataValidations.add).toHaveBeenCalledWith("D2:D100", {
        type: "whole",
        operator: "between",
        formulae: [1, 100],
      });
    });
  });

  describe("conditional formatting with tuples", () => {
    it("addCellIsRule accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addCellIsRule([2, 3, 4, 10], "greaterThan", [100]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "B3:D10",
        rules: [{ type: "cellIs", operator: "greaterThan", formulae: [100] }],
      });
    });

    it("addExpressionRule accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addExpressionRule([1, 1, 5, 5], "A1>0");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "A1:E5",
        rules: [{ type: "expression", formulae: ["A1>0"] }],
      });
    });

    it("addDataBar accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addDataBar([2, 2, 2, 20]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "B2:B20",
        rules: [{ type: "dataBar" }],
      });
    });

    it("addColorScale accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addColorScale([3, 1, 3, 10], [{ type: "min" }, { type: "max" }]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "C1:C10",
        rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "max" }] }],
      });
    });

    it("addIconSet accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addIconSet([4, 1, 4, 10], "3TrafficLights1");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "D1:D10",
        rules: [{ type: "iconSet", iconSet: "3TrafficLights1" }],
      });
    });

    it("addTop10Rule accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addTop10Rule([5, 1, 5, 10], 5);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "E1:E10",
        rules: [{ type: "top10", rank: 5, percent: false }],
      });
    });

    it("addAboveAverageRule accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addAboveAverageRule([6, 1, 6, 10]);
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "F1:F10",
        rules: [{ type: "aboveAverage" }],
      });
    });

    it("addContainsTextRule accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addContainsTextRule([7, 1, 7, 10], "urgent");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "G1:G10",
        rules: [{ type: "containsText", text: "urgent" }],
      });
    });

    it("addTimePeriodRule accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addConditionalFormatting");
      sheet.addTimePeriodRule([8, 1, 8, 10], "thisMonth");
      expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
        ref: "H1:H10",
        rules: [{ type: "timePeriod", timePeriod: "thisMonth" }],
      });
    });
  });
});
