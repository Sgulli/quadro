import { Workbook } from "@cj-tech-master/excelts";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { SheetBuilder } from "../sheet-builder.js";
import type { NoteConfig, SheetOptions } from "../types.js";

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
  describe("headers", () => {
    it("sets column definitions and writes header row", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.headers([{ key: "a", header: "A" }]);
      expect(ws.columns).toHaveLength(1);
      expect(ws.columns[0].key).toBe("a");
      expect(ws.columns[0].width).toBe(15);
      expect(ws.columns[0].hidden).toBe(false);
      expect(ws.addRow).toHaveBeenCalledWith(["A"]);
    });

    it("calls addRow on the underlying worksheet", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.headers([{ key: "x", header: "X" }]);
      expect(ws.addRow).toHaveBeenCalledWith(["X"]);
    });
  });

  describe("addRow", () => {
    it("writes array data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRow(["a", "b"]);
      expect(ws.addRow).toHaveBeenCalledWith(["a", "b"]);
    });

    it("writes object data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.columns([
        { key: "name", header: "Name" },
        { key: "val", header: "Val" },
      ]);
      sheet.addRow({ name: "Alice", val: 42 });
      expect(ws.addRow).toHaveBeenCalledWith({ name: "Alice", val: 42 });
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRow([{ formula: "SUM(A1:A10)" }]);
      expect(ws.addRow).toHaveBeenCalledWith([{ formula: "SUM(A1:A10)" }]);
    });
  });

  describe("addRows", () => {
    it("adds multiple rows", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRows([[1], [2], [3]]);
      expect(ws.addRow).toHaveBeenCalledTimes(3);
    });
  });

  describe("setCell", () => {
    it("accepts a tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell([4, 5], 42);
      expect(ws.getCell("D5").value).toBe(42);
    });

    it("sets a value at the given address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("B3", 42);
      expect(ws.getCell("B3").value).toBe(42);
    });

    it("applies style when provided", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "hello", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
    });

    it("skips undefined value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1");
      expect(ws.getCell("A1").value).toBeFalsy();
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("C5", { formula: "SUM(A1:A10)" });
      expect(ws.getCell("C5").value).toEqual({ formula: "SUM(A1:A10)" });
    });
  });

  describe("styleRange", () => {
    it("applies style to a range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "a");
      sheet.setCell("B2", "b");
      sheet.styleRange("A1:B2", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
      expect(ws.getCell("B2").font?.bold).toBe(true);
    });

    it("handles single cell range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "a");
      const result = sheet.styleRange("A1:A1", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
      expect(result).toBe(sheet);
    });
  });

  describe("merge", () => {
    it("merges cells and writes value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "mergeCells");
      sheet.merge({ range: "A1:C1", value: "Title", style: { font: { bold: true } } });
      expect(ws.mergeCells).toHaveBeenCalledWith("A1:C1");
      expect(ws.getCell("A1").value).toBe("Title");
      expect(ws.getCell("A1").font?.bold).toBe(true);
    });

    it("accepts a tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "mergeCells");
      sheet.merge([1, 1, 3, 1], { value: "Title" });
      expect(ws.mergeCells).toHaveBeenCalledWith("A1:C1");
    });
  });

  describe("rowHeight / colWidth", () => {
    it("sets row height", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.rowHeight(1, 30);
      expect(ws.getRow(1).height).toBe(30);
    });

    it("sets column width", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.colWidth("A", 40);
      expect(ws.getColumn("A").width).toBe(40);
    });
  });

  describe("freeze", () => {
    it("sets freeze pane with positional args", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.freeze(3);
      const view = ws.views?.[0] ?? null;
      expect(view?.state).toBe("frozen");
      if (view?.state === "frozen") {
        expect(view.ySplit).toBe(3);
      }
    });

    it("sets freeze pane with object overload", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.freeze({ row: 2, col: 1 });
      const view = ws.views?.[0] ?? null;
      expect(view?.state).toBe("frozen");
      if (view?.state === "frozen") {
        expect(view.ySplit).toBe(2);
        expect(view.xSplit).toBe(1);
      }
    });
  });

  describe("autoFilter", () => {
    it("sets auto filter range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.autoFilter("A1:C1");
      expect(ws.autoFilter).toBe("A1:C1");
    });
  });

  describe("autoFitColumns", () => {
    it("auto-fits columns", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "autoFitColumns");
      sheet.autoFitColumns();
      expect(ws.autoFitColumns).toHaveBeenCalledWith(undefined, undefined);
    });
  });

  describe("rowCount / columnRange", () => {
    it("tracks row count after addRow", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow([1]);
      expect(sheet.rowCount).toBe(1);
      sheet.addRow([2]);
      expect(sheet.rowCount).toBe(2);
    });

    it("tracks row count after addRows", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRows([[1], [2], [3]]);
      expect(sheet.rowCount).toBe(3);
    });

    it("columnRange resolves column key to range string", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet
        .headers([
          { key: "name", header: "Name" },
          { key: "score", header: "Score" },
        ])
        .addRows([
          { name: "Alice", score: 85 },
          { name: "Bob", score: 42 },
        ]);
      expect(sheet.columnRange("name")).toBe("A2:A3");
      expect(sheet.columnRange("score")).toBe("B2:B3");
    });

    it("columnRange accepts custom startRow", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.headers([{ key: "x", header: "X" }]);
      sheet.addRow(["a"]);
      expect(sheet.columnRange("x", 1)).toBe("A1:A1");
    });

    it("columnRange throws for unknown key", () => {
      const { sheet } = makeSheet({ name: "Test" });
      expect(() => sheet.columnRange("missing")).toThrow("No column with key");
    });
  });

  describe("auto-inferred columns", () => {
    it("infers columns from first object row", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow({ name: "Alice", score: 85 });
      expect(sheet.columnRange("name")).toBe("A1:A1");
      expect(sheet.columnRange("score")).toBe("B1:B1");
    });

    it("columnRange works after addRows with inferred columns", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRows([
        { product: "Widget", quantity: 10 },
        { product: "Gadget", quantity: 20 },
        { product: "Doohickey", quantity: 30 },
      ]);
      expect(sheet.rowCount).toBe(3);
      expect(sheet.columnRange("product")).toBe("A1:A3");
      expect(sheet.columnRange("quantity")).toBe("B1:B3");
    });

    it("explicit columns take precedence over inference", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.headers([{ key: "custom", header: "Custom" }]);
      sheet.addRow({ name: "Alice", score: 85 });
      expect(sheet.columnRange("custom")).toBe("A2:A2");
      expect(() => sheet.columnRange("name")).toThrow("No column with key");
    });
  });

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

  describe("tables", () => {
    it("addTable delegates to ws.addTable", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTable("MyTable", "A1:C10", [{ name: "Name" }, { name: "Score" }]);
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({
        name: "MyTable",
        ref: "A1:C10",
        headerRow: true,
      });
    });

    it("addTable with rows and options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTable("Sales", "A1:D5", [{ name: "Product" }, { name: "Revenue" }], {
        rows: [
          ["Widget", 100],
          ["Gadget", 200],
        ],
        totalsRow: true,
        style: { theme: "TableStyleMedium9", showRowStripes: true },
      });
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({
        name: "Sales",
        ref: "A1:D5",
        totalsRow: true,
        style: { theme: "TableStyleMedium9", showRowStripes: true },
      });
    });

    it("addTable accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTable("Data", [1, 1, 3, 10], [{ name: "A" }, { name: "B" }, { name: "C" }]);
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({ name: "Data", ref: "A1:C10" });
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addTable("T1", "A1:B2", [{ name: "X" }])
        .addTable("T2", "D1:E2", [{ name: "Y" }]);
      expect(result).toBe(sheet);
    });
  });

  describe("hyperlinks", () => {
    it("setCellHyperlink sets cell value with hyperlink", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCellHyperlink("A1", "https://example.com", "Click here", "Visit site");
      const cell = ws.getCell("A1");
      expect(cell.isHyperlink).toBe(true);
      expect(cell.hyperlink).toBe("https://example.com");
    });

    it("setCellHyperlink uses hyperlink as default text", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCellHyperlink("B2", "https://example.com");
      const cell = ws.getCell("B2");
      expect(cell.isHyperlink).toBe(true);
      expect(cell.hyperlink).toBe("https://example.com");
    });

    it("setCellHyperlink accepts tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getCell");
      sheet.setCellHyperlink([3, 5], "https://example.com");
      expect(ws.getCell).toHaveBeenCalledWith("C5");
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .setCellHyperlink("A1", "https://a.com")
        .setCellHyperlink([2, 1], "https://b.com");
      expect(result).toBe(sheet);
    });
  });

  describe("comments / notes", () => {
    it("addNote sets cell note text", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addNote("A1", "Reviewed by John");
      const cell = ws.getCell("A1");
      expect(cell.note).toBeDefined();
    });

    it("addNote accepts tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getCell");
      sheet.addNote([2, 3], "Note text");
      expect(ws.getCell).toHaveBeenCalledWith("B3");
    });

    it("addThreadedComment pushes to ws.threadedComments", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.threadedComments, "push");
      sheet.addThreadedComment("A1", { personId: "{person-1}", text: "Great work!" });
      expect(ws.threadedComments.push).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.addNote("A1", "Note 1").addNote([2, 2], "Note 2");
      expect(result).toBe(sheet);
    });
  });

  describe("rich text", () => {
    it("setCellRichText sets rich text value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCellRichText("A1", [{ text: "Hello ", font: { bold: true } }, { text: "World" }]);
      const cell = ws.getCell("A1");
      expect(cell.value).toBeDefined();
    });

    it("setCellRichText accepts tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getCell");
      sheet.setCellRichText([1, 5], [{ text: "Test" }]);
      expect(ws.getCell).toHaveBeenCalledWith("A5");
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.setCellRichText("A1", [{ text: "Hi" }]);
      expect(result).toBe(sheet);
    });
  });

  describe("images", () => {
    it("addImage delegates to ws.addImage", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addImage");
      sheet.addImage(1, "A1:B5");
      expect(ws.addImage).toHaveBeenCalledWith(1, "A1:B5");
    });

    it("addBackgroundImage delegates to ws.addBackgroundImage", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addBackgroundImage");
      sheet.addBackgroundImage(1);
      expect(ws.addBackgroundImage).toHaveBeenCalledWith(1);
    });

    it("addWatermark delegates to ws.addWatermark", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addWatermark");
      sheet.addWatermark({ imageId: 1, opacity: 0.15 });
      expect(ws.addWatermark).toHaveBeenCalledWith({ imageId: 1, opacity: 0.15 });
    });

    it("removeWatermark delegates to ws.removeWatermark", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "removeWatermark");
      sheet.removeWatermark();
      expect(ws.removeWatermark).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.addImage(1, "A1:B5").addBackgroundImage(2);
      expect(result).toBe(sheet);
    });
  });

  describe("sparklines", () => {
    it("addSparklineGroup delegates to ws.addSparklineGroup", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addSparklineGroup");
      sheet.addSparklineGroup({
        type: "line",
        sparklines: [{ dataRef: "A1:C1", cellRef: "D1" }],
      });
      expect(ws.addSparklineGroup).toHaveBeenCalledWith({
        type: "line",
        sparklines: [{ dataRef: "A1:C1", cellRef: "D1" }],
      });
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.addSparklineGroup({
        type: "column",
        sparklines: [{ dataRef: "A1:B1", cellRef: "C1" }],
      });
      expect(result).toBe(sheet);
    });
  });

  describe("charts", () => {
    it("addChart delegates to ws.addChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addChart");
      sheet.addChart({ type: "bar", series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addChart).toHaveBeenCalledWith(
        { type: "bar", series: [{ values: "Sheet1!$B$2:$B$5" }] },
        "A1:D15",
      );
    });

    it("addBarChart delegates to ws.addBarChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addBarChart");
      sheet.addBarChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addBarChart).toHaveBeenCalledWith(
        { series: [{ values: "Sheet1!$B$2:$B$5" }] },
        "A1:D15",
      );
    });

    it("addColumnChart delegates to ws.addColumnChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addColumnChart");
      sheet.addColumnChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addColumnChart).toHaveBeenCalledWith(
        { series: [{ values: "Sheet1!$B$2:$B$5" }] },
        "A1:D15",
      );
    });

    it("addLineChart delegates to ws.addLineChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addLineChart");
      sheet.addLineChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addLineChart).toHaveBeenCalled();
    });

    it("addPieChart delegates to ws.addPieChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addPieChart");
      sheet.addPieChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addPieChart).toHaveBeenCalled();
    });

    it("addComboChart delegates to ws.addComboChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addComboChart");
      sheet.addComboChart(
        {
          groups: [{ type: "bar", series: [{ values: "Sheet1!$B$2:$B$5" }] }],
        },
        "A1:D15",
      );
      expect(ws.addComboChart).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addChart({ type: "bar", series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15")
        .addLineChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(result).toBe(sheet);
    });
  });

  // ── v0.5: Form Controls ──────────────────────────────────────────────────────

  describe("form controls", () => {
    it("addFormCheckbox delegates to ws.addFormCheckbox", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addFormCheckbox");
      sheet.addFormCheckbox("B2:C3", { text: "Approve", checked: true });
      expect(ws.addFormCheckbox).toHaveBeenCalledWith("B2:C3", { text: "Approve", checked: true });
    });

    it("getFormCheckboxes delegates to ws.getFormCheckboxes", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getFormCheckboxes");
      sheet.addFormCheckbox("A1", { text: "Yes" });
      sheet.getFormCheckboxes();
      expect(ws.getFormCheckboxes).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addFormCheckbox("A1", { text: "X" })
        .addFormCheckbox("B1", { text: "Y" });
      expect(result).toBe(sheet);
    });
  });

  // ── v0.5: Note Overloads ─────────────────────────────────────────────────────

  describe("addNote (v0.5 richer)", () => {
    it("accepts NoteConfig with texts", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      const config: NoteConfig = {
        texts: [{ text: "Line 1" }, { text: "Line 2" }],
      };
      sheet.addNote("A1", config);
      const cell = ws.getCell("A1");
      expect(cell.note).toBeDefined();
    });

    it("accepts string directly (backward compat)", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addNote("B2", "Simple note");
      expect(ws.getCell("B2").note).toBe("Simple note");
    });
  });

  // ── v0.5: Row Operations ─────────────────────────────────────────────────────

  describe("insertRow", () => {
    it("inserts a row at the given position", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "insertRow");
      sheet.insertRow(1, ["a", "b"]);
      expect(ws.insertRow).toHaveBeenCalledWith(1, ["a", "b"]);
    });

    it("inserts a row and updates rowCount", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow([1]);
      expect(sheet.rowCount).toBe(1);
      sheet.insertRow(1, ["a"]);
      expect(sheet.rowCount).toBe(2);
    });

    it("accepts row options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.insertRow(1, ["data"], { height: 30 });
      expect(ws.getRow(1).height).toBe(30);
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.insertRow(1, ["a"]).insertRow(2, ["b"]);
      expect(result).toBe(sheet);
    });
  });

  describe("duplicateRow", () => {
    it("duplicates a row and updates rowCount", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "duplicateRow");
      sheet.addRow([1]);
      sheet.duplicateRow(1, 2, true);
      expect(ws.duplicateRow).toHaveBeenCalledWith(1, 2, true);
      expect(sheet.rowCount).toBe(3);
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow([1]);
      const result = sheet.duplicateRow(1, 1);
      expect(result).toBe(sheet);
      expect(sheet.rowCount).toBe(2);
    });
  });

  describe("removeRow", () => {
    it("removes rows and updates rowCount", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "spliceRows");
      sheet.addRows([[1], [2], [3]]);
      expect(sheet.rowCount).toBe(3);
      sheet.removeRow(2, 2);
      expect(ws.spliceRows).toHaveBeenCalledWith(2, 2);
      expect(sheet.rowCount).toBe(1);
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRows([[1], [2]]);
      const result = sheet.removeRow(1, 1);
      expect(result).toBe(sheet);
    });
  });

  // ── v0.5: Column Operations ──────────────────────────────────────────────────

  describe("insertColumn / removeColumn", () => {
    it("insertColumn delegates to ws.spliceColumns", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "spliceColumns");
      sheet.insertColumn(2, 1);
      expect(ws.spliceColumns).toHaveBeenCalledWith(2, 1);
    });

    it("insertColumn with data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "spliceColumns");
      sheet.insertColumn(2, 1, ["a"], ["b"]);
      expect(ws.spliceColumns).toHaveBeenCalledWith(2, 1, ["a"], ["b"]);
    });

    it("removeColumn delegates to ws.spliceColumns", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "spliceColumns");
      sheet.removeColumn(3, 2);
      expect(ws.spliceColumns).toHaveBeenCalledWith(3, 2);
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.insertColumn(1, 1).removeColumn(2, 1);
      expect(result).toBe(sheet);
    });
  });

  // ── v0.5: Page Breaks ────────────────────────────────────────────────────────

  describe("page breaks", () => {
    it("addPageBreak adds a row page break", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      const row = ws.getRow(5);
      vi.spyOn(row, "addPageBreak");
      sheet.addPageBreak(5);
      expect(row.addPageBreak).toHaveBeenCalled();
    });

    it("addColumnPageBreak adds a column page break", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      const col = ws.getColumn(3);
      vi.spyOn(col, "addPageBreak");
      sheet.addColumnPageBreak(3);
      expect(col.addPageBreak).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.addPageBreak(10).addColumnPageBreak(5);
      expect(result).toBe(sheet);
    });
  });

  // ── v0.5: Ignored Errors ─────────────────────────────────────────────────────

  describe("addIgnoredError", () => {
    it("pushes to ws.ignoredErrors", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.ignoredErrors, "push");
      sheet.addIgnoredError("A1:A100", { numberStoredAsText: true });
      expect(ws.ignoredErrors.push).toHaveBeenCalledWith({
        ref: "A1:A100",
        numberStoredAsText: true,
      });
    });

    it("accepts multiple error types", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.ignoredErrors, "push");
      sheet.addIgnoredError("C1:C10", {
        formula: true,
        evalError: true,
        emptyCellReference: true,
      });
      expect(ws.ignoredErrors.push).toHaveBeenCalledWith({
        ref: "C1:C10",
        formula: true,
        evalError: true,
        emptyCellReference: true,
      });
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addIgnoredError("A1:A10", { numberStoredAsText: true })
        .addIgnoredError("B1:B10", { formula: true });
      expect(result).toBe(sheet);
    });
  });

  // ── v0.5: worksheet getter ───────────────────────────────────────────────────

  describe("worksheet getter", () => {
    it("returns the underlying excelts worksheet", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      expect(sheet.worksheet).toBe(ws);
    });
  });
});
