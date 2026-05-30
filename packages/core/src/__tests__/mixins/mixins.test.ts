import { Workbook } from "@cj-tech-master/excelts";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { SheetBuilder } from "../../builders/sheet-builder.js";
import type { SheetOptions } from "../../types.js";
import "../../mixins/charts.js";
import "../../mixins/conditional-formatting.js";
import "../../mixins/data-validation.js";
import "../../mixins/media.js";

function makeSheet(opts: SheetOptions) {
  const wb = new Workbook();
  const ws = wb.addWorksheet(opts.name);
  const sheet = new SheetBuilder(ws, opts);
  return { ws, sheet };
}

beforeEach(() => {
  vi.restoreAllMocks();
});

describe("charts", () => {
  it("addChart delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addChart").mockImplementation(() => 1);
    sheet.addChart({ type: "line" }, { tl: "A1", br: "B5" });
    expect(ws.addChart).toHaveBeenCalled();
  });

  it("addColumnChart delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addColumnChart").mockImplementation(() => 1);
    sheet.addColumnChart({}, { tl: "A1", br: "B5" });
    expect(ws.addColumnChart).toHaveBeenCalled();
  });

  it("addBarChart delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addBarChart").mockImplementation(() => 1);
    sheet.addBarChart({}, { tl: "A1", br: "B5" });
    expect(ws.addBarChart).toHaveBeenCalled();
  });

  it("addLineChart delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addLineChart").mockImplementation(() => 1);
    sheet.addLineChart({}, { tl: "A1", br: "B5" });
    expect(ws.addLineChart).toHaveBeenCalled();
  });

  it("addPieChart delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addPieChart").mockImplementation(() => 1);
    sheet.addPieChart({}, { tl: "A1", br: "B5" });
    expect(ws.addPieChart).toHaveBeenCalled();
  });

  it("addScatterChart delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addScatterChart").mockImplementation(() => 1);
    sheet.addScatterChart({}, { tl: "A1", br: "B5" });
    expect(ws.addScatterChart).toHaveBeenCalled();
  });
});

describe("conditional formatting", () => {
  it("addConditionalFormatting delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addConditionalFormatting");
    const cf = {
      ref: "A1:A10",
      rules: [{ type: "cellIs" as const, operator: "greaterThan" as const, formulae: [0] }],
    };
    sheet.addConditionalFormatting(cf);
    expect(ws.addConditionalFormatting).toHaveBeenCalledWith(cf);
  });

  it("removeConditionalFormatting delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "removeConditionalFormatting");
    sheet.removeConditionalFormatting(0);
    expect(ws.removeConditionalFormatting).toHaveBeenCalledWith(0);
  });

  it("addCellIsRule creates and adds rule", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addConditionalFormatting");
    sheet.addCellIsRule("A1:A5", "greaterThan", [10], { font: { bold: true } });
    expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
      ref: "A1:A5",
      rules: [
        {
          type: "cellIs",
          operator: "greaterThan",
          formulae: [10],
          style: { font: { bold: true } },
        },
      ],
    });
  });

  it("addExpressionRule creates expression rule", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addConditionalFormatting");
    sheet.addExpressionRule("B1:B10", "B1>100", { font: { italic: true } });
    expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
      ref: "B1:B10",
      rules: [{ type: "expression", formulae: ["B1>100"], style: { font: { italic: true } } }],
    });
  });

  it("addDataBar creates with optional color", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addConditionalFormatting");
    sheet.addDataBar("C1:C10", { argb: "FF0000FF" });
    expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
      ref: "C1:C10",
      rules: [{ type: "dataBar", color: { argb: "FF0000FF" } }],
    });
  });

  it("addTop10Rule with percent and bottom options", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addConditionalFormatting");
    sheet.addTop10Rule("D1:D10", 5, { percent: true, bottom: true });
    expect(ws.addConditionalFormatting).toHaveBeenCalledWith({
      ref: "D1:D10",
      rules: [{ type: "top10", rank: 5, percent: true, bottom: true }],
    });
  });
});

describe("data validation", () => {
  it("addDataValidation delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws.dataValidations, "add");
    sheet.addDataValidation("A1", { type: "whole", operator: "between", formulae: [1, 10] });
    expect(ws.dataValidations.add).toHaveBeenCalledWith("A1", {
      type: "whole",
      operator: "between",
      formulae: [1, 10],
    });
  });

  it("removeDataValidation delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws.dataValidations, "remove");
    sheet.removeDataValidation("A1");
    expect(ws.dataValidations.remove).toHaveBeenCalledWith("A1");
  });

  it("addListValidation formats list values", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws.dataValidations, "add");
    sheet.addListValidation("A1:A5", ["Yes", "No"]);
    expect(ws.dataValidations.add).toHaveBeenCalledWith("A1:A5", {
      type: "list",
      formulae: ['"Yes"', '"No"'],
    });
  });

  it("addRangeValidation delegates", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws.dataValidations, "add");
    const validation = { type: "whole" as const, operator: "between" as const, formulae: [0, 100] };
    sheet.addRangeValidation("A1:A10", validation);
    expect(ws.dataValidations.add).toHaveBeenCalledWith("A1:A10", validation);
  });
});

describe("media features", () => {
  it("addNote with string delegates to cell", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    sheet.addNote("A1", "My note");
    expect(ws.getCell("A1").note).toBe("My note");
  });

  it("setCellHyperlink sets hyperlink value", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    sheet.setCellHyperlink("A1", "https://example.com", "Click here", "Tooltip");
    const cell = ws.getCell("A1");
    expect(cell.value).toMatchObject({
      text: "Click here",
      hyperlink: "https://example.com",
      tooltip: "Tooltip",
    });
  });

  it("addImage delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addImage");
    sheet.addImage(1, { tl: { col: 1, row: 1 }, ext: { width: 200, height: 150 } });
    expect(ws.addImage).toHaveBeenCalledWith(1, {
      tl: { col: 1, row: 1 },
      ext: { width: 200, height: 150 },
    });
  });

  it("addFormCheckbox delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addFormCheckbox");
    sheet.addFormCheckbox({ tl: { col: 2, row: 3 }, br: { col: 4, row: 5 } });
    expect(ws.addFormCheckbox).toHaveBeenCalledWith(
      { tl: { col: 2, row: 3 }, br: { col: 4, row: 5 } },
      undefined,
    );
  });

  it("addSparklineGroup delegates to worksheet", () => {
    const { ws, sheet } = makeSheet({ name: "Test" });
    vi.spyOn(ws, "addSparklineGroup");
    const options = { type: "line" as const, sparklines: [{ dataRef: "A1:A5", cellRef: "B1" }] };
    sheet.addSparklineGroup(options);
    expect(ws.addSparklineGroup).toHaveBeenCalledWith(options);
  });
});
